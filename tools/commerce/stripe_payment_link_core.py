from __future__ import annotations

import hashlib
import json
import re
from dataclasses import dataclass, field
from typing import Any
from urllib.parse import urlparse


SUPPORTED_PRODUCT_KINDS = {
    "app_one_time",
    "pack_one_time_manual_delivery",
    "digital_product_one_time",
    "web_service_one_time_manual_fulfillment",
    "web_service_one_time_redirect",
}
SUPPORTED_PRICE_MODELS = {"one_time"}
SUPPORTED_CURRENCIES = {"jpy"}
CHECKOUT_NOTICE_MAX_LENGTH = 1200
PRODUCT_PLACEHOLDER = "__PRODUCT_ID_FROM_LIVE_PRODUCT__"
PRICE_PLACEHOLDER = "__PRICE_ID_FROM_LIVE_PRICE__"
SECRET_PATTERN = re.compile(r"sk_(?:test|live)_[A-Za-z0-9]{16,}|whsec_[A-Za-z0-9]{16,}")
PRIVATE_HOSTS = {"localhost", "127.0.0.1", "0.0.0.0", "::1"}


class CommerceSafetyStop(RuntimeError):
    pass


@dataclass(frozen=True)
class CommerceProductSpec:
    product_id: str
    product_kind: str
    title: str
    description: str
    amount: int
    currency: str
    price_model: str
    tax_code: str
    quantity: int = 1
    metadata: dict[str, Any] = field(default_factory=dict)
    checkout_notice: str = ""
    fulfillment_mode: str = ""
    after_completion_url: str | None = None


@dataclass(frozen=True)
class StripePaymentLinkPayloads:
    product_payload: dict[str, Any]
    price_payload: dict[str, Any]
    payment_link_payload: dict[str, Any]


def canonical_json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def sha256_payload(value: Any) -> str:
    return hashlib.sha256(canonical_json(value).encode("utf-8")).hexdigest()


def contains_secret_like_value(value: Any) -> bool:
    return bool(SECRET_PATTERN.search(canonical_json(value)))


def safe_product_id(product_id: str) -> str:
    safe = re.sub(r"[^A-Za-z0-9_-]+", "_", product_id).strip("_")
    return safe or "item"


def idempotency_key(prefix: str, product_id: str, payload: Any) -> str:
    return f"commerce-{prefix}-v1-{safe_product_id(product_id)}-{sha256_payload(payload)[:12]}"


def is_private_or_local_url(value: str) -> bool:
    parsed = urlparse(value)
    host = (parsed.hostname or "").lower()
    if parsed.scheme != "https":
        return True
    if host in PRIVATE_HOSTS:
        return True
    if host.startswith("10.") or host.startswith("192.168."):
        return True
    if re.match(r"^172\.(1[6-9]|2[0-9]|3[0-1])\.", host):
        return True
    if host.endswith(".local"):
        return True
    return False


def metadata_strings(metadata: dict[str, Any]) -> dict[str, str]:
    return {key: str(value) for key, value in metadata.items() if str(value) != ""}


def payment_metadata_for(spec: CommerceProductSpec) -> dict[str, str]:
    product_metadata = metadata_strings(spec.metadata)
    keys = (
        "dake_item_id",
        "dake_type",
        "source_original",
        "purchase_delivery_method",
        "delivery_policy",
    )
    payment_metadata = {key: product_metadata[key] for key in keys if product_metadata.get(key)}
    if payment_metadata:
        return payment_metadata
    return {
        "commerce_product_id": spec.product_id,
        "commerce_product_kind": spec.product_kind,
    }


def validate_spec(spec: CommerceProductSpec) -> None:
    if not spec.product_id.strip():
        raise CommerceSafetyStop("product_id is required")
    if spec.product_kind not in SUPPORTED_PRODUCT_KINDS:
        raise CommerceSafetyStop(f"unsupported product_kind: {spec.product_kind}")
    if spec.price_model not in SUPPORTED_PRICE_MODELS:
        raise CommerceSafetyStop("unsupported_in_commerce_v1")
    if not spec.title.strip():
        raise CommerceSafetyStop("title is required")
    if not isinstance(spec.amount, int) or spec.amount <= 0:
        raise CommerceSafetyStop("amount must be a positive integer")
    if spec.currency.lower() not in SUPPORTED_CURRENCIES:
        raise CommerceSafetyStop(f"unsupported currency: {spec.currency}")
    if not spec.tax_code.strip():
        raise CommerceSafetyStop("tax_code is required")
    if spec.quantity != 1:
        raise CommerceSafetyStop("commerce v1 supports quantity=1 only")
    if len(spec.checkout_notice) > CHECKOUT_NOTICE_MAX_LENGTH:
        raise CommerceSafetyStop("checkout notice length overflow")
    if contains_secret_like_value(spec.metadata):
        raise CommerceSafetyStop("secret-like metadata is not allowed")
    if spec.after_completion_url:
        if is_private_or_local_url(spec.after_completion_url):
            raise CommerceSafetyStop("private URL exposure is not allowed")
    if spec.product_kind == "web_service_one_time_redirect" and not spec.after_completion_url:
        raise CommerceSafetyStop("after_completion_url is required for redirect products")


def build_stripe_payment_link_payloads(
    spec: CommerceProductSpec,
    *,
    product_placeholder: str = PRODUCT_PLACEHOLDER,
    price_placeholder: str = PRICE_PLACEHOLDER,
) -> StripePaymentLinkPayloads:
    validate_spec(spec)
    product_metadata = metadata_strings(spec.metadata)
    payment_metadata = payment_metadata_for(spec)

    product_payload = {
        "name": spec.title,
        "description": spec.description,
        "active": True,
        "tax_code": spec.tax_code,
        "metadata": product_metadata,
    }
    price_payload = {
        "currency": spec.currency.lower(),
        "unit_amount": spec.amount,
        "product": product_placeholder,
        "metadata": {"dake_item_id": product_metadata.get("dake_item_id", spec.product_id)},
    }
    payment_link_payload: dict[str, Any] = {
        "line_items": [
            {
                "price": price_placeholder,
                "quantity": spec.quantity,
            }
        ],
        "metadata": payment_metadata,
        "payment_intent_data": {
            "metadata": payment_metadata,
        },
    }
    if spec.checkout_notice:
        payment_link_payload["custom_text"] = {
            "submit": {
                "message": spec.checkout_notice,
            },
        }
    if spec.after_completion_url:
        payment_link_payload["after_completion"] = {
            "type": "redirect",
            "redirect": {"url": spec.after_completion_url},
        }
    return StripePaymentLinkPayloads(
        product_payload=product_payload,
        price_payload=price_payload,
        payment_link_payload=payment_link_payload,
    )
