"""Normal DAKE Stripe Commerce flow.

Dry-run is read-only and does not import Stripe or read STRIPE_SECRET_KEY.
Live execution is entered by one explicit CLI intent and reconciles Stripe from
live read-back before creating anything or updating ORIGINAL.md.
"""

from __future__ import annotations

import copy
import json
import os
import re
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Callable

try:
    from store.stripe_release_core import (
        SafetyStop,
        build_release_payload,
        contains_secret_like_value,
        create_with_idempotency,
        object_to_dict,
        validate_dry_run_payload,
        validate_live_secret_key,
        write_json_atomic,
        write_text_atomic,
    )
except ModuleNotFoundError:  # Imported as tools.commerce.stripe_commerce_v2.
    from tools.store.stripe_release_core import (
        SafetyStop,
        build_release_payload,
        contains_secret_like_value,
        create_with_idempotency,
        object_to_dict,
        validate_dry_run_payload,
        validate_live_secret_key,
        write_json_atomic,
        write_text_atomic,
    )


ROOT = Path(__file__).resolve().parents[2]
ARTIFACT_ROOT = ROOT / "tools" / "reports" / "release_artifacts"
JST = timezone(timedelta(hours=9))
BUY_STRIPE_PREFIX = "https://buy.stripe.com/"
PRODUCT_PLACEHOLDER = "__PRODUCT_ID_FROM_LIVE_PRODUCT__"
PRICE_PLACEHOLDER = "__PRICE_ID_FROM_LIVE_PRICE__"


def now_jst() -> str:
    return datetime.now(JST).isoformat(timespec="seconds")


def safe_id(product_id: str) -> str:
    return re.sub(r"[^A-Za-z0-9_-]+", "_", product_id).strip("_") or "item"


def artifact_dir(product_id: str) -> Path:
    return ARTIFACT_ROOT / safe_id(product_id)


def plan_path(product_id: str) -> Path:
    return artifact_dir(product_id) / "commerce_plan.json"


def state_path(product_id: str) -> Path:
    return artifact_dir(product_id) / "commerce_state.json"


def result_path(product_id: str) -> Path:
    return artifact_dir(product_id) / "commerce_result.json"


def import_stripe_module() -> Any:
    try:
        import stripe  # type: ignore[import-not-found]
    except ImportError as exc:
        raise SafetyStop("stripe Python SDK is not installed. Install with: pip install stripe") from exc
    return stripe


def safe_error_message(exc: BaseException) -> str:
    message = str(exc)
    if contains_secret_like_value(message):
        return "redacted secret-like value"
    return message[:500]


def assert_artifact_safe(data: Any) -> None:
    if contains_secret_like_value(data):
        raise SafetyStop("secret-like value refused in Commerce artifact")


def write_safe_json(path: Path, data: dict[str, Any]) -> None:
    assert_artifact_safe(data)
    write_json_atomic(path, data)


def build_dry_run_plan(product_id: str) -> tuple[dict[str, Any], dict[str, Any]]:
    """Build the normal plan without reading secrets, importing Stripe, or calling an API."""
    payload = build_release_payload(product_id)
    errors = validate_dry_run_payload(payload)
    plan = {
        "mode": "dry-run",
        "created_at": now_jst(),
        "product_id": payload["product_id"],
        "source_original": payload["source_original"],
        "title": payload["title"],
        "amount": payload["price"],
        "currency": payload["currency"],
        "ready": not errors,
        "secret_read": "no",
        "live_api_called": "no",
        "operations": {
            "product": "create_if_missing_else_reuse_exact",
            "price": "create_if_no_exact_match_else_reuse_exact",
            "payment_link": "create_if_no_exact_match_else_reuse_exact",
            "original": "update_only_after_verified_read_back",
        },
        "idempotency_keys": {
            "product": payload["product_idempotency_key"],
            "price": payload["price_idempotency_key"],
            "payment_link": payload["payment_link_idempotency_key"],
        },
        "errors": errors,
    }
    write_safe_json(plan_path(product_id), plan)
    return payload, plan


def metadata_contains(actual: Any, expected: Any) -> bool:
    if not isinstance(actual, dict) or not isinstance(expected, dict):
        return False
    return all(str(actual.get(key, "")) == str(value) for key, value in expected.items())


def product_mismatches(product: dict[str, Any], expected: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    for key in ("name", "description", "tax_code", "active"):
        if product.get(key) != expected.get(key):
            errors.append(f"Product {key} mismatch")
    if not metadata_contains(product.get("metadata"), expected.get("metadata")):
        errors.append("Product metadata mismatch")
    if product.get("livemode") is not True:
        errors.append("Product livemode is not true")
    return errors


def select_product(products: list[dict[str, Any]], product_id: str, expected: dict[str, Any]) -> tuple[dict[str, Any] | None, str]:
    matches = [
        item
        for item in products
        if isinstance(item.get("metadata"), dict)
        and item["metadata"].get("dake_item_id") == product_id
    ]
    if not matches:
        return None, "created"
    if len(matches) > 1:
        raise SafetyStop("ambiguous Stripe Products: multiple metadata.dake_item_id matches")
    errors = product_mismatches(matches[0], expected)
    if errors:
        raise SafetyStop("existing Stripe Product does not exactly match: " + "; ".join(errors))
    return matches[0], "reused"


def price_mismatches(price: dict[str, Any], product_id: str, expected: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    expected_fields = {
        "product": product_id,
        "currency": expected.get("currency"),
        "unit_amount": expected.get("unit_amount"),
    }
    for key, value in expected_fields.items():
        actual = price.get(key)
        if isinstance(actual, dict):
            actual = actual.get("id")
        if actual != value:
            errors.append(f"Price {key} mismatch")
    if price.get("type", "one_time") != "one_time" or price.get("recurring") not in (None, {}):
        errors.append("Price is not one_time")
    if price.get("active") is not True:
        errors.append("Price active is not true")
    if not metadata_contains(price.get("metadata"), expected.get("metadata")):
        errors.append("Price metadata mismatch")
    if price.get("livemode") is not True:
        errors.append("Price livemode is not true")
    return errors


def select_price(prices: list[dict[str, Any]], product_id: str, expected: dict[str, Any]) -> tuple[dict[str, Any] | None, str]:
    exact = [price for price in prices if not price_mismatches(price, product_id, expected)]
    if not exact:
        return None, "created"
    if len(exact) > 1:
        raise SafetyStop("ambiguous Stripe Prices: multiple exact matches")
    return exact[0], "reused"


def line_item_values(line_item: dict[str, Any]) -> tuple[str, int]:
    price = line_item.get("price")
    price_id = price.get("id") if isinstance(price, dict) else price
    return str(price_id or ""), int(line_item.get("quantity") or 0)


def payment_link_mismatches(
    link: dict[str, Any],
    line_items: list[dict[str, Any]],
    price_id: str,
    expected: dict[str, Any],
) -> list[str]:
    errors: list[str] = []
    if link.get("active") is not True:
        errors.append("Payment Link active is not true")
    if link.get("livemode") is not True:
        errors.append("Payment Link livemode is not true")
    if not metadata_contains(link.get("metadata"), expected.get("metadata")):
        errors.append("Payment Link metadata mismatch")
    expected_lines = expected.get("line_items") or []
    expected_quantity = int(expected_lines[0].get("quantity") or 0) if len(expected_lines) == 1 else 0
    if len(line_items) != 1 or line_item_values(line_items[0]) != (price_id, expected_quantity):
        errors.append("Payment Link line item mismatch")
    for key in ("custom_text", "after_completion"):
        if key in expected and link.get(key) != expected.get(key):
            errors.append(f"Payment Link {key} mismatch")
    return errors


def select_payment_link(
    links: list[tuple[dict[str, Any], list[dict[str, Any]]]],
    price_id: str,
    expected: dict[str, Any],
) -> tuple[dict[str, Any] | None, str]:
    exact = [link for link, lines in links if not payment_link_mismatches(link, lines, price_id, expected)]
    if not exact:
        return None, "created"
    if len(exact) > 1:
        raise SafetyStop("ambiguous Stripe Payment Links: multiple exact matches")
    return exact[0], "reused"


def iter_list(result: Any) -> list[dict[str, Any]]:
    iterator = result.auto_paging_iter() if hasattr(result, "auto_paging_iter") else result
    return [object_to_dict(item) for item in iterator]


def list_products(stripe: Any) -> list[dict[str, Any]]:
    return iter_list(stripe.Product.list(limit=100))


def list_prices(stripe: Any, product_id: str) -> list[dict[str, Any]]:
    return iter_list(stripe.Price.list(product=product_id, active=True, limit=100))


def payment_link_with_lines(stripe: Any, link: dict[str, Any]) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    return link, iter_list(stripe.PaymentLink.list_line_items(str(link.get("id")), limit=100))


def list_payment_links(stripe: Any) -> list[tuple[dict[str, Any], list[dict[str, Any]]]]:
    links = iter_list(stripe.PaymentLink.list(active=True, limit=100))
    return [payment_link_with_lines(stripe, link) for link in links]


def retrieve_and_verify(
    retrieve: Callable[..., Any],
    object_id: str,
    validator: Callable[[dict[str, Any]], list[str]],
    label: str,
) -> dict[str, Any]:
    value = object_to_dict(retrieve(object_id))
    errors = validator(value)
    if errors:
        raise SafetyStop(f"{label} read-back mismatch: " + "; ".join(errors))
    return value


def replace_or_append_metadata(text: str, key: str, value: str) -> str:
    pattern = re.compile(rf"^(\s*[-*]\s*{re.escape(key)}\s*:\s*)(.*?)\s*$", re.MULTILINE)
    matches = list(pattern.finditer(text))
    if len(matches) > 1:
        raise SafetyStop(f"ambiguous ORIGINAL.md: multiple {key} lines")
    if matches:
        return pattern.sub(rf"\g<1>{value}", text, count=1)
    separator = "" if text.endswith("\n") else "\n"
    heading = "\n## Commerce\n\n" if "## Commerce" not in text else "\n"
    return f"{text}{separator}{heading}- {key}: {value}\n"


def metadata_line(text: str, key: str) -> str:
    pattern = re.compile(rf"^\s*[-*]\s*{re.escape(key)}\s*:\s*(.+?)\s*$", re.MULTILINE)
    match = pattern.search(text)
    if not match:
        return ""
    value = match.group(1).strip()
    return value[1:-1].strip() if len(value) >= 2 and value.startswith("`") and value.endswith("`") else value


def update_original_atomic(source_original: str, payment_link_url: str) -> str:
    if not payment_link_url.startswith(BUY_STRIPE_PREFIX) or "test_" in payment_link_url:
        raise SafetyStop("verified Payment Link URL is not a live buy.stripe.com URL")
    path = (ROOT / source_original).resolve()
    try:
        path.relative_to(ROOT.resolve())
    except ValueError as exc:
        raise SafetyStop("source ORIGINAL.md is outside the repository") from exc
    original = path.read_text(encoding="utf-8")
    current_status = metadata_line(original, "payment_status")
    current_link = metadata_line(original, "stripe_payment_link")
    unset = {"", "not set", "none", "null", "-", "unset", "未設定"}
    if current_link.lower() not in unset and current_link != payment_link_url:
        raise SafetyStop("ORIGINAL.md already contains a different Stripe Payment Link")
    if current_status not in {"", "booth_only", "preparing", "stripe_ready"}:
        raise SafetyStop(f"unsupported ORIGINAL.md payment_status: {current_status}")
    if current_status == "stripe_ready" and current_link == payment_link_url:
        return "unchanged"
    updated = replace_or_append_metadata(original, "stripe_payment_link", payment_link_url)
    updated = replace_or_append_metadata(updated, "payment_status", "stripe_ready")
    try:
        write_text_atomic(path, updated)
        verified = path.read_text(encoding="utf-8")
        if metadata_line(verified, "payment_status") != "stripe_ready":
            raise SafetyStop("ORIGINAL.md payment_status write-back verification failed")
        if metadata_line(verified, "stripe_payment_link") != payment_link_url:
            raise SafetyStop("ORIGINAL.md Payment Link write-back verification failed")
    except Exception:
        write_text_atomic(path, original)
        raise
    return "updated"


def save_state(product_id: str, state: dict[str, Any], status: str) -> None:
    state["status"] = status
    state["updated_at"] = now_jst()
    write_safe_json(state_path(product_id), state)


def execute_live(payload: dict[str, Any], *, stripe_client: Any | None = None, secret_key: str | None = None) -> dict[str, Any]:
    errors = validate_dry_run_payload(payload)
    if errors:
        raise SafetyStop("dry-run plan is not live-ready: " + "; ".join(errors))
    key = validate_live_secret_key() if secret_key is None else secret_key.strip()
    if key.startswith("sk_test_"):
        raise SafetyStop("Refusing to run: test mode secret keys are not allowed for live execution")
    if not key.startswith("sk_live_"):
        raise SafetyStop("Refusing to run: STRIPE_SECRET_KEY must start with sk_live_")
    stripe = stripe_client or import_stripe_module()
    stripe.api_key = key

    product_id = payload["product_id"]
    state: dict[str, Any] = {
        "mode": "live",
        "product_id": product_id,
        "source_original": payload["source_original"],
        "started_at": now_jst(),
        "status": "reconciling",
        "actions": {},
        "product_id_on_stripe": None,
        "price_id": None,
        "payment_link_id": None,
        "payment_link_url": None,
        "error": None,
    }
    save_state(product_id, state, "reconciling")
    step = "product"
    try:
        product, product_action = select_product(list_products(stripe), product_id, payload["product_payload"])
        if product is None:
            product = create_with_idempotency(stripe.Product.create, payload["product_payload"], payload["product_idempotency_key"])
        product_id_on_stripe = str(product.get("id") or "")
        if not product_id_on_stripe:
            raise SafetyStop("Stripe Product id is missing")
        product = retrieve_and_verify(
            stripe.Product.retrieve,
            product_id_on_stripe,
            lambda item: product_mismatches(item, payload["product_payload"]),
            "Product",
        )
        state["actions"]["product"] = product_action
        state["product_id_on_stripe"] = product_id_on_stripe
        save_state(product_id, state, "product_verified")

        step = "price"
        expected_price = copy.deepcopy(payload["price_payload"])
        expected_price["product"] = product_id_on_stripe
        price, price_action = select_price(list_prices(stripe, product_id_on_stripe), product_id_on_stripe, expected_price)
        if price is None:
            price = create_with_idempotency(stripe.Price.create, expected_price, payload["price_idempotency_key"])
        price_id = str(price.get("id") or "")
        if not price_id:
            raise SafetyStop("Stripe Price id is missing")
        price = retrieve_and_verify(
            stripe.Price.retrieve,
            price_id,
            lambda item: price_mismatches(item, product_id_on_stripe, expected_price),
            "Price",
        )
        state["actions"]["price"] = price_action
        state["price_id"] = price_id
        save_state(product_id, state, "price_verified")

        step = "payment_link"
        expected_link = copy.deepcopy(payload["payment_link_payload"])
        expected_link["line_items"] = [{**line, "price": price_id} for line in expected_link.get("line_items", [])]
        link, link_action = select_payment_link(list_payment_links(stripe), price_id, expected_link)
        if link is None:
            link = create_with_idempotency(stripe.PaymentLink.create, expected_link, payload["payment_link_idempotency_key"])
        link_id = str(link.get("id") or "")
        if not link_id:
            raise SafetyStop("Stripe Payment Link id is missing")
        link = object_to_dict(stripe.PaymentLink.retrieve(link_id))
        link_lines = iter_list(stripe.PaymentLink.list_line_items(link_id, limit=100))
        link_errors = payment_link_mismatches(link, link_lines, price_id, expected_link)
        if link_errors:
            raise SafetyStop("Payment Link read-back mismatch: " + "; ".join(link_errors))
        payment_link_url = str(link.get("url") or "")
        if not payment_link_url.startswith(BUY_STRIPE_PREFIX) or "test_" in payment_link_url:
            raise SafetyStop("Payment Link read-back returned a non-live URL")
        state["actions"]["payment_link"] = link_action
        state["payment_link_id"] = link_id
        state["payment_link_url"] = payment_link_url
        save_state(product_id, state, "stripe_verified")

        step = "original"
        original_action = update_original_atomic(payload["source_original"], payment_link_url)
        state["actions"]["original"] = original_action
        save_state(product_id, state, "completed")
        result = {
            "mode": "live",
            "status": "completed",
            "completed_at": now_jst(),
            "product_id": product_id,
            "title": payload["title"],
            "amount": payload["price"],
            "currency": payload["currency"],
            "product_id_on_stripe": product_id_on_stripe,
            "price_id": price_id,
            "payment_link_id": link_id,
            "payment_link_url": payment_link_url,
            "actions": state["actions"],
            "original_verified": True,
            "commerce_status": "READY",
            "errors": [],
        }
        write_safe_json(result_path(product_id), result)
        return result
    except Exception as exc:
        state["error"] = {"step": step, "message": safe_error_message(exc)}
        save_state(product_id, state, "stopped")
        if isinstance(exc, SafetyStop):
            raise
        raise SafetyStop(safe_error_message(exc)) from exc
