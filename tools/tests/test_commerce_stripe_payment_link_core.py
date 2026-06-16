from __future__ import annotations

import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

from tools.commerce.stripe_payment_link_core import (  # noqa: E402
    CommerceProductSpec,
    CommerceSafetyStop,
    build_stripe_payment_link_payloads,
    idempotency_key,
)


def spec(**overrides: object) -> CommerceProductSpec:
    values = {
        "product_id": "generic_item",
        "product_kind": "digital_product_one_time",
        "title": "Generic Digital Product",
        "description": "One-time digital product.",
        "amount": 500,
        "currency": "jpy",
        "price_model": "one_time",
        "tax_code": "txcd_10202003",
        "quantity": 1,
        "metadata": {
            "dake_item_id": "generic_item",
            "dake_type": "digital_product",
            "source_repo": "fixture",
            "source_original": "fixtures/ORIGINAL.md",
            "store_url": "https://store.dakeapp.com/product/?id=generic_item",
        },
        "checkout_notice": "",
        "fulfillment_mode": "",
        "after_completion_url": None,
    }
    values.update(overrides)
    return CommerceProductSpec(**values)  # type: ignore[arg-type]


def must_stop(expected: str, **overrides: object) -> None:
    try:
        build_stripe_payment_link_payloads(spec(**overrides))
    except CommerceSafetyStop as exc:
        assert expected in str(exc), str(exc)
    else:
        raise AssertionError("expected CommerceSafetyStop")


def test_supported_product_modes() -> None:
    cases = [
        ("app_one_time", {}),
        ("pack_one_time_manual_delivery", {"fulfillment_mode": "manual_email_private_download"}),
        ("digital_product_one_time", {}),
        ("web_service_one_time_manual_fulfillment", {"fulfillment_mode": "manual"}),
        (
            "web_service_one_time_redirect",
            {"after_completion_url": "https://example.com/complete"},
        ),
    ]
    for product_kind, overrides in cases:
        payloads = build_stripe_payment_link_payloads(spec(product_kind=product_kind, **overrides))
        assert payloads.product_payload["name"]
        assert payloads.price_payload["unit_amount"] == 500
        assert payloads.payment_link_payload["line_items"][0]["quantity"] == 1
        assert "sk_live" not in str(payloads)
        assert "whsec" not in str(payloads)
    redirected = build_stripe_payment_link_payloads(
        spec(product_kind="web_service_one_time_redirect", after_completion_url="https://example.com/complete")
    )
    assert redirected.payment_link_payload["after_completion"]["type"] == "redirect"


def test_rejections() -> None:
    must_stop("unsupported_in_commerce_v1", price_model="recurring")
    must_stop("amount must be", amount=0)
    must_stop("unsupported currency", currency="usd")
    must_stop("tax_code is required", tax_code="")
    must_stop("checkout notice length overflow", checkout_notice="x" * 1201)
    must_stop("secret-like metadata", metadata={"dake_item_id": "x", "key": "sk_live_1234567890abcdef"})
    must_stop(
        "private URL exposure",
        product_kind="web_service_one_time_redirect",
        after_completion_url="http://localhost/private",
    )
    must_stop("after_completion_url is required", product_kind="web_service_one_time_redirect")


def test_idempotency_stability() -> None:
    payloads = build_stripe_payment_link_payloads(spec())
    first = idempotency_key("product", "generic_item", payloads.product_payload)
    second = idempotency_key("product", "generic_item", payloads.product_payload)
    assert first == second
    assert first.startswith("commerce-product-v1-generic_item-")


def main() -> int:
    tests = [
        test_supported_product_modes,
        test_rejections,
        test_idempotency_stability,
    ]
    for test in tests:
        test()
        print(f"{test.__name__}=passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
