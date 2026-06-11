from __future__ import annotations

import argparse
import json
import os
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any


SAFETY_NOTICE = "TEST MODE ONLY. Default mode is dry-run and does not call Stripe API."
ROOT = Path(__file__).resolve().parents[2]
DEFAULT_PAYLOAD_JSON = ROOT / "tools" / "reports" / "stripe_payment_link_pilot10_payloads.json"
DEFAULT_RESULT_JSON = ROOT / "tools" / "reports" / "stripe_payment_link_pilot10_test_result.json"
DEFAULT_RESULT_MD = ROOT / "tools" / "reports" / "stripe_payment_link_pilot10_test_result.md"
JST = timezone(timedelta(hours=9))


def load_payload(path: Path) -> dict[str, Any]:
    data = json.loads(path.read_text(encoding="utf-8"))
    if data.get("dry_run") is not True:
        raise ValueError("payload JSON must be a dry-run payload")
    if data.get("errors"):
        raise ValueError("payload JSON contains errors; refuse to execute")
    items = data.get("items")
    if not isinstance(items, list) or len(items) != 10:
        raise ValueError("payload JSON must contain exactly 10 items")
    return data


def validate_secret_key() -> str:
    value = os.environ.get("STRIPE_SECRET_KEY", "").strip()
    if not value:
        raise RuntimeError("STRIPE_SECRET_KEY is not set. Set a Stripe test mode secret key in the environment.")
    if value.startswith("sk_live_"):
        raise RuntimeError("Refusing to run: live mode secret keys are not allowed.")
    if not value.startswith("sk_test_"):
        raise RuntimeError("Refusing to run: STRIPE_SECRET_KEY must start with sk_test_.")
    return value


def print_dry_run(payload: dict[str, Any]) -> None:
    print(SAFETY_NOTICE)
    print("No Stripe API call is made.")
    print("No Stripe Secret Key is read.")
    print(f"items={len(payload['items'])}")
    print("This script creates new Stripe test objects only when --execute-test is passed.")
    for item in payload["items"]:
        product = item["product_payload"]
        price = item["price_payload"]
        metadata = product.get("metadata", {})
        print(
            f"- {item['id']}: name={product['name']} "
            f"amount={price['unit_amount']} {price['currency']} "
            f"tax_code={product.get('tax_code')} "
            f"metadata_keys={','.join(metadata.keys())}"
        )


def import_stripe_module() -> Any:
    try:
        import stripe  # type: ignore[import-not-found]
    except ImportError as exc:
        raise RuntimeError("stripe Python SDK is not installed. Install with: pip install stripe") from exc
    return stripe


def create_one(stripe: Any, item: dict[str, Any]) -> dict[str, Any]:
    product_payload = dict(item["product_payload"])
    price_payload = dict(item["price_payload"])
    payment_link_payload = dict(item["payment_link_payload"])

    print(f"creating test objects for {item['id']}")
    product = stripe.Product.create(**product_payload)
    price_payload["product"] = product.id
    price = stripe.Price.create(**price_payload)
    payment_link_payload["line_items"] = [
        {
            "price": price.id,
            "quantity": line_item.get("quantity", 1),
        }
        for line_item in payment_link_payload.get("line_items", [])
    ]
    payment_link = stripe.PaymentLink.create(**payment_link_payload)

    metadata = dict(product_payload.get("metadata", {}))
    return {
        "id": item["id"],
        "title": item["title"],
        "product_id": product.id,
        "price_id": price.id,
        "payment_link_id": payment_link.id,
        "payment_link_url": payment_link.url,
        "metadata": metadata,
        "livemode": bool(getattr(payment_link, "livemode", False)),
    }


def write_result_json(path: Path, result: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def md(value: object) -> str:
    text = "" if value is None else str(value)
    text = text.replace("|", "\\|").replace("\r\n", "\n").replace("\r", "\n")
    return "<br>".join(part.strip() for part in text.split("\n") if part.strip()) or "-"


def table(headers: list[str], rows: list[list[object]]) -> str:
    lines = [
        "| " + " | ".join(headers) + " |",
        "| " + " | ".join("---" for _ in headers) + " |",
    ]
    if not rows:
        rows = [["-" for _ in headers]]
    for row in rows:
        lines.append("| " + " | ".join(md(cell) for cell in row) + " |")
    return "\n".join(lines)


def write_result_markdown(path: Path, result: dict[str, Any]) -> None:
    item_rows = [
        [
            item["id"],
            item["title"],
            item["product_id"],
            item["price_id"],
            item["payment_link_id"],
            item["payment_link_url"],
        ]
        for item in result["items"]
    ]
    error_rows = [[error] for error in result["errors"]]
    content = f"""# Stripe Payment Link Pilot10 Test Result

## Summary

- mode: test
- count: {result['count']}
- errors: {len(result['errors'])}
- live mode used: no
- Secret saved: no

## Items

{table(['id', 'title', 'product_id', 'price_id', 'payment_link_id', 'payment_link_url'], item_rows)}

## Errors

{table(['error'], error_rows)}

## Safety Notes

- Stripe Secret Key was read only from the `STRIPE_SECRET_KEY` environment variable.
- The secret value is not written to this file.
- Live mode keys are rejected.
- `ORIGINAL.md`, generated JSON, dake-store-site, and Store production were not updated.
"""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8", newline="\n")


def execute_test(payload: dict[str, Any], result_json: Path, result_md: Path) -> int:
    secret_key = validate_secret_key()
    stripe = import_stripe_module()
    stripe.api_key = secret_key

    print("Executing Stripe test mode creation.")
    print("This script creates new test Product, Price, and Payment Link objects.")
    print("Secret value is not printed or stored.")

    items: list[dict[str, Any]] = []
    errors: list[str] = []
    for item in payload["items"]:
        try:
            created = create_one(stripe, item)
            if created.get("livemode"):
                errors.append(f"{item['id']}: Stripe returned livemode=true; refusing to treat this as success")
            items.append(created)
        except Exception as exc:  # Stripe SDK raises several concrete exception types.
            errors.append(f"{item['id']}: {exc}")

    result = {
        "mode": "test",
        "created_at": datetime.now(JST).isoformat(timespec="seconds"),
        "count": len(items),
        "items": items,
        "errors": errors,
    }
    write_result_json(result_json, result)
    write_result_markdown(result_md, result)
    print(f"created_items={len(items)}")
    print(f"errors={len(errors)}")
    print(f"wrote {result_json}")
    print(f"wrote {result_md}")
    return 0 if not errors and len(items) == 10 else 1


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=SAFETY_NOTICE)
    parser.add_argument("--payload-json", type=Path, default=DEFAULT_PAYLOAD_JSON)
    parser.add_argument("--result-json", type=Path, default=DEFAULT_RESULT_JSON)
    parser.add_argument("--result-md", type=Path, default=DEFAULT_RESULT_MD)
    parser.add_argument("--execute-test", action="store_true", help="Create Stripe test mode objects. Requires STRIPE_SECRET_KEY=sk_test_...")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    payload = load_payload(args.payload_json)
    if not args.execute_test:
        print_dry_run(payload)
        return 0
    return execute_test(payload, args.result_json, args.result_md)


if __name__ == "__main__":
    raise SystemExit(main())
