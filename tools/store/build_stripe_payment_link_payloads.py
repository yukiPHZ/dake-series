from __future__ import annotations

import argparse
import csv
import json
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any


DRY_RUN_NOTICE = "DRY RUN ONLY. This file does not call Stripe API."
STORE_PRODUCT_BASE_URL = "https://store.dakeapp.com/product/"
JST = timezone(timedelta(hours=9))

ROOT = Path(__file__).resolve().parents[2]
DEFAULT_INPUT_CSV = ROOT / "tools" / "reports" / "stripe_payment_link_pilot10_selection.csv"
DEFAULT_GENERATED_JSON = ROOT / "tools" / "generated" / "store_products.generated.json"
DEFAULT_OUTPUT_JSON = ROOT / "tools" / "reports" / "stripe_payment_link_pilot10_payloads.json"
DEFAULT_OUTPUT_MD = ROOT / "tools" / "reports" / "stripe_payment_link_pilot10_payloads.md"

PRODUCT_ID_PLACEHOLDER = "__PRODUCT_ID_FROM_CREATED_PRODUCT__"
PRICE_ID_PLACEHOLDER = "__PRICE_ID_FROM_CREATED_PRICE__"
EXPECTED_COUNT = 10


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as file:
        return list(csv.DictReader(file))


def read_generated_items(path: Path) -> dict[str, dict[str, Any]]:
    with path.open("r", encoding="utf-8") as file:
        data = json.load(file)
    items = data.get("items")
    if not isinstance(items, list):
        raise ValueError("generated Store JSON must contain an items array")
    return {str(item.get("id")): item for item in items if isinstance(item, dict) and item.get("id")}


def parse_price(value: str) -> int | None:
    text = str(value or "").replace(",", "").strip()
    if not text:
        return None
    try:
        return int(text)
    except ValueError:
        return None


def clean_description(value: object, fallback: str) -> str:
    text = str(value or "").strip()
    if not text:
        text = fallback.strip()
    text = " ".join(text.replace("\r\n", "\n").replace("\r", "\n").split())
    if len(text) > 240:
        return text[:237].rstrip() + "..."
    return text


def store_url(item_id: str) -> str:
    return f"{STORE_PRODUCT_BASE_URL}?id={item_id}"


def metadata(row: dict[str, str], generated_item: dict[str, Any]) -> dict[str, str]:
    values = {
        "dake_item_id": row["id"],
        "dake_type": row["type"],
        "source_repo": str(generated_item.get("source_repo") or ""),
        "source_original": row["source_original"],
        "store_url": store_url(row["id"]),
        "booth_url": row["booth_url"],
        "github_release_url": row["github_release_url"],
    }
    return {key: value for key, value in values.items() if value != ""}


def validate_row(index: int, row: dict[str, str], generated_items: dict[str, dict[str, Any]]) -> list[str]:
    label = row.get("id") or f"row {index + 1}"
    errors: list[str] = []
    if not row.get("id"):
        errors.append(f"{label}: id is missing")
    if not row.get("title"):
        errors.append(f"{label}: title is missing")
    price = parse_price(row.get("price", ""))
    if price is None:
        errors.append(f"{label}: price is not an integer")
    if row.get("currency", "").strip().lower() != "jpy":
        errors.append(f"{label}: currency must be JPY")
    if not row.get("tax_code_candidate"):
        errors.append(f"{label}: tax_code_candidate is missing")
    if row.get("metadata_ready") != "yes":
        errors.append(f"{label}: metadata_ready must be yes")
    if row.get("review_result") != "create":
        errors.append(f"{label}: review_result must be create")
    if row.get("creation_method") != "api_candidate":
        errors.append(f"{label}: creation_method must be api_candidate")
    if row.get("price_check") != "price_ok":
        errors.append(f"{label}: price_check must be price_ok")
    if row.get("id") and row["id"] not in generated_items:
        errors.append(f"{label}: id is missing from generated Store JSON")
    return errors


def build_item(row: dict[str, str], generated_item: dict[str, Any]) -> dict[str, Any]:
    item_id = row["id"]
    price = parse_price(row["price"])
    if price is None:
        raise ValueError(f"price is invalid for {item_id}")
    product_name = row.get("stripe_product_name") or row["title"]
    product_description = clean_description(
        generated_item.get("description"),
        f"{row['title']} / {row.get('category', '')}".strip(" /"),
    )
    full_metadata = metadata(row, generated_item)
    core_metadata = {
        "dake_item_id": full_metadata["dake_item_id"],
        "dake_type": full_metadata["dake_type"],
        "source_original": full_metadata["source_original"],
    }

    product_payload = {
        "name": product_name,
        "description": product_description,
        "metadata": full_metadata,
        "tax_code": row["tax_code_candidate"],
    }
    price_payload = {
        "currency": row["currency"].lower(),
        "unit_amount": price,
        "product": PRODUCT_ID_PLACEHOLDER,
        "metadata": {"dake_item_id": item_id},
    }
    payment_link_payload = {
        "line_items": [
            {
                "price": PRICE_ID_PLACEHOLDER,
                "quantity": 1,
            }
        ],
        "metadata": full_metadata,
        "payment_intent_data": {
            "metadata": core_metadata,
        },
    }
    return {
        "id": item_id,
        "title": row["title"],
        "product_payload": product_payload,
        "price_payload": price_payload,
        "payment_link_payload": payment_link_payload,
    }


def build_payload(rows: list[dict[str, str]], generated_items: dict[str, dict[str, Any]]) -> dict[str, Any]:
    errors: list[str] = []
    if len(rows) != EXPECTED_COUNT:
        errors.append(f"expected {EXPECTED_COUNT} input rows, got {len(rows)}")
    for index, row in enumerate(rows):
        errors.extend(validate_row(index, row, generated_items))

    items: list[dict[str, Any]] = []
    if not errors:
        for row in rows:
            items.append(build_item(row, generated_items[row["id"]]))

    return {
        "dry_run": True,
        "notice": DRY_RUN_NOTICE,
        "source": "tools/reports/stripe_payment_link_pilot10_selection.csv",
        "created_at": datetime.now(JST).isoformat(timespec="seconds"),
        "count": len(items),
        "items": items,
        "errors": errors,
        "safety": [
            "No Stripe API call is made.",
            "No Stripe Secret Key is read.",
            "No Payment Link is created.",
            "No Product is created.",
            "No Price is created.",
        ],
    }


def write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def md(value: object) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("|", "\\|")
    text = "<br>".join(part.strip() for part in text.split("\n") if part.strip())
    return text or "-"


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


def metadata_summary(item: dict[str, Any]) -> str:
    keys = item["product_payload"]["metadata"].keys()
    return ", ".join(keys)


def write_markdown(path: Path, payload: dict[str, Any], output_json: Path) -> None:
    item_rows = [
        [
            item["id"],
            item["title"],
            item["price_payload"]["unit_amount"],
            item["price_payload"]["currency"],
            item["product_payload"]["tax_code"],
            metadata_summary(item),
        ]
        for item in payload["items"]
    ]
    error_rows = [[error] for error in payload["errors"]]
    content = f"""# Stripe Payment Link Pilot10 Payloads

## Purpose

Build dry-run payloads for Stripe Product / Price / Payment Link creation for the Phase 13A pilot 10 items.

## Dry-run Notice

{DRY_RUN_NOTICE}

No Stripe API call is made. No Stripe Secret Key is read. No Payment Link is created. No Product is created. No Price is created.

## Input

- `tools/reports/stripe_payment_link_pilot10_selection.csv`
- `tools/generated/store_products.generated.json`

## Output

- `{output_json.relative_to(ROOT).as_posix()}`
- `{path.relative_to(ROOT).as_posix()}`

## Count

- items: {payload['count']}
- errors: {len(payload['errors'])}

## Items

{table(['id', 'title', 'price', 'currency', 'tax_code', 'metadata'], item_rows)}

## Errors

{table(['error'], error_rows)}

## Safety Notes

- No Stripe API call is made.
- No Stripe Secret Key is read.
- No Payment Link is created.
- No Product is created.
- No Price is created.
- Metadata must not contain personal information, card information, buyer information, Stripe Secret Key, Webhook Secret, or internal tokens.
- `tax_code` is a candidate and should be reviewed before live execution.
- Stripe Payment Links are Stripe-hosted checkout URLs. The payload here only describes what would be sent in a later phase.
- Metadata stores structured key-value information on Stripe objects. Payment Link `payment_intent_data.metadata` is included so generated Payment Intents can carry DAKE identifiers.

## Next Phase

Use this JSON as the reviewed input for a test mode implementation phase. That next phase should still default to dry-run, require explicit human approval before any execution path, and keep Stripe Secret Key in environment variables only.
"""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8", newline="\n")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=DRY_RUN_NOTICE)
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT_CSV)
    parser.add_argument("--generated-json", type=Path, default=DEFAULT_GENERATED_JSON)
    parser.add_argument("--output-json", type=Path, default=DEFAULT_OUTPUT_JSON)
    parser.add_argument("--output-md", type=Path, default=DEFAULT_OUTPUT_MD)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    rows = read_csv(args.input)
    generated_items = read_generated_items(args.generated_json)
    payload = build_payload(rows, generated_items)
    write_json(args.output_json, payload)
    write_markdown(args.output_md, payload, args.output_json)
    print(DRY_RUN_NOTICE)
    print(f"input_rows={len(rows)}")
    print(f"payload_items={payload['count']}")
    print(f"errors={len(payload['errors'])}")
    print(f"wrote {args.output_json}")
    print(f"wrote {args.output_md}")
    return 0 if not payload["errors"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
