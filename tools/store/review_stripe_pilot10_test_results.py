from __future__ import annotations

import csv
import json
import os
import re
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[2]
DEFAULT_TEST_RESULT_JSON = ROOT / "tools" / "reports" / "stripe_payment_link_pilot10_test_result.json"
DEFAULT_PAYLOAD_JSON = ROOT / "tools" / "reports" / "stripe_payment_link_pilot10_payloads.json"
DEFAULT_SELECTION_CSV = ROOT / "tools" / "reports" / "stripe_payment_link_pilot10_selection.csv"
DEFAULT_REVIEW_CSV = ROOT / "tools" / "reports" / "stripe_payment_link_pilot10_test_review.csv"
DEFAULT_REVIEW_MD = ROOT / "tools" / "reports" / "stripe_payment_link_pilot10_test_review.md"
JST = timezone(timedelta(hours=9))

EXPECTED_TAX_CODE = "txcd_10202003"
REQUIRED_FULL_METADATA_KEYS = [
    "dake_item_id",
    "dake_type",
    "source_repo",
    "source_original",
    "store_url",
    "booth_url",
    "github_release_url",
]
SECRET_PATTERN = re.compile(r"(sk_(?:test|live)_[A-Za-z0-9]{16,}|whsec_[A-Za-z0-9]{16,})")
CSV_COLUMNS = [
    "id",
    "expected_title",
    "actual_product_name",
    "description_check",
    "expected_price",
    "actual_unit_amount",
    "currency",
    "price_type",
    "recurring_check",
    "product_match",
    "product_active",
    "price_active",
    "payment_link_active",
    "livemode_check",
    "payment_link_url",
    "metadata_check",
    "tax_code_actual",
    "tax_code_check",
    "browser_check",
    "technical_ready",
    "tax_business_review",
    "live_ready",
    "notes",
]


def read_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def read_selection(path: Path) -> dict[str, dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return {row["id"]: row for row in csv.DictReader(handle)}


def object_to_dict(obj: Any) -> Any:
    """Convert Stripe SDK objects and nested containers to plain Python values."""
    to_dict = getattr(obj, "to_dict", None)
    if callable(to_dict):
        return object_to_dict(to_dict())

    legacy_to_dict_recursive = getattr(obj, "to_dict_recursive", None)
    if callable(legacy_to_dict_recursive):
        return object_to_dict(legacy_to_dict_recursive())

    if isinstance(obj, dict):
        return {key: object_to_dict(value) for key, value in obj.items()}
    if isinstance(obj, list):
        return [object_to_dict(value) for value in obj]
    if isinstance(obj, tuple):
        return [object_to_dict(value) for value in obj]
    return obj


def validate_secret_key() -> str:
    value = os.environ.get("STRIPE_SECRET_KEY", "").strip()
    if not value:
        raise RuntimeError("STRIPE_SECRET_KEY is not set. Set a Stripe test mode secret key in the environment.")
    if value.startswith("sk_live_"):
        raise RuntimeError("Refusing to run: live mode secret keys are not allowed.")
    if not value.startswith("sk_test_"):
        raise RuntimeError("Refusing to run: STRIPE_SECRET_KEY must start with sk_test_.")
    return value


def import_stripe_module() -> Any:
    try:
        import stripe  # type: ignore[import-not-found]
    except ImportError as exc:
        raise RuntimeError("stripe Python SDK is not installed. Install with: pip install stripe") from exc
    return stripe


def validate_local_result(result: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    if result.get("mode") != "test":
        errors.append("local result mode is not test")
    if result.get("count") != 10:
        errors.append("local result count is not 10")
    if result.get("errors"):
        errors.append("local result contains errors")
    items = result.get("items")
    if not isinstance(items, list) or len(items) != 10:
        errors.append("local result must contain exactly 10 items")
        return errors
    for item in items:
        item_id = item.get("id", "<missing>")
        for key in ["product_id", "price_id", "payment_link_id", "payment_link_url"]:
            if not item.get(key):
                errors.append(f"{item_id}: local result missing {key}")
        if item.get("livemode") is not False:
            errors.append(f"{item_id}: local result livemode is not false")
    return errors


def index_payloads(payload: dict[str, Any]) -> dict[str, dict[str, Any]]:
    return {item["id"]: item for item in payload.get("items", [])}


def get_metadata(obj: dict[str, Any]) -> dict[str, str]:
    metadata = obj.get("metadata") or {}
    return {str(key): "" if value is None else str(value) for key, value in dict(metadata).items()}


def expected_full_metadata(result_item: dict[str, Any], payload_item: dict[str, Any] | None) -> dict[str, str]:
    if payload_item:
        product_meta = payload_item.get("product_payload", {}).get("metadata") or {}
        if product_meta:
            return {str(key): "" if value is None else str(value) for key, value in dict(product_meta).items()}
    return {str(key): "" if value is None else str(value) for key, value in dict(result_item.get("metadata") or {}).items()}


def metadata_matches(actual: dict[str, str], expected: dict[str, str]) -> bool:
    return all(actual.get(key) == value for key, value in expected.items())


def full_metadata_ok(actual: dict[str, str], expected: dict[str, str], item_id: str) -> bool:
    if not all(key in actual for key in REQUIRED_FULL_METADATA_KEYS):
        return False
    if actual.get("dake_item_id") != item_id:
        return False
    return metadata_matches(actual, expected)


def check_description(name: str, description: str) -> str:
    text = (description or "").strip()
    if not text:
        return "ng"
    if len(text) > 500:
        return "review"
    if SECRET_PATTERN.search(text):
        return "ng"
    if not name.strip():
        return "ng"
    return "ok"


def stringify(value: Any) -> str:
    if value is None:
        return ""
    return str(value)


def retrieve_one(
    stripe: Any,
    result_item: dict[str, Any],
    payload_item: dict[str, Any] | None,
    selection_item: dict[str, str] | None,
    browser_check: str,
) -> tuple[dict[str, str], dict[str, int]]:
    item_id = result_item["id"]
    expected_meta = expected_full_metadata(result_item, payload_item)
    expected_price = int(selection_item["price"]) if selection_item and selection_item.get("price") else None
    if expected_price is None and payload_item:
        expected_price = payload_item.get("price_payload", {}).get("unit_amount")
    expected_title = ""
    if selection_item:
        expected_title = selection_item.get("stripe_product_name") or selection_item.get("title") or ""
    if not expected_title and payload_item:
        expected_title = payload_item.get("product_payload", {}).get("name", "")
    if not expected_title:
        expected_title = result_item.get("title", "")

    product = object_to_dict(stripe.Product.retrieve(result_item["product_id"]))
    price = object_to_dict(stripe.Price.retrieve(result_item["price_id"]))
    payment_link = object_to_dict(stripe.PaymentLink.retrieve(result_item["payment_link_id"]))

    product_metadata = get_metadata(product)
    price_metadata = get_metadata(price)
    link_metadata = get_metadata(payment_link)
    payment_intent_data = payment_link.get("payment_intent_data") or {}
    payment_intent_metadata = {
        str(key): "" if value is None else str(value)
        for key, value in dict((payment_intent_data or {}).get("metadata") or {}).items()
    }

    expected_price_metadata = {}
    expected_intent_metadata = {}
    if payload_item:
        expected_price_metadata = {
            str(key): "" if value is None else str(value)
            for key, value in dict(payload_item.get("price_payload", {}).get("metadata") or {}).items()
        }
        expected_intent_metadata = {
            str(key): "" if value is None else str(value)
            for key, value in dict(
                payload_item.get("payment_link_payload", {}).get("payment_intent_data", {}).get("metadata") or {}
            ).items()
        }
    if not expected_price_metadata:
        expected_price_metadata = {"dake_item_id": item_id}
    if not expected_intent_metadata:
        expected_intent_metadata = {
            "dake_item_id": item_id,
            "dake_type": expected_meta.get("dake_type", ""),
            "source_original": expected_meta.get("source_original", ""),
        }

    product_metadata_ok = full_metadata_ok(product_metadata, expected_meta, item_id)
    price_metadata_ok = metadata_matches(price_metadata, expected_price_metadata)
    link_metadata_ok = full_metadata_ok(link_metadata, expected_meta, item_id)
    intent_metadata_ok = metadata_matches(payment_intent_metadata, expected_intent_metadata)
    if product_metadata_ok and price_metadata_ok and link_metadata_ok and intent_metadata_ok:
        metadata_check = "ok"
    elif any([product_metadata_ok, price_metadata_ok, link_metadata_ok, intent_metadata_ok]):
        metadata_check = "partial"
    else:
        metadata_check = "ng"

    product_active = bool(product.get("active"))
    price_active = bool(price.get("active"))
    link_active = bool(payment_link.get("active"))
    product_livemode = bool(product.get("livemode"))
    price_livemode = bool(price.get("livemode"))
    link_livemode = bool(payment_link.get("livemode"))
    livemode_check = "ok" if not any([product_livemode, price_livemode, link_livemode]) else "ng"

    price_product = price.get("product")
    if isinstance(price_product, dict):
        price_product = price_product.get("id")
    product_match = "ok" if price_product == result_item["product_id"] else "ng"

    actual_unit_amount = price.get("unit_amount")
    currency = stringify(price.get("currency")).lower()
    price_type = stringify(price.get("type"))
    recurring = price.get("recurring")
    recurring_check = "ok" if price_type == "one_time" and recurring is None else "ng"
    price_ok = expected_price == actual_unit_amount and currency == "jpy"

    actual_product_name = stringify(product.get("name"))
    description_check = check_description(actual_product_name, stringify(product.get("description")))
    tax_code_actual = stringify(product.get("tax_code"))
    tax_code_check = "actual_matches_candidate" if tax_code_actual == EXPECTED_TAX_CODE else ("missing" if not tax_code_actual else "mismatch")

    url = stringify(payment_link.get("url"))
    url_ok = bool(url) and url == result_item.get("payment_link_url")
    link_currency = stringify(payment_link.get("currency")).lower()
    link_currency_ok = link_currency in {"", "jpy"}

    notes: list[str] = []
    if actual_product_name != expected_title:
        notes.append("product name differs from expected Stripe product name")
    if description_check != "ok":
        notes.append("description needs review")
    if not price_ok:
        notes.append("price or currency mismatch")
    if not url_ok:
        notes.append("payment link URL differs from local result")
    if not link_currency_ok:
        notes.append("payment link currency is not jpy")
    if metadata_check != "ok":
        notes.append("metadata mismatch or partial metadata")
    if browser_check != "ok":
        notes.append("browser checkout page is not verified by this script")

    technical_ready_bool = all(
        [
            actual_product_name == expected_title,
            description_check == "ok",
            price_ok,
            recurring_check == "ok",
            product_match == "ok",
            product_active,
            price_active,
            link_active,
            livemode_check == "ok",
            url_ok,
            metadata_check == "ok",
            tax_code_check == "actual_matches_candidate",
        ]
    )
    live_ready = "conditional" if technical_ready_bool and browser_check == "ok" and tax_code_check == "actual_matches_candidate" else "no"

    row = {
        "id": item_id,
        "expected_title": expected_title,
        "actual_product_name": actual_product_name,
        "description_check": description_check,
        "expected_price": stringify(expected_price),
        "actual_unit_amount": stringify(actual_unit_amount),
        "currency": currency,
        "price_type": price_type,
        "recurring_check": recurring_check,
        "product_match": product_match,
        "product_active": "true" if product_active else "false",
        "price_active": "true" if price_active else "false",
        "payment_link_active": "true" if link_active else "false",
        "livemode_check": livemode_check,
        "payment_link_url": url,
        "metadata_check": metadata_check,
        "tax_code_actual": tax_code_actual,
        "tax_code_check": tax_code_check,
        "browser_check": browser_check,
        "technical_ready": "yes" if technical_ready_bool else "no",
        "tax_business_review": "pending",
        "live_ready": live_ready,
        "notes": "; ".join(notes) if notes else "ok",
    }
    counters = {
        "product_retrieved": 1,
        "price_retrieved": 1,
        "payment_link_retrieved": 1,
        "price_ok": 1 if price_ok else 0,
        "one_time_ok": 1 if recurring_check == "ok" else 0,
        "metadata_ok": 1 if metadata_check == "ok" else 0,
        "browser_ok": 1 if browser_check == "ok" else 0,
        "browser_manual_review": 1 if browser_check == "manual_review" else 0,
        "technical_ready_yes": 1 if technical_ready_bool else 0,
        "live_ready_conditional": 1 if live_ready == "conditional" else 0,
        "live_ready_no": 1 if live_ready == "no" else 0,
    }
    return row, counters


def write_csv(path: Path, rows: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=CSV_COLUMNS)
        writer.writeheader()
        writer.writerows(rows)


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


def write_markdown(path: Path, rows: list[dict[str, str]], summary: dict[str, int], local_errors: list[str]) -> None:
    item_rows = [
        [
            row["id"],
            row["expected_title"],
            row["actual_product_name"],
            row["expected_price"],
            row["actual_unit_amount"],
            row["currency"],
            row["metadata_check"],
            row["tax_code_check"],
            row["browser_check"],
            row["technical_ready"],
            row["live_ready"],
            row["notes"],
        ]
        for row in rows
    ]
    local_error_rows = [[error] for error in local_errors]
    content = f"""# Stripe Payment Link Pilot10 Test Review

## Purpose

Review the 10 Stripe test mode Product, Price, and Payment Link objects created in Phase 13C using read-only Stripe API retrieval.

## Safety Conditions

- Stripe Secret Key is read only from `STRIPE_SECRET_KEY`.
- `sk_live_` keys are rejected.
- This script calls only retrieve APIs for Product, Price, and Payment Link.
- No Stripe object is created, updated, disabled, or deleted.
- `ORIGINAL.md`, generated JSON, dake-store-site, and Store production are not updated.
- tax_code is checked only as a configured candidate; tax business review remains pending.

## Target

- review count: {len(rows)}
- local result mode: test
- expected items: 10

## Local Structure Check

{table(['local_error'], local_error_rows)}

## Stripe Test Mode Retrieval Summary

- Product retrieved: {summary.get('product_retrieved', 0)}
- Price retrieved: {summary.get('price_retrieved', 0)}
- Payment Link retrieved: {summary.get('payment_link_retrieved', 0)}
- price ok: {summary.get('price_ok', 0)}
- one_time ok: {summary.get('one_time_ok', 0)}
- metadata ok: {summary.get('metadata_ok', 0)}
- browser ok: {summary.get('browser_ok', 0)}
- browser manual_review: {summary.get('browser_manual_review', 0)}
- technical_ready yes: {summary.get('technical_ready_yes', 0)}
- live_ready conditional: {summary.get('live_ready_conditional', 0)}
- live_ready no: {summary.get('live_ready_no', 0)}

## Item Review

{table(['id', 'expected title', 'actual product name', 'expected price', 'actual amount', 'currency', 'metadata', 'tax_code', 'browser', 'technical', 'live_ready', 'notes'], item_rows)}

## Product Name And Description Review

Product names are compared with the pilot selection Stripe product name. Descriptions must be non-empty, reasonably short, and free of secret-like tokens.

## Price And Billing Review

Each Price must be active, `currency=jpy`, `type=one_time`, `recurring=null`, and linked to the expected Product.

## Payment Link Review

Each Payment Link must be active, test mode only, and its URL must match the Phase 13C local result JSON.

## Metadata Review

Product and Payment Link metadata must match the full DAKE metadata. Price metadata is checked against the Phase 13B Price payload, which contains `dake_item_id`.

## Tax Code Review

Expected candidate: `{EXPECTED_TAX_CODE}`. This is not a final tax determination. `tax_business_review` remains `pending`.

## Browser Manual Review

The script does not open checkout pages. Rows remain `manual_review` unless a separate browser pass confirms the pages.

## Secret Leak Check

The secret value is not written to this report. Use the repository regex check to confirm no key-like token appears in report files.

## Corrections Before Live

Rows with `technical_ready=no`, `browser_check!=ok`, or `tax_code_check` not matching the candidate must be reviewed before any live-mode rollout.

## Judgment

`live_ready` is `conditional` only when technical checks pass, browser check is `ok`, and the candidate tax code matches. This phase never marks a row as final live approval.

## Next Phase Proposal

After human browser review and tax review, decide whether to write selected test Payment Link URLs back to the source planning files or proceed to a live-mode dry-run plan.

## Not Done In This Phase

- No Stripe object creation.
- No Stripe object update.
- No Stripe object deletion.
- No live mode API call.
- No `sk_live_` use.
- No Store production update.
- No `ORIGINAL.md` or generated JSON update.
"""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8", newline="\n")


def main() -> int:
    result = read_json(DEFAULT_TEST_RESULT_JSON)
    payload = read_json(DEFAULT_PAYLOAD_JSON)
    selection = read_selection(DEFAULT_SELECTION_CSV)
    local_errors = validate_local_result(result)
    if local_errors:
        raise RuntimeError("Local Phase 13C result failed validation: " + "; ".join(local_errors))

    secret_key = validate_secret_key()
    stripe = import_stripe_module()
    stripe.api_key = secret_key

    payloads = index_payloads(payload)
    rows: list[dict[str, str]] = []
    summary = {
        "product_retrieved": 0,
        "price_retrieved": 0,
        "payment_link_retrieved": 0,
        "price_ok": 0,
        "one_time_ok": 0,
        "metadata_ok": 0,
        "browser_ok": 0,
        "browser_manual_review": 0,
        "technical_ready_yes": 0,
        "live_ready_conditional": 0,
        "live_ready_no": 0,
    }
    browser_check = "manual_review"
    for item in result["items"]:
        row, counters = retrieve_one(
            stripe=stripe,
            result_item=item,
            payload_item=payloads.get(item["id"]),
            selection_item=selection.get(item["id"]),
            browser_check=browser_check,
        )
        rows.append(row)
        for key, value in counters.items():
            summary[key] = summary.get(key, 0) + value

    write_csv(DEFAULT_REVIEW_CSV, rows)
    write_markdown(DEFAULT_REVIEW_MD, rows, summary, local_errors)
    print("Stripe test mode retrieval review complete.")
    print(f"review_count={len(rows)}")
    print(f"product_retrieved={summary['product_retrieved']}")
    print(f"price_retrieved={summary['price_retrieved']}")
    print(f"payment_link_retrieved={summary['payment_link_retrieved']}")
    print(f"technical_ready_yes={summary['technical_ready_yes']}")
    print(f"live_ready_conditional={summary['live_ready_conditional']}")
    print(f"live_ready_no={summary['live_ready_no']}")
    print(f"wrote {DEFAULT_REVIEW_CSV}")
    print(f"wrote {DEFAULT_REVIEW_MD}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except RuntimeError as exc:
        print(f"ERROR: {exc}")
        raise SystemExit(1) from None
