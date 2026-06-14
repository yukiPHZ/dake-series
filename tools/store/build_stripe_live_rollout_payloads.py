from __future__ import annotations

import csv
import hashlib
import json
import re
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[2]
ROLLOUT_REVIEW_CSV = ROOT / "tools" / "reports" / "stripe_payment_link_rollout_review.csv"
GENERATED_STORE_JSON = ROOT / "tools" / "generated" / "store_products.generated.json"
FINAL_PILOT_REVIEW_CSV = ROOT / "tools" / "reports" / "stripe_payment_link_pilot10_final_review.csv"
OUTPUT_JSON = ROOT / "tools" / "reports" / "stripe_payment_link_live_rollout_payloads.json"
OUTPUT_CSV = ROOT / "tools" / "reports" / "stripe_payment_link_live_rollout_payloads.csv"
OUTPUT_MD = ROOT / "tools" / "reports" / "stripe_payment_link_live_rollout_payloads.md"
JST = timezone(timedelta(hours=9))

DRY_RUN_NOTICE = "DRY RUN ONLY. NO STRIPE API CALL. NO LIVE OBJECT IS CREATED."
PRODUCT_PLACEHOLDER = "__PRODUCT_ID_FROM_LIVE_PRODUCT__"
PRICE_PLACEHOLDER = "__PRICE_ID_FROM_LIVE_PRICE__"
STORE_PRODUCT_BASE_URL = "https://store.dakeapp.com/product/"
EXPECTED_CANDIDATE_COUNT = 45
GAME_IDS = {"game_alien_road", "game_diver_catch"}

CSV_COLUMNS = [
    "id",
    "type",
    "title",
    "price",
    "currency",
    "tax_code_candidate",
    "source_original",
    "booth_url",
    "github_release_url",
    "product_name",
    "description_check",
    "metadata_ready",
    "product_payload_sha256",
    "price_payload_sha256",
    "payment_link_payload_sha256",
    "product_idempotency_key",
    "price_idempotency_key",
    "payment_link_idempotency_key",
    "live_dry_run_ready",
    "notes",
]


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def read_generated_items(path: Path) -> dict[str, dict[str, Any]]:
    data = json.loads(path.read_text(encoding="utf-8"))
    items = data.get("items")
    if not isinstance(items, list):
        raise ValueError("generated Store JSON must contain an items array")
    return {str(item.get("id")): item for item in items if isinstance(item, dict) and item.get("id")}


def parse_price(value: str) -> int | None:
    text = str(value or "").replace(",", "").strip()
    if not text:
        return None
    try:
        price = int(text)
    except ValueError:
        return None
    return price if price > 0 else None


def store_url(item_id: str) -> str:
    return f"{STORE_PRODUCT_BASE_URL}?id={item_id}"


def normalize_space(value: object) -> str:
    return " ".join(str(value or "").replace("\r\n", "\n").replace("\r", "\n").split()).strip()


def clean_description(generated_item: dict[str, Any], fallback: str) -> tuple[str, str]:
    text = normalize_space(generated_item.get("description"))
    if not text:
        text = normalize_space(fallback)
    if not text:
        return "", "ng"
    if len(text) > 240:
        return text[:237].rstrip() + "...", "ok"
    return text, "ok"


def canonical_json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def sha256_payload(value: Any) -> str:
    return hashlib.sha256(canonical_json(value).encode("utf-8")).hexdigest()


def safe_id(item_id: str) -> str:
    safe = re.sub(r"[^A-Za-z0-9_-]+", "_", item_id).strip("_")
    return safe or "item"


def idempotency_key(kind: str, item_id: str, payload_hash: str) -> str:
    return f"dake-live-{kind}-v1-{safe_id(item_id)}-{payload_hash[:12]}"


def full_metadata(row: dict[str, str], generated_item: dict[str, Any]) -> dict[str, str]:
    values = {
        "dake_item_id": row["id"],
        "dake_type": row["type"],
        "source_repo": str(generated_item.get("source_repo") or "DAKE_series"),
        "source_original": row["source_original"],
        "store_url": store_url(row["id"]),
        "booth_url": row["booth_url"],
        "github_release_url": row["github_release_url"],
    }
    return {key: value for key, value in values.items() if value}


def core_metadata(row: dict[str, str]) -> dict[str, str]:
    return {
        "dake_item_id": row["id"],
        "dake_type": row["type"],
        "source_original": row["source_original"],
    }


def select_candidates(rows: list[dict[str, str]]) -> tuple[list[dict[str, str]], list[dict[str, str]]]:
    candidates: list[dict[str, str]] = []
    excluded: list[dict[str, str]] = []
    for row in rows:
        is_candidate = (
            row.get("review_result") == "create"
            and row.get("creation_method") == "api_candidate"
            and row.get("price_check") == "price_ok"
            and row.get("metadata_ready") == "yes"
        )
        if is_candidate:
            candidates.append(row)
        else:
            excluded.append(row)
    return candidates, excluded


def validate_final_pilot(path: Path) -> list[str]:
    rows = read_csv(path)
    errors: list[str] = []
    if len(rows) != 10:
        errors.append(f"final pilot review row count is {len(rows)}, expected 10")
    passed = sum(row.get("pilot_validation") == "passed" for row in rows)
    conditional = sum(row.get("live_ready") == "conditional" for row in rows)
    if passed != 10:
        errors.append(f"pilot_validation passed is {passed}, expected 10")
    if conditional != 10:
        errors.append(f"live_ready conditional is {conditional}, expected 10")
    return errors


def build_item(row: dict[str, str], generated_item: dict[str, Any]) -> dict[str, Any]:
    item_id = row["id"]
    price = parse_price(row.get("price", ""))
    product_name = row.get("stripe_product_name") or row.get("title") or item_id
    description, description_check = clean_description(
        generated_item,
        f"{row.get('title', '')} / {generated_item.get('category') or ''}",
    )
    metadata = full_metadata(row, generated_item)
    payment_metadata = core_metadata(row)

    product_payload = {
        "name": product_name,
        "description": description,
        "active": True,
        "tax_code": row.get("tax_code_candidate", ""),
        "metadata": metadata,
    }
    price_payload = {
        "currency": "jpy",
        "unit_amount": price,
        "product": PRODUCT_PLACEHOLDER,
        "metadata": {
            "dake_item_id": item_id,
        },
    }
    payment_link_payload = {
        "line_items": [
            {
                "price": PRICE_PLACEHOLDER,
                "quantity": 1,
            }
        ],
        "metadata": payment_metadata,
        "payment_intent_data": {
            "metadata": payment_metadata,
        },
    }

    product_hash = sha256_payload(product_payload)
    price_hash = sha256_payload(price_payload)
    link_hash = sha256_payload(payment_link_payload)

    notes: list[str] = []
    if description_check != "ok":
        notes.append("description missing")
    if generated_item.get("source_original") and generated_item.get("source_original") != row.get("source_original"):
        notes.append("source_original differs from generated JSON")
    if generated_item.get("booth_url") and generated_item.get("booth_url") != row.get("booth_url"):
        notes.append("booth_url differs from generated JSON")
    if generated_item.get("github_release_url") and generated_item.get("github_release_url") != row.get("github_release_url"):
        notes.append("github_release_url differs from generated JSON")

    required_ready = {
        "id": bool(item_id),
        "title": bool(row.get("title")),
        "price": price is not None,
        "currency": row.get("currency", "").lower() == "jpy",
        "tax_code_candidate": bool(row.get("tax_code_candidate")),
        "source_original": bool(row.get("source_original")),
        "booth_url": bool(row.get("booth_url")),
        "github_release_url": bool(row.get("github_release_url")),
        "metadata_ready": row.get("metadata_ready") == "yes",
        "description": description_check == "ok",
        "payload_hash": all([product_hash, price_hash, link_hash]),
    }
    missing = [key for key, ok in required_ready.items() if not ok]
    if missing:
        notes.append("missing or invalid: " + ", ".join(missing))

    ready = "yes" if not missing else "no"
    tax_review_flag = "yes" if row.get("tax_code_review_required") == "yes" or item_id in GAME_IDS else "no"

    return {
        "id": item_id,
        "type": row.get("type", ""),
        "title": row.get("title", ""),
        "price": price,
        "currency": "jpy",
        "category": generated_item.get("category") or "",
        "source_original": row.get("source_original", ""),
        "booth_url": row.get("booth_url", ""),
        "github_release_url": row.get("github_release_url", ""),
        "product_name": product_name,
        "description_check": description_check,
        "metadata_ready": row.get("metadata_ready", ""),
        "tax_code_candidate": row.get("tax_code_candidate", ""),
        "tax_code_source": "rollout_review",
        "tax_candidate_review": "operator_confirmation_required",
        "tax_code_review_required": tax_review_flag,
        "product_payload": product_payload,
        "price_payload": price_payload,
        "payment_link_payload": payment_link_payload,
        "product_payload_sha256": product_hash,
        "price_payload_sha256": price_hash,
        "payment_link_payload_sha256": link_hash,
        "product_idempotency_key": idempotency_key("product", item_id, product_hash),
        "price_idempotency_key": idempotency_key("price", item_id, price_hash),
        "payment_link_idempotency_key": idempotency_key("link", item_id, link_hash),
        "live_dry_run_ready": ready,
        "notes": "; ".join(notes) if notes else "ok",
    }


def build_payload(
    rollout_rows: list[dict[str, str]],
    generated_items: dict[str, dict[str, Any]],
    final_pilot_errors: list[str],
) -> dict[str, Any]:
    candidates, excluded = select_candidates(rollout_rows)
    errors: list[str] = list(final_pilot_errors)
    if len(candidates) != EXPECTED_CANDIDATE_COUNT:
        errors.append(f"expected {EXPECTED_CANDIDATE_COUNT} live dry-run candidates, got {len(candidates)}")

    seen_ids: set[str] = set()
    duplicate_ids: set[str] = set()
    items: list[dict[str, Any]] = []
    for row in candidates:
        item_id = row.get("id", "")
        if item_id in seen_ids:
            duplicate_ids.add(item_id)
        seen_ids.add(item_id)
        generated_item = generated_items.get(item_id)
        if not generated_item:
            errors.append(f"{item_id}: missing from generated Store JSON")
            continue
        items.append(build_item(row, generated_item))

    duplicate_keys: list[str] = []
    for field in ["product_idempotency_key", "price_idempotency_key", "payment_link_idempotency_key"]:
        values = [item[field] for item in items]
        duplicate_keys.extend(sorted({value for value in values if values.count(value) > 1}))
    if duplicate_ids:
        errors.append("duplicate dake_item_id: " + ", ".join(sorted(duplicate_ids)))
    if duplicate_keys:
        errors.append("duplicate idempotency key: " + ", ".join(sorted(duplicate_keys)))

    return {
        "dry_run": True,
        "notice": DRY_RUN_NOTICE,
        "created_at": datetime.now(JST).isoformat(timespec="seconds"),
        "sources": {
            "rollout_review": "tools/reports/stripe_payment_link_rollout_review.csv",
            "generated_store": "tools/generated/store_products.generated.json",
            "final_pilot_review": "tools/reports/stripe_payment_link_pilot10_final_review.csv",
        },
        "candidate_count": len(candidates),
        "count": len(items),
        "items": items,
        "excluded": [
            {
                "id": row.get("id", ""),
                "title": row.get("title", ""),
                "review_result": row.get("review_result", ""),
                "creation_method": row.get("creation_method", ""),
                "reason": "not create/api_candidate/price_ok/metadata_ready",
            }
            for row in excluded
        ],
        "errors": errors,
        "safety": [
            "DRY RUN ONLY.",
            "NO STRIPE API CALL.",
            "NO LIVE OBJECT IS CREATED.",
            "No Stripe Secret Key is read.",
            "No Product, Price, or Payment Link is created.",
            "No ORIGINAL.md, generated JSON, dake-store-site, or Store production update is performed.",
        ],
        "duplicate_strategy_for_next_phase": [
            "Retrieve live Product list.",
            "Match by metadata.dake_item_id.",
            "If zero matches, create a new live object candidate.",
            "If one match, reuse existing object or stop for operator confirmation.",
            "If two or more matches, stop as an abnormal state.",
            "Use the idempotency key on each POST.",
            "Save progress incrementally.",
        ],
    }


def write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def csv_row(item: dict[str, Any]) -> dict[str, str]:
    return {
        "id": item["id"],
        "type": item["type"],
        "title": item["title"],
        "price": str(item["price"]),
        "currency": item["currency"],
        "tax_code_candidate": item["tax_code_candidate"],
        "source_original": item["source_original"],
        "booth_url": item["booth_url"],
        "github_release_url": item["github_release_url"],
        "product_name": item["product_name"],
        "description_check": item["description_check"],
        "metadata_ready": item["metadata_ready"],
        "product_payload_sha256": item["product_payload_sha256"],
        "price_payload_sha256": item["price_payload_sha256"],
        "payment_link_payload_sha256": item["payment_link_payload_sha256"],
        "product_idempotency_key": item["product_idempotency_key"],
        "price_idempotency_key": item["price_idempotency_key"],
        "payment_link_idempotency_key": item["payment_link_idempotency_key"],
        "live_dry_run_ready": item["live_dry_run_ready"],
        "notes": item["notes"],
    }


def write_csv(path: Path, items: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=CSV_COLUMNS)
        writer.writeheader()
        writer.writerows(csv_row(item) for item in items)


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


def write_markdown(path: Path, payload: dict[str, Any]) -> None:
    items = payload["items"]
    ready_yes = sum(item["live_dry_run_ready"] == "yes" for item in items)
    ready_no = sum(item["live_dry_run_ready"] == "no" for item in items)
    game_apps = sum(item["id"] in GAME_IDS for item in items)
    normal_apps = len(items) - game_apps
    price_errors = sum(item["price"] is None for item in items)
    metadata_errors = sum(item["metadata_ready"] != "yes" for item in items)
    hash_missing = sum(
        not (item["product_payload_sha256"] and item["price_payload_sha256"] and item["payment_link_payload_sha256"])
        for item in items
    )
    idempotency_values: list[str] = []
    for item in items:
        idempotency_values.extend(
            [
                item["product_idempotency_key"],
                item["price_idempotency_key"],
                item["payment_link_idempotency_key"],
            ]
        )
    duplicate_key_count = len({value for value in idempotency_values if idempotency_values.count(value) > 1})
    item_rows = [
        [
            item["id"],
            item["title"],
            item["price"],
            item["tax_code_candidate"],
            item["live_dry_run_ready"],
            item["notes"],
        ]
        for item in items
    ]
    excluded_rows = [
        [
            item["id"],
            item["title"],
            item["review_result"],
            item["creation_method"],
            item["reason"],
        ]
        for item in payload["excluded"]
    ]
    error_rows = [[error] for error in payload["errors"]]
    content = f"""# Stripe Payment Link Live Rollout Dry Run

## Purpose

Generate live-mode Stripe Product, Price, and Payment Link payloads for the 45 API creation candidates without calling Stripe.

## Safety Notice

{DRY_RUN_NOTICE}

- No Stripe Secret Key is read.
- No Product, Price, or Payment Link is created.
- No Stripe Product list is retrieved.
- No `ORIGINAL.md`, generated JSON, dake-store-site, or Store production update is performed.

## Inputs

- `tools/reports/stripe_payment_link_rollout_review.csv`
- `tools/generated/store_products.generated.json`
- `tools/reports/stripe_payment_link_pilot10_final_review.csv`

## Target Conditions

Rows must satisfy `review_result=create`, `creation_method=api_candidate`, `price_check=price_ok`, and `metadata_ready=yes`.

## Excluded

Pack products and preparing products are not included in this live dry-run.

{table(['id', 'title', 'review_result', 'creation_method', 'reason'], excluded_rows)}

## Summary

- total candidates: {payload['candidate_count']}
- payloads generated: {payload['count']}
- live_dry_run_ready yes: {ready_yes}
- live_dry_run_ready no: {ready_no}
- normal apps: {normal_apps}
- game apps: {game_apps}
- price errors: {price_errors}
- metadata errors: {metadata_errors}
- tax candidate review required: {len(items)}
- payload hash missing: {hash_missing}
- idempotency key duplicates: {duplicate_key_count}
- errors: {len(payload['errors'])}

## Target 45

{table(['id', 'title', 'price', 'tax_code', 'ready', 'notes'], item_rows)}

## Product Payload Policy

Products are built with `name`, a concise generated-store description, `active=true`, candidate `tax_code`, and full DAKE metadata.

## Price Payload Policy

Prices use `currency=jpy`, integer `unit_amount`, one-time billing by omitting recurring fields, and `metadata.dake_item_id`.

## Payment Link Payload Policy

Payment Links use one line item with quantity 1. Link metadata and payment intent metadata contain `dake_item_id`, `dake_type`, and `source_original`.

## Metadata Policy

Product metadata includes `dake_item_id`, `dake_type`, `source_repo`, `source_original`, `store_url`, `booth_url`, and `github_release_url`. No personal information or secret values are included.

## Payload Hash

Each Product, Price, and Payment Link payload is canonicalized with sorted JSON keys and hashed with SHA-256.

## Idempotency Keys

Idempotency keys are generated from the safe item id and the first 12 characters of each payload hash. These keys are not secrets and are not sent to Stripe in this phase.

## Duplicate Avoidance Policy

Next phase should retrieve live Products, match by `metadata.dake_item_id`, create only when there is no match, reuse or stop when one match exists, and stop on multiple matches.

## Tax Code Candidate

Tax codes come from `stripe_payment_link_rollout_review.csv`. `tax_candidate_review=operator_confirmation_required` for every item; this dry-run does not make a final tax determination.

## Stop Conditions Before Live Execution

- Target count is not 45.
- Any `live_dry_run_ready=no`.
- Price, currency, metadata, source, BOOTH URL, GitHub Release URL, payload hash, or idempotency key errors exist.
- Duplicate idempotency keys or duplicate DAKE item ids exist.
- Any secret-like value appears in output files.

## Next Phase Proposal

Use this dry-run as the review artifact for a separate live-mode execution script with explicit operator approval and duplicate checks.

## Not Done In This Phase

- No Stripe API call.
- No Stripe Secret Key read.
- No `sk_test_` or `sk_live_` use.
- No Product, Price, or Payment Link creation.
- No Payment Link URL write-back.
- No Store production update.

## Errors

{table(['error'], error_rows)}
"""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8", newline="\n")


def main() -> int:
    rollout_rows = read_csv(ROLLOUT_REVIEW_CSV)
    generated_items = read_generated_items(GENERATED_STORE_JSON)
    final_pilot_errors = validate_final_pilot(FINAL_PILOT_REVIEW_CSV)
    payload = build_payload(rollout_rows, generated_items, final_pilot_errors)
    write_json(OUTPUT_JSON, payload)
    write_csv(OUTPUT_CSV, payload["items"])
    write_markdown(OUTPUT_MD, payload)
    ready_yes = sum(item["live_dry_run_ready"] == "yes" for item in payload["items"])
    ready_no = sum(item["live_dry_run_ready"] == "no" for item in payload["items"])
    game_apps = sum(item["id"] in GAME_IDS for item in payload["items"])
    print(f"live dry-run candidate rows={payload['candidate_count']}")
    print(f"Product payload={payload['count']}")
    print(f"Price payload={payload['count']}")
    print(f"Payment Link payload={payload['count']}")
    print(f"live_dry_run_ready yes={ready_yes}")
    print(f"live_dry_run_ready no={ready_no}")
    print(f"normal apps={payload['count'] - game_apps}")
    print(f"game apps={game_apps}")
    print(f"errors={len(payload['errors'])}")
    print(f"wrote {OUTPUT_JSON}")
    print(f"wrote {OUTPUT_CSV}")
    print(f"wrote {OUTPUT_MD}")
    return 0 if not payload["errors"] and payload["count"] == EXPECTED_CANDIDATE_COUNT and ready_no == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
