from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path
from typing import Any


DEFAULT_INPUT = Path(__file__).resolve().parents[1] / "generated" / "store_products.generated.json"
DEFAULT_OUTPUT = Path(__file__).resolve().parents[1] / "reports" / "stripe_payment_link_candidates.csv"
STORE_PRODUCT_BASE_URL = "https://store.dakeapp.com/product/"

CSV_FIELDS = [
    "store_id",
    "type",
    "title",
    "price",
    "currency",
    "description",
    "payment_status",
    "classification",
    "category",
    "booth_url",
    "github_release_url",
    "source_repo",
    "source_original",
    "source_original_exists",
    "stripe_product_name",
    "stripe_price_amount",
    "stripe_payment_link_target",
    "tax_code_candidate",
    "tax_code_memo",
    "store_url",
    "metadata_json",
    "memo",
]


def load_items(path: Path) -> list[dict[str, Any]]:
    with path.open("r", encoding="utf-8") as file:
        data = json.load(file)
    items = data.get("items")
    if not isinstance(items, list):
        raise ValueError("store_products.generated.json must contain an items array")
    return [item for item in items if isinstance(item, dict)]


def store_url(item: dict[str, Any]) -> str:
    value = item.get("store_url")
    if isinstance(value, str) and value.strip():
        return value.strip()
    return f"{STORE_PRODUCT_BASE_URL}?id={item.get('id', '')}"


def source_original_exists(item: dict[str, Any], repo_root: Path) -> bool:
    source = str(item.get("source_original") or "").strip()
    if not source:
        return False
    path = Path(source)
    if not path.is_absolute():
        source_repo = str(item.get("source_repo") or "").strip()
        if source_repo == "DAKE_series":
            path = repo_root / source
    return path.exists()


def tax_code_candidate(item: dict[str, Any]) -> tuple[str, str]:
    text = " ".join(
        str(item.get(key) or "")
        for key in ("id", "type", "title", "category")
    ).lower()
    tags = item.get("tags")
    if isinstance(tags, list):
        text += " " + " ".join(str(tag).lower() for tag in tags)
    if "game" in text or "ゲーム" in text:
        return "txcd_10201000", "downloadable video game candidate; confirm before Stripe Tax use"
    return "txcd_10202003", "downloadable prewritten software for business use candidate; confirm tax treatment"


def metadata_for(item: dict[str, Any]) -> dict[str, str]:
    keys = {
        "dake_item_id": item.get("id"),
        "dake_type": item.get("type"),
        "source_repo": item.get("source_repo"),
        "source_original": item.get("source_original"),
        "store_url": store_url(item),
        "booth_url": item.get("booth_url"),
        "github_release_url": item.get("github_release_url"),
    }
    return {key: str(value) for key, value in keys.items() if value not in (None, "")}


def classify(item: dict[str, Any], repo_root: Path) -> tuple[str, bool, str]:
    status = str(item.get("status") or "").strip().lower()
    payment_status = str(item.get("payment_status") or "").strip().lower()
    if status != "available":
        return "not_available", False, "not an available Store item"
    if payment_status == "stripe_ready":
        return "stripe_ready", False, "already has a Stripe Payment Link"
    if payment_status == "preparing":
        return "preparing", False, "preparing item; do not create Stripe link yet"

    missing: list[str] = []
    if not item.get("id"):
        missing.append("id")
    if not item.get("title"):
        missing.append("title")
    if not isinstance(item.get("price"), int) or int(item.get("price") or 0) <= 0:
        missing.append("price")
    if str(item.get("currency") or "").upper() != "JPY":
        missing.append("currency")
    if not item.get("source_original"):
        missing.append("source_original")
    if not source_original_exists(item, repo_root):
        missing.append("source_original_exists")
    if missing:
        return "needs_review", False, "missing " + ", ".join(missing)

    memo_parts: list[str] = []
    if not item.get("booth_url"):
        memo_parts.append("booth_url missing")
    if not item.get("github_release_url"):
        memo_parts.append("github_release_url missing; acceptable for packs if delivery policy is separate")
    memo = "; ".join(memo_parts) if memo_parts else "ready for Stripe Payment Link planning"
    return "stripe_candidate", True, memo


def candidate_rows(items: list[dict[str, Any]], repo_root: Path) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for item in items:
        status = str(item.get("status") or "").strip().lower()
        payment_status = str(item.get("payment_status") or "").strip().lower()
        if status != "available" or payment_status == "stripe_ready":
            continue
        classification, create_link, memo = classify(item, repo_root)
        tax_code, tax_memo = tax_code_candidate(item)
        rows.append(
            {
                "store_id": str(item.get("id") or ""),
                "type": str(item.get("type") or ""),
                "title": str(item.get("title") or ""),
                "price": str(item.get("price") or ""),
                "currency": str(item.get("currency") or ""),
                "description": str(item.get("description") or ""),
                "payment_status": str(item.get("payment_status") or ""),
                "classification": classification,
                "category": str(item.get("category") or ""),
                "booth_url": str(item.get("booth_url") or ""),
                "github_release_url": str(item.get("github_release_url") or ""),
                "source_repo": str(item.get("source_repo") or ""),
                "source_original": str(item.get("source_original") or ""),
                "source_original_exists": "yes" if source_original_exists(item, repo_root) else "no",
                "stripe_product_name": str(item.get("title") or ""),
                "stripe_price_amount": str(item.get("price") or ""),
                "stripe_payment_link_target": "yes" if create_link else "no",
                "tax_code_candidate": tax_code,
                "tax_code_memo": tax_memo,
                "store_url": store_url(item),
                "metadata_json": json.dumps(metadata_for(item), ensure_ascii=False, sort_keys=True),
                "memo": memo,
            }
        )
    return rows


def write_csv(rows: list[dict[str, str]], output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=CSV_FIELDS, lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Export Stripe Payment Link planning candidates without calling Stripe APIs."
    )
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    repo_root = Path(__file__).resolve().parents[2]
    rows = candidate_rows(load_items(args.input), repo_root)
    write_csv(rows, args.output)
    counts: dict[str, int] = {}
    for row in rows:
        counts[row["classification"]] = counts.get(row["classification"], 0) + 1
    print(f"wrote {len(rows)} rows to {args.output}")
    print("classification: " + ", ".join(f"{key}={counts[key]}" for key in sorted(counts)))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
