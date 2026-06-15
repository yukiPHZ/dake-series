from __future__ import annotations

import csv
import hashlib
import json
import re
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
PACK_IDS = ("DAKE_Pack_Document", "DAKE_Pack_Memo")
STORE_PRODUCTS = ROOT / "tools" / "generated" / "store_products.generated.json"
RULE_FILE = ROOT / "00_core" / "DAKE_PACK_MANUAL_DELIVERY_RULE.md"
EMAIL_TEMPLATE = ROOT / "tools" / "templates" / "stripe_pack_manual_delivery_email.txt"
LOG_TEMPLATE = ROOT / "tools" / "templates" / "stripe_pack_manual_delivery_log.example.csv"
CSV_OUTPUT = ROOT / "tools" / "reports" / "stripe_pack2_manual_delivery_ready.csv"
MD_OUTPUT = ROOT / "tools" / "reports" / "stripe_pack2_manual_delivery_ready.md"

SECRET_RE = re.compile(r"sk_(?:test|live)_[A-Za-z0-9]{16,}|whsec_[A-Za-z0-9]{16,}")
PRIVATE_URL_RE = re.compile(r"https?://(?!peakheadz\.booth\.pm|peakheadz\.com|github\.com|store\.dakeapp\.com)[^\s)`]+")

CSV_COLUMNS = [
    "id",
    "title",
    "price",
    "currency",
    "source_original",
    "booth_url",
    "distribution_file",
    "distribution_path",
    "zip_size",
    "zip_sha256",
    "purchase_delivery_method",
    "purchase_delivery_ready",
    "stripe_creation_method",
    "review_result",
    "payment_status",
    "stripe_payment_link",
    "checks",
    "notes",
]


def read_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def load_store_items() -> dict[str, dict]:
    data = read_json(STORE_PRODUCTS)
    return {item["id"]: item for item in data.get("items", [])}


def has_no_secret(path: Path) -> bool:
    return SECRET_RE.search(path.read_text(encoding="utf-8")) is None


def has_no_private_url(path: Path) -> bool:
    return PRIVATE_URL_RE.search(path.read_text(encoding="utf-8")) is None


def review_pack(pack_id: str, store_items: dict[str, dict]) -> dict[str, str]:
    pack_dir = ROOT / "04_packs" / pack_id
    original_path = pack_dir / "ORIGINAL.md"
    manifest = read_json(pack_dir / "pack_manifest.json")
    original = original_path.read_text(encoding="utf-8")
    store_item = store_items.get(pack_id, {})

    pack_zip = manifest.get("pack_zip", "")
    zip_path = ROOT / pack_zip
    zip_exists = zip_path.exists()
    actual_size = zip_path.stat().st_size if zip_exists else None
    expected_size = manifest.get("pack_zip_size")
    actual_hash = sha256_file(zip_path) if zip_exists else ""
    expected_hash = str(manifest.get("pack_zip_sha256", "")).lower()

    checks = {
        "original_exists": original_path.exists(),
        "price_exists": bool(store_item.get("price") or manifest.get("price")),
        "booth_url_exists": bool(store_item.get("booth_url") or manifest.get("booth_url")),
        "pack_zip_exists": zip_exists,
        "pack_zip_size_ok": actual_size == expected_size,
        "pack_zip_sha256_ok": actual_hash.lower() == expected_hash,
        "manual_delivery_method": "manual_email_private_download" in original,
        "delivery_window": "next business day" in original or "次営業日以内" in original,
        "buyer_notice": "Buyer notice" in original,
        "resend_policy": "Resend and failure handling" in original,
        "rule_reference": "00_core/DAKE_PACK_MANUAL_DELIVERY_RULE.md" in original,
        "payment_status_booth_only": store_item.get("payment_status") == "booth_only",
        "stripe_payment_link_empty": not store_item.get("stripe_payment_link"),
    }
    ready = all(checks.values())

    return {
        "id": pack_id,
        "title": store_item.get("title") or manifest.get("display_name") or pack_id,
        "price": str(store_item.get("price") or manifest.get("price") or ""),
        "currency": store_item.get("currency") or "JPY",
        "source_original": store_item.get("source_original") or f"04_packs/{pack_id}/ORIGINAL.md",
        "booth_url": store_item.get("booth_url") or manifest.get("booth_url") or "",
        "distribution_file": Path(pack_zip).name,
        "distribution_path": pack_zip,
        "zip_size": str(actual_size or ""),
        "zip_sha256": actual_hash.lower(),
        "purchase_delivery_method": "manual_email_private_download",
        "purchase_delivery_ready": "yes" if ready else "no",
        "stripe_creation_method": "manual_dashboard_ready" if ready else "hold",
        "review_result": "ready" if ready else "hold",
        "payment_status": store_item.get("payment_status") or "",
        "stripe_payment_link": str(store_item.get("stripe_payment_link") or ""),
        "checks": "; ".join(f"{key}={value}" for key, value in checks.items()),
        "notes": "Manual delivery is defined; no Stripe API call is made by this check." if ready else "Manual delivery readiness is incomplete.",
    }


def write_csv(rows: list[dict[str, str]]) -> None:
    CSV_OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    with CSV_OUTPUT.open("w", encoding="utf-8", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=CSV_COLUMNS)
        writer.writeheader()
        writer.writerows(rows)


def write_markdown(rows: list[dict[str, str]], safety_checks: dict[str, bool]) -> None:
    ready = sum(1 for row in rows if row["review_result"] == "ready")
    hold = sum(1 for row in rows if row["review_result"] == "hold")
    conditional = sum(1 for row in rows if row["review_result"] == "conditional")
    delivery_ready = sum(1 for row in rows if row["purchase_delivery_ready"] == "yes")

    table = [
        "| id | title | price | file | sha256 | delivery_ready | method | review |",
        "|---|---|---:|---|---|---|---|---|",
    ]
    for row in rows:
        table.append(
            f"| {row['id']} | {row['title']} | {row['price']} {row['currency']} | "
            f"{row['distribution_file']} | `{row['zip_sha256']}` | {row['purchase_delivery_ready']} | "
            f"{row['stripe_creation_method']} | {row['review_result']} |"
        )

    safety = "\n".join(f"- {key}: {value}" for key, value in safety_checks.items())
    content = f"""# Stripe Pack2 Manual Delivery Ready

## Purpose

Confirm that the two DAKE Pack products have a documented manual delivery operation before Stripe Payment Link registration.

## Safety Scope

- No Stripe API call is made.
- No Stripe Secret Key is read.
- No Product, Price, or Payment Link is created.
- Payment Link URLs are not written back.
- generated JSON and Store files are not regenerated.
- Pack ZIP files are not rebuilt or moved.
- No buyer information is stored.
- No public download URL is added.

## Summary

- reviewed_packs: {len(rows)}
- ready: {ready}
- conditional: {conditional}
- hold: {hold}
- purchase_delivery_ready_yes: {delivery_ready}
- generated_at: {datetime.now().isoformat(timespec="seconds")}

## Pack Results

{chr(10).join(table)}

## Safety Checks

{safety}

## Common Operation

- delivery_method: `manual_email_private_download`
- delivery_window: within the next business day after payment confirmation
- payment_confirmation: Stripe Dashboard manual confirmation
- delivery_record: secure local log outside Git
- resend: verify payment, buyer email, Pack, and previous delivery record
- email_failure: record `delivery_failed`, verify email and payment information, then follow existing DAKE Store refund/support policy
- personal_information: do not store buyer data in Git, generated JSON, public Store files, or Markdown reports

## Next Phase

Create Stripe Dashboard Payment Links manually for the two Packs after final human confirmation, then write only the confirmed Payment Link URLs back to ORIGINAL.md in a later phase.
"""
    MD_OUTPUT.write_text(content, encoding="utf-8", newline="\n")


def main() -> None:
    store_items = load_store_items()
    rows = [review_pack(pack_id, store_items) for pack_id in PACK_IDS]
    files_to_scan = [
        RULE_FILE,
        EMAIL_TEMPLATE,
        LOG_TEMPLATE,
        ROOT / "04_packs" / "DAKE_Pack_Document" / "ORIGINAL.md",
        ROOT / "04_packs" / "DAKE_Pack_Memo" / "ORIGINAL.md",
    ]
    safety_checks = {
        "rule_exists": RULE_FILE.exists(),
        "email_template_exists": EMAIL_TEMPLATE.exists(),
        "log_template_exists": LOG_TEMPLATE.exists(),
        "no_secret_pattern": all(has_no_secret(path) for path in files_to_scan if path.exists()),
        "no_unapproved_public_download_url": all(has_no_private_url(path) for path in files_to_scan if path.exists()),
        "payment_status_unchanged": all(row["payment_status"] == "booth_only" for row in rows),
        "payment_link_not_set": all(not row["stripe_payment_link"] for row in rows),
    }
    write_csv(rows)
    write_markdown(rows, safety_checks)

    ready = sum(1 for row in rows if row["review_result"] == "ready")
    hold = sum(1 for row in rows if row["review_result"] == "hold")
    print(f"ready={ready}")
    print(f"hold={hold}")
    print(f"purchase_delivery_ready=yes:{sum(1 for row in rows if row['purchase_delivery_ready'] == 'yes')}")
    for key, value in safety_checks.items():
        print(f"{key}={value}")
    if ready != len(PACK_IDS) or hold:
        raise SystemExit(1)
    if not all(safety_checks.values()):
        raise SystemExit(1)


if __name__ == "__main__":
    main()
