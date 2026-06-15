from __future__ import annotations

import csv
import hashlib
import json
import subprocess
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
PACK_IDS = ("DAKE_Pack_Document", "DAKE_Pack_Memo")
STORE_PRODUCTS = ROOT / "tools" / "generated" / "store_products.generated.json"
ROLLOUT_REVIEW = ROOT / "tools" / "reports" / "stripe_payment_link_rollout_review.csv"
CSV_OUTPUT = ROOT / "tools" / "reports" / "stripe_pack2_preflight_review.csv"
MD_OUTPUT = ROOT / "tools" / "reports" / "stripe_pack2_preflight_review.md"

CSV_COLUMNS = [
    "id",
    "title",
    "price",
    "currency",
    "source_original",
    "booth_url",
    "github_release_url",
    "distribution_file",
    "distribution_path",
    "purchase_delivery_method",
    "purchase_delivery_ready",
    "tax_code_candidate",
    "stripe_creation_method",
    "review_result",
    "notes",
]


def read_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def read_rollout_rows() -> dict[str, dict[str, str]]:
    if not ROLLOUT_REVIEW.exists():
        return {}
    with ROLLOUT_REVIEW.open("r", encoding="utf-8-sig", newline="") as file:
        return {row["id"]: row for row in csv.DictReader(file)}


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def is_git_tracked(relative_path: str) -> bool:
    result = subprocess.run(
        ["git", "ls-files", "--error-unmatch", relative_path],
        cwd=ROOT,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        check=False,
    )
    return result.returncode == 0


def load_store_items() -> dict[str, dict]:
    data = read_json(STORE_PRODUCTS)
    return {item["id"]: item for item in data.get("items", [])}


def included_app_summary(manifest: dict) -> str:
    apps = manifest.get("included_apps", [])
    return ", ".join(app.get("folder", "") for app in apps if app.get("folder"))


def review_pack(pack_id: str, store_items: dict[str, dict], rollout_rows: dict[str, dict[str, str]]) -> dict[str, str]:
    pack_dir = ROOT / "04_packs" / pack_id
    manifest_path = pack_dir / "pack_manifest.json"
    manifest = read_json(manifest_path)
    store_item = store_items.get(pack_id, {})
    rollout = rollout_rows.get(pack_id, {})

    pack_zip = manifest.get("pack_zip", "")
    pack_zip_path = ROOT / pack_zip
    exists = pack_zip_path.exists()
    actual_size = pack_zip_path.stat().st_size if exists else None
    expected_size = manifest.get("pack_zip_size")
    actual_hash = sha256_file(pack_zip_path) if exists else ""
    expected_hash = str(manifest.get("pack_zip_sha256", "")).lower()
    hash_ok = bool(actual_hash and expected_hash and actual_hash.lower() == expected_hash)
    size_ok = actual_size == expected_size
    tracked = is_git_tracked(pack_zip) if pack_zip else False

    source_original = store_item.get("source_original") or f"04_packs/{pack_id}/ORIGINAL.md"
    booth_url = store_item.get("booth_url") or manifest.get("booth_url") or rollout.get("booth_url", "")
    github_release_url = store_item.get("github_release_url") or rollout.get("github_release_url", "")
    tax_code = rollout.get("tax_code_candidate") or "txcd_10202003"

    notes = [
        f"pack_zip_exists={exists}",
        f"pack_zip_size_ok={size_ok}",
        f"pack_zip_sha256_ok={hash_ok}",
        f"pack_zip_git_tracked={tracked}",
        f"included_apps={included_app_summary(manifest)}",
        "Stripe post-payment delivery route is not confirmed in ORIGINAL.md",
        "Define manual fulfillment or a private download route before live Stripe registration",
    ]

    return {
        "id": pack_id,
        "title": rollout.get("title") or store_item.get("title") or manifest.get("display_name") or pack_id,
        "price": str(store_item.get("price") or manifest.get("price") or ""),
        "currency": store_item.get("currency") or "JPY",
        "source_original": source_original,
        "booth_url": booth_url,
        "github_release_url": github_release_url or "",
        "distribution_file": Path(pack_zip).name if pack_zip else "",
        "distribution_path": pack_zip,
        "purchase_delivery_method": "BOOTH delivery exists; Stripe post-payment fulfillment not confirmed",
        "purchase_delivery_ready": "no",
        "tax_code_candidate": tax_code,
        "stripe_creation_method": "hold",
        "review_result": "hold",
        "notes": "; ".join(notes),
    }


def write_csv(rows: list[dict[str, str]]) -> None:
    CSV_OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    with CSV_OUTPUT.open("w", encoding="utf-8", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=CSV_COLUMNS)
        writer.writeheader()
        writer.writerows(rows)


def markdown_table(rows: list[dict[str, str]]) -> str:
    lines = [
        "| id | title | price | booth | distribution | delivery_ready | review |",
        "|---|---|---:|---|---|---|---|",
    ]
    for row in rows:
        booth = "yes" if row["booth_url"] else "no"
        lines.append(
            f"| {row['id']} | {row['title']} | {row['price']} {row['currency']} | "
            f"{booth} | {row['distribution_file']} | {row['purchase_delivery_ready']} | {row['review_result']} |"
        )
    return "\n".join(lines)


def write_markdown(rows: list[dict[str, str]]) -> None:
    ready = sum(1 for row in rows if row["review_result"] == "ready")
    conditional = sum(1 for row in rows if row["review_result"] == "conditional")
    hold = sum(1 for row in rows if row["review_result"] == "hold")

    detail_lines: list[str] = []
    for row in rows:
        detail_lines.extend(
            [
                f"### {row['id']}",
                "",
                f"- title: {row['title']}",
                f"- price: {row['price']} {row['currency']}",
                f"- source_original: `{row['source_original']}`",
                f"- booth_url: {row['booth_url'] or '(empty)'}",
                f"- github_release_url: {row['github_release_url'] or '(empty)'}",
                f"- distribution_file: `{row['distribution_file']}`",
                f"- distribution_path: `{row['distribution_path']}`",
                f"- purchase_delivery_method: {row['purchase_delivery_method']}",
                f"- purchase_delivery_ready: {row['purchase_delivery_ready']}",
                f"- tax_code_candidate: {row['tax_code_candidate']}",
                f"- stripe_creation_method: {row['stripe_creation_method']}",
                f"- review_result: {row['review_result']}",
                f"- notes: {row['notes']}",
                "",
            ]
        )

    content = f"""# Stripe Pack2 Preflight Review

## Purpose

Review the two DAKE Pack products before Stripe Payment Link registration.

## Safety Scope

- No Stripe API call is made.
- No Stripe Secret Key is read.
- No Product, Price, or Payment Link is created, updated, or deleted.
- ORIGINAL.md, generated JSON, dake-store-site, BOOTH, and Pack ZIP files are not updated.

## Inputs

- `tools/generated/store_products.generated.json`
- `tools/reports/stripe_payment_link_rollout_review.csv`
- `04_packs/DAKE_Pack_Document/pack_manifest.json`
- `04_packs/DAKE_Pack_Memo/pack_manifest.json`

## Outputs

- `tools/reports/stripe_pack2_preflight_review.csv`
- `tools/reports/stripe_pack2_preflight_review.md`

## Summary

- reviewed_packs: {len(rows)}
- ready: {ready}
- conditional: {conditional}
- hold: {hold}
- generated_at: {datetime.now().isoformat(timespec="seconds")}

## Pack Review

{markdown_table(rows)}

## Details

{chr(10).join(detail_lines)}
## Pre-live Required Actions

- Define the post-Stripe purchase delivery route for each Pack.
- If fulfillment is manual, document the operator workflow and purchase message before live registration.
- If fulfillment uses a private download URL or GitHub Release asset, add the confirmed route to the source of truth before Store reflection.
- Review the tax code candidate before live execution. Current candidate: `txcd_10202003`.

## Recommendation

Both Pack ZIPs and BOOTH routes exist, but Stripe post-payment fulfillment is not confirmed in the source of truth. Keep both Packs on hold for live Stripe registration until the delivery route is defined.

## Next Phase

After the delivery route is confirmed, create manual Dashboard Payment Links or a dedicated Pack payload flow, then write confirmed Payment Link URLs back to ORIGINAL.md.
"""
    MD_OUTPUT.write_text(content, encoding="utf-8")


def main() -> None:
    store_items = load_store_items()
    rollout_rows = read_rollout_rows()
    rows = [review_pack(pack_id, store_items, rollout_rows) for pack_id in PACK_IDS]
    write_csv(rows)
    write_markdown(rows)
    print(f"Wrote {CSV_OUTPUT.relative_to(ROOT)}")
    print(f"Wrote {MD_OUTPUT.relative_to(ROOT)}")
    print(f"Reviewed packs: {len(rows)}")
    print(f"Hold: {sum(1 for row in rows if row['review_result'] == 'hold')}")


if __name__ == "__main__":
    main()
