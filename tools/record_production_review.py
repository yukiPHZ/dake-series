from __future__ import annotations

import argparse
from typing import Any

from store.release_pipeline_core import ReleasePipeline, metadata_line, now_jst, write_json_atomic, write_text_atomic


COMMON_CHECKS = [
    ("product_page_http_ok", "Product page returns HTTP 200 / visible page"),
    ("product_name_visible", "Product name is correct and visible"),
    ("price_visible", "Price is correct and visible"),
    ("stripe_button_visible", "Stripe purchase button is visible"),
    ("payment_link_matches_source", "Stripe link matches source of truth"),
    ("booth_link_visible", "BOOTH link is visible"),
    ("booth_link_correct", "BOOTH link matches source of truth"),
    ("manual_delivery_notice_visible", "Manual delivery notice is visible"),
    ("next_business_day_notice_visible", "Next-business-day notice is visible"),
    ("test_url_detected", "No test_ URL is visible", True),
    ("private_url_exposed", "No private download URL is exposed", True),
    ("local_path_exposed", "No local path is exposed", True),
    ("zip_url_exposed", "No ZIP URL is exposed", True),
    ("actual_payment_completed", "No actual payment was completed", True),
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Record a DAKE Store production page review for a DAKE product.")
    parser.add_argument("product_id", help="DAKE product id")
    parser.add_argument("--replace-failed-review", action="store_true", help="Allow replacing an existing failed production review.")
    return parser.parse_args()


def ask_bool(prompt: str) -> bool:
    while True:
        answer = input(f"{prompt} [y/N]: ").strip().lower()
        if answer in {"y", "yes"}:
            return True
        if answer in {"", "n", "no"}:
            return False
        print("Please answer y or n.")


def expected_value(inverted: bool) -> bool:
    return False if inverted else True


def product_specific_notice_required(pipeline: ReleasePipeline, product_id: str) -> bool:
    sources, duplicates = pipeline.discover_sources()
    if product_id in duplicates:
        return False
    source = sources.get(product_id)
    if source is None:
        return False
    return metadata_line(source.text, "checkout_notice_required").lower() == "yes"


def render_review_md(review: dict[str, Any]) -> str:
    rows = "\n".join(f"- {key}: {value}" for key, value in review.items() if key != "notes")
    notes = "\n".join(f"- {note}" for note in review.get("notes", [])) or "- none"
    return f"""# Production Review

## Summary

{rows}

## Notes

{notes}
"""


def main() -> int:
    args = parse_args()
    pipeline = ReleasePipeline()
    status = pipeline.status(args.product_id)
    if status["current_stage"] in {"SOURCE_INVALID", "INCONSISTENT", "PREPARING_BLOCKED"}:
        print(f"refusing production review at stage={status['current_stage']}")
        for error in status.get("errors", []):
            print(f"- {error}")
        return 1

    artifact_dir = pipeline.artifact_dir(args.product_id)
    json_path = artifact_dir / "production_review.json"
    md_path = artifact_dir / "production_review.md"
    existing = None
    if json_path.exists():
        existing, read_error = pipeline.read_optional_json(json_path)
        if read_error:
            print(read_error)
            return 1
        if existing and existing.get("review_status") == "failed" and not args.replace_failed_review:
            print("existing failed production review found; pass --replace-failed-review to replace it")
            return 1
        if existing and existing.get("review_status") != "failed":
            print("existing production review is not failed; refusing to overwrite")
            return 1

    checks: dict[str, Any] = {
        "product_id": args.product_id,
        "reviewed_at": now_jst(),
        "product_page_url": f"https://store.dakeapp.com/product/?id={args.product_id}",
        "product_name": status.get("title") or args.product_id,
        "price": status.get("price"),
        "currency": "jpy",
    }
    failures: list[str] = []

    for entry in COMMON_CHECKS:
        key = entry[0]
        prompt = entry[1]
        inverted = bool(entry[2]) if len(entry) > 2 else False
        answer = ask_bool(prompt)
        value = not answer if inverted else answer
        checks[key] = value
        if value != expected_value(inverted):
            failures.append(key)

    checks["page_status"] = 200 if checks["product_page_http_ok"] else 0
    notice_required = product_specific_notice_required(pipeline, args.product_id)
    checks["product_specific_notice_required"] = notice_required
    if notice_required:
        value = ask_bool("Product-specific notice is visible")
        checks["product_specific_notice_visible"] = value
        if not value:
            failures.append("product_specific_notice_visible")
    else:
        checks["product_specific_notice_visible"] = "not_applicable"

    console_ok = ask_bool("Console has no errors")
    checks["console_errors"] = 0 if console_ok else 1
    if not console_ok:
        failures.append("console_errors")

    notes_text = input("Notes (optional): ").strip()
    checks["notes"] = [notes_text] if notes_text else []
    checks["review_status"] = "passed" if not failures else "failed"
    checks["failed_checks"] = failures

    write_json_atomic(json_path, checks)
    write_text_atomic(md_path, render_review_md(checks))
    print(f"review_status={checks['review_status']}")
    print(f"output_json={json_path}")
    print(f"output_md={md_path}")
    return 0 if checks["review_status"] == "passed" else 1


if __name__ == "__main__":
    raise SystemExit(main())
