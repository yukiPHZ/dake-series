from __future__ import annotations

import csv
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
TECHNICAL_REVIEW_CSV = ROOT / "tools" / "reports" / "stripe_payment_link_pilot10_test_review.csv"
BROWSER_REVIEW_CSV = ROOT / "tools" / "reports" / "stripe_payment_link_pilot10_browser_review.csv"
OUTPUT_CSV = ROOT / "tools" / "reports" / "stripe_payment_link_pilot10_final_review.csv"
OUTPUT_MD = ROOT / "tools" / "reports" / "stripe_payment_link_pilot10_final_review.md"
EXPECTED_COUNT = 10

FINAL_COLUMNS = [
    "id",
    "title",
    "technical_ready",
    "browser_check",
    "tax_code_check",
    "pilot_validation",
    "live_ready",
    "technical_notes",
    "browser_notes",
]

TECHNICAL_REQUIRED = {
    "technical_ready": "yes",
    "livemode_check": "ok",
    "currency": "jpy",
    "price_type": "one_time",
    "recurring_check": "ok",
    "metadata_check": "ok",
    "tax_code_check": "actual_matches_candidate",
}

BROWSER_REQUIRED = {
    "browser_check": "ok",
    "title_check": "ok",
    "price_check": "ok",
    "currency_check": "ok",
    "one_time_display_check": "ok",
    "wording_check": "ok",
}


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def index_by_id(rows: list[dict[str, str]], label: str) -> dict[str, dict[str, str]]:
    indexed: dict[str, dict[str, str]] = {}
    duplicates: list[str] = []
    for row in rows:
        item_id = row.get("id", "").strip()
        if not item_id:
            raise ValueError(f"{label}: row without id")
        if item_id in indexed:
            duplicates.append(item_id)
        indexed[item_id] = row
    if duplicates:
        raise ValueError(f"{label}: duplicate ids: {', '.join(sorted(duplicates))}")
    return indexed


def price_matches(row: dict[str, str]) -> bool:
    return row.get("expected_price", "").strip() == row.get("actual_unit_amount", "").strip()


def technical_passed(row: dict[str, str]) -> bool:
    checks = [row.get(key, "").strip() == expected for key, expected in TECHNICAL_REQUIRED.items()]
    checks.append(price_matches(row))
    return all(checks)


def browser_passed(row: dict[str, str]) -> bool:
    return all(row.get(key, "").strip() == expected for key, expected in BROWSER_REQUIRED.items())


def technical_notes(row: dict[str, str]) -> str:
    notes: list[str] = []
    for key, expected in TECHNICAL_REQUIRED.items():
        actual = row.get(key, "").strip()
        if actual != expected:
            notes.append(f"{key}={actual or '-'}")
    if not price_matches(row):
        notes.append("price mismatch")
    source_notes = row.get("notes", "").strip()
    if source_notes and source_notes != "ok":
        notes.append(source_notes)
    return "; ".join(notes) if notes else "ok"


def browser_notes(row: dict[str, str]) -> str:
    notes: list[str] = []
    for key, expected in BROWSER_REQUIRED.items():
        actual = row.get(key, "").strip()
        if actual != expected:
            notes.append(f"{key}={actual or '-'}")
    source_notes = row.get("notes", "").strip()
    if source_notes:
        notes.append(source_notes)
    method = row.get("review_method", "").strip()
    if method != "human_browser":
        notes.append(f"review_method={method or '-'}")
    return "; ".join(notes) if notes else "ok"


def build_final_rows(technical_rows: list[dict[str, str]], browser_rows: list[dict[str, str]]) -> list[dict[str, str]]:
    if len(technical_rows) != EXPECTED_COUNT:
        raise ValueError(f"expected {EXPECTED_COUNT} technical review rows, got {len(technical_rows)}")
    if len(browser_rows) != EXPECTED_COUNT:
        raise ValueError(f"expected {EXPECTED_COUNT} browser review rows, got {len(browser_rows)}")

    technical_by_id = index_by_id(technical_rows, "technical review")
    browser_by_id = index_by_id(browser_rows, "browser review")
    missing_browser = sorted(set(technical_by_id) - set(browser_by_id))
    extra_browser = sorted(set(browser_by_id) - set(technical_by_id))
    if missing_browser or extra_browser:
        raise ValueError(f"browser review id mismatch: missing={missing_browser}, extra={extra_browser}")

    final_rows: list[dict[str, str]] = []
    for tech in technical_rows:
        item_id = tech["id"]
        browser = browser_by_id[item_id]
        tech_ok = technical_passed(tech)
        browser_ok = browser_passed(browser)
        pilot_validation = "passed" if tech_ok and browser_ok else "failed"
        live_ready = "conditional" if pilot_validation == "passed" else "no"
        final_rows.append(
            {
                "id": item_id,
                "title": tech.get("expected_title") or tech.get("actual_product_name") or item_id,
                "technical_ready": "yes" if tech_ok else "no",
                "browser_check": browser.get("browser_check", ""),
                "tax_code_check": tech.get("tax_code_check", ""),
                "pilot_validation": pilot_validation,
                "live_ready": live_ready,
                "technical_notes": technical_notes(tech),
                "browser_notes": browser_notes(browser),
            }
        )
    return final_rows


def write_csv(path: Path, rows: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=FINAL_COLUMNS)
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


def write_markdown(path: Path, rows: list[dict[str, str]]) -> None:
    technical_yes = sum(row["technical_ready"] == "yes" for row in rows)
    browser_ok = sum(row["browser_check"] == "ok" for row in rows)
    pilot_passed = sum(row["pilot_validation"] == "passed" for row in rows)
    live_conditional = sum(row["live_ready"] == "conditional" for row in rows)
    live_no = sum(row["live_ready"] == "no" for row in rows)
    item_rows = [
        [
            row["id"],
            row["title"],
            row["technical_ready"],
            row["browser_check"],
            row["pilot_validation"],
            row["live_ready"],
        ]
        for row in rows
    ]
    content = f"""# Stripe Payment Link Pilot10 Final Review

## Purpose

Combine the Phase 13D technical review with the independent human browser review record for the 10 Stripe sandbox Payment Links.

## Inputs

- `tools/reports/stripe_payment_link_pilot10_test_review.csv`
- `tools/reports/stripe_payment_link_pilot10_browser_review.csv`

## Technical Review Result

- technical_ready yes: {technical_yes}
- tax_code remains a candidate and requires operator confirmation before live execution.

## Human Browser Review Result

- browser_check ok: {browser_ok}
- review method: human_browser
- review date: 2026-06-14
- source note: user confirmed all 10 Stripe sandbox checkout pages.

## Summary

- target: {len(rows)}
- technical_ready yes: {technical_yes}
- browser_check ok: {browser_ok}
- pilot_validation passed: {pilot_passed}
- live_ready conditional: {live_conditional}
- live_ready no: {live_no}

## Items

{table(['id', 'title', 'technical', 'browser', 'pilot_validation', 'live_ready'], item_rows)}

## Tax Code Handling

`tax_code_check=actual_matches_candidate` confirms that Stripe test objects match the candidate tax code. It is not a final tax determination.

## Judgment

`pilot_validation=passed` means the pilot technical checks and the human browser display checks both passed. `live_ready=conditional` means the item can proceed to live dry-run planning, not final live approval.

## Conditions Before Live Execution

- Human operator approval is still required.
- Tax candidate review remains required.
- No live Stripe object should be created from this report alone.
- Idempotency and duplicate checks must be performed in the live execution phase.

## Not Done In This Phase

- No Stripe API call.
- No Stripe Secret Key read.
- No Product, Price, or Payment Link creation.
- No `ORIGINAL.md`, generated JSON, dake-store-site, or Store production update.
"""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8", newline="\n")


def main() -> int:
    technical_rows = read_csv(TECHNICAL_REVIEW_CSV)
    browser_rows = read_csv(BROWSER_REVIEW_CSV)
    final_rows = build_final_rows(technical_rows, browser_rows)
    write_csv(OUTPUT_CSV, final_rows)
    write_markdown(OUTPUT_MD, final_rows)
    print(f"final pilot review rows={len(final_rows)}")
    print(f"pilot_validation passed={sum(row['pilot_validation'] == 'passed' for row in final_rows)}")
    print(f"live_ready conditional={sum(row['live_ready'] == 'conditional' for row in final_rows)}")
    print(f"live_ready no={sum(row['live_ready'] == 'no' for row in final_rows)}")
    print(f"wrote {OUTPUT_CSV}")
    print(f"wrote {OUTPUT_MD}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
