from __future__ import annotations

import argparse
import csv
import json
import os
import re
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Any


EXPECTED_COUNT = 45
STORE_STATUS_LINE = "- Store雋ｩ螢ｲ迥ｶ諷・ stripe_ready"
STRIPE_LINK_PREFIX = "https://buy.stripe.com/"
TEST_LINK_MARKER = "https://buy.stripe.com/test_"


def repo_root() -> Path:
    return Path(__file__).resolve().parents[2]


ROOT = repo_root()
REPORT_DIR = ROOT / "tools" / "reports"
RESULT_JSON = REPORT_DIR / "stripe_payment_link_live_execution_result.json"
STATE_JSON = REPORT_DIR / "stripe_payment_link_live_execution_state.json"
PLAN_CSV = REPORT_DIR / "stripe_live_results_writeback_plan.csv"
PLAN_MD = REPORT_DIR / "stripe_live_results_writeback_plan.md"


@dataclass
class PlanRow:
    item_id: str
    title: str
    source_original: str
    current_payment_status: str
    current_stripe_payment_link: str
    new_payment_status: str
    new_stripe_payment_link: str
    action: str
    validation_result: str
    notes: str
    absolute_path: Path


def read_json(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as handle:
        data = json.load(handle)
    if not isinstance(data, dict):
        raise ValueError(f"JSON root must be an object: {path}")
    return data


def write_text_atomic(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_name(f"{path.name}.tmp")
    with temp_path.open("w", encoding="utf-8", newline="") as handle:
        handle.write(text)
        handle.flush()
        os.fsync(handle.fileno())
    os.replace(temp_path, path)


def normalize_path_label(path: Path) -> str:
    return path.relative_to(ROOT).as_posix()


def resolve_source_original(source_original: str) -> Path:
    if not source_original:
        raise ValueError("source_original is empty")
    source_path = Path(source_original)
    if source_path.is_absolute():
        raise ValueError(f"source_original must be repo-relative: {source_original}")
    if any(part == ".." for part in source_path.parts):
        raise ValueError(f"source_original must not contain '..': {source_original}")
    resolved = (ROOT / source_path).resolve()
    try:
        resolved.relative_to(ROOT.resolve())
    except ValueError as exc:
        raise ValueError(f"source_original escapes repo root: {source_original}") from exc
    if resolved.name != "ORIGINAL.md":
        raise ValueError(f"source_original must point to ORIGINAL.md: {source_original}")
    if not resolved.exists():
        raise FileNotFoundError(f"source_original does not exist: {source_original}")
    return resolved


def metadata_source_original(item: dict[str, Any]) -> str:
    metadata = item.get("metadata")
    if not isinstance(metadata, dict):
        return ""
    return str(metadata.get("source_original") or "")


def validate_inputs(result: dict[str, Any], state: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    if result.get("mode") != "live":
        errors.append("result.mode must be live")
    if result.get("count") != EXPECTED_COUNT:
        errors.append(f"result.count must be {EXPECTED_COUNT}")
    if result.get("errors") not in ([], None):
        errors.append("result.errors must be empty")
    if state.get("mode") != "live":
        errors.append("state.mode must be live")
    if state.get("expected_count") != EXPECTED_COUNT:
        errors.append(f"state.expected_count must be {EXPECTED_COUNT}")
    if state.get("status") != "completed":
        errors.append("state.status must be completed")

    result_items = result.get("items")
    state_items = state.get("items")
    if not isinstance(result_items, list):
        errors.append("result.items must be an array")
        result_items = []
    if not isinstance(state_items, list):
        errors.append("state.items must be an array")
        state_items = []
    if len(result_items) != EXPECTED_COUNT:
        errors.append(f"result.items must contain {EXPECTED_COUNT} items")
    if len(state_items) != EXPECTED_COUNT:
        errors.append(f"state.items must contain {EXPECTED_COUNT} items")

    ids = [str(item.get("id") or "") for item in result_items if isinstance(item, dict)]
    urls = [str(item.get("payment_link_url") or "") for item in result_items if isinstance(item, dict)]
    sources = [metadata_source_original(item) for item in result_items if isinstance(item, dict)]
    for label, values in {"id": ids, "payment_link_url": urls, "source_original": sources}.items():
        duplicates = sorted(value for value, count in Counter(values).items() if value and count > 1)
        if duplicates:
            errors.append(f"duplicate {label}: {', '.join(duplicates)}")

    state_by_id = {str(item.get("id") or ""): item for item in state_items if isinstance(item, dict)}
    for item in result_items:
        if not isinstance(item, dict):
            errors.append("result item must be an object")
            continue
        item_id = str(item.get("id") or "")
        source_original = metadata_source_original(item)
        payment_link_url = str(item.get("payment_link_url") or "")
        for field in ["product_id", "price_id", "payment_link_id"]:
            if not item.get(field):
                errors.append(f"{item_id}: {field} is missing")
        if item.get("status") != "completed":
            errors.append(f"{item_id}: status must be completed")
        if item.get("livemode") is not True:
            errors.append(f"{item_id}: livemode must be true")
        if item.get("manual_resolution_required") not in (False, None):
            errors.append(f"{item_id}: manual_resolution_required must be false")
        if item.get("error") not in (None, ""):
            errors.append(f"{item_id}: error must be null")
        if not payment_link_url.startswith(STRIPE_LINK_PREFIX):
            errors.append(f"{item_id}: payment_link_url must start with {STRIPE_LINK_PREFIX}")
        if payment_link_url.startswith(TEST_LINK_MARKER):
            errors.append(f"{item_id}: test Payment Link URL is not allowed")
        try:
            resolve_source_original(source_original)
        except Exception as exc:
            errors.append(f"{item_id}: {exc}")

        state_item = state_by_id.get(item_id)
        if not state_item:
            errors.append(f"{item_id}: missing in state")
        else:
            if state_item.get("status") != "completed":
                errors.append(f"{item_id}: state status must be completed")
            if state_item.get("livemode") is not True:
                errors.append(f"{item_id}: state livemode must be true")
    return errors


def line_ending(text: str) -> str:
    return "\r\n" if "\r\n" in text else "\n"


def split_keep_end(text: str) -> tuple[list[str], str, bool]:
    newline = line_ending(text)
    has_trailing = text.endswith(("\n", "\r"))
    normalized = text.replace("\r\n", "\n")
    return normalized.split("\n"), newline, has_trailing


def find_store_section(lines: list[str]) -> tuple[int, int] | None:
    start = None
    for index, line in enumerate(lines):
        if re.match(r"^##\s+.*Store.*$", line):
            start = index
            break
    if start is None:
        return None
    end = len(lines)
    for index in range(start + 1, len(lines)):
        if re.match(r"^##\s+", lines[index]):
            end = index
            break
    return start, end


def current_payment_status(current_link: str) -> str:
    return "stripe_ready" if current_link else "booth_only"


def read_current_link(section_lines: list[str]) -> str:
    for line in section_lines:
        match = re.match(r"^\s*-\s+Stripe Payment Link:\s*(.*?)\s*$", line)
        if match:
            value = match.group(1).strip()
            return value if value.startswith("http") else ""
    return ""


def has_ready_status(section_lines: list[str]) -> bool:
    return any("stripe_ready" in line and ("Store" in line or "payment_status" in line) for line in section_lines)


def plan_for_item(item: dict[str, Any]) -> PlanRow:
    item_id = str(item["id"])
    title = str(item.get("title") or "")
    source_original = metadata_source_original(item)
    absolute_path = resolve_source_original(source_original)
    payment_link_url = str(item["payment_link_url"])
    text = absolute_path.read_text(encoding="utf-8")
    lines, _, _ = split_keep_end(text)
    section = find_store_section(lines)
    if section is None:
        return PlanRow(
            item_id,
            title,
            source_original,
            "missing_field",
            "",
            "stripe_ready",
            payment_link_url,
            "stop_missing",
            "missing_store_section",
            "Store section was not found",
            absolute_path,
        )
    start, end = section
    section_lines = lines[start:end]
    current_link = read_current_link(section_lines)
    current_status = current_payment_status(current_link)
    if current_link and current_link != payment_link_url:
        return PlanRow(
            item_id,
            title,
            source_original,
            current_status,
            current_link,
            "stripe_ready",
            payment_link_url,
            "stop_conflict",
            "conflict",
            "Existing Stripe Payment Link differs from live result",
            absolute_path,
        )
    if current_link == payment_link_url and has_ready_status(section_lines):
        action = "already_same"
        note = "Already contains the live Payment Link and stripe_ready status"
    else:
        action = "update"
        note = "Write live Payment Link and stripe_ready status"
    return PlanRow(
        item_id,
        title,
        source_original,
        current_status,
        current_link,
        "stripe_ready",
        payment_link_url,
        action,
        "ok",
        note,
        absolute_path,
    )


def update_original(path: Path, payment_link_url: str) -> None:
    text = path.read_text(encoding="utf-8")
    lines, newline, has_trailing = split_keep_end(text)
    section = find_store_section(lines)
    if section is None:
        raise ValueError(f"Store section was not found: {path}")
    start, end = section
    link_index = None
    status_index = None
    for index in range(start + 1, end):
        if re.match(r"^\s*-\s+Stripe Payment Link:\s*", lines[index]):
            link_index = index
        elif "stripe_ready" in lines[index] and ("Store" in lines[index] or "payment_status" in lines[index]):
            status_index = index
        elif lines[index].lstrip().startswith("- Store雋ｩ螢ｲ迥ｶ諷・"):
            status_index = index
    link_line = f"- Stripe Payment Link: {payment_link_url}"
    if link_index is None:
        insert_at = end
        while insert_at > start + 1 and not lines[insert_at - 1].strip():
            insert_at -= 1
        lines.insert(insert_at, link_line)
        end += 1
        link_index = insert_at
    else:
        lines[link_index] = link_line

    if status_index is None:
        lines.insert(link_index + 1, STORE_STATUS_LINE)
    else:
        lines[status_index] = STORE_STATUS_LINE

    output = newline.join(lines)
    if has_trailing and not output.endswith(newline):
        output += newline
    write_text_atomic(path, output)


def write_plan(rows: list[PlanRow], validation_errors: list[str]) -> None:
    PLAN_CSV.parent.mkdir(parents=True, exist_ok=True)
    with PLAN_CSV.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=[
                "id",
                "title",
                "source_original",
                "current_payment_status",
                "current_stripe_payment_link",
                "new_payment_status",
                "new_stripe_payment_link",
                "action",
                "validation_result",
                "notes",
            ],
        )
        writer.writeheader()
        for row in rows:
            writer.writerow(
                {
                    "id": row.item_id,
                    "title": row.title,
                    "source_original": row.source_original,
                    "current_payment_status": row.current_payment_status,
                    "current_stripe_payment_link": row.current_stripe_payment_link,
                    "new_payment_status": row.new_payment_status,
                    "new_stripe_payment_link": row.new_stripe_payment_link,
                    "action": row.action,
                    "validation_result": row.validation_result,
                    "notes": row.notes,
                }
            )

    counts = Counter(row.action for row in rows)
    lines = [
        "# Stripe Live Results Writeback Plan",
        "",
        "## Purpose",
        "",
        "Plan for writing Stripe live Payment Link URLs from the live execution result back to source `ORIGINAL.md` files.",
        "",
        "## Input",
        "",
        "- `tools/reports/stripe_payment_link_live_execution_result.json`",
        "- `tools/reports/stripe_payment_link_live_execution_state.json`",
        "",
        "## Input Validation",
        "",
        f"- validation errors: {len(validation_errors)}",
    ]
    if validation_errors:
        lines.extend(f"- {error}" for error in validation_errors)
    lines.extend(
        [
            "",
            "## Summary",
            "",
            f"- result items: {len(rows)}",
            f"- source originals found: {sum(1 for row in rows if row.absolute_path.exists())}",
            f"- update: {counts.get('update', 0)}",
            f"- already_same: {counts.get('already_same', 0)}",
            f"- conflicts: {counts.get('stop_conflict', 0)}",
            f"- missing: {counts.get('stop_missing', 0)}",
            f"- errors: {len(validation_errors)}",
            "",
            "## Planned Updates",
            "",
            "| id | source_original | action | payment_link |",
            "|---|---|---|---|",
        ]
    )
    for row in rows:
        lines.append(f"| {row.item_id} | `{row.source_original}` | {row.action} | {row.new_stripe_payment_link} |")
    lines.extend(
        [
            "",
            "## Excluded",
            "",
            "- Existing Stripe-ready items outside the 45 live result items are not changed.",
            "- `DAKE_Pack_Document` and `DAKE_Pack_Memo` remain BOOTH-only.",
            "- `video_shorts_cut` remains preparing.",
            "",
            "## Stop Conditions",
            "",
            "- Any input validation error.",
            "- Missing source original or Store section.",
            "- Existing different Stripe Payment Link in an ORIGINAL.",
            "- `--apply` without `--confirm-count 45`.",
            "",
            "## Not Done In This Phase",
            "",
            "- No Stripe API call.",
            "- No Stripe Secret Key read.",
            "- No Stripe Product, Price, or Payment Link creation/update/deletion.",
            "- No Pack Stripe writeback.",
        ]
    )
    write_text_atomic(PLAN_MD, "\n".join(lines) + "\n")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Write Stripe live Payment Link results back to ORIGINAL.md files.")
    parser.add_argument("--result-json", type=Path, default=RESULT_JSON)
    parser.add_argument("--state-json", type=Path, default=STATE_JSON)
    parser.add_argument("--apply", action="store_true", help="Write changes to ORIGINAL.md files.")
    parser.add_argument("--confirm-count", type=int, default=None, help="Required with --apply; must be 45.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    result = read_json(args.result_json)
    state = read_json(args.state_json)
    validation_errors = validate_inputs(result, state)
    items = result.get("items") if isinstance(result.get("items"), list) else []
    rows = [plan_for_item(item) for item in items if isinstance(item, dict)] if not validation_errors else []
    write_plan(rows, validation_errors)

    counts = Counter(row.action for row in rows)
    print(f"mode={'apply' if args.apply else 'dry-run'}")
    print(f"result_items={len(items)}")
    print(f"source_originals_found={sum(1 for row in rows if row.absolute_path.exists())}")
    print(f"update={counts.get('update', 0)}")
    print(f"already_same={counts.get('already_same', 0)}")
    print(f"conflicts={counts.get('stop_conflict', 0)}")
    print(f"missing={counts.get('stop_missing', 0)}")
    print(f"errors={len(validation_errors)}")
    print(f"plan_csv={PLAN_CSV.relative_to(ROOT).as_posix()}")
    print(f"plan_md={PLAN_MD.relative_to(ROOT).as_posix()}")
    print("stripe_api_called=no")
    print("secret_read=no")

    stop_actions = {"stop_conflict", "stop_missing"}
    if validation_errors or any(row.action in stop_actions for row in rows):
        return 1
    if len(rows) != EXPECTED_COUNT:
        print(f"expected {EXPECTED_COUNT} rows, got {len(rows)}")
        return 1

    if not args.apply:
        return 0
    if args.confirm_count != EXPECTED_COUNT:
        print(f"--confirm-count must be {EXPECTED_COUNT}")
        return 1

    for row in rows:
        if row.action == "update":
            update_original(row.absolute_path, row.new_stripe_payment_link)
    print(f"applied_updates={counts.get('update', 0)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
