from __future__ import annotations

import argparse
import json
import os
import re
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
ARTIFACT_ROOT = ROOT / "tools" / "reports" / "release_artifacts"
JST = timezone(timedelta(hours=9))
BUY_STRIPE_PREFIX = "https://buy.stripe.com/"
REQUIRED_CONFIRMATION = "FINALIZE LIVE RELEASE"


class SafetyStop(RuntimeError):
    pass


def now_jst() -> str:
    return datetime.now(JST).isoformat(timespec="seconds")


def validate_product_id(product_id: str) -> None:
    if not product_id or not product_id.strip():
        raise SafetyStop("product_id is required")
    if product_id != product_id.strip():
        raise SafetyStop("product_id must not contain surrounding whitespace")
    if any(part in product_id for part in ["..", "/", "\\"]):
        raise SafetyStop("product_id must not contain path separators or traversal")
    if Path(product_id).is_absolute():
        raise SafetyStop("product_id must not be an absolute path")
    if any(ord(char) < 32 for char in product_id):
        raise SafetyStop("product_id must not contain control characters")


def read_json(path: Path) -> dict[str, Any]:
    if not path.exists():
        raise SafetyStop(f"missing file: {relative_to_root(path)}")
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise SafetyStop(f"JSON root must be an object: {relative_to_root(path)}")
    return data


def write_json_atomic(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(f"{path.name}.{os.getpid()}.tmp")
    with temp.open("w", encoding="utf-8", newline="\n") as handle:
        handle.write(json.dumps(data, ensure_ascii=False, indent=2) + "\n")
        handle.flush()
        os.fsync(handle.fileno())
    os.replace(temp, path)


def write_text_atomic(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(f"{path.name}.{os.getpid()}.tmp")
    with temp.open("w", encoding="utf-8", newline="\n") as handle:
        handle.write(text)
        handle.flush()
        os.fsync(handle.fileno())
    os.replace(temp, path)


def relative_to_root(path: Path) -> str:
    try:
        return path.resolve().relative_to(ROOT).as_posix()
    except ValueError:
        return str(path)


def original_paths() -> list[Path]:
    return list((ROOT / "01_apps").glob("**/ORIGINAL.md")) + list((ROOT / "04_packs").glob("**/ORIGINAL.md"))


def normalize_value(value: str) -> str:
    text = value.strip()
    if len(text) >= 2 and text.startswith("`") and text.endswith("`"):
        text = text[1:-1].strip()
    return text


def metadata_line(text: str, key: str) -> str:
    match = re.search(rf"^\s*[-*]\s*{re.escape(key)}\s*:\s*(.+?)\s*$", text, re.MULTILINE)
    return normalize_value(match.group(1)) if match else ""


def parse_original(path: Path) -> dict[str, Any]:
    text = path.read_text(encoding="utf-8")
    relative = relative_to_root(path)
    item_id = metadata_line(text, "pack_id") if relative.startswith("04_packs/") else path.parent.name
    if not item_id:
        item_id = path.parent.name
    return {
        "id": item_id,
        "source_original": relative,
        "path": path,
        "text": text,
        "payment_status": metadata_line(text, "payment_status"),
        "stripe_payment_link": metadata_line(text, "stripe_payment_link"),
    }


def discover_original_items() -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    duplicates: dict[str, list[str]] = {}
    for path in original_paths():
        item = parse_original(path)
        item_id = str(item["id"])
        if item_id in result:
            duplicates.setdefault(item_id, [result[item_id]["source_original"]]).append(item["source_original"])
        result[item_id] = item
    if duplicates:
        detail = "; ".join(f"{item_id}: {', '.join(paths)}" for item_id, paths in sorted(duplicates.items()))
        raise SafetyStop("duplicate product id in ORIGINAL.md files: " + detail)
    return result


def artifact_dir(product_id: str) -> Path:
    return ARTIFACT_ROOT / product_id


def result_path(product_id: str) -> Path:
    return artifact_dir(product_id) / "stripe_release_result.json"


def state_path(product_id: str) -> Path:
    return artifact_dir(product_id) / "stripe_release_state.json"


def checkout_review_path(product_id: str) -> Path:
    return artifact_dir(product_id) / "checkout_review.json"


def plan_path(product_id: str) -> Path:
    return artifact_dir(product_id) / "finalize_release_plan.json"


def plan_md_path(product_id: str) -> Path:
    return artifact_dir(product_id) / "finalize_release_plan.md"


def result_out_path(product_id: str) -> Path:
    return artifact_dir(product_id) / "finalize_release_result.json"


def result_out_md_path(product_id: str) -> Path:
    return artifact_dir(product_id) / "finalize_release_result.md"


def is_unset_link(value: str) -> bool:
    compact = (value or "").strip().lower()
    return compact in {"", "not set", "none", "null", "-", "unset"}


def validate_stripe_result(product_id: str, result: dict[str, Any], state: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    if result.get("product_id") != product_id:
        errors.append("result product_id does not match target")
    if state.get("product_id") != product_id:
        errors.append("state product_id does not match target")
    if result.get("livemode") is not True:
        errors.append("result livemode must be true")
    if state.get("livemode") is not True:
        errors.append("state livemode must be true")
    if state.get("status") != "completed":
        errors.append("state status must be completed")
    if result.get("errors") not in ([], None):
        errors.append("result errors must be empty")

    for key in ["product_id_on_stripe", "price_id", "payment_link_id", "payment_link_url"]:
        if not result.get(key):
            errors.append(f"result {key} is missing")
        if not state.get(key):
            errors.append(f"state {key} is missing")
        if result.get(key) and state.get(key) and result.get(key) != state.get(key):
            errors.append(f"result/state mismatch: {key}")

    payment_link_url = str(result.get("payment_link_url") or "")
    if not payment_link_url.startswith(BUY_STRIPE_PREFIX):
        errors.append("payment_link_url must start with https://buy.stripe.com/")
    if "test_" in payment_link_url:
        errors.append("payment_link_url must not contain test_")
    return errors


def validate_checkout_review(product_id: str, review: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    expected = {
        "product_id": product_id,
        "review_status": "passed",
        "livemode": True,
        "actual_payment_completed": False,
        "email_field_present": True,
        "email_required": True,
        "quantity_change_ui": False,
        "shipping_address_required": False,
        "manual_delivery_notice_visible": True,
        "next_business_day_notice_visible": True,
        "test_url_detected": False,
        "private_download_url_exposed": False,
        "local_path_exposed": False,
    }
    for key, value in expected.items():
        if review.get(key) != value:
            errors.append(f"checkout_review {key} must be {value!r}")
    if int(review.get("console_errors") or 0) != 0:
        errors.append("checkout_review console_errors must be 0")
    return errors


def validate_unique_live_results() -> list[str]:
    errors: list[str] = []
    seen: dict[str, dict[str, str]] = {
        "product_id_on_stripe": {},
        "price_id": {},
        "payment_link_id": {},
        "payment_link_url": {},
    }
    for path in ARTIFACT_ROOT.glob("*/stripe_release_result.json"):
        try:
            data = read_json(path)
        except SafetyStop:
            continue
        item_id = str(data.get("product_id") or path.parent.name)
        for key, values in seen.items():
            value = str(data.get(key) or "")
            if not value:
                continue
            if value in values and values[value] != item_id:
                errors.append(f"duplicate {key}: {value} used by {values[value]} and {item_id}")
            values[value] = item_id
    return errors


def replace_single_metadata_line(text: str, key: str, value: str) -> str:
    pattern = re.compile(rf"^(\s*[-*]\s*{re.escape(key)}\s*:\s*)(.*?)\s*$", re.MULTILINE)
    matches = list(pattern.finditer(text))
    if len(matches) != 1:
        raise SafetyStop(f"expected exactly one {key} line, found {len(matches)}")
    return pattern.sub(rf"\g<1>{value}", text, count=1)


def build_updated_original(text: str, payment_link_url: str) -> str:
    updated = replace_single_metadata_line(text, "stripe_payment_link", payment_link_url)
    updated = replace_single_metadata_line(updated, "payment_status", "stripe_ready")
    return updated


def build_plan(product_id: str) -> dict[str, Any]:
    validate_product_id(product_id)
    originals = discover_original_items()
    if product_id not in originals:
        raise SafetyStop(f"Product id not found in ORIGINAL.md files: {product_id}")
    original = originals[product_id]
    result = read_json(result_path(product_id))
    state = read_json(state_path(product_id))
    checkout = read_json(checkout_review_path(product_id))

    result_errors = validate_stripe_result(product_id, result, state)
    checkout_errors = validate_checkout_review(product_id, checkout)
    uniqueness_errors = validate_unique_live_results()
    errors = result_errors + checkout_errors + uniqueness_errors

    payment_status_before = original["payment_status"]
    stripe_payment_link_before = original["stripe_payment_link"]
    payment_link_after = str(result.get("payment_link_url") or "")
    already_same = (
        payment_status_before == "stripe_ready"
        and stripe_payment_link_before == payment_link_after
        and not errors
    )
    can_update = (
        payment_status_before == "booth_only"
        and is_unset_link(stripe_payment_link_before)
        and not errors
    )

    if errors:
        action = "stop_invalid_result"
    elif already_same:
        action = "already_same"
    elif can_update:
        action = "update"
    else:
        action = "stop_conflict"
        errors.append("ORIGINAL.md payment state is not booth_only/not set and not already_same")

    return {
        "created_at": now_jst(),
        "product_id": product_id,
        "source_original": original["source_original"],
        "result_file": relative_to_root(result_path(product_id)),
        "state_file": relative_to_root(state_path(product_id)),
        "checkout_review_file": relative_to_root(checkout_review_path(product_id)),
        "payment_status_before": payment_status_before,
        "stripe_payment_link_before": "" if is_unset_link(stripe_payment_link_before) else stripe_payment_link_before,
        "payment_status_after": "stripe_ready",
        "stripe_payment_link_after": payment_link_after,
        "result_validation": "passed" if not result_errors else "failed",
        "state_validation": "passed" if not result_errors else "failed",
        "checkout_validation": "passed" if not checkout_errors else "failed",
        "uniqueness_validation": "passed" if not uniqueness_errors else "failed",
        "action": action,
        "errors": errors,
        "dry_run": True,
    }


def render_plan_md(plan: dict[str, Any]) -> str:
    errors = "\n".join(f"- {error}" for error in plan["errors"]) or "- none"
    return f"""# Finalize Product Release Plan

## Summary

- product_id: {plan['product_id']}
- source_original: {plan['source_original']}
- action: {plan['action']}
- errors: {len(plan['errors'])}

## Before

- payment_status: {plan['payment_status_before']}
- stripe_payment_link: {plan['stripe_payment_link_before'] or 'not set'}

## After

- payment_status: {plan['payment_status_after']}
- stripe_payment_link: {plan['stripe_payment_link_after']}

## Validation

- result: {plan['result_validation']}
- state: {plan['state_validation']}
- checkout: {plan['checkout_validation']}
- uniqueness: {plan['uniqueness_validation']}

## Errors

{errors}
"""


def render_result_md(result: dict[str, Any]) -> str:
    errors = "\n".join(f"- {error}" for error in result["errors"]) or "- none"
    return f"""# Finalize Product Release Result

## Summary

- product_id: {result['product_id']}
- source_original: {result['source_original']}
- action: {result['action']}
- applied: {result['applied']}
- errors: {len(result['errors'])}

## ORIGINAL.md

- payment_status: {result['payment_status_after']}
- stripe_payment_link: {result['stripe_payment_link_after']}

## Errors

{errors}
"""


def apply_plan(plan: dict[str, Any]) -> dict[str, Any]:
    if plan["action"] == "already_same":
        return {**plan, "applied": False, "dry_run": False, "completed_at": now_jst()}
    if plan["action"] != "update":
        raise SafetyStop("cannot apply plan with action=" + plan["action"])
    source_path = ROOT / plan["source_original"]
    original_text = source_path.read_text(encoding="utf-8")
    updated_text = build_updated_original(original_text, plan["stripe_payment_link_after"])
    write_text_atomic(source_path, updated_text)
    verified = parse_original(source_path)
    errors: list[str] = []
    if verified["payment_status"] != "stripe_ready":
        errors.append("post-apply payment_status is not stripe_ready")
    if verified["stripe_payment_link"] != plan["stripe_payment_link_after"]:
        errors.append("post-apply stripe_payment_link does not match result URL")
    return {
        **plan,
        "payment_status_after_verified": verified["payment_status"],
        "stripe_payment_link_after_verified": verified["stripe_payment_link"],
        "applied": not errors,
        "dry_run": False,
        "completed_at": now_jst(),
        "errors": errors,
        "action": "updated" if not errors else "stop_conflict",
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Finalize a DAKE product release after live Stripe Checkout review.")
    parser.add_argument("product_id", nargs="?", help="DAKE product id")
    parser.add_argument("--apply", action="store_true", help="Write stripe_payment_link and payment_status to ORIGINAL.md.")
    parser.add_argument("--confirm-product-id", default="", help="Must exactly match product_id when --apply is used.")
    parser.add_argument("--confirm-checkout-reviewed", action="store_true", help="Confirm live Checkout review was completed.")
    parser.add_argument("--confirmation-text", default="", help="Must be FINALIZE LIVE RELEASE <product_id> when --apply is used.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if not args.product_id:
        print("usage: python tools\\finalize_product_release.py <product_id> [--apply]")
        return 2
    try:
        plan = build_plan(args.product_id)
        write_json_atomic(plan_path(args.product_id), plan)
        write_text_atomic(plan_md_path(args.product_id), render_plan_md(plan))
        print(f"product_id={plan['product_id']}")
        print(f"action={plan['action']}")
        print(f"errors={len(plan['errors'])}")
        print(f"plan_json={plan_path(args.product_id)}")
        print(f"plan_md={plan_md_path(args.product_id)}")
        if plan["errors"]:
            for error in plan["errors"]:
                print(f"- {error}")
            return 1
        if not args.apply:
            return 0
        expected_text = f"{REQUIRED_CONFIRMATION} {args.product_id}"
        if args.confirm_product_id != args.product_id:
            raise SafetyStop(f"--confirm-product-id must be {args.product_id}")
        if not args.confirm_checkout_reviewed:
            raise SafetyStop("--confirm-checkout-reviewed is required")
        if args.confirmation_text != expected_text:
            raise SafetyStop(f'--confirmation-text must be "{expected_text}"')
        final_result = apply_plan(plan)
        write_json_atomic(result_out_path(args.product_id), final_result)
        write_text_atomic(result_out_md_path(args.product_id), render_result_md(final_result))
        print(f"applied={final_result['applied']}")
        print(f"result_json={result_out_path(args.product_id)}")
        print(f"result_md={result_out_md_path(args.product_id)}")
        return 0 if not final_result["errors"] else 1
    except SafetyStop as exc:
        print(str(exc))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
