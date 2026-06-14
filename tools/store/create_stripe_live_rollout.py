from __future__ import annotations

import argparse
import copy
import csv
import hashlib
import json
import os
import re
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Callable


ROOT = Path(__file__).resolve().parents[2]
DEFAULT_PAYLOAD_JSON = ROOT / "tools" / "reports" / "stripe_payment_link_live_rollout_payloads.json"
DEFAULT_STATE_JSON = ROOT / "tools" / "reports" / "stripe_payment_link_live_execution_state.json"
DEFAULT_RESULT_JSON = ROOT / "tools" / "reports" / "stripe_payment_link_live_execution_result.json"
DEFAULT_RESULT_CSV = ROOT / "tools" / "reports" / "stripe_payment_link_live_execution_result.csv"
DEFAULT_RESULT_MD = ROOT / "tools" / "reports" / "stripe_payment_link_live_execution_result.md"

EXPECTED_COUNT = 45
CONFIRMATION_TEXT = "CREATE 45 DAKE LIVE PAYMENT LINKS"
GAME_IDS = {"game_alien_road", "game_diver_catch"}
SECRET_PATTERN = re.compile(r"sk_(?:test|live)_[A-Za-z0-9]{16,}|whsec_[A-Za-z0-9]{16,}")
JST = timezone(timedelta(hours=9))


class SafetyStop(RuntimeError):
    pass


def now_jst() -> str:
    return datetime.now(JST).isoformat(timespec="seconds")


def read_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def canonical_json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def sha256_payload(value: Any) -> str:
    return hashlib.sha256(canonical_json(value).encode("utf-8")).hexdigest()


def object_to_dict(obj: Any) -> Any:
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


def contains_secret_like_value(value: Any) -> bool:
    return bool(SECRET_PATTERN.search(canonical_json(value)))


def validate_payload(payload: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    if payload.get("dry_run") is not True:
        errors.append("payload dry_run must be true")
    if payload.get("candidate_count") != EXPECTED_COUNT:
        errors.append(f"candidate_count must be {EXPECTED_COUNT}")
    if payload.get("count") != EXPECTED_COUNT:
        errors.append(f"count must be {EXPECTED_COUNT}")
    if payload.get("errors"):
        errors.append("payload contains upstream errors")
    items = payload.get("items")
    if not isinstance(items, list):
        errors.append("payload items must be a list")
        return errors
    if len(items) != EXPECTED_COUNT:
        errors.append(f"items length must be {EXPECTED_COUNT}")

    seen_ids: set[str] = set()
    idempotency_keys: list[str] = []
    for index, item in enumerate(items):
        label = str(item.get("id") or f"row {index + 1}")
        item_id = str(item.get("id") or "")
        if not item_id:
            errors.append(f"{label}: id is missing")
        if item_id in seen_ids:
            errors.append(f"{label}: duplicate dake_item_id")
        seen_ids.add(item_id)
        if item.get("live_dry_run_ready") != "yes":
            errors.append(f"{label}: live_dry_run_ready must be yes")
        if not item.get("title"):
            errors.append(f"{label}: title is missing")
        if not isinstance(item.get("price"), int) or item.get("price") <= 0:
            errors.append(f"{label}: price must be a positive integer")
        if item.get("currency") != "jpy":
            errors.append(f"{label}: currency must be jpy")
        for field in ["source_original", "booth_url", "github_release_url", "tax_code_candidate"]:
            if not item.get(field):
                errors.append(f"{label}: {field} is missing")
        for payload_key, hash_key in [
            ("product_payload", "product_payload_sha256"),
            ("price_payload", "price_payload_sha256"),
            ("payment_link_payload", "payment_link_payload_sha256"),
        ]:
            payload_part = item.get(payload_key)
            if not isinstance(payload_part, dict):
                errors.append(f"{label}: {payload_key} must be an object")
                continue
            actual_hash = sha256_payload(payload_part)
            expected_hash = item.get(hash_key)
            if not expected_hash:
                errors.append(f"{label}: {hash_key} is missing")
            elif actual_hash != expected_hash:
                errors.append(f"{label}: {hash_key} mismatch")
        for key_field in [
            "product_idempotency_key",
            "price_idempotency_key",
            "payment_link_idempotency_key",
        ]:
            key_value = item.get(key_field)
            if not key_value:
                errors.append(f"{label}: {key_field} is missing")
            else:
                idempotency_keys.append(str(key_value))
        product_payload = item.get("product_payload") or {}
        price_payload = item.get("price_payload") or {}
        payment_link_payload = item.get("payment_link_payload") or {}
        if product_payload.get("tax_code") != item.get("tax_code_candidate"):
            errors.append(f"{label}: product tax_code does not match candidate")
        if price_payload.get("currency") != "jpy":
            errors.append(f"{label}: price payload currency must be jpy")
        if price_payload.get("unit_amount") != item.get("price"):
            errors.append(f"{label}: price payload unit_amount mismatch")
        if payment_link_payload.get("line_items", [{}])[0].get("quantity") != 1:
            errors.append(f"{label}: payment link quantity must be 1")
        if contains_secret_like_value(item):
            errors.append(f"{label}: secret-like value detected in payload")
    duplicate_keys = sorted({key for key in idempotency_keys if idempotency_keys.count(key) > 1})
    if duplicate_keys:
        errors.append("duplicate idempotency keys: " + ", ".join(duplicate_keys))
    return errors


def summarize_items(items: list[dict[str, Any]]) -> dict[str, int]:
    game_apps = sum(item.get("id") in GAME_IDS for item in items)
    return {
        "candidate_count": len(items),
        "normal_apps": len(items) - game_apps,
        "game_apps": game_apps,
        "product_create_planned": len(items),
        "price_create_planned": len(items),
        "payment_link_create_planned": len(items),
    }


def print_plan(payload: dict[str, Any], validation_errors: list[str]) -> int:
    items = payload.get("items", [])
    summary = summarize_items(items)
    ready_yes = sum(item.get("live_dry_run_ready") == "yes" for item in items)
    print("mode=dry-run")
    print(f"candidate_count={summary['candidate_count']}")
    print(f"normal_apps={summary['normal_apps']}")
    print(f"game_apps={summary['game_apps']}")
    print(f"product_create_planned={summary['product_create_planned']}")
    print(f"price_create_planned={summary['price_create_planned']}")
    print(f"payment_link_create_planned={summary['payment_link_create_planned']}")
    print(f"live_dry_run_ready_yes={ready_yes}")
    print(f"errors={len(validation_errors)}")
    print("tax_confirmation_required=yes")
    print("live_api_called=no")
    print("secret_read=no")
    print("This confirms configured tax candidates only.")
    print("This is not a tax or legal determination.")
    for item in items:
        print(
            "item "
            f"id={item['id']} "
            f"title={item['title']} "
            f"price={item['price']} "
            f"tax_code_candidate={item['tax_code_candidate']} "
            f"product_key={item['product_idempotency_key']} "
            f"price_key={item['price_idempotency_key']} "
            f"link_key={item['payment_link_idempotency_key']}"
        )
    if validation_errors:
        print("validation_errors:")
        for error in validation_errors:
            print(f"- {error}")
        return 1
    return 0


def validate_execute_gates(args: argparse.Namespace) -> None:
    if args.confirm_count != EXPECTED_COUNT:
        raise SafetyStop(f"--confirm-count must be {EXPECTED_COUNT}")
    if not args.confirm_tax_codes:
        raise SafetyStop("--confirm-tax-codes is required")
    if args.confirmation_text != CONFIRMATION_TEXT:
        raise SafetyStop(f'--confirmation-text must be "{CONFIRMATION_TEXT}"')


def validate_live_secret_key() -> str:
    value = os.environ.get("STRIPE_SECRET_KEY", "").strip()
    if not value:
        raise SafetyStop("STRIPE_SECRET_KEY is not set")
    if value.startswith("sk_test_"):
        raise SafetyStop("Refusing to run: test mode secret keys are not allowed for live execution")
    if not value.startswith("sk_live_"):
        raise SafetyStop("Refusing to run: STRIPE_SECRET_KEY must start with sk_live_")
    return value


def import_stripe_module() -> Any:
    try:
        import stripe  # type: ignore[import-not-found]
    except ImportError as exc:
        raise SafetyStop("stripe Python SDK is not installed. Install with: pip install stripe") from exc
    return stripe


def initial_state(items: list[dict[str, Any]]) -> dict[str, Any]:
    return {
        "mode": "live",
        "started_at": now_jst(),
        "updated_at": now_jst(),
        "expected_count": EXPECTED_COUNT,
        "status": "in_progress",
        "items": [
            {
                "id": item["id"],
                "title": item["title"],
                "status": "pending",
                "product_id": None,
                "price_id": None,
                "payment_link_id": None,
                "payment_link_url": None,
                "existing_product_id": None,
                "existing_price_ids": [],
                "existing_payment_link_ids": [],
                "manual_resolution_required": False,
                "error": None,
            }
            for item in items
        ],
    }


def load_or_create_state(path: Path, items: list[dict[str, Any]], resume: bool) -> dict[str, Any]:
    if path.exists():
        if not resume:
            raise SafetyStop(f"state file exists; pass --resume to inspect it: {path}")
        state = read_json(path)
        if state.get("status") == "completed":
            raise SafetyStop("state is already completed; refusing to run")
        return state
    if resume:
        raise SafetyStop(f"--resume was passed but state file does not exist: {path}")
    state = initial_state(items)
    write_json(path, state)
    return state


def save_state(path: Path, state: dict[str, Any], status: str | None = None) -> None:
    if status:
        state["status"] = status
    state["updated_at"] = now_jst()
    write_json(path, state)


def state_item(state: dict[str, Any], item_id: str) -> dict[str, Any]:
    for item in state["items"]:
        if item["id"] == item_id:
            return item
    raise SafetyStop(f"state is missing item {item_id}")


def create_with_idempotency(create_func: Callable[..., Any], payload: dict[str, Any], idempotency_key: str) -> dict[str, Any]:
    params = copy.deepcopy(payload)
    params["idempotency_key"] = idempotency_key
    return object_to_dict(create_func(**params))


def find_existing_products(stripe: Any) -> dict[str, list[dict[str, Any]]]:
    by_item_id: dict[str, list[dict[str, Any]]] = {}
    for product in stripe.Product.list(limit=100).auto_paging_iter():
        data = object_to_dict(product)
        item_id = str((data.get("metadata") or {}).get("dake_item_id") or "")
        if item_id:
            by_item_id.setdefault(item_id, []).append(data)
    return by_item_id


def find_prices_for_product(stripe: Any, product_id: str) -> list[str]:
    price_ids: list[str] = []
    for price in stripe.Price.list(product=product_id, active=True, limit=100).auto_paging_iter():
        data = object_to_dict(price)
        if data.get("id"):
            price_ids.append(str(data["id"]))
    return price_ids


def find_payment_links_for_prices(stripe: Any, price_ids: set[str]) -> list[str]:
    if not price_ids:
        return []
    found: list[str] = []
    for link in stripe.PaymentLink.list(limit=100).auto_paging_iter():
        link_data = object_to_dict(link)
        link_id = link_data.get("id")
        if not link_id:
            continue
        line_items = stripe.PaymentLink.list_line_items(str(link_id), limit=100)
        for line_item in line_items.auto_paging_iter():
            line_data = object_to_dict(line_item)
            price = line_data.get("price") or {}
            price_id = price.get("id") if isinstance(price, dict) else price
            if price_id in price_ids:
                found.append(str(link_id))
                break
    return found


def write_result_files(result_json: Path, result_csv: Path, result_md: Path, state: dict[str, Any]) -> None:
    completed = [item for item in state["items"] if item.get("status") == "completed"]
    result = {
        "mode": "live",
        "created_at": now_jst(),
        "count": len(completed),
        "items": completed,
        "errors": [item for item in state["items"] if item.get("error")],
    }
    write_json(result_json, result)
    with result_csv.open("w", encoding="utf-8", newline="") as handle:
        fieldnames = [
            "id",
            "title",
            "product_id",
            "price_id",
            "payment_link_id",
            "payment_link_url",
            "livemode",
            "status",
            "created_at",
            "metadata",
        ]
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for item in completed:
            writer.writerow(
                {
                    "id": item.get("id"),
                    "title": item.get("title"),
                    "product_id": item.get("product_id"),
                    "price_id": item.get("price_id"),
                    "payment_link_id": item.get("payment_link_id"),
                    "payment_link_url": item.get("payment_link_url"),
                    "livemode": "true",
                    "status": item.get("status"),
                    "created_at": result["created_at"],
                    "metadata": json.dumps(item.get("metadata") or {}, ensure_ascii=False, sort_keys=True),
                }
            )
    rows = "\n".join(
        f"| {item.get('id')} | {item.get('title')} | {item.get('product_id')} | {item.get('price_id')} | {item.get('payment_link_id')} | {item.get('payment_link_url')} |"
        for item in completed
    )
    if not rows:
        rows = "| - | - | - | - | - | - |"
    result_md.write_text(
        f"""# Stripe Payment Link Live Execution Result

## Summary

- mode: live
- completed: {len(completed)}
- errors: {len(result['errors'])}
- Secret saved: no

## Items

| id | title | product_id | price_id | payment_link_id | payment_link_url |
|---|---|---|---|---|---|
{rows}
""",
        encoding="utf-8",
        newline="\n",
    )


def execute_live(args: argparse.Namespace, payload: dict[str, Any], validation_errors: list[str]) -> int:
    if validation_errors:
        raise SafetyStop("payload validation failed; refusing live execution")
    validate_execute_gates(args)
    secret_key = validate_live_secret_key()
    stripe = import_stripe_module()
    stripe.api_key = secret_key

    items = payload["items"]
    state = load_or_create_state(args.state_json, items, args.resume)
    existing_products = find_existing_products(stripe)

    for item in items:
        entry = state_item(state, item["id"])
        if entry.get("status") == "completed":
            continue
        if entry.get("status") in {"product_created", "price_created"}:
            raise SafetyStop(f"{item['id']}: partial state requires manual reconciliation before resume")
        try:
            matches = existing_products.get(item["id"], [])
            if len(matches) > 1:
                entry["status"] = "failed"
                entry["error"] = "multiple existing live Products with matching metadata.dake_item_id"
                save_state(args.state_json, state, "failed")
                return 1
            if len(matches) == 1:
                product_id = matches[0].get("id")
                price_ids = find_prices_for_product(stripe, str(product_id))
                link_ids = find_payment_links_for_prices(stripe, set(price_ids))
                entry.update(
                    {
                        "status": "existing_object_detected",
                        "existing_product_id": product_id,
                        "existing_price_ids": price_ids,
                        "existing_payment_link_ids": link_ids,
                        "manual_resolution_required": True,
                        "error": "existing live Product detected; manual resolution required",
                    }
                )
                save_state(args.state_json, state, "failed")
                return 1

            product = create_with_idempotency(
                stripe.Product.create,
                item["product_payload"],
                item["product_idempotency_key"],
            )
            entry["product_id"] = product["id"]
            entry["status"] = "product_created"
            save_state(args.state_json, state)

            price_payload = copy.deepcopy(item["price_payload"])
            price_payload["product"] = product["id"]
            price = create_with_idempotency(
                stripe.Price.create,
                price_payload,
                item["price_idempotency_key"],
            )
            entry["price_id"] = price["id"]
            entry["status"] = "price_created"
            save_state(args.state_json, state)

            payment_link_payload = copy.deepcopy(item["payment_link_payload"])
            payment_link_payload["line_items"] = [
                {
                    **line,
                    "price": price["id"],
                }
                for line in payment_link_payload.get("line_items", [])
            ]
            payment_link = create_with_idempotency(
                stripe.PaymentLink.create,
                payment_link_payload,
                item["payment_link_idempotency_key"],
            )
            entry["payment_link_id"] = payment_link["id"]
            entry["payment_link_url"] = payment_link.get("url")
            entry["metadata"] = item["product_payload"].get("metadata", {})
            entry["status"] = "completed"
            save_state(args.state_json, state)
        except Exception as exc:
            entry["status"] = "failed"
            entry["error"] = str(exc)
            save_state(args.state_json, state, "failed")
            return 1

    save_state(args.state_json, state, "completed")
    write_result_files(args.result_json, args.result_csv, args.result_md, state)
    return 0


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Prepare or execute DAKE Stripe live Payment Link rollout.")
    parser.add_argument("--payload-json", type=Path, default=DEFAULT_PAYLOAD_JSON)
    parser.add_argument("--state-json", type=Path, default=DEFAULT_STATE_JSON)
    parser.add_argument("--result-json", type=Path, default=DEFAULT_RESULT_JSON)
    parser.add_argument("--result-csv", type=Path, default=DEFAULT_RESULT_CSV)
    parser.add_argument("--result-md", type=Path, default=DEFAULT_RESULT_MD)
    parser.add_argument("--execute-live", action="store_true")
    parser.add_argument("--confirm-count", type=int)
    parser.add_argument("--confirm-tax-codes", action="store_true")
    parser.add_argument("--confirmation-text", default="")
    parser.add_argument("--resume", action="store_true")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    payload = read_json(args.payload_json)
    validation_errors = validate_payload(payload)
    if not args.execute_live:
        return print_plan(payload, validation_errors)
    return execute_live(args, payload, validation_errors)


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except SafetyStop as exc:
        print(str(exc))
        raise SystemExit(1) from None
