from __future__ import annotations

import argparse
import copy
from typing import Any

from store.stripe_release_core import (
    SafetyStop,
    build_release_payload,
    canonical_json,
    import_stripe_module,
    now_jst,
    object_to_dict,
    read_json,
    relative_to_root,
    result_json_path,
    sha256_payload,
    state_json_path,
    validate_live_secret_key,
    write_json_atomic,
    write_text_atomic,
)


EXPECTED_CONFIRMATION = "UPDATE LIVE CHECKOUT NOTICE"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Update only the Checkout submit notice on an existing live Stripe Payment Link.")
    parser.add_argument("product_id", help="DAKE product id")
    parser.add_argument("--apply-live", action="store_true", help="Update the live Payment Link after strict confirmation.")
    parser.add_argument("--confirm-product-id", default="", help="Must match product_id when --apply-live is used.")
    parser.add_argument("--confirm-payment-link-id", default="", help="Must match the existing live Payment Link id.")
    parser.add_argument("--confirmation-text", default="", help="Must be UPDATE LIVE CHECKOUT NOTICE <product_id>.")
    return parser.parse_args()


def artifact_paths(product_id: str) -> dict[str, Any]:
    base = result_json_path(product_id).parent
    return {
        "plan_json": base / "checkout_notice_update_plan.json",
        "plan_md": base / "checkout_notice_update_plan.md",
        "result_json": base / "checkout_notice_update_result.json",
        "result_md": base / "checkout_notice_update_result.md",
    }


def notice_from_payload(payload: dict[str, Any]) -> str:
    custom_text = payload.get("payment_link_payload", {}).get("custom_text") or {}
    submit = custom_text.get("submit") if isinstance(custom_text, dict) else {}
    message = submit.get("message") if isinstance(submit, dict) else ""
    return str(message or "")


def validate_local_artifacts(product_id: str, payload: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any], list[str]]:
    errors: list[str] = []
    state = read_json(state_json_path(product_id))
    result = read_json(result_json_path(product_id))
    if state.get("status") != "completed":
        errors.append("state.status must be completed")
    if result.get("status") != "completed":
        errors.append("result.status must be completed")
    if state.get("livemode") is not True or result.get("livemode") is not True:
        errors.append("state/result livemode must be true")
    if result.get("errors") not in ([], None):
        errors.append("result errors must be empty")
    expected_pairs = [
        ("product_id", product_id),
        ("product_id_on_stripe", None),
        ("price_id", None),
        ("payment_link_id", None),
        ("payment_link_url", None),
    ]
    for key, expected in expected_pairs:
        if state.get(key) != result.get(key):
            errors.append(f"state/result mismatch: {key}")
        if expected is not None and result.get(key) != expected:
            errors.append(f"unexpected {key}: {result.get(key)}")
    if result.get("product_id") != product_id or state.get("product_id") != product_id:
        errors.append("state/result product_id must match target")
    if result.get("price_id") != payload.get("expected_price_id") and payload.get("expected_price_id"):
        errors.append("result price_id does not match expected dry-run value")
    if payload.get("price") != 780 and product_id == "DAKE_Pack_Mail":
        errors.append("price must be 780 for DAKE_Pack_Mail")
    if payload.get("currency") != "jpy":
        errors.append("currency must be jpy")
    if payload.get("payment_status_before") != "booth_only":
        errors.append("payment_status_before must be booth_only")
    if payload.get("stripe_payment_link_before") not in ("", None):
        errors.append("stripe_payment_link_before must be unset")
    return state, result, errors


def validate_notice(message: str) -> list[str]:
    errors: list[str] = []
    required = [
        "Windows",
        "Microsoft Outlook Classic",
        "New Outlook",
        "Web Outlook",
        "自動送信されません",
        "下書き",
        "確認してから",
        "自動ダウンロードではありません",
        "購入時に入力されたメールアドレス",
        "次営業日以内",
    ]
    if not message:
        errors.append("checkout notice message is missing")
    if len(message) > 1200:
        errors.append("checkout notice message exceeds 1200 characters")
    for term in required:
        if term not in message:
            errors.append(f"checkout notice missing required term: {term}")
    return errors


def checkout_notice_idempotency_key(product_id: str, payment_link_id: str, notice_sha256: str) -> str:
    return f"dake-update-checkout-notice-v1-{product_id}-{payment_link_id}-{notice_sha256[:12]}"


def build_plan(product_id: str) -> dict[str, Any]:
    payload = build_release_payload(product_id)
    message = notice_from_payload(payload)
    state, result, artifact_errors = validate_local_artifacts(product_id, payload)
    errors = list(payload.get("errors") or []) + artifact_errors + validate_notice(message)
    payment_link_id = str(result.get("payment_link_id") or "")
    notice_sha256 = sha256_payload(message)
    key = checkout_notice_idempotency_key(product_id, payment_link_id, notice_sha256)
    plan = {
        "mode": "dry-run",
        "product_id": product_id,
        "created_at": now_jst(),
        "payment_link_id": payment_link_id,
        "payment_link_url": result.get("payment_link_url"),
        "product_id_on_stripe": result.get("product_id_on_stripe"),
        "price_id": result.get("price_id"),
        "state_file": relative_to_root(state_json_path(product_id)),
        "result_file": relative_to_root(result_json_path(product_id)),
        "action": "update_checkout_notice",
        "new_notice": message,
        "new_notice_length": len(message),
        "new_notice_sha256": notice_sha256,
        "idempotency_key_hash": sha256_payload(key),
        "ready_for_live_update": "yes" if not errors else "no",
        "errors": errors,
        "secret_read": "no",
        "live_api_called": "no",
        "state_status": state.get("status"),
        "result_status": result.get("status"),
        "livemode": result.get("livemode"),
    }
    return plan


def render_plan(plan: dict[str, Any]) -> str:
    errors = "\n".join(f"- {error}" for error in plan["errors"]) or "- none"
    return f"""# Checkout Notice Update Plan

## Summary

- product_id: {plan['product_id']}
- payment_link_id: {plan['payment_link_id']}
- payment_link_url: {plan['payment_link_url']}
- action: {plan['action']}
- ready_for_live_update: {plan['ready_for_live_update']}
- errors: {len(plan['errors'])}
- secret_read: {plan['secret_read']}
- live_api_called: {plan['live_api_called']}

## Notice

- length: {plan['new_notice_length']}
- sha256: {plan['new_notice_sha256']}

```text
{plan['new_notice']}
```

## Errors

{errors}
"""


def payment_link_submit_message(link: dict[str, Any]) -> str:
    custom_text = link.get("custom_text") or {}
    submit = custom_text.get("submit") if isinstance(custom_text, dict) else {}
    return str(submit.get("message") or "") if isinstance(submit, dict) else ""


def retrieve_payment_link(stripe: Any, payment_link_id: str) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    link = object_to_dict(stripe.PaymentLink.retrieve(payment_link_id))
    line_items = [
        object_to_dict(item)
        for item in stripe.PaymentLink.list_line_items(payment_link_id, limit=100).auto_paging_iter()
    ]
    return link, line_items


def line_item_price_id(line_item: dict[str, Any]) -> str:
    price = line_item.get("price")
    if isinstance(price, dict):
        return str(price.get("id") or "")
    return str(price or "")


def validate_live_payment_link(link: dict[str, Any], line_items: list[dict[str, Any]], plan: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    if link.get("id") != plan["payment_link_id"]:
        errors.append("Payment Link ID mismatch")
    if link.get("livemode") is not True:
        errors.append("Payment Link livemode must be true")
    if link.get("active") is not True:
        errors.append("Payment Link active must be true")
    if link.get("url") != plan["payment_link_url"]:
        errors.append("Payment Link URL mismatch")
    metadata = link.get("metadata") or {}
    if not isinstance(metadata, dict) or metadata.get("dake_item_id") != plan["product_id"]:
        errors.append("Payment Link metadata.dake_item_id mismatch")
    if len(line_items) != 1:
        errors.append("Payment Link must have exactly one line item")
    else:
        item = line_items[0]
        if line_item_price_id(item) != plan["price_id"]:
            errors.append("Payment Link Price ID mismatch")
        if int(item.get("quantity") or 0) != 1:
            errors.append("Payment Link quantity must be 1")
    return errors


def render_result(result: dict[str, Any]) -> str:
    errors = "\n".join(f"- {error}" for error in result["errors"]) or "- none"
    return f"""# Checkout Notice Update Result

## Summary

- product_id: {result['product_id']}
- status: {result['status']}
- livemode: {result['livemode']}
- payment_link_id: {result['payment_link_id']}
- payment_link_url_before: {result['payment_link_url_before']}
- payment_link_url_after: {result['payment_link_url_after']}
- active_before: {result['active_before']}
- active_after: {result['active_after']}
- errors: {len(result['errors'])}

## Notice

- notice_sha256: {result['notice_sha256']}
- notice_length: {len(result['custom_text_after'])}
- idempotency_key_hash: {result['idempotency_key_hash']}

## Errors

{errors}
"""


def apply_live_update(plan: dict[str, Any], args: argparse.Namespace) -> dict[str, Any]:
    product_id = plan["product_id"]
    if args.confirm_product_id != product_id:
        raise SafetyStop(f"--confirm-product-id must be {product_id}")
    if args.confirm_payment_link_id != plan["payment_link_id"]:
        raise SafetyStop(f"--confirm-payment-link-id must be {plan['payment_link_id']}")
    expected_text = f"{EXPECTED_CONFIRMATION} {product_id}"
    if args.confirmation_text != expected_text:
        raise SafetyStop(f'--confirmation-text must be "{expected_text}"')
    if plan["ready_for_live_update"] != "yes":
        raise SafetyStop("dry-run plan is not ready_for_live_update=yes")

    secret_key = validate_live_secret_key()
    stripe = import_stripe_module()
    stripe.api_key = secret_key

    before, before_items = retrieve_payment_link(stripe, plan["payment_link_id"])
    preflight_errors = validate_live_payment_link(before, before_items, plan)
    if preflight_errors:
        raise SafetyStop("live Payment Link preflight failed: " + "; ".join(preflight_errors))

    before_message = payment_link_submit_message(before)
    status = "updated"
    live_api_called = "yes"
    if before_message == plan["new_notice"]:
        status = "already_same"
        live_api_called = "retrieve_only"
        updated = before
    else:
        idempotency_hash = plan["idempotency_key_hash"]
        idempotency_key = checkout_notice_idempotency_key(product_id, plan["payment_link_id"], plan["new_notice_sha256"])
        updated = object_to_dict(
            stripe.PaymentLink.modify(
                plan["payment_link_id"],
                custom_text={"submit": {"message": plan["new_notice"]}},
                idempotency_key=idempotency_key,
            )
        )
        if sha256_payload(idempotency_key) != idempotency_hash:
            raise SafetyStop("idempotency key hash mismatch")

    after, after_items = retrieve_payment_link(stripe, plan["payment_link_id"])
    post_errors = validate_live_payment_link(after, after_items, plan)
    if payment_link_submit_message(after) != plan["new_notice"]:
        post_errors.append("custom_text.submit.message was not updated to the generated notice")
    if after.get("url") != before.get("url"):
        post_errors.append("Payment Link URL changed")
    if line_item_price_id(after_items[0]) != line_item_price_id(before_items[0]):
        post_errors.append("Price ID changed")
    if int(after_items[0].get("quantity") or 0) != int(before_items[0].get("quantity") or 0):
        post_errors.append("quantity changed")

    if post_errors:
        status = "failed"

    return {
        "product_id": product_id,
        "payment_link_id": plan["payment_link_id"],
        "payment_link_url_before": before.get("url"),
        "payment_link_url_after": after.get("url"),
        "active_before": before.get("active"),
        "active_after": after.get("active"),
        "custom_text_before": before_message,
        "custom_text_after": payment_link_submit_message(after),
        "notice_sha256": plan["new_notice_sha256"],
        "idempotency_key_hash": plan["idempotency_key_hash"],
        "livemode": after.get("livemode"),
        "status": status,
        "errors": post_errors,
        "live_api_called": live_api_called,
        "stripe_object_id_unchanged": updated.get("id") == plan["payment_link_id"],
        "product_created": False,
        "price_created": False,
        "payment_link_created": False,
        "updated_at": now_jst(),
    }


def main() -> int:
    args = parse_args()
    try:
        paths = artifact_paths(args.product_id)
        plan = build_plan(args.product_id)
        write_json_atomic(paths["plan_json"], plan)
        write_text_atomic(paths["plan_md"], render_plan(plan))
        print(f"product_id={plan['product_id']}")
        print(f"payment_link_id={plan['payment_link_id']}")
        print(f"payment_link_url={plan['payment_link_url']}")
        print(f"action={plan['action']}")
        print(f"new_notice_length={plan['new_notice_length']}")
        print(f"ready_for_live_update={plan['ready_for_live_update']}")
        print(f"errors={len(plan['errors'])}")
        print(f"secret_read={plan['secret_read']}")
        print(f"live_api_called={plan['live_api_called']}")
        print(f"output_json={paths['plan_json']}")
        print(f"output_md={paths['plan_md']}")
        if plan["errors"]:
            for error in plan["errors"]:
                print(f"- {error}")
            return 1
        if not args.apply_live:
            return 0
        result = apply_live_update(plan, args)
        write_json_atomic(paths["result_json"], result)
        write_text_atomic(paths["result_md"], render_result(result))
        print(f"status={result['status']}")
        print(f"livemode={result['livemode']}")
        print(f"errors={len(result['errors'])}")
        print(f"live_api_called={result['live_api_called']}")
        print(f"output_json={paths['result_json']}")
        print(f"output_md={paths['result_md']}")
        return 0 if result["status"] in {"updated", "already_same"} and not result["errors"] else 1
    except SafetyStop as exc:
        print(str(exc))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
