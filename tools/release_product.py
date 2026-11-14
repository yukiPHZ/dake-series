"""Single normal entrypoint for DAKE Stripe Commerce.

The default is a local dry-run. ``--apply-live`` is the one explicit live
intent. This command never advances SHIP, BOOTH, Store, Cloudflare, SNS, or
Buffer work.
"""

from __future__ import annotations

import argparse

from commerce.stripe_commerce_v2 import SafetyStop, build_dry_run_plan, execute_live, safe_error_message


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Prepare one DAKE product for Stripe Commerce from ORIGINAL.md.")
    parser.add_argument("product_id", nargs="?", help="DAKE product id")
    parser.add_argument(
        "--apply-live",
        action="store_true",
        help="Explicitly create or reuse verified live Stripe objects and update ORIGINAL.md.",
    )
    # Compatibility aliases are accepted for old operator notes, but are not
    # part of the normal Commerce v2 command shown in documentation.
    parser.add_argument("--execute-live", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--confirm-product-id", default="", help=argparse.SUPPRESS)
    parser.add_argument("--confirm-tax-code", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--confirmation-text", default="", help=argparse.SUPPRESS)
    parser.add_argument("--resume", action="store_true", help=argparse.SUPPRESS)
    return parser.parse_args()


def print_plan(plan: dict) -> None:
    print(plan["title"])
    print(f"{plan['amount']} {str(plan['currency']).upper()}")
    print("Stripe: dry-run")
    print("Product: reconcile")
    print("Price: reconcile")
    print("Payment Link: reconcile")
    print("ORIGINAL: unchanged")
    print(f"COMMERCE: {'PLAN READY' if plan['ready'] else 'STOPPED'}")
    if plan["errors"]:
        for error in plan["errors"]:
            print(f"- {error}")


def print_result(result: dict) -> None:
    actions = result["actions"]
    print(result["title"])
    print(f"{result['amount']} {str(result['currency']).upper()}")
    print("Stripe: ready")
    print(f"Product: {actions['product']}")
    print(f"Price: {actions['price']}")
    print(f"Payment Link: {actions['payment_link']}")
    print(f"ORIGINAL: {actions['original']}")
    print("COMMERCE: READY")


def main() -> int:
    args = parse_args()
    if not args.product_id:
        print("usage: python tools\\release_product.py <product_id> [--apply-live]")
        return 2
    if args.execute_live:
        print("COMMERCE: STOPPED - --execute-live is legacy; use --apply-live")
        return 1
    try:
        payload, plan = build_dry_run_plan(args.product_id)
        print_plan(plan)
        if not plan["ready"]:
            return 1
        if not args.apply_live:
            return 0
        result = execute_live(payload)
        print_result(result)
        return 0
    except SafetyStop as exc:
        print(f"COMMERCE: STOPPED - {safe_error_message(exc)}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
