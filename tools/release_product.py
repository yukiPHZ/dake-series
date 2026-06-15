from __future__ import annotations

import argparse

from store.stripe_release_core import (
    SafetyStop,
    build_release_payload,
    execute_live_release,
    validate_dry_run_payload,
    write_dry_run_files,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Prepare a DAKE product for Stripe release from ORIGINAL.md.",
    )
    parser.add_argument("product_id", nargs="?", help="DAKE product id, for example DAKE_Pack_Document")
    parser.add_argument("--execute-live", action="store_true", help="Create live Stripe objects after explicit confirmations.")
    parser.add_argument("--confirm-product-id", default="", help="Must exactly match product_id for live execution.")
    parser.add_argument("--confirm-tax-code", action="store_true", help="Confirm tax_code candidate review before live execution.")
    parser.add_argument("--confirmation-text", default="", help="Must be CREATE LIVE PAYMENT LINK <product_id> for live execution.")
    parser.add_argument("--resume", action="store_true", help="Resume only after safe state review.")
    return parser.parse_args()


def print_payload_summary(payload: dict, json_path: object, md_path: object) -> None:
    print(f"mode={payload['mode']}")
    print(f"product_id={payload['product_id']}")
    print(f"product_type={payload['product_type']}")
    print(f"price={payload['price']}")
    print(f"currency={payload['currency']}")
    print(f"tax_code_candidate={payload['tax_code_candidate']}")
    print(f"payment_status_before={payload['payment_status_before']}")
    print(f"stripe_payment_link_before={payload['stripe_payment_link_before'] or ''}")
    print(f"purchase_delivery_ready={payload['purchase_delivery_ready']}")
    print(f"purchase_delivery_method={payload['purchase_delivery_method']}")
    print(f"distribution_file={payload['distribution_file']}")
    print(f"ready_for_live_execution={payload['ready_for_live_execution']}")
    print(f"errors={len(payload['errors'])}")
    print(f"secret_read={payload['secret_read']}")
    print(f"live_api_called={payload['live_api_called']}")
    print(f"output_json={json_path}")
    print(f"output_md={md_path}")
    if payload["errors"]:
        print("validation_errors:")
        for error in payload["errors"]:
            print(f"- {error}")


def main() -> int:
    args = parse_args()
    if not args.product_id:
        print("usage: python tools\\release_product.py <product_id> [--execute-live]")
        return 2
    try:
        payload = build_release_payload(args.product_id)
        json_path, md_path = write_dry_run_files(payload)
        print_payload_summary(payload, json_path, md_path)
        validation_errors = validate_dry_run_payload(payload)
        if validation_errors:
            return 1
        if not args.execute_live:
            return 0
        result = execute_live_release(payload, args)
        print("live_api_called=yes")
        return result
    except SafetyStop as exc:
        print(str(exc))
        print("live_api_called=no")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
