"""Legacy-compatible DISTRIBUTION / COMMERCE status pipeline.

This tool does not build an app, validate its main function, create a GitHub
Release, or decide SHIP completion. Normal sales preparation uses the BOOTH,
dakeapp, and Store tools independently and stops before external registration,
publication, or Commerce operations. Cross-layer automatic ``advance`` is
disabled unless the explicit Legacy compatibility flag is supplied.
"""

from __future__ import annotations

import argparse

from store.release_pipeline_core import (
    ReleasePipeline,
    add_common_args,
    cli_json,
    print_status,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Inspect the Legacy DAKE DISTRIBUTION / COMMERCE pipeline without changing SHIP.",
    )
    add_common_args(parser)
    parser.add_argument(
        "--allow-legacy-commerce-advance",
        action="store_true",
        help="Explicitly allow the old cross-layer advance behavior; use only for a separately requested Commerce task.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    pipeline = ReleasePipeline()
    status = pipeline.status(args.product_id)

    if args.command == "status":
        if args.save_report:
            json_path, md_path = pipeline.save_status_report(status)
            status = {**status, "saved_report_json": str(json_path), "saved_report_md": str(md_path)}
        if args.json:
            print(cli_json(status))
        else:
            print_status(status)
        return 0 if status["current_stage"] not in {"SOURCE_INVALID", "INCONSISTENT"} else 1

    if args.command == "next":
        if args.json:
            print(cli_json(status))
        else:
            print(f"product_id: {status['product_id']}")
            print(f"current_stage: {status['current_stage']}")
            print(f"next_action: {status['next_action'] or 'none'}")
            print(f"next_legacy_action: {status.get('next_formal_action') or 'none'}")
        return 0 if status["current_stage"] not in {"SOURCE_INVALID", "INCONSISTENT"} else 1

    if not args.allow_legacy_commerce_advance:
        message = (
            "automatic cross-layer advance is disabled; use make_booth_ready.py, "
            "sync_dakeapp_apps_json.py, or store/sync_store_to_site.py for the requested "
            "DISTRIBUTION layer, and use Commerce tools only on explicit request"
        )
        if args.json:
            print(cli_json({"status": status, "advanced": False, "message": message}))
        else:
            print(f"product_id: {args.product_id}")
            print(f"current_stage: {status['current_stage']}")
            print("advanced: False")
            print(f"message: {message}")
        return 1

    result = pipeline.advance(args.product_id)
    if args.json:
        print(cli_json({"status": status, "advance": result}))
    else:
        print(f"product_id: {args.product_id}")
        print(f"current_stage: {status['current_stage']}")
        print(f"advanced: {result.get('advanced')}")
        print(f"message: {result.get('message') or ''}")
        if result.get("command"):
            command = result["command"]
            if isinstance(command, list):
                print("command: " + " ".join(str(part) for part in command))
            else:
                print(f"command: {command}")
        if "returncode" in result:
            print(f"returncode: {result['returncode']}")
    return 0 if result.get("advanced") or status["current_stage"] in {"BOOTH_REGISTRATION_PENDING", "STRIPE_DRY_RUN_READY", "CHECKOUT_REVIEW_PENDING", "STORE_GENERATED", "STORE_SYNC_PENDING", "STORE_SYNCED", "PRODUCTION_VERIFICATION_PENDING", "RELEASE_COMPLETE", "LEGACY_COMPLETE"} else 1


if __name__ == "__main__":
    raise SystemExit(main())
