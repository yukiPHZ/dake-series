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
        description="Inspect and advance the DAKE product release pipeline one product at a time.",
    )
    add_common_args(parser)
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
        return 0 if status["current_stage"] not in {"SOURCE_INVALID", "INCONSISTENT"} else 1

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
    return 0 if result.get("advanced") or status["current_stage"] in {"STRIPE_DRY_RUN_READY", "CHECKOUT_REVIEW_PENDING", "STORE_GENERATED", "STORE_SYNC_PENDING", "STORE_SYNCED", "PRODUCTION_VERIFICATION_PENDING", "RELEASE_COMPLETE", "LEGACY_COMPLETE"} else 1


if __name__ == "__main__":
    raise SystemExit(main())
