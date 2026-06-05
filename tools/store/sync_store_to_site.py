from __future__ import annotations

import argparse
import json
import subprocess
import sys
from collections import Counter
from pathlib import Path
from typing import Any


SOURCE_POLICY_MARKER = "ORIGINAL.md is the source of truth"
DEFAULT_STORE_SITE = Path(r"C:\Users\yukiz\devlop\dake-store-site")
EXPECTED_REFERENCE = {
    "items": 53,
    "type_counts": {"app": 50, "pack": 2, "shimarisu_pack": 1},
    "payment_counts": {"stripe_ready": 5, "booth_only": 47, "preparing": 1},
}


def series_root() -> Path:
    return Path(__file__).resolve().parents[2]


def load_json(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as file:
        data = json.load(file)
    if not isinstance(data, dict):
        raise ValueError(f"JSON root must be an object: {path}")
    return data


def semantic_data(data: dict[str, Any]) -> dict[str, Any]:
    cleaned = dict(data)
    cleaned.pop("generated_at", None)
    return cleaned


def counters(data: dict[str, Any]) -> dict[str, Any]:
    items = data.get("items")
    if not isinstance(items, list):
        raise ValueError("generated JSON must contain an items array")
    type_counts = Counter(item.get("type") for item in items)
    payment_counts = Counter(item.get("payment_status") for item in items)
    stripe_link_count = sum(1 for item in items if item.get("stripe_payment_link"))
    booth_url_count = sum(1 for item in items if item.get("booth_url"))
    return {
        "items": len(items),
        "type_counts": type_counts,
        "payment_counts": payment_counts,
        "stripe_link_count": stripe_link_count,
        "booth_url_count": booth_url_count,
        "preparing_count": payment_counts.get("preparing", 0),
        "has_shimarisu_pack": any(item.get("type") == "shimarisu_pack" for item in items),
    }


def validate_generated(data: dict[str, Any]) -> dict[str, Any]:
    source_policy = data.get("source_policy")
    if not isinstance(source_policy, str) or SOURCE_POLICY_MARKER not in source_policy:
        raise ValueError("source_policy is missing or does not mention ORIGINAL.md")
    if data.get("do_not_edit") is not True:
        raise ValueError("do_not_edit must be true")
    result = counters(data)
    if not result["has_shimarisu_pack"]:
        raise ValueError("shimarisu_pack item is missing")
    return result


def format_counter(counter: Counter[Any], keys: list[str]) -> str:
    parts = [f"{key}={counter.get(key, 0)}" for key in keys]
    extras = sorted(key for key in counter if key not in keys and key is not None)
    parts.extend(f"{key}={counter.get(key, 0)}" for key in extras)
    return ", ".join(parts)


def print_summary(title: str, summary: dict[str, Any]) -> None:
    print(f"\n{title}")
    print("-" * len(title))
    print(f"items: {summary['items']}")
    print(
        "type: "
        + format_counter(summary["type_counts"], ["app", "pack", "shimarisu_pack"])
    )
    print(
        "payment_status: "
        + format_counter(
            summary["payment_counts"],
            ["stripe_ready", "booth_only", "preparing", "free_download", "not_for_sale"],
        )
    )
    print(f"stripe_payment_link: {summary['stripe_link_count']}")
    print(f"booth_url: {summary['booth_url_count']}")
    print(f"preparing: {summary['preparing_count']}")
    print(f"shimarisu_pack: {'yes' if summary['has_shimarisu_pack'] else 'no'}")


def print_reference(summary: dict[str, Any]) -> None:
    print("\nReference counts (not enforced)")
    print("-------------------------------")
    print(f"items: expected {EXPECTED_REFERENCE['items']}, current {summary['items']}")
    for key, expected in EXPECTED_REFERENCE["type_counts"].items():
        print(f"type.{key}: expected {expected}, current {summary['type_counts'].get(key, 0)}")
    for key, expected in EXPECTED_REFERENCE["payment_counts"].items():
        print(f"payment_status.{key}: expected {expected}, current {summary['payment_counts'].get(key, 0)}")


def run_generator(root: Path) -> None:
    generator = root / "tools" / "store" / "generate_store_products.py"
    if not generator.exists():
        raise FileNotFoundError(f"generator not found: {generator}")
    subprocess.run([sys.executable, str(generator)], cwd=root, check=True)


def next_steps(root: Path, store_site: Path, source_changed: bool, store_changed: bool) -> None:
    print("\nNext git steps")
    print("--------------")
    print(f"cd {root}")
    print("git status")
    if source_changed:
        print("git add tools/generated/store_products.generated.json")
        print('git commit -m "Regenerate Store products data"')
        print("git push origin main")
    else:
        print("No DAKE_series generated JSON semantic change. Commit is not required for generated JSON.")
    print("")
    print(f"cd {store_site}")
    print("git status")
    if store_changed:
        print("git add public/assets/data/store_products.generated.json")
        print('git commit -m "Sync Store products data"')
        print("git push origin main")
    else:
        print("No dake-store-site JSON change. Commit is not required.")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Regenerate Store products JSON and sync it to dake-store-site.")
    parser.add_argument(
        "--store-site",
        type=Path,
        default=DEFAULT_STORE_SITE,
        help="Path to the dake-store-site repository.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    root = series_root()
    store_site = args.store_site.resolve()
    source_json = root / "tools" / "generated" / "store_products.generated.json"
    generator_report = root / "tools" / "reports" / "original_phase52_stripe_payment_link_ready.md"
    store_json = store_site / "public" / "assets" / "data" / "store_products.generated.json"

    if not store_site.exists():
        raise FileNotFoundError(f"store site path does not exist: {store_site}")

    old_source_data = load_json(source_json) if source_json.exists() else None
    old_source_bytes = source_json.read_bytes() if source_json.exists() else None
    old_report_bytes = generator_report.read_bytes() if generator_report.exists() else None
    old_store_data = load_json(store_json) if store_json.exists() else None
    old_store_bytes = store_json.read_bytes() if store_json.exists() else None

    print(f"DAKE_series: {root}")
    print(f"dake-store-site: {store_site}")
    print("\nRegenerating Store products data...")
    run_generator(root)

    new_source_data = load_json(source_json)
    summary = validate_generated(new_source_data)
    print_summary("Generated JSON", summary)
    print_reference(summary)

    source_changed = True
    if old_source_data is not None and semantic_data(old_source_data) == semantic_data(new_source_data):
        source_changed = False
        if old_source_bytes is not None:
            source_json.write_bytes(old_source_bytes)
        print("\nNo semantic change in DAKE_series generated JSON; restored existing file to avoid timestamp-only diff.")
    else:
        print("\nDAKE_series generated JSON changed semantically.")

    if old_report_bytes is not None and generator_report.exists():
        generator_report.write_bytes(old_report_bytes)
        print("Preserved existing generator report; Phase 10 report is the sync audit record.")

    synced_data = load_json(source_json)
    validate_generated(synced_data)
    store_changed = True
    if old_store_data is not None and semantic_data(old_store_data) == semantic_data(synced_data):
        store_changed = False
        print("No semantic change in dake-store-site JSON; left existing Store file untouched.")
    else:
        serialized = json.dumps(synced_data, ensure_ascii=False, indent=2).encode("utf-8") + b"\n"
        store_changed = old_store_bytes != serialized
        store_json.parent.mkdir(parents=True, exist_ok=True)
        store_json.write_bytes(serialized)

    copied_data = load_json(store_json)
    copied_summary = validate_generated(copied_data)
    if semantic_data(copied_data) != semantic_data(synced_data):
        raise ValueError("copied Store JSON does not match DAKE_series JSON")
    print_summary("Synced Store JSON", copied_summary)
    print(f"\nStore JSON copied to: {store_json}")
    print(f"Store JSON changed: {'yes' if store_changed else 'no'}")
    print("Validation: source_policy OK, do_not_edit OK, shimarisu_pack OK, JSON OK")
    next_steps(root, store_site, source_changed, store_changed)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
