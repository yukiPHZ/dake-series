# -*- coding: utf-8 -*-
"""Sync generated DAKE/Store link data to public site repos.

The synced JSON files are derived views. Update ORIGINAL.md first, regenerate
the DAKE_series generated files, then run this tool.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
GENERATED_DIR = ROOT / "tools" / "generated"
DEFAULT_DAKEAPP_SITE = Path(r"C:\Users\yukiz\devlop\dakeapp-site")
DEFAULT_STORE_SITE = Path(r"C:\Users\yukiz\devlop\dake-store-site")


def load_json(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as file:
        data = json.load(file)
    if not isinstance(data, dict):
        raise ValueError(f"JSON root must be an object: {path}")
    return data


def semantic(data: dict[str, Any]) -> dict[str, Any]:
    cleaned = dict(data)
    cleaned.pop("generated_at", None)
    return cleaned


def write_if_changed(path: Path, data: dict[str, Any]) -> bool:
    serialized = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8") + b"\n"
    if path.exists():
        old_bytes = path.read_bytes()
        try:
            old = json.loads(old_bytes.decode("utf-8"))
        except json.JSONDecodeError:
            old = None
        if old is not None and semantic(old) == semantic(data):
            return False
        if old_bytes == serialized:
            return False
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(serialized)
    return True


def validate_apps(data: dict[str, Any]) -> int:
    if data.get("do_not_edit") is not True:
        raise ValueError("apps data must be marked do_not_edit")
    apps = data.get("apps")
    if not isinstance(apps, list):
        raise ValueError("apps data must contain apps array")
    return len(apps)


def validate_store(data: dict[str, Any]) -> int:
    if data.get("do_not_edit") is not True:
        raise ValueError("store products data must be marked do_not_edit")
    items = data.get("items")
    if not isinstance(items, list):
        raise ValueError("store products data must contain items array")
    return len(items)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Sync generated DAKE/Store link data to site repos.")
    parser.add_argument("--dakeapp-site", type=Path, default=DEFAULT_DAKEAPP_SITE)
    parser.add_argument("--store-site", type=Path, default=DEFAULT_STORE_SITE)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    apps_data = load_json(GENERATED_DIR / "dakeapp_apps.generated.json")
    store_data = load_json(GENERATED_DIR / "store_products.generated.json")
    apps_count = validate_apps(apps_data)
    store_count = validate_store(store_data)

    dakeapp_store_json = args.dakeapp_site / "public" / "assets" / "data" / "store_products.generated.json"
    store_apps_json = args.store_site / "public" / "assets" / "data" / "apps.generated.json"

    dakeapp_changed = write_if_changed(dakeapp_store_json, store_data)
    store_changed = write_if_changed(store_apps_json, apps_data)

    print("DAKE web link data sync")
    print(f"apps: {apps_count}")
    print(f"store items: {store_count}")
    print(f"dakeapp-site store data: {dakeapp_store_json} changed={dakeapp_changed}")
    print(f"dake-store-site app data: {store_apps_json} changed={store_changed}")
    print("Validation: generated JSON OK, do_not_edit OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
