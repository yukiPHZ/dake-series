# -*- coding: utf-8 -*-
"""Generate and sync dakeapp.com app data from DAKE_series.

apps.json is a derived view. Do not edit the site copy manually; update
ORIGINAL.md first, then run this tool.
"""

from __future__ import annotations

import argparse
import json
import re
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Any

from release_source_policy import SOURCE_POLICY_TEXT, app_url_for, read_app_source


ROOT = Path(__file__).resolve().parents[1]
APPS_DIR = ROOT / "01_apps"
GENERATED_DIR = ROOT / "tools" / "generated"
DEFAULT_SITE_ROOT = Path(r"C:\Users\yukiz\devlop\dakeapp-site")
JST = timezone(timedelta(hours=9))


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return path.read_text(encoding="utf-8", errors="replace")


def text_between(text: str, pattern: str) -> str:
    match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
    return re.sub(r"\s+", " ", match.group(1)).strip() if match else ""


def site_pages(site_root: Path) -> dict[str, str]:
    pages: dict[str, str] = {}
    for page in sorted((site_root / "public" / "apps").glob("*/index.html")):
        pages[page.parent.name] = read_text(page)
    return pages


def source_by_slug(site_root: Path) -> dict[str, Any]:
    mapping: dict[str, Any] = {}
    for app_dir in sorted(path for path in APPS_DIR.iterdir() if path.is_dir() and path.name.startswith("DAKE_")):
        source = read_app_source(app_dir, ROOT)
        if source.meta:
            slug = app_url_for(app_dir, source.meta, site_root).rstrip("/").split("/")[-1]
            mapping[slug] = source
    return mapping


def build_data(site_root: Path) -> dict[str, Any]:
    pages = site_pages(site_root)
    sources = source_by_slug(site_root)
    apps = []
    for slug, html in pages.items():
        source = sources.get(slug)
        meta = source.meta if source else {}
        release_url = text_between(html, r'<a[^>]+class="button"[^>]+href="([^"]+)"')
        screenshot_src = text_between(html, r'<img[^>]+class="screenshot"[^>]+src="([^"]+)"')
        title = text_between(html, r"<h1[^>]*>(.*?)</h1>")
        apps.append(
            {
                "slug": slug,
                "app_key": meta.get("app_key", ""),
                "folder_name": meta.get("folder_name", source.app_dir.name if source else ""),
                "display_name": meta.get("display_name", title),
                "detail_url": f"https://dakeapp.com/apps/{slug}/",
                "release_url": meta.get("release_url") or release_url,
                "screenshot_path": screenshot_src,
                "demo_video_path": meta.get("demo_video_path", "release_artifacts/demo.mp4"),
                "demo_video_url": meta.get("demo_video_url", ""),
                "video_cloudinary_url": "",
                "video_r2_url": "",
                "social_release_path": meta.get("social_release_path", "release_artifacts/social_release.json"),
                "source_kind": source.source_kind if source else "site_only",
                "source_original": source.source_label if source and source.source_kind == "original" else "",
                "source_readme": source.source_label if source and source.source_kind == "readme_fallback" else "",
                "original_missing": bool(source.original_missing) if source else True,
                "meta_derivative_mismatch": source.derivative_mismatches if source else [],
            }
        )
    return {
        "schema": "dakeapp.apps.v1",
        "do_not_edit": True,
        "source_policy": SOURCE_POLICY_TEXT,
        "generated_at": datetime.now(JST).isoformat(timespec="seconds"),
        "apps": sorted(apps, key=lambda item: item["slug"]),
    }


def semantic(data: dict[str, Any]) -> dict[str, Any]:
    cleaned = dict(data)
    cleaned.pop("generated_at", None)
    return cleaned


def write_json_if_changed(path: Path, data: dict[str, Any], preserve_timestamp_only: bool) -> bool:
    old = None
    old_bytes = None
    if path.exists():
        old_bytes = path.read_bytes()
        old = json.loads(old_bytes.decode("utf-8"))
    if preserve_timestamp_only and old is not None and semantic(old) == semantic(data):
        return False
    serialized = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8") + b"\n"
    path.parent.mkdir(parents=True, exist_ok=True)
    if old_bytes == serialized:
        return False
    path.write_bytes(serialized)
    return True


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate and sync dakeapp.com apps.json from DAKE_series.")
    parser.add_argument("--site-root", type=Path, default=DEFAULT_SITE_ROOT)
    parser.add_argument("--preserve-timestamp-only", action=argparse.BooleanOptionalAction, default=False)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    data = build_data(args.site_root)
    source_json = GENERATED_DIR / "dakeapp_apps.generated.json"
    site_json = args.site_root / "public" / "assets" / "data" / "apps.json"
    source_changed = write_json_if_changed(source_json, data, args.preserve_timestamp_only)
    site_changed = write_json_if_changed(site_json, data, args.preserve_timestamp_only)
    original_missing = sum(1 for item in data["apps"] if item.get("original_missing"))
    mismatches = sum(1 for item in data["apps"] if item.get("meta_derivative_mismatch"))
    print("DAKE dakeapp apps.json sync")
    print(f"apps: {len(data['apps'])}")
    print(f"original_missing: {original_missing}")
    print(f"meta_derivative_mismatch: {mismatches}")
    print(f"source: {source_json} changed={source_changed}")
    print(f"site: {site_json} changed={site_changed}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
