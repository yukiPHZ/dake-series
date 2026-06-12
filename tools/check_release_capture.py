# -*- coding: utf-8 -*-
"""Check DAKE release capture artifacts.

The check is intentionally strict about the evidence video:
``release_artifacts/release.json`` must record ``video_local_path`` and that
file must exist. A lone video file is not enough to close formal shipping.
"""

from __future__ import annotations

import argparse
import csv
import json
import re
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

from release_source_policy import read_app_source


ROOT = Path(__file__).resolve().parents[1]
APPS_DIR = ROOT / "01_apps"
REPORT_DIR = ROOT / "tools" / "reports" / "release_artifacts"
SHIPPING_STATUS = "available"


@dataclass
class CaptureCheck:
    app: str
    display_name: str = ""
    status: str = "unknown"
    release_json: bool = False
    release_stage: str = ""
    release_reason: str = ""
    video_local_path: str = ""
    video_exists: bool = False
    demo_mp4: bool = False
    demo_webm: bool = False
    screenshot: bool = False
    booth_thumbnail: bool = False
    social_release_path: str = ""
    source_kind: str = ""
    source_path: str = ""
    original_missing: bool = False
    meta_derivative_mismatch: str = ""
    actions: list[str] = field(default_factory=list)

    @property
    def shipping_candidate(self) -> bool:
        return self.status == SHIPPING_STATUS

    @property
    def ok(self) -> bool:
        return (
            self.shipping_candidate
            and self.release_json
            and self.release_stage == "complete"
            and bool(self.video_local_path)
            and self.video_exists
            and (self.demo_mp4 or self.demo_webm)
        )


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return path.read_text(encoding="utf-8", errors="replace")


def read_json(path: Path) -> dict[str, Any]:
    return json.loads(read_text(path))


def read_meta(app_dir: Path) -> tuple[dict[str, Any], str]:
    source = read_app_source(app_dir, ROOT)
    return source.meta, source.error


def app_dirs() -> list[Path]:
    return sorted(path for path in APPS_DIR.iterdir() if path.is_dir() and path.name.startswith("DAKE_"))


def check_app(app_dir: Path, only_available: bool) -> CaptureCheck:
    source = read_app_source(app_dir, ROOT)
    meta = source.meta
    meta_error = source.error
    status = str(meta.get("status") or "unknown")
    result = CaptureCheck(
        app=app_dir.name,
        display_name=str(meta.get("display_name") or meta.get("site_title") or app_dir.name),
        status=status,
        source_kind=source.source_kind,
        source_path=source.source_label,
        original_missing=source.original_missing,
        meta_derivative_mismatch=",".join(source.derivative_mismatches),
    )
    if source.original_missing:
        result.actions.append("original_missing")
    if source.derivative_mismatches:
        result.actions.append("DAKE_META derivative mismatch: " + ",".join(source.derivative_mismatches))
    if only_available and status != SHIPPING_STATUS:
        result.actions.append(f"skipped status={status}")
        return result
    if meta_error:
        result.actions.append(meta_error)

    release_dir = app_dir / "release_artifacts"
    release_json_path = release_dir / "release.json"
    result.demo_mp4 = (release_dir / "demo.mp4").exists()
    result.demo_webm = (release_dir / "demo.webm").exists()
    result.screenshot = (app_dir / str(meta.get("screenshot_path") or "assets/screenshot.webp")).exists()
    result.booth_thumbnail = (app_dir / "assets" / "booth_thumbnail.jpg").exists()

    if release_json_path.exists():
        result.release_json = True
        try:
            record = read_json(release_json_path)
            result.release_stage = str(record.get("stage") or "")
            result.release_reason = str(record.get("reason") or "")
            result.video_local_path = str(record.get("video_local_path") or "")
            result.social_release_path = str(record.get("social_release_path") or "")
            if result.video_local_path:
                result.video_exists = (app_dir / result.video_local_path).exists()
        except Exception as exc:
            result.actions.append(f"release.json parse failed: {exc}")
    else:
        result.actions.append("create release_artifacts/release.json with tools/release_capture.py")

    if not result.video_local_path:
        result.actions.append("record video_local_path in release.json")
    if result.video_local_path and not result.video_exists:
        result.actions.append(f"video_local_path missing on disk: {result.video_local_path}")
    if not (result.demo_mp4 or result.demo_webm):
        result.actions.append("save demo.mp4 or demo.webm")
    if result.release_json and result.release_stage != "complete":
        result.actions.append(f"release stage is {result.release_stage or 'empty'}: {result.release_reason}")
    return result


def write_csv(results: list[CaptureCheck], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fields = [
        "app",
        "display_name",
        "status",
        "ok",
        "release_json",
        "release_stage",
        "release_reason",
        "video_local_path",
        "video_exists",
        "demo_mp4",
        "demo_webm",
        "screenshot",
        "booth_thumbnail",
        "social_release_path",
        "source_kind",
        "source_path",
        "original_missing",
        "meta_derivative_mismatch",
        "actions",
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        for item in results:
            writer.writerow({field: getattr(item, field) if field != "actions" else "; ".join(item.actions) for field in fields})


def write_markdown(results: list[CaptureCheck], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    shipping = [item for item in results if item.shipping_candidate]
    ok = [item for item in shipping if item.ok]
    not_ok = [item for item in shipping if not item.ok]
    counts = Counter(item.release_stage or "missing" for item in shipping)
    lines = [
        "# DAKE Release Capture Check",
        "",
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "",
        "## Summary",
        "",
        f"- checked: {len(results)}",
        f"- available checked: {len(shipping)}",
        f"- original_missing: {sum(item.original_missing for item in shipping)}",
        f"- meta_derivative_mismatch: {sum(bool(item.meta_derivative_mismatch) for item in shipping)}",
        f"- capture ok: {len(ok)}",
        f"- capture missing or failed: {len(not_ok)}",
        f"- demo.mp4/webm missing: {sum(not (item.demo_mp4 or item.demo_webm) for item in shipping)}",
        f"- video_local_path missing: {sum(not item.video_local_path for item in shipping)}",
        "",
        "## Stage Counts",
        "",
    ]
    for stage, count in sorted(counts.items()):
        lines.append(f"- {stage}: {count}")
    lines.extend([
        "",
        "## Not Complete",
        "",
        "| app | stage | video_local_path | video exists | action |",
        "| --- | --- | --- | --- | --- |",
    ])
    for item in not_ok:
        lines.append(
            f"| {item.app} | {item.release_stage or '-'} | {item.video_local_path or '-'} | {item.video_exists} | {'<br>'.join(item.actions) or '-'} |"
        )
    lines.extend(["", "## Complete", "", "| app | video |", "| --- | --- |"])
    for item in ok:
        lines.append(f"| {item.app} | {item.video_local_path} |")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Check DAKE release capture artifacts.")
    parser.add_argument("--only-available", action="store_true")
    parser.add_argument("--report-dir", type=Path, default=REPORT_DIR)
    parser.add_argument("--fail-on-error", action="store_true")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    results = [check_app(app_dir, args.only_available) for app_dir in app_dirs()]
    md_path = args.report_dir / "release_capture_check.md"
    csv_path = args.report_dir / "release_capture_check.csv"
    write_markdown(results, md_path)
    write_csv(results, csv_path)
    shipping = [item for item in results if item.shipping_candidate]
    failed = [item for item in shipping if not item.ok]
    print("DAKE Release Capture Check")
    print(f"checked: {len(results)}")
    print(f"available checked: {len(shipping)}")
    print(f"original_missing: {sum(item.original_missing for item in shipping)}")
    print(f"meta_derivative_mismatch: {sum(bool(item.meta_derivative_mismatch) for item in shipping)}")
    print(f"capture ok: {len(shipping) - len(failed)}")
    print(f"capture missing or failed: {len(failed)}")
    print(f"report: {md_path}")
    print(f"csv: {csv_path}")
    if failed:
        print("not complete:")
        for item in failed:
            print(f"- {item.app}: {', '.join(item.actions)}")
    return 1 if args.fail_on_error and failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
