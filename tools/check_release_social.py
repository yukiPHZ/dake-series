# -*- coding: utf-8 -*-
"""Check DAKE release social artifacts."""

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

from release_source_policy import app_url_for, read_app_source


ROOT = Path(__file__).resolve().parents[1]
APPS_DIR = ROOT / "01_apps"
REPORT_DIR = ROOT / "tools" / "reports" / "release_artifacts"
SHIPPING_STATUS = "available"
PLATFORMS = ("x", "threads", "instagram")


@dataclass
class SocialCheck:
    app: str
    display_name: str = ""
    status: str = "unknown"
    social_posts: bool = False
    social_release: bool = False
    social_stage: str = ""
    social_reason: str = ""
    social_posts_valid: bool = False
    forbidden_links: bool = False
    buffer_x_id: str = ""
    buffer_threads_id: str = ""
    buffer_instagram_id: str = ""
    source_kind: str = ""
    source_path: str = ""
    original_missing: bool = False
    meta_derivative_mismatch: str = ""
    actions: list[str] = field(default_factory=list)

    @property
    def shipping_candidate(self) -> bool:
        return self.status == SHIPPING_STATUS

    @property
    def social_draft_ready(self) -> bool:
        return self.shipping_candidate and self.social_posts and self.social_posts_valid

    @property
    def social_buffer_ready(self) -> bool:
        return (
            self.social_draft_ready
            and self.social_stage == "complete"
            and bool(self.buffer_x_id)
            and bool(self.buffer_threads_id)
            and bool(self.buffer_instagram_id)
        )

    @property
    def ok(self) -> bool:
        return self.social_buffer_ready


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return path.read_text(encoding="utf-8", errors="replace")


def read_meta(app_dir: Path) -> tuple[dict[str, Any], str]:
    source = read_app_source(app_dir, ROOT)
    return source.meta, source.error


def buffer_id(record: dict[str, Any], platform: str) -> str:
    item = record.get("buffer", {}).get(platform, {})
    if not isinstance(item, dict):
        return ""
    for key in ("buffer_post_id", "buffer_update_id", "update_id", "id"):
        if item.get(key):
            return str(item[key])
    return ""


def validate_social_posts(path: Path, app_url: str) -> tuple[bool, bool]:
    if not path.exists():
        return False, False
    text = read_text(path)
    urls = re.findall(r"https?://[^\s)]+", text)
    expected = app_url.rstrip("/")
    forbidden = any(("github.com" in url.lower() or "booth.pm" in url.lower() or url.rstrip(".,").rstrip("/") != expected) for url in urls)
    required = ("## X", "## THREADS", "## INSTAGRAM")
    has_sections = all(section in text for section in required)
    return has_sections and not forbidden, forbidden


def check_app(app_dir: Path, only_available: bool) -> SocialCheck:
    source = read_app_source(app_dir, ROOT)
    meta = source.meta
    meta_error = source.error
    status = str(meta.get("status") or "unknown")
    result = SocialCheck(
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
    app_url = app_url_for(app_dir, meta, ROOT.parent / "dakeapp-site")
    release_dir = app_dir / "release_artifacts"
    posts = release_dir / "social_posts.md"
    release = release_dir / "social_release.json"
    result.social_posts = posts.exists()
    result.social_posts_valid, result.forbidden_links = validate_social_posts(posts, app_url)
    result.social_release = release.exists()
    if release.exists():
        try:
            record = json.loads(read_text(release))
            result.social_stage = str(record.get("stage") or "")
            result.social_reason = str(record.get("reason") or "")
            result.buffer_x_id = buffer_id(record, "x")
            result.buffer_threads_id = buffer_id(record, "threads")
            result.buffer_instagram_id = buffer_id(record, "instagram")
        except Exception as exc:
            result.actions.append(f"social_release.json parse failed: {exc}")
    if not result.social_posts:
        result.actions.append("generate social_posts.md with tools/release_social.py")
    elif not result.social_posts_valid:
        result.actions.append("fix social post body/link policy")
    if not result.social_release:
        result.actions.append("generate dry-run social_release.json with tools/release_social.py")
    if result.social_release and result.social_stage not in {"dry_run", "complete"}:
        result.actions.append(f"social stage is {result.social_stage or 'empty'}: {result.social_reason}")
    if result.social_draft_ready and not result.social_buffer_ready:
        result.actions.append("create Buffer posts with tools/release_social.py --post-to-buffer")
    for platform, value in (("x", result.buffer_x_id), ("threads", result.buffer_threads_id), ("instagram", result.buffer_instagram_id)):
        if not value:
            result.actions.append(f"{platform} Buffer ID missing")
    return result


def app_dirs() -> list[Path]:
    return sorted(path for path in APPS_DIR.iterdir() if path.is_dir() and path.name.startswith("DAKE_"))


def write_csv(results: list[SocialCheck], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fields = [
        "app",
        "display_name",
        "status",
        "ok",
        "social_draft_ready",
        "social_buffer_ready",
        "social_posts",
        "social_posts_valid",
        "forbidden_links",
        "social_release",
        "social_stage",
        "social_reason",
        "buffer_x_id",
        "buffer_threads_id",
        "buffer_instagram_id",
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


def write_markdown(results: list[SocialCheck], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    shipping = [item for item in results if item.shipping_candidate]
    ok = [item for item in shipping if item.ok]
    not_ok = [item for item in shipping if not item.ok]
    draft_ready = [item for item in shipping if item.social_draft_ready]
    buffer_ready = [item for item in shipping if item.social_buffer_ready]
    counts = Counter(item.social_stage or "missing" for item in shipping)
    lines = [
        "# DAKE Release Social Check",
        "",
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "",
        "## Summary",
        "",
        f"- checked: {len(results)}",
        f"- available checked: {len(shipping)}",
        f"- original_missing: {sum(item.original_missing for item in shipping)}",
        f"- meta_derivative_mismatch: {sum(bool(item.meta_derivative_mismatch) for item in shipping)}",
        f"- social ok: {len(ok)}",
        f"- social_draft_ready: {len(draft_ready)}",
        f"- social_buffer_ready: {len(buffer_ready)}",
        f"- social missing or failed: {len(not_ok)}",
        f"- social_posts.md missing: {sum(not item.social_posts for item in shipping)}",
        f"- social_release.json missing: {sum(not item.social_release for item in shipping)}",
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
        "| app | stage | draft | buffer | X | Threads | Instagram | action |",
        "| --- | --- | --- | --- | --- | --- | --- | --- |",
    ])
    for item in not_ok:
        lines.append(
            f"| {item.app} | {item.social_stage or '-'} | {item.social_draft_ready} | {item.social_buffer_ready} | {bool(item.buffer_x_id)} | {bool(item.buffer_threads_id)} | {bool(item.buffer_instagram_id)} | {'<br>'.join(item.actions) or '-'} |"
        )
    lines.extend(["", "## Complete", "", "| app | X | Threads | Instagram |", "| --- | --- | --- | --- |"])
    for item in ok:
        lines.append(f"| {item.app} | {item.buffer_x_id} | {item.buffer_threads_id} | {item.buffer_instagram_id} |")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Check DAKE release social artifacts.")
    parser.add_argument("--only-available", action="store_true")
    parser.add_argument("--report-dir", type=Path, default=REPORT_DIR)
    parser.add_argument("--fail-on-error", action="store_true")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    results = [check_app(app_dir, args.only_available) for app_dir in app_dirs()]
    md_path = args.report_dir / "release_social_check.md"
    csv_path = args.report_dir / "release_social_check.csv"
    write_markdown(results, md_path)
    write_csv(results, csv_path)
    shipping = [item for item in results if item.shipping_candidate]
    failed = [item for item in shipping if not item.ok]
    print("DAKE Release Social Check")
    print(f"checked: {len(results)}")
    print(f"available checked: {len(shipping)}")
    print(f"original_missing: {sum(item.original_missing for item in shipping)}")
    print(f"meta_derivative_mismatch: {sum(bool(item.meta_derivative_mismatch) for item in shipping)}")
    print(f"social ok: {len(shipping) - len(failed)}")
    print(f"social_draft_ready: {sum(item.social_draft_ready for item in shipping)}")
    print(f"social_buffer_ready: {sum(item.social_buffer_ready for item in shipping)}")
    print(f"social missing or failed: {len(failed)}")
    print(f"report: {md_path}")
    print(f"csv: {csv_path}")
    if failed:
        print("not complete:")
        for item in failed:
            print(f"- {item.app}: {', '.join(item.actions)}")
    return 1 if args.fail_on_error and failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
