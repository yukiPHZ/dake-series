"""Check DAKE formal shipping assets.

Version 2 treats only ``status: available`` apps as normal shipping
candidates. Frozen, draft, experimental, private, and internal apps are
listed, but they are not counted as ordinary missing BOOTH assets.

Formal shipping is not closed by GitHub Release alone. An available app is
closed only when release assets, BOOTH ready assets, a GitHub Release URL,
the BOOTH ``# URL`` field, local operation evidence, and Buffer social
post IDs are all present.
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

from release_source_policy import app_url_for, read_app_source


ROOT = Path(__file__).resolve().parents[1]
APPS_DIR = ROOT / "01_apps"
REPORT_DIR = ROOT / "tools" / "reports"
COMMON_ICON_PARTS = ("02_assets", "dake_icon.ico")
SHIPPING_STATUS = "available"
KNOWN_NON_SHIPPING_STATUSES = {"frozen", "draft", "experimental", "private", "internal"}
STATUS_DISPLAY_ORDER = ["available", "frozen", "draft", "experimental", "private", "internal", "unknown"]
ROOT_GITIGNORE_REQUIRED = [
    "build/",
    "dist/",
    "*.spec",
    "*_config.json",
    "__pycache__/",
    "*.pyc",
    "playwright_profile/",
]


@dataclass
class AppCheck:
    app: str
    readme: bool = False
    dake_meta: bool = False
    meta_error: str = ""
    status: str = "unknown"
    show_in_launcher: bool | None = None
    show_on_site: bool | None = None
    release_url: str = ""
    release_body: bool = False
    screenshot: bool = False
    thumbnail: bool = False
    booth_product: bool = False
    booth_product_path: str = ""
    booth_url_state: str = "booth_product_missing"
    booth_url: str = ""
    booth_ready: bool = False
    zip_file: bool = False
    build_bat: bool = False
    main_py: bool = False
    dist_exe: bool = False
    icon_build: bool = False
    icon_main: bool = False
    release_json: bool = False
    release_stage: str = ""
    video_local_path: str = ""
    demo_video: bool = False
    social_posts: bool = False
    social_release: bool = False
    social_stage: str = ""
    social_posts_valid: bool = False
    social_forbidden_links: bool = False
    buffer_x_id: bool = False
    buffer_threads_id: bool = False
    buffer_instagram_id: bool = False
    source_kind: str = ""
    source_path: str = ""
    original_missing: bool = False
    meta_derivative_mismatch: str = ""
    actions: list[str] = field(default_factory=list)

    @property
    def shipping_candidate(self) -> bool:
        return self.status == SHIPPING_STATUS

    @property
    def flag_conflict(self) -> bool:
        return not self.shipping_candidate and (
            self.show_in_launcher is True or self.show_on_site is True
        )

    @property
    def asset_ready(self) -> bool:
        return all(
            [
                self.readme,
                self.dake_meta,
                self.release_body,
                self.screenshot,
                self.thumbnail,
                self.booth_product,
                self.booth_ready,
                self.zip_file,
                self.build_bat,
                self.main_py,
                self.dist_exe,
                self.icon_build,
                self.icon_main,
            ]
        )

    @property
    def release_ready(self) -> bool:
        return bool(self.release_url)

    @property
    def booth_url_ready(self) -> bool:
        return self.booth_url_state == "set"

    @property
    def release_capture_ready(self) -> bool:
        return (
            self.release_json
            and self.release_stage == "complete"
            and bool(self.video_local_path)
            and self.demo_video
        )

    @property
    def social_draft_ready(self) -> bool:
        return self.social_posts and self.social_posts_valid

    @property
    def social_buffer_ready(self) -> bool:
        return (
            self.social_draft_ready
            and self.social_stage == "complete"
            and self.buffer_x_id
            and self.buffer_threads_id
            and self.buffer_instagram_id
        )

    @property
    def social_ready(self) -> bool:
        return self.social_buffer_ready

    @property
    def v2_pending_reason(self) -> str:
        reasons: list[str] = []
        if not self.legacy_closed:
            reasons.append("legacy")
        if not self.release_capture_ready:
            reasons.append("capture")
        if not self.social_draft_ready:
            reasons.append("social_draft")
        if not self.social_buffer_ready:
            reasons.append("social_buffer")
        return ", ".join(reasons)

    @property
    def legacy_closed(self) -> bool:
        return (
            self.shipping_candidate
            and self.asset_ready
            and self.release_ready
            and self.booth_url_ready
        )

    @property
    def v2_closed(self) -> bool:
        return self.legacy_closed and self.release_capture_ready and self.social_ready

    @property
    def closed_ready(self) -> bool:
        return self.legacy_closed

    @property
    def ok(self) -> bool:
        return self.legacy_closed


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return path.read_text(encoding="utf-8", errors="replace")


def normalize_path_text(text: str) -> str:
    return text.replace("/", "\\").lower()


def has_common_icon(text: str) -> bool:
    normalized = normalize_path_text(text)
    return all(part.lower() in normalized for part in COMMON_ICON_PARTS)


def has_release_body(readme_text: str) -> bool:
    return "## RELEASE_BODY" in readme_text and bool(
        readme_text.split("## RELEASE_BODY", 1)[1].strip()
    )


def parse_dake_meta(readme_text: str) -> tuple[dict[str, Any] | None, str]:
    if "## DAKE_META" not in readme_text or "```json" not in readme_text:
        return None, "missing DAKE_META block"
    try:
        block = (
            readme_text.split("## DAKE_META", 1)[1]
            .split("```json", 1)[1]
            .split("```", 1)[0]
        )
        return json.loads(block), ""
    except IndexError:
        return None, "broken DAKE_META fence"
    except json.JSONDecodeError as exc:
        return None, f"invalid DAKE_META JSON: {exc}"


def find_booth_product(app_dir: Path) -> Path | None:
    candidates = [
        app_dir / "booth_ready" / "booth_product.txt",
        app_dir / "booth_product.txt",
    ]
    return next((path for path in candidates if path.exists()), None)


def read_booth_url(app_dir: Path) -> tuple[str, str, str]:
    product_path = find_booth_product(app_dir)
    if product_path is None:
        return "booth_product_missing", "", ""
    text = read_text(product_path)
    match = re.search(r"(?ms)^# URL\s*\n(.*?)(?=\n# |\Z)", text)
    if not match:
        return "url_section_missing", "", str(product_path.relative_to(app_dir))
    url = match.group(1).strip()
    if not url:
        return "empty", "", str(product_path.relative_to(app_dir))
    return "set", url, str(product_path.relative_to(app_dir))


def read_json_file(path: Path) -> dict[str, Any]:
    return json.loads(read_text(path))


def buffer_update_id(record: dict[str, Any], platform: str) -> str:
    item = record.get("buffer", {}).get(platform, {})
    if not isinstance(item, dict):
        return ""
    for key in ("buffer_post_id", "buffer_update_id", "update_id", "id"):
        if item.get(key):
            return str(item[key])
    return ""


def validate_social_posts(path: Path) -> tuple[bool, bool]:
    if not path.exists():
        return False, False
    text = read_text(path)
    urls = re.findall(r"https?://[^\s)]+", text)
    forbidden = any(("github.com" in url.lower() or "booth.pm" in url.lower() or not url.rstrip(".,").startswith("https://dakeapp.com/apps/")) for url in urls)
    required = ("## X", "## THREADS", "## INSTAGRAM")
    has_sections = all(section in text for section in required)
    return has_sections and not forbidden, forbidden


def apply_release_artifact_flags(app_dir: Path, result: AppCheck) -> None:
    release_dir = app_dir / "release_artifacts"
    release_json = release_dir / "release.json"
    social_release = release_dir / "social_release.json"
    result.demo_video = (release_dir / "demo.mp4").exists() or (release_dir / "demo.webm").exists()
    result.release_json = release_json.exists()
    if release_json.exists():
        try:
            release_record = read_json_file(release_json)
            result.release_stage = str(release_record.get("stage") or "")
            result.video_local_path = str(release_record.get("video_local_path") or "")
            if result.video_local_path:
                result.demo_video = result.demo_video and (app_dir / result.video_local_path).exists()
        except Exception as exc:
            result.release_stage = f"invalid: {exc}"
    social_posts_path = release_dir / "social_posts.md"
    result.social_posts = social_posts_path.exists()
    result.social_posts_valid, result.social_forbidden_links = validate_social_posts(social_posts_path)
    result.social_release = social_release.exists()
    if social_release.exists():
        try:
            social_record = read_json_file(social_release)
            result.social_stage = str(social_record.get("stage") or "")
            result.buffer_x_id = bool(buffer_update_id(social_record, "x"))
            result.buffer_threads_id = bool(buffer_update_id(social_record, "threads"))
            result.buffer_instagram_id = bool(buffer_update_id(social_record, "instagram"))
        except Exception as exc:
            result.social_stage = f"invalid: {exc}"


def check_icon_build(build_text: str) -> bool:
    return "--icon" in build_text and has_common_icon(build_text)


def check_icon_main(main_text: str) -> bool:
    is_tk_app = "tkinter" in main_text or "tk.Tk" in main_text or "tk." in main_text
    is_win32_app = "CreateWindowEx" in main_text or "WNDCLASSEX" in main_text
    if is_win32_app:
        return has_common_icon(main_text) and "LoadImage" in main_text and "SendMessage" in main_text
    if not is_tk_app:
        return True
    if "iconbitmap" not in main_text or not has_common_icon(main_text):
        return False
    icon_positions = [match.start() for match in re.finditer(r"iconbitmap", main_text)]
    for position in icon_positions:
        window = main_text[max(0, position - 700) : min(len(main_text), position + 700)]
        if "try:" in window and "except" in window:
            return True
    return False


def add_action(result: AppCheck, condition: bool, action: str) -> None:
    if not condition:
        result.actions.append(action)


def check_app(app_dir: Path) -> AppCheck:
    result = AppCheck(app=app_dir.name)

    source = read_app_source(app_dir, ROOT)
    result.source_kind = source.source_kind
    result.source_path = source.source_label
    result.original_missing = source.original_missing
    result.meta_derivative_mismatch = ",".join(source.derivative_mismatches)
    readme_path = app_dir / "README.md"
    release_body_path = app_dir / "release_body.md"
    screenshot_path = app_dir / "assets" / "screenshot.webp"
    thumbnail_path = app_dir / "assets" / "booth_thumbnail.jpg"
    booth_ready_dir = app_dir / "booth_ready"
    build_path = app_dir / "build.bat"
    main_path = app_dir / "main.py"
    dist_dir = app_dir / "dist"

    readme_has_release = False
    result.readme = readme_path.exists()
    if result.readme:
        readme_text = read_text(readme_path)
        meta, meta_error = parse_dake_meta(readme_text)
        result.dake_meta = meta is not None
        result.meta_error = meta_error
        readme_has_release = has_release_body(readme_text)
        if meta is not None:
            pass
    else:
        result.meta_error = "README.md missing"

    if source.meta:
        result.status = str(source.meta.get("status") or "unknown")
        result.show_in_launcher = source.meta.get("show_in_launcher")
        result.show_on_site = source.meta.get("show_on_site")
        result.release_url = str(source.meta.get("release_url") or "")
    elif source.error:
        result.meta_error = source.error

    result.release_body = release_body_path.exists()
    result.screenshot = screenshot_path.exists()
    result.thumbnail = thumbnail_path.exists()
    product_path = find_booth_product(app_dir)
    result.booth_product = product_path is not None
    result.booth_url_state, result.booth_url, result.booth_product_path = read_booth_url(app_dir)
    result.booth_ready = booth_ready_dir.exists()
    result.zip_file = booth_ready_dir.exists() and any(booth_ready_dir.glob("*.zip"))
    result.build_bat = build_path.exists()
    result.main_py = main_path.exists()
    result.dist_exe = dist_dir.exists() and any(dist_dir.glob("*.exe"))
    apply_release_artifact_flags(app_dir, result)

    if result.build_bat:
        result.icon_build = check_icon_build(read_text(build_path))

    if result.main_py:
        result.icon_main = check_icon_main(read_text(main_path))

    if result.original_missing:
        result.actions.append("original_missing")
    if result.meta_derivative_mismatch:
        result.actions.append("sync README DAKE_META from ORIGINAL.md: " + result.meta_derivative_mismatch)

    if result.flag_conflict:
        result.actions.append("fix show_in_launcher/show_on_site for non-available status")

    if not result.shipping_candidate:
        if not result.flag_conflict:
            result.actions.append("excluded from regular shipping check")
        return result

    add_action(result, result.readme, "create README.md")
    add_action(result, result.dake_meta, "add or fix DAKE_META JSON")
    if not result.release_body:
        if result.readme and readme_has_release:
            result.actions.append("generate release_body.md from README RELEASE_BODY")
        else:
            result.actions.append("add RELEASE_BODY to README and generate release_body.md")
    add_action(result, result.screenshot, "need screenshot: create assets/screenshot.webp")
    add_action(result, result.thumbnail, "generate assets/booth_thumbnail.jpg with tools/make_booth_ready.py")
    add_action(result, result.booth_product, "generate booth_ready/booth_product.txt with tools/make_booth_ready.py")
    add_action(result, result.booth_ready, "generate booth_ready/ with tools/make_booth_ready.py")
    if not result.zip_file:
        if result.dist_exe:
            result.actions.append("generate BOOTH ready zip from dist/*.exe")
        else:
            result.actions.append("need build: create dist/*.exe before zip")
    add_action(result, result.build_bat, "create build.bat")
    add_action(result, result.main_py, "create main.py")
    add_action(result, result.dist_exe, "need build: create dist/*.exe")
    add_action(result, result.icon_build, "add common --icon setting to build.bat")
    add_action(result, result.icon_main, "add safe common icon setting to main.py")
    add_action(result, result.release_ready, "fill README DAKE_META.release_url after GitHub Release")
    add_action(result, result.release_capture_ready, "capture operation evidence with tools/release_capture.py")
    add_action(result, result.social_draft_ready, "generate SNS dry-run posts with tools/release_social.py")
    add_action(result, result.social_buffer_ready, "create SNS Buffer posts with tools/release_social.py --post-to-buffer")

    if result.booth_url_state == "url_section_missing":
        result.actions.append("add # URL section to booth_product.txt")
    elif result.booth_url_state == "empty":
        result.actions.append("fill BOOTH URL in booth_product.txt after publication")
    elif result.booth_url_state == "booth_product_missing" and result.booth_product:
        result.actions.append("check booth_product.txt URL field")

    return result


def root_gitignore_missing() -> list[str]:
    path = ROOT / ".gitignore"
    if not path.exists():
        return ROOT_GITIGNORE_REQUIRED
    text = read_text(path)
    return [pattern for pattern in ROOT_GITIGNORE_REQUIRED if pattern not in text]


def yn(value: bool) -> str:
    return "OK" if value else "NG"


def status_counts(results: list[AppCheck]) -> Counter[str]:
    return Counter(result.status for result in results)


def available_results(results: list[AppCheck]) -> list[AppCheck]:
    return [result for result in results if result.shipping_candidate]


def non_shipping_results(results: list[AppCheck]) -> list[AppCheck]:
    return [result for result in results if not result.shipping_candidate]


def flag_conflicts(results: list[AppCheck]) -> list[AppCheck]:
    return [result for result in results if result.flag_conflict]


def release_url_empty_available(results: list[AppCheck]) -> list[AppCheck]:
    return [result for result in available_results(results) if not result.release_url]


def site_publish_candidates(results: list[AppCheck]) -> list[AppCheck]:
    return [
        result
        for result in available_results(results)
        if result.show_on_site is True and not result.release_url
    ]


def booth_url_summary(results: list[AppCheck]) -> Counter[str]:
    return Counter(result.booth_url_state for result in available_results(results))


def write_markdown(results: list[AppCheck], missing_gitignore: list[str], path: Path) -> None:
    shipping = available_results(results)
    closed = [result for result in shipping if result.legacy_closed]
    not_closed = [result for result in shipping if not result.legacy_closed]
    excluded = non_shipping_results(results)
    counts = status_counts(results)
    booth_counts = booth_url_summary(results)
    conflicts = flag_conflicts(results)
    release_empty = release_url_empty_available(results)
    site_candidates = site_publish_candidates(results)

    lines: list[str] = [
        "# DAKE BOOTH Ready Check v2",
        "",
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "",
        "## Summary",
        "",
        f"- checked: {len(results)}",
        f"- available checked: {len(shipping)}",
        f"- legacy_closed ok: {len(closed)}",
        f"- legacy_not_closed: {len(not_closed)}",
        f"- v2_closed ok: {sum(result.v2_closed for result in shipping)}",
        f"- v2_pending: {sum(not result.v2_closed for result in shipping)}",
        f"- original_missing: {sum(result.original_missing for result in shipping)}",
        f"- meta_derivative_mismatch: {sum(bool(result.meta_derivative_mismatch) for result in shipping)}",
        f"- excluded by status: {len(excluded)}",
        f"- release_url empty available: {len(release_empty)}",
        f"- show flag conflicts: {len(conflicts)}",
        f"- dakeapp.com publish candidates: {len(site_candidates)}",
        "",
        "## Status Counts",
        "",
    ]
    for status in STATUS_DISPLAY_ORDER:
        lines.append(f"- {status}: {counts.get(status, 0)}")
    for status, count in sorted(counts.items()):
        if status not in set(STATUS_DISPLAY_ORDER):
            lines.append(f"- {status}: {count}")

    lines.extend(
        [
            "",
            "## Available Asset Summary",
            "",
            f"- screenshot.webp missing: {sum(not r.screenshot for r in shipping)}",
            f"- booth_thumbnail.jpg missing: {sum(not r.thumbnail for r in shipping)}",
            f"- booth_product.txt missing: {sum(not r.booth_product for r in shipping)}",
            f"- booth_ready/ missing: {sum(not r.booth_ready for r in shipping)}",
            f"- zip missing: {sum(not r.zip_file for r in shipping)}",
            f"- icon build missing: {sum(not r.icon_build for r in shipping)}",
            f"- icon main missing: {sum(not r.icon_main for r in shipping)}",
            f"- release capture missing: {sum(not r.release_capture_ready for r in shipping)}",
            f"- social draft missing: {sum(not r.social_draft_ready for r in shipping)}",
            f"- social buffer missing: {sum(not r.social_buffer_ready for r in shipping)}",
            "",
            "## BOOTH URL",
            "",
            f"- set: {booth_counts.get('set', 0)}",
            f"- empty: {booth_counts.get('empty', 0)}",
            f"- url section missing: {booth_counts.get('url_section_missing', 0)}",
            f"- booth_product missing: {booth_counts.get('booth_product_missing', 0)}",
            "",
            "## Git Ignore",
            "",
        ]
    )
    if missing_gitignore:
        lines.extend(f"- missing: `{pattern}`" for pattern in missing_gitignore)
    else:
        lines.append("- OK")

    lines.extend(["", "## Show Flag Conflicts", ""])
    if conflicts:
        lines.extend(
            [
                "| app | status | show_in_launcher | show_on_site |",
                "| --- | --- | --- | --- |",
            ]
        )
        for result in conflicts:
            lines.append(
                f"| {result.app} | {result.status} | {result.show_in_launcher} | {result.show_on_site} |"
            )
    else:
        lines.append("- none")

    lines.extend(["", "## Release URL Empty Available", ""])
    if release_empty:
        for result in release_empty:
            lines.append(f"- {result.app}")
    else:
        lines.append("- none")

    lines.extend(["", "## dakeapp.com Publish Candidates", ""])
    if site_candidates:
        for result in site_candidates:
            lines.append(f"- {result.app}")
    else:
        lines.append("- none")

    lines.extend(
        [
            "",
            "## Legacy Not Closed Available Apps",
            "",
            "| app | asset_ready | release_capture | release_url | booth_url | social | next_action |",
            "| --- | --- | --- | --- | --- | --- | --- |",
        ]
    )
    for result in not_closed:
        action = "<br>".join(result.actions) if result.actions else "-"
        lines.append(
            "| "
            + " | ".join(
                [
                    result.app,
                    yn(result.asset_ready),
                    yn(result.release_capture_ready),
                    yn(result.release_ready),
                    result.booth_url_state,
                    yn(result.social_ready),
                    action,
                ]
            )
            + " |"
        )

    lines.extend(
        [
            "",
            "## Excluded By Status",
            "",
            "| app | status | show_in_launcher | show_on_site | note |",
            "| --- | --- | --- | --- | --- |",
        ]
    )
    for result in excluded:
        note = "<br>".join(result.actions) if result.actions else "excluded"
        lines.append(
            f"| {result.app} | {result.status} | {result.show_in_launcher} | {result.show_on_site} | {note} |"
        )

    lines.extend(["", "## V2 Readiness", "", "| app | legacy_closed | release_capture_ready | social_draft_ready | social_buffer_ready | v2_closed | pending_reason |", "| --- | --- | --- | --- | --- | --- | --- |"])
    for result in shipping:
        lines.append(f"| {result.app} | {yn(result.legacy_closed)} | {yn(result.release_capture_ready)} | {yn(result.social_draft_ready)} | {yn(result.social_buffer_ready)} | {yn(result.v2_closed)} | {result.v2_pending_reason or '-'} |")

    lines.extend(["", "## Legacy Closed OK", "", "| app | status |", "| --- | --- |"])
    for result in closed:
        lines.append(f"| {result.app} | LEGACY_CLOSED |")

    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def write_csv(results: list[AppCheck], path: Path) -> None:
    fields = [
        "app",
        "status",
        "shipping_candidate",
        "legacy_closed",
        "v2_closed",
        "asset_ready",
        "release_url",
        "booth_url_state",
        "booth_url",
        "show_in_launcher",
        "show_on_site",
        "flag_conflict",
        "readme",
        "dake_meta",
        "release_body",
        "screenshot",
        "thumbnail",
        "booth_product",
        "booth_ready",
        "zip",
        "build_bat",
        "main_py",
        "dist_exe",
        "icon_build",
        "icon_main",
        "release_json",
        "release_stage",
        "video_local_path",
        "demo_video",
        "social_posts",
        "social_posts_valid",
        "social_forbidden_links",
        "social_release",
        "social_stage",
        "social_draft_ready",
        "social_buffer_ready",
        "v2_pending_reason",
        "buffer_x_id",
        "buffer_threads_id",
        "buffer_instagram_id",
        "source_kind",
        "source_path",
        "original_missing",
        "meta_derivative_mismatch",
        "actions",
    ]
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        for result in results:
            writer.writerow(
                {
                    "app": result.app,
                    "status": result.status,
                    "shipping_candidate": result.shipping_candidate,
                    "legacy_closed": result.legacy_closed,
                    "v2_closed": result.v2_closed,
                    "asset_ready": result.asset_ready,
                    "release_url": result.release_url,
                    "booth_url_state": result.booth_url_state,
                    "booth_url": result.booth_url,
                    "show_in_launcher": result.show_in_launcher,
                    "show_on_site": result.show_on_site,
                    "flag_conflict": result.flag_conflict,
                    "readme": result.readme,
                    "dake_meta": result.dake_meta,
                    "release_body": result.release_body,
                    "screenshot": result.screenshot,
                    "thumbnail": result.thumbnail,
                    "booth_product": result.booth_product,
                    "booth_ready": result.booth_ready,
                    "zip": result.zip_file,
                    "build_bat": result.build_bat,
                    "main_py": result.main_py,
                    "dist_exe": result.dist_exe,
                    "icon_build": result.icon_build,
                    "icon_main": result.icon_main,
                    "release_json": result.release_json,
                    "release_stage": result.release_stage,
                    "video_local_path": result.video_local_path,
                    "demo_video": result.demo_video,
                    "social_posts": result.social_posts,
                    "social_posts_valid": result.social_posts_valid,
                    "social_forbidden_links": result.social_forbidden_links,
                    "social_release": result.social_release,
                    "social_stage": result.social_stage,
                    "social_draft_ready": result.social_draft_ready,
                    "social_buffer_ready": result.social_buffer_ready,
                    "v2_pending_reason": result.v2_pending_reason,
                    "buffer_x_id": result.buffer_x_id,
                    "buffer_threads_id": result.buffer_threads_id,
                    "buffer_instagram_id": result.buffer_instagram_id,
                    "source_kind": result.source_kind,
                    "source_path": result.source_path,
                    "original_missing": result.original_missing,
                    "meta_derivative_mismatch": result.meta_derivative_mismatch,
                    "actions": "; ".join(result.actions),
                }
            )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Check DAKE formal shipping assets.")
    parser.add_argument("--report-dir", type=Path, default=REPORT_DIR)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    report_dir = args.report_dir
    report_dir.mkdir(parents=True, exist_ok=True)

    app_dirs = sorted(path for path in APPS_DIR.iterdir() if path.is_dir() and path.name.startswith("DAKE_"))
    results = [check_app(app_dir) for app_dir in app_dirs]
    missing_gitignore = root_gitignore_missing()

    md_path = report_dir / "booth_ready_check.md"
    csv_path = report_dir / "booth_ready_check.csv"
    write_markdown(results, missing_gitignore, md_path)
    write_csv(results, csv_path)

    shipping = available_results(results)
    closed_count = sum(result.legacy_closed for result in shipping)
    not_closed_count = len(shipping) - closed_count
    counts = status_counts(results)
    booth_counts = booth_url_summary(results)
    conflicts = flag_conflicts(results)
    release_empty = release_url_empty_available(results)
    site_candidates = site_publish_candidates(results)

    print("DAKE BOOTH Ready Check v2")
    print(f"checked: {len(results)}")
    print("status counts:")
    for status in STATUS_DISPLAY_ORDER:
        print(f"  {status}: {counts.get(status, 0)}")
    print(f"available checked: {len(shipping)}")
    print(f"legacy_closed ok: {closed_count}")
    print(f"v2_closed ok: {sum(result.v2_closed for result in shipping)}")
    print(f"legacy_not_closed: {not_closed_count}")
    print(f"v2_pending: {sum(not result.v2_closed for result in shipping)}")
    print(f"screenshot.webp missing: {sum(not r.screenshot for r in shipping)}")
    print(f"booth_thumbnail.jpg missing: {sum(not r.thumbnail for r in shipping)}")
    print(f"booth_product.txt missing: {sum(not r.booth_product for r in shipping)}")
    print(f"booth_ready/ missing: {sum(not r.booth_ready for r in shipping)}")
    print(f"zip missing: {sum(not r.zip_file for r in shipping)}")
    print(f"icon build missing: {sum(not r.icon_build for r in shipping)}")
    print(f"icon main missing: {sum(not r.icon_main for r in shipping)}")
    print(f"release capture missing: {sum(not r.release_capture_ready for r in shipping)}")
    print(f"social draft missing: {sum(not r.social_draft_ready for r in shipping)}")
    print(f"social buffer missing: {sum(not r.social_buffer_ready for r in shipping)}")
    print(f"original_missing: {sum(r.original_missing for r in shipping)}")
    print(f"meta_derivative_mismatch: {sum(bool(r.meta_derivative_mismatch) for r in shipping)}")
    print("BOOTH URL:")
    print(f"  set: {booth_counts.get('set', 0)}")
    print(f"  empty: {booth_counts.get('empty', 0)}")
    print(f"  url section missing: {booth_counts.get('url_section_missing', 0)}")
    print(f"  booth_product missing: {booth_counts.get('booth_product_missing', 0)}")
    print(f"release_url empty available: {len(release_empty)}")
    if release_empty:
        print("release_url empty apps: " + ", ".join(result.app for result in release_empty))
    print(f"show flag conflicts: {len(conflicts)}")
    if conflicts:
        for result in conflicts:
            print(
                f"- {result.app}: status={result.status}, "
                f"show_in_launcher={result.show_in_launcher}, show_on_site={result.show_on_site}"
            )
    print(f"dakeapp.com publish candidates: {len(site_candidates)}")
    if site_candidates:
        print("dakeapp.com candidates: " + ", ".join(result.app for result in site_candidates))
    if missing_gitignore:
        print("gitignore missing: " + ", ".join(missing_gitignore))
    else:
        print("gitignore: OK")
    print(f"report: {md_path}")
    print(f"csv: {csv_path}")
    if not_closed_count:
        print("legacy not closed available apps:")
        for result in shipping:
            if not result.legacy_closed:
                print(f"- {result.app}: {', '.join(result.actions)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
