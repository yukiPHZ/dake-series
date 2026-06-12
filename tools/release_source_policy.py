from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


RELEASE_META_DEFAULTS = {
    "demo_video_path": "release_artifacts/demo.mp4",
    "demo_video_url": "",
    "social_release_path": "release_artifacts/social_release.json",
}
RELEASE_META_KEYS = tuple(RELEASE_META_DEFAULTS)
SOURCE_POLICY_TEXT = (
    "ORIGINAL.md is the source of truth. README.md and DAKE_META are derived "
    "views. Generated JSON must not be edited manually."
)
SLUG_OVERRIDES = {
    "game_alien_road": "alien-road",
    "game_diver_catch": "diver-catch",
    "time_advanced_timer": "advanced-timer",
    "dake_pdf_checkstamp": "check-mark",
    "dake_image_heictojpg": "heic-to-jpg",
    "dake_image_iphonetopc": "iphone-to-pc",
}


@dataclass
class AppSource:
    app_dir: Path
    meta: dict[str, Any] = field(default_factory=dict)
    source_kind: str = "missing"
    source_path: Path | None = None
    source_label: str = ""
    original_exists: bool = False
    original_missing: bool = False
    readme_exists: bool = False
    readme_meta: dict[str, Any] = field(default_factory=dict)
    error: str = ""
    original_error: str = ""
    readme_error: str = ""
    derivative_mismatches: list[str] = field(default_factory=list)

    @property
    def status(self) -> str:
        return str(self.meta.get("status") or "unknown")

    @property
    def display_name(self) -> str:
        return str(
            self.meta.get("display_name")
            or self.meta.get("site_title")
            or self.meta.get("launcher_title")
            or self.app_dir.name
        )


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return path.read_text(encoding="utf-8", errors="replace")


def extract_json_section(text: str, heading: str) -> dict[str, Any] | None:
    if heading == "DAKE_META":
        heading_pattern = r"(?m)^\s*##\s+DAKE_META[^\n]*\n"
    else:
        heading_pattern = rf"(?m)^\s*##\s+{re.escape(heading)}\s*\n"
    heading_match = re.search(heading_pattern, text)
    if not heading_match:
        return None
    section_start = heading_match.end()
    next_heading = re.search(r"(?m)^\s*##\s+", text[section_start:])
    section_end = section_start + next_heading.start() if next_heading else len(text)
    section = text[section_start:section_end]
    code_match = re.search(r"(?s)```json\s*(.*?)\s*```", section)
    if not code_match:
        return None
    loaded = json.loads(code_match.group(1))
    return loaded if isinstance(loaded, dict) else None


def read_json_heading(path: Path, headings: tuple[str, ...]) -> tuple[dict[str, Any], str]:
    if not path.exists():
        return {}, f"{path.name} missing"
    text = read_text(path)
    for heading in headings:
        try:
            meta = extract_json_section(text, heading)
        except json.JSONDecodeError as exc:
            return {}, f"invalid {heading} JSON: {exc}"
        if meta is not None:
            return meta, ""
    return {}, f"{'/'.join(headings)} missing"


def with_release_defaults(meta: dict[str, Any]) -> dict[str, Any]:
    normalized = dict(meta)
    for key, value in RELEASE_META_DEFAULTS.items():
        normalized.setdefault(key, value)
    return normalized


def source_label(path: Path, root: Path) -> str:
    try:
        return path.relative_to(root).as_posix()
    except ValueError:
        return path.as_posix()


def read_app_source(app_dir: Path, root: Path | None = None) -> AppSource:
    root = root or app_dir.parents[1]
    original = app_dir / "ORIGINAL.md"
    readme = app_dir / "README.md"
    original_meta: dict[str, Any] = {}
    readme_meta: dict[str, Any] = {}
    original_error = ""
    readme_error = ""

    if original.exists():
        original_meta, original_error = read_json_heading(original, ("DAKE_META生成用情報", "DAKE_META"))
    if readme.exists():
        readme_meta, readme_error = read_json_heading(readme, ("DAKE_META",))

    source = AppSource(
        app_dir=app_dir,
        original_exists=original.exists(),
        original_missing=not original.exists(),
        readme_exists=readme.exists(),
        readme_meta=with_release_defaults(readme_meta) if readme_meta else {},
        original_error=original_error,
        readme_error=readme_error,
    )

    if original_meta:
        source.meta = with_release_defaults(original_meta)
        source.source_kind = "original"
        source.source_path = original
        source.source_label = source_label(original, root)
        for key in RELEASE_META_KEYS:
            if readme_meta and str(source.meta.get(key, "")) != str(with_release_defaults(readme_meta).get(key, "")):
                source.derivative_mismatches.append(key)
        return source

    if original.exists() and original_error:
        source.error = original_error
    if readme_meta:
        source.meta = with_release_defaults(readme_meta)
        source.source_kind = "readme_fallback"
        source.source_path = readme
        source.source_label = source_label(readme, root)
        if not source.error:
            source.error = "original_missing" if source.original_missing else original_error
        return source

    source.meta = {}
    source.source_kind = "missing"
    source.source_path = None
    if source.error and readme_error:
        source.error = f"{source.error}; {readme_error}"
    elif readme_error:
        source.error = readme_error
    elif not source.error:
        source.error = "metadata missing"
    return source


def app_dirs(apps_dir: Path) -> list[Path]:
    return sorted(path for path in apps_dir.iterdir() if path.is_dir() and path.name.startswith("DAKE_"))


def find_app(apps_dir: Path, identifier: str, root: Path | None = None) -> Path:
    needle = identifier.lower()
    for app_dir in app_dirs(apps_dir):
        if app_dir.name.lower() == needle:
            return app_dir
        source = read_app_source(app_dir, root)
        values = [
            str(source.meta.get("app_key") or ""),
            str(source.meta.get("display_name") or ""),
            str(source.meta.get("folder_name") or ""),
        ]
        if any(value.lower() == needle for value in values if value):
            return app_dir
    raise SystemExit(f"app not found: {identifier}")


def camel_slug(value: str) -> str:
    value = value.strip()
    value = re.sub(r"^DAKE_", "", value, flags=re.IGNORECASE)
    value = value.replace("HEICtoJPG", "HEIC_To_JPG")
    value = value.replace("iPhoneToPC", "iPhone_To_PC")
    value = value.replace("CheckStamp", "Check_Stamp")
    value = value.replace("AdvancedTimer", "Advanced_Timer")
    value = value.replace("SplitOne", "Split_One").replace("SplitSelect", "Split_Select")
    value = re.sub(r"([a-z0-9])([A-Z])", r"\1_\2", value)
    value = value.replace("_", "-").lower()
    value = re.sub(r"[^a-z0-9-]+", "-", value)
    value = re.sub(r"-+", "-", value).strip("-")
    return value


def site_slug(app_dir: Path, meta: dict[str, Any], site_root: Path) -> str:
    explicit = str(meta.get("site_slug") or "").strip()
    if explicit:
        return explicit
    apps_root = site_root / "public" / "apps"
    override = SLUG_OVERRIDES.get(str(meta.get("app_key") or "").strip())
    if override and (apps_root / override / "index.html").exists():
        return override
    release_url = str(meta.get("release_url") or "").strip()
    if release_url and apps_root.exists():
        for page in sorted(apps_root.glob("*/index.html")):
            try:
                if release_url in read_text(page):
                    return page.parent.name
            except Exception:
                pass
    candidates = [
        camel_slug(app_dir.name),
        camel_slug(str(meta.get("app_key") or "")),
        camel_slug(str(meta.get("folder_name") or "")),
    ]
    for candidate in candidates:
        if candidate and (apps_root / candidate / "index.html").exists():
            return candidate
    return next((candidate for candidate in candidates if candidate), camel_slug(app_dir.name))


def app_url_for(app_dir: Path, meta: dict[str, Any], site_root: Path) -> str:
    explicit = str(meta.get("app_url") or "").strip()
    if explicit:
        return explicit
    return f"https://dakeapp.com/apps/{site_slug(app_dir, meta, site_root)}/"
