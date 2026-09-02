# -*- coding: utf-8 -*-
"""Release-preparation checks derived from ORIGINAL.md."""

from __future__ import annotations

import json
import re
from pathlib import Path

from PIL import Image


APP_DIR = Path(__file__).resolve().parents[1]

EXPECTED_LICENSE_FILES = {
    "BUILD_LICENSES/abseil.txt",
    "BUILD_LICENSES/agg23.txt",
    "BUILD_LICENSES/fast_float.txt",
    "BUILD_LICENSES/freetype.txt",
    "BUILD_LICENSES/icu.txt",
    "BUILD_LICENSES/lcms.txt",
    "BUILD_LICENSES/libjpeg_turbo.ijg",
    "BUILD_LICENSES/libjpeg_turbo.md",
    "BUILD_LICENSES/libopenjpeg.txt",
    "BUILD_LICENSES/libpng.txt",
    "BUILD_LICENSES/libtiff.txt",
    "BUILD_LICENSES/llvm-libc.txt",
    "BUILD_LICENSES/pdfium-binaries.txt",
    "BUILD_LICENSES/pdfium.txt",
    "BUILD_LICENSES/simdutf.txt",
    "BUILD_LICENSES/zlib.txt",
    "LICENSES/Apache-2.0.txt",
    "LICENSES/BSD-3-Clause.txt",
    "LICENSES/CC-BY-4.0.txt",
}


def _read(relative: str) -> str:
    return (APP_DIR / relative).read_text(encoding="utf-8")


def _section(text: str, heading: str) -> str:
    match = re.search(rf"(?ms)^## {re.escape(heading)}\s*\n(.*?)(?=^## |\Z)", text)
    assert match, heading
    return match.group(1).strip()


def _meta(text: str) -> dict[str, object]:
    block = _section(text, "DAKE_META")
    match = re.search(r"(?s)```json\s*(\{.*?\})\s*```", block)
    assert match
    return json.loads(match.group(1))


def test_readme_meta_matches_original_and_keeps_release_unpublished() -> None:
    meta = _meta(_read("README.md"))
    original_block = _section(_read("ORIGINAL.md"), "DAKE_META生成用情報")
    original_match = re.search(r"(?s)```json\s*(\{.*?\})\s*```", original_block)
    assert original_match
    assert meta == json.loads(original_match.group(1))
    assert meta["app_type"] == "market"
    assert meta["completion_goal"] == "formal_release"
    assert meta["status"] == "draft"
    assert meta["release_url"] == ""
    assert meta["show_in_launcher"] is False
    assert meta["show_on_site"] is False


def test_release_body_matches_readme_derived_view() -> None:
    readme_body = _section(_read("README.md"), "RELEASE_BODY")
    assert _read("release_body.md").strip() == readme_body
    bullets = [line for line in readme_body.splitlines() if line.startswith("- ")]
    assert 3 <= len(bullets) <= 5


def test_booth_views_keep_price_and_urls_unset() -> None:
    canonical = _read("booth_product.txt")
    ready = _read("booth_ready/booth_product.txt")
    assert canonical == ready
    for heading in ("価格案", "GitHub Release", "URL"):
        match = re.search(rf"(?ms)^# {re.escape(heading)}\s*\n(.*?)(?=^# |\Z)", canonical)
        assert match, heading
        assert match.group(1).strip() == ""


def test_wheel_license_file_set_is_complete() -> None:
    root = APP_DIR / "third_party_licenses" / "pypdfium2-5.13.0"
    actual = {path.relative_to(root).as_posix() for path in root.rglob("*") if path.is_file()}
    assert actual == EXPECTED_LICENSE_FILES
    assert all((root / relative).stat().st_size > 0 for relative in actual)
    ready_root = APP_DIR / "booth_ready" / "third_party_licenses" / "pypdfium2-5.13.0"
    ready_actual = {
        path.relative_to(ready_root).as_posix()
        for path in ready_root.rglob("*")
        if path.is_file()
    }
    assert ready_actual == EXPECTED_LICENSE_FILES
    assert all((ready_root / relative).read_bytes() == (root / relative).read_bytes() for relative in actual)


def test_release_images_follow_dake_dimensions() -> None:
    with Image.open(APP_DIR / "assets" / "screenshot.webp") as screenshot:
        assert screenshot.format == "WEBP"
        assert screenshot.size == (1182, 812)
        assert screenshot.width <= 1200
    with Image.open(APP_DIR / "assets" / "screenshot.jpg") as screenshot_jpg:
        assert screenshot_jpg.format == "JPEG"
        assert screenshot_jpg.size == (1182, 812)
    with Image.open(APP_DIR / "assets" / "booth_thumbnail.jpg") as thumbnail:
        assert thumbnail.format == "JPEG"
        assert thumbnail.size == (1200, 1200)


def test_booth_ready_copies_match_release_assets() -> None:
    for relative in (
        "screenshot.jpg",
        "booth_thumbnail.jpg",
        "THIRD_PARTY_NOTICES.txt",
    ):
        source = APP_DIR / (relative if relative == "THIRD_PARTY_NOTICES.txt" else f"assets/{relative}")
        assert (APP_DIR / "booth_ready" / relative).read_bytes() == source.read_bytes()


def test_public_assets_contain_no_local_user_paths() -> None:
    public_texts = [
        "README.md",
        "release_body.md",
        "booth_product.txt",
        "THIRD_PARTY_NOTICES.txt",
        "booth_ready/README.txt",
        "booth_ready/注意事項.txt",
        "booth_ready/booth_product.txt",
    ]
    combined = "\n".join(_read(path) for path in public_texts).lower()
    assert "c:\\users\\" not in combined
    assert "appdata\\local" not in combined
