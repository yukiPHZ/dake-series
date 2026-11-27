import struct
import sys
from pathlib import Path


sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "00_core"))

from dake_quality_engine.quality_check import (
    check_icons,
    check_original_md,
    detect_mojibake_lines,
    parse_ico_resolutions,
)
from dake_quality_engine.ui_guard import check_source


def test_current_footer_is_accepted() -> None:
    source = """
APP_NAME = "A"
WINDOW_TITLE = "A"
KAITORI_URL = "https://dakeapp.com/go/kaitori?app=dake_sample"
UI_TEXT = {
    "footer_kaitori": "戸建買取査定",
    "footer_author": "by Yukihiko Kikuta",
}
"""

    assert check_source(source) == []


def test_legacy_footer_keys_fail() -> None:
    source = """
APP_NAME = "A"
WINDOW_TITLE = "A"
KAITORI_URL = "https://dakeapp.com/go/kaitori?app=dake_sample"
UI_TEXT = {
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_link_1": "戸建買取査定",
    "footer_author": "by Yukihiko Kikuta",
}
"""

    fails = [issue for issue in check_source(source) if issue.level == "FAIL"]
    assert any("footer_left" in issue.message for issue in fails)
    assert any("footer_link_1" in issue.message for issue in fails)


def test_forbidden_strings_fail() -> None:
    source = """
APP_NAME = "A"
WINDOW_TITLE = "A"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"
KAITORI_URL = "https://dakeapp.com/go/kaitori?app=dake_sample"
OLD_URL = "https://sakurayk.notion.site/12345"
SNS = "Instagram"
UI_TEXT = {
    "footer_kaitori": "戸建買取査定",
    "footer_author": "by Yukihiko Kikuta",
}
"""

    fails = [issue.message for issue in check_source(source) if issue.level == "FAIL"]
    assert any("Vibe-Coded" in m for m in fails)
    assert any("Instagram" in m for m in fails)
    assert any("sakurayk.notion.site" in m for m in fails)
    assert any("しまりす不動産" in m for m in fails)


def test_missing_canonical_footer_parts_fail() -> None:
    source = """
APP_NAME = "A"
WINDOW_TITLE = "A"
UI_TEXT = {
    "title": "サンプル",
}
"""

    fails = [issue.message for issue in check_source(source) if issue.level == "FAIL"]
    assert any("戸建買取査定" in m for m in fails)
    assert any("by Yukihiko Kikuta" in m for m in fails)
    assert any("https://dakeapp.com/go/kaitori?app=<app_id>" in m for m in fails)


def test_mojibake_detection_on_original_md(tmp_path: Path) -> None:
    bad = tmp_path / "ORIGINAL.md"
    bad.write_text("# ORIGINAL.md\n- Store雋ｩ螢ｲ迥ｶ諷・ stripe_ready\n", encoding="utf-8")
    res_bad = check_original_md(tmp_path)
    assert any(r.level == "FAIL" for r in res_bad)

    bad.write_text("# ORIGINAL.md\n- Store販売状態: stripe_ready\n", encoding="utf-8")
    res_ok = check_original_md(tmp_path)
    assert all(r.level == "OK" for r in res_ok)
    assert detect_mojibake_lines("壊れた文字: \ufffd") != []


def test_icon_checks_detect_missing_and_bad_ico(tmp_path: Path) -> None:
    (tmp_path / "main.py").write_text("import tkinter as tk\nroot = tk.Tk()\n", encoding="utf-8")
    (tmp_path / "build.bat").write_text("pyinstaller --onefile --noconsole main.py\n", encoding="utf-8")
    # Single-size ICO (missing 32, 48, 256)
    assets = tmp_path / "assets"
    assets.mkdir()
    single_ico = assets / "app.ico"
    single_ico.write_bytes(struct.pack("<HHH", 0, 1, 1) + struct.pack("<BBBBHHII", 16, 16, 0, 0, 1, 32, 40, 22))
    # Dummy exe without PE header
    dist = tmp_path / "dist"
    dist.mkdir()
    fake_exe = dist / "Fake.exe"
    fake_exe.write_bytes(b"NOT_A_PE_FILE")

    entries, err = parse_ico_resolutions(single_ico)
    assert err is None and entries == [(16, 16, 32)]

    results = check_icons(tmp_path, [fake_exe])
    failed_items = {r.item for r in results if r.level == "FAIL"}
    assert "Icon (build --icon)" in failed_items
    assert "Icon (.ico resolutions)" in failed_items
    assert "Icon (main.py window)" in failed_items
    assert "Icon (EXE resource: Fake.exe)" in failed_items
