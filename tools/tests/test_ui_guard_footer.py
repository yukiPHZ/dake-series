import sys
from pathlib import Path


sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "00_core"))

from dake_quality_engine.ui_guard import check_source


def test_current_footer_is_accepted() -> None:
    source = """
APP_NAME = "A"
WINDOW_TITLE = "A"
KAITORI_URL = "https://dakeapp.com/go/kaitori?app=a"
UI_TEXT = {
    "footer_kaitori": "戸建買取査定",
    "footer_author": "by Yukihiko Kikuta",
}
"""

    assert check_source(source) == []


def test_legacy_footer_remains_compatible_during_incremental_migration() -> None:
    source = """
APP_NAME = "A"
WINDOW_TITLE = "A"
COPYRIGHT = "legacy"
UI_TEXT = {
    "footer_left": "legacy",
    "footer_link_1": "legacy",
}
"""

    assert not [issue for issue in check_source(source) if issue.level == "NG"]
