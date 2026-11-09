import sys
from pathlib import Path


sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from release_capture import StepState, ship_failures


def test_distribution_checks_do_not_block_ship() -> None:
    checks = {
        "ORIGINAL / DAKE_META": StepState(True, "ok"),
        "アプリ起動確認": StepState(True, "ok"),
        "主要操作の確認": StepState(True, "ok"),
        "build / dist 生成": StepState(True, "ok"),
        "GitHub Release": StepState(True, "ok"),
        "screenshot.webp 保存": StepState(False, "missing"),
        "BOOTH URL確認": StepState(False, "missing"),
        "dakeapp.com詳細ページ更新": StepState(False, "missing"),
        "Cloudflare本番反映確認": StepState(False, "not requested"),
        "操作動画保存": StepState(False, "not requested"),
    }

    assert ship_failures(checks) == []


def test_ship_failure_is_limited_to_ship_gate() -> None:
    checks = {
        "アプリ起動確認": StepState(False, "exe did not start"),
        "BOOTH URL確認": StepState(False, "missing"),
    }

    assert ship_failures(checks) == ["アプリ起動確認: exe did not start"]
