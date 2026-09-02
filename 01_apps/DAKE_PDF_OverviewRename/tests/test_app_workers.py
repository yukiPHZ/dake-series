# -*- coding: utf-8 -*-
from __future__ import annotations

import ast
import queue
import threading
import time
from pathlib import Path

import main
import pytest
from main import RenderPool, RenderRequest, scan_pdf_folder
from rename_core import FileSnapshot


def make_snapshot(tmp_path: Path, name: str) -> FileSnapshot:
    path = tmp_path / name
    path.write_bytes(b"fixture")
    return FileSnapshot.capture(path)


def test_scan_is_sorted_non_recursive_and_pdf_only(tmp_path: Path) -> None:
    (tmp_path / "subfolder").mkdir()
    (tmp_path / "subfolder" / "hidden.pdf").write_bytes(b"hidden")
    (tmp_path / "ignore.txt").write_text("ignore", encoding="utf-8")
    for name in ("b.PDF", "A.pdf", "c.pdf"):
        (tmp_path / name).write_bytes(name.encode())
    assert [item.path.name for item in scan_pdf_folder(tmp_path)] == ["A.pdf", "b.PDF", "c.pdf"]


def test_render_pool_replaces_not_started_old_jobs(monkeypatch, tmp_path: Path) -> None:
    active = make_snapshot(tmp_path, "active.pdf")
    old_one = make_snapshot(tmp_path, "old_1.pdf")
    old_two = make_snapshot(tmp_path, "old_2.pdf")
    latest = make_snapshot(tmp_path, "latest.pdf")
    started = threading.Event()
    release = threading.Event()
    calls: list[str] = []

    def fake_render(path: Path, _box: tuple[int, int]):
        calls.append(path.name)
        if path.name == "active.pdf":
            started.set()
            assert release.wait(3)
        return object(), 1

    monkeypatch.setattr(main, "render_first_page", fake_render)
    pool = RenderPool(worker_count=1)
    try:
        pool.replace(1, [RenderRequest(1, "thumbnail", 1, active, (10, 10))])
        assert started.wait(2)
        pool.replace(
            2,
            [
                RenderRequest(2, "thumbnail", 2, old_one, (10, 10)),
                RenderRequest(2, "thumbnail", 3, old_two, (10, 10)),
            ],
        )
        pool.replace(3, [RenderRequest(3, "thumbnail", 4, latest, (10, 10))])
        release.set()
        results = [pool.results.get(timeout=3), pool.results.get(timeout=3)]
        assert {result.request.identifier for result in results} == {1, 4}
        assert calls == ["active.pdf", "latest.pdf"]
        with pytest.raises(queue.Empty):
            pool.results.get(timeout=0.1)
    finally:
        release.set()
        pool.shutdown()


def test_render_failure_is_isolated(monkeypatch, tmp_path: Path) -> None:
    broken = make_snapshot(tmp_path, "broken.pdf")
    valid = make_snapshot(tmp_path, "valid.pdf")

    def fake_render(path: Path, _box: tuple[int, int]):
        if path.name == "broken.pdf":
            raise ValueError("broken")
        return object(), 1

    monkeypatch.setattr(main, "render_first_page", fake_render)
    pool = RenderPool(worker_count=1)
    try:
        pool.replace(
            1,
            [
                RenderRequest(1, "thumbnail", 1, broken, (10, 10)),
                RenderRequest(1, "thumbnail", 2, valid, (10, 10)),
            ],
        )
        first = pool.results.get(timeout=3)
        second = pool.results.get(timeout=3)
        assert first.error is not None
        assert second.error is None and second.image is not None
    finally:
        pool.shutdown()


def test_ui_widget_text_has_no_inline_japanese() -> None:
    source = Path(main.__file__).read_text(encoding="utf-8")
    tree = ast.parse(source)
    offenders: list[tuple[int, str]] = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        for keyword in node.keywords:
            if keyword.arg != "text" or not isinstance(keyword.value, ast.Constant):
                continue
            value = keyword.value.value
            if isinstance(value, str) and any(ord(character) > 127 for character in value):
                offenders.append((node.lineno, value))
    assert offenders == []


def test_build_uses_common_icon_for_exe_and_onefile_data() -> None:
    build = Path(main.__file__).with_name("build.bat").read_text(encoding="utf-8")
    assert "--onefile" in build
    assert "--noconsole" in build
    assert "--icon=..\\..\\02_assets\\dake_icon.ico" in build
    assert "--add-data=..\\..\\02_assets\\dake_icon.ico;." in build
