"""v1.0.3 DnD, unneeded-name and keyboard-preview regressions."""

from __future__ import annotations

import tkinter as tk
import time
import hashlib
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import main
import pytest
from rename_core import FileSnapshot, RenameRequest, RenameTransactionError, build_rename_plan, rename_batch, undo_rename
from viewport_helpers import settle_view


class Variable:
    def __init__(self, value: str):
        self.value = value

    def get(self) -> str:
        return self.value


def card(identifier: int, name: str, checked: bool = False, value: str | None = None):
    return SimpleNamespace(
        identifier=identifier, original_name=name, variable=Variable(value or Path(name).stem), unneeded=checked
    )


def tk_root(factory=tk.Tk):
    for attempt in range(5):
        try:
            return factory()
        except tk.TclError:
            if attempt == 4:
                raise
            time.sleep(.2)


def test_unneeded_allocates_in_display_order_and_reserves_manual_targets() -> None:
    cards = [
        card(0, "scan1.pdf", True),
        card(1, "scan2.pdf", True),
        card(2, "manual.pdf", False, "不要2"),
        card(3, "scan3.pdf", True),
    ]
    disk = {item.original_name for item in cards} | {"不要.pdf", "不要1.pdf"}
    assert main.allocate_unneeded_names(cards, disk) == {0: "不要3", 1: "不要4", 3: "不要5"}
    cards[1].unneeded = False
    assert main.allocate_unneeded_names(cards, disk) == {0: "不要3", 3: "不要4"}


def test_unneeded_uses_vacated_name_and_noop_is_not_renamed(tmp_path: Path) -> None:
    first = tmp_path / "不要.pdf"
    second = tmp_path / "scan.pdf"
    first.write_bytes(b"one")
    second.write_bytes(b"two")
    cards = [card(0, first.name, True), card(1, second.name, True)]
    assigned = main.allocate_unneeded_names(cards, {first.name, second.name})
    assert assigned == {0: "不要", 1: "不要1"}
    requests = [RenameRequest(FileSnapshot.capture(path), assigned[i]) for i, path in enumerate((first, second))]
    assert len(build_rename_plan(requests).entries) == 1
    _, record = rename_batch(requests)
    assert (tmp_path / "不要1.pdf").read_bytes() == b"two"
    undo_rename(record)
    assert first.read_bytes() == b"one" and second.read_bytes() == b"two"


def test_unneeded_100_checks_stay_unique_with_casefolded_collisions() -> None:
    cards = [card(i, f"scan{i:03d}.pdf", True) for i in range(100)]
    assigned = main.allocate_unneeded_names(cards, {"不 要.pdf", "不要.PDF", "不要1.pdf"})
    assert len(assigned) == 100
    assert assigned[0] == "不要2" and assigned[99] == "不要101"
    assert len({name.casefold() for name in assigned.values()}) == 100


@pytest.mark.parametrize("checked_count", [1, 10, 100])
def test_unneeded_mixed_batch_hashes_undo_and_rollback(tmp_path: Path, checked_count: int) -> None:
    manual = tmp_path / "manual.pdf"
    manual.write_bytes(b"manual-data")
    original_paths = []
    for i in range(checked_count):
        path = tmp_path / f"scan{i:03d}.pdf"
        path.write_bytes(f"pdf-{i}".encode())
        original_paths.append(path)
    source_hashes = {path.name: hashlib.sha256(path.read_bytes()).hexdigest() for path in (manual, *original_paths)}
    cards = [card(0, manual.name, False, "重要事項説明書")]
    cards.extend(card(i + 1, path.name, True) for i, path in enumerate(original_paths))
    assigned = main.allocate_unneeded_names(cards, {path.name for path in (manual, *original_paths)})
    requests = [RenameRequest(FileSnapshot.capture(manual), "重要事項説明書")]
    requests.extend(RenameRequest(FileSnapshot.capture(path), assigned[i + 1]) for i, path in enumerate(original_paths))
    assert len(build_rename_plan(requests).entries) == checked_count + 1
    if checked_count == 1:
        calls = 0

        def fail_second_move(source, destination):
            nonlocal calls
            calls += 1
            if calls == 2:
                raise OSError("injected rename failure")
            source.rename(destination)

        with pytest.raises(RenameTransactionError):
            rename_batch(requests, move=fail_second_move)
        assert {path.name: hashlib.sha256(path.read_bytes()).hexdigest() for path in (manual, *original_paths)} == source_hashes
    _, record = rename_batch(requests)
    assert len(record.entries) == checked_count + 1
    assert (tmp_path / "重要事項説明書.pdf").read_bytes() == b"manual-data"
    assert (tmp_path / "不要.pdf").read_bytes() == b"pdf-0"
    undo_rename(record)
    assert {path.name: hashlib.sha256(path.read_bytes()).hexdigest() for path in (manual, *original_paths)} == source_hashes


def test_drop_rejects_multiple_and_files_without_mutating_folder(tmp_path: Path) -> None:
    app = main.OverviewRenameApp.__new__(main.OverviewRenameApp)
    app.busy = False
    app.root = object()
    app.folder = tmp_path
    app._confirm_discard = Mock(return_value=False)
    app._start_load = Mock()
    one = tmp_path / "日本語 (資料)"
    two = tmp_path / "other"
    one.mkdir()
    two.mkdir()
    pdf = tmp_path / "test.pdf"
    pdf.write_bytes(b"PDF")
    original = main.messagebox.showinfo
    main.messagebox.showinfo = Mock()
    try:
        assert not app._load_dropped_paths([str(one), str(two)])
        assert not app._load_dropped_paths([str(pdf)])
        assert not app._load_dropped_paths([str(one)])  # discard cancelled
        app._start_load.assert_not_called()
        app._confirm_discard.return_value = True
        assert app._load_dropped_paths([str(one)])
        app._start_load.assert_called_once_with(one)
    finally:
        main.messagebox.showinfo = original


def test_native_tkdnd_load_and_tcl_path_parsing() -> None:
    from tkinterdnd2 import TkinterDnD

    root = tk_root(TkinterDnD.Tk)
    try:
        app = main.OverviewRenameApp(root)
        assert app._dnd_enabled
        assert root.tk.splitlist("{C:/日本語 資料 (1)} {C:/other folder}") == (
            "C:/日本語 資料 (1)", "C:/other folder"
        )
    finally:
        app.closing = True
        if app._poll_after is not None:
            root.after_cancel(app._poll_after)
        app.scanner.shutdown()
        app.render_pool.shutdown()
        app.preview_worker.shutdown()
        root.destroy()


def test_checked_noop_stays_pending_without_a_rename(tmp_path: Path) -> None:
    root = tk_root()
    app = main.OverviewRenameApp(root)
    app.folder = tmp_path
    path = tmp_path / "不要.pdf"
    path.write_bytes(b"original")
    try:
        app._create_card(FileSnapshot.capture(path))
        app._layout_cards()
        settle_view(root, app)
        selected = app.cards[0]
        selected.unneeded_var.set(True)
        app._toggle_unneeded(selected)
        assert selected.variable.get() == "不要"
        assert app._pending_ids == {selected.identifier}
        assert "変更待ち1件" in app.status_var.get()
        assert not build_rename_plan([RenameRequest(selected.snapshot, selected.variable.get())]).entries
        selected.unneeded_var.set(False)
        app._toggle_unneeded(selected)
        assert not app._pending_ids
    finally:
        app.closing = True
        if app._poll_after is not None:
            root.after_cancel(app._poll_after)
        app.scanner.shutdown()
        app.render_pool.shutdown()
        app.preview_worker.shutdown()
        root.destroy()


@pytest.mark.parametrize("width", [900, 1920])
def test_800_viewport_all_sizes_positions_and_unneeded_state(tmp_path: Path, width: int) -> None:
    root = tk_root()
    root.geometry(f"{width}x900+0+0")
    app = main.OverviewRenameApp(root)
    app._reprioritize_unrendered = Mock()
    sample = tmp_path / "sample.pdf"
    sample.write_bytes(b"synthetic")
    snapshot = FileSnapshot.capture(sample)
    try:
        root.update()
        for i in range(800):
            app._create_card(snapshot)
            app.cards[-1].original_name = f"scan{i:04d}.pdf"
            app.cards[-1].variable.set(f"scan{i:04d}")
        app._layout_cards()
        app.cards[799].unneeded = True
        app.cards[799].variable.set("不要")
        app.cards[799].manual_value = "最後の手入力"
        for size in main.SIZE_CONFIG:
            app.size_var.set(size)
            started = time.perf_counter()
            app.change_size()
            assert time.perf_counter() - started < .5
            settle_view(root, app)
            for position in (0, .5, 1):
                app._on_scrollbar("moveto", position)
                settle_view(root, app)
                top = max(0, app.canvas.canvasy(0))
                bottom = top + app.canvas.winfo_height()
                visible = []
                for card in app.cards:
                    y = card.identifier // app._current_columns * app._row_height + 7
                    if y < bottom and y + app._row_height - 14 > top:
                        visible.append(card.identifier)
                        assert card.frame is not None and card.frame.winfo_ismapped()
                        assert card.entry.cget("state") == ("disabled" if card.unneeded else "normal")
                        assert card.name_label.cget("text") == card.original_name
                assert len(app._mounted) < 100
                if position == 1:
                    assert 799 in visible
                    assert app.cards[799].entry.cget("state") == "disabled"
                    assert app.cards[799].variable.get() == "不要"
    finally:
        app.closing = True
        if app._poll_after is not None:
            root.after_cancel(app._poll_after)
        app.scanner.shutdown()
        app.render_pool.shutdown()
        app.preview_worker.shutdown()
        root.destroy()


def test_real_tk_unneeded_state_view_reuse_and_tab_preview_focus(tmp_path: Path) -> None:
    root = tk_root()
    root.geometry("900x620+2000+100")
    app = main.OverviewRenameApp(root)
    app.folder = tmp_path
    app._reprioritize_unrendered = Mock()
    try:
        for i in range(32):
            path = tmp_path / f"scan{i:03d}.pdf"
            path.write_bytes(b"synthetic")
            app._create_card(FileSnapshot.capture(path))
        app._layout_cards()
        settle_view(root, app)
        first = app.cards[0]
        first.variable.set("手入力")
        first.unneeded_var.set(True)
        app._toggle_unneeded(first)
        assert first.unneeded and first.manual_value == "手入力"
        assert first.entry.cget("state") == "disabled"
        assert first.variable.get() == "不要"
        app.cards[1].variable.set("不要")
        time.sleep(.07)
        root.update()
        assert first.variable.get() == "不要1"  # a manual target always takes priority
        app.cards[1].variable.set("scan001")
        time.sleep(.07)
        root.update()
        assert first.variable.get() == "不要"
        app.size_var.set("xlarge")
        app.change_size()
        settle_view(root, app)
        app._on_scrollbar("moveto", 1)
        settle_view(root, app)
        app._on_scrollbar("moveto", 0)
        settle_view(root, app)
        assert first.unneeded and first.variable.get() == "不要"
        first.unneeded_var.set(False)
        app._toggle_unneeded(first)
        assert first.variable.get() == "手入力"
        assert first.entry.cget("state") == "normal"
        first.entry.focus_force()
        root.update()
        app._preview_window = tk.Toplevel(root)
        app._preview_snapshot = first.snapshot
        app._set_preview_target = Mock()
        app._focus_adjacent(first, 1)
        root.update()
        assert root.focus_get() is app.cards[1].entry
        app._set_preview_target.assert_called_with(app.cards[1], activate=False)
        app.cards[1].entry.insert("end", "次")
        assert app.cards[1].variable.get().endswith("次")
        app._focus_adjacent(app.cards[1], -1)
        root.update()
        assert root.focus_get() is first.entry
        for i in range(1, 32):
            app._focus_adjacent(app.cards[i - 1], 1)
            root.update()
            assert root.focus_get() is app.cards[i].entry
            assert app.cards[i].frame is not None and app.cards[i].frame.winfo_ismapped()
    finally:
        app.closing = True
        if app._poll_after is not None:
            root.after_cancel(app._poll_after)
        app.scanner.shutdown()
        app.render_pool.shutdown()
        app.preview_worker.shutdown()
        root.destroy()
