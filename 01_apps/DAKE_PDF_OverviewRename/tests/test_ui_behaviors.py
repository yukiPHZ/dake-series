# -*- coding: utf-8 -*-
from __future__ import annotations

import tkinter as tk
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import main
import pytest
from PIL import Image
from main import OverviewRenameApp, UI_TEXT
from rename_core import FileSnapshot, RenameEntry, RenamePlan, UndoEntry, UndoRecord


class FakeVariable:
    def __init__(self, value: str) -> None:
        self.value = value

    def set(self, value: str) -> None:
        self.value = value

    def get(self) -> str:
        return self.value


def test_root_scoped_wheel_routes_all_main_list_surfaces() -> None:
    root = object()
    canvas = Mock()
    app = OverviewRenameApp.__new__(OverviewRenameApp)
    app.root = SimpleNamespace(after_idle=Mock())
    app.canvas = canvas
    app._reprioritize_unrendered = Mock()

    for surface in ("canvas", "thumbnail", "name_label", "entry"):
        widget = SimpleNamespace(winfo_toplevel=lambda: app.root, surface=surface)
        event = SimpleNamespace(widget=widget, delta=-120)
        assert app._route_mousewheel(event) == "break"

    assert canvas.yview_scroll.call_count == 4
    canvas.yview_scroll.assert_called_with(1, "units")
    assert app.root.after_idle.call_count == 4


def test_root_scoped_wheel_ignores_preview_toplevel() -> None:
    app = OverviewRenameApp.__new__(OverviewRenameApp)
    app.root = SimpleNamespace(after_idle=Mock())
    app.canvas = Mock()
    preview = object()
    event = SimpleNamespace(widget=SimpleNamespace(winfo_toplevel=lambda: preview), delta=-120)

    assert app._route_mousewheel(event) is None
    app.canvas.yview_scroll.assert_not_called()
    app.root.after_idle.assert_not_called()


def test_real_tk_wheel_binding_and_refresh_integration(monkeypatch, tmp_path: Path) -> None:
    root = _create_tk_root()
    root.geometry("900x620+2500+100")
    app = OverviewRenameApp(root)
    first_surfaces: tuple[tk.Widget, tk.Widget, tk.Widget] | None = None
    for index in range(48):
        card = tk.Frame(app.cards_frame)
        card.pack(fill="x", pady=2)
        thumbnail = tk.Label(card, text=f"thumbnail {index}", height=2)
        thumbnail.pack(fill="x")
        name_label = tk.Label(card, text=f"source_{index:04d}.pdf")
        name_label.pack(fill="x")
        entry = tk.Entry(card)
        entry.insert(0, f"source_{index:04d}")
        entry.pack(fill="x")
        if first_surfaces is None:
            first_surfaces = (thumbnail, name_label, entry)
    root.update_idletasks()
    root.update()
    app._update_scrollregion()

    try:
        assert first_surfaces is not None
        for widget in (app.canvas, *first_surfaces):
            app.canvas.yview_moveto(0.0)
            root.update()
            before = app.canvas.yview()[0]
            widget.event_generate("<MouseWheel>", delta=-120, when="tail")
            root.update()
            assert app.canvas.yview()[0] > before

        preview = tk.Toplevel(root)
        preview_entry = tk.Entry(preview)
        preview_entry.pack()
        root.update()
        before = app.canvas.yview()[0]
        preview_entry.event_generate("<MouseWheel>", delta=-120)
        root.update()
        assert app.canvas.yview()[0] == before
        preview.destroy()

        selected_folder = tmp_path.resolve()
        pending_card = SimpleNamespace(pending=True)
        app.folder = selected_folder
        app.cards = [pending_card]
        app.undo_record = object()
        app.path_var.set(str(selected_folder))
        app.status_var.set("loaded")
        app._preview_window = tk.Toplevel(root)
        app._preview_label = tk.Label(app._preview_window)
        app._preview_label.pack()
        answers = iter((False, True))
        monkeypatch.setattr(
            main.messagebox,
            "askyesno",
            lambda *_args, **_kwargs: next(answers),
        )
        root.update()

        app.refresh()
        assert app.folder == selected_folder
        assert app.cards == [pending_card]
        assert app.undo_record is not None
        assert app._preview_window is not None

        app.refresh()
        root.update()
        assert app.folder is None
        assert app.cards == []
        assert app.undo_record is None
        assert app.path_var.get() == UI_TEXT["folder_unselected"]
        assert app.status_var.get() == UI_TEXT["status_empty"]
        assert app._preview_window is None
        assert app.refresh_button.cget("state") == "disabled"
        assert app.select_button.cget("state") == "normal"
    finally:
        app.closing = True
        if app._poll_after is not None:
            root.after_cancel(app._poll_after)
        app.scanner.shutdown()
        app.render_pool.shutdown()
        app.preview_worker.shutdown()
        root.destroy()


def test_refresh_is_safe_noop_when_folder_is_unselected() -> None:
    app = OverviewRenameApp.__new__(OverviewRenameApp)
    app.busy = False
    app.folder = None
    app._confirm_discard = Mock()
    app._reset_to_initial = Mock()

    app.refresh()

    app._confirm_discard.assert_not_called()
    app._reset_to_initial.assert_not_called()


def test_refresh_keeps_state_when_pending_discard_is_rejected(tmp_path: Path) -> None:
    app = OverviewRenameApp.__new__(OverviewRenameApp)
    app.busy = False
    app.folder = tmp_path
    app._confirm_discard = Mock(return_value=False)
    app._reset_to_initial = Mock()

    app.refresh()

    app._confirm_discard.assert_called_once_with()
    app._reset_to_initial.assert_not_called()


def test_refresh_resets_folder_jobs_cards_preview_and_undo(tmp_path: Path) -> None:
    app = OverviewRenameApp.__new__(OverviewRenameApp)
    app.busy = False
    app.folder = tmp_path
    app.scan_token = 4
    app.generation = 7
    app.preview_generation = 10
    app.undo_record = object()
    app.scanner = SimpleNamespace(cancel=Mock())
    app.render_pool = SimpleNamespace(cancel=Mock())
    app._close_preview = Mock(side_effect=lambda: setattr(app, "preview_generation", 11))
    app._clear_cards = Mock()
    app.canvas = SimpleNamespace(yview_moveto=Mock())
    app.path_var = FakeVariable(str(tmp_path))
    app.status_var = FakeVariable("loaded")
    app._sync_controls = Mock()
    app._confirm_discard = Mock(return_value=True)

    app.refresh()

    assert app.scan_token == 5
    assert app.generation == 8
    app.scanner.cancel.assert_called_once_with(5)
    app.render_pool.cancel.assert_called_once_with(8)
    app._close_preview.assert_called_once_with()
    assert app.preview_generation == 11
    assert app.folder is None
    assert app.undo_record is None
    app._clear_cards.assert_called_once_with()
    app.canvas.yview_moveto.assert_called_once_with(0.0)
    assert app.path_var.value == UI_TEXT["folder_unselected"]
    assert app.status_var.value == UI_TEXT["status_empty"]
    app._sync_controls.assert_called_once_with()


def test_reload_control_is_disabled_without_folder_and_enabled_with_folder(tmp_path: Path) -> None:
    app = OverviewRenameApp.__new__(OverviewRenameApp)
    app.cards = []
    app.busy = False
    app.folder = None
    app.undo_record = None
    app.apply_button = Mock()
    app.undo_button = Mock()
    app.refresh_button = Mock()
    app.reload_button = Mock()
    app.select_button = Mock()

    app._sync_controls()
    assert app.reload_button.configure.call_args.kwargs["state"] == "disabled"

    app.folder = tmp_path
    app._sync_controls()
    assert app.reload_button.configure.call_args.kwargs["state"] == "normal"


def test_reload_is_safe_noop_when_folder_is_unselected_or_busy(tmp_path: Path) -> None:
    app = OverviewRenameApp.__new__(OverviewRenameApp)
    app.busy = False
    app.folder = None
    app._confirm_discard = Mock()
    app._close_preview = Mock()
    app._start_load = Mock()
    app.canvas = SimpleNamespace(yview_moveto=Mock())

    app.reload()

    app._confirm_discard.assert_not_called()
    app._start_load.assert_not_called()

    app.folder = tmp_path
    app.busy = True
    app.reload()
    app._confirm_discard.assert_not_called()
    app._start_load.assert_not_called()


def test_reload_rejection_preserves_all_state(tmp_path: Path) -> None:
    app = OverviewRenameApp.__new__(OverviewRenameApp)
    folder = tmp_path.resolve()
    cards = [SimpleNamespace(pending=True)]
    undo_record = object()
    preview_window = object()
    app.busy = False
    app.folder = folder
    app.cards = cards
    app.undo_record = undo_record
    app._preview_window = preview_window
    app.size_var = FakeVariable("large")
    app._confirm_discard = Mock(return_value=False)
    app._close_preview = Mock()
    app._start_load = Mock()
    app.canvas = SimpleNamespace(yview_moveto=Mock())

    app.reload()

    assert app.folder == folder
    assert app.cards is cards
    assert app.undo_record is undo_record
    assert app._preview_window is preview_window
    assert app.size_var.get() == "large"
    app._close_preview.assert_not_called()
    app._start_load.assert_not_called()
    app.canvas.yview_moveto.assert_not_called()


def test_reload_approval_rescans_same_folder_and_resets_old_state(tmp_path: Path) -> None:
    app = OverviewRenameApp.__new__(OverviewRenameApp)
    folder = tmp_path.resolve()
    preview_window = SimpleNamespace(destroy=Mock())
    app.busy = False
    app.folder = folder
    app.cards = [SimpleNamespace(pending=True)]
    app.undo_record = object()
    app.scan_token = 4
    app.generation = 7
    app.preview_generation = 10
    app.rendered_count = 1
    app._preview_window = preview_window
    app._preview_label = object()
    app._preview_photo = object()
    app.size_var = FakeVariable("large")
    app.path_var = FakeVariable(str(folder))
    app.status_var = FakeVariable("loaded")
    app._confirm_discard = Mock(return_value=True)
    app.scanner = SimpleNamespace(request=Mock())
    app.render_pool = SimpleNamespace(cancel=Mock())
    app.preview_worker = SimpleNamespace(cancel=Mock())
    app._clear_cards = Mock(side_effect=lambda: app.cards.clear())
    app._sync_controls = Mock()
    app.canvas = SimpleNamespace(yview_moveto=Mock())

    app.reload()

    assert app.folder == folder
    assert app.size_var.get() == "large"
    assert app.undo_record is None
    assert app._preview_window is None
    assert app.scan_token == 5
    assert app.generation == 8
    assert app.preview_generation == 12
    app.scanner.request.assert_called_once_with(5, folder)
    app.render_pool.cancel.assert_called_once_with(8)
    assert [call.args for call in app.preview_worker.cancel.call_args_list] == [(11,), (12,)]
    app._clear_cards.assert_called_once_with()
    app.canvas.yview_moveto.assert_called_once_with(0.0)


def test_reload_stale_generation_thumbnail_is_ignored() -> None:
    card = SimpleNamespace(
        snapshot=SimpleNamespace(path=Path("same.pdf")),
        rendered=False,
    )
    app = OverviewRenameApp.__new__(OverviewRenameApp)
    app.generation = 8
    app.cards = [card]
    app.rendered_count = 0
    app._sync_status = Mock()
    stale_request = main.RenderRequest(7, "thumbnail", 0, card.snapshot, (20, 20))
    stale_result = main.RenderResult(stale_request, object(), 1, None)

    app._accept_thumbnail(stale_result)

    assert card.rendered is False
    assert app.rendered_count == 0
    app._sync_status.assert_not_called()


def test_real_tk_reload_48_cards_reflects_add_delete_and_external_rename(tmp_path: Path) -> None:
    root = _create_tk_root()
    root.geometry("900x620+2500+100")
    app = OverviewRenameApp(root)
    app.scanner.request = Mock()
    app._reprioritize_unrendered = Mock()
    folder = tmp_path.resolve()
    for index in range(48):
        (folder / f"scan_{index:04d}.pdf").write_bytes(b"%PDF-1.4\n%%EOF\n")

    def complete_scan(expected: int) -> None:
        token, requested_folder = app.scanner.request.call_args.args
        app._accept_scan(token, requested_folder, main.scan_pdf_folder(requested_folder), None)
        deadline = main.time.monotonic() + 3
        while len(app.cards) < expected and main.time.monotonic() < deadline:
            root.update()
        root.update_idletasks()
        assert len(app.cards) == expected

    try:
        assert app.reload_button.cget("state") == "disabled"
        app._start_load(folder)
        complete_scan(48)
        assert app.reload_button.cget("state") == "normal"

        app.size_var.set("large")
        app.undo_record = object()
        app._preview_window = tk.Toplevel(root)
        app._preview_label = tk.Label(app._preview_window)
        app._preview_label.pack()
        (folder / "added.pdf").write_bytes(b"%PDF-1.4\n%%EOF\n")
        app.reload()
        complete_scan(49)
        assert app.folder == folder
        assert app.size_var.get() == "large"
        assert app.undo_record is None
        assert app._preview_window is None
        assert app.cards[0].snapshot.path.name == "added.pdf"

        (folder / "scan_0000.pdf").unlink()
        app.reload()
        complete_scan(48)
        assert all(card.snapshot.path.name != "scan_0000.pdf" for card in app.cards)

        (folder / "scan_0001.pdf").rename(folder / "externally_renamed.pdf")
        app.reload()
        complete_scan(48)
        names = [card.snapshot.path.name for card in app.cards]
        assert "externally_renamed.pdf" in names
        assert "scan_0001.pdf" not in names
    finally:
        app.closing = True
        if app._poll_after is not None:
            root.after_cancel(app._poll_after)
        app.scanner.shutdown()
        app.render_pool.shutdown()
        app.preview_worker.shutdown()
        root.destroy()


def _layout_test_app() -> OverviewRenameApp:
    app = OverviewRenameApp.__new__(OverviewRenameApp)
    app._layout_after = None
    app._current_columns = 0
    app._laid_out_count = 0
    app.canvas = SimpleNamespace(winfo_width=lambda: 900)
    app.size_var = FakeVariable("normal")
    app.cards = []
    app.cards_frame = SimpleNamespace(grid_columnconfigure=Mock())
    app._update_scrollregion = Mock()
    return app


def test_empty_layout_then_32_cards_grids_every_frame() -> None:
    app = _layout_test_app()
    app._layout_cards()
    assert app._current_columns > 0

    app.cards = [SimpleNamespace(frame=Mock()) for _ in range(32)]
    app._layout_cards()

    assert app._laid_out_count == 32
    assert all(card.frame.grid.call_count == 1 for card in app.cards)


def test_layout_adds_second_batch_when_columns_are_unchanged() -> None:
    app = _layout_test_app()
    app.cards = [SimpleNamespace(frame=Mock()) for _ in range(main.CARD_BATCH_SIZE)]
    app._layout_cards()
    first_batch = list(app.cards)
    second_batch = [SimpleNamespace(frame=Mock()) for _ in range(8)]
    app.cards.extend(second_batch)

    app._layout_cards()

    assert app._laid_out_count == 32
    assert all(card.frame.grid.call_count == 1 for card in first_batch)
    assert all(card.frame.grid.call_count == 1 for card in second_batch)


def test_xlarge_uses_sufficient_base_resolution() -> None:
    _, image_width, image_height = main.SIZE_CONFIG["xlarge"]
    assert main.UI_TEXT["size_xlarge"] == "特大"
    assert main.THUMB_RENDER_BOX[0] >= image_width
    assert main.THUMB_RENDER_BOX[1] >= image_height


def _success_test_app() -> OverviewRenameApp:
    app = OverviewRenameApp.__new__(OverviewRenameApp)
    app.busy = True
    app.success_var = FakeVariable("")
    app.status_var = FakeVariable("working")
    app._status_stacked = None
    app.root = SimpleNamespace(after_idle=Mock())
    app.status_row = object()
    app._responsive_status = Mock()
    app._reschedule_unrendered = Mock()
    app._sync_status = Mock()
    app._style_card = Mock()
    return app


def test_rename_success_uses_separate_non_modal_feedback(tmp_path: Path) -> None:
    original_path = tmp_path / "before.pdf"
    renamed_path = tmp_path / "after.pdf"
    original_path.write_bytes(b"before")
    renamed_path.write_bytes(b"after")
    original_snapshot = FileSnapshot.capture(original_path)
    renamed_snapshot = FileSnapshot.capture(renamed_path)
    plan = RenamePlan(tmp_path, (RenameEntry(original_snapshot, renamed_path),))
    record = UndoRecord(tmp_path, (UndoEntry(original_path, renamed_snapshot),))
    card = SimpleNamespace(
        snapshot=original_snapshot,
        original_name=original_path.name,
        variable=FakeVariable(original_path.stem),
        name_label=Mock(),
    )
    app = _success_test_app()
    app.cards = [card]
    app.undo_record = None

    app._accept_operation("rename", record, plan)

    assert app.status_var.get() == "working"
    assert app.success_var.get() == UI_TEXT["success_rename"].format(count=1)
    assert app.undo_record is record


def test_undo_success_uses_same_feedback_channel(tmp_path: Path) -> None:
    renamed_path = tmp_path / "after.pdf"
    original_path = tmp_path / "before.pdf"
    renamed_path.write_bytes(b"after")
    renamed_snapshot = FileSnapshot.capture(renamed_path)
    original_path.write_bytes(b"before")
    plan = RenamePlan(tmp_path, (RenameEntry(renamed_snapshot, original_path),))
    record = UndoRecord(tmp_path, (UndoEntry(original_path, renamed_snapshot),))
    card = SimpleNamespace(
        snapshot=renamed_snapshot,
        original_name=renamed_path.name,
        variable=FakeVariable(renamed_path.stem),
        name_label=Mock(),
    )
    app = _success_test_app()
    app.cards = [card]
    app.undo_record = record

    app._accept_operation("undo", plan, record)

    assert app.success_var.get() == UI_TEXT["success_undo"].format(count=1)
    assert app.undo_record is None


def test_thumbnail_progress_does_not_clear_success_feedback() -> None:
    app = OverviewRenameApp.__new__(OverviewRenameApp)
    app.cards = [SimpleNamespace(pending=False, entry=Mock())]
    app.rendered_count = 1
    app.busy = False
    app.folder = Path("folder")
    app.undo_record = None
    app.status_var = FakeVariable("")
    app.success_var = FakeVariable("keep success")
    app.apply_button = Mock()
    app.undo_button = Mock()
    app.refresh_button = Mock()
    app.reload_button = Mock()
    app.select_button = Mock()

    app._sync_status()

    assert app.success_var.get() == "keep success"


def test_next_name_edit_clears_success_feedback() -> None:
    app = OverviewRenameApp.__new__(OverviewRenameApp)
    app.success_var = FakeVariable("completed")
    app._status_stacked = False
    app.root = SimpleNamespace(after_idle=Mock())
    app.status_row = object()
    app._responsive_status = Mock()
    app._style_card = Mock()
    app._sync_status = Mock()
    card = SimpleNamespace(variable=FakeVariable("next name"), hint_label=Mock())

    app._on_name_changed(card)

    assert app.success_var.get() == ""


def _make_ui_pdf(path: Path, index: int, total: int) -> None:
    image = Image.new("RGB", (240, 340), "white")
    image.save(path, "PDF", resolution=72.0)


def _create_tk_root() -> tk.Tk:
    error: tk.TclError | None = None
    for _ in range(5):
        try:
            return tk.Tk()
        except tk.TclError as exc:
            error = exc
            main.time.sleep(0.2)
    pytest.skip(f"Tk display is unavailable: {error}")


def test_real_tk_first_load_progressive_layout_reload_and_xlarge(tmp_path: Path) -> None:
    root = _create_tk_root()
    root.geometry("900x620+2500+100")
    app = OverviewRenameApp(root)
    for index in range(48):
        _make_ui_pdf(tmp_path / f"first_{index:04d}.pdf", index, 48)
    progressive_states: list[bool] = []
    accept_thumbnail = app._accept_thumbnail

    def track_thumbnail(result) -> None:
        accept_thumbnail(result)
        if 0 < app.rendered_count < 48:
            progressive_states.append(
                any(card.photo is not None for card in app.cards)
                and all(card.frame.winfo_manager() == "grid" for card in app.cards)
            )

    app._accept_thumbnail = track_thumbnail

    def wait_for_complete() -> bool:
        deadline = main.time.monotonic() + 30
        while app.rendered_count < 48 and main.time.monotonic() < deadline:
            root.update()
        root.update_idletasks()
        assert app.rendered_count == 48
        assert len(app.cards) == 48
        assert all(card.frame.winfo_manager() == "grid" for card in app.cards)
        return any(progressive_states)

    try:
        root.update()
        app._layout_cards()
        assert app.cards == []
        assert app._current_columns > 0

        app._start_load(tmp_path)
        assert wait_for_complete()

        progressive_states.clear()
        app.reload()
        assert wait_for_complete()

        first_card = app.cards[0]
        first_card.variable.set("pending_name")
        undo_marker = object()
        app.undo_record = undo_marker
        app.render_pool.replace = Mock()
        app.size_var.set("xlarge")
        app.change_size()
        root.update()
        assert first_card.variable.get() == "pending_name"
        assert first_card.pending
        assert app.undo_record is undo_marker
        assert all(card.frame.winfo_manager() == "grid" for card in app.cards)
        assert all(call.args[1] == [] for call in app.render_pool.replace.call_args_list)
        assert app.size_buttons["xlarge"].cget("text") == UI_TEXT["size_xlarge"]

        app.size_var.set("normal")
        app.change_size()
        root.update()
        assert first_card.variable.get() == "pending_name"
        assert first_card.pending
        assert app.undo_record is undo_marker
        assert all(card.frame.winfo_manager() == "grid" for card in app.cards)
    finally:
        app.closing = True
        if app._poll_after is not None:
            root.after_cancel(app._poll_after)
        if app._layout_after is not None:
            root.after_cancel(app._layout_after)
        app._close_preview()
        app.scanner.shutdown()
        app.render_pool.shutdown()
        app.preview_worker.shutdown()
        root.destroy()


def test_preview_fit_zoom_clamp_pan_quality_and_reset(monkeypatch, tmp_path: Path) -> None:
    root = _create_tk_root()
    root.geometry("900x620+2500+100")
    app = OverviewRenameApp(root)
    request = Mock()
    cancel = Mock()
    monkeypatch.setattr(app.preview_worker, "request", request)
    monkeypatch.setattr(app.preview_worker, "cancel", cancel)
    first_path = tmp_path / "first.pdf"
    second_path = tmp_path / "second.pdf"
    first_path.write_bytes(b"first")
    second_path.write_bytes(b"second")
    first = SimpleNamespace(identifier=0, snapshot=FileSnapshot.capture(first_path))
    second = SimpleNamespace(identifier=1, snapshot=FileSnapshot.capture(second_path))

    try:
        app.show_preview(first)
        root.update()
        initial_request = request.call_args.args[0]
        app._accept_preview(
            main.RenderResult(initial_request, Image.new("RGB", main.PREVIEW_RENDER_BOX, "white"), 1, None)
        )
        root.update()
        assert app._preview_zoom == 1.0
        assert app.preview_zoom_var.get() == "100%"
        assert app._preview_canvas is not None
        assert app._preview_image_item is not None

        main_scroll = app.canvas.yview()
        center_x = app._preview_canvas.winfo_width() // 2
        center_y = app._preview_canvas.winfo_height() // 2
        event = SimpleNamespace(delta=120, x=center_x, y=center_y)
        app._preview_canvas.event_generate("<MouseWheel>", delta=120, x=center_x, y=center_y)
        root.update()
        assert app._preview_zoom > 1.0
        assert app.preview_zoom_var.get() != "100%"
        assert app.canvas.yview() == main_scroll
        event.delta = -120
        app._on_preview_wheel(event)
        assert app._preview_zoom == pytest.approx(1.0)

        event.delta = 120
        for _ in range(20):
            app._on_preview_wheel(event)
        assert app._preview_zoom == main.PREVIEW_ZOOM_MAX
        assert app._preview_display_size[0] <= app._preview_base_image.size[0]
        assert app._preview_display_size[1] <= app._preview_base_image.size[1]
        before_pan = (app._preview_canvas.xview(), app._preview_canvas.yview())
        app._on_preview_pan_start(SimpleNamespace(x=center_x, y=center_y))
        app._on_preview_pan_move(SimpleNamespace(x=center_x - 180, y=center_y - 140))
        after_pan = (app._preview_canvas.xview(), app._preview_canvas.yview())
        assert after_pan != before_pan

        event.delta = -120
        for _ in range(40):
            app._on_preview_wheel(event)
        assert app._preview_zoom == main.PREVIEW_ZOOM_MIN

        accepted_base = app._preview_base_image
        stale_request = main.RenderRequest(
            app.preview_generation - 1, "preview", 0, first.snapshot, main.PREVIEW_RENDER_BOX
        )
        app._accept_preview(main.RenderResult(stale_request, Image.new("RGB", (20, 20), "red"), 1, None))
        assert app._preview_base_image is accepted_base

        app.show_preview(second)
        root.update()
        assert app._preview_zoom == 1.0
        assert app.preview_zoom_var.get() == "100%"
        assert app._preview_base_image is None
        assert app._preview_canvas.xview()[0] == pytest.approx(0.0)
        assert app._preview_canvas.yview()[0] == pytest.approx(0.0)

        second_request = request.call_args.args[0]
        app._accept_preview(
            main.RenderResult(second_request, Image.new("RGB", (100, 100), "white"), 1, None)
        )
        root.update()
        app._preview_requested_box = (100, 100)
        request.reset_mock()
        scheduled: list[str] = []
        for _ in range(5):
            app._schedule_preview_quality()
            assert app._preview_zoom_after is not None
            scheduled.append(app._preview_zoom_after)
        assert len(set(scheduled)) == 5
        assert request.call_count == 0
        root.after_cancel(app._preview_zoom_after)
        app._preview_zoom_after = None
        app._request_preview_quality()
        assert request.call_count == 1
        quality_request = request.call_args.args[0]
        assert quality_request.box[0] > 100 or quality_request.box[1] > 100
        assert quality_request.box[0] <= main.PREVIEW_MAX_RENDER_BOX[0]
        assert quality_request.box[1] <= main.PREVIEW_MAX_RENDER_BOX[1]

        app._schedule_preview_quality()
        assert app._preview_zoom_after is not None
        app._close_preview()
        assert app._preview_zoom_after is None
        assert app._preview_resize_after is None
        assert app._preview_canvas is None
        assert app._preview_snapshot is None
    finally:
        app.closing = True
        if app._poll_after is not None:
            root.after_cancel(app._poll_after)
        if app._layout_after is not None:
            root.after_cancel(app._layout_after)
        if app._preview_window is not None:
            app._close_preview()
        app.scanner.shutdown()
        app.render_pool.shutdown()
        app.preview_worker.shutdown()
        root.destroy()
