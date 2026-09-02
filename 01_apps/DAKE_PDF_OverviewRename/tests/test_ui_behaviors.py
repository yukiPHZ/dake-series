# -*- coding: utf-8 -*-
from __future__ import annotations

import tkinter as tk
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import main
import pytest
from main import OverviewRenameApp, UI_TEXT


class FakeVariable:
    def __init__(self, value: str) -> None:
        self.value = value

    def set(self, value: str) -> None:
        self.value = value


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
    try:
        root = tk.Tk()
    except tk.TclError as exc:
        pytest.skip(f"Tk display is unavailable: {exc}")
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
