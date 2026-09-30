from pathlib import Path
import time
import tkinter as tk
from types import SimpleNamespace
from unittest.mock import Mock

import pytest
from PIL import Image
import main
from rename_core import FileSnapshot
from viewport_helpers import settle_view, assert_visible_view
from human_review_fixture import generate


def tk_root():
    for attempt in range(5):
        try:
            return tk.Tk()
        except tk.TclError:
            if attempt == 4:
                raise
            time.sleep(.2)


@pytest.mark.parametrize("scale", [1, 1.25, 1.5])
@pytest.mark.parametrize("width", [900, 1180, 1920])
def test_400_matrix_identity_geometry_latest_size_and_editing(scale, width, tmp_path):
    root = tk_root()
    root.tk.call("tk", "scaling", 96/72*scale)
    app = main.OverviewRenameApp(root)
    root.geometry(f"{width}x1080+0+0")
    app._reprioritize_unrendered = Mock()
    path = tmp_path / "sample.pdf"
    path.write_bytes(b"synthetic")
    snapshot = FileSnapshot.capture(path)
    try:
        root.update()
        app._layout_cards()  # empty configure before first batch
        for i in range(400):
            app._create_card(snapshot)
            app.cards[-1].original_name = f"sample-{i:04d}.pdf"
            app.cards[-1].variable.set(f"sample-{i:04d}")
        app._layout_cards()
        for loaded in (False, True):
            if loaded:
                for card in app.cards:
                    rgb = (card.identifier % 251, card.identifier // 251, 127)
                    card.expected_rgb = rgb
                    request = main.RenderRequest(app.generation, "thumbnail", card.identifier, snapshot, main.THUMB_RENDER_BOX)
                    app._accept_thumbnail(main.RenderResult(request, Image.new("RGB", main.THUMB_RENDER_BOX, rgb), card.identifier+1, None))
            for size in main.SIZE_CONFIG:
                app.size_var.set(size)
                app.change_size()
                settle_view(root, app)
                for position in (0, .5, 1):
                    app._on_scrollbar("moveto", position)
                    settle_view(root, app)
                    ids = assert_visible_view(app, loaded)
                    if position == 1:
                        assert 399 in ids
                    for card in app._mounted.values():
                        assert card.entry.master.winfo_y()+card.entry.master.winfo_height() <= card.body.winfo_height()
            # User's small-tail -> xlarge, preserving a visible document.
            app.size_var.set("small")
            app.change_size()
            settle_view(root, app)
            app._on_scrollbar("moveto", 1)
            settle_view(root, app)
            anchor, offset = app._view_anchor()
            app.size_var.set("xlarge")
            started = time.perf_counter()
            app.change_size()
            assert time.perf_counter()-started < .5
            settle_view(root, app)
            assert anchor in assert_visible_view(app, loaded)
            # Intermediate requests are replaced, not queued per file/size.
            for _ in range(3):
                for size in ("small", "xlarge", "normal", "large", "small"):
                    app.size_var.set(size)
                    app.change_size()
            settle_view(root, app)
            assert_visible_view(app, loaded)
        last = app.cards[-1]
        app._on_scrollbar("moveto", 1)
        settle_view(root, app)
        last.entry.focus_force()
        root.update()
        last.variable.set("日本語 編集中")
        last.entry.icursor(3)
        last.entry.selection_range(1, 3)
        entry = last.entry
        app._on_scrollbar("moveto", 0)
        settle_view(root, app)
        assert last.entry is entry  # focused/IME entry is pinned, never rebound
        assert last.variable.get() == "日本語 編集中"
        assert 399 in app._pending_ids
        assert last.entry.index("insert") == 3
        assert last.entry.selection_present()
        app.size_var.set("xlarge")
        app.change_size()
        settle_view(root, app)
        assert last.entry is entry
        assert last.entry.index("insert") == 3
        app.canvas.focus_force()
        root.update()
        settle_view(root, app)
        assert last.frame is None
        app._on_scrollbar("moveto", 1)
        settle_view(root, app)
        assert last.variable.get() == "日本語 編集中"
        assert last.entry.index("insert") == 3
        assert last.entry.selection_present()
    finally:
        app._confirm_discard = lambda: True
        app.on_close()


def test_thumbnail_progress_never_updates_all_entries(tmp_path):
    root = tk_root()
    app = main.OverviewRenameApp(root)
    path = tmp_path / "sample.pdf"
    path.write_bytes(b"synthetic")
    try:
        for _ in range(400):
            app._create_card(FileSnapshot.capture(path))
        app._sync_controls = Mock()
        for card in app.cards:
            request = main.RenderRequest(app.generation, "thumbnail", card.identifier, card.snapshot, main.THUMB_RENDER_BOX)
            app._accept_thumbnail(main.RenderResult(request, None, None, "test failure"))
        app._sync_controls.assert_not_called()
        assert app.rendered_count == app.thumbnail_failures == 400
    finally:
        app.on_close()


def test_real_400_pdfs_progressive_switch_folder_tail_rename_undo_reload(monkeypatch, tmp_path):
    folder = tmp_path / "400"
    expected = generate(folder, 400)
    other = tmp_path / "other"
    generate(other, 1)
    root = tk_root()
    app = main.OverviewRenameApp(root)
    root.geometry("1920x1080+0+0")
    monkeypatch.setattr(main.messagebox, "askyesno", lambda *a, **kw: True)
    errors = []
    monkeypatch.setattr(main.messagebox, "showerror", lambda *a, **kw: errors.append(a))
    def wait(predicate, timeout=45):
        deadline = time.monotonic()+timeout
        while not predicate() and time.monotonic()<deadline:
            root.update()
            time.sleep(.001)
        assert predicate()
        assert not errors
    try:
        root.update()
        app._start_load(folder)
        wait(lambda: 0 < app.rendered_count < 400 and any(c.photo is not None for c in app._mounted.values()))
        assert_visible_view(app)
        app.size_var.set("xlarge")
        app.change_size()
        app._on_scrollbar("moveto", .5)
        # Old generation scan/render results must not overwrite this selection.
        app._start_load(other)
        wait(lambda: len(app.cards)==1 and app.rendered_count==1)
        assert app.cards[0].snapshot.path.parent == other
        app._start_load(folder)
        wait(lambda: app.rendered_count==400)
        settle_view(root, app)
        assert {c.original_name for c in app.cards} == {m["name"] for m in expected}
        assert all(c.page_count==m["pages"] and c.snapshot.size==m["bytes"] for c,m in zip(app.cards,expected))
        app.render_pool.replace = Mock(wraps=app.render_pool.replace)
        app.size_var.set("small")
        app.change_size()
        settle_view(root, app)
        for position in (0, .5, 1):
            target = int(app._logical_height*position)
            direction = 120 if target < app.canvas.canvasy(0) else -120
            for _ in range(app._logical_height//30+2):
                if (direction < 0 and app.canvas.canvasy(0) >= min(target, app._logical_height-app.canvas.winfo_height())) or (direction > 0 and app.canvas.canvasy(0) <= target):
                    break
                app._on_mousewheel(SimpleNamespace(delta=direction))
            settle_view(root, app)
            ids = assert_visible_view(app, True)
            if position==1:
                assert 399 in ids
        assert all(call.args[1] == [] for call in app.render_pool.replace.call_args_list)
        last = app.cards[-1]
        last.variable.set("renamed_tail")
        app.apply()
        wait(lambda: not app.busy)
        assert (folder/"renamed_tail.pdf").exists()
        assert app.success_var.get() == main.UI_TEXT["success_rename"].format(count=1)
        assert not app._pending_ids
        app.undo()
        wait(lambda: not app.busy)
        assert (folder/expected[-1]["name"]).exists()
        assert last.page_count == expected[-1]["pages"]
        assert app.success_var.get() == main.UI_TEXT["success_undo"].format(count=1)
        app.reload()
        wait(lambda: app.rendered_count==400)
        app._on_scrollbar("moveto", 1)
        settle_view(root, app)
        assert 399 in assert_visible_view(app, True)
        assert {c.original_name for c in app.cards} == {m["name"] for m in expected}
    finally:
        app._confirm_discard = lambda: True
        app.on_close()


def test_lossless_thumbnail_storage_and_resize_does_not_reopen_pdf():
    image = Image.new("RGB", (350,455), (17,53,129))
    source = main.ThumbnailSource.encode(image)
    assert len(source.pixels) < 350*455*3
    assert source.copy().tobytes() == image.tobytes()


def test_reused_view_binds_once_without_retaining_previous_card(tmp_path):
    root = tk_root()
    app = main.OverviewRenameApp(root)
    path = tmp_path / "sample.pdf"
    path.write_bytes(b"synthetic")
    try:
        for _ in range(2):
            app._create_card(FileSnapshot.capture(path))
        app._mount_card(app.cards[0])
        entry = app.cards[0].entry
        image = app.cards[0].image_label
        commands = (len(entry._tclCommands),len(image._tclCommands))
        for i in range(100):
            old, new = app.cards[i%2], app.cards[(i+1)%2]
            app._unmount_card(old)
            assert entry.current_card is image.current_card is None
            app._mount_card(new)
            assert entry.current_card is image.current_card is new
            assert commands == (len(entry._tclCommands),len(image._tclCommands))
    finally:
        app.on_close()
