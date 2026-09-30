"""Source Tk checks: visibility/intersections and image identity, not OS screenshots."""
import time


def settle_view(root, app, timeout=5):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        root.update()
        if app._view_after is None and app._image_after is None and app._layout_after is None:
            return
        time.sleep(.002)
    raise AssertionError("viewport did not settle")


def assert_visible_view(app, require_images=False):
    top = max(0, app.canvas.canvasy(0))
    bottom = top + app.canvas.winfo_height()
    columns = max(1, app._current_columns)
    origin = app.canvas.coords(app.cards_window)[1]
    visible = []
    for card in app.cards:
        y = card.identifier // columns * app._row_height + 7
        if y < bottom and y + app._row_height - 14 > top:
            visible.append(card.identifier)
            assert card.frame is not None, card.identifier
            assert card.frame.winfo_ismapped(), card.identifier
            assert abs(origin + card.frame.winfo_y() - y) <= 1
            assert card.name_label.cget("text") == card.original_name
            assert card.entry.cget("textvariable") == str(card.variable)
            assert card.entry.cget("state") == ("disabled" if app.busy else "normal")
            assert card.page_label.cget("text") == app._metadata(card)
            if require_images and not card.thumbnail_failed:
                assert card.photo is not None
                assert card.photo_size == app.size_var.get()
                assert str(card.image_label.cget("image")) == str(card.photo)
                if hasattr(card, "expected_rgb"):
                    pixel = app.root.tk.call(str(card.photo), "get", 0, 0)
                    assert tuple(map(int, pixel)) == card.expected_rgb
    assert visible or not app.cards
    assert len(app._mounted) <= (int(app.canvas.winfo_height()/app._row_height)+4)*columns+1
    assert app.cards_frame.winfo_height() < 10000
    return visible
