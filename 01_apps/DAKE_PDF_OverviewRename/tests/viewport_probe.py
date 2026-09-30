"""Source-Tk timing/geometry probe. Not a packaged-exe visual acceptance test."""
import json
import argparse
import ctypes
import subprocess
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import tkinter as tk
from PIL import Image
import main
from rename_core import FileSnapshot
from synthetic_pdf_trial import working_set_bytes
from viewport_helpers import settle_view, assert_visible_view


def run():
    parser = argparse.ArgumentParser()
    parser.add_argument("--baseline", action="store_true")
    parser.add_argument("--out", type=Path)
    args = parser.parse_args()
    if args.baseline:
        source = subprocess.check_output(["git", "show", "31a31f3513d12e6c0bb7b94be01e6a4323f61c73:01_apps/DAKE_PDF_OverviewRename/main.py"])
        exec(compile(source, "baseline-main.py", "exec"), main.__dict__)
    root = tk.Tk()
    app = main.OverviewRenameApp(root)
    root.geometry("1920x1080+0+0")
    root.update()
    app._reprioritize_unrendered = lambda: None
    snapshot = FileSnapshot.capture(Path(main.__file__))
    memory_initial = working_set_bytes()
    gaps = []
    last_tick = time.perf_counter()
    def tick():
        nonlocal last_tick
        now = time.perf_counter()
        gaps.append(now-last_tick)
        last_tick = now
        root.after(10, tick)
    tick()
    for _ in range(400):
        app._create_card(snapshot)
    image = Image.new("RGB", main.THUMB_RENDER_BOX, "white")
    app._layout_cards()
    root.update()
    applies = []
    tk_image_times = []
    apply_image = app._apply_card_image
    def timed_apply(card):
        start = time.perf_counter()
        apply_image(card)
        applies.append(time.perf_counter()-start)
    app._apply_card_image = timed_apply
    _, (_, ImageTk) = main.load_preview_dependencies()
    photo_constructor = ImageTk.PhotoImage
    def timed_photo(*a, **kw):
        start = time.perf_counter()
        result = photo_constructor(*a, **kw)
        tk_image_times.append(time.perf_counter()-start)
        return result
    ImageTk.PhotoImage = timed_photo
    started = time.perf_counter()
    for card in app.cards:
        request = main.RenderRequest(app.generation, "thumbnail", card.identifier, snapshot, main.THUMB_RENDER_BOX)
        app._accept_thumbnail(main.RenderResult(request, image.copy(), 1, None))
    root.update()
    rows = [{"accept_400_seconds": time.perf_counter() - started, "memory_initial_mb":memory_initial/1024**2, "memory_loaded_mb":working_set_bytes()/1024**2}]
    for size in ("small", "xlarge", "normal", "large", "small")*3:
        app.canvas.yview_moveto(1)
        root.update()
        if not args.baseline:
            settle_view(root, app)
        app.size_var.set(size)
        applies.clear()
        tk_image_times.clear()
        gaps.clear()
        last_tick = time.perf_counter()
        started = time.perf_counter()
        app.change_size()
        handler = time.perf_counter() - started
        root.update()
        if not args.baseline:
            settle_view(root, app)
            assert_visible_view(app, True)
        rows.append(dict(size=size, handler_ms=round(handler*1000, 2), settle_ms=round((time.perf_counter()-started)*1000, 2), max_event_gap_ms=round(max(gaps or [0])*1000,2), image_updates=len(applies), image_total_ms=round(sum(applies)*1000,2), tk_photo_ms=round(sum(tk_image_times)*1000,2), memory_mb=round(working_set_bytes()/1024**2,2), columns=app._current_columns, frame_requested_height=app.cards_frame.winfo_reqheight(), frame_actual_height=app.cards_frame.winfo_height(), scrollregion=app.canvas.cget("scrollregion"), embedded_y=app.canvas.coords(app.cards_window), views=sum(c.frame is not None for c in app.cards), photos=sum(c.photo is not None for c in app.cards)))
    ImageTk.PhotoImage = photo_constructor
    app.on_close()
    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(json.dumps(rows, indent=2), encoding="utf-8")
    print(json.dumps(rows, indent=2))


if __name__ == "__main__":
    run()
