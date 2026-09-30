"""Measure actual synthetic PDFs in source Tk; never substitutes for native review."""
import argparse
import json
import sys
import time
import tkinter as tk
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import main
from synthetic_pdf_trial import working_set_bytes
from viewport_helpers import settle_view, assert_visible_view


def run():
    parser = argparse.ArgumentParser()
    parser.add_argument("folder", type=Path)
    parser.add_argument("out", type=Path)
    args = parser.parse_args()
    expected = json.loads((args.folder / "expected.json").read_text(encoding="utf-8"))
    root = tk.Tk()
    app = main.OverviewRenameApp(root)
    root.geometry("1920x1080+0+0")
    root.update()
    started = last = time.perf_counter()
    gaps = []
    def tick():
        nonlocal last
        now = time.perf_counter()
        gaps.append(now-last)
        last = now
        root.after(10, tick)
    tick()
    first = None
    memory = []
    app._start_load(args.folder)
    while app.rendered_count < len(expected):
        root.update()
        now = time.perf_counter()
        if first is None and any(c.photo for c in app._mounted.values()):
            first = {"seconds":now-started,"processed":app.rendered_count}
        memory.append(working_set_bytes()/1024**2)
        assert now-started < 120
        time.sleep(.001)
    settle_view(root, app)
    result = {"first_visible":first,"complete_seconds":time.perf_counter()-started,
              "load_max_gap_ms":max(gaps)*1000,"load_peak_mb":max(memory),"switches":[]}
    assert {c.original_name for c in app.cards} == {e["name"] for e in expected}
    assert all(c.page_count==e["pages"] and c.snapshot.size==e["bytes"] for c,e in zip(app.cards,expected))
    for size in ("small","xlarge","normal","large","small")*3:
        app._on_scrollbar("moveto",1)
        settle_view(root,app)
        app.size_var.set(size)
        start = time.perf_counter()
        gaps.clear()
        last = start
        app.change_size()
        settle_view(root,app)
        assert_visible_view(app, True)
        result["switches"].append({"size":size,"settle_ms":(time.perf_counter()-start)*1000,
                                   "max_event_gap_ms":max(gaps or [0])*1000,
                                   "memory_mb":working_set_bytes()/1024**2,"views":len(app._mounted)})
    app.on_close()
    args.out.write_text(json.dumps(result,indent=2),encoding="utf-8")
    print(json.dumps(result,indent=2))


if __name__ == "__main__":
    run()
