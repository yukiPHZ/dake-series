# -*- coding: utf-8 -*-
"""Programmatic Windows Tk layout check at 100%, 125%, and 150% scaling."""

from __future__ import annotations

import json
import sys
import tkinter as tk
from pathlib import Path

APP_DIR = Path(__file__).resolve().parents[1]
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

from main import OverviewRenameApp


def check(scale_factor: float, width: int) -> dict[str, object]:
    root = tk.Tk()
    root.tk.call("tk", "scaling", (96.0 / 72.0) * scale_factor)
    root.geometry(f"{width}x780+2500+100")
    app = OverviewRenameApp(root)
    root.update_idletasks()
    root.update()
    toolbar = app.select_button.master
    toolbar_right = app.apply_button.winfo_x() + app.apply_button.winfo_width()
    header_right = app.description_label.winfo_x() + app.description_label.winfo_width()
    footer_stacked = app.footer_right.winfo_y() > app.footer_brand.winfo_y()
    title_center = app.title_label.winfo_y() + app.title_label.winfo_height() / 2
    description_center = app.description_label.winfo_y() + app.description_label.winfo_height() / 2
    result = {
        "scale_percent": int(scale_factor * 100),
        "window_width": width,
        "toolbar_fits": toolbar_right <= toolbar.winfo_width(),
        "header_horizontal": abs(description_center - title_center) <= 2,
        "header_fits": header_right <= app.header.winfo_width(),
        "footer_stacked": footer_stacked,
        "footer_fits": (
            app.footer_right.winfo_x() + app.footer_right.winfo_width() <= app.footer.winfo_width()
            and app.footer_brand.winfo_x() + app.footer_brand.winfo_width() <= app.footer.winfo_width()
        ),
    }
    app.closing = True
    if app._poll_after is not None:
        root.after_cancel(app._poll_after)
    app.scanner.shutdown()
    app.render_pool.shutdown()
    app.preview_worker.shutdown()
    root.destroy()
    return result


def main() -> None:
    results = [check(scale, width) for scale in (1.0, 1.25, 1.5) for width in (1180, 900)]
    print(json.dumps(results, ensure_ascii=False, indent=2))
    assert all(
        result["toolbar_fits"]
        and result["header_horizontal"]
        and result["header_fits"]
        and result["footer_fits"]
        for result in results
    ), results


if __name__ == "__main__":
    main()
