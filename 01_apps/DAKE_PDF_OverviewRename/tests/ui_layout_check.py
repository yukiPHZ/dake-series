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
    app._set_success("✓ 300件の名前変更が完了しました")
    root.update_idletasks()
    root.update()
    toolbar = app.select_button.master
    toolbar_right = app.apply_button.winfo_x() + app.apply_button.winfo_width()
    toolbar_order = (
        app.select_button.winfo_x()
        < app.refresh_button.winfo_x()
        < app.reload_button.winfo_x()
    )
    header_right = app.description_label.winfo_x() + app.description_label.winfo_width()
    footer_stacked = app.footer_right.winfo_y() > app.footer_brand.winfo_y()
    title_center = app.title_label.winfo_y() + app.title_label.winfo_height() / 2
    description_center = app.description_label.winfo_y() + app.description_label.winfo_height() / 2
    status_stacked = app.success_label.winfo_y() > app.status_label.winfo_y()
    status_fits = (
        app.status_label.winfo_x() + app.status_label.winfo_width() <= app.status_row.winfo_width()
        and app.success_label.winfo_x() + app.success_label.winfo_width() <= app.status_row.winfo_width()
    )
    status_no_overlap = status_stacked or (
        app.status_label.winfo_x() + app.status_label.winfo_width() <= app.success_label.winfo_x()
    )
    result = {
        "scale_percent": int(scale_factor * 100),
        "window_width": width,
        "toolbar_fits": toolbar_right <= toolbar.winfo_width(),
        "toolbar_order": toolbar_order,
        "reload_visible": app.reload_button.winfo_ismapped(),
        "xlarge_visible": app.size_buttons["xlarge"].winfo_ismapped(),
        "header_horizontal": abs(description_center - title_center) <= 2,
        "header_fits": header_right <= app.header.winfo_width(),
        "status_stacked": status_stacked,
        "status_fits": status_fits,
        "status_no_overlap": status_no_overlap,
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
        and result["toolbar_order"]
        and result["reload_visible"]
        and result["xlarge_visible"]
        and result["header_horizontal"]
        and result["header_fits"]
        and result["status_fits"]
        and result["status_no_overlap"]
        and result["footer_fits"]
        for result in results
    ), results


if __name__ == "__main__":
    main()
