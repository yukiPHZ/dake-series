"""Read-only Tk view. PDF handles and rasterization stay in the existing worker."""

import math
import tkinter as tk
from tkinter import ttk


PREVIEW_PIXEL_LIMIT = 12_000_000


class PagePreview:
    def __init__(self, app, text, colors, icon_path):
        self.app, self.text, self.colors = app, text, colors
        self.window = tk.Toplevel(app)
        self.window.configure(bg=colors["bg"])
        # Screen-capped initial geometry leaves room for taskbar and title bar.
        width = min(900, max(320, app.winfo_screenwidth() - 80))
        height = min(700, max(240, app.winfo_screenheight() - 100))
        self.window.geometry(f"{width}x{height}")
        # Keep every supported 300% raster within budget, without silently
        # reducing the displayed zoom when a very large monitor is used.
        self.window.maxsize(min(1400, max(320, app.winfo_screenwidth() - 80)),
                            min(1000, max(240, app.winfo_screenheight() - 100)))
        if icon_path.exists():
            self.window.iconbitmap(str(icon_path))
        bar = ttk.Frame(self.window, style="Surface.TFrame", padding=(12, 8))
        bar.pack(fill="x")
        self.zoom_var = tk.StringVar(self.window)
        ttk.Label(bar, textvariable=self.zoom_var, style="PanelHint.TLabel").pack(side="left")
        ttk.Label(bar, text=text["preview_hint"], style="PanelHint.TLabel").pack(side="right", padx=12)
        self.canvas = tk.Canvas(self.window, bg=colors["bg"], highlightthickness=0, cursor="fleur")
        self.canvas.pack(fill="both", expand=True, padx=12, pady=12)
        self.canvas.bind("<MouseWheel>", self.on_wheel)
        self.canvas.bind("<ButtonPress-1>", self.pan_start)
        self.canvas.bind("<B1-Motion>", self.pan_move)
        self.canvas.bind("<Configure>", self.on_resize)
        self.window.bind("<Escape>", lambda _event: app._close_preview())
        self.window.bind("<F5>", lambda _event: app.refresh_all())
        self.window.bind("<Control-o>", lambda _event: app._choose_pdf())
        self.window.protocol("WM_DELETE_WINDOW", app._close_preview)
        self.job = None
        self.photo = None
        self.item = None
        self.page = None
        self.zoom = 1.0
        self.offset = (0, 0)
        self.anchor = None
        self.drag = None

    def open(self, page):
        self.cancel_callback()
        self.app.render_worker.cancel_preview()
        self.app.preview_request_id += 1
        self.page, self.zoom = page, 1.0
        self.photo = self.item = self.anchor = None
        self.canvas.delete("all")
        self.window.title(self.text["preview_title"].format(page=page + 1))
        self.window.deiconify()
        self.window.lift()
        self.window.focus_set()
        self.zoom_var.set(self.text["preview_zoom"].format(percent=100))
        self.message(self.text["preview_loading"])
        self.schedule(0)

    def message(self, message):
        self.canvas.delete("message")
        self.canvas.create_text(max(1, self.canvas.winfo_width()) / 2,
                                max(1, self.canvas.winfo_height()) / 2,
                                text=message, fill=self.colors["muted"], tags="message",
                                width=max(160, self.canvas.winfo_width() - 32),
                                font=(self.app.font_family, 10))

    def cancel_callback(self):
        if self.job is not None:
            self.app.after_cancel(self.job)
            self.job = None

    def schedule(self, delay=100):
        self.cancel_callback()
        # Invalidate immediately, not only when the debounce timer expires.
        self.app.preview_request_id += 1
        self.app.render_worker.cancel_preview()
        self.job = self.app.after(delay, self.request)

    def request(self):
        self.job = None
        if self.page is None:
            return
        width = max(1, self.canvas.winfo_width()) * self.zoom
        height = max(1, self.canvas.winfo_height()) * self.zoom
        factor = min(1.0, math.sqrt(PREVIEW_PIXEL_LIMIT / max(1, width * height)))
        box = (max(1, int(width * factor)), max(1, int(height * factor)))
        self.app.render_worker.request_preview(self.app.render_generation,
                                               self.app.preview_request_id, self.page, box)

    def point_anchor(self, x, y):
        if self.photo is None:
            return None
        return (min(1, max(0, (x - self.offset[0]) / self.photo.width())),
                min(1, max(0, (y - self.offset[1]) / self.photo.height())), x, y)

    def position(self, x, y):
        if self.photo is None:
            return
        w, h = self.canvas.winfo_width(), self.canvas.winfo_height()
        iw, ih = self.photo.width(), self.photo.height()
        x = (w - iw) / 2 if iw <= w else max(w - iw, min(0, x))
        y = (h - ih) / 2 if ih <= h else max(h - ih, min(0, y))
        self.offset = (x, y)
        self.canvas.coords(self.item, x, y)

    def accept(self, event):
        kind, generation, request_id, page, payload = event
        if (generation != self.app.render_generation or request_id != self.app.preview_request_id
                or page != self.page):
            return
        if kind == "preview_failed":
            self.message(self.text["preview_error"].format(reason=payload))
            return
        self.photo = tk.PhotoImage(master=self.window, data=payload, format="PPM")
        if self.item is None:
            self.item = self.canvas.create_image(0, 0, anchor="nw", image=self.photo)
        else:
            self.canvas.itemconfigure(self.item, image=self.photo)
        self.canvas.delete("message")
        if self.anchor is None:
            self.position(0, 0)
        else:
            rx, ry, x, y = self.anchor
            self.position(x - rx * self.photo.width(), y - ry * self.photo.height())
        self.anchor = None

    def on_wheel(self, event):
        if self.photo is None or not event.delta:
            return "break"
        zoom = max(.5, min(3.0, self.zoom * (1.2 if event.delta > 0 else 1 / 1.2)))
        if zoom == self.zoom:
            return "break"
        self.anchor = self.point_anchor(event.x, event.y)
        self.zoom = zoom
        self.zoom_var.set(self.text["preview_zoom"].format(percent=round(self.zoom * 100)))
        self.schedule()
        return "break"

    def pan_start(self, event):
        self.drag = (event.x, event.y, *self.offset)
        return "break"

    def pan_move(self, event):
        if self.drag is not None:
            x, y, ox, oy = self.drag
            self.position(ox + event.x - x, oy + event.y - y)
            self.anchor = self.point_anchor(event.x, event.y)
        return "break"

    def on_resize(self, event):
        if self.page is not None:
            self.anchor = self.point_anchor(event.width / 2, event.height / 2)
            self.schedule()

    def close(self):
        self.cancel_callback()
        self.app.preview_request_id += 1
        self.app.render_worker.cancel_preview()
        self.photo = None
        self.window.destroy()
