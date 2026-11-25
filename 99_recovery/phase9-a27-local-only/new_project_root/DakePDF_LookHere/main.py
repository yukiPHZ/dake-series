# -*- coding: utf-8 -*-
from __future__ import annotations

import base64
import ctypes
import math
import os
import subprocess
import sys
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, font as tkfont, messagebox, ttk

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None  # type: ignore[assignment]

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore

    DND_ENABLED = True
except Exception:
    DND_FILES = None
    TkinterDnD = None
    DND_ENABLED = False


APP_NAME = "DakePDFここ見て"
WINDOW_TITLE = "DakePDFここ見て"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "brand_series": "シンプルそれDAKEシリーズ",
    "main_title": "PDFにここ見てを付ける",
    "main_description": "PDFに丸と矢印だけを付けて、確認箇所を伝えます。",
    "button_open": "PDFを開く",
    "button_circle": "○ 丸",
    "button_arrow": "→ 矢印",
    "button_undo": "戻す",
    "button_save": "保存",
    "button_prev": "前へ",
    "button_next": "次へ",
    "empty_title": "PDFを開いてください",
    "empty_subtitle": "PDFを選択すると、ここに表示されます。",
    "status_idle": "未選択",
    "status_loading": "読み込み中",
    "status_ready": "準備完了",
    "status_circle": "丸を付ける場所をドラッグしてください",
    "status_arrow": "矢印を引く方向へドラッグしてください",
    "status_saving": "保存中",
    "status_complete": "保存完了",
    "status_error": "エラー",
    "status_undo": "1つ戻しました",
    "status_no_undo": "戻すものがありません",
    "status_page": "{current} / {total} ページ",
    "dialog_open_title": "PDFを選択",
    "dialog_save_title": "保存先を選択",
    "dialog_open_error_title": "PDFを開けませんでした",
    "dialog_save_error_title": "保存できませんでした",
    "dialog_complete_title": "保存しました",
    "dialog_complete_message": "PDFを保存しました。",
    "dialog_pdf_filter_label": "PDFファイル",
    "message_non_pdf": "PDFファイルを選択してください。",
    "message_no_pdf": "先にPDFを開いてください。",
    "message_pymupdf_missing": "PyMuPDF が見つかりません。pip install pymupdf を実行してください。",
    "message_no_pages": "PDFのページを読み込めませんでした。",
    "message_same_file": "元PDFとは別名で保存してください。",
    "message_unknown_error": "原因を特定できませんでした。",
    "save_suffix": "_ここ見て",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_link_1": "戸建買取査定",
    "footer_link_2": "Instagram",
    "footer_separator": " ｜ ",
    "footer_copyright": COPYRIGHT,
}

THEME = {
    "background": "#F6F7F9",
    "card": "#FFFFFF",
    "text": "#1E2430",
    "muted": "#667085",
    "border": "#E6EAF0",
    "accent": "#2F6FED",
    "accent_hover": "#2458BF",
    "selection_bg": "#EAF2FF",
    "disabled_bg": "#EEF2F7",
    "disabled_text": "#98A2B3",
    "danger": "#D92D20",
    "danger_bg": "#FDECEC",
    "success": "#12B76A",
    "success_bg": "#EAFBF3",
}

FONT_CANDIDATES = ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo")
WINDOW_SIZE = "1080x780"
WINDOW_MIN_SIZE = (860, 640)
PAGE_MARGIN = 24
MAX_RENDER_SCALE = 4.0
MIN_RENDER_SCALE = 0.25
ZOOM_STEP = 1.12
RED_HEX = "#E11919"
RED_RGB = (1.0, 0.0, 0.0)
PDF_LINE_WIDTH = 3.0
DISPLAY_LINE_WIDTH = 3
CLICK_CIRCLE_RADIUS = 18.0
ARROW_HEAD_ANGLE = math.radians(28)
APP_USER_MODEL_ID = "Shimarisu.DakePDFLookHere"


@dataclass
class Mark:
    kind: str
    page_index: int
    rect: tuple[float, float, float, float] | None = None
    start: tuple[float, float] | None = None
    end: tuple[float, float] | None = None


def make_root() -> tk.Tk:
    if DND_ENABLED and TkinterDnD is not None:
        return TkinterDnD.Tk()
    return tk.Tk()


def set_windows_app_id() -> None:
    if not sys.platform.startswith("win"):
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)
    except Exception:
        pass


def detect_font_family(root: tk.Misc) -> str:
    try:
        families = set(tkfont.families(root))
    except Exception:
        families = set()
    for name in FONT_CANDIDATES:
        if name in families:
            return name
    return "TkDefaultFont"


def icon_candidates() -> list[Path]:
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        return [
            exe_dir / ".." / ".." / ".." / "02_assets" / "dake_icon.ico",
            exe_dir / ".." / ".." / "02_assets" / "dake_icon.ico",
            Path(getattr(sys, "_MEIPASS", exe_dir)) / "dake_icon.ico",
        ]
    base = Path(__file__).resolve().parent
    return [
        base / ".." / ".." / "02_assets" / "dake_icon.ico",
        Path("..") / ".." / "02_assets" / "dake_icon.ico",
    ]


def apply_window_icon(root: tk.Tk) -> None:
    for candidate in icon_candidates():
        try:
            icon_path = candidate.resolve()
        except Exception:
            icon_path = candidate
        if not icon_path.exists():
            continue
        try:
            root.iconbitmap(str(icon_path))
            root.iconbitmap(default=str(icon_path))
            return
        except Exception:
            continue


def humanize_error(exc: Exception) -> str:
    text = str(exc).strip().replace("\n", " ")
    return text or UI_TEXT["message_unknown_error"]


def open_folder(path: Path) -> None:
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception:
        pass


def normalize_rect(rect: tuple[float, float, float, float]) -> tuple[float, float, float, float]:
    x0, y0, x1, y1 = rect
    return min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1)


def clamp_point_to_rect(point: tuple[float, float], rect: "fitz.Rect") -> tuple[float, float]:
    x, y = point
    return min(max(x, rect.x0), rect.x1), min(max(y, rect.y0), rect.y1)


def clamp_rect_to_page(
    raw_rect: tuple[float, float, float, float], page_rect: "fitz.Rect"
) -> tuple[float, float, float, float]:
    x0, y0, x1, y1 = normalize_rect(raw_rect)
    return (
        min(max(x0, page_rect.x0), page_rect.x1),
        min(max(y0, page_rect.y0), page_rect.y1),
        min(max(x1, page_rect.x0), page_rect.x1),
        min(max(y1, page_rect.y0), page_rect.y1),
    )


class LookHereApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.font_family = detect_font_family(root)
        self.doc: "fitz.Document | None" = None
        self.pdf_path: Path | None = None
        self.page_index = 0
        self.zoom = 1.0
        self.render_scale = 1.0
        self.page_rect: "fitz.Rect | None" = None
        self.image_x = PAGE_MARGIN
        self.image_y = PAGE_MARGIN
        self.image_width = 0
        self.image_height = 0
        self.page_image: tk.PhotoImage | None = None
        self.marks: list[Mark] = []
        self.mode: str | None = None
        self.busy = False
        self.drag_start_pdf: tuple[float, float] | None = None
        self.preview_id: int | None = None
        self.resize_after_id: str | None = None

        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(*WINDOW_MIN_SIZE)
        self.root.configure(bg=THEME["background"])
        apply_window_icon(self.root)

        self._build_styles()
        self._build_ui()
        self._bind_events()
        self._update_status("status_idle")
        self._update_buttons()
        self._show_empty_canvas()

    def _build_styles(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(
            "LookHere.Vertical.TScrollbar",
            gripcount=0,
            background=THEME["border"],
            troughcolor=THEME["card"],
            bordercolor=THEME["card"],
            arrowcolor=THEME["muted"],
        )
        style.configure(
            "LookHere.Horizontal.TScrollbar",
            gripcount=0,
            background=THEME["border"],
            troughcolor=THEME["card"],
            bordercolor=THEME["card"],
            arrowcolor=THEME["muted"],
        )

    def _build_ui(self) -> None:
        base_font = (self.font_family, 10)
        self.title_font = (self.font_family, 18, "bold")
        self.subtitle_font = (self.font_family, 10)
        self.button_font = (self.font_family, 10, "bold")
        self.small_font = (self.font_family, 9)
        self.status_font = (self.font_family, 9, "bold")
        self.empty_title_font = (self.font_family, 17, "bold")

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        header = tk.Frame(self.root, bg=THEME["background"])
        header.grid(row=0, column=0, sticky="ew", padx=22, pady=(18, 10))
        header.columnconfigure(0, weight=1)

        title_block = tk.Frame(header, bg=THEME["background"])
        title_block.grid(row=0, column=0, sticky="w")

        tk.Label(
            title_block,
            text=UI_TEXT["brand_series"],
            font=self.small_font,
            fg=THEME["accent"],
            bg=THEME["background"],
        ).pack(anchor="w")
        tk.Label(
            title_block,
            text=UI_TEXT["main_title"],
            font=self.title_font,
            fg=THEME["text"],
            bg=THEME["background"],
        ).pack(anchor="w", pady=(2, 0))
        tk.Label(
            title_block,
            text=UI_TEXT["main_description"],
            font=self.subtitle_font,
            fg=THEME["muted"],
            bg=THEME["background"],
        ).pack(anchor="w", pady=(4, 0))

        toolbar = tk.Frame(header, bg=THEME["card"], highlightthickness=1, highlightbackground=THEME["border"])
        toolbar.grid(row=1, column=0, sticky="ew", pady=(14, 0))
        for index in range(8):
            toolbar.columnconfigure(index, weight=0)
        toolbar.columnconfigure(8, weight=1)

        self.open_button = self._make_button(toolbar, "button_open", self.open_pdf, primary=True)
        self.open_button.grid(row=0, column=0, padx=(12, 6), pady=10)
        self.circle_button = self._make_button(toolbar, "button_circle", lambda: self.set_mode("circle"))
        self.circle_button.grid(row=0, column=1, padx=6, pady=10)
        self.arrow_button = self._make_button(toolbar, "button_arrow", lambda: self.set_mode("arrow"))
        self.arrow_button.grid(row=0, column=2, padx=6, pady=10)
        self.undo_button = self._make_button(toolbar, "button_undo", self.undo)
        self.undo_button.grid(row=0, column=3, padx=6, pady=10)
        self.save_button = self._make_button(toolbar, "button_save", self.save_pdf, primary=True)
        self.save_button.grid(row=0, column=4, padx=(6, 14), pady=10)

        divider = tk.Frame(toolbar, width=1, bg=THEME["border"])
        divider.grid(row=0, column=5, sticky="ns", pady=12)

        self.prev_button = self._make_button(toolbar, "button_prev", self.prev_page)
        self.prev_button.grid(row=0, column=6, padx=(14, 6), pady=10)
        self.next_button = self._make_button(toolbar, "button_next", self.next_page)
        self.next_button.grid(row=0, column=7, padx=6, pady=10)
        self.page_label = tk.Label(
            toolbar,
            text="",
            font=base_font,
            fg=THEME["muted"],
            bg=THEME["card"],
        )
        self.page_label.grid(row=0, column=8, sticky="e", padx=(8, 14))

        viewer = tk.Frame(self.root, bg=THEME["card"], highlightthickness=1, highlightbackground=THEME["border"])
        viewer.grid(row=1, column=0, sticky="nsew", padx=22, pady=(0, 10))
        viewer.columnconfigure(0, weight=1)
        viewer.rowconfigure(0, weight=1)

        self.canvas = tk.Canvas(
            viewer,
            bg=THEME["card"],
            bd=0,
            highlightthickness=0,
            xscrollincrement=1,
            yscrollincrement=1,
        )
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.v_scroll = ttk.Scrollbar(
            viewer,
            orient="vertical",
            command=self.canvas.yview,
            style="LookHere.Vertical.TScrollbar",
        )
        self.v_scroll.grid(row=0, column=1, sticky="ns")
        self.h_scroll = ttk.Scrollbar(
            viewer,
            orient="horizontal",
            command=self.canvas.xview,
            style="LookHere.Horizontal.TScrollbar",
        )
        self.h_scroll.grid(row=1, column=0, sticky="ew")
        self.canvas.configure(yscrollcommand=self.v_scroll.set, xscrollcommand=self.h_scroll.set)

        footer = tk.Frame(self.root, bg=THEME["background"])
        footer.grid(row=2, column=0, sticky="ew", padx=22, pady=(0, 14))
        footer.columnconfigure(1, weight=1)

        self.status_badge = tk.Label(
            footer,
            text="",
            font=self.status_font,
            fg=THEME["muted"],
            bg=THEME["disabled_bg"],
            padx=12,
            pady=4,
        )
        self.status_badge.grid(row=0, column=0, sticky="w")
        self.status_detail = tk.Label(
            footer,
            text="",
            font=self.small_font,
            fg=THEME["muted"],
            bg=THEME["background"],
        )
        self.status_detail.grid(row=0, column=1, sticky="w", padx=(10, 0))

        footer_text = (
            UI_TEXT["footer_left"]
            + UI_TEXT["footer_separator"]
            + UI_TEXT["footer_link_1"]
            + UI_TEXT["footer_separator"]
            + UI_TEXT["footer_link_2"]
            + UI_TEXT["footer_separator"]
            + UI_TEXT["footer_copyright"]
        )
        tk.Label(
            footer,
            text=footer_text,
            font=self.small_font,
            fg=THEME["muted"],
            bg=THEME["background"],
        ).grid(row=0, column=2, sticky="e")

    def _make_button(
        self,
        parent: tk.Misc,
        text_key: str,
        command: object,
        primary: bool = False,
    ) -> tk.Button:
        return tk.Button(
            parent,
            text=UI_TEXT[text_key],
            command=command,
            font=self.button_font,
            fg=THEME["card"] if primary else THEME["text"],
            bg=THEME["accent"] if primary else THEME["card"],
            activeforeground=THEME["card"] if primary else THEME["text"],
            activebackground=THEME["accent_hover"] if primary else THEME["selection_bg"],
            disabledforeground=THEME["disabled_text"],
            relief="flat",
            bd=0,
            padx=14,
            pady=7,
            cursor="hand2",
        )

    def _bind_events(self) -> None:
        self.canvas.bind("<ButtonPress-1>", self._on_press)
        self.canvas.bind("<B1-Motion>", self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind("<Control-MouseWheel>", self._on_ctrl_mousewheel)
        self.root.bind("<Control-MouseWheel>", self._on_ctrl_mousewheel)
        self.canvas.bind("<Control-Button-4>", lambda event: self._zoom_at_mouse(1))
        self.canvas.bind("<Control-Button-5>", lambda event: self._zoom_at_mouse(-1))
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        if DND_ENABLED and DND_FILES is not None and hasattr(self.root, "drop_target_register"):
            try:
                self.root.drop_target_register(DND_FILES)
                self.root.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                pass

    def _on_drop(self, event: object) -> None:
        data = getattr(event, "data", "")
        try:
            paths = self.root.tk.splitlist(data)
        except Exception:
            paths = [data]
        if not paths:
            return
        self.load_pdf(Path(paths[0]))

    def _set_busy(self, busy: bool) -> None:
        self.busy = busy
        self._update_buttons()
        self.root.update_idletasks()

    def _update_status(self, key: str, detail: str | None = None) -> None:
        text = UI_TEXT.get(key, key)
        if key == "status_error":
            bg, fg = THEME["danger_bg"], THEME["danger"]
        elif key == "status_complete":
            bg, fg = THEME["success_bg"], THEME["success"]
        elif key in {"status_loading", "status_saving", "status_circle", "status_arrow"}:
            bg, fg = THEME["selection_bg"], THEME["accent"]
        else:
            bg, fg = THEME["disabled_bg"], THEME["muted"]
        self.status_badge.configure(text=text, bg=bg, fg=fg)
        self.status_detail.configure(text=detail or text)

    def _update_buttons(self) -> None:
        has_doc = self.doc is not None
        normal_if_doc = tk.NORMAL if has_doc and not self.busy else tk.DISABLED
        self.open_button.configure(state=tk.DISABLED if self.busy else tk.NORMAL)
        self.circle_button.configure(state=normal_if_doc)
        self.arrow_button.configure(state=normal_if_doc)
        self.undo_button.configure(state=tk.NORMAL if self.marks and not self.busy else tk.DISABLED)
        self.save_button.configure(state=normal_if_doc)

        page_count = self.doc.page_count if self.doc is not None else 0
        self.prev_button.configure(state=tk.NORMAL if has_doc and self.page_index > 0 and not self.busy else tk.DISABLED)
        self.next_button.configure(
            state=tk.NORMAL if has_doc and self.page_index < page_count - 1 and not self.busy else tk.DISABLED
        )
        if has_doc:
            self.page_label.configure(
                text=UI_TEXT["status_page"].format(current=self.page_index + 1, total=page_count)
            )
        else:
            self.page_label.configure(text="")

        self._style_mode_button(self.circle_button, self.mode == "circle")
        self._style_mode_button(self.arrow_button, self.mode == "arrow")

    def _style_mode_button(self, button: tk.Button, selected: bool) -> None:
        if selected and str(button.cget("state")) != tk.DISABLED:
            button.configure(bg=THEME["selection_bg"], fg=THEME["accent"], activebackground=THEME["selection_bg"])
        else:
            button.configure(bg=THEME["card"], fg=THEME["text"], activebackground=THEME["selection_bg"])

    def open_pdf(self) -> None:
        filename = filedialog.askopenfilename(
            parent=self.root,
            title=UI_TEXT["dialog_open_title"],
            filetypes=[(UI_TEXT["dialog_pdf_filter_label"], "*.pdf")],
        )
        if filename:
            self.load_pdf(Path(filename))

    def load_pdf(self, path: Path) -> None:
        if self.busy:
            return
        if path.suffix.lower() != ".pdf":
            messagebox.showerror(UI_TEXT["dialog_open_error_title"], UI_TEXT["message_non_pdf"], parent=self.root)
            self._update_status("status_error", UI_TEXT["message_non_pdf"])
            return
        if fitz is None:
            messagebox.showerror(
                UI_TEXT["dialog_open_error_title"], UI_TEXT["message_pymupdf_missing"], parent=self.root
            )
            self._update_status("status_error", UI_TEXT["message_pymupdf_missing"])
            return

        self._set_busy(True)
        self._update_status("status_loading")
        try:
            new_doc = fitz.open(str(path))
            if new_doc.page_count < 1:
                new_doc.close()
                raise ValueError(UI_TEXT["message_no_pages"])
            if self.doc is not None:
                self.doc.close()
            self.doc = new_doc
            self.pdf_path = path
            self.page_index = 0
            self.zoom = 1.0
            self.marks.clear()
            self.mode = "circle"
            self.render_page()
            self._update_status("status_circle", path.name)
        except Exception as exc:
            message = humanize_error(exc)
            messagebox.showerror(UI_TEXT["dialog_open_error_title"], message, parent=self.root)
            self._update_status("status_error", message)
            self.doc = None
            self.pdf_path = None
            self.page_index = 0
            self.mode = None
            self._show_empty_canvas()
        finally:
            self._set_busy(False)

    def set_mode(self, mode: str) -> None:
        if self.doc is None or self.busy:
            return
        self.mode = mode
        if mode == "circle":
            self._update_status("status_circle")
        elif mode == "arrow":
            self._update_status("status_arrow")
        self._update_buttons()

    def undo(self) -> None:
        if self.busy:
            return
        if not self.marks:
            self._update_status("status_no_undo")
            return
        self.marks.pop()
        self.render_page()
        self._update_status("status_undo")
        self._update_buttons()

    def prev_page(self) -> None:
        if self.doc is None or self.busy or self.page_index <= 0:
            return
        self.page_index -= 1
        self.render_page()
        self._update_status("status_ready")

    def next_page(self) -> None:
        if self.doc is None or self.busy or self.page_index >= self.doc.page_count - 1:
            return
        self.page_index += 1
        self.render_page()
        self._update_status("status_ready")

    def save_pdf(self) -> None:
        if self.busy:
            return
        if self.doc is None or self.pdf_path is None:
            messagebox.showerror(UI_TEXT["dialog_save_error_title"], UI_TEXT["message_no_pdf"], parent=self.root)
            return

        initial_name = f"{self.pdf_path.stem}{UI_TEXT['save_suffix']}.pdf"
        filename = filedialog.asksaveasfilename(
            parent=self.root,
            title=UI_TEXT["dialog_save_title"],
            initialdir=str(self.pdf_path.parent),
            initialfile=initial_name,
            defaultextension=".pdf",
            filetypes=[(UI_TEXT["dialog_pdf_filter_label"], "*.pdf")],
        )
        if not filename:
            return

        output_path = Path(filename)
        try:
            if output_path.resolve() == self.pdf_path.resolve():
                messagebox.showerror(UI_TEXT["dialog_save_error_title"], UI_TEXT["message_same_file"], parent=self.root)
                self._update_status("status_error", UI_TEXT["message_same_file"])
                return
        except Exception:
            pass

        self._set_busy(True)
        self._update_status("status_saving")
        try:
            with fitz.open(str(self.pdf_path)) as output_doc:
                for mark in self.marks:
                    if mark.page_index < 0 or mark.page_index >= output_doc.page_count:
                        continue
                    page = output_doc.load_page(mark.page_index)
                    if mark.kind == "circle" and mark.rect is not None:
                        rect_values = clamp_rect_to_page(mark.rect, page.rect)
                        page.draw_oval(fitz.Rect(rect_values), color=RED_RGB, width=PDF_LINE_WIDTH, overlay=True)
                    elif mark.kind == "arrow" and mark.start is not None and mark.end is not None:
                        start = clamp_point_to_rect(mark.start, page.rect)
                        end = clamp_point_to_rect(mark.end, page.rect)
                        self._draw_pdf_arrow(page, start, end)
                output_doc.save(str(output_path), garbage=4, deflate=True)

            self._update_status("status_complete", str(output_path))
            messagebox.showinfo(
                UI_TEXT["dialog_complete_title"],
                UI_TEXT["dialog_complete_message"],
                parent=self.root,
            )
            open_folder(output_path.parent)
        except Exception as exc:
            message = humanize_error(exc)
            messagebox.showerror(UI_TEXT["dialog_save_error_title"], message, parent=self.root)
            self._update_status("status_error", message)
        finally:
            self._set_busy(False)

    def render_page(self) -> None:
        if self.doc is None:
            self._show_empty_canvas()
            return
        page = self.doc.load_page(self.page_index)
        self.page_rect = page.rect
        canvas_width = max(self.canvas.winfo_width(), 480)
        available_width = max(canvas_width - PAGE_MARGIN * 2, 240)
        fit_scale = available_width / max(page.rect.width, 1.0)
        self.render_scale = min(max(fit_scale * self.zoom, MIN_RENDER_SCALE), MAX_RENDER_SCALE)
        matrix = fitz.Matrix(self.render_scale, self.render_scale)
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        png_data = pix.tobytes("png")
        encoded = base64.b64encode(png_data).decode("ascii")
        self.page_image = tk.PhotoImage(data=encoded, format="PNG")
        self.image_width = self.page_image.width()
        self.image_height = self.page_image.height()

        self.canvas.delete("all")
        visible_width = max(self.canvas.winfo_width(), self.image_width + PAGE_MARGIN * 2)
        self.image_x = max(PAGE_MARGIN, (visible_width - self.image_width) // 2)
        self.image_y = PAGE_MARGIN

        shadow_offset = 2
        self.canvas.create_rectangle(
            self.image_x + shadow_offset,
            self.image_y + shadow_offset,
            self.image_x + self.image_width + shadow_offset,
            self.image_y + self.image_height + shadow_offset,
            fill=THEME["border"],
            outline="",
        )
        self.canvas.create_rectangle(
            self.image_x - 1,
            self.image_y - 1,
            self.image_x + self.image_width + 1,
            self.image_y + self.image_height + 1,
            fill=THEME["card"],
            outline=THEME["border"],
        )
        self.canvas.create_image(self.image_x, self.image_y, anchor="nw", image=self.page_image)
        self._draw_marks()
        self.canvas.configure(
            scrollregion=(
                0,
                0,
                max(visible_width, self.image_x + self.image_width + PAGE_MARGIN),
                self.image_y + self.image_height + PAGE_MARGIN,
            )
        )
        self._update_buttons()

    def _show_empty_canvas(self) -> None:
        self.canvas.delete("all")
        width = max(self.canvas.winfo_width(), 480)
        height = max(self.canvas.winfo_height(), 360)
        self.canvas.configure(scrollregion=(0, 0, width, height))
        self.canvas.create_text(
            width / 2,
            height / 2 - 18,
            text=UI_TEXT["empty_title"],
            font=self.empty_title_font,
            fill=THEME["text"],
        )
        self.canvas.create_text(
            width / 2,
            height / 2 + 18,
            text=UI_TEXT["empty_subtitle"],
            font=self.subtitle_font,
            fill=THEME["muted"],
        )

    def _draw_marks(self) -> None:
        for mark in self.marks:
            if mark.page_index != self.page_index:
                continue
            if mark.kind == "circle" and mark.rect is not None:
                x0, y0, x1, y1 = self._page_rect_to_canvas(mark.rect)
                self.canvas.create_oval(x0, y0, x1, y1, outline=RED_HEX, width=DISPLAY_LINE_WIDTH)
            elif mark.kind == "arrow" and mark.start is not None and mark.end is not None:
                x0, y0 = self._page_point_to_canvas(mark.start)
                x1, y1 = self._page_point_to_canvas(mark.end)
                self.canvas.create_line(
                    x0,
                    y0,
                    x1,
                    y1,
                    fill=RED_HEX,
                    width=DISPLAY_LINE_WIDTH,
                    arrow=tk.LAST,
                    arrowshape=(14, 18, 6),
                    capstyle=tk.ROUND,
                )

    def _page_point_to_canvas(self, point: tuple[float, float]) -> tuple[float, float]:
        if self.page_rect is None:
            return self.image_x, self.image_y
        x, y = point
        return (
            self.image_x + (x - self.page_rect.x0) * self.render_scale,
            self.image_y + (y - self.page_rect.y0) * self.render_scale,
        )

    def _page_rect_to_canvas(self, rect: tuple[float, float, float, float]) -> tuple[float, float, float, float]:
        x0, y0 = self._page_point_to_canvas((rect[0], rect[1]))
        x1, y1 = self._page_point_to_canvas((rect[2], rect[3]))
        return x0, y0, x1, y1

    def _event_to_page_point(self, event: tk.Event, require_inside: bool) -> tuple[float, float] | None:
        if self.page_rect is None or self.image_width <= 0 or self.image_height <= 0:
            return None
        x = self.canvas.canvasx(event.x) - self.image_x
        y = self.canvas.canvasy(event.y) - self.image_y
        inside = 0 <= x <= self.image_width and 0 <= y <= self.image_height
        if require_inside and not inside:
            return None
        x = min(max(x, 0), self.image_width)
        y = min(max(y, 0), self.image_height)
        return (
            self.page_rect.x0 + x / self.render_scale,
            self.page_rect.y0 + y / self.render_scale,
        )

    def _on_press(self, event: tk.Event) -> None:
        if self.doc is None or self.busy or self.mode is None:
            return
        point = self._event_to_page_point(event, require_inside=True)
        if point is None:
            return
        self.drag_start_pdf = point
        self._clear_preview()

    def _on_drag(self, event: tk.Event) -> None:
        if self.drag_start_pdf is None or self.mode is None:
            return
        current = self._event_to_page_point(event, require_inside=False)
        if current is None:
            return
        self._clear_preview()
        if self.mode == "circle":
            rect = self._circle_rect_from_points(self.drag_start_pdf, current)
            x0, y0, x1, y1 = self._page_rect_to_canvas(rect)
            self.preview_id = self.canvas.create_oval(
                x0,
                y0,
                x1,
                y1,
                outline=RED_HEX,
                width=DISPLAY_LINE_WIDTH,
                dash=(6, 4),
            )
        elif self.mode == "arrow":
            start_x, start_y = self._page_point_to_canvas(self.drag_start_pdf)
            end_x, end_y = self._page_point_to_canvas(current)
            self.preview_id = self.canvas.create_line(
                start_x,
                start_y,
                end_x,
                end_y,
                fill=RED_HEX,
                width=DISPLAY_LINE_WIDTH,
                arrow=tk.LAST,
                arrowshape=(14, 18, 6),
                dash=(6, 4),
                capstyle=tk.ROUND,
            )

    def _on_release(self, event: tk.Event) -> None:
        if self.drag_start_pdf is None or self.mode is None:
            return
        end_point = self._event_to_page_point(event, require_inside=False)
        if end_point is None:
            self.drag_start_pdf = None
            self._clear_preview()
            return

        if self.mode == "circle":
            rect = self._circle_rect_from_points(self.drag_start_pdf, end_point)
            self.marks.append(Mark(kind="circle", page_index=self.page_index, rect=rect))
            self._update_status("status_circle")
        elif self.mode == "arrow":
            start, end = self._arrow_points_from_drag(self.drag_start_pdf, end_point)
            self.marks.append(Mark(kind="arrow", page_index=self.page_index, start=start, end=end))
            self._update_status("status_arrow")

        self.drag_start_pdf = None
        self._clear_preview()
        self.render_page()
        self._update_buttons()

    def _circle_rect_from_points(
        self, start: tuple[float, float], end: tuple[float, float]
    ) -> tuple[float, float, float, float]:
        if self.page_rect is None:
            return start[0], start[1], end[0], end[1]
        dx = abs(end[0] - start[0]) * self.render_scale
        dy = abs(end[1] - start[1]) * self.render_scale
        if dx < 5 and dy < 5:
            radius = max(CLICK_CIRCLE_RADIUS, 16.0 / max(self.render_scale, 0.1))
            raw = (start[0] - radius, start[1] - radius, start[0] + radius, start[1] + radius)
        else:
            raw = (start[0], start[1], end[0], end[1])
        return clamp_rect_to_page(raw, self.page_rect)

    def _arrow_points_from_drag(
        self, start: tuple[float, float], end: tuple[float, float]
    ) -> tuple[tuple[float, float], tuple[float, float]]:
        if self.page_rect is None:
            return start, end
        dx = (end[0] - start[0]) * self.render_scale
        dy = (end[1] - start[1]) * self.render_scale
        if math.hypot(dx, dy) < 8:
            end = (start[0] + 48.0 / max(self.render_scale, 0.1), start[1])
        return clamp_point_to_rect(start, self.page_rect), clamp_point_to_rect(end, self.page_rect)

    def _draw_pdf_arrow(
        self,
        page: "fitz.Page",
        start: tuple[float, float],
        end: tuple[float, float],
    ) -> None:
        sx, sy = start
        ex, ey = end
        dx = ex - sx
        dy = ey - sy
        length = math.hypot(dx, dy)
        if length < 1:
            return

        page.draw_line(fitz.Point(sx, sy), fitz.Point(ex, ey), color=RED_RGB, width=PDF_LINE_WIDTH, overlay=True)
        angle = math.atan2(dy, dx)
        head_len = min(18.0, max(8.0, length * 0.25))
        for sign in (-1, 1):
            head_angle = angle + math.pi + sign * ARROW_HEAD_ANGLE
            hx = ex + head_len * math.cos(head_angle)
            hy = ey + head_len * math.sin(head_angle)
            page.draw_line(
                fitz.Point(ex, ey),
                fitz.Point(hx, hy),
                color=RED_RGB,
                width=PDF_LINE_WIDTH,
                overlay=True,
            )

    def _clear_preview(self) -> None:
        if self.preview_id is not None:
            try:
                self.canvas.delete(self.preview_id)
            except Exception:
                pass
            self.preview_id = None

    def _on_ctrl_mousewheel(self, event: tk.Event) -> str:
        direction = 1 if getattr(event, "delta", 0) > 0 else -1
        self._zoom_at_mouse(direction)
        return "break"

    def _zoom_at_mouse(self, direction: int) -> None:
        if self.doc is None or self.busy:
            return
        if direction > 0:
            self.zoom = min(self.zoom * ZOOM_STEP, 4.0)
        else:
            self.zoom = max(self.zoom / ZOOM_STEP, 0.35)
        self.render_page()

    def _on_canvas_configure(self, _event: tk.Event) -> None:
        if self.resize_after_id is not None:
            try:
                self.root.after_cancel(self.resize_after_id)
            except Exception:
                pass
        self.resize_after_id = self.root.after(120, self._rerender_after_resize)

    def _rerender_after_resize(self) -> None:
        self.resize_after_id = None
        if self.doc is None:
            self._show_empty_canvas()
        else:
            self.render_page()

    def _on_close(self) -> None:
        if self.doc is not None:
            try:
                self.doc.close()
            except Exception:
                pass
        self.root.destroy()


def main() -> None:
    set_windows_app_id()
    root = make_root()
    LookHereApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
