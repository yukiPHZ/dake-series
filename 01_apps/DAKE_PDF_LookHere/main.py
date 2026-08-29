# -*- coding: utf-8 -*-
from __future__ import annotations

import base64
import ctypes
import math
import os
import queue
import subprocess
import sys
import threading
import webbrowser
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, font as tkfont, messagebox, ttk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore

    DND_ENABLED = True
except Exception:
    DND_FILES = None
    TkinterDnD = None
    DND_ENABLED = False


APP_NAME = "DakePDFここ見て"
WINDOW_TITLE = "PDFここ見て"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "brand_series": "シンプルそれDAKEシリーズ",
    "main_title": "PDFにここ見てを付ける",
    "main_description": "丸と矢印だけで確認箇所を伝えます。",
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
    "footer_tagline": "止まらない、迷わない、すぐ終わる。",
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
FOOTER_BREAKPOINT = 960

_FITZ_MODULE: object | None = None

LINKS = {
    "assessment": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "instagram": "https://instagram.com/kikuta.shimarisu_fudosan",
}


def get_fitz() -> object:
    global _FITZ_MODULE
    if _FITZ_MODULE is None:
        import fitz as fitz_module  # PyInstallerの解析対象に保ったまま利用時だけ読み込む

        _FITZ_MODULE = fitz_module
    return _FITZ_MODULE


@dataclass(frozen=True)
class Mark:
    kind: str
    page_index: int
    rect: tuple[float, float, float, float] | None = None
    start: tuple[float, float] | None = None
    end: tuple[float, float] | None = None


@dataclass(frozen=True)
class RenderRequest:
    generation: int
    pdf_path: Path
    page_index: int
    zoom: float
    canvas_width: int
    canvas_height: int
    opens_document: bool


@dataclass(frozen=True)
class RenderResult:
    request: RenderRequest
    page_count: int = 0
    page_rect: tuple[float, float, float, float] | None = None
    render_scale: float = 1.0
    image_width: int = 0
    image_height: int = 0
    png_data: bytes | None = None
    error: Exception | None = None


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


def rect_values(rect: object) -> tuple[float, float, float, float]:
    if isinstance(rect, tuple):
        return rect
    return (
        float(getattr(rect, "x0")),
        float(getattr(rect, "y0")),
        float(getattr(rect, "x1")),
        float(getattr(rect, "y1")),
    )


def clamp_point_to_rect(point: tuple[float, float], rect: object) -> tuple[float, float]:
    x, y = point
    x0, y0, x1, y1 = rect_values(rect)
    return min(max(x, x0), x1), min(max(y, y0), y1)


def clamp_rect_to_page(
    raw_rect: tuple[float, float, float, float], page_rect: object
) -> tuple[float, float, float, float]:
    x0, y0, x1, y1 = normalize_rect(raw_rect)
    page_x0, page_y0, page_x1, page_y1 = rect_values(page_rect)
    return (
        min(max(x0, page_x0), page_x1),
        min(max(y0, page_y0), page_y1),
        min(max(x1, page_x0), page_x1),
        min(max(y1, page_y0), page_y1),
    )


class LookHereApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.font_family = detect_font_family(root)
        self.pdf_path: Path | None = None
        self.page_count = 0
        self.page_index = 0
        self.zoom = 1.0
        self.render_scale = 1.0
        self.page_rect: tuple[float, float, float, float] | None = None
        self.image_x = PAGE_MARGIN
        self.image_y = PAGE_MARGIN
        self.image_width = 0
        self.image_height = 0
        self.page_image: tk.PhotoImage | None = None
        self.page_shadow_id: int | None = None
        self.page_paper_id: int | None = None
        self.page_image_id: int | None = None
        self.marks: list[Mark] = []
        self.mark_item_ids: dict[int, int] = {}
        self.mode: str | None = None
        self.busy = False
        self.opening_pdf = False
        self.render_pending = False
        self.shutting_down = False
        self.drag_start_pdf: tuple[float, float] | None = None
        self.preview_id: int | None = None
        self.resize_after_id: str | None = None
        self.render_dispatch_after_id: str | None = None
        self.render_after_id: str | None = None
        self.save_after_id: str | None = None
        self.scheduled_render_request: RenderRequest | None = None
        self.pending_render_request: RenderRequest | None = None
        self.render_condition = threading.Condition()
        self.render_results: queue.Queue[RenderResult] = queue.Queue()
        self.save_results: queue.Queue[tuple[Path, Exception | None]] = queue.Queue()
        self.render_stop_requested = False
        self.render_generation = 0
        self.displayed_generation = 0
        self.displayed_pdf_path: Path | None = None
        self.displayed_page_index = -1
        self.render_requests_submitted = 0
        self.render_jobs_replaced = 0
        self.render_jobs_started = 0
        self.render_results_discarded = 0
        self.render_documents_opened = 0
        self.save_worker: threading.Thread | None = None
        self.render_worker = threading.Thread(target=self._render_worker_loop, daemon=True)

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
        self.render_worker.start()
        self.render_after_id = self.root.after(30, self._poll_render_results)

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
        title_block.grid(row=0, column=0, sticky="ew")
        title_block.columnconfigure(0, weight=0)
        title_block.columnconfigure(1, weight=1)

        self.header_title = tk.Label(
            title_block,
            text=UI_TEXT["main_title"],
            font=self.title_font,
            fg=THEME["text"],
            bg=THEME["background"],
        )
        self.header_title.grid(row=0, column=0, sticky="w")
        self.header_description = tk.Label(
            title_block,
            text=UI_TEXT["main_description"],
            font=self.subtitle_font,
            fg=THEME["muted"],
            bg=THEME["background"],
        )
        self.header_description.grid(row=0, column=1, sticky="w", padx=(14, 0), pady=(4, 0))

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
        footer.columnconfigure(0, weight=1)

        self.status_row = tk.Frame(footer, bg=THEME["background"])
        self.status_row.grid(row=0, column=0, sticky="ew")
        self.status_row.columnconfigure(1, weight=1)

        self.status_badge = tk.Label(
            self.status_row,
            text="",
            font=self.status_font,
            fg=THEME["muted"],
            bg=THEME["disabled_bg"],
            padx=12,
            pady=4,
        )
        self.status_badge.grid(row=0, column=0, sticky="w")
        self.status_detail = tk.Label(
            self.status_row,
            text="",
            font=self.small_font,
            fg=THEME["muted"],
            bg=THEME["background"],
        )
        self.status_detail.grid(row=0, column=1, sticky="w", padx=(10, 0))

        self.footer_container = tk.Frame(footer, bg=THEME["background"])
        self.footer_container.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        self.footer_container.columnconfigure(0, weight=1)
        self.footer_container.columnconfigure(1, weight=1)

        self.footer_left_line = tk.Frame(self.footer_container, bg=THEME["background"])
        self.footer_left_text = tk.Label(
            self.footer_left_line,
            text=UI_TEXT["footer_left"] + UI_TEXT["footer_separator"] + UI_TEXT["footer_tagline"],
            font=self.small_font,
            fg=THEME["muted"],
            bg=THEME["background"],
        )
        self.footer_left_text.pack(side="left")

        self.footer_right_line = tk.Frame(self.footer_container, bg=THEME["background"])
        self.footer_link_1 = self._make_footer_link(
            self.footer_right_line,
            UI_TEXT["footer_link_1"],
            LINKS["assessment"],
        )
        self.footer_link_1.pack(side="left")
        self._make_footer_separator(self.footer_right_line).pack(side="left")
        self.footer_link_2 = self._make_footer_link(
            self.footer_right_line,
            UI_TEXT["footer_link_2"],
            LINKS["instagram"],
        )
        self.footer_link_2.pack(side="left")
        self._make_footer_separator(self.footer_right_line).pack(side="left")
        tk.Label(
            self.footer_right_line,
            text=UI_TEXT["footer_copyright"],
            font=self.small_font,
            fg=THEME["muted"],
            bg=THEME["background"],
        ).pack(side="left")
        self.root.after(80, lambda: self._layout_footer(self.root.winfo_width()))

    def _make_footer_separator(self, parent: tk.Misc) -> tk.Label:
        return tk.Label(
            parent,
            text=UI_TEXT["footer_separator"],
            font=self.small_font,
            fg=THEME["muted"],
            bg=THEME["background"],
        )

    def _make_footer_link(self, parent: tk.Misc, text: str, url: str) -> tk.Label:
        label = tk.Label(
            parent,
            text=text,
            font=self.small_font,
            fg=THEME["muted"],
            bg=THEME["background"],
            cursor="hand2",
        )
        label.bind("<Button-1>", lambda _event: self._open_footer_link(url))
        label.bind("<Enter>", lambda _event: label.configure(fg=THEME["accent"]))
        label.bind("<Leave>", lambda _event: label.configure(fg=THEME["muted"]))
        return label

    def _open_footer_link(self, url: str) -> None:
        try:
            webbrowser.open(url)
        except Exception:
            pass

    def _layout_footer(self, width: int) -> None:
        if not hasattr(self, "footer_left_line"):
            return
        self.footer_left_line.grid_forget()
        self.footer_right_line.grid_forget()
        if width >= FOOTER_BREAKPOINT:
            self.footer_container.columnconfigure(0, weight=1)
            self.footer_container.columnconfigure(1, weight=1)
            self.footer_left_line.grid(row=0, column=0, sticky="w")
            self.footer_right_line.grid(row=0, column=1, sticky="e")
        else:
            self.footer_container.columnconfigure(0, weight=1)
            self.footer_container.columnconfigure(1, weight=0)
            self.footer_left_line.grid(row=0, column=0, sticky="")
            self.footer_right_line.grid(row=1, column=0, sticky="", pady=(4, 0))

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
        self.root.bind("<Configure>", self._on_root_configure, add="+")

        if DND_ENABLED and DND_FILES is not None and hasattr(self.root, "drop_target_register"):
            try:
                self.root.drop_target_register(DND_FILES)
                self.root.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                pass

    def _on_root_configure(self, event: tk.Event) -> None:
        if event.widget is self.root:
            self._layout_footer(event.width)

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
        has_doc = self.pdf_path is not None and self.page_count > 0
        stable_preview = (
            has_doc
            and not self.busy
            and not self.opening_pdf
            and not self.render_pending
            and self.displayed_pdf_path == self.pdf_path
            and self.displayed_page_index == self.page_index
        )
        normal_if_stable = tk.NORMAL if stable_preview else tk.DISABLED
        self.open_button.configure(state=tk.DISABLED if self.busy or self.opening_pdf else tk.NORMAL)
        self.circle_button.configure(state=normal_if_stable)
        self.arrow_button.configure(state=normal_if_stable)
        self.undo_button.configure(state=tk.NORMAL if self.marks and not self.busy else tk.DISABLED)
        self.save_button.configure(state=tk.NORMAL if has_doc and not self.busy and not self.opening_pdf else tk.DISABLED)

        page_count = self.page_count
        self.prev_button.configure(
            state=tk.NORMAL if has_doc and self.page_index > 0 and not self.busy and not self.opening_pdf else tk.DISABLED
        )
        self.next_button.configure(
            state=tk.NORMAL
            if has_doc and self.page_index < page_count - 1 and not self.busy and not self.opening_pdf
            else tk.DISABLED
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
        self._update_status("status_loading")
        self.opening_pdf = True
        self._submit_render_request(path, page_index=0, zoom=1.0, opens_document=True)

    def set_mode(self, mode: str) -> None:
        if self.pdf_path is None or self.busy or self.opening_pdf or self.render_pending:
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
        removed_index = len(self.marks) - 1
        self.marks.pop()
        item_id = self.mark_item_ids.pop(removed_index, None)
        if item_id is not None:
            self.canvas.delete(item_id)
        self._update_status("status_undo")
        self._update_buttons()

    def prev_page(self) -> None:
        if self.pdf_path is None or self.busy or self.opening_pdf or self.page_index <= 0:
            return
        self.page_index -= 1
        self.render_page()
        self._update_status("status_loading")

    def next_page(self) -> None:
        if self.pdf_path is None or self.busy or self.opening_pdf or self.page_index >= self.page_count - 1:
            return
        self.page_index += 1
        self.render_page()
        self._update_status("status_loading")

    def save_pdf(self) -> None:
        if self.busy:
            return
        if self.pdf_path is None or self.page_count < 1:
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
        marks = list(self.marks)
        self.save_worker = threading.Thread(
            target=self._save_worker,
            args=(self.pdf_path, output_path, marks),
            daemon=False,
        )
        self.save_worker.start()
        self._poll_save_result()

    def _save_worker(self, input_path: Path, output_path: Path, marks: list[Mark]) -> None:
        error: Exception | None = None
        try:
            fitz_module = get_fitz()
            with fitz_module.open(str(input_path)) as output_doc:
                for mark in marks:
                    if mark.page_index < 0 or mark.page_index >= output_doc.page_count:
                        continue
                    page = output_doc.load_page(mark.page_index)
                    if mark.kind == "circle" and mark.rect is not None:
                        rect_values = clamp_rect_to_page(mark.rect, page.rect)
                        page.draw_oval(
                            fitz_module.Rect(rect_values),
                            color=RED_RGB,
                            width=PDF_LINE_WIDTH,
                            overlay=True,
                        )
                    elif mark.kind == "arrow" and mark.start is not None and mark.end is not None:
                        start = clamp_point_to_rect(mark.start, page.rect)
                        end = clamp_point_to_rect(mark.end, page.rect)
                        self._draw_pdf_arrow(page, start, end)
                output_doc.save(str(output_path), garbage=4, deflate=True)
        except Exception as exc:
            error = exc
        self.save_results.put((output_path, error))

    def _poll_save_result(self) -> None:
        self.save_after_id = None
        try:
            output_path, error = self.save_results.get_nowait()
        except queue.Empty:
            if not self.shutting_down:
                self.save_after_id = self.root.after(50, self._poll_save_result)
            return

        self._set_busy(False)
        if error is None:
            self._update_status("status_complete", str(output_path))
            messagebox.showinfo(
                UI_TEXT["dialog_complete_title"],
                UI_TEXT["dialog_complete_message"],
                parent=self.root,
            )
            open_folder(output_path.parent)
        else:
            message = humanize_error(error)
            messagebox.showerror(UI_TEXT["dialog_save_error_title"], message, parent=self.root)
            self._update_status("status_error", message)

    def render_page(self) -> None:
        if self.pdf_path is None:
            self._show_empty_canvas()
            return
        self._submit_render_request(
            self.pdf_path,
            page_index=self.page_index,
            zoom=self.zoom,
            opens_document=False,
        )

    def _submit_render_request(
        self,
        pdf_path: Path,
        page_index: int,
        zoom: float,
        opens_document: bool,
    ) -> None:
        self.render_generation += 1
        request = RenderRequest(
            generation=self.render_generation,
            pdf_path=pdf_path,
            page_index=page_index,
            zoom=zoom,
            canvas_width=max(self.canvas.winfo_width(), 480),
            canvas_height=max(self.canvas.winfo_height(), 360),
            opens_document=opens_document,
        )
        self.render_requests_submitted += 1
        self.render_pending = True
        if opens_document:
            self._cancel_render_dispatch()
            if self.scheduled_render_request is not None:
                self.render_jobs_replaced += 1
                self.scheduled_render_request = None
            self._enqueue_render_request(request)
        else:
            if self.scheduled_render_request is not None:
                self.render_jobs_replaced += 1
            self.scheduled_render_request = request
            self._cancel_render_dispatch()
            self.render_dispatch_after_id = self.root.after(16, self._dispatch_scheduled_render)
        self._update_buttons()

    def _cancel_render_dispatch(self) -> None:
        if self.render_dispatch_after_id is None:
            return
        try:
            self.root.after_cancel(self.render_dispatch_after_id)
        except Exception:
            pass
        self.render_dispatch_after_id = None

    def _dispatch_scheduled_render(self) -> None:
        self.render_dispatch_after_id = None
        request = self.scheduled_render_request
        self.scheduled_render_request = None
        if request is not None:
            self._enqueue_render_request(request)

    def _enqueue_render_request(self, request: RenderRequest) -> None:
        with self.render_condition:
            if self.pending_render_request is not None:
                self.render_jobs_replaced += 1
            self.pending_render_request = request
            self.render_condition.notify()

    def _render_worker_loop(self) -> None:
        document = None
        document_path: Path | None = None
        try:
            while True:
                with self.render_condition:
                    while self.pending_render_request is None and not self.render_stop_requested:
                        self.render_condition.wait()
                    if self.render_stop_requested:
                        return
                    request = self.pending_render_request
                    self.pending_render_request = None
                    self.render_jobs_started += 1

                if request is None:
                    continue
                try:
                    fitz_module = get_fitz()
                    if document is None or document_path != request.pdf_path:
                        if document is not None:
                            document.close()
                        document = fitz_module.open(str(request.pdf_path))
                        document_path = request.pdf_path
                        self.render_documents_opened += 1
                    if document.page_count < 1:
                        raise ValueError(UI_TEXT["message_no_pages"])
                    if request.page_index < 0 or request.page_index >= document.page_count:
                        raise IndexError(UI_TEXT["message_no_pages"])

                    page = document.load_page(request.page_index)
                    page_rect = rect_values(page.rect)
                    page_width = max(page_rect[2] - page_rect[0], 1.0)
                    available_width = max(request.canvas_width - PAGE_MARGIN * 2, 240)
                    fit_scale = available_width / page_width
                    render_scale = min(max(fit_scale * request.zoom, MIN_RENDER_SCALE), MAX_RENDER_SCALE)
                    pix = page.get_pixmap(
                        matrix=fitz_module.Matrix(render_scale, render_scale),
                        alpha=False,
                    )
                    result = RenderResult(
                        request=request,
                        page_count=document.page_count,
                        page_rect=page_rect,
                        render_scale=render_scale,
                        image_width=pix.width,
                        image_height=pix.height,
                        png_data=pix.tobytes("png"),
                    )
                except Exception as exc:
                    result = RenderResult(request=request, error=exc)
                self.render_results.put(result)
        finally:
            if document is not None:
                try:
                    document.close()
                except Exception:
                    pass

    def _poll_render_results(self) -> None:
        self.render_after_id = None
        try:
            while True:
                result = self.render_results.get_nowait()
                if result.request.generation != self.render_generation:
                    self.render_results_discarded += 1
                    continue
                self._apply_render_result(result)
        except queue.Empty:
            pass
        if not self.render_stop_requested and not self.shutting_down:
            self.render_after_id = self.root.after(30, self._poll_render_results)

    def _apply_render_result(self, result: RenderResult) -> None:
        request = result.request
        self.render_pending = False
        if result.error is not None or result.png_data is None or result.page_rect is None:
            self.opening_pdf = False
            if isinstance(result.error, ModuleNotFoundError):
                message = UI_TEXT["message_pymupdf_missing"]
            else:
                message = humanize_error(result.error or RuntimeError(UI_TEXT["message_unknown_error"]))
            messagebox.showerror(UI_TEXT["dialog_open_error_title"], message, parent=self.root)
            self._update_status("status_error", message)
            if request.opens_document:
                self.pdf_path = None
                self.page_count = 0
                self.page_index = 0
                self.zoom = 1.0
                self.mode = None
                self.marks.clear()
                self._show_empty_canvas()
            self._update_buttons()
            return

        if request.opens_document:
            self.pdf_path = request.pdf_path
            self.page_count = result.page_count
            self.page_index = request.page_index
            self.zoom = request.zoom
            self.marks.clear()
            self._clear_mark_overlay()
            self.mode = "circle"
            self.opening_pdf = False

        self.page_rect = result.page_rect
        self.render_scale = result.render_scale
        self.image_width = result.image_width
        self.image_height = result.image_height
        encoded = base64.b64encode(result.png_data).decode("ascii")
        self.page_image = tk.PhotoImage(data=encoded, format="PNG")
        self._update_page_canvas(request.canvas_width, request.canvas_height)
        self.displayed_generation = request.generation
        self.displayed_pdf_path = request.pdf_path
        self.displayed_page_index = request.page_index
        self._sync_mark_overlay()
        if request.opens_document:
            self._update_status("status_circle", request.pdf_path.name)
        else:
            self._update_status("status_ready")
        self._update_buttons()

    def _update_page_canvas(self, canvas_width: int, canvas_height: int) -> None:
        visible_width = max(canvas_width, self.image_width + PAGE_MARGIN * 2)
        visible_height = max(canvas_height, self.image_height + PAGE_MARGIN * 2)
        self.image_x = max(PAGE_MARGIN, (visible_width - self.image_width) // 2)
        self.image_y = PAGE_MARGIN
        shadow_offset = 2
        shadow_coords = (
            self.image_x + shadow_offset,
            self.image_y + shadow_offset,
            self.image_x + self.image_width + shadow_offset,
            self.image_y + self.image_height + shadow_offset,
        )
        paper_coords = (
            self.image_x - 1,
            self.image_y - 1,
            self.image_x + self.image_width + 1,
            self.image_y + self.image_height + 1,
        )
        if self.page_shadow_id is None or self.page_paper_id is None or self.page_image_id is None:
            self.canvas.delete("all")
            self.mark_item_ids.clear()
            self.preview_id = None
            self.page_shadow_id = self.canvas.create_rectangle(
                *shadow_coords,
                fill=THEME["border"],
                outline="",
            )
            self.page_paper_id = self.canvas.create_rectangle(
                *paper_coords,
                fill=THEME["card"],
                outline=THEME["border"],
            )
            self.page_image_id = self.canvas.create_image(
                self.image_x,
                self.image_y,
                anchor="nw",
                image=self.page_image,
            )
        else:
            self.canvas.coords(self.page_shadow_id, *shadow_coords)
            self.canvas.coords(self.page_paper_id, *paper_coords)
            self.canvas.coords(self.page_image_id, self.image_x, self.image_y)
            self.canvas.itemconfigure(self.page_image_id, image=self.page_image)
        self.canvas.configure(scrollregion=(0, 0, visible_width, visible_height))

    def _show_empty_canvas(self) -> None:
        self.canvas.delete("all")
        self.page_shadow_id = None
        self.page_paper_id = None
        self.page_image_id = None
        self.mark_item_ids.clear()
        self.preview_id = None
        self.page_image = None
        self.page_rect = None
        self.image_width = 0
        self.image_height = 0
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

    def _clear_mark_overlay(self) -> None:
        self.canvas.delete("mark")
        self.mark_item_ids.clear()

    def _sync_mark_overlay(self) -> None:
        if self.pdf_path is None or self.page_rect is None:
            return
        visible_indices = {index for index, mark in enumerate(self.marks) if mark.page_index == self.page_index}
        for mark_index in tuple(self.mark_item_ids):
            if mark_index in visible_indices:
                continue
            self.canvas.delete(self.mark_item_ids.pop(mark_index))
        for mark_index in sorted(visible_indices):
            self._update_mark_item(mark_index)

    def _update_mark_item(self, mark_index: int) -> None:
        mark = self.marks[mark_index]
        if mark.page_index != self.page_index:
            return
        item_id = self.mark_item_ids.get(mark_index)
        if mark.kind == "circle" and mark.rect is not None:
            coords = self._page_rect_to_canvas(mark.rect)
            if item_id is None:
                self.mark_item_ids[mark_index] = self.canvas.create_oval(
                    *coords,
                    outline=RED_HEX,
                    width=DISPLAY_LINE_WIDTH,
                    tags=("mark",),
                )
            else:
                self.canvas.coords(item_id, *coords)
        elif mark.kind == "arrow" and mark.start is not None and mark.end is not None:
            start_x, start_y = self._page_point_to_canvas(mark.start)
            end_x, end_y = self._page_point_to_canvas(mark.end)
            coords = (start_x, start_y, end_x, end_y)
            if item_id is None:
                self.mark_item_ids[mark_index] = self.canvas.create_line(
                    *coords,
                    fill=RED_HEX,
                    width=DISPLAY_LINE_WIDTH,
                    arrow=tk.LAST,
                    arrowshape=(14, 18, 6),
                    capstyle=tk.ROUND,
                    tags=("mark",),
                )
            else:
                self.canvas.coords(item_id, *coords)

    def _page_point_to_canvas(self, point: tuple[float, float]) -> tuple[float, float]:
        if self.page_rect is None:
            return self.image_x, self.image_y
        x, y = point
        return (
            self.image_x + (x - self.page_rect[0]) * self.render_scale,
            self.image_y + (y - self.page_rect[1]) * self.render_scale,
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
            self.page_rect[0] + x / self.render_scale,
            self.page_rect[1] + y / self.render_scale,
        )

    def _on_press(self, event: tk.Event) -> None:
        if (
            self.pdf_path is None
            or self.busy
            or self.opening_pdf
            or self.render_pending
            or self.mode is None
            or self.displayed_page_index != self.page_index
        ):
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
        if self.mode == "circle":
            rect = self._circle_rect_from_points(self.drag_start_pdf, current)
            x0, y0, x1, y1 = self._page_rect_to_canvas(rect)
            if self.preview_id is None:
                self.preview_id = self.canvas.create_oval(
                    x0,
                    y0,
                    x1,
                    y1,
                    outline=RED_HEX,
                    width=DISPLAY_LINE_WIDTH,
                    dash=(6, 4),
                )
            else:
                self.canvas.coords(self.preview_id, x0, y0, x1, y1)
        elif self.mode == "arrow":
            start_x, start_y = self._page_point_to_canvas(self.drag_start_pdf)
            end_x, end_y = self._page_point_to_canvas(current)
            if self.preview_id is None:
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
            else:
                self.canvas.coords(self.preview_id, start_x, start_y, end_x, end_y)

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
        self._update_mark_item(len(self.marks) - 1)
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
        page: object,
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

        fitz_module = get_fitz()
        page.draw_line(
            fitz_module.Point(sx, sy),
            fitz_module.Point(ex, ey),
            color=RED_RGB,
            width=PDF_LINE_WIDTH,
            overlay=True,
        )
        angle = math.atan2(dy, dx)
        head_len = min(18.0, max(8.0, length * 0.25))
        for sign in (-1, 1):
            head_angle = angle + math.pi + sign * ARROW_HEAD_ANGLE
            hx = ex + head_len * math.cos(head_angle)
            hy = ey + head_len * math.sin(head_angle)
            page.draw_line(
                fitz_module.Point(ex, ey),
                fitz_module.Point(hx, hy),
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
        if self.pdf_path is None or self.busy or self.opening_pdf:
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
        if self.pdf_path is None:
            self._show_empty_canvas()
        else:
            self.render_page()

    def _cancel_after(self, attr_name: str) -> None:
        after_id = getattr(self, attr_name)
        if after_id is None:
            return
        try:
            self.root.after_cancel(after_id)
        except Exception:
            pass
        setattr(self, attr_name, None)

    def _stop_render_worker(self) -> None:
        self.render_stop_requested = True
        self.render_generation += 1
        self.scheduled_render_request = None
        with self.render_condition:
            self.pending_render_request = None
            self.render_condition.notify_all()
        if self.render_worker.is_alive() and self.render_worker is not threading.current_thread():
            self.render_worker.join(timeout=1.5)

    def _on_close(self) -> None:
        if self.shutting_down:
            return
        self.shutting_down = True
        for attr_name in (
            "resize_after_id",
            "render_dispatch_after_id",
            "render_after_id",
            "save_after_id",
        ):
            self._cancel_after(attr_name)
        self._stop_render_worker()
        if self.save_worker is not None and self.save_worker.is_alive():
            self.save_worker.join()
        self.root.destroy()


def main() -> None:
    set_windows_app_id()
    root = make_root()
    LookHereApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
