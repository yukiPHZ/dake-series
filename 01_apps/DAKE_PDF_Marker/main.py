# -*- coding: utf-8 -*-
from __future__ import annotations

import base64
import ctypes
import os
import queue
import subprocess
import sys
import threading
import urllib.parse
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, font as tkfont, messagebox
import tkinter as tk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore

    DND_ENABLED = True
except Exception:
    DND_FILES = None
    TkinterDnD = None
    DND_ENABLED = False


APP_NAME = "PDFマーカー"
WINDOW_TITLE = "PDFマーカー"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "brand_series": "シンプルそれDAKEシリーズ",
    "main_title": "PDFにマーカーを付ける",
    "main_description": "ピンクの半透明マーカーだけを付けます。",
    "button_add_pdf": "PDFを追加",
    "button_prev": "前へ",
    "button_next": "次へ",
    "button_save": "保存",
    "button_reset": "リセット",
    "status_idle": "未選択",
    "status_loading": "読み込み中",
    "status_ready": "準備完了",
    "status_processing": "処理中",
    "status_saving": "保存中",
    "status_complete": "保存完了",
    "status_error": "エラー",
    "empty_title": "PDFを追加してください",
    "empty_subtitle": "表示されたページ上をドラッグします。",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_tagline": "止まらない、迷わない、すぐ終わる。",
    "footer_link_1": "戸建買取査定",
    "footer_link_2": "Instagram",
    "footer_separator": " ｜ ",
    "footer_copyright": COPYRIGHT,
    "dialog_open_title": "PDFを選択",
    "dialog_error_title": "エラー",
    "dialog_complete_title": "保存完了",
    "dialog_pdf_filter_label": "PDFファイル",
    "dialog_all_filter_label": "すべてのファイル",
    "message_dependency_missing": "PyMuPDF が見つかりません。pymupdf をインストールしてください。",
    "message_invalid_file": "PDFファイルを選択してください。",
    "message_open_failed": "PDFを開けませんでした。",
    "message_save_failed": "保存できませんでした。",
    "message_save_complete": "PDFを保存しました。",
    "message_no_pdf": "先にPDFを追加してください。",
    "message_no_pages": "PDFのページを読み込めませんでした。",
    "message_no_folder_access": "保存先にアクセスできませんでした。",
    "message_unknown_error": "原因を特定できませんでした。",
    "status_marker_added": "マーカーを付けました",
    "status_reset": "マーカーを消しました",
    "status_drag": "ドラッグして範囲を選択してください",
    "status_page": "{current} / {total} ページ",
    "save_suffix": "_marker",
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
    "selection_border": "#7AA7FF",
    "disabled_bg": "#EEF2F7",
    "disabled_text": "#98A2B3",
    "success": "#12B76A",
    "success_bg": "#EAFBF3",
    "danger": "#D92D20",
    "danger_bg": "#FDECEC",
}

LINKS = {
    "footer_link_1": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "footer_link_2": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

FONT_CANDIDATES = ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo")
WINDOW_SIZE = "1040x740"
WINDOW_MIN_SIZE = (760, 620)
APP_USER_MODEL_ID = "Shimarisu.DakePDFMarker"
HEADER_NARROW_WIDTH = 820
FOOTER_NARROW_WIDTH = 900

PAGE_MARGIN = 24
MIN_RENDER_SCALE = 0.25
MAX_RENDER_SCALE = 5.0
ZOOM_MIN = 0.35
ZOOM_MAX = 4.0
ZOOM_STEP = 1.12
MIN_MARKER_DISPLAY_SIZE = 4
MARKER_RGB = (1.0, 0.29, 0.62)
MARKER_ALPHA = 0.35
MARKER_FILL_HEX = "#FF85BD"
MARKER_OUTLINE_HEX = "#F04493"

_FITZ_MODULE: object | None = None


def get_fitz() -> object:
    global _FITZ_MODULE
    if _FITZ_MODULE is None:
        import fitz as fitz_module  # PyInstallerの解析対象に保ったまま利用時だけ読み込む

        _FITZ_MODULE = fitz_module
    return _FITZ_MODULE


@dataclass(frozen=True)
class Marker:
    page_index: int
    rect: tuple[float, float, float, float]


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
        try:
            return TkinterDnD.Tk()
        except Exception:
            pass
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
        meipass = Path(getattr(sys, "_MEIPASS", exe_dir))
        return [
            exe_dir / ".." / ".." / ".." / "02_assets" / "dake_icon.ico",
            exe_dir / ".." / ".." / "02_assets" / "dake_icon.ico",
            meipass / "dake_icon.ico",
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


def parse_dropped_files(root: tk.Tk, raw_data: str) -> list[Path]:
    paths: list[Path] = []
    for raw_item in root.tk.splitlist(raw_data):
        value = raw_item.strip().strip("{}")
        if value.startswith("file:"):
            parsed = urllib.parse.urlparse(value)
            value = urllib.parse.unquote(parsed.path)
            if value.startswith("/") and len(value) > 3 and value[2] == ":":
                value = value[1:]
        if value:
            paths.append(Path(value))
    return paths


def is_pdf_path(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() == ".pdf"


def humanize_error(exc: Exception) -> str:
    detail = str(exc).strip().replace("\n", " ")
    return detail or UI_TEXT["message_unknown_error"]


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


def clamp_rect_to_page(
    raw_rect: tuple[float, float, float, float], page_rect: object
) -> tuple[float, float, float, float]:
    x0, y0, x1, y1 = normalize_rect(raw_rect)
    if isinstance(page_rect, tuple):
        page_x0, page_y0, page_x1, page_y1 = page_rect
    else:
        page_x0 = float(getattr(page_rect, "x0"))
        page_y0 = float(getattr(page_rect, "y0"))
        page_x1 = float(getattr(page_rect, "x1"))
        page_y1 = float(getattr(page_rect, "y1"))
    return (
        min(max(x0, page_x0), page_x1),
        min(max(y0, page_y0), page_y1),
        min(max(x1, page_x0), page_x1),
        min(max(y1, page_y0), page_y1),
    )


def marker_output_path(pdf_path: Path) -> Path:
    base_name = f"{pdf_path.stem}{UI_TEXT['save_suffix']}"
    candidate = pdf_path.with_name(f"{base_name}.pdf")
    index = 2
    while candidate.exists():
        candidate = pdf_path.with_name(f"{base_name}_{index}.pdf")
        index += 1
    return candidate


def draw_pdf_marker(page: object, marker: Marker) -> None:
    rect_values = clamp_rect_to_page(marker.rect, page.rect)
    x0, y0, x1, y1 = rect_values
    if x1 <= x0 or y1 <= y0:
        return
    fitz_module = get_fitz()
    page.draw_rect(
        fitz_module.Rect(rect_values),
        color=None,
        fill=MARKER_RGB,
        width=0,
        overlay=True,
        fill_opacity=MARKER_ALPHA,
        stroke_opacity=0,
    )


class PDFMarkerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.font_family = detect_font_family(root)
        self.pdf_path: Path | None = None
        self.page_count = 0
        self.page_index = 0
        self.page_rect: tuple[float, float, float, float] | None = None
        self.page_image: tk.PhotoImage | None = None
        self.image_x = PAGE_MARGIN
        self.image_y = PAGE_MARGIN
        self.image_width = 0
        self.image_height = 0
        self.render_scale = 1.0
        self.zoom = 1.0
        self.markers: list[Marker] = []
        self.marker_item_ids: dict[int, int] = {}
        self.busy = False
        self.render_pending = False
        self.opening_pdf = False
        self.drag_start_pdf: tuple[float, float] | None = None
        self.preview_id: int | None = None
        self.page_shadow_id: int | None = None
        self.page_paper_id: int | None = None
        self.page_image_id: int | None = None
        self.resize_after_id: str | None = None
        self.worker_after_id: str | None = None
        self.render_after_id: str | None = None
        self.render_dispatch_after_id: str | None = None
        self.header_mode: str | None = None
        self.footer_mode: str | None = None
        self.worker_results: "queue.Queue[tuple[str, Path | None, Exception | None]]" = queue.Queue()
        self.render_results: "queue.Queue[RenderResult]" = queue.Queue()
        self.render_condition = threading.Condition()
        self.pending_render_request: RenderRequest | None = None
        self.scheduled_render_request: RenderRequest | None = None
        self.render_generation = 0
        self.displayed_generation = 0
        self.displayed_pdf_path: Path | None = None
        self.displayed_page_index = -1
        self.render_stop_requested = False
        self.render_requests_submitted = 0
        self.render_jobs_started = 0
        self.render_jobs_replaced = 0
        self.render_results_discarded = 0
        self.render_documents_opened = 0
        self.render_worker = threading.Thread(target=self._render_worker_loop, daemon=True)

        self.status_var = tk.StringVar(value=UI_TEXT["status_idle"])
        self.status_detail_var = tk.StringVar(value=UI_TEXT["status_idle"])
        self.page_var = tk.StringVar(value=UI_TEXT["status_page"].format(current=0, total=0))

        self._configure_root()
        self._build_ui()
        self._setup_bindings()
        self._show_empty_canvas()
        self._update_controls()
        self.render_worker.start()
        self.render_after_id = self.root.after(30, self._poll_render_results)

    def _configure_root(self) -> None:
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(*WINDOW_MIN_SIZE)
        self.root.configure(bg=THEME["background"])
        apply_window_icon(self.root)

    def _build_ui(self) -> None:
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=1)

        self.title_font = (self.font_family, 20, "bold")
        self.subtitle_font = (self.font_family, 10)
        self.button_font = (self.font_family, 10, "bold")
        self.body_font = (self.font_family, 10)
        self.small_font = (self.font_family, 9)
        self.status_font = (self.font_family, 9, "bold")

        self.header = tk.Frame(self.root, bg=THEME["background"])
        self.header.grid(row=0, column=0, sticky="ew", padx=22, pady=(16, 10))
        self.header.grid_columnconfigure(0, weight=0)
        self.header.grid_columnconfigure(1, weight=1)

        self.header_title_label = tk.Label(
            self.header,
            text=UI_TEXT["main_title"],
            bg=THEME["background"],
            fg=THEME["text"],
            font=self.title_font,
            anchor="w",
        )
        self.header_title_label.grid(row=0, column=0, sticky="w")
        self.header_description_label = tk.Label(
            self.header,
            text=UI_TEXT["main_description"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=self.subtitle_font,
            anchor="w",
        )
        self.header_description_label.grid(row=0, column=1, sticky="w", padx=(16, 0), pady=(4, 0))

        toolbar = tk.Frame(self.header, bg=THEME["card"], highlightthickness=1, highlightbackground=THEME["border"])
        toolbar.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        toolbar.grid_columnconfigure(8, weight=1)

        self.add_button = self._make_button(toolbar, "button_add_pdf", self.select_pdf, primary=True)
        self.prev_button = self._make_button(toolbar, "button_prev", self.prev_page)
        self.next_button = self._make_button(toolbar, "button_next", self.next_page)
        self.reset_button = self._make_button(toolbar, "button_reset", self.reset_markers)
        self.save_button = self._make_button(toolbar, "button_save", self.save_pdf, primary=True)

        self.add_button.grid(row=0, column=0, padx=(12, 8), pady=10, sticky="w")
        self.prev_button.grid(row=0, column=1, padx=(0, 6), pady=10, sticky="w")
        self.page_label = tk.Label(
            toolbar,
            textvariable=self.page_var,
            bg=THEME["card"],
            fg=THEME["text"],
            font=(self.font_family, 10, "bold"),
            width=13,
            anchor="center",
        )
        self.page_label.grid(row=0, column=2, padx=(0, 6), pady=10, sticky="w")
        self.next_button.grid(row=0, column=3, padx=(0, 12), pady=10, sticky="w")

        divider = tk.Frame(toolbar, width=1, bg=THEME["border"])
        divider.grid(row=0, column=4, sticky="ns", pady=12)

        self.reset_button.grid(row=0, column=5, padx=(12, 8), pady=10, sticky="w")
        self.save_button.grid(row=0, column=6, padx=(0, 12), pady=10, sticky="w")

        viewer = tk.Frame(self.root, bg=THEME["card"], highlightthickness=1, highlightbackground=THEME["border"])
        viewer.grid(row=1, column=0, sticky="nsew", padx=24, pady=(0, 10))
        viewer.grid_columnconfigure(0, weight=1)
        viewer.grid_rowconfigure(0, weight=1)

        self.canvas = tk.Canvas(
            viewer,
            bg=THEME["card"],
            bd=0,
            highlightthickness=0,
            xscrollincrement=12,
            yscrollincrement=12,
        )
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.v_scroll = tk.Scrollbar(viewer, orient="vertical", command=self.canvas.yview)
        self.h_scroll = tk.Scrollbar(viewer, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=self.v_scroll.set, xscrollcommand=self.h_scroll.set)
        self.v_scroll.grid(row=0, column=1, sticky="ns")
        self.h_scroll.grid(row=1, column=0, sticky="ew")

        bottom = tk.Frame(self.root, bg=THEME["background"])
        bottom.grid(row=2, column=0, sticky="ew", padx=24, pady=(0, 14))
        bottom.grid_columnconfigure(0, weight=1)

        status_row = tk.Frame(bottom, bg=THEME["background"])
        status_row.grid(row=0, column=0, sticky="ew")
        status_row.grid_columnconfigure(1, weight=1)

        self.status_badge = tk.Label(
            status_row,
            textvariable=self.status_var,
            bg=THEME["disabled_bg"],
            fg=THEME["muted"],
            font=self.status_font,
            padx=12,
            pady=4,
        )
        self.status_badge.grid(row=0, column=0, sticky="w")
        tk.Label(
            status_row,
            textvariable=self.status_detail_var,
            bg=THEME["background"],
            fg=THEME["muted"],
            font=self.small_font,
            anchor="w",
        ).grid(row=0, column=1, sticky="w", padx=(10, 0))

        self.footer = tk.Frame(bottom, bg=THEME["background"])
        self.footer.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        self.footer.grid_columnconfigure(0, weight=0)
        self.footer.grid_columnconfigure(1, weight=1)
        self.footer.grid_columnconfigure(2, weight=0)

        footer_left_text = UI_TEXT["footer_left"] + UI_TEXT["footer_separator"] + UI_TEXT["footer_tagline"]
        self.footer_left_label = tk.Label(
            self.footer,
            text=footer_left_text,
            bg=THEME["background"],
            fg=THEME["muted"],
            font=self.small_font,
            justify="left",
            anchor="w",
        )
        self.footer_left_label.grid(row=0, column=0, sticky="w")

        self.right_footer = tk.Frame(self.footer, bg=THEME["background"])
        self.right_footer.grid(row=0, column=2, sticky="e")
        self._make_footer_link(self.right_footer, "footer_link_1").grid(row=0, column=0, sticky="e")
        self._make_footer_separator(self.right_footer).grid(row=0, column=1, sticky="e")
        self._make_footer_link(self.right_footer, "footer_link_2").grid(row=0, column=2, sticky="e")
        self._make_footer_separator(self.right_footer).grid(row=0, column=3, sticky="e")
        self.footer_copyright_label = tk.Label(
            self.right_footer,
            text=UI_TEXT["footer_copyright"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=self.small_font,
            anchor="e",
        )
        self.footer_copyright_label.grid(row=0, column=4, sticky="e")
        self._update_responsive_layout()

    def _make_button(self, parent: tk.Misc, text_key: str, command: object, primary: bool = False) -> tk.Button:
        bg = THEME["accent"] if primary else THEME["card"]
        fg = THEME["card"] if primary else THEME["text"]
        active_bg = THEME["accent_hover"] if primary else THEME["selection_bg"]
        return tk.Button(
            parent,
            text=UI_TEXT[text_key],
            command=command,
            bg=bg,
            fg=fg,
            activebackground=active_bg,
            activeforeground=fg,
            disabledforeground=THEME["disabled_text"],
            relief="flat",
            bd=0,
            padx=16,
            pady=8,
            cursor="hand2",
            font=self.button_font,
            highlightthickness=1,
            highlightbackground=THEME["border"],
            highlightcolor=THEME["selection_border"],
        )

    def _make_footer_link(self, parent: tk.Misc, key: str) -> tk.Label:
        label = tk.Label(
            parent,
            text=UI_TEXT[key],
            bg=THEME["background"],
            fg=THEME["accent"],
            font=self.small_font,
            cursor="hand2",
        )
        label.bind("<Button-1>", lambda _event, link_key=key: self._open_link(link_key))
        label.bind("<Enter>", lambda _event, item=label: item.configure(fg=THEME["accent_hover"]))
        label.bind("<Leave>", lambda _event, item=label: item.configure(fg=THEME["accent"]))
        return label

    def _make_footer_separator(self, parent: tk.Misc) -> tk.Label:
        return tk.Label(
            parent,
            text=UI_TEXT["footer_separator"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=self.small_font,
        )

    def _update_responsive_layout(self, width: int | None = None) -> None:
        current_width = width or max(self.root.winfo_width(), WINDOW_MIN_SIZE[0])

        next_header_mode = "narrow" if current_width < HEADER_NARROW_WIDTH else "wide"
        if next_header_mode != self.header_mode:
            self.header_mode = next_header_mode
            if next_header_mode == "narrow":
                self.header_title_label.grid_configure(row=0, column=0, columnspan=2, sticky="w")
                self.header_description_label.grid_configure(
                    row=1,
                    column=0,
                    columnspan=2,
                    sticky="w",
                    padx=(0, 0),
                    pady=(3, 0),
                )
            else:
                self.header_title_label.grid_configure(row=0, column=0, columnspan=1, sticky="w")
                self.header_description_label.grid_configure(
                    row=0,
                    column=1,
                    columnspan=1,
                    sticky="w",
                    padx=(16, 0),
                    pady=(4, 0),
                )

        next_footer_mode = "narrow" if current_width < FOOTER_NARROW_WIDTH else "wide"
        if next_footer_mode == self.footer_mode:
            return

        self.footer_mode = next_footer_mode
        if next_footer_mode == "narrow":
            self.footer.grid_columnconfigure(0, weight=1)
            self.footer.grid_columnconfigure(1, weight=0)
            self.footer.grid_columnconfigure(2, weight=1)
            self.footer_left_label.configure(anchor="center", justify="center")
            self.footer_left_label.grid_configure(row=0, column=0, columnspan=3, sticky="")
            self.right_footer.grid_configure(row=1, column=0, columnspan=3, sticky="", pady=(2, 0))
        else:
            self.footer.grid_columnconfigure(0, weight=0)
            self.footer.grid_columnconfigure(1, weight=1)
            self.footer.grid_columnconfigure(2, weight=0)
            self.footer_left_label.configure(anchor="w", justify="left")
            self.footer_left_label.grid_configure(row=0, column=0, columnspan=1, sticky="w")
            self.right_footer.grid_configure(row=0, column=2, columnspan=1, sticky="e", pady=(0, 0))

    def _setup_bindings(self) -> None:
        self.canvas.bind("<ButtonPress-1>", self._on_press)
        self.canvas.bind("<B1-Motion>", self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind("<Control-MouseWheel>", self._on_ctrl_mousewheel)
        self.root.bind("<Control-MouseWheel>", self._on_ctrl_mousewheel)
        self.canvas.bind("<Control-Button-4>", lambda _event: self._zoom_at_mouse(1))
        self.canvas.bind("<Control-Button-5>", lambda _event: self._zoom_at_mouse(-1))
        self.root.bind("<Configure>", self._on_root_configure, add="+")
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        if DND_ENABLED and DND_FILES is not None:
            for widget in (self.root, self.canvas):
                try:
                    widget.drop_target_register(DND_FILES)
                    widget.dnd_bind("<<Drop>>", self._on_drop)
                except Exception:
                    pass

    def _open_link(self, key: str) -> None:
        url = LINKS.get(key)
        if not url:
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(url)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", url])
            else:
                subprocess.Popen(["xdg-open", url])
        except Exception:
            pass

    def select_pdf(self) -> None:
        if self.busy:
            return
        filename = filedialog.askopenfilename(
            parent=self.root,
            title=UI_TEXT["dialog_open_title"],
            filetypes=[
                (UI_TEXT["dialog_pdf_filter_label"], "*.pdf"),
                (UI_TEXT["dialog_all_filter_label"], "*.*"),
            ],
        )
        if filename:
            self.open_pdf(Path(filename))

    def open_pdf(self, path: Path) -> None:
        if self.busy:
            return
        if not is_pdf_path(path):
            self._show_error(UI_TEXT["message_invalid_file"])
            return

        self._set_status("status_loading", path.name)
        self.opening_pdf = True
        self._submit_render_request(path, page_index=0, zoom=1.0, opens_document=True)

    def save_pdf(self) -> None:
        if self.busy:
            return
        if self.pdf_path is None or self.page_count < 1:
            self._show_error(UI_TEXT["message_no_pdf"])
            return

        input_path = self.pdf_path
        output_path = marker_output_path(input_path)
        markers = list(self.markers)

        self._set_busy(True)
        self._set_status("status_saving", output_path.name)
        worker = threading.Thread(
            target=self._save_worker,
            args=(input_path, output_path, markers),
            daemon=True,
        )
        worker.start()
        self._poll_worker_results()

    def _save_worker(self, input_path: Path, output_path: Path, markers: list[Marker]) -> None:
        error: Exception | None = None
        try:
            fitz_module = get_fitz()
            with fitz_module.open(str(input_path)) as output_doc:
                for marker in markers:
                    if marker.page_index < 0 or marker.page_index >= output_doc.page_count:
                        continue
                    page = output_doc.load_page(marker.page_index)
                    draw_pdf_marker(page, marker)
                output_doc.save(str(output_path), garbage=4, deflate=True)
        except Exception as exc:
            error = exc
        self.worker_results.put(("save", output_path, error))

    def _poll_worker_results(self) -> None:
        try:
            kind, path, error = self.worker_results.get_nowait()
        except queue.Empty:
            self.worker_after_id = self.root.after(80, self._poll_worker_results)
            return

        self.worker_after_id = None
        if kind == "save":
            self._set_busy(False)
            if error is None and path is not None:
                self._set_status("status_complete", str(path))
                messagebox.showinfo(
                    UI_TEXT["dialog_complete_title"],
                    UI_TEXT["message_save_complete"],
                    parent=self.root,
                )
                open_folder(path.parent)
            else:
                detail = humanize_error(error) if error is not None else UI_TEXT["message_save_failed"]
                message = f"{UI_TEXT['message_save_failed']}\n\n{detail}"
                messagebox.showerror(UI_TEXT["dialog_error_title"], message, parent=self.root)
                self._set_status("status_error", UI_TEXT["message_save_failed"])
            self._update_controls()

    def reset_markers(self) -> None:
        if self.busy:
            return
        self.markers.clear()
        self.drag_start_pdf = None
        self._clear_preview()
        self._clear_marker_overlay()
        if self.pdf_path is None:
            self._show_empty_canvas()
        self._set_status("status_reset")
        self._update_controls()

    def prev_page(self) -> None:
        if self.pdf_path is None or self.busy or self.opening_pdf or self.page_index <= 0:
            return
        self.page_index -= 1
        self.render_page()
        self._set_status("status_loading", self.page_var.get())
        self._update_controls()

    def next_page(self) -> None:
        if self.pdf_path is None or self.busy or self.opening_pdf:
            return
        if self.page_index >= self.page_count - 1:
            return
        self.page_index += 1
        self.render_page()
        self._set_status("status_loading", self.page_var.get())
        self._update_controls()

    def render_page(self) -> None:
        if self.pdf_path is None:
            self._show_empty_canvas()
            return
        self._submit_render_request(self.pdf_path, self.page_index, self.zoom, opens_document=False)

    def _submit_render_request(
        self,
        pdf_path: Path,
        page_index: int,
        zoom: float,
        opens_document: bool,
    ) -> None:
        canvas_width = max(self.canvas.winfo_width(), 520)
        canvas_height = max(self.canvas.winfo_height(), 360)
        self.render_generation += 1
        request = RenderRequest(
            generation=self.render_generation,
            pdf_path=pdf_path,
            page_index=page_index,
            zoom=zoom,
            canvas_width=canvas_width,
            canvas_height=canvas_height,
            opens_document=opens_document,
        )
        self.render_requests_submitted += 1
        self.render_pending = True
        if opens_document:
            if self.render_dispatch_after_id is not None:
                try:
                    self.root.after_cancel(self.render_dispatch_after_id)
                except Exception:
                    pass
                self.render_dispatch_after_id = None
            if self.scheduled_render_request is not None:
                self.render_jobs_replaced += 1
                self.scheduled_render_request = None
            self._enqueue_render_request(request)
        else:
            if self.scheduled_render_request is not None:
                self.render_jobs_replaced += 1
            self.scheduled_render_request = request
            if self.render_dispatch_after_id is not None:
                try:
                    self.root.after_cancel(self.render_dispatch_after_id)
                except Exception:
                    pass
            self.render_dispatch_after_id = self.root.after(16, self._dispatch_scheduled_render)
        self._update_controls()

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
                        document = None
                        document_path = None
                        document = fitz_module.open(str(request.pdf_path))
                        document_path = request.pdf_path
                        self.render_documents_opened += 1
                    if document.page_count < 1:
                        raise ValueError(UI_TEXT["message_no_pages"])
                    if request.page_index < 0 or request.page_index >= document.page_count:
                        raise IndexError(UI_TEXT["message_no_pages"])

                    page = document.load_page(request.page_index)
                    page_rect = (
                        float(page.rect.x0),
                        float(page.rect.y0),
                        float(page.rect.x1),
                        float(page.rect.y1),
                    )
                    page_width = max(page_rect[2] - page_rect[0], 1.0)
                    page_height = max(page_rect[3] - page_rect[1], 1.0)
                    available_width = max(request.canvas_width - PAGE_MARGIN * 2, 240)
                    available_height = max(request.canvas_height - PAGE_MARGIN * 2, 240)
                    fit_scale = min(available_width / page_width, available_height / page_height)
                    render_scale = min(max(fit_scale * request.zoom, MIN_RENDER_SCALE), MAX_RENDER_SCALE)
                    pix = page.get_pixmap(matrix=fitz_module.Matrix(render_scale, render_scale), alpha=False)
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
        if not self.render_stop_requested:
            self.render_after_id = self.root.after(30, self._poll_render_results)

    def _apply_render_result(self, result: RenderResult) -> None:
        request = result.request
        self.render_pending = False
        if result.error is not None or result.png_data is None or result.page_rect is None:
            self.opening_pdf = False
            if isinstance(result.error, ModuleNotFoundError):
                message = UI_TEXT["message_dependency_missing"]
            else:
                detail = humanize_error(result.error) if result.error is not None else UI_TEXT["message_open_failed"]
                message = f"{UI_TEXT['message_open_failed']}\n\n{detail}"
            messagebox.showerror(UI_TEXT["dialog_error_title"], message, parent=self.root)
            self._set_status("status_error", UI_TEXT["message_open_failed"])
            self._update_controls()
            return

        if request.opens_document:
            self.pdf_path = request.pdf_path
            self.page_count = result.page_count
            self.page_index = request.page_index
            self.zoom = request.zoom
            self.markers.clear()
            self._clear_marker_overlay()
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
        self._sync_marker_overlay()
        self._set_status("status_ready", UI_TEXT["status_drag"])
        self._update_page_label()
        self._update_controls()

    def _update_page_canvas(self, canvas_width: int, canvas_height: int) -> None:
        content_width = max(canvas_width, self.image_width + PAGE_MARGIN * 2)
        content_height = max(canvas_height, self.image_height + PAGE_MARGIN * 2)
        self.image_x = max(PAGE_MARGIN, (content_width - self.image_width) // 2)
        self.image_y = PAGE_MARGIN

        shadow_coords = (
            self.image_x + 3,
            self.image_y + 3,
            self.image_x + self.image_width + 3,
            self.image_y + self.image_height + 3,
        )
        paper_coords = (
            self.image_x - 1,
            self.image_y - 1,
            self.image_x + self.image_width + 1,
            self.image_y + self.image_height + 1,
        )
        if self.page_shadow_id is None or self.page_paper_id is None or self.page_image_id is None:
            self.canvas.delete("all")
            self.marker_item_ids.clear()
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
                image=self.page_image,
                anchor="nw",
            )
        else:
            self.canvas.coords(self.page_shadow_id, *shadow_coords)
            self.canvas.coords(self.page_paper_id, *paper_coords)
            self.canvas.coords(self.page_image_id, self.image_x, self.image_y)
            self.canvas.itemconfigure(self.page_image_id, image=self.page_image)
        self.canvas.configure(scrollregion=(0, 0, content_width, content_height))

    def _show_empty_canvas(self) -> None:
        self.canvas.delete("all")
        self.page_shadow_id = None
        self.page_paper_id = None
        self.page_image_id = None
        self.marker_item_ids.clear()
        self.preview_id = None
        width = max(self.canvas.winfo_width(), 600)
        height = max(self.canvas.winfo_height(), 360)
        self.canvas.configure(scrollregion=(0, 0, width, height))
        self.canvas.create_rectangle(
            24,
            24,
            width - 24,
            height - 24,
            fill=THEME["card"],
            outline=THEME["border"],
            width=1,
            dash=(6, 5),
        )
        self.canvas.create_text(
            width / 2,
            height / 2 - 20,
            text=UI_TEXT["empty_title"],
            fill=THEME["text"],
            font=(self.font_family, 18, "bold"),
        )
        self.canvas.create_text(
            width / 2,
            height / 2 + 18,
            text=UI_TEXT["empty_subtitle"],
            fill=THEME["muted"],
            font=self.subtitle_font,
        )
        self._update_page_label()

    def _clear_marker_overlay(self) -> None:
        self.canvas.delete("marker")
        self.marker_item_ids.clear()

    def _redraw_marker_overlay(self) -> None:
        self._sync_marker_overlay()

    def _sync_marker_overlay(self) -> None:
        if self.pdf_path is None or self.page_rect is None:
            return
        visible_indices = {
            index for index, marker in enumerate(self.markers) if marker.page_index == self.page_index
        }
        for marker_index in tuple(self.marker_item_ids):
            if marker_index in visible_indices:
                continue
            self.canvas.delete(self.marker_item_ids.pop(marker_index))

        for marker_index in sorted(visible_indices):
            self._update_marker_item(marker_index)

    def _update_marker_item(self, marker_index: int) -> None:
        marker = self.markers[marker_index]
        if marker.page_index != self.page_index:
            return
        x0, y0, x1, y1 = self._page_rect_to_canvas(marker.rect)
        item_id = self.marker_item_ids.get(marker_index)
        if item_id is None:
            self.marker_item_ids[marker_index] = self.canvas.create_rectangle(
                x0,
                y0,
                x1,
                y1,
                fill=MARKER_FILL_HEX,
                outline=MARKER_OUTLINE_HEX,
                width=1,
                stipple="gray50",
                tags=("marker",),
            )
        else:
            self.canvas.coords(item_id, x0, y0, x1, y1)

    def _event_to_page_point(self, event: tk.Event, require_inside: bool) -> tuple[float, float] | None:
        if self.pdf_path is None or self.page_rect is None or self.image_width <= 0 or self.image_height <= 0:
            return None
        local_x = self.canvas.canvasx(event.x) - self.image_x
        local_y = self.canvas.canvasy(event.y) - self.image_y
        inside = 0 <= local_x <= self.image_width and 0 <= local_y <= self.image_height
        if require_inside and not inside:
            return None
        local_x = min(max(local_x, 0), self.image_width)
        local_y = min(max(local_y, 0), self.image_height)
        page_rect = self.page_rect
        return (
            page_rect[0] + local_x / self.render_scale,
            page_rect[1] + local_y / self.render_scale,
        )

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

    def _on_press(self, event: tk.Event) -> str:
        if self.pdf_path is None or self.busy or self.render_pending or self.opening_pdf:
            return "break"
        point = self._event_to_page_point(event, require_inside=True)
        if point is None:
            return "break"
        self.drag_start_pdf = point
        self._clear_preview()
        self._set_status("status_processing", UI_TEXT["status_drag"])
        return "break"

    def _on_drag(self, event: tk.Event) -> str:
        if self.drag_start_pdf is None:
            return "break"
        current = self._event_to_page_point(event, require_inside=False)
        if current is None:
            return "break"
        raw_rect = (self.drag_start_pdf[0], self.drag_start_pdf[1], current[0], current[1])
        x0, y0, x1, y1 = self._page_rect_to_canvas(normalize_rect(raw_rect))
        if self.preview_id is None:
            self.preview_id = self.canvas.create_rectangle(
                x0,
                y0,
                x1,
                y1,
                fill=MARKER_FILL_HEX,
                outline=MARKER_OUTLINE_HEX,
                width=1,
                stipple="gray50",
                dash=(5, 3),
            )
        else:
            self.canvas.coords(self.preview_id, x0, y0, x1, y1)
        return "break"

    def _on_release(self, event: tk.Event) -> str:
        if self.drag_start_pdf is None:
            return "break"
        end_point = self._event_to_page_point(event, require_inside=False)
        start_point = self.drag_start_pdf
        self.drag_start_pdf = None
        self._clear_preview()
        if end_point is None or self.page_rect is None:
            return "break"

        raw_rect = (start_point[0], start_point[1], end_point[0], end_point[1])
        canvas_rect = self._page_rect_to_canvas(normalize_rect(raw_rect))
        if (
            abs(canvas_rect[2] - canvas_rect[0]) < MIN_MARKER_DISPLAY_SIZE
            or abs(canvas_rect[3] - canvas_rect[1]) < MIN_MARKER_DISPLAY_SIZE
        ):
            self._set_status("status_ready", UI_TEXT["status_drag"])
            return "break"

        rect = clamp_rect_to_page(raw_rect, self.page_rect)
        self.markers.append(Marker(page_index=self.page_index, rect=rect))
        self._update_marker_item(len(self.markers) - 1)
        self._set_status("status_marker_added", UI_TEXT["status_drag"])
        self._update_controls()
        return "break"

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

    def _zoom_at_mouse(self, direction: int) -> str:
        if self.pdf_path is None or self.busy or self.opening_pdf:
            return "break"
        if direction > 0:
            self.zoom = min(self.zoom * ZOOM_STEP, ZOOM_MAX)
        else:
            self.zoom = max(self.zoom / ZOOM_STEP, ZOOM_MIN)
        self.render_page()
        return "break"

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

    def _on_root_configure(self, event: tk.Event) -> None:
        if event.widget is self.root:
            self._update_responsive_layout(getattr(event, "width", None))

    def _on_drop(self, event: tk.Event) -> None:
        paths = parse_dropped_files(self.root, getattr(event, "data", ""))
        if not paths:
            return
        self.open_pdf(paths[0])

    def _set_status(self, key: str, detail: str | None = None) -> None:
        text = UI_TEXT.get(key, key)
        self.status_var.set(text)
        self.status_detail_var.set(detail or text)
        if key == "status_complete":
            bg, fg = THEME["success_bg"], THEME["success"]
        elif key == "status_error":
            bg, fg = THEME["danger_bg"], THEME["danger"]
        elif key in {"status_loading", "status_processing", "status_saving", "status_marker_added"}:
            bg, fg = THEME["selection_bg"], THEME["accent"]
        else:
            bg, fg = THEME["disabled_bg"], THEME["muted"]
        self.status_badge.configure(bg=bg, fg=fg)

    def _show_error(self, message: str) -> None:
        messagebox.showerror(UI_TEXT["dialog_error_title"], message, parent=self.root)
        self._set_status("status_error", message)

    def _set_busy(self, busy: bool) -> None:
        self.busy = busy
        self.canvas.configure(cursor="watch" if busy else "")
        self._update_controls()
        self.root.update_idletasks()

    def _update_page_label(self) -> None:
        if self.pdf_path is None or self.page_count < 1:
            self.page_var.set(UI_TEXT["status_page"].format(current=0, total=0))
            return
        self.page_var.set(UI_TEXT["status_page"].format(current=self.page_index + 1, total=self.page_count))

    def _update_controls(self) -> None:
        has_pdf = self.pdf_path is not None and self.page_count > 0
        stable_preview = has_pdf and not self.render_pending and not self.opening_pdf
        self.add_button.configure(state=tk.DISABLED if self.busy else tk.NORMAL)
        self.save_button.configure(state=tk.NORMAL if stable_preview and not self.busy else tk.DISABLED)
        self.reset_button.configure(state=tk.NORMAL if self.markers and stable_preview and not self.busy else tk.DISABLED)
        self.prev_button.configure(
            state=tk.NORMAL if has_pdf and not self.busy and not self.opening_pdf and self.page_index > 0 else tk.DISABLED
        )
        self.next_button.configure(
            state=(
                tk.NORMAL
                if has_pdf and not self.busy and not self.opening_pdf and self.page_index < self.page_count - 1
                else tk.DISABLED
            )
        )
        self._update_page_label()

    def _close_document(self) -> None:
        self.pdf_path = None
        self.page_count = 0
        self.page_rect = None
        self.page_image = None

    def _stop_render_worker(self) -> None:
        if self.render_dispatch_after_id is not None:
            try:
                self.root.after_cancel(self.render_dispatch_after_id)
            except Exception:
                pass
            self.render_dispatch_after_id = None
        self.scheduled_render_request = None
        if self.render_after_id is not None:
            try:
                self.root.after_cancel(self.render_after_id)
            except Exception:
                pass
            self.render_after_id = None
        with self.render_condition:
            self.render_stop_requested = True
            self.pending_render_request = None
            self.render_condition.notify_all()
        if self.render_worker.is_alive() and self.render_worker is not threading.current_thread():
            self.render_worker.join(timeout=1.5)

    def _on_close(self) -> None:
        if self.resize_after_id is not None:
            try:
                self.root.after_cancel(self.resize_after_id)
            except Exception:
                pass
        if self.worker_after_id is not None:
            try:
                self.root.after_cancel(self.worker_after_id)
            except Exception:
                pass
        self._stop_render_worker()
        self._close_document()
        self.root.destroy()

    def run(self) -> None:
        self.root.mainloop()
        self._stop_render_worker()
        self._close_document()


def main() -> None:
    set_windows_app_id()
    root = make_root()
    app = PDFMarkerApp(root)
    app.run()


if __name__ == "__main__":
    main()
