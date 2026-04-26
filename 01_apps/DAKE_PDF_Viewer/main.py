# -*- coding: utf-8 -*-
from __future__ import annotations

import ctypes
import os
import queue
import shutil
import sys
import tempfile
import threading
import time
import tkinter as tk
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, font as tkfont, messagebox, ttk

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None  # type: ignore[assignment]

try:
    from PIL import Image, ImageDraw, ImageTk
except Exception:
    Image = None  # type: ignore[assignment]
    ImageDraw = None  # type: ignore[assignment]
    ImageTk = None  # type: ignore[assignment]

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore

    DND_ENABLED = True
except Exception:
    DND_FILES = None
    TkinterDnD = None
    DND_ENABLED = False


APP_NAME = "DakePDF見る"
WINDOW_TITLE = "DakePDF見る"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "brand_series": "シンプルそれDAKEシリーズ",
    "main_title": "PDFを見る",
    "main_description": "PDFを選んで、すぐ確認します。",
    "button_select_folder": "フォルダを選ぶ",
    "button_add_pdf": "PDFを追加",
    "button_clear": "一覧をクリア",
    "file_list_title": "PDF一覧",
    "column_file": "ファイル名",
    "column_pages": "ページ",
    "column_modified": "更新日時",
    "viewer_empty_title": "PDFを選んでください",
    "viewer_empty_subtitle": "フォルダを選ぶか、ドラッグ＆ドロップ",
    "viewer_empty_hint": "Ctrl + ホイールで拡大・縮小",
    "search_placeholder": "検索",
    "search_not_found": "見つかりませんでした",
    "search_not_available": "このPDFは文字検索できません",
    "status_idle": "未選択",
    "status_loading": "読み込み中",
    "status_ready": "表示できます",
    "status_viewing": "表示中",
    "status_zoom": "倍率",
    "status_rotate": "表示を回転しました",
    "status_reset_rotate": "回転をリセットしました",
    "status_print_current": "現在ページを印刷します",
    "status_error": "エラー",
    "status_page": "{current} / {total} ページ",
    "error_open_pdf": "PDFを開けませんでした。",
    "error_no_pdf": "PDFが見つかりませんでした。",
    "error_invalid_file": "PDFファイルを選んでください。",
    "error_print": "印刷できませんでした。",
    "error_dependency": "PyMuPDF と Pillow が必要です。",
    "dialog_folder_title": "フォルダを選択",
    "dialog_pdf_title": "PDFを追加",
    "dialog_pdf_filter": "PDFファイル",
    "dialog_error_title": "エラー",
    "page_count_unknown": "-",
    "page_count_format": "{count}",
    "modified_unknown": "-",
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
    "selection_border": "#7AA7FF",
    "soft": "#EEF2F7",
    "error": "#D92D20",
    "error_bg": "#FDECEC",
}

FONT_CANDIDATES = ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo")
WINDOW_SIZE = "1120x760"
WINDOW_MIN_SIZE = (900, 620)
APP_USER_MODEL_ID = "Shimarisu.DakePDFViewer"
ZOOM_MIN = 0.5
ZOOM_MAX = 3.0
ZOOM_STEP = 1.12
FIT_PADDING = 56
RENDER_DPI_SCALE = 2.0
PAGE_GAP = 18
POLL_INTERVAL_MS = 60
ZOOM_DEBOUNCE_MS = 140


@dataclass
class PdfItem:
    path: Path
    page_count: int | None = None
    modified: str | None = None
    error: str | None = None

    @property
    def ready(self) -> bool:
        return self.error is None and self.page_count is not None and self.page_count > 0


@dataclass
class RenderResult:
    token: int
    page_index: int
    page_count: int
    zoom: float
    image: object
    width: int
    height: int
    hit_rects: list[tuple[float, float, float, float]]


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
    base = Path(__file__).resolve().parent
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        meipass = Path(getattr(sys, "_MEIPASS", exe_dir))
        return [
            exe_dir / ".." / ".." / "02_assets" / "dake_icon.ico",
            exe_dir / ".." / ".." / ".." / "02_assets" / "dake_icon.ico",
            meipass / "dake_icon.ico",
        ]
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
    message = str(exc).strip().replace("\n", " ")
    return message or UI_TEXT["status_error"]


def is_pdf_path(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() == ".pdf"


def format_modified(path: Path) -> str:
    try:
        return datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y/%m/%d %H:%M")
    except Exception:
        return UI_TEXT["modified_unknown"]


def read_pdf_summary(path: Path) -> tuple[int | None, str | None]:
    if fitz is None:
        return None, UI_TEXT["error_dependency"]
    try:
        with fitz.open(str(path)) as doc:
            return int(doc.page_count), None
    except Exception as exc:
        return None, humanize_error(exc)


def split_drop_files(root: tk.Tk, raw_value: str) -> list[Path]:
    try:
        values = root.tk.splitlist(raw_value)
    except Exception:
        values = raw_value.split()
    return [Path(value) for value in values]


def cleanup_later(path: Path, delay_seconds: int = 90) -> None:
    def worker() -> None:
        time.sleep(delay_seconds)
        try:
            path.unlink(missing_ok=True)
        except Exception:
            pass

    threading.Thread(target=worker, daemon=True).start()


class DakePdfViewerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.font_family = detect_font_family(root)
        self.items: dict[str, PdfItem] = {}
        self.current_path: Path | None = None
        self.current_page_index = 0
        self.current_page_count = 0
        self.zoom = 1.0
        self.rotation = 0
        self.fit_to_width = True
        self.search_visible = False
        self.search_term = ""
        self.search_results: list[tuple[int, list[object]]] = []
        self.search_result_index = -1
        self.current_highlight_rects: list[object] = []
        self.photo_image: object | None = None
        self.render_token = 0
        self.render_after_id: str | None = None
        self.result_queue: queue.Queue[tuple[str, object]] = queue.Queue()

        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(*WINDOW_MIN_SIZE)
        self.root.configure(bg=THEME["background"])
        apply_window_icon(self.root)
        self.configure_styles()
        self.build_ui()
        self.bind_events()
        self.poll_results()
        self.show_empty_view()
        self.set_status(UI_TEXT["status_idle"])

    def configure_styles(self) -> None:
        default_font = (self.font_family, 10)
        self.root.option_add("*Font", default_font)
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TFrame", background=THEME["background"])
        style.configure("Card.TFrame", background=THEME["card"])
        style.configure("Header.TLabel", background=THEME["background"], foreground=THEME["text"])
        style.configure("Muted.TLabel", background=THEME["background"], foreground=THEME["muted"])
        style.configure("Card.TLabel", background=THEME["card"], foreground=THEME["text"])
        style.configure("MutedCard.TLabel", background=THEME["card"], foreground=THEME["muted"])
        style.configure(
            "Primary.TButton",
            background=THEME["accent"],
            foreground="#FFFFFF",
            bordercolor=THEME["accent"],
            focusthickness=0,
            padding=(14, 8),
        )
        style.map("Primary.TButton", background=[("active", THEME["accent_hover"])])
        style.configure("Quiet.TButton", padding=(12, 8))
        style.configure(
            "Treeview",
            background=THEME["card"],
            fieldbackground=THEME["card"],
            foreground=THEME["text"],
            rowheight=30,
            bordercolor=THEME["border"],
            lightcolor=THEME["border"],
            darkcolor=THEME["border"],
        )
        style.configure("Treeview.Heading", background=THEME["soft"], foreground=THEME["text"])
        style.map("Treeview", background=[("selected", THEME["selection_bg"])], foreground=[("selected", THEME["text"])])
        style.configure("Vertical.TScrollbar", background=THEME["soft"], troughcolor=THEME["background"])

    def build_ui(self) -> None:
        outer = ttk.Frame(self.root, style="TFrame", padding=(18, 16, 18, 10))
        outer.pack(fill=tk.BOTH, expand=True)

        header = ttk.Frame(outer, style="TFrame")
        header.pack(fill=tk.X, pady=(0, 12))
        title_block = ttk.Frame(header, style="TFrame")
        title_block.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Label(
            title_block,
            text=UI_TEXT["main_title"],
            style="Header.TLabel",
            font=(self.font_family, 18, "bold"),
        ).pack(anchor=tk.W)
        ttk.Label(
            title_block,
            text=UI_TEXT["main_description"],
            style="Muted.TLabel",
            font=(self.font_family, 10),
        ).pack(anchor=tk.W, pady=(3, 0))

        action_bar = ttk.Frame(header, style="TFrame")
        action_bar.pack(side=tk.RIGHT)
        ttk.Button(
            action_bar,
            text=UI_TEXT["button_select_folder"],
            style="Primary.TButton",
            command=self.select_folder,
        ).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(
            action_bar,
            text=UI_TEXT["button_add_pdf"],
            style="Quiet.TButton",
            command=self.select_pdf_files,
        ).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(
            action_bar,
            text=UI_TEXT["button_clear"],
            style="Quiet.TButton",
            command=self.clear_list,
        ).pack(side=tk.LEFT)

        self.main_pane = ttk.Panedwindow(outer, orient=tk.HORIZONTAL)
        self.main_pane.pack(fill=tk.BOTH, expand=True)

        self.left_frame = ttk.Frame(self.main_pane, style="Card.TFrame", padding=(12, 12))
        self.right_frame = ttk.Frame(self.main_pane, style="Card.TFrame", padding=(0, 0))
        self.main_pane.add(self.left_frame, weight=0)
        self.main_pane.add(self.right_frame, weight=1)

        ttk.Label(
            self.left_frame,
            text=UI_TEXT["file_list_title"],
            style="Card.TLabel",
            font=(self.font_family, 11, "bold"),
        ).pack(anchor=tk.W, pady=(0, 8))

        list_area = ttk.Frame(self.left_frame, style="Card.TFrame")
        list_area.pack(fill=tk.BOTH, expand=True)
        columns = ("pages", "modified")
        self.tree = ttk.Treeview(list_area, columns=columns, show="tree headings", selectmode="browse")
        self.tree.heading("#0", text=UI_TEXT["column_file"], anchor=tk.W)
        self.tree.heading("pages", text=UI_TEXT["column_pages"], anchor=tk.CENTER)
        self.tree.heading("modified", text=UI_TEXT["column_modified"], anchor=tk.W)
        self.tree.column("#0", width=190, minwidth=120, stretch=True)
        self.tree.column("pages", width=54, minwidth=44, anchor=tk.CENTER, stretch=False)
        self.tree.column("modified", width=112, minwidth=90, anchor=tk.W, stretch=False)
        tree_scroll = ttk.Scrollbar(list_area, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.search_frame = ttk.Frame(self.right_frame, style="Card.TFrame", padding=(12, 10, 12, 8))
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(self.search_frame, textvariable=self.search_var, font=(self.font_family, 11))
        self.search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.search_status = ttk.Label(self.search_frame, text="", style="MutedCard.TLabel")
        self.search_status.pack(side=tk.LEFT, padx=(10, 0))

        viewer_shell = ttk.Frame(self.right_frame, style="Card.TFrame")
        viewer_shell.pack(fill=tk.BOTH, expand=True)
        self.canvas = tk.Canvas(
            viewer_shell,
            bg=THEME["background"],
            highlightthickness=0,
            borderwidth=0,
            takefocus=True,
        )
        self.v_scroll = ttk.Scrollbar(viewer_shell, orient=tk.VERTICAL, command=self.canvas.yview)
        self.h_scroll = ttk.Scrollbar(viewer_shell, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=self.v_scroll.set, xscrollcommand=self.h_scroll.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.v_scroll.grid(row=0, column=1, sticky="ns")
        self.h_scroll.grid(row=1, column=0, sticky="ew")
        viewer_shell.rowconfigure(0, weight=1)
        viewer_shell.columnconfigure(0, weight=1)

        footer = ttk.Frame(outer, style="TFrame")
        footer.pack(fill=tk.X, pady=(8, 0))
        self.status_var = tk.StringVar(value=UI_TEXT["status_idle"])
        self.status_label = ttk.Label(
            footer,
            textvariable=self.status_var,
            style="Muted.TLabel",
            font=(self.font_family, 9),
        )
        self.status_label.pack(side=tk.LEFT)
        footer_text = (
            UI_TEXT["footer_left"]
            + UI_TEXT["footer_separator"]
            + UI_TEXT["footer_link_1"]
            + UI_TEXT["footer_separator"]
            + UI_TEXT["footer_link_2"]
            + UI_TEXT["footer_separator"]
            + UI_TEXT["footer_copyright"]
        )
        ttk.Label(
            footer,
            text=footer_text,
            style="Muted.TLabel",
            font=(self.font_family, 9),
        ).pack(side=tk.RIGHT)

    def bind_events(self) -> None:
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        self.canvas.bind("<Button-1>", lambda _event: self.canvas.focus_set())
        self.canvas.bind("<Configure>", self.on_canvas_configure)
        self.canvas.bind("<MouseWheel>", self.on_mouse_wheel)
        self.canvas.bind("<Button-4>", self.on_mouse_wheel)
        self.canvas.bind("<Button-5>", self.on_mouse_wheel)
        self.root.bind_all("<Up>", self.scroll_up)
        self.root.bind_all("<Down>", self.scroll_down)
        self.root.bind_all("<Left>", self.previous_page)
        self.root.bind_all("<Right>", self.next_page)
        self.root.bind_all("<Prior>", self.previous_page)
        self.root.bind_all("<Next>", self.next_page)
        self.root.bind_all("<Home>", self.first_page)
        self.root.bind_all("<End>", self.last_page)
        self.root.bind_all("<KeyPress-r>", self.rotate_right)
        self.root.bind_all("<KeyPress-R>", self.rotate_right)
        self.root.bind_all("<KeyPress-l>", self.rotate_left)
        self.root.bind_all("<KeyPress-L>", self.rotate_left)
        self.root.bind_all("<KeyPress-0>", self.reset_rotation)
        self.root.bind_all("<Control-KeyPress-0>", self.fit_width)
        self.root.bind_all("<Control-KeyPress-f>", self.show_search)
        self.root.bind_all("<Control-KeyPress-F>", self.show_search)
        self.root.bind_all("<Control-KeyPress-p>", self.print_current_page)
        self.root.bind_all("<Control-KeyPress-P>", self.print_current_page)
        self.root.bind_all("<Escape>", self.on_escape)
        self.search_entry.bind("<Return>", self.search_next)
        self.search_entry.bind("<Shift-Return>", self.search_previous)

        if DND_ENABLED and DND_FILES is not None:
            for widget in (self.root, self.tree, self.canvas):
                try:
                    widget.drop_target_register(DND_FILES)
                    widget.dnd_bind("<<Drop>>", self.on_drop)
                except Exception:
                    pass

    def poll_results(self) -> None:
        while True:
            try:
                kind, payload = self.result_queue.get_nowait()
            except queue.Empty:
                break
            if kind == "summary":
                self.apply_summary(payload)  # type: ignore[arg-type]
            elif kind == "render":
                self.apply_render(payload)  # type: ignore[arg-type]
            elif kind == "render_error":
                self.apply_render_error(payload)  # type: ignore[arg-type]
            elif kind == "search":
                self.apply_search_results(payload)  # type: ignore[arg-type]
            elif kind == "print_error":
                self.set_status(UI_TEXT["error_print"])
        self.root.after(POLL_INTERVAL_MS, self.poll_results)

    def set_status(self, message: str) -> None:
        self.status_var.set(message)

    def select_folder(self) -> None:
        folder = filedialog.askdirectory(title=UI_TEXT["dialog_folder_title"])
        if not folder:
            return
        paths = sorted(
            [path for path in Path(folder).iterdir() if path.is_file() and path.suffix.lower() == ".pdf"],
            key=lambda p: p.name.casefold(),
        )
        if not paths:
            self.set_status(UI_TEXT["error_no_pdf"])
            return
        self.add_pdf_paths(paths, replace=True)

    def select_pdf_files(self) -> None:
        files = filedialog.askopenfilenames(
            title=UI_TEXT["dialog_pdf_title"],
            filetypes=[(UI_TEXT["dialog_pdf_filter"], "*.pdf"), ("PDF", "*.PDF")],
        )
        if files:
            self.add_pdf_paths([Path(file) for file in files], replace=False)

    def on_drop(self, event: tk.Event) -> None:
        paths = split_drop_files(self.root, getattr(event, "data", ""))
        self.add_pdf_paths(paths, replace=False)

    def add_pdf_paths(self, paths: list[Path], replace: bool = False) -> None:
        if replace:
            self.clear_list(clear_view=True)
        valid_paths = sorted({path.resolve() for path in paths if is_pdf_path(path)}, key=lambda p: p.name.casefold())
        if not valid_paths:
            self.set_status(UI_TEXT["error_invalid_file"])
            return
        for path in valid_paths:
            key = str(path)
            if key in self.items:
                continue
            item = PdfItem(path=path, modified=format_modified(path))
            self.items[key] = item
            self.tree.insert(
                "",
                tk.END,
                iid=key,
                text=path.name,
                values=(UI_TEXT["page_count_unknown"], item.modified or UI_TEXT["modified_unknown"]),
            )
            self.load_summary_async(path)
        self.sort_tree_by_name()
        self.set_status(UI_TEXT["status_loading"])

    def sort_tree_by_name(self) -> None:
        rows = sorted(self.tree.get_children(""), key=lambda iid: self.items[str(iid)].path.name.casefold())
        for index, iid in enumerate(rows):
            self.tree.move(iid, "", index)

    def clear_list(self, clear_view: bool = True) -> None:
        for iid in self.tree.get_children(""):
            self.tree.delete(iid)
        self.items.clear()
        if clear_view:
            self.close_current_pdf()
            self.show_empty_view()
            self.set_status(UI_TEXT["status_idle"])

    def load_summary_async(self, path: Path) -> None:
        def worker() -> None:
            page_count, error = read_pdf_summary(path)
            self.result_queue.put(("summary", (str(path), page_count, error, format_modified(path))))

        threading.Thread(target=worker, daemon=True).start()

    def apply_summary(self, payload: tuple[str, int | None, str | None, str]) -> None:
        key, page_count, error, modified = payload
        item = self.items.get(key)
        if item is None:
            return
        item.page_count = page_count
        item.error = error
        item.modified = modified
        pages = UI_TEXT["page_count_format"].format(count=page_count) if page_count else UI_TEXT["page_count_unknown"]
        if self.tree.exists(key):
            self.tree.item(key, values=(pages, modified))
        if self.current_path is None:
            self.set_status(UI_TEXT["status_ready"])

    def on_tree_select(self, _event: tk.Event) -> None:
        selected = self.tree.selection()
        if not selected:
            return
        item = self.items.get(str(selected[0]))
        if item is None:
            return
        self.open_pdf(item.path)

    def close_current_pdf(self) -> None:
        self.current_path = None
        self.current_page_index = 0
        self.current_page_count = 0
        self.current_highlight_rects = []
        self.search_results = []
        self.search_result_index = -1
        self.search_term = ""
        self.search_var.set("")
        self.rotation = 0
        self.fit_to_width = True
        self.photo_image = None

    def open_pdf(self, path: Path) -> None:
        if not path.exists():
            self.set_status(UI_TEXT["error_open_pdf"])
            return
        self.current_path = path
        self.current_page_index = 0
        self.current_page_count = self.items.get(str(path), PdfItem(path)).page_count or 0
        self.rotation = 0
        self.fit_to_width = True
        self.zoom = self.compute_fit_zoom(path, 0)
        self.search_results = []
        self.search_result_index = -1
        self.search_term = ""
        self.search_var.set("")
        self.hide_search()
        self.canvas.focus_set()
        self.request_render()

    def show_empty_view(self) -> None:
        self.canvas.delete("all")
        self.photo_image = None
        width = max(self.canvas.winfo_width(), 640)
        height = max(self.canvas.winfo_height(), 420)
        self.canvas.configure(scrollregion=(0, 0, width, height))
        center_x = width / 2
        center_y = height / 2 - 16
        self.canvas.create_text(
            center_x,
            center_y,
            text=UI_TEXT["viewer_empty_title"],
            fill=THEME["text"],
            font=(self.font_family, 18, "bold"),
        )
        self.canvas.create_text(
            center_x,
            center_y + 36,
            text=UI_TEXT["viewer_empty_subtitle"],
            fill=THEME["muted"],
            font=(self.font_family, 11),
        )
        self.canvas.create_text(
            center_x,
            center_y + 64,
            text=UI_TEXT["viewer_empty_hint"],
            fill=THEME["muted"],
            font=(self.font_family, 10),
        )

    def on_canvas_configure(self, _event: tk.Event) -> None:
        if self.current_path is None:
            self.show_empty_view()
            return
        if self.fit_to_width:
            self.zoom = self.compute_fit_zoom(self.current_path, self.current_page_index)
            self.schedule_render()

    def compute_fit_zoom(self, path: Path, page_index: int) -> float:
        if fitz is None:
            return 1.0
        try:
            with fitz.open(str(path)) as doc:
                if doc.page_count == 0:
                    return 1.0
                page = doc.load_page(max(0, min(page_index, doc.page_count - 1)))
                rect = page.rect
                page_width = rect.height if self.rotation in (90, 270) else rect.width
                canvas_width = max(self.canvas.winfo_width(), 400)
                zoom = (canvas_width - FIT_PADDING) / max(page_width, 1)
                return max(ZOOM_MIN, min(ZOOM_MAX, zoom))
        except Exception:
            return 1.0

    def schedule_render(self) -> None:
        if self.render_after_id is not None:
            self.root.after_cancel(self.render_after_id)
        self.render_after_id = self.root.after(ZOOM_DEBOUNCE_MS, self.request_render)

    def request_render(self) -> None:
        self.render_after_id = None
        if self.current_path is None:
            return
        if fitz is None or Image is None or ImageTk is None:
            self.set_status(UI_TEXT["error_dependency"])
            return
        path = self.current_path
        page_index = self.current_page_index
        zoom = self.zoom
        rotation = self.rotation
        highlight_rects = list(self.current_highlight_rects)
        self.render_token += 1
        token = self.render_token
        self.set_status(UI_TEXT["status_loading"])

        def worker() -> None:
            try:
                result = self.render_page(path, page_index, zoom, rotation, highlight_rects, token)
                self.result_queue.put(("render", result))
            except Exception as exc:
                self.result_queue.put(("render_error", (token, humanize_error(exc))))

        threading.Thread(target=worker, daemon=True).start()

    def render_page(
        self,
        path: Path,
        page_index: int,
        zoom: float,
        rotation: int,
        highlight_rects: list[object],
        token: int,
    ) -> RenderResult:
        if fitz is None or Image is None or ImageDraw is None:
            raise RuntimeError(UI_TEXT["error_dependency"])
        with fitz.open(str(path)) as doc:
            if doc.page_count == 0:
                raise RuntimeError(UI_TEXT["error_open_pdf"])
            safe_index = max(0, min(page_index, doc.page_count - 1))
            page = doc.load_page(safe_index)
            scale = max(ZOOM_MIN, min(ZOOM_MAX, zoom)) * RENDER_DPI_SCALE
            matrix = fitz.Matrix(scale, scale).prerotate(rotation)
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            mode = "RGB" if pix.n < 4 else "RGBA"
            image = Image.frombytes(mode, (pix.width, pix.height), pix.samples)
            draw = ImageDraw.Draw(image, "RGBA")
            drawn_rects: list[tuple[float, float, float, float]] = []
            for raw_rect in highlight_rects:
                try:
                    rect = fitz.Rect(raw_rect) * matrix
                    x0 = float(rect.x0 - pix.x)
                    y0 = float(rect.y0 - pix.y)
                    x1 = float(rect.x1 - pix.x)
                    y1 = float(rect.y1 - pix.y)
                    drawn_rects.append((x0, y0, x1, y1))
                    draw.rectangle((x0, y0, x1, y1), fill=(255, 232, 122, 92), outline=(232, 174, 0, 150))
                except Exception:
                    continue
            return RenderResult(
                token=token,
                page_index=safe_index,
                page_count=int(doc.page_count),
                zoom=zoom,
                image=image,
                width=pix.width,
                height=pix.height,
                hit_rects=drawn_rects,
            )

    def apply_render(self, result: RenderResult) -> None:
        if result.token != self.render_token:
            return
        if ImageTk is None:
            self.set_status(UI_TEXT["error_dependency"])
            return
        self.current_page_index = result.page_index
        self.current_page_count = result.page_count
        self.photo_image = ImageTk.PhotoImage(result.image)
        self.canvas.delete("all")
        canvas_width = max(self.canvas.winfo_width(), result.width + FIT_PADDING)
        x = max((canvas_width - result.width) // 2, PAGE_GAP)
        y = PAGE_GAP
        self.canvas.create_rectangle(
            x - 1,
            y - 1,
            x + result.width + 1,
            y + result.height + 1,
            fill=THEME["card"],
            outline=THEME["border"],
        )
        self.canvas.create_image(x, y, anchor=tk.NW, image=self.photo_image)
        region_width = max(canvas_width, result.width + (PAGE_GAP * 2))
        region_height = result.height + (PAGE_GAP * 2)
        self.canvas.configure(scrollregion=(0, 0, region_width, region_height))
        self.set_status(
            UI_TEXT["status_viewing"]
            + "  "
            + UI_TEXT["status_page"].format(current=result.page_index + 1, total=result.page_count)
            + f"  {UI_TEXT['status_zoom']} {int(result.zoom * 100)}%"
        )

    def apply_render_error(self, payload: tuple[int, str]) -> None:
        token, message = payload
        if token != self.render_token:
            return
        self.canvas.delete("all")
        self.set_status(UI_TEXT["error_open_pdf"])
        self.canvas.create_text(
            max(self.canvas.winfo_width(), 640) / 2,
            max(self.canvas.winfo_height(), 420) / 2,
            text=UI_TEXT["error_open_pdf"] + "\n" + message,
            fill=THEME["error"],
            font=(self.font_family, 12),
            justify=tk.CENTER,
        )

    def on_mouse_wheel(self, event: tk.Event) -> str:
        if getattr(event, "state", 0) & 0x0004:
            delta = self.normalized_wheel_delta(event)
            if delta > 0:
                self.change_zoom(ZOOM_STEP)
            else:
                self.change_zoom(1 / ZOOM_STEP)
            return "break"
        delta = self.normalized_wheel_delta(event)
        self.canvas.yview_scroll(-delta * 3, "units")
        return "break"

    def normalized_wheel_delta(self, event: tk.Event) -> int:
        num = getattr(event, "num", None)
        if num == 4:
            return 1
        if num == 5:
            return -1
        delta = getattr(event, "delta", 0)
        return 1 if delta > 0 else -1

    def change_zoom(self, factor: float) -> None:
        if self.current_path is None:
            return
        self.fit_to_width = False
        self.zoom = max(ZOOM_MIN, min(ZOOM_MAX, self.zoom * factor))
        self.set_status(f"{UI_TEXT['status_zoom']} {int(self.zoom * 100)}%")
        self.schedule_render()

    def fit_width(self, event: tk.Event | None = None) -> str:
        if self.current_path is None:
            return "break"
        self.fit_to_width = True
        self.zoom = self.compute_fit_zoom(self.current_path, self.current_page_index)
        self.set_status(f"{UI_TEXT['status_zoom']} {int(self.zoom * 100)}%")
        self.request_render()
        return "break"

    def scroll_up(self, event: tk.Event | None = None) -> str:
        if self.search_entry == self.root.focus_get():
            return ""
        self.canvas.yview_scroll(-4, "units")
        return "break"

    def scroll_down(self, event: tk.Event | None = None) -> str:
        if self.search_entry == self.root.focus_get():
            return ""
        self.canvas.yview_scroll(4, "units")
        return "break"

    def previous_page(self, event: tk.Event | None = None) -> str:
        if self.current_path is None or self.search_entry == self.root.focus_get():
            return ""
        if self.current_page_index > 0:
            self.current_page_index -= 1
            self.current_highlight_rects = []
            self.request_render()
        return "break"

    def next_page(self, event: tk.Event | None = None) -> str:
        if self.current_path is None or self.search_entry == self.root.focus_get():
            return ""
        if self.current_page_count and self.current_page_index < self.current_page_count - 1:
            self.current_page_index += 1
            self.current_highlight_rects = []
            self.request_render()
        return "break"

    def first_page(self, event: tk.Event | None = None) -> str:
        if self.current_path is None or self.search_entry == self.root.focus_get():
            return ""
        self.current_page_index = 0
        self.current_highlight_rects = []
        self.request_render()
        return "break"

    def last_page(self, event: tk.Event | None = None) -> str:
        if self.current_path is None or self.search_entry == self.root.focus_get():
            return ""
        if self.current_page_count:
            self.current_page_index = self.current_page_count - 1
            self.current_highlight_rects = []
            self.request_render()
        return "break"

    def rotate_right(self, event: tk.Event | None = None) -> str:
        if self.current_path is None or self.search_entry == self.root.focus_get():
            return ""
        self.rotation = (self.rotation + 90) % 360
        if self.fit_to_width:
            self.zoom = self.compute_fit_zoom(self.current_path, self.current_page_index)
        self.set_status(UI_TEXT["status_rotate"])
        self.request_render()
        return "break"

    def rotate_left(self, event: tk.Event | None = None) -> str:
        if self.current_path is None or self.search_entry == self.root.focus_get():
            return ""
        self.rotation = (self.rotation - 90) % 360
        if self.fit_to_width:
            self.zoom = self.compute_fit_zoom(self.current_path, self.current_page_index)
        self.set_status(UI_TEXT["status_rotate"])
        self.request_render()
        return "break"

    def reset_rotation(self, event: tk.Event | None = None) -> str:
        if self.current_path is None:
            return "break"
        if getattr(event, "state", 0) & 0x0004:
            return self.fit_width(event)
        self.rotation = 0
        if self.fit_to_width:
            self.zoom = self.compute_fit_zoom(self.current_path, self.current_page_index)
        self.set_status(UI_TEXT["status_reset_rotate"])
        self.request_render()
        return "break"

    def show_search(self, event: tk.Event | None = None) -> str:
        if self.current_path is None:
            return "break"
        if not self.search_visible:
            self.search_frame.pack(fill=tk.X, before=self.canvas.master)
            self.search_visible = True
        self.search_entry.focus_set()
        self.search_entry.selection_range(0, tk.END)
        self.search_entry.icursor(tk.END)
        return "break"

    def hide_search(self) -> None:
        if self.search_visible:
            self.search_frame.pack_forget()
        self.search_visible = False
        self.search_status.configure(text="")
        self.canvas.focus_set()

    def on_escape(self, event: tk.Event | None = None) -> str:
        if self.search_visible:
            self.hide_search()
            return "break"
        self.canvas.focus_set()
        return "break"

    def search_next(self, event: tk.Event | None = None) -> str:
        self.run_search(direction=1)
        return "break"

    def search_previous(self, event: tk.Event | None = None) -> str:
        self.run_search(direction=-1)
        return "break"

    def run_search(self, direction: int) -> None:
        if self.current_path is None:
            return
        term = self.search_var.get().strip()
        if not term:
            return
        if term == self.search_term and self.search_results:
            self.move_search_result(direction)
            return
        self.search_term = term
        self.search_status.configure(text=UI_TEXT["status_loading"])
        path = self.current_path

        def worker() -> None:
            results: list[tuple[int, list[object]]] = []
            text_seen = False
            try:
                if fitz is None:
                    raise RuntimeError(UI_TEXT["error_dependency"])
                with fitz.open(str(path)) as doc:
                    for page_index in range(doc.page_count):
                        page = doc.load_page(page_index)
                        if page.get_text("text").strip():
                            text_seen = True
                        rects = page.search_for(term)
                        if rects:
                            results.append((page_index, rects))
                self.result_queue.put(("search", (str(path), term, direction, text_seen, results)))
            except Exception:
                self.result_queue.put(("search", (str(path), term, direction, False, [])))

        threading.Thread(target=worker, daemon=True).start()

    def apply_search_results(
        self,
        payload: tuple[str, str, int, bool, list[tuple[int, list[object]]]],
    ) -> None:
        key, term, direction, text_seen, results = payload
        if self.current_path is None or key != str(self.current_path) or term != self.search_term:
            return
        self.search_results = results
        self.search_result_index = -1 if direction > 0 else 0
        if not text_seen:
            self.search_status.configure(text=UI_TEXT["search_not_available"])
            self.current_highlight_rects = []
            self.request_render()
            return
        if not results:
            self.search_status.configure(text=UI_TEXT["search_not_found"])
            self.current_highlight_rects = []
            self.request_render()
            return
        self.move_search_result(direction)

    def move_search_result(self, direction: int) -> None:
        if not self.search_results:
            return
        self.search_result_index = (self.search_result_index + direction) % len(self.search_results)
        page_index, rects = self.search_results[self.search_result_index]
        self.current_page_index = page_index
        self.current_highlight_rects = rects
        self.search_status.configure(text=UI_TEXT["status_page"].format(current=page_index + 1, total=self.current_page_count))
        self.request_render()

    def print_current_page(self, event: tk.Event | None = None) -> str:
        if self.current_path is None or fitz is None:
            self.set_status(UI_TEXT["error_print"])
            return "break"
        path = self.current_path
        page_index = self.current_page_index
        self.set_status(UI_TEXT["status_print_current"])

        def worker() -> None:
            temp_path: Path | None = None
            try:
                with fitz.open(str(path)) as src:
                    safe_index = max(0, min(page_index, src.page_count - 1))
                    out = fitz.open()
                    out.insert_pdf(src, from_page=safe_index, to_page=safe_index)
                    temp_dir = Path(tempfile.gettempdir())
                    temp_path = temp_dir / f"dake_pdf_viewer_print_{os.getpid()}_{int(time.time())}.pdf"
                    out.save(str(temp_path))
                    out.close()
                if sys.platform.startswith("win"):
                    os.startfile(str(temp_path), "print")  # type: ignore[attr-defined]
                else:
                    opener = shutil.which("lp") or shutil.which("lpr")
                    if opener is None:
                        raise RuntimeError(UI_TEXT["error_print"])
                    os.spawnlp(os.P_NOWAIT, opener, opener, str(temp_path))
                cleanup_later(temp_path)
            except Exception:
                if temp_path is not None:
                    try:
                        temp_path.unlink(missing_ok=True)
                    except Exception:
                        pass
                self.result_queue.put(("print_error", None))

        threading.Thread(target=worker, daemon=True).start()
        return "break"


def main() -> None:
    set_windows_app_id()
    root = make_root()
    DakePdfViewerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
