# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import queue
import re
import sys
import threading
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable

import fitz  # PyMuPDF
from pypdf import PdfReader, PdfWriter
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except Exception:  # pragma: no cover - optional runtime fallback
    DND_FILES = None
    TkinterDnD = None


APP_NAME = "DakePDF結合mini"
WINDOW_TITLE = "PDF結合mini"
COMPANY_NAME = "KIKUTA YUKIHIKO"
FILE_DESCRIPTION = "PDF結合mini"
PRODUCT_NAME = "DAKE PDF Merge Mini"
LEGAL_COPYRIGHT = "© 2026 KIKUTA YUKIHIKO"
ORIGINAL_FILENAME = "DakePDF_Merge_Mini.exe"
INTERNAL_NAME = "DakePDF_Merge_Mini"

UI_TEXT = {
    "brand_series": "シンプルそれDAKEシリーズ",
    "header_subtitle": "止まらない、迷わない、すぐ終わる。",
    "main_title": "PDFを結合する",
    "main_description": "最大5件・並び替えて結合",
    "button_add": "PDFを追加",
    "button_refresh": "リフレッシュ",
    "button_execute": "結合して保存",
    "button_remove": "×",
    "empty_title": "PDFを追加",
    "empty_subtitle": "ドラッグ＆ドロップ または クリック",
    "status_idle": "PDFを追加してください",
    "status_loading": "読み込み中...",
    "status_ready": "準備完了",
    "status_refreshed": "リフレッシュしました",
    "status_dragging": "移動先で離してください",
    "status_reordered": "並び替えました",
    "status_processing": "結合中...",
    "status_saving": "保存中...",
    "status_complete": "保存しました",
    "status_complete_with_name": "保存しました：{filename}",
    "status_error": "エラー",
    "status_count": "{count}件のPDFを読み込みました",
    "page_count": "{pages}ページ",
    "error_title": "確認してください",
    "error_too_many_files": "PDFは5ファイルまでです。ファイル数を減らしてから追加してください。",
    "error_too_large_file": "このPDFは大きすぎるため処理できません。",
    "error_too_many_pages": "ページ数が多すぎるため処理できません。",
    "error_total_too_large": "合計サイズが大きすぎるため処理できません。",
    "error_total_too_many_pages": "合計ページ数が多すぎるため処理できません。",
    "error_invalid_pdf": "PDFファイルを読み込めませんでした。",
    "error_no_pdf": "PDFを追加してください。",
    "error_busy": "処理中です。しばらくお待ちください。",
    "error_save_failed": "保存できませんでした。",
    "error_merge_failed": "結合できませんでした。",
    "error_unique_name_failed": "保存ファイル名を作成できませんでした。保存先を変更してください。",
    "mini_policy": "社内配布用mini版では、軽いPDFだけを対象にしています。",
    "reduce_request": "ファイル数・容量・ページ数を減らしてください。",
    "dialog_save_title": "結合PDFを保存",
    "dialog_open_title": "PDFを選択",
    "output_suffix": "結合",
    "fallback_output_base": "PDF結合",
    "filetype_pdf": "PDFファイル",
    "filetype_all": "すべてのファイル",
    "complete_title": "完了",
    "complete_message": "保存しました。",
}

MAX_FILES = 5
MAX_FILE_SIZE = 50 * 1024 * 1024
MAX_TOTAL_SIZE = 150 * 1024 * 1024
MAX_PAGES_PER_FILE = 100
MAX_TOTAL_PAGES = 200
THUMBNAIL_WIDTH = 116
THUMBNAIL_HEIGHT = 150

FONT_CANDIDATES = ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo")
COLOR_BASE = "#F6F7F9"
COLOR_CARD = "#FFFFFF"
COLOR_TEXT = "#1E2430"
COLOR_MUTED = "#667085"
COLOR_BORDER = "#E6EAF0"
COLOR_ACCENT = "#2F6FED"
COLOR_ACCENT_HOVER = "#2458BF"
COLOR_SELECTED_BG = "#EAF2FF"
COLOR_SELECTED_BORDER = "#7AA7FF"
COLOR_SUCCESS = "#12B76A"
COLOR_ERROR = "#D92D20"
COLOR_DISABLED_BG = "#E9EDF3"


@dataclass(frozen=True)
class PdfItem:
    path: Path
    name: str
    size: int
    pages: int
    thumbnail: "tk.PhotoImage"


class LimitError(Exception):
    def __init__(self, message: str, file_name: str | None = None) -> None:
        self.message = message
        self.file_name = file_name
        super().__init__(message)


def format_size(size: int) -> str:
    mb = size / (1024 * 1024)
    if mb >= 1:
        return f"{mb:.1f}MB"
    return f"{max(1, round(size / 1024))}KB"


def display_filename(name: str, limit: int = 28) -> str:
    if len(name) <= limit:
        return name
    stem = Path(name).stem
    suffix = Path(name).suffix
    keep = max(8, limit - len(suffix) - 3)
    head = max(4, keep // 2)
    tail = max(4, keep - head)
    return f"{stem[:head]}...{stem[-tail:]}{suffix}"


def sanitize_filename(name: str) -> str:
    sanitized = re.sub(r'[\\/:*?"<>|]', "_", name).strip(" .")
    if not sanitized:
        return UI_TEXT["fallback_output_base"]
    return sanitized[:80].rstrip(" .") or UI_TEXT["fallback_output_base"]


def build_default_output_name(first_file: Path | None = None) -> str:
    if first_file:
        base = sanitize_filename(first_file.stem)
    else:
        base = UI_TEXT["fallback_output_base"]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{base}_{UI_TEXT['output_suffix']}_{timestamp}.pdf"


def normalize_output_path(path: Path) -> Path:
    if path.suffix.lower() != ".pdf":
        return path.with_suffix(".pdf")
    return path


def get_unique_path(path: Path) -> Path:
    output_path = normalize_output_path(path)
    if not output_path.exists():
        return output_path

    stem = output_path.stem
    suffix = output_path.suffix or ".pdf"
    parent = output_path.parent
    for number in range(2, 1000):
        candidate = parent / f"{stem}_{number}{suffix}"
        if not candidate.exists():
            return candidate
    raise LimitError(UI_TEXT["error_unique_name_failed"], output_path.name)


def resolve_font_family(root: tk.Tk) -> str:
    available = set(tkfont.families(root))
    for family in FONT_CANDIDATES:
        if family in available:
            return family
    return "TkDefaultFont"


def find_dake_icon_path() -> Path | None:
    search_roots = [Path(__file__).resolve().parent, Path.cwd().resolve()]
    seen: set[Path] = set()
    for root in search_roots:
        for parent in (root, *root.parents):
            if parent in seen:
                continue
            seen.add(parent)
            icon_path = parent / "02_assets" / "dake_icon.ico"
            if icon_path.is_file():
                return icon_path
    return None


def apply_window_icon(root: tk.Tk) -> Path | None:
    icon_path = find_dake_icon_path()
    if not icon_path:
        return None
    try:
        root.iconbitmap(default=str(icon_path))
    except Exception:
        pass
    try:
        root.iconbitmap(str(icon_path))
    except Exception:
        pass
    return icon_path


def build_error_message(message: str, file_name: str | None = None, include_policy: bool = False) -> str:
    parts = []
    if file_name:
        parts.append(file_name)
    parts.append(message)
    if include_policy:
        parts.append(UI_TEXT["mini_policy"])
        parts.append(UI_TEXT["reduce_request"])
    return "\n".join(parts)


def normalize_pdf_paths(paths: Iterable[str]) -> list[Path]:
    result: list[Path] = []
    for raw_path in paths:
        if not raw_path:
            continue
        path = Path(raw_path).expanduser()
        if path.suffix.lower() != ".pdf" or not path.is_file():
            raise LimitError(UI_TEXT["error_invalid_pdf"], path.name if path.name else None)
        result.append(path)
    return result


def read_pdf_info(path: Path) -> tuple[int, int, bytes, int, int]:
    size = path.stat().st_size
    if size > MAX_FILE_SIZE:
        raise LimitError(UI_TEXT["error_too_large_file"], path.name)

    try:
        with fitz.open(path) as doc:
            pages = doc.page_count
            if pages <= 0:
                raise LimitError(UI_TEXT["error_invalid_pdf"], path.name)
            if pages > MAX_PAGES_PER_FILE:
                raise LimitError(UI_TEXT["error_too_many_pages"], path.name)
            page = doc.load_page(0)
            rect = page.rect
            scale = min(THUMBNAIL_WIDTH / rect.width, THUMBNAIL_HEIGHT / rect.height, 0.35)
            pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
            return size, pages, pix.tobytes("ppm"), pix.width, pix.height
    except LimitError:
        raise
    except Exception as exc:
        raise LimitError(UI_TEXT["error_invalid_pdf"], path.name) from exc


class PdfMergeMiniApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.items: list[PdfItem] = []
        self.cards: list[tk.Frame] = []
        self.worker_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.busy = False
        self.drag_index: int | None = None
        self.drag_source_index: int | None = None
        self.dragging = False

        self.root.title(WINDOW_TITLE)
        self.icon_path = apply_window_icon(self.root)
        self.root.geometry("760x480")
        self.root.minsize(760, 480)
        self.root.configure(bg=COLOR_BASE)
        self.font_family = resolve_font_family(self.root)

        self._configure_style()
        self._build_ui()
        self._setup_drop_target()
        self._set_status(UI_TEXT["status_idle"])
        self.root.after(80, self._handle_worker_queue)

    def _configure_style(self) -> None:
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TLabel", background=COLOR_BASE, foreground=COLOR_TEXT, font=(self.font_family, 10))

    def _build_ui(self) -> None:
        header = tk.Frame(self.root, bg=COLOR_BASE)
        header.pack(fill="x", padx=34, pady=(24, 20))
        header.grid_columnconfigure(0, weight=1)
        header.grid_columnconfigure(1, weight=0)
        title_area = tk.Frame(header, bg=COLOR_BASE)
        title_area.grid(row=0, column=0, sticky="ew")

        tk.Label(
            title_area,
            text=UI_TEXT["main_title"],
            bg=COLOR_BASE,
            fg=COLOR_TEXT,
            font=(self.font_family, 18, "bold"),
        ).pack(side="left", anchor="w")
        tk.Label(
            title_area,
            text=UI_TEXT["main_description"],
            bg=COLOR_BASE,
            fg=COLOR_MUTED,
            font=(self.font_family, 10),
        ).pack(side="left", anchor="w", padx=(20, 0), pady=(3, 0))

        header_actions = tk.Frame(header, bg=COLOR_BASE)
        header_actions.grid(row=0, column=1, sticky="e", padx=(16, 0))
        refresh_holder = tk.Frame(header_actions, bg=COLOR_BASE, width=118, height=34)
        refresh_holder.pack(side="left")
        refresh_holder.pack_propagate(False)
        self.refresh_button = self._create_secondary_button(refresh_holder, UI_TEXT["button_refresh"], self.refresh_files)
        self.refresh_button.pack(fill="both", expand=True)
        add_holder = tk.Frame(header_actions, bg=COLOR_BASE, width=132, height=34)
        add_holder.pack(side="left", padx=(8, 0))
        add_holder.pack_propagate(False)
        self.add_button = self._create_primary_button(add_holder, UI_TEXT["button_add"], self.choose_files)
        self.add_button.pack(fill="both", expand=True)

        self.center_frame = tk.Frame(self.root, bg=COLOR_BASE)
        self.center_frame.pack(fill="both", expand=True, padx=34, pady=(0, 16))

        self.empty_frame = tk.Frame(self.center_frame, bg=COLOR_CARD, highlightthickness=1, highlightbackground=COLOR_BORDER)
        self.empty_frame.pack(expand=True, pady=(0, 20))
        self.empty_frame.configure(width=440, height=198)
        self.empty_frame.pack_propagate(False)
        self.empty_frame.bind("<Button-1>", lambda _event: self.choose_files())

        empty_inner = tk.Frame(self.empty_frame, bg=COLOR_CARD)
        empty_inner.place(relx=0.5, rely=0.5, anchor="center")
        empty_title = tk.Label(
            empty_inner,
            text=UI_TEXT["empty_title"],
            bg=COLOR_CARD,
            fg=COLOR_TEXT,
            font=(self.font_family, 19, "bold"),
            cursor="hand2",
        )
        empty_title.pack()
        empty_subtitle = tk.Label(
            empty_inner,
            text=UI_TEXT["empty_subtitle"],
            bg=COLOR_CARD,
            fg=COLOR_MUTED,
            font=(self.font_family, 10),
            cursor="hand2",
        )
        empty_subtitle.pack(pady=(7, 0))
        empty_title.bind("<Button-1>", lambda _event: self.choose_files())
        empty_subtitle.bind("<Button-1>", lambda _event: self.choose_files())

        self.cards_frame = tk.Frame(self.center_frame, bg=COLOR_BASE)

        bottom_bar = tk.Frame(self.root, bg=COLOR_BASE)
        bottom_bar.pack(fill="x", padx=34, pady=(0, 20))
        self.status_label = tk.Label(
            bottom_bar,
            text="",
            bg=COLOR_BASE,
            fg=COLOR_MUTED,
            font=(self.font_family, 10),
        )
        self.status_label.pack(side="left", fill="x", expand=True)
        self.merge_button = self._create_primary_button(bottom_bar, UI_TEXT["button_execute"], self.ask_save_and_merge)
        self.merge_button.pack(side="right")
        self._update_buttons()

    def _create_primary_button(self, parent: tk.Widget, text: str, command: object) -> tk.Button:
        button = tk.Button(
            parent,
            text=text,
            command=command,
            height=1,
            bg=COLOR_ACCENT,
            fg=COLOR_CARD,
            activebackground=COLOR_ACCENT_HOVER,
            activeforeground=COLOR_CARD,
            disabledforeground=COLOR_MUTED,
            relief="flat",
            bd=0,
            highlightthickness=0,
            padx=10,
            pady=7,
            cursor="hand2",
            font=(self.font_family, 10, "bold"),
        )
        button.bind("<Enter>", lambda _event, target=button: self._set_button_hover(target, True))
        button.bind("<Leave>", lambda _event, target=button: self._set_button_hover(target, False))
        return button

    def _create_secondary_button(self, parent: tk.Widget, text: str, command: object) -> tk.Button:
        button = tk.Button(
            parent,
            text=text,
            command=command,
            height=1,
            bg=COLOR_CARD,
            fg=COLOR_MUTED,
            activebackground="#F0F3F8",
            activeforeground=COLOR_TEXT,
            disabledforeground=COLOR_MUTED,
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=COLOR_BORDER,
            highlightcolor=COLOR_BORDER,
            padx=10,
            pady=7,
            cursor="hand2",
            font=(self.font_family, 10, "bold"),
        )
        button.bind("<Enter>", lambda _event, target=button: self._set_button_hover(target, True))
        button.bind("<Leave>", lambda _event, target=button: self._set_button_hover(target, False))
        return button

    def _set_button_hover(self, button: tk.Button, hover: bool) -> None:
        if str(button.cget("state")) == "disabled":
            return
        if button is self.refresh_button:
            button.configure(bg="#F0F3F8" if hover else COLOR_CARD)
        else:
            button.configure(bg=COLOR_ACCENT_HOVER if hover else COLOR_ACCENT)

    def _set_primary_button_enabled(self, button: tk.Button, enabled: bool) -> None:
        if enabled:
            button.configure(
                state="normal",
                bg=COLOR_ACCENT,
                fg=COLOR_CARD,
                activebackground=COLOR_ACCENT_HOVER,
                cursor="hand2",
            )
        else:
            button.configure(
                state="disabled",
                bg=COLOR_DISABLED_BG,
                fg=COLOR_MUTED,
                activebackground=COLOR_DISABLED_BG,
                cursor="arrow",
            )

    def _set_secondary_button_enabled(self, button: tk.Button, enabled: bool) -> None:
        if enabled:
            button.configure(
                state="normal",
                bg=COLOR_CARD,
                fg=COLOR_MUTED,
                activebackground="#F0F3F8",
                cursor="hand2",
            )
        else:
            button.configure(
                state="disabled",
                bg=COLOR_DISABLED_BG,
                fg=COLOR_MUTED,
                activebackground=COLOR_DISABLED_BG,
                cursor="arrow",
            )

    def _setup_drop_target(self) -> None:
        if not DND_FILES or not hasattr(self.root, "drop_target_register"):
            return
        for target in (self.root, self.center_frame, self.empty_frame, self.cards_frame):
            target.drop_target_register(DND_FILES)
            target.dnd_bind("<<Drop>>", self._on_files_dropped)

    def _on_files_dropped(self, event: object) -> None:
        if self.busy:
            self._show_error(UI_TEXT["error_busy"])
            return
        data = getattr(event, "data", "")
        paths = list(self.root.tk.splitlist(data))
        self.empty_frame.configure(bg=COLOR_CARD, highlightbackground=COLOR_BORDER)
        self.add_paths(paths)

    def choose_files(self) -> None:
        if self.busy:
            self._show_error(UI_TEXT["error_busy"])
            return
        selected = filedialog.askopenfilenames(
            title=UI_TEXT["dialog_open_title"],
            filetypes=((UI_TEXT["filetype_pdf"], "*.pdf"), (UI_TEXT["filetype_all"], "*.*")),
        )
        if selected:
            self.add_paths(selected)

    def add_paths(self, raw_paths: Iterable[str]) -> None:
        if self.busy:
            self._show_error(UI_TEXT["error_busy"])
            return
        try:
            paths = normalize_pdf_paths(raw_paths)
        except LimitError as exc:
            self._show_error(build_error_message(exc.message, exc.file_name))
            return
        if not paths:
            return
        if len(self.items) + len(paths) > MAX_FILES:
            self._show_error(UI_TEXT["error_too_many_files"])
            return

        self._set_busy(True)
        self._set_status(UI_TEXT["status_loading"])
        existing_size = sum(item.size for item in self.items)
        existing_pages = sum(item.pages for item in self.items)
        threading.Thread(
            target=self._load_files_worker,
            args=(paths, existing_size, existing_pages),
            daemon=True,
        ).start()

    def _load_files_worker(self, paths: list[Path], existing_size: int, existing_pages: int) -> None:
        try:
            loaded: list[tuple[Path, int, int, bytes, int, int]] = []
            total_size = existing_size
            total_pages = existing_pages
            for path in paths:
                size, pages, thumb_data, width, height = read_pdf_info(path)
                total_size += size
                total_pages += pages
                if total_size > MAX_TOTAL_SIZE:
                    raise LimitError(UI_TEXT["error_total_too_large"], path.name)
                if total_pages > MAX_TOTAL_PAGES:
                    raise LimitError(UI_TEXT["error_total_too_many_pages"], path.name)
                loaded.append((path, size, pages, thumb_data, width, height))
            self.worker_queue.put(("loaded", loaded))
        except LimitError as exc:
            self.worker_queue.put(("error_policy", (exc.message, exc.file_name)))
        except Exception:
            self.worker_queue.put(("error_policy", (UI_TEXT["error_invalid_pdf"], None)))

    def ask_save_and_merge(self) -> None:
        if self.busy:
            self._show_error(UI_TEXT["error_busy"])
            return
        if not self.items:
            self._show_error(UI_TEXT["error_no_pdf"])
            return
        output_path = filedialog.asksaveasfilename(
            title=UI_TEXT["dialog_save_title"],
            initialfile=build_default_output_name(self.items[0].path),
            defaultextension=".pdf",
            filetypes=((UI_TEXT["filetype_pdf"], "*.pdf"),),
            confirmoverwrite=False,
        )
        if not output_path:
            return
        try:
            output_path = get_unique_path(Path(output_path))
        except LimitError as exc:
            self._show_error(build_error_message(exc.message, exc.file_name))
            return
        self._set_busy(True)
        self._set_status(UI_TEXT["status_saving"])
        paths = [item.path for item in self.items]
        threading.Thread(target=self._merge_worker, args=(paths, output_path), daemon=True).start()

    def _merge_worker(self, paths: list[Path], output_path: Path) -> None:
        try:
            output_path = get_unique_path(output_path)
            writer = PdfWriter()
            total_size = 0
            total_pages = 0
            for path in paths:
                size = path.stat().st_size
                if size > MAX_FILE_SIZE:
                    raise LimitError(UI_TEXT["error_too_large_file"], path.name)
                total_size += size
                if total_size > MAX_TOTAL_SIZE:
                    raise LimitError(UI_TEXT["error_total_too_large"], path.name)

                reader = PdfReader(str(path))
                page_count = len(reader.pages)
                if page_count > MAX_PAGES_PER_FILE:
                    raise LimitError(UI_TEXT["error_too_many_pages"], path.name)
                total_pages += page_count
                if total_pages > MAX_TOTAL_PAGES:
                    raise LimitError(UI_TEXT["error_total_too_many_pages"], path.name)
                for page in reader.pages:
                    writer.add_page(page)

            with output_path.open("xb") as output_file:
                writer.write(output_file)
            self.worker_queue.put(("merged", output_path))
        except LimitError as exc:
            self.worker_queue.put(("error_policy", (exc.message, exc.file_name)))
        except FileExistsError:
            self.worker_queue.put(("error", UI_TEXT["error_unique_name_failed"]))
        except Exception:
            self.worker_queue.put(("error", UI_TEXT["error_merge_failed"]))

    def _handle_worker_queue(self) -> None:
        try:
            while True:
                kind, payload = self.worker_queue.get_nowait()
                if kind == "loaded":
                    loaded = payload
                    for path, size, pages, thumb_data, _width, _height in loaded:
                        thumbnail = tk.PhotoImage(data=thumb_data)
                        self.items.append(PdfItem(path=path, name=path.name, size=size, pages=pages, thumbnail=thumbnail))
                    self._render_cards()
                    self._set_status(UI_TEXT["status_count"].format(count=len(self.items)))
                    self._set_busy(False)
                elif kind == "merged":
                    output_path = payload
                    self._set_status(UI_TEXT["status_complete_with_name"].format(filename=output_path.name))
                    self._set_busy(False)
                    self._open_folder(output_path)
                    messagebox.showinfo(
                        UI_TEXT["complete_title"],
                        UI_TEXT["status_complete_with_name"].format(filename=output_path.name),
                        parent=self.root,
                    )
                elif kind == "error_policy":
                    message, file_name = payload
                    self._set_status(UI_TEXT["status_error"])
                    self._set_busy(False)
                    self._show_error(build_error_message(message, file_name, include_policy=True))
                elif kind == "error":
                    self._set_status(UI_TEXT["status_error"])
                    self._set_busy(False)
                    self._show_error(str(payload))
        except queue.Empty:
            pass
        self.root.after(80, self._handle_worker_queue)

    def _render_cards(self) -> None:
        self.empty_frame.pack_forget()
        self.cards_frame.pack(fill="both", expand=True)
        for child in self.cards_frame.winfo_children():
            child.destroy()
        self.cards.clear()

        if not self.items:
            self.cards_frame.pack_forget()
            self.empty_frame.configure(bg=COLOR_CARD, highlightbackground=COLOR_BORDER)
            self.empty_frame.pack(expand=True, pady=(0, 20))
            return

        self.cards_frame.grid_columnconfigure(tuple(range(MAX_FILES)), weight=1, uniform="cards")
        for index, item in enumerate(self.items):
            card = tk.Frame(
                self.cards_frame,
                bg=COLOR_CARD,
                highlightthickness=1,
                highlightbackground=COLOR_BORDER,
                width=130,
                height=244,
            )
            card.grid(row=0, column=index, padx=(0 if index == 0 else 10, 0), pady=10, sticky="n")
            card.grid_propagate(False)
            self.cards.append(card)

            remove_button = tk.Button(
                card,
                text=UI_TEXT["button_remove"],
                command=lambda idx=index: self.remove_item(idx),
                bg=COLOR_CARD,
                fg=COLOR_MUTED,
                activebackground="#F2F5FA",
                activeforeground=COLOR_TEXT,
                relief="flat",
                bd=0,
                highlightthickness=0,
                font=(self.font_family, 12, "bold"),
                cursor="hand2",
            )
            remove_button.place(x=100, y=6, width=22, height=22)

            image_label = tk.Label(card, image=item.thumbnail, bg=COLOR_BASE, width=THUMBNAIL_WIDTH, height=THUMBNAIL_HEIGHT)
            image_label.place(x=7, y=34)
            name_label = tk.Label(
                card,
                text=display_filename(item.name),
                bg=COLOR_CARD,
                fg=COLOR_TEXT,
                font=(self.font_family, 9, "bold"),
                wraplength=112,
                justify="center",
            )
            name_label.place(x=7, y=190, width=116, height=32)
            info_label = tk.Label(
                card,
                text=f"{UI_TEXT['page_count'].format(pages=item.pages)} / {format_size(item.size)}",
                bg=COLOR_CARD,
                fg=COLOR_MUTED,
                font=(self.font_family, 8),
            )
            info_label.place(x=7, y=222, width=116, height=18)

            for widget in (card, image_label, name_label, info_label):
                widget.bind("<ButtonPress-1>", lambda event, idx=index: self._start_card_drag(event, idx))
                widget.bind("<B1-Motion>", self._move_card_drag)
                widget.bind("<ButtonRelease-1>", self._finish_card_drag)

    def _start_card_drag(self, _event: tk.Event, index: int) -> None:
        if self.busy:
            return
        self.drag_index = index
        self.drag_source_index = index
        self.dragging = False
        self.cards[index].configure(highlightbackground=COLOR_SELECTED_BORDER, bg=COLOR_SELECTED_BG)
        self._set_status(UI_TEXT["status_dragging"])

    def _move_card_drag(self, _event: tk.Event) -> None:
        if self.drag_index is not None:
            self.dragging = True

    def _calculate_insert_index(self, pointer_x: int) -> int:
        if not self.cards:
            return 0
        for index, card in enumerate(self.cards):
            left = card.winfo_rootx()
            width = card.winfo_width()
            if pointer_x < left + (width / 2):
                return index
            if pointer_x <= left + width:
                return index + 1
        return len(self.items)

    def _move_item(self, from_index: int, insert_index: int) -> bool:
        if not 0 <= from_index < len(self.items):
            return False
        insert_index = max(0, min(insert_index, len(self.items)))
        adjusted_index = insert_index
        if from_index < adjusted_index:
            adjusted_index -= 1
        if adjusted_index == from_index:
            return False

        item = self.items.pop(from_index)
        adjusted_index = max(0, min(adjusted_index, len(self.items)))
        self.items.insert(adjusted_index, item)
        return True

    def _clear_drag_state(self, reset_cards: bool = True) -> None:
        self.drag_index = None
        self.drag_source_index = None
        self.dragging = False
        if reset_cards:
            for card in self.cards:
                card.configure(highlightbackground=COLOR_BORDER, bg=COLOR_CARD)

    def _finish_card_drag(self, event: tk.Event) -> None:
        if self.busy or self.drag_index is None:
            return
        source = self.drag_index
        was_dragging = self.dragging
        insert_index = self._calculate_insert_index(event.x_root)
        self._clear_drag_state()
        if not was_dragging:
            self._set_status(UI_TEXT["status_count"].format(count=len(self.items)) if self.items else UI_TEXT["status_idle"])
            return

        if self._move_item(source, insert_index):
            self._render_cards()
            self._set_status(UI_TEXT["status_reordered"])
        else:
            self._set_status(UI_TEXT["status_count"].format(count=len(self.items)))

    def remove_item(self, index: int) -> None:
        if self.busy:
            self._show_error(UI_TEXT["error_busy"])
            return
        if 0 <= index < len(self.items):
            self._clear_drag_state()
            self.items.pop(index)
            self._render_cards()
            if self.items:
                self._set_status(UI_TEXT["status_count"].format(count=len(self.items)))
            else:
                self._set_status(UI_TEXT["status_idle"])
            self._update_buttons()

    def refresh_files(self) -> None:
        if self.busy:
            self._show_error(UI_TEXT["error_busy"])
            return
        self.items.clear()
        self._clear_drag_state()
        self._drain_worker_queue()
        self._render_cards()
        self._set_status(UI_TEXT["status_idle"])
        self._update_buttons()

    def _drain_worker_queue(self) -> None:
        try:
            while True:
                self.worker_queue.get_nowait()
        except queue.Empty:
            return

    def _set_busy(self, busy: bool) -> None:
        if busy:
            self._clear_drag_state()
        self.busy = busy
        self._update_buttons()

    def _update_buttons(self) -> None:
        self._set_primary_button_enabled(self.add_button, not self.busy)
        self._set_secondary_button_enabled(self.refresh_button, not self.busy and bool(self.items))
        self._set_primary_button_enabled(self.merge_button, not self.busy and bool(self.items))

    def _set_status(self, text: str) -> None:
        if text == UI_TEXT["status_error"]:
            color = COLOR_ERROR
        elif text == UI_TEXT["status_complete"] or text.startswith(UI_TEXT["status_complete"]):
            color = COLOR_SUCCESS
        elif text in (
            UI_TEXT["status_loading"],
            UI_TEXT["status_processing"],
            UI_TEXT["status_saving"],
            UI_TEXT["status_dragging"],
            UI_TEXT["status_reordered"],
        ) or text.startswith(str(len(self.items))):
            color = COLOR_ACCENT
        else:
            color = COLOR_MUTED
        self.status_label.configure(text=text, fg=color)

    def _show_error(self, message: str) -> None:
        messagebox.showerror(UI_TEXT["error_title"], message, parent=self.root)

    def _open_folder(self, output_path: Path) -> None:
        try:
            os.startfile(str(output_path.parent))
        except Exception:
            pass


def create_root() -> tk.Tk:
    if TkinterDnD:
        return TkinterDnD.Tk()
    return tk.Tk()


def main() -> int:
    root = create_root()
    PdfMergeMiniApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
