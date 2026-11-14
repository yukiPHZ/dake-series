# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import queue
import re
import sys
import threading
import time
_MAIN_ENTRY_AT = time.perf_counter()
import traceback
import webbrowser
from dataclasses import dataclass
from pathlib import Path
from typing import Any
import tempfile

from thumbnail_worker import ThumbnailWorker
from tkinter import filedialog, font as tkfont, messagebox, ttk
import tkinter as tk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    BASE_WINDOW = TkinterDnD.Tk
    DND_READY = True
except Exception:
    DND_FILES = None
    BASE_WINDOW = tk.Tk
    DND_READY = False


APP_NAME = "DakePDF差し込み"
WINDOW_TITLE = "DakePDF差し込み"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "brand_series": "シンプルそれDAKEシリーズ",
    "header_subtitle": "止まらない、迷わない、すぐ終わる。",
    "main_title": "PDFを差し込む",
    "main_description": "元PDFの好きなページ間に、別のPDFを差し込みます。",
    "source_title": "元になるPDF",
    "insert_title": "差し込むPDF",
    "position_title": "差し込み位置",
    "thumbnail_size": "サムネイルサイズ",
    "button_add_source": "元PDFを追加",
    "button_add_insert": "差し込むPDFを追加",
    "button_select_folder": "保存先を選ぶ",
    "button_refresh": "リフレッシュ",
    "button_cancel": "キャンセル",
    "button_execute": "差し込んで保存",
    "button_change": "変更",
    "button_position_after": "この後へ",
    "source_empty_title": "元PDFを追加してください",
    "insert_empty_title": "差し込むPDFを追加してください",
    "empty_title_drop": "PDFをドロップしてください",
    "empty_subtitle": "ドラッグ＆ドロップ または クリックして追加",
    "file_loaded": "{name}",
    "file_pages": "{pages}ページ",
    "file_loading": "読み込み中",
    "file_not_selected": "未選択",
    "position_before_first": "先頭に差し込む",
    "position_after_page": "{page}ページ目の後",
    "position_after_last": "末尾に差し込む",
    "position_selected": "選択中：{position}",
    "position_selected_none": "選択中：未選択",
    "position_plus_first": "＋先頭",
    "position_plus": "＋",
    "position_page_label": "{page}ページ",
    "position_loading_thumbnail": "読み込み中",
    "result_page_count": "完成後：全{pages}ページ",
    "result_page_count_empty": "完成後：-",
    "save_folder_unselected": "保存先：未選択",
    "save_folder_value": "保存先：{path}",
    "status_idle": "未選択",
    "status_loading": "読み込み中",
    "status_ready": "準備完了",
    "status_processing": "差し込み中",
    "status_saving": "保存中",
    "status_complete": "保存完了",
    "status_error": "エラー",
    "status_processing_base": "処理中",
    "status_processing_dots": [
        "処理中.",
        "処理中..",
        "処理中...",
    ],
    "status_phrase_1": "Simple",
    "status_phrase_2": "Simple, fast",
    "status_phrase_3": "Simple, fast, for real work.",
    "detail_idle": "元PDFと差し込むPDFを選んでください。",
    "detail_loading": "PDFを読み込んでいます。",
    "detail_thumbnail_loading": "元PDFのサムネイルを準備しています。",
    "detail_choose_insert_pdf": "差し込むPDFを選んでください。",
    "detail_choose_source_pdf": "元PDFを選んでください。",
    "detail_choose_position": "差し込み位置を選んでください。",
    "detail_ready": "保存できます。",
    "detail_processing": "ページ順を組み立てています。",
    "detail_saving": "新しいPDFを書き出しています。",
    "detail_complete": "保存フォルダを開きます。",
    "detail_error": "内容を確認してください。",
    "dialog_pdf_filter": "PDFファイル",
    "dialog_all_filter": "すべてのファイル",
    "dialog_source_title": "元PDFを選択",
    "dialog_insert_title": "差し込むPDFを選択",
    "dialog_folder_title": "保存先を選択",
    "dialog_complete_title": "保存が完了しました",
    "dialog_complete_message": "PDFを差し込んで保存しました。\n\n保存先：\n{path}",
    "dialog_warning_title": "確認してください",
    "dialog_error_title": "エラー",
    "error_not_pdf": "PDFファイルを指定してください。",
    "error_file_missing": "ファイルが見つかりませんでした。",
    "error_pdf_open": "PDFを読み込めませんでした。",
    "error_pdf_empty": "ページのないPDFは指定できません。",
    "error_pdf_protected": "パスワード保護されたPDFは、このアプリでは開けません。",
    "error_pdf_broken": "PDFの内容を確認できませんでした。",
    "error_source_missing": "元PDFを選んでください。",
    "error_insert_missing": "差し込むPDFを選んでください。",
    "error_position_missing": "差し込み位置を選んでください。",
    "error_same_file": "元PDFと差し込むPDFは、別のファイルを選んでください。",
    "error_save_folder_missing": "保存先フォルダを選んでください。",
    "error_save_folder_invalid": "保存先フォルダが見つかりませんでした。",
    "error_save_denied": "保存先に書き込めませんでした。",
    "error_save_failed": "保存できませんでした。",
    "error_thumbnail_resize": "サムネイルサイズを変更できませんでした。",
    "error_open_folder_failed": "保存先フォルダを開けませんでした。手動で確認してください。",
    "error_dependency": "必要なPDFライブラリを読み込めませんでした。",
    "error_drop_multiple": "PDFは1ファイルずつ、それぞれの追加欄へドロップしてください。",
    "thumbnail_size_value": "{size}px",
    "error_drop_failed": "ドロップされた内容を読み取れませんでした。",
    "error_detail": "{message}\n\n詳細：{detail}",
    "output_suffix": "_差し込み済み",
    "fallback_output_base": "差し込み済みPDF",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_subtitle": "止まらない、迷わない、すぐ終わる。",
    "footer_link_1": "戸建買取査定",
    "footer_link_2": "Instagram",
    "footer_separator": " ｜ ",
    "footer_copyright": COPYRIGHT,
    "window_geometry": "1080x780",
    "window_min_size": "920x680",
}

LINKS = {
    "assessment": "https://www.shimarisu-fudosan.com/",
    "instagram": "https://www.instagram.com/shimarisu.fudosan/",
}

THEME = {
    "background": "#F7F8FA",
    "surface": "#FFFFFF",
    "surface_alt": "#F1F3F5",
    "border": "#D9DEE5",
    "text": "#20242A",
    "muted": "#69727D",
    "soft": "#EEF1F5",
    "accent": "#2563EB",
    "accent_dark": "#1D4ED8",
    "selection_bg": "#EAF1FF",
    "success": "#168457",
    "danger": "#B42318",
}

THUMB_MIN_WIDTH = 110
THUMB_MAX_WIDTH = 260
THUMB_DEFAULT_WIDTH = 165
THUMB_RENDER_WIDTH = 360
THUMB_ASPECT = 1.42
POLL_INTERVAL_MS = 16
BUSY_INTERVAL_MS = 420
RESIZE_BATCH_SIZE = 8
RESIZE_DEBOUNCE_MS = 120
FOOTER_COMPACT_WIDTH = 1160


def _probe(event: str, at: float | None = None) -> None:
    target = os.environ.get("DAKE_INSERT_PROBE")
    if target:
        import json
        with open(target, "a", encoding="utf-8") as handle:
            handle.write(json.dumps({"event": event, "at": time.perf_counter() if at is None else at}) + "\n")


class PdfInsertError(Exception):
    pass


class SaveCancelled(Exception):
    pass


@dataclass(frozen=True)
class PdfInfo:
    path: Path
    pages: int


def shorten_text(value: str, max_chars: int = 52) -> str:
    if len(value) <= max_chars:
        return value
    head = max_chars // 2 - 2
    tail = max_chars - head - 3
    return f"{value[:head]}...{value[-tail:]}"


def sanitize_filename(value: str) -> str:
    sanitized = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value).strip(" .")
    if not sanitized:
        sanitized = UI_TEXT["fallback_output_base"]
    return sanitized[:90].rstrip(" .") or UI_TEXT["fallback_output_base"]


def icon_path() -> Path | None:
    meipass = getattr(sys, "_MEIPASS", None)
    candidates: list[Path] = []
    if meipass:
        candidates.append(Path(meipass) / "dake_icon.ico")
    base = Path(__file__).resolve().parent
    candidates.extend(
        [
            base / ".." / ".." / "02_assets" / "dake_icon.ico",
            base / ".." / ".." / ".." / "02_assets" / "dake_icon.ico",
            Path.cwd() / ".." / ".." / "02_assets" / "dake_icon.ico",
        ]
    )
    for candidate in candidates:
        try:
            resolved = candidate.resolve()
        except OSError:
            continue
        if resolved.exists():
            return resolved
    return None


def apply_window_icon(window: tk.Tk) -> None:
    ico = icon_path()
    if ico is None:
        return
    try:
        window.iconbitmap(str(ico))
    except Exception:
        pass


def choose_font_family() -> str:
    try:
        families = set(tkfont.families())
    except Exception:
        return "TkDefaultFont"
    for name in ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo"):
        if name in families:
            return name
    return "TkDefaultFont"


def parse_drop_paths(root: tk.Misc, data: str) -> list[Path]:
    parts = root.tk.splitlist(data)
    return [Path(part) for part in parts if part]

def validate_pdf_path(path: Path) -> None:
    if path.suffix.lower() != ".pdf":
        raise PdfInsertError(UI_TEXT["error_not_pdf"])
    if not path.exists():
        raise PdfInsertError(UI_TEXT["error_file_missing"])
    if not path.is_file():
        raise PdfInsertError(UI_TEXT["error_file_missing"])


def read_pdf_info(path: Path) -> PdfInfo:
    import fitz
    validate_pdf_path(path)
    try:
        with fitz.open(str(path)) as document:
            if document.needs_pass:
                raise PdfInsertError(UI_TEXT["error_pdf_protected"])
            count = document.page_count
        if count < 1:
            raise PdfInsertError(UI_TEXT["error_pdf_empty"])
        return PdfInfo(path=path, pages=count)
    except PdfInsertError:
        raise
    except Exception as exc:
        raise PdfInsertError(UI_TEXT["error_pdf_open"]) from exc

def make_output_path(source_pdf: Path, output_dir: Path) -> Path:
    base_name = sanitize_filename(source_pdf.stem)
    candidate = output_dir / f"{base_name}{UI_TEXT['output_suffix']}.pdf"
    source_resolved = source_pdf.resolve()
    index = 2
    while candidate.exists() or candidate.resolve() == source_resolved:
        candidate = output_dir / f"{base_name}{UI_TEXT['output_suffix']}_{index}.pdf"
        index += 1
    return candidate


def insert_pdf_pages(source_pdf: Path, insert_pdf: Path, output_pdf: Path, insert_after_pages: int, cancel_event: threading.Event | None = None, on_saving: Any = None) -> int:
    from pypdf import PdfReader, PdfWriter
    def check_cancelled():
        if cancel_event is not None and cancel_event.is_set():
            raise SaveCancelled()
    check_cancelled()
    if source_pdf.resolve() == insert_pdf.resolve():
        raise PdfInsertError(UI_TEXT["error_same_file"])
    if output_pdf.resolve() in {source_pdf.resolve(), insert_pdf.resolve()}:
        raise PdfInsertError(UI_TEXT["error_same_file"])
    temporary_path = None
    try:
        with source_pdf.open("rb") as source_handle, insert_pdf.open("rb") as insert_handle:
            source_reader = PdfReader(source_handle)
            insert_reader = PdfReader(insert_handle)
            if source_reader.is_encrypted or insert_reader.is_encrypted:
                raise PdfInsertError(UI_TEXT["error_pdf_protected"])
            source_count = len(source_reader.pages)
            insert_count = len(insert_reader.pages)
            if source_count <= 0 or insert_count <= 0:
                raise PdfInsertError(UI_TEXT["error_pdf_empty"])
            if not 0 <= insert_after_pages <= source_count:
                raise PdfInsertError(UI_TEXT["error_position_missing"])
            writer = PdfWriter()
            try:
                for page_index in range(insert_after_pages):
                    check_cancelled()
                    writer.add_page(source_reader.pages[page_index])
                for page in insert_reader.pages:
                    check_cancelled()
                    writer.add_page(page)
                for page_index in range(insert_after_pages, source_count):
                    check_cancelled()
                    writer.add_page(source_reader.pages[page_index])
                output_pdf.parent.mkdir(parents=True, exist_ok=True)
                with tempfile.NamedTemporaryFile(prefix=".dake-insert-", suffix=".tmp", dir=output_pdf.parent, delete=False) as handle:
                    temporary_path = Path(handle.name)
                    check_cancelled()
                    if on_saving is not None:
                        on_saving()
                    writer.write(handle)
                    handle.flush()
                    os.fsync(handle.fileno())
            finally:
                writer.close()
        total = source_count + insert_count
        with temporary_path.open("rb") as handle:
            if len(PdfReader(handle).pages) != total:
                raise PdfInsertError(UI_TEXT["error_save_failed"])
        check_cancelled()
        # Windows rename publishes only a complete PDF and refuses an existing destination.
        os.rename(temporary_path, output_pdf)
        temporary_path = None
        return total
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)

def open_folder(path: Path) -> bool:
    try:
        os.startfile(str(path))
        return True
    except Exception:
        return False


def save_pdf_worker(results: queue.Queue, cancelled: threading.Event, source_path: Path, insert_path: Path, output_dir: Path, position: int) -> None:
    try:
        if not output_dir.is_dir():
            raise FileNotFoundError(str(output_dir))
        output_path = make_output_path(source_path, output_dir)
        def saving():
            while not cancelled.is_set():
                try:
                    results.put(("save_status", "status_saving", UI_TEXT["detail_saving"]), timeout=.025)
                    return
                except queue.Full:
                    continue
            raise SaveCancelled()
        total = insert_pdf_pages(source_path, insert_path, output_path, position, cancelled, saving)
        result = ("save_done", output_path, total)
    except SaveCancelled:
        return
    except Exception as exc:
        result = ("save_failed", exc)
    while not cancelled.is_set():
        try:
            results.put(result, timeout=.025)
            return
        except queue.Full:
            continue


def format_error(message: str, exc: BaseException | None = None) -> str:
    if exc is None or isinstance(exc, PdfInsertError):
        return message
    detail = "".join(traceback.format_exception_only(type(exc), exc)).strip()
    return UI_TEXT["error_detail"].format(message=message, detail=detail)


class DakePdfInsertApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.geometry(UI_TEXT["window_geometry"])
        min_width, min_height = UI_TEXT["window_min_size"].split("x")
        self.root.minsize(int(min_width), int(min_height))
        self.root.configure(bg=THEME["background"])
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        apply_window_icon(self.root)

        self.font_family = choose_font_family()
        self.queue: queue.Queue[tuple[Any, ...]] = queue.Queue(maxsize=32)
        self.thumbnail_worker: ThumbnailWorker | None = None
        self.source_load_id = 0
        self.insert_load_id = 0
        self.source_info: PdfInfo | None = None
        self.insert_info: PdfInfo | None = None
        self.source_loading = False
        self.insert_loading = False
        self.thumbnail_loading = False
        self.thumbnail_ready = False
        self.is_busy = False
        self._save_cancel = threading.Event()
        self.selected_position: int | None = None
        self.save_dir: Path | None = None
        self.save_dir_manual = False

        self.thumbnail_photos: dict[int, Any] = {}
        self._photo_sizes: dict[int, tuple[int, int]] = {}
        self._photo_items: dict[int, int] = {}
        self._placeholder_items: dict[int, int] = {}
        self._visible_pages: tuple[int, ...] = ()
        self._nearby_pages: tuple[int, ...] = ()
        self._view_job: str | None = None
        self._view_pending = False
        self._view_signature = None
        self._image_height = int(THUMB_DEFAULT_WIDTH * THUMB_ASPECT)
        self._wheel_remainder = 0.0
        self.thumbnail_display_size = THUMB_DEFAULT_WIDTH
        self.requested_thumbnail_size = THUMB_DEFAULT_WIDTH
        self.resize_request_id = 0
        self.resize_debounce_job: str | None = None
        self.queue_poll_job: str | None = None
        self.busy_animation_job: str | None = None
        self.footer_compact: bool | None = None
        self._closing = False
        self._pending_paths: dict[str, Path] = {}
        self._load_started_at: dict[str, float] = {}
        self._diagnostics: dict[str, float] = {}
        self._profile = bool(os.environ.get("DAKE_INSERT_PROFILE"))

        self.busy_key: str | None = None
        self.busy_dot_index = 0
        self.busy_started_at = 0.0

        self.source_name_var = tk.StringVar(value=UI_TEXT["source_empty_title"])
        self.source_pages_var = tk.StringVar(value=UI_TEXT["empty_subtitle"])
        self.insert_name_var = tk.StringVar(value=UI_TEXT["insert_empty_title"])
        self.insert_pages_var = tk.StringVar(value=UI_TEXT["empty_subtitle"])
        self.position_var = tk.StringVar(value=UI_TEXT["position_selected_none"])
        self.result_pages_var = tk.StringVar(value=UI_TEXT["result_page_count_empty"])
        self.save_folder_var = tk.StringVar(value=UI_TEXT["save_folder_unselected"])
        self.status_var = tk.StringVar(value=UI_TEXT["status_idle"])
        self.detail_var = tk.StringVar(value=UI_TEXT["detail_idle"])
        self.thumbnail_size_var = tk.DoubleVar(value=THUMB_DEFAULT_WIDTH)
        self.thumbnail_size_text = tk.StringVar(value=UI_TEXT["thumbnail_size_value"].format(size=THUMB_DEFAULT_WIDTH))

        self._configure_styles()
        self._build_ui()
        self._update_footer_layout(self.root.winfo_width())
        self._update_action_state()
        self._setup_dnd()
        self.queue_poll_job = self.root.after(POLL_INTERVAL_MS, self._poll_queue)


    def _configure_styles(self) -> None:
        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")
        except Exception:
            pass
        self.style.configure(
            "Dake.Horizontal.TScale",
            background=THEME["surface"],
            troughcolor=THEME["surface_alt"],
            bordercolor=THEME["border"],
            lightcolor=THEME["border"],
            darkcolor=THEME["border"],
            sliderlength=14,
            borderwidth=0,
        )

    def font(self, size: int, weight: str = "normal") -> tuple[str, int, str]:
        return (self.font_family, size, weight)

    def _build_ui(self) -> None:
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(0, weight=1)

        shell = tk.Frame(self.root, bg=THEME["background"])
        shell.grid(row=0, column=0, sticky="nsew", padx=24, pady=(22, 10))
        shell.grid_columnconfigure(0, weight=1)
        shell.grid_rowconfigure(2, weight=1)

        self._build_header(shell)
        self._build_pdf_inputs(shell)
        self._build_position_area(shell)
        self._build_bottom_bar(shell)
        self._build_footer()

    def _build_header(self, parent: tk.Frame) -> None:
        header = tk.Frame(parent, bg=THEME["background"])
        header.grid(row=0, column=0, sticky="ew", pady=(0, 16))
        header.grid_columnconfigure(0, weight=1)

        tk.Label(
            header,
            text=UI_TEXT["main_title"],
            bg=THEME["background"],
            fg=THEME["text"],
            font=self.font(22, "bold"),
            anchor="w",
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            header,
            text=UI_TEXT["main_description"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=self.font(10),
            anchor="w",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

    def _build_pdf_inputs(self, parent: tk.Frame) -> None:
        inputs = tk.Frame(parent, bg=THEME["background"])
        inputs.grid(row=1, column=0, sticky="ew", pady=(0, 14))
        inputs.grid_columnconfigure(0, weight=1, uniform="pdf-input")
        inputs.grid_columnconfigure(1, weight=1, uniform="pdf-input")

        self.source_card = self._create_pdf_card(
            inputs,
            UI_TEXT["source_title"],
            self.source_name_var,
            self.source_pages_var,
            UI_TEXT["button_add_source"],
            lambda: self._choose_pdf("source"),
        )
        self.source_card.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        self.insert_card = self._create_pdf_card(
            inputs,
            UI_TEXT["insert_title"],
            self.insert_name_var,
            self.insert_pages_var,
            UI_TEXT["button_add_insert"],
            lambda: self._choose_pdf("insert"),
        )
        self.insert_card.grid(row=0, column=1, sticky="ew", padx=(8, 0))

    def _create_pdf_card(
        self,
        parent: tk.Frame,
        title: str,
        name_var: tk.StringVar,
        pages_var: tk.StringVar,
        button_text: str,
        command: Any,
    ) -> tk.Frame:
        card = tk.Frame(
            parent,
            bg=THEME["surface"],
            highlightthickness=1,
            highlightbackground=THEME["border"],
            padx=16,
            pady=14,
        )
        card.grid_columnconfigure(0, weight=1)
        self._bind_card_children = getattr(self, "_bind_card_children", [])
        tk.Label(
            card,
            text=title,
            bg=THEME["surface"],
            fg=THEME["text"],
            font=self.font(11, "bold"),
            anchor="w",
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            card,
            textvariable=name_var,
            width=1,
            bg=THEME["surface"],
            fg=THEME["text"],
            font=self.font(10, "bold"),
            anchor="w",
        ).grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        tk.Label(
            card,
            textvariable=pages_var,
            bg=THEME["surface"],
            fg=THEME["muted"],
            font=self.font(9),
            anchor="w",
        ).grid(row=2, column=0, sticky="ew", pady=(3, 0))
        button = tk.Button(
            card,
            text=button_text,
            command=command,
            bg=THEME["surface_alt"],
            fg=THEME["text"],
            activebackground=THEME["selection_bg"],
            activeforeground=THEME["accent"],
            relief="flat",
            padx=14,
            pady=7,
            font=self.font(9, "bold"),
            cursor="hand2",
        )
        button.grid(row=0, column=1, sticky="e", padx=(12, 0))
        self._bind_card_children.append((card, command))
        for widget in (card,):
            widget.bind("<Button-1>", lambda _event, action=command: action())
        return card

    def _build_position_area(self, parent: tk.Frame) -> None:
        area = tk.Frame(
            parent,
            bg=THEME["surface"],
            highlightthickness=1,
            highlightbackground=THEME["border"],
            padx=16,
            pady=14,
        )
        area.grid(row=2, column=0, sticky="nsew")
        area.grid_columnconfigure(0, weight=1)
        area.grid_rowconfigure(2, weight=1)

        top = tk.Frame(area, bg=THEME["surface"])
        top.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        top.grid_columnconfigure(0, weight=1)

        tk.Label(
            top,
            text=UI_TEXT["position_title"],
            bg=THEME["surface"],
            fg=THEME["text"],
            font=self.font(12, "bold"),
            anchor="w",
        ).grid(row=0, column=0, sticky="w")

        slider_frame = tk.Frame(top, bg=THEME["surface"])
        slider_frame.grid(row=0, column=1, sticky="e")
        tk.Label(
            slider_frame,
            text=UI_TEXT["thumbnail_size"],
            bg=THEME["surface"],
            fg=THEME["muted"],
            font=self.font(9),
        ).pack(side="left", padx=(0, 8))
        self.thumbnail_scale = ttk.Scale(
            slider_frame,
            from_=THUMB_MIN_WIDTH,
            to=THUMB_MAX_WIDTH,
            orient="horizontal",
            variable=self.thumbnail_size_var,
            command=self._on_thumbnail_scale,
            style="Dake.Horizontal.TScale",
            length=180,
            state="disabled",
        )
        self.thumbnail_scale.pack(side="left")
        tk.Label(slider_frame, textvariable=self.thumbnail_size_text, width=5, bg=THEME["surface"], fg=THEME["muted"], font=self.font(8)).pack(side="left", padx=(8, 0))

        summary = tk.Frame(area, bg=THEME["surface"])
        summary.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        summary.grid_columnconfigure(0, weight=1)
        tk.Label(
            summary,
            textvariable=self.position_var,
            bg=THEME["selection_bg"],
            fg=THEME["accent_dark"],
            font=self.font(10, "bold"),
            anchor="w",
            padx=12,
            pady=7,
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            summary,
            textvariable=self.result_pages_var,
            bg=THEME["surface"],
            fg=THEME["muted"],
            font=self.font(10, "bold"),
            anchor="e",
        ).grid(row=0, column=1, sticky="e")

        canvas_frame = tk.Frame(area, bg=THEME["surface"])
        canvas_frame.grid(row=2, column=0, sticky="nsew")
        canvas_frame.grid_columnconfigure(0, weight=1)
        canvas_frame.grid_rowconfigure(0, weight=1)

        self.thumbnail_canvas = tk.Canvas(
            canvas_frame,
            bg=THEME["surface"],
            highlightthickness=0,
            height=250,
            xscrollincrement=1,
        )
        self.thumbnail_canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(canvas_frame, orient="horizontal", command=self._on_scrollbar)
        self.thumbnail_scrollbar = scrollbar
        scrollbar.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        self.thumbnail_canvas.configure(xscrollcommand=scrollbar.set)

        self.thumbnail_canvas.bind("<Configure>", self._on_view_configure)
        self.thumbnail_canvas.bind("<MouseWheel>", self._scroll_thumbnails)
        self.thumbnail_canvas.bind("<Shift-MouseWheel>", self._scroll_thumbnails)
        self.thumbnail_canvas.bind("<Button-1>", self._on_thumbnail_click)
        self._rebuild_thumbnail_area()

    def _build_bottom_bar(self, parent: tk.Frame) -> None:
        bar = tk.Frame(parent, bg=THEME["background"])
        bar.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        bar.grid_columnconfigure(0, weight=1)
        self.save_folder_label = tk.Label(bar, textvariable=self.save_folder_var, bg=THEME["background"], fg=THEME["muted"], font=self.font(9), anchor="w", width=1)
        self.save_folder_label.grid(row=0, column=0, sticky="ew", padx=(0, 12))
        actions = tk.Frame(bar, bg=THEME["background"])
        actions.grid(row=0, column=1, sticky="e")
        self.folder_button = self._make_button(actions, UI_TEXT["button_select_folder"], self._choose_save_folder, primary=False)
        self.folder_button.pack(side="left", padx=(0, 8))
        self.execute_button = self._make_button(actions, UI_TEXT["button_execute"], self._execute_insert, primary=True)
        self.execute_button.pack(side="left")
        status = tk.Frame(bar, bg=THEME["background"])
        status.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        status.grid_columnconfigure(1, weight=1)
        self.status_badge = tk.Label(status, textvariable=self.status_var, bg=THEME["soft"], fg=THEME["muted"], font=self.font(9, "bold"), padx=10, pady=5)
        self.status_badge.grid(row=0, column=0, sticky="w")
        self.detail_label = tk.Label(status, textvariable=self.detail_var, bg=THEME["background"], fg=THEME["muted"], font=self.font(9), anchor="w", justify="left", width=1)
        self.detail_label.grid(row=0, column=1, sticky="ew", padx=(10, 0))

    def _build_footer(self) -> None:
        self.footer = tk.Frame(self.root, bg=THEME["background"])
        self.footer.grid(row=1, column=0, sticky="ew", padx=24, pady=(0, 14))
        self.footer.grid_columnconfigure(0, weight=1)
        self.footer.grid_columnconfigure(1, weight=1)

        self.footer_left = tk.Frame(self.footer, bg=THEME["background"])
        self.footer_left.grid(row=0, column=0, sticky="w")
        tk.Label(
            self.footer_left,
            text=UI_TEXT["footer_left"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=self.font(8),
        ).pack(side="left")
        tk.Label(
            self.footer_left,
            text=UI_TEXT["footer_separator"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=self.font(8),
        ).pack(side="left")
        tk.Label(
            self.footer_left,
            text=UI_TEXT["footer_subtitle"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=self.font(8),
        ).pack(side="left")

        self.footer_right = tk.Frame(self.footer, bg=THEME["background"])
        self.footer_right.grid(row=0, column=1, sticky="e")
        self._make_footer_link(self.footer_right, UI_TEXT["footer_link_1"], LINKS["assessment"]).pack(side="left")
        self._make_footer_text(self.footer_right, UI_TEXT["footer_separator"]).pack(side="left")
        self._make_footer_link(self.footer_right, UI_TEXT["footer_link_2"], LINKS["instagram"]).pack(side="left")
        self._make_footer_text(self.footer_right, UI_TEXT["footer_separator"]).pack(side="left")
        self._make_footer_text(self.footer_right, UI_TEXT["footer_copyright"]).pack(side="left")
        self.root.bind("<Configure>", self._on_root_configure, add="+")

    def _make_button(self, parent: tk.Misc, text: str, command: Any, primary: bool) -> tk.Button:
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=THEME["accent"] if primary else THEME["surface"],
            fg="#FFFFFF" if primary else THEME["text"],
            activebackground=THEME["accent_dark"] if primary else THEME["selection_bg"],
            activeforeground="#FFFFFF" if primary else THEME["accent"],
            disabledforeground="#A0A8B2",
            relief="flat",
            padx=16,
            pady=9,
            font=self.font(10, "bold"),
            cursor="hand2",
        )

    def _make_footer_text(self, parent: tk.Misc, text: str) -> tk.Label:
        return tk.Label(
            parent,
            text=text,
            bg=THEME["background"],
            fg=THEME["muted"],
            font=self.font(8),
        )

    def _make_footer_link(self, parent: tk.Misc, text: str, url: str) -> tk.Label:
        label = tk.Label(
            parent,
            text=text,
            bg=THEME["background"],
            fg=THEME["accent"],
            font=self.font(8),
            cursor="hand2",
        )
        label.bind("<Button-1>", lambda _event: webbrowser.open_new_tab(url))
        return label

    def _on_root_configure(self, event: tk.Event) -> None:
        if event.widget == self.root:
            self._update_footer_layout(event.width)

    def _update_footer_layout(self, width: int) -> None:
        required = self.footer_left.winfo_reqwidth() + self.footer_right.winfo_reqwidth() + 80
        compact = width < required
        if compact == self.footer_compact:
            return
        self.footer_compact = compact
        if compact:
            self.footer_left.grid_configure(row=0, column=0, columnspan=2, sticky="")
            self.footer_right.grid_configure(row=1, column=0, columnspan=2, sticky="", pady=(4, 0))
        else:
            self.footer_left.grid_configure(row=0, column=0, columnspan=1, sticky="w")
            self.footer_right.grid_configure(row=0, column=1, columnspan=1, sticky="e", pady=0)

    def _setup_dnd(self) -> None:
        for card, command in self._bind_card_children:
            for child in card.winfo_children():
                if not isinstance(child, tk.Button):
                    child.bind("<Button-1>", lambda _event, action=command: action())
        if not DND_READY or DND_FILES is None:
            return
        targets = [(self.thumbnail_canvas, "source")]
        for card, kind in ((self.source_card, "source"), (self.insert_card, "insert")):
            targets.append((card, kind))
            targets.extend((child, kind) for child in card.winfo_children())
        for widget, kind in targets:
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", lambda event, target=kind: self._on_drop(event, target))

    def _on_view_configure(self, _event: tk.Event) -> None:
        self._schedule_view()

    def _on_scrollbar(self, *args) -> None:
        self.thumbnail_canvas.xview(*args)
        self._schedule_view()

    def _scroll_thumbnails(self, event: tk.Event) -> str:
        self._wheel_remainder += -event.delta / 120 * 60
        pixels = int(self._wheel_remainder)
        self._wheel_remainder -= pixels
        if pixels:
            self.thumbnail_canvas.xview_scroll(pixels, "units")
            self._schedule_view()
        return "break"

    def _choose_pdf(self, kind: str) -> None:
        title_key = "dialog_source_title" if kind == "source" else "dialog_insert_title"
        path = filedialog.askopenfilename(
            title=UI_TEXT[title_key],
            filetypes=[
                (UI_TEXT["dialog_pdf_filter"], "*.pdf"),
                (UI_TEXT["dialog_all_filter"], "*.*"),
            ],
        )
        if path:
            self._load_pdf(kind, Path(path))

    def _choose_save_folder(self) -> None:
        initial = str(self.save_dir) if self.save_dir else str(Path.home())
        path = filedialog.askdirectory(title=UI_TEXT["dialog_folder_title"], initialdir=initial)
        if not path:
            return
        self.save_dir = Path(path)
        self.save_dir_manual = True
        self._refresh_save_folder()
        self._update_readiness_status()

    def _on_drop(self, event: tk.Event, kind: str | None) -> str:
        try:
            paths = parse_drop_paths(self.root, str(event.data))
        except tk.TclError:
            self._set_status("status_error", UI_TEXT["error_drop_failed"])
            return "break"
        if len(paths) != 1:
            key = "error_drop_multiple" if paths else "error_drop_failed"
            self._set_status("status_error", UI_TEXT[key])
            return "break"
        self._load_pdf(kind or "source", paths[0])
        return "break"

    def _load_pdf(self, kind: str, path: Path) -> None:
        if self.is_busy or self._closing:
            return
        if path.suffix.lower() != ".pdf":
            self._set_status("status_error", UI_TEXT["error_not_pdf"])
            return
        other_kind = "insert" if kind == "source" else "source"
        other = self._pending_paths.get(other_kind)
        if other is not None and os.path.normcase(os.path.abspath(path)) == os.path.normcase(os.path.abspath(other)):
            self._set_status("status_error", UI_TEXT["error_same_file"])
            return
        self._pending_paths[kind] = path
        if self._profile:
            self._load_started_at[kind] = time.perf_counter()
        if self.thumbnail_worker is None:
            self.thumbnail_worker = ThumbnailWorker(self.queue)
        if kind == "source":
            self.source_load_id += 1
            load_id = self.source_load_id
            self._cancel_thumbnail_resize()
            self.source_loading = True
            self.thumbnail_loading = False
            self.thumbnail_ready = False
            self.source_info = None
            self.selected_position = None
            self.thumbnail_photos.clear()
            self._photo_sizes.clear()
            self._view_signature = None
            self._diagnostics.clear()
            self._rebuild_thumbnail_area()
            self.source_name_var.set(shorten_text(path.name))
            self.source_pages_var.set(UI_TEXT["file_loading"])
        else:
            self.insert_load_id += 1
            load_id = self.insert_load_id
            self.insert_loading = True
            self.insert_info = None
            self.insert_name_var.set(shorten_text(path.name))
            self.insert_pages_var.set(UI_TEXT["file_loading"])
        self._set_status("status_loading", UI_TEXT["detail_loading"])
        self._update_summary()
        self._update_action_state()
        self.thumbnail_worker.load(kind, load_id, path)

    def _poll_queue(self) -> None:
        self.queue_poll_job = None
        if self._closing:
            return
        deadline = time.perf_counter() + .008
        while time.perf_counter() < deadline:
            try:
                message = self.queue.get_nowait()
            except queue.Empty:
                break
            self._handle_queue_message(message)
        if not self._closing:
            self.queue_poll_job = self.root.after(POLL_INTERVAL_MS, self._poll_queue)

    def _handle_queue_message(self, message: tuple[Any, ...]) -> None:
        if self._closing:
            return
        kind = message[0]
        if kind == "pdf_metadata":
            _, pdf_kind, load_id, path, count = message
            if not self._is_current_load(pdf_kind, load_id):
                return
            info = PdfInfo(path, count)
            self._record_load_timing("page_count_s", pdf_kind)
            if pdf_kind == "source":
                self.source_info = info
                self.source_loading = False
                self.thumbnail_loading = True
                if not self.save_dir or not self.save_dir_manual:
                    self.save_dir = info.path.parent
                    self.save_dir_manual = False
                    self._refresh_save_folder()
                self._set_source_loaded_text(info)
                self._rebuild_thumbnail_area()
            else:
                self.insert_info = info
                self.insert_loading = False
                self._set_insert_loaded_text(info)
            self._update_summary()
            self._update_readiness_status()
            return
        if kind in {"pdf_load_failed", "pdf_rejected"}:
            _, pdf_kind, load_id, error = message
            if not self._is_current_load(pdf_kind, load_id):
                return
            self._pending_paths.pop(pdf_kind, None)
            if pdf_kind == "source":
                self.source_loading = self.thumbnail_loading = self.thumbnail_ready = False
                self.source_info = None
                self._set_source_empty_text()
                self._rebuild_thumbnail_area()
            else:
                self.insert_loading = False
                self.insert_info = None
                self._set_insert_empty_text()
            detail = UI_TEXT[error] if kind == "pdf_rejected" else self._load_error_message(error)
            self._set_status("status_error", detail)
            self._update_summary()
            self._update_action_state()
            return
        if kind == "view_thumbnail":
            _, load_id, request_id, page_index, width, height, image = message
            if load_id != self.source_load_id or request_id != self.resize_request_id:
                return
            self._draw_thumbnail(page_index, image)
            self._photo_sizes[page_index] = (width, height)
            self._record_load_timing("first_thumbnail_s", "source")
            self._check_visible_ready()
            return
        if kind in {"view_done", "view_failed"}:
            _, load_id, request_id, *detail = message
            if load_id != self.source_load_id or request_id != self.resize_request_id:
                return
            self._view_pending = False
            if kind == "view_failed":
                self.thumbnail_loading = False
                self._set_status("status_error", format_error(UI_TEXT["error_thumbnail_resize"], detail[0]))
                self._update_action_state()
            else:
                self._check_visible_ready()
            return
        if kind == "save_status":
            _, status_key, detail = message
            self._set_status(status_key, detail)
            return

        if kind == "save_done":
            _, output_path, total_pages = message
            self.is_busy = False
            self._set_status("status_complete", UI_TEXT["detail_complete"])
            self._update_action_state()
            messagebox.showinfo(
                UI_TEXT["dialog_complete_title"],
                UI_TEXT["dialog_complete_message"].format(path=str(output_path)),
                parent=self.root,
            )
            if not open_folder(Path(output_path).parent):
                self._set_status("status_error", UI_TEXT["error_open_folder_failed"])
            self.result_pages_var.set(UI_TEXT["result_page_count"].format(pages=total_pages))
            return

        if kind == "save_failed":
            _, exc = message
            self.is_busy = False
            error_message = self._save_error_message(exc)
            self._set_status("status_error", error_message)
            self._update_action_state()
            messagebox.showerror(UI_TEXT["dialog_error_title"], error_message, parent=self.root)
            return

    def _is_current_load(self, kind: str, load_id: int) -> bool:
        return (kind == "source" and load_id == self.source_load_id) or (
            kind == "insert" and load_id == self.insert_load_id
        )

    def _record_load_timing(self, name: str, kind: str) -> None:
        if self._profile and kind in self._load_started_at:
            self._diagnostics.setdefault(name, time.perf_counter() - self._load_started_at[kind])

    def _set_source_loading_text(self) -> None:
        self.source_name_var.set(UI_TEXT["file_loading"])
        self.source_pages_var.set(UI_TEXT["empty_subtitle"])

    def _set_insert_loading_text(self) -> None:
        self.insert_name_var.set(UI_TEXT["file_loading"])
        self.insert_pages_var.set(UI_TEXT["empty_subtitle"])

    def _set_source_loaded_text(self, info: PdfInfo) -> None:
        self.source_name_var.set(UI_TEXT["file_loaded"].format(name=shorten_text(info.path.name)))
        self.source_pages_var.set(UI_TEXT["file_pages"].format(pages=info.pages))

    def _set_insert_loaded_text(self, info: PdfInfo) -> None:
        self.insert_name_var.set(UI_TEXT["file_loaded"].format(name=shorten_text(info.path.name)))
        self.insert_pages_var.set(UI_TEXT["file_pages"].format(pages=info.pages))

    def _set_source_empty_text(self) -> None:
        self.source_name_var.set(UI_TEXT["source_empty_title"])
        self.source_pages_var.set(UI_TEXT["empty_subtitle"])

    def _set_insert_empty_text(self) -> None:
        self.insert_name_var.set(UI_TEXT["insert_empty_title"])
        self.insert_pages_var.set(UI_TEXT["empty_subtitle"])

    def _load_error_message(self, exc: BaseException) -> str:
        if isinstance(exc, PdfInsertError):
            return str(exc)
        if isinstance(exc, FileNotFoundError):
            return UI_TEXT["error_file_missing"]
        return format_error(UI_TEXT["error_pdf_broken"], exc)

    def _save_error_message(self, exc: BaseException) -> str:
        if isinstance(exc, PermissionError):
            return UI_TEXT["error_save_denied"]
        if isinstance(exc, FileNotFoundError):
            return UI_TEXT["error_save_folder_invalid"]
        if isinstance(exc, PdfInsertError):
            return str(exc)
        return format_error(UI_TEXT["error_save_failed"], exc)

    def _set_status(self, status_key: str, detail: str | None = None) -> None:
        self.busy_key = status_key if status_key in {"status_loading", "status_processing", "status_saving"} else None
        if self.busy_key:
            self.busy_started_at = time.monotonic()
            self.busy_dot_index = 0
        self.status_var.set(UI_TEXT[status_key])
        if detail is not None:
            self.detail_var.set(detail)
        self._style_status_badge(status_key)
        if self.busy_key and self.busy_animation_job is None:
            self.busy_animation_job = self.root.after(BUSY_INTERVAL_MS, self._animate_busy_status)
        elif self.busy_key is None and self.busy_animation_job is not None:
            self.root.after_cancel(self.busy_animation_job)
            self.busy_animation_job = None

    def _style_status_badge(self, status_key: str) -> None:
        if status_key == "status_complete":
            self.status_badge.configure(bg="#EAFBF3", fg=THEME["success"])
        elif status_key == "status_error":
            self.status_badge.configure(bg="#FDECEC", fg=THEME["danger"])
        elif status_key in {"status_ready", "status_processing", "status_saving", "status_loading"}:
            self.status_badge.configure(bg=THEME["selection_bg"], fg=THEME["accent_dark"])
        else:
            self.status_badge.configure(bg=THEME["soft"], fg=THEME["muted"])

    def _animate_busy_status(self) -> None:
        self.busy_animation_job = None
        if self._closing or self.busy_key is None:
            return
        dots = UI_TEXT["status_processing_dots"]
        self.status_var.set(UI_TEXT[self.busy_key])
        self.detail_var.set(dots[self.busy_dot_index % len(dots)])
        self.busy_dot_index += 1
        self.busy_animation_job = self.root.after(BUSY_INTERVAL_MS, self._animate_busy_status)

    def _update_readiness_status(self) -> None:
        if self.is_busy:
            return
        if self.source_loading or self.insert_loading:
            self._set_status("status_loading", UI_TEXT["detail_loading"])
        elif self.thumbnail_loading:
            self._set_status("status_loading", UI_TEXT["detail_thumbnail_loading"])
        elif self.source_info is None:
            self._set_status("status_idle", UI_TEXT["detail_choose_source_pdf"])
        elif self.insert_info is None:
            self._set_status("status_ready", UI_TEXT["detail_choose_insert_pdf"])
        elif self.selected_position is None:
            self._set_status("status_idle", UI_TEXT["detail_choose_position"])
        else:
            self._set_status("status_ready", UI_TEXT["detail_ready"])
        self._update_action_state()

    def _update_action_state(self) -> None:
        can_execute = (
            self.source_info is not None
            and self.insert_info is not None
            and self.selected_position is not None
            and not self.is_busy
            and not self.source_loading
            and not self.insert_loading
            and self.thumbnail_ready
        )
        self.execute_button.configure(state="normal" if can_execute else "disabled")
        self.folder_button.configure(state="disabled" if self.is_busy else "normal")
        can_resize = (
            self.source_info is not None
            and self.thumbnail_ready
            and not self.source_loading
            and not self.thumbnail_loading
            and not self.is_busy
            and not self._closing
        )
        self.thumbnail_scale.configure(state="normal" if can_resize else "disabled")

    def _refresh_save_folder(self) -> None:
        if self.save_dir is None:
            self.save_folder_var.set(UI_TEXT["save_folder_unselected"])
            return
        self.save_folder_var.set(UI_TEXT["save_folder_value"].format(path=shorten_text(str(self.save_dir), 56)))

    def _select_position(self, position: int) -> None:
        if self.source_info is None or position < 0 or position > self.source_info.pages:
            return
        self.selected_position = position
        self._update_summary()
        self._draw_viewport()
        self._update_readiness_status()

    def _position_label(self, position: int) -> str:
        if position == 0:
            return UI_TEXT["position_before_first"]
        if self.source_info is not None and position == self.source_info.pages:
            return UI_TEXT["position_after_last"]
        return UI_TEXT["position_after_page"].format(page=position)

    def _update_summary(self) -> None:
        if self.selected_position is None:
            self.position_var.set(UI_TEXT["position_selected_none"])
        else:
            self.position_var.set(
                UI_TEXT["position_selected"].format(position=self._position_label(self.selected_position))
            )
        if self.source_info is not None and self.insert_info is not None:
            total = self.source_info.pages + self.insert_info.pages
            self.result_pages_var.set(UI_TEXT["result_page_count"].format(pages=total))
        else:
            self.result_pages_var.set(UI_TEXT["result_page_count_empty"])

    def _rebuild_thumbnail_area(self) -> None:
        self._view_signature = None
        self.thumbnail_canvas.xview_moveto(0)
        self._schedule_view()

    def _schedule_view(self) -> None:
        if not self._closing and self._view_job is None:
            self._view_job = self.root.after(16, self._update_view)

    def _update_view(self) -> None:
        self._view_job = None
        if self._closing:
            return
        canvas = self.thumbnail_canvas
        width = self.thumbnail_display_size
        self._image_height = min(int(width * THUMB_ASPECT), max(90, canvas.winfo_height() - 86))
        if self.source_info is None:
            canvas.configure(scrollregion=(0, 0, canvas.winfo_width(), canvas.winfo_height()))
            canvas.delete("all")
            self._photo_items.clear()
            canvas.create_text(canvas.winfo_width()/2, canvas.winfo_height()/2 - 14,
                               text=UI_TEXT["source_empty_title"], fill=THEME["muted"], font=self.font(13, "bold"))
            canvas.create_text(canvas.winfo_width()/2, canvas.winfo_height()/2 + 14,
                               text=UI_TEXT["empty_subtitle"], fill=THEME["muted"], font=self.font(10))
            return
        stride = width + 70
        total = 88 + self.source_info.pages * stride
        canvas.configure(scrollregion=(0, 0, total, canvas.winfo_height()))
        left = canvas.canvasx(0)
        right = left + canvas.winfo_width()
        first = max(0, int((left-88)//stride))
        last = min(self.source_info.pages-1, int((right-88)//stride))
        self._visible_pages = tuple(range(first, last+1))
        self._nearby_pages = tuple(range(max(0, first-2), min(self.source_info.pages, last+3)))
        keep = set(self._nearby_pages)
        for page in list(self.thumbnail_photos):
            if page not in keep:
                self.thumbnail_photos.pop(page)
                self._photo_sizes.pop(page, None)
        self._draw_viewport()
        signature = (self.source_load_id, width, self._image_height, self._nearby_pages)
        if signature != self._view_signature:
            self._view_signature = signature
            self.resize_request_id += 1
            missing = tuple(p for p in self._visible_pages + tuple(p for p in self._nearby_pages if p not in self._visible_pages)
                            if self._photo_sizes.get(p) != (width, self._image_height))
            self._view_pending = bool(missing)
            if missing and self.thumbnail_worker is not None:
                self.thumbnail_worker.render(self.source_load_id, self.resize_request_id,
                                             self.source_info.path, missing, width, self._image_height)
        self._check_visible_ready()

    def _draw_viewport(self) -> None:
        if self.source_info is None or self._closing:
            return
        canvas = self.thumbnail_canvas
        canvas.delete("all")
        self._photo_items.clear()
        self._placeholder_items.clear()
        width = self.thumbnail_display_size
        height = self._image_height
        stride = width + 70
        top = max(8, (canvas.winfo_height() - height - 64)//2)
        center_y = top + (height+64)/2
        selected = self.selected_position == 0
        canvas.create_rectangle(4, center_y-18, 74, center_y+18,
                                fill=THEME["accent"] if selected else THEME["surface_alt"], outline="")
        canvas.create_text(39, center_y, text=UI_TEXT["position_plus_first"],
                           fill="#FFFFFF" if selected else THEME["accent"], font=self.font(10, "bold"))
        for page_index in self._nearby_pages:
            x = 88 + page_index * stride
            selected = self.selected_position == page_index + 1
            fill = THEME["selection_bg"] if selected else THEME["surface"]
            canvas.create_rectangle(x, top, x+width+16, top+height+64,
                                    fill=fill, outline=THEME["accent"] if selected else THEME["border"], width=2 if selected else 1)
            photo = self.thumbnail_photos.get(page_index)
            if photo is not None:
                self._photo_items[page_index] = canvas.create_image(x+8+width/2, top+8+height/2, image=photo)
            else:
                self._placeholder_items[page_index] = canvas.create_text(x+8+width/2, top+8+height/2,
                    text=UI_TEXT["position_loading_thumbnail"], fill=THEME["muted"], font=self.font(9))
            canvas.create_text(x+8+width/2, top+height+23, text=UI_TEXT["position_page_label"].format(page=page_index+1),
                               fill=THEME["text"], font=self.font(9, "bold"))
            canvas.create_rectangle(x+8, top+height+36, x+width+8, top+height+58,
                                    fill=THEME["accent"] if selected else THEME["surface_alt"], outline="")
            canvas.create_text(x+8+width/2, top+height+47, text=UI_TEXT["button_position_after"],
                               fill="#FFFFFF" if selected else THEME["text"], font=self.font(8, "bold"))
            plus_x = x+width+42
            canvas.create_rectangle(plus_x-16, center_y-18, plus_x+16, center_y+18,
                                    fill=THEME["accent"] if selected else THEME["surface_alt"], outline="")
            canvas.create_text(plus_x, center_y, text=UI_TEXT["position_plus"],
                               fill="#FFFFFF" if selected else THEME["accent"], font=self.font(12))
        canvas.configure(cursor="hand2")

    def _draw_thumbnail(self, page_index: int, rendered: Any) -> None:
        from PIL import ImageTk
        # The only PhotoImage creation site runs in the main-thread queue poll.
        photo = ImageTk.PhotoImage(rendered, master=self.root)
        self.thumbnail_photos[page_index] = photo
        item = self._photo_items.get(page_index)
        if item is not None:
            self.thumbnail_canvas.itemconfigure(item, image=photo)
        elif page_index in self._nearby_pages:
            x = 88 + page_index * (self.thumbnail_display_size+70)
            top = max(8, (self.thumbnail_canvas.winfo_height()-self._image_height-64)//2)
            self._photo_items[page_index] = self.thumbnail_canvas.create_image(
                x+8+self.thumbnail_display_size/2, top+8+self._image_height/2, image=photo)
        placeholder = self._placeholder_items.pop(page_index, None)
        if placeholder is not None:
            self.thumbnail_canvas.delete(placeholder)

    def _check_visible_ready(self) -> None:
        size = (self.thumbnail_display_size, self._image_height)
        ready = bool(self._visible_pages) and all(self._photo_sizes.get(page) == size for page in self._visible_pages)
        if ready and not self.thumbnail_ready:
            self.thumbnail_ready = True
            self.thumbnail_loading = False
            self._record_load_timing("initial_ready_s", "source")
            self._update_readiness_status()

    def _on_thumbnail_click(self, event: tk.Event) -> None:
        if self.source_info is None:
            return
        x = self.thumbnail_canvas.canvasx(event.x)
        top = max(8, (self.thumbnail_canvas.winfo_height()-self._image_height-64)//2)
        center_y = top+(self._image_height+64)/2
        if x < 80:
            if 4 <= x <= 74 and abs(event.y-center_y) <= 18:
                self._select_position(0)
            return
        page_index = int((x-88)//(self.thumbnail_display_size+70))
        if 0 <= page_index < self.source_info.pages:
            local_x = x-88-page_index*(self.thumbnail_display_size+70)
            on_card = 0 <= local_x <= self.thumbnail_display_size+16 and top <= event.y <= top+self._image_height+64
            on_plus = abs(local_x-self.thumbnail_display_size-42) <= 16 and abs(event.y-center_y) <= 18
            if on_card or on_plus:
                self._select_position(page_index+1)

    def _on_thumbnail_scale(self, value: str) -> None:
        if self._closing or self.source_info is None or not self.thumbnail_ready or self.is_busy:
            return
        requested = max(THUMB_MIN_WIDTH, min(THUMB_MAX_WIDTH, round(float(value))))
        self.thumbnail_size_text.set(UI_TEXT["thumbnail_size_value"].format(size=requested))
        self.requested_thumbnail_size = requested
        self.resize_request_id += 1
        if self.thumbnail_worker is not None:
            self.thumbnail_worker.invalidate()
        if self.resize_debounce_job is not None:
            self.root.after_cancel(self.resize_debounce_job)
        self.resize_debounce_job = self.root.after(RESIZE_DEBOUNCE_MS, self._start_thumbnail_resize)

    def _start_thumbnail_resize(self) -> None:
        self.resize_debounce_job = None
        if self._closing or self.source_info is None:
            return
        old_stride = self.thumbnail_display_size + 70
        anchor = max(0, self.thumbnail_canvas.canvasx(0)-88) / old_stride
        self.thumbnail_display_size = self.requested_thumbnail_size
        stride = self.thumbnail_display_size+70
        total = 88 + self.source_info.pages * stride
        self.thumbnail_canvas.configure(scrollregion=(0, 0, total, self.thumbnail_canvas.winfo_height()))
        target = 0 if anchor == 0 else 88 + anchor*stride
        self.thumbnail_canvas.xview_moveto(target/total)
        self._view_signature = None
        self._schedule_view()

    def _cancel_thumbnail_resize(self) -> None:
        self.resize_request_id += 1
        if self.thumbnail_worker is not None:
            self.thumbnail_worker.invalidate()
        if self.resize_debounce_job is not None:
            self.root.after_cancel(self.resize_debounce_job)
            self.resize_debounce_job = None
        self._view_pending = False

    def _execute_insert(self) -> None:
        if self.source_info is None:
            self._set_status("status_error", UI_TEXT["error_source_missing"])
            return
        if self.insert_info is None:
            self._set_status("status_error", UI_TEXT["error_insert_missing"])
            return
        if self.selected_position is None:
            self._set_status("status_error", UI_TEXT["error_position_missing"])
            return
        if os.path.normcase(os.path.abspath(self.source_info.path)) == os.path.normcase(os.path.abspath(self.insert_info.path)):
            self._set_status("status_error", UI_TEXT["error_same_file"])
            return
        if self.save_dir is None:
            self._set_status("status_error", UI_TEXT["error_save_folder_missing"])
            return
        output_path = self.save_dir
        self.is_busy = True
        self._set_status("status_processing", UI_TEXT["detail_processing"])
        self._update_action_state()
        self._save_cancel.clear()
        threading.Thread(
            target=save_pdf_worker,
            args=(self.queue, self._save_cancel, self.source_info.path, self.insert_info.path, output_path, self.selected_position),
            daemon=False,
            name="pdf-save",
        ).start()

    def _on_close(self) -> None:
        if self._closing:
            return
        self._closing = True
        self._save_cancel.set()
        self.source_load_id += 1
        self.insert_load_id += 1
        self._cancel_thumbnail_resize()
        if self.thumbnail_worker is not None:
            self.thumbnail_worker.close()
        if self._view_job is not None:
            self.root.after_cancel(self._view_job)
            self._view_job = None
        for job_name in ("queue_poll_job", "busy_animation_job"):
            job = getattr(self, job_name)
            if job is not None:
                try:
                    self.root.after_cancel(job)
                except tk.TclError:
                    pass
                setattr(self, job_name, None)
        self.root.destroy()


def main() -> int:
    _probe("module_entry", _MAIN_ENTRY_AT)
    root = BASE_WINDOW()
    app = DakePdfInsertApp(root)
    mapped = [False]
    def first_map(event):
        if event.widget == root and not mapped[0]:
            mapped[0] = True
            _probe("first_map")
            root.after_idle(lambda: _probe("interactive"))
    root.bind("<Map>", first_map, add="+")
    if "--launch-check" in sys.argv:
        root.after(900, app._on_close)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
