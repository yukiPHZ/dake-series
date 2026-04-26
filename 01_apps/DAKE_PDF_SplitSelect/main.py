import os
import queue
import re
import subprocess
import sys
import threading
import webbrowser
from dataclasses import dataclass
from enum import Enum
from itertools import count
from pathlib import Path
from typing import Optional

import fitz
from PIL import Image, ImageTk
from pypdf import PdfReader, PdfWriter
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    BASE_WINDOW = TkinterDnD.Tk
    DND_READY = True
except ImportError:
    DND_FILES = ""
    BASE_WINDOW = tk.Tk
    DND_READY = False


APP_NAME = "PDF分割Select"
WINDOW_TITLE = "PDF分割Select"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "button_add_pdf": "PDF追加",
    "button_choose_save_dir": "保存先選択",
    "button_select_save_folder": "保存先選択",
    "button_refresh": "リフレッシュ",
    "button_clear_selection": "選択解除",
    "button_extract_merged": "抽出する",
    "button_extract_single": "1ページずつ出力",
    "button_extract": "抽出する",
    "button_split_each": "1ページずつ出力",
    "label_file_name": "ファイル名",
    "label_total_pages": "総ページ数",
    "label_range_input": "範囲入力",
    "label_range_example": "形式: 1-3,5,8-10",
    "label_selected_pages": "選択中 {count} ページ",
    "label_save_dir_default": "保存先: 未選択",
    "label_save_dir_value": "保存先: {path}",
    "label_file_default": "未選択",
    "label_total_pages_default": "0 ページ",
    "label_total_pages_value": "{count} ページ",
    "label_thumbnail_empty": "PDFを追加するか、ここへドロップしてください",
    "label_thumbnail_loading": "サムネイルを準備しています",
    "label_thumbnail_page": "P.{page}",
    "label_thumbnail_selected": "選択中",
    "label_drop_multiple_error": "PDFは1ファイルだけ指定してください。",
    "label_drop_file_error": "PDFファイルを指定してください。",
    "label_status_unloaded": "未読込",
    "label_status_loading": "読込中",
    "label_status_ready": "準備完了",
    "label_status_selecting": "選択中",
    "label_status_processing": "処理中...",
    "label_status_complete": "完了",
    "label_status_error": "エラー",
    "message_unloaded": "PDFを追加してください。",
    "message_loading": "PDFを読み込んでいます。",
    "message_ready": "見て選んで抜く準備ができました。",
    "message_selecting": "抽出するページを選んでください。",
    "message_processing_extract": "選択されたページを書き出しています。",
    "message_complete_merged": "{count} ページを1つのPDFにまとめて保存しました。",
    "message_complete_single": "{count} ページを1ページずつ保存しました。",
    "message_range_secondary": "サムネイル選択を優先しています。",
    "message_error_save_dir": "保存先フォルダを選んでください。",
    "message_error_selection": "抽出するページを選んでください。",
    "message_error_range_format": "範囲入力は 1-3,5,8-10 の形式で入力してください。",
    "message_error_range_reverse": "範囲の始まりは終わり以下で入力してください。",
    "message_error_range_zero": "ページ番号は 1 以上で入力してください。",
    "message_error_range_over": "指定されたページがPDFの総ページ数を超えています。",
    "message_error_open_folder": "保存フォルダを開けませんでした。保存先を手動でご確認ください。",
    "message_error_no_pdf": "先にPDFを追加してください。",
    "message_error_open_pdf_detail": "PDF読み込みに失敗しました。{reason}",
    "message_error_render_pdf_detail": "サムネイルの準備に失敗しました。{reason}",
    "message_error_processing_detail": "保存に失敗しました。{reason}",
    "message_error_drop_detail": "ドロップされた内容を読み取れませんでした。{reason}",
    "message_complete_title": "完了",
    "dialog_complete_with_path": "{summary}\n\n保存先:\n{path}",
    "file_dialog_title_pdf": "PDFを選択",
    "file_dialog_title_save_dir": "保存先を選択",
    "file_dialog_pdf_filter": "PDFファイル",
    "reason_file_busy": "ファイルにアクセスできません。開いたままの可能性があります。",
    "reason_file_not_found": "ファイルが見つかりません。",
    "reason_pdf_invalid": "PDFの内容を読み取れませんでした。",
    "reason_pdf_protected": "保護されたPDFの可能性があります。",
    "reason_save_denied": "保存先に書き込めませんでした。",
    "reason_folder_missing": "保存先フォルダが見つかりませんでした。",
    "reason_unknown": "詳しい原因を確認できませんでした。",
    "reason_detail": "詳細: {detail}",
    "reason_with_detail": "{reason} {detail}",
    "output_merged_suffix": "_selected",
    "output_single_suffix": "_page",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_left_suffix": " / 止まらない、迷わない、すぐ終わる。",
    "footer_right_prefix": "戸建買取査定",
    "footer_right_separator": " ｜ ",
    "footer_right_instagram": "Instagram",
    "footer_right_suffix": COPYRIGHT,
    "main_title": "PDFのページを選んで保存する",
    "main_description": "必要なページだけ選んで、1つのPDFにまとめます。",
    "window_minimum_size": "1100x720",
    "status_idle": "未選択",
    "status_loading": "読み込み中",
    "status_ready": "準備完了",
    "status_selecting": "選択中",
    "status_message_idle": "PDFを追加してください",
    "status_message_loading": "PDFを読み込んでいます...",
    "status_message_ready": "抽出するページを選んでください。",
    "status_message_selecting": "抽出するページを選んでください。",
}

FOOTER_URLS = {
    "assessment": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "instagram": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

COLORS = {
    "bg": "#F6F7F9",
    "panel": "#ffffff",
    "panel_alt": "#ffffff",
    "border": "#E6EAF0",
    "text": "#1E2430",
    "muted": "#667085",
    "accent": "#2F6FED",
    "accent_hover": "#2458BF",
    "accent_soft": "#EAF2FF",
    "accent_border": "#7AA7FF",
    "danger": "#c75050",
    "button_secondary": "#ffffff",
    "button_disabled": "#E6EAF0",
    "scrollbar_track": "#EEF1F5",
    "scrollbar_thumb": "#C9D2E0",
    "scrollbar_thumb_hover": "#AEBACC",
}

THUMBNAIL_WIDTH = 152
THUMBNAIL_HEIGHT = 214
CARD_WIDTH = 176
CARD_HEIGHT = 264
CARD_GAP_X = 20
CARD_GAP_Y = 20
CANVAS_PADDING_X = 16
CANVAS_PADDING_Y = 16
SHIFT_MASK = 0x0001


def pick_font_family(root: tk.Misc) -> str:
    available = set(tkfont.families(root))
    for family in ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo"):
        if family in available:
            return family
    return "TkDefaultFont"


class RangeParseError(ValueError):
    pass


class AppState(Enum):
    UNLOADED = "unloaded"
    LOADING = "loading"
    READY = "ready"
    SELECTING = "selecting"
    PROCESSING = "processing"
    COMPLETE = "complete"
    ERROR = "error"


@dataclass
class StatusPayload:
    state: AppState
    message: str


def parse_range_expression(expression: str, page_count: Optional[int] = None) -> set[int]:
    expression = expression.strip()
    if not expression:
        return set()

    pattern = re.compile(r"^\d+(?:-\d+)?(?:,\d+(?:-\d+)?)*$")
    if not pattern.fullmatch(expression):
        raise RangeParseError(UI_TEXT["message_error_range_format"])

    pages: set[int] = set()
    for token in expression.split(","):
        if "-" in token:
            start_text, end_text = token.split("-", 1)
            start = int(start_text)
            end = int(end_text)
            if start < 1 or end < 1:
                raise RangeParseError(UI_TEXT["message_error_range_zero"])
            if start > end:
                raise RangeParseError(UI_TEXT["message_error_range_reverse"])
            if page_count is not None and end > page_count:
                raise RangeParseError(UI_TEXT["message_error_range_over"])
            pages.update(range(start, end + 1))
        else:
            page = int(token)
            if page < 1:
                raise RangeParseError(UI_TEXT["message_error_range_zero"])
            if page_count is not None and page > page_count:
                raise RangeParseError(UI_TEXT["message_error_range_over"])
            pages.add(page)
    return pages


def make_available_path(target_path: Path) -> Path:
    if not target_path.exists():
        return target_path

    index = 1
    while True:
        candidate = target_path.with_name(f"{target_path.stem}_{index}{target_path.suffix}")
        if not candidate.exists():
            return candidate
        index += 1


def open_directory(path: str) -> None:
    if os.name == "nt":
        os.startfile(path)
        return
    if hasattr(subprocess, "run"):
        subprocess.run(["xdg-open", path], check=False)


def resource_path(name: str) -> str:
    base_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return str(base_dir / name)


def get_common_icon_path() -> Path:
    return Path(__file__).resolve().parent / ".." / ".." / "02_assets" / "dake_icon.ico"


def normalize_error_detail(error: object) -> str:
    return " ".join(str(error).strip().split())


def with_error_detail(reason: str, detail: str) -> str:
    if not detail:
        return reason
    return UI_TEXT["reason_with_detail"].format(
        reason=reason,
        detail=UI_TEXT["reason_detail"].format(detail=detail),
    )


def build_pdf_load_error_message(error: object) -> str:
    detail = normalize_error_detail(error)
    lower_detail = detail.lower()
    reason = UI_TEXT["reason_unknown"]

    if "permission" in lower_detail or "denied" in lower_detail or "used by another process" in lower_detail:
        reason = UI_TEXT["reason_file_busy"]
    elif "no such file" in lower_detail or "not found" in lower_detail or "cannot find the file" in lower_detail:
        reason = UI_TEXT["reason_file_not_found"]
    elif "password" in lower_detail or "encrypted" in lower_detail:
        reason = UI_TEXT["reason_pdf_protected"]
    elif "pdf" in lower_detail or "document" in lower_detail or "broken" in lower_detail:
        reason = UI_TEXT["reason_pdf_invalid"]

    return UI_TEXT["message_error_open_pdf_detail"].format(reason=with_error_detail(reason, detail))


def build_thumbnail_error_message(error: object) -> str:
    detail = normalize_error_detail(error)
    lower_detail = detail.lower()
    reason = UI_TEXT["reason_unknown"]

    if "password" in lower_detail or "encrypted" in lower_detail:
        reason = UI_TEXT["reason_pdf_protected"]
    elif detail:
        reason = UI_TEXT["reason_pdf_invalid"]

    return UI_TEXT["message_error_render_pdf_detail"].format(reason=with_error_detail(reason, detail))


def build_save_error_message(error: object) -> str:
    detail = normalize_error_detail(error)
    lower_detail = detail.lower()
    reason = UI_TEXT["reason_unknown"]

    if "permission" in lower_detail or "denied" in lower_detail or "used by another process" in lower_detail:
        reason = UI_TEXT["reason_save_denied"]
    elif "no such file" in lower_detail or "not found" in lower_detail or "cannot find the path" in lower_detail:
        reason = UI_TEXT["reason_folder_missing"]
    elif "password" in lower_detail or "encrypted" in lower_detail:
        reason = UI_TEXT["reason_pdf_protected"]
    elif "pdf" in lower_detail or "document" in lower_detail or "broken" in lower_detail:
        reason = UI_TEXT["reason_pdf_invalid"]

    return UI_TEXT["message_error_processing_detail"].format(reason=with_error_detail(reason, detail))


def build_drop_error_message(error: object) -> str:
    detail = normalize_error_detail(error)
    reason = UI_TEXT["reason_unknown"] if detail else UI_TEXT["label_drop_file_error"]
    return UI_TEXT["message_error_drop_detail"].format(reason=with_error_detail(reason, detail))


class ThumbnailViewport(ttk.Frame):
    def __init__(self, master, request_thumbnail_callback, toggle_page_callback, document_ready_callback=None):
        super().__init__(master, style="Surface.TFrame")
        self.request_thumbnail_callback = request_thumbnail_callback
        self.toggle_page_callback = toggle_page_callback
        self.document_ready_callback = document_ready_callback
        self.page_count = 0
        self.columns = 1
        self.thumbnail_cache: dict[int, ImageTk.PhotoImage] = {}
        self.requested_pages: set[int] = set()
        self.selected_pages: set[int] = set()
        self.empty_message = UI_TEXT["label_thumbnail_empty"]
        self.ready_reported = False
        self._redraw_job = None

        self.canvas = tk.Canvas(
            self,
            bg=COLORS["panel"],
            bd=0,
            highlightthickness=0,
            relief="flat",
        )
        self.scrollbar = ttk.Scrollbar(
            self,
            orient="vertical",
            command=self._on_scrollbar,
            style="App.Vertical.TScrollbar",
        )
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self.canvas.bind("<Configure>", self._on_configure)
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Button-1>", self._on_click)

    def reset(self, message: Optional[str] = None) -> None:
        self.page_count = 0
        self.columns = 1
        self.thumbnail_cache.clear()
        self.requested_pages.clear()
        self.selected_pages.clear()
        self.ready_reported = False
        if message is not None:
            self.empty_message = message
        self.canvas.delete("all")
        self.canvas.configure(scrollregion=(0, 0, 0, 0))
        self._draw_empty_state()

    def set_document(self, page_count: int) -> None:
        self.page_count = page_count
        self.thumbnail_cache.clear()
        self.requested_pages.clear()
        self.selected_pages.clear()
        self.ready_reported = False
        self._schedule_redraw()

    def set_selected_pages(self, pages: set[int]) -> None:
        self.selected_pages = set(pages)
        self._schedule_redraw()

    def clear_cache(self) -> None:
        self.thumbnail_cache.clear()
        self.requested_pages.clear()
        self._schedule_redraw()

    def set_thumbnail(self, page_index: int, photo_image: ImageTk.PhotoImage) -> None:
        self.thumbnail_cache[page_index] = photo_image
        self.requested_pages.discard(page_index)
        self._schedule_redraw()

    def _on_scrollbar(self, *args) -> None:
        self.canvas.yview(*args)
        self._schedule_redraw()

    def _on_mousewheel(self, event) -> str:
        delta = 0
        if event.delta:
            delta = -1 * int(event.delta / 120)
        if delta:
            self.canvas.yview_scroll(delta, "units")
            self._schedule_redraw()
        return "break"

    def _on_configure(self, _event) -> None:
        self._schedule_redraw()

    def _on_click(self, event) -> None:
        if self.page_count == 0:
            return
        canvas_x = self.canvas.canvasx(event.x)
        canvas_y = self.canvas.canvasy(event.y)
        index = self._page_index_from_point(canvas_x, canvas_y)
        if index is None:
            return
        self.toggle_page_callback(index + 1, bool(event.state & SHIFT_MASK))

    def _page_index_from_point(self, x: float, y: float) -> Optional[int]:
        if self.page_count == 0:
            return None

        content_width = max(self.canvas.winfo_width(), CARD_WIDTH + (CANVAS_PADDING_X * 2))
        self.columns = self._compute_columns(content_width)
        x -= CANVAS_PADDING_X
        y -= CANVAS_PADDING_Y
        if x < 0 or y < 0:
            return None

        column = int(x // (CARD_WIDTH + CARD_GAP_X))
        row = int(y // (CARD_HEIGHT + CARD_GAP_Y))
        if column < 0 or column >= self.columns:
            return None

        local_x = x - (column * (CARD_WIDTH + CARD_GAP_X))
        local_y = y - (row * (CARD_HEIGHT + CARD_GAP_Y))
        if local_x > CARD_WIDTH or local_y > CARD_HEIGHT:
            return None

        index = row * self.columns + column
        if index >= self.page_count:
            return None
        return index

    def _compute_columns(self, canvas_width: int) -> int:
        usable_width = max(canvas_width - (CANVAS_PADDING_X * 2), CARD_WIDTH)
        return max(1, (usable_width + CARD_GAP_X) // (CARD_WIDTH + CARD_GAP_X))

    def _schedule_redraw(self) -> None:
        if self._redraw_job:
            self.after_cancel(self._redraw_job)
        self._redraw_job = self.after(16, self._redraw)

    def _redraw(self) -> None:
        self._redraw_job = None
        self.canvas.delete("all")

        if self.page_count == 0:
            self._draw_empty_state()
            return

        canvas_width = max(self.canvas.winfo_width(), CARD_WIDTH + (CANVAS_PADDING_X * 2))
        canvas_height = max(self.canvas.winfo_height(), CARD_HEIGHT)
        self.columns = self._compute_columns(canvas_width)
        rows = (self.page_count + self.columns - 1) // self.columns
        total_width = (self.columns * CARD_WIDTH) + ((self.columns - 1) * CARD_GAP_X) + (CANVAS_PADDING_X * 2)
        total_height = (rows * CARD_HEIGHT) + ((rows - 1) * CARD_GAP_Y) + (CANVAS_PADDING_Y * 2)
        self.canvas.configure(scrollregion=(0, 0, total_width, total_height))

        y0 = self.canvas.canvasy(0)
        y1 = y0 + canvas_height
        row_height = CARD_HEIGHT + CARD_GAP_Y
        start_row = max(0, int(y0 // row_height) - 1)
        end_row = min(rows - 1, int(y1 // row_height) + 1)

        visible_indices: list[int] = []
        for row in range(start_row, end_row + 1):
            for column in range(self.columns):
                index = (row * self.columns) + column
                if index >= self.page_count:
                    continue
                visible_indices.append(index)

        for page_index in visible_indices:
            self._draw_page_card(page_index)

        visible_complete = bool(visible_indices) and all(
            page_index in self.thumbnail_cache for page_index in visible_indices
        )
        if visible_complete != self.ready_reported:
            self.ready_reported = visible_complete
            if self.document_ready_callback is not None:
                self.after_idle(self.document_ready_callback, visible_complete)

    def _draw_empty_state(self) -> None:
        canvas_width = max(self.canvas.winfo_width(), 360)
        canvas_height = max(self.canvas.winfo_height(), 300)
        self.canvas.configure(scrollregion=(0, 0, canvas_width, canvas_height))
        self.canvas.create_text(
            canvas_width / 2,
            canvas_height / 2,
            text=self.empty_message,
            fill=COLORS["muted"],
            font=(self.winfo_toplevel().font_family, 12),
            width=420,
            justify="center",
        )

    def _draw_page_card(self, page_index: int) -> None:
        row = page_index // self.columns
        column = page_index % self.columns
        x1 = CANVAS_PADDING_X + (column * (CARD_WIDTH + CARD_GAP_X))
        y1 = CANVAS_PADDING_Y + (row * (CARD_HEIGHT + CARD_GAP_Y))
        x2 = x1 + CARD_WIDTH
        y2 = y1 + CARD_HEIGHT
        is_selected = (page_index + 1) in self.selected_pages

        border_color = COLORS["accent_border"] if is_selected else COLORS["border"]
        fill_color = COLORS["accent_soft"] if is_selected else COLORS["panel"]

        self.canvas.create_rectangle(
            x1,
            y1,
            x2,
            y2,
            outline=border_color,
            width=1,
            fill=fill_color,
        )
        preview_x1 = x1 + 12
        preview_y1 = y1 + 12
        preview_x2 = preview_x1 + THUMBNAIL_WIDTH
        preview_y2 = preview_y1 + THUMBNAIL_HEIGHT

        self.canvas.create_rectangle(
            preview_x1,
            preview_y1,
            preview_x2,
            preview_y2,
            outline=COLORS["border"],
            width=1,
            fill=COLORS["panel_alt"],
        )

        cached_image = self.thumbnail_cache.get(page_index)
        if cached_image is None:
            self.canvas.create_text(
                (preview_x1 + preview_x2) / 2,
                (preview_y1 + preview_y2) / 2,
                text=UI_TEXT["label_thumbnail_loading"],
                fill=COLORS["muted"],
                font=(self.winfo_toplevel().font_family, 10),
                width=THUMBNAIL_WIDTH - 20,
                justify="center",
            )
            if page_index not in self.requested_pages:
                self.requested_pages.add(page_index)
                self.request_thumbnail_callback(page_index)
        else:
            self.canvas.create_image(
                (preview_x1 + preview_x2) / 2,
                (preview_y1 + preview_y2) / 2,
                image=cached_image,
            )

        self.canvas.create_text(
            x1 + 14,
            y2 - 18,
            anchor="w",
            text=UI_TEXT["label_thumbnail_page"].format(page=page_index + 1),
            fill=COLORS["text"],
            font=(self.winfo_toplevel().font_family, 10, "bold"),
        )

        if is_selected:
            self.canvas.create_text(
                x2 - 14,
                y2 - 18,
                anchor="e",
                text=UI_TEXT["label_thumbnail_selected"],
                fill=COLORS["accent"],
                font=(self.winfo_toplevel().font_family, 9),
            )


class DakePdfSplitSelectApp(BASE_WINDOW):
    def __init__(self):
        super().__init__()
        self.title(WINDOW_TITLE)
        min_width, min_height = UI_TEXT["window_minimum_size"].split("x")
        self.minsize(int(min_width), int(min_height))
        self.geometry("1280x820")
        self.configure(bg=COLORS["bg"])
        self.font_family = pick_font_family(self)
        icon_path = get_common_icon_path().resolve()
        if not getattr(sys, "frozen", False) and icon_path.exists():
            try:
                self.iconbitmap(str(icon_path))
            except tk.TclError:
                pass

        self.queue_events: queue.Queue = queue.Queue()
        self.render_queue: queue.PriorityQueue = queue.PriorityQueue()
        self.render_order = count()
        self.render_generation = 0
        self.pdf_path: Optional[str] = None
        self.page_count = 0
        self.current_pdf_path: Optional[str] = None
        self.current_pdf_name = ""
        self.total_pages = 0
        self.save_dir: Optional[str] = None
        self.save_dir_is_manual = False
        self.selected_pages: set[int] = set()
        self.thumbnail_selected_pages: set[int] = set()
        self.range_selected_pages: set[int] = set()
        self.range_error_message = ""
        self.current_status = StatusPayload(AppState.UNLOADED, UI_TEXT["status_message_idle"])
        self.is_processing = False
        self.completed_state_message = ""
        self.document_ready = False
        self.visible_thumbnails_ready = False
        self.generated_thumbnail_count = 0
        self.generated_thumbnail_pages: set[int] = set()
        self.queued_thumbnail_pages: set[int] = set()
        self.thumbnail_generation_complete = False
        self.thumbnail_selection_anchor: Optional[int] = None

        self.file_name_var = tk.StringVar(value=UI_TEXT["label_file_default"])
        self.total_pages_var = tk.StringVar(value=UI_TEXT["label_total_pages_default"])
        self.save_dir_var = tk.StringVar(value=UI_TEXT["label_save_dir_default"])
        self.range_var = tk.StringVar(value="")
        self.range_input_var = self.range_var
        self.selected_count_var = tk.StringVar(value=UI_TEXT["label_selected_pages"].format(count=0))
        self.state_var = tk.StringVar(value=UI_TEXT["status_idle"])
        self.status_message_var = tk.StringVar(value=UI_TEXT["status_message_idle"])
        self.range_error_var = tk.StringVar(value="")

        self._configure_styles()
        self._build_layout()
        self.set_status("status_idle")
        self._bind_events()
        self._start_background_workers()
        self._refresh_ui_state()
        self._poll_worker_events()

    def _configure_styles(self) -> None:
        style = ttk.Style(self)
        if "clam" in style.theme_names():
            style.theme_use("clam")

        base_font = (self.font_family, 10)
        style.configure(".", background=COLORS["bg"], foreground=COLORS["text"], font=base_font)
        style.configure("App.TFrame", background=COLORS["bg"])
        style.configure("Surface.TFrame", background=COLORS["panel"], relief="flat")
        style.configure("Status.TFrame", background=COLORS["panel"])
        style.configure("Footer.TFrame", background=COLORS["panel"])
        style.configure(
            "FieldLabel.TLabel",
            background=COLORS["bg"],
            foreground=COLORS["muted"],
            font=(self.font_family, 9),
        )
        style.configure(
            "Value.TLabel",
            background=COLORS["bg"],
            foreground=COLORS["text"],
            font=(self.font_family, 10),
        )
        style.configure(
            "PanelTitle.TLabel",
            background=COLORS["panel"],
            foreground=COLORS["text"],
            font=(self.font_family, 11, "bold"),
        )
        style.configure(
            "PanelHint.TLabel",
            background=COLORS["panel"],
            foreground=COLORS["muted"],
            font=(self.font_family, 9),
        )
        style.configure(
            "MainTitle.TLabel",
            background=COLORS["bg"],
            foreground=COLORS["text"],
            font=(self.font_family, 14, "bold"),
        )
        style.configure(
            "MainDescription.TLabel",
            background=COLORS["bg"],
            foreground=COLORS["muted"],
            font=(self.font_family, 10),
        )
        style.configure(
            "StatusState.TLabel",
            background=COLORS["panel"],
            foreground=COLORS["text"],
            font=(self.font_family, 11, "bold"),
        )
        style.configure(
            "StatusMessage.TLabel",
            background=COLORS["panel"],
            foreground=COLORS["muted"],
            font=(self.font_family, 10),
        )
        style.configure(
            "Primary.TButton",
            background=COLORS["accent"],
            foreground="#ffffff",
            padding=(18, 11),
            borderwidth=1,
            bordercolor=COLORS["accent"],
            focusthickness=0,
            relief="flat",
            font=(self.font_family, 10, "bold"),
        )
        style.map(
            "Primary.TButton",
            background=[
                ("disabled", COLORS["button_disabled"]),
                ("active", COLORS["accent_hover"]),
            ],
            foreground=[("disabled", "#ffffff"), ("active", "#ffffff")],
            bordercolor=[
                ("disabled", COLORS["button_disabled"]),
                ("active", COLORS["accent_hover"]),
            ],
        )
        style.configure(
            "Secondary.TButton",
            background=COLORS["button_secondary"],
            foreground=COLORS["text"],
            padding=(16, 10),
            borderwidth=1,
            bordercolor=COLORS["border"],
            focusthickness=0,
            relief="flat",
            font=(self.font_family, 10),
        )
        style.map(
            "Secondary.TButton",
            background=[
                ("disabled", COLORS["button_disabled"]),
                ("active", COLORS["button_secondary"]),
            ],
            foreground=[("disabled", COLORS["muted"]), ("active", COLORS["text"])],
            bordercolor=[
                ("disabled", COLORS["border"]),
                ("active", COLORS["border"]),
            ],
        )
        focusless_button_layout = [
            (
                "Button.border",
                {
                    "sticky": "nswe",
                    "border": "1",
                    "children": [
                        (
                            "Button.padding",
                            {
                                "sticky": "nswe",
                                "children": [("Button.label", {"sticky": "nswe"})],
                            },
                        )
                    ],
                },
            )
        ]
        style.layout("Primary.TButton", focusless_button_layout)
        style.layout("Secondary.TButton", focusless_button_layout)
        base_scrollbar_layout = style.layout("Vertical.TScrollbar")
        if base_scrollbar_layout:
            style.layout("App.Vertical.TScrollbar", base_scrollbar_layout)
        style.configure(
            "App.Vertical.TScrollbar",
            background=COLORS["scrollbar_thumb"],
            troughcolor=COLORS["scrollbar_track"],
            bordercolor=COLORS["scrollbar_track"],
            lightcolor=COLORS["scrollbar_track"],
            darkcolor=COLORS["scrollbar_track"],
            arrowcolor=COLORS["scrollbar_thumb"],
            relief="flat",
            borderwidth=0,
            arrowsize=10,
            gripcount=0,
            width=10,
        )
        style.map(
            "App.Vertical.TScrollbar",
            background=[("active", COLORS["scrollbar_thumb_hover"]), ("pressed", COLORS["scrollbar_thumb_hover"])],
            arrowcolor=[("active", COLORS["scrollbar_thumb_hover"]), ("pressed", COLORS["scrollbar_thumb_hover"])],
        )
        style.configure(
            "TEntry",
            fieldbackground=COLORS["panel"],
            foreground=COLORS["text"],
            padding=8,
            bordercolor=COLORS["border"],
            lightcolor=COLORS["border"],
            darkcolor=COLORS["border"],
            relief="solid",
            font=(self.font_family, 10),
        )

    def _build_layout(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self.rowconfigure(1, weight=0)

        app_frame = ttk.Frame(self, style="App.TFrame", padding=(24, 24, 24, 16))
        app_frame.grid(row=0, column=0, sticky="nsew")
        app_frame.columnconfigure(0, weight=1)
        app_frame.rowconfigure(2, weight=1)

        top_frame = ttk.Frame(app_frame, style="App.TFrame")
        top_frame.grid(row=0, column=0, sticky="ew", pady=(0, 20))
        top_frame.columnconfigure(5, weight=1)

        header_frame = ttk.Frame(top_frame, style="App.TFrame")
        header_frame.grid(row=0, column=0, columnspan=6, sticky="ew", pady=(0, 14))
        header_frame.columnconfigure(1, weight=1)

        ttk.Label(header_frame, text=UI_TEXT["main_title"], style="MainTitle.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(header_frame, text=UI_TEXT["main_description"], style="MainDescription.TLabel").grid(
            row=0, column=1, sticky="w", padx=(12, 0)
        )

        self.add_pdf_button = ttk.Button(
            top_frame,
            text=UI_TEXT["button_add_pdf"],
            command=self._choose_pdf,
            style="Primary.TButton",
        )
        self.add_pdf_button.grid(row=1, column=0, padx=(0, 10), sticky="w")

        self.choose_save_dir_button = ttk.Button(
            top_frame,
            text=UI_TEXT["button_select_save_folder"],
            command=self._choose_save_dir,
            style="Secondary.TButton",
        )
        self.choose_save_dir_button.grid(row=1, column=1, padx=(0, 10), sticky="w")

        self.refresh_button = ttk.Button(
            top_frame,
            text=UI_TEXT["button_refresh"],
            command=self._refresh_pdf,
            style="Secondary.TButton",
        )
        self.refresh_button.grid(row=1, column=2, padx=(0, 12), sticky="w")

        self.clear_selection_button = ttk.Button(
            top_frame,
            text=UI_TEXT["button_clear_selection"],
            command=self.clear_selection,
            style="Secondary.TButton",
        )
        self.clear_selection_button.grid(row=1, column=3, padx=(0, 16), sticky="w")

        file_info_frame = ttk.Frame(top_frame, style="App.TFrame")
        file_info_frame.grid(row=1, column=5, sticky="ew")
        file_info_frame.columnconfigure(1, weight=1)

        ttk.Label(file_info_frame, text=UI_TEXT["label_file_name"], style="FieldLabel.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0, 12)
        )
        ttk.Label(file_info_frame, textvariable=self.file_name_var, style="Value.TLabel").grid(
            row=0, column=1, sticky="w"
        )

        ttk.Label(file_info_frame, text=UI_TEXT["label_total_pages"], style="FieldLabel.TLabel").grid(
            row=1, column=0, sticky="w", padx=(0, 12), pady=(6, 0)
        )
        ttk.Label(file_info_frame, textvariable=self.total_pages_var, style="Value.TLabel").grid(
            row=1, column=1, sticky="w", pady=(6, 0)
        )

        body_frame = ttk.Frame(app_frame, style="App.TFrame")
        body_frame.grid(row=2, column=0, sticky="nsew")
        body_frame.columnconfigure(0, weight=1)
        body_frame.columnconfigure(1, weight=0)
        body_frame.rowconfigure(0, weight=1)

        thumbnail_frame = ttk.Frame(body_frame, style="Surface.TFrame", padding=16)
        thumbnail_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 16))
        thumbnail_frame.columnconfigure(0, weight=1)
        thumbnail_frame.rowconfigure(0, weight=1)

        self.thumbnail_viewport = ThumbnailViewport(
            thumbnail_frame,
            request_thumbnail_callback=self._request_thumbnail,
            toggle_page_callback=self._toggle_thumbnail_page,
            document_ready_callback=self._on_thumbnail_view_ready,
        )
        self.thumbnail_viewport.grid(row=0, column=0, sticky="nsew")

        control_frame = ttk.Frame(body_frame, style="Surface.TFrame", padding=20)
        control_frame.grid(row=0, column=1, sticky="ns")
        control_frame.columnconfigure(0, weight=1)

        ttk.Label(control_frame, text=UI_TEXT["label_range_input"], style="PanelTitle.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        self.range_entry = ttk.Entry(control_frame, textvariable=self.range_var, width=28)
        self.range_entry.grid(row=1, column=0, sticky="ew", pady=(10, 6))
        ttk.Label(control_frame, text=UI_TEXT["label_range_example"], style="PanelHint.TLabel").grid(
            row=2, column=0, sticky="w"
        )
        self.range_error_label = tk.Label(
            control_frame,
            textvariable=self.range_error_var,
            bg=COLORS["panel"],
            fg=COLORS["danger"],
            font=(self.font_family, 9),
            wraplength=240,
            justify="left",
        )
        self.range_error_label.grid(row=3, column=0, sticky="w", pady=(8, 18))

        ttk.Label(control_frame, textvariable=self.selected_count_var, style="PanelTitle.TLabel").grid(
            row=4, column=0, sticky="w", pady=(0, 8)
        )
        self.save_dir_label = ttk.Label(
            control_frame,
            textvariable=self.save_dir_var,
            style="PanelHint.TLabel",
            wraplength=260,
            justify="left",
        )
        self.save_dir_label.grid(row=5, column=0, sticky="w", pady=(0, 20))

        self.extract_merged_button = ttk.Button(
            control_frame,
            text=UI_TEXT["button_extract"],
            command=lambda: self._start_extract(mode="merged"),
            style="Primary.TButton",
        )
        self.extract_merged_button.grid(row=6, column=0, sticky="ew", pady=(0, 10))

        self.extract_single_button = ttk.Button(
            control_frame,
            text=UI_TEXT["button_split_each"],
            command=lambda: self._start_extract(mode="single"),
            style="Primary.TButton",
        )
        self.extract_single_button.grid(row=7, column=0, sticky="ew")

        button_widgets = (
            self.add_pdf_button,
            self.choose_save_dir_button,
            self.refresh_button,
            self.clear_selection_button,
            self.extract_merged_button,
            self.extract_single_button,
        )
        for button in button_widgets:
            self._configure_button_widget(button)

        status_frame = ttk.Frame(app_frame, style="Status.TFrame", padding=(16, 14))
        status_frame.grid(row=3, column=0, sticky="ew", pady=(16, 0))
        status_frame.columnconfigure(1, weight=1)

        ttk.Label(status_frame, textvariable=self.state_var, style="StatusState.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0, 16)
        )
        ttk.Label(status_frame, textvariable=self.status_message_var, style="StatusMessage.TLabel").grid(
            row=0, column=1, sticky="w"
        )

        footer_frame = ttk.Frame(self, style="Footer.TFrame", padding=(24, 10, 24, 14))
        footer_frame.grid(row=1, column=0, sticky="ew")
        footer_frame.columnconfigure(0, weight=1)
        footer_frame.columnconfigure(1, weight=0)

        left_footer_frame = ttk.Frame(footer_frame, style="Footer.TFrame")
        left_footer_frame.grid(row=0, column=0, sticky="w")

        tk.Label(
            left_footer_frame,
            text=UI_TEXT["footer_left"],
            bg=COLORS["panel"],
            fg=COLORS["muted"],
            font=(self.font_family, 8),
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            left_footer_frame,
            text=UI_TEXT["footer_left_suffix"],
            bg=COLORS["panel"],
            fg=COLORS["muted"],
            font=(self.font_family, 8),
        ).grid(row=0, column=1, sticky="w")

        right_footer = ttk.Frame(footer_frame, style="Footer.TFrame")
        right_footer.grid(row=0, column=1, sticky="e")

        self._build_footer_links(right_footer)

    def _build_footer_links(self, parent) -> None:
        parts = [
            ("assessment", UI_TEXT["footer_right_prefix"], True),
            ("sep1", UI_TEXT["footer_right_separator"], False),
            ("instagram", UI_TEXT["footer_right_instagram"], True),
            ("sep2", UI_TEXT["footer_right_separator"], False),
            ("copyright", UI_TEXT["footer_right_suffix"], False),
        ]

        column = 0
        for key, text, is_link in parts:
            label = tk.Label(
                parent,
                text=text,
                bg=COLORS["panel"],
                fg=COLORS["muted"],
                font=(self.font_family, 8),
                cursor="hand2" if is_link else "",
            )
            label.grid(row=0, column=column, sticky="w")
            if is_link:
                url = FOOTER_URLS["assessment"] if key == "assessment" else FOOTER_URLS["instagram"]
                label.bind("<Button-1>", lambda _event, link=url: webbrowser.open_new_tab(link))
                label.bind("<Enter>", lambda event: event.widget.configure(fg=COLORS["text"]))
                label.bind("<Leave>", lambda event: event.widget.configure(fg=COLORS["muted"]))
            column += 1

    def _bind_events(self) -> None:
        self.range_var.trace_add("write", self._on_range_changed)
        if DND_READY and hasattr(self, "drop_target_register"):
            self.drop_target_register(DND_FILES)
            self.dnd_bind("<<Drop>>", self._on_drop)
            if hasattr(self.thumbnail_viewport.canvas, "drop_target_register"):
                self.thumbnail_viewport.canvas.drop_target_register(DND_FILES)
                self.thumbnail_viewport.canvas.dnd_bind("<<Drop>>", self._on_drop)

    def _start_background_workers(self) -> None:
        render_thread = threading.Thread(target=self._render_worker, daemon=True)
        render_thread.start()

    def _queue_render_task(self, priority: int, task: dict) -> None:
        self.render_queue.put((priority, next(self.render_order), task))

    def _set_status(self, state: AppState, message: str) -> None:
        self.current_status = StatusPayload(state=state, message=message)
        self._refresh_status_widgets()

    def set_status(self, status_key: str) -> None:
        if status_key == "status_idle":
            self.current_status = StatusPayload(AppState.UNLOADED, UI_TEXT["status_message_idle"])
            self.state_var.set(UI_TEXT["status_idle"])
            self.status_message_var.set(UI_TEXT["status_message_idle"])
        elif status_key == "status_loading":
            self.current_status = StatusPayload(AppState.LOADING, UI_TEXT["status_message_loading"])
            self.state_var.set(UI_TEXT["status_loading"])
            self.status_message_var.set(UI_TEXT["status_message_loading"])
        elif status_key == "status_ready":
            self.current_status = StatusPayload(AppState.READY, UI_TEXT["status_message_ready"])
            self.state_var.set(UI_TEXT["status_ready"])
            self.status_message_var.set(UI_TEXT["status_message_ready"])
        elif status_key == "status_selecting":
            self.current_status = StatusPayload(AppState.SELECTING, UI_TEXT["status_message_selecting"])
            self.state_var.set(UI_TEXT["status_selecting"])
            self.status_message_var.set(UI_TEXT["status_message_selecting"])

    def _configure_button_widget(self, button: ttk.Button) -> None:
        button.configure(takefocus=False)
        button.bind("<ButtonRelease-1>", self._release_button_focus, add="+")

    def _release_button_focus(self, _event=None) -> None:
        self.after_idle(self.focus_set)

    def _reset_thumbnail_generation_state(self) -> None:
        self.document_ready = False
        self.visible_thumbnails_ready = False
        self.generated_thumbnail_count = 0
        self.generated_thumbnail_pages.clear()
        self.queued_thumbnail_pages.clear()
        self.thumbnail_generation_complete = False

    def _update_document_ready_state(self) -> None:
        is_ready = bool(
            self.current_pdf_path
            and self.page_count > 0
            and self.thumbnail_generation_complete
            and self.visible_thumbnails_ready
        )
        if self.document_ready == is_ready:
            return
        self.document_ready = is_ready
        self._refresh_ui_state()

    def _on_thumbnail_view_ready(self, visible_complete: bool) -> None:
        self.visible_thumbnails_ready = visible_complete
        self._update_document_ready_state()

    def _refresh_status_widgets(self) -> None:
        state = self.current_status.state
        if state == AppState.UNLOADED:
            self.state_var.set(UI_TEXT["status_idle"])
        elif state == AppState.LOADING:
            self.state_var.set(UI_TEXT["status_loading"])
        elif state == AppState.READY:
            self.state_var.set(UI_TEXT["status_ready"])
        elif state == AppState.SELECTING:
            self.state_var.set(UI_TEXT["status_selecting"])
        elif state == AppState.COMPLETE:
            self.state_var.set(UI_TEXT["label_status_complete"])
        elif state == AppState.ERROR:
            self.state_var.set(UI_TEXT["label_status_error"])
        elif state == AppState.PROCESSING:
            self.state_var.set(UI_TEXT["label_status_processing"])

        self.status_message_var.set(self.current_status.message)

    def update_selection_ui(self) -> None:
        self.selected_pages = set(self._effective_pages())
        self.selected_count_var.set(UI_TEXT["label_selected_pages"].format(count=len(self.selected_pages)))
        self.thumbnail_viewport.set_selected_pages(self.thumbnail_selected_pages)

    def update_action_buttons(self) -> None:
        if self.is_processing:
            self.extract_merged_button.configure(state="disabled")
            self.extract_single_button.configure(state="disabled")
            self.clear_selection_button.configure(state="disabled")
            return

        if not self.current_pdf_path:
            self.extract_merged_button.configure(state="disabled")
            self.extract_single_button.configure(state="disabled")
            self.clear_selection_button.configure(state="disabled")
            return

        if len(self.selected_pages) == 0:
            self.extract_merged_button.configure(state="disabled")
            self.extract_single_button.configure(state="disabled")
            self.clear_selection_button.configure(state="disabled")
        else:
            self.extract_merged_button.configure(state="normal")
            self.extract_single_button.configure(state="normal")
            self.clear_selection_button.configure(state="normal")

    def _sync_selection_status(self) -> None:
        if self.current_pdf_path and not self.document_ready:
            self.set_status("status_loading")
        elif self.selected_pages:
            self.set_status("status_selecting")
        else:
            self.set_status("status_ready")

    def clear_selection(self) -> None:
        self.thumbnail_selected_pages.clear()
        self.range_selected_pages.clear()
        self.selected_pages.clear()
        self.thumbnail_selection_anchor = None
        self.range_error_message = ""
        self.range_error_var.set("")
        self.range_input_var.set("")
        self.update_selection_ui()
        self.update_action_buttons()
        self._sync_selection_status()

    def refresh_all(self) -> None:
        self.render_generation += 1
        self.current_pdf_path = None
        self.current_pdf_name = ""
        self.total_pages = 0
        self.pdf_path = None
        self.page_count = 0
        self._reset_thumbnail_generation_state()
        self.thumbnail_selection_anchor = None
        self.completed_state_message = ""
        self.range_error_message = ""
        self.thumbnail_selected_pages.clear()
        self.range_selected_pages.clear()
        self.selected_pages.clear()
        self.range_error_var.set("")
        self.range_input_var.set("")
        self.thumbnail_viewport.thumbnail_cache.clear()
        self.thumbnail_viewport.requested_pages.clear()
        self.thumbnail_viewport.reset(UI_TEXT["label_thumbnail_empty"])
        self.file_name_var.set("")
        self.total_pages_var.set("")
        self.update_selection_ui()
        self.update_action_buttons()
        self.set_status("status_idle")

    def _choose_pdf(self) -> None:
        path = filedialog.askopenfilename(
            title=UI_TEXT["file_dialog_title_pdf"],
            filetypes=[(UI_TEXT["file_dialog_pdf_filter"], "*.pdf")],
        )
        if path:
            self._load_pdf_async(path)

    def _choose_save_dir(self) -> None:
        path = filedialog.askdirectory(title=UI_TEXT["file_dialog_title_save_dir"])
        if not path:
            return
        self.save_dir = path
        self.save_dir_is_manual = True
        self.completed_state_message = ""
        self._refresh_save_dir_text()
        self.update_action_buttons()
        self._refresh_ui_state()

    def _refresh_pdf(self) -> None:
        self.refresh_all()

    def _load_pdf_async(self, path: str, keep_current_selection: bool = False) -> None:
        target_path = Path(path)
        self.render_generation += 1
        generation = self.render_generation
        self.pdf_path = str(target_path)
        self.current_pdf_path = str(target_path)
        self.current_pdf_name = target_path.name
        self.page_count = 0
        self.total_pages = 0
        self._reset_thumbnail_generation_state()
        self.thumbnail_selection_anchor = None
        self.thumbnail_viewport.reset(UI_TEXT["label_thumbnail_loading"])

        if not keep_current_selection:
            self.thumbnail_selected_pages.clear()
            self.range_selected_pages.clear()
            self.selected_pages.clear()
            self.range_var.set("")
        else:
            self.thumbnail_selected_pages = set(self.thumbnail_selected_pages)

        self.range_error_message = ""
        self.range_error_var.set("")
        self.file_name_var.set(target_path.name)
        self.total_pages_var.set(UI_TEXT["label_total_pages_default"])
        self.set_status("status_loading")
        self._refresh_ui_state()

        thread = threading.Thread(
            target=self._load_pdf_worker,
            args=(str(target_path), generation),
            daemon=True,
        )
        thread.start()

    def _load_pdf_worker(self, path: str, generation: int) -> None:
        try:
            with fitz.open(path) as document:
                page_count = document.page_count
            self.queue_events.put(("pdf_loaded", generation, path, page_count))
        except Exception as error:
            self.queue_events.put(("pdf_load_failed", generation, build_pdf_load_error_message(error)))

    def _enqueue_thumbnail_render(self, generation: int, page_index: int, priority: int) -> None:
        if generation != self.render_generation or not self.pdf_path:
            return
        if page_index < 0 or page_index >= self.page_count:
            return
        if page_index in self.generated_thumbnail_pages or page_index in self.queued_thumbnail_pages:
            return
        self.queued_thumbnail_pages.add(page_index)
        self._queue_render_task(
            priority=priority,
            task={
                "kind": "render",
                "generation": generation,
                "path": self.pdf_path,
                "page_index": page_index,
            },
        )

    def _enqueue_thumbnail_generation_batch(self, generation: int, start_index: int = 0, batch_size: int = 24) -> None:
        if generation != self.render_generation or self.page_count == 0:
            return
        end_index = min(start_index + batch_size, self.page_count)
        for page_index in range(start_index, end_index):
            self._enqueue_thumbnail_render(generation, page_index, priority=2)
        if end_index < self.page_count:
            self.after(1, self._enqueue_thumbnail_generation_batch, generation, end_index, batch_size)

    def _request_thumbnail(self, page_index: int) -> None:
        self._enqueue_thumbnail_render(self.render_generation, page_index, priority=1)

    def _render_worker(self) -> None:
        current_path = None
        current_generation = -1
        document = None
        while True:
            _, _, task = self.render_queue.get()
            kind = task.get("kind")
            generation = task.get("generation", -1)
            path = task.get("path")

            if kind != "render":
                continue

            try:
                if path != current_path or generation != current_generation:
                    if document is not None:
                        document.close()
                        document = None
                    current_path = path
                    current_generation = generation
                    document = fitz.open(path)

                if document is None or generation != current_generation:
                    continue

                page = document.load_page(task["page_index"])
                pixmap = page.get_pixmap(matrix=fitz.Matrix(1.0, 1.0), alpha=False)
                image = Image.frombytes("RGB", [pixmap.width, pixmap.height], pixmap.samples)
                image.thumbnail((THUMBNAIL_WIDTH, THUMBNAIL_HEIGHT), Image.Resampling.LANCZOS)
                self.queue_events.put(("thumbnail_ready", generation, task["page_index"], image))
            except Exception as error:
                self.queue_events.put(
                    ("thumbnail_failed", generation, task["page_index"], build_thumbnail_error_message(error))
                )

    def _toggle_thumbnail_page(self, page_number: int, shift_pressed: bool = False) -> None:
        self.completed_state_message = ""
        if shift_pressed and self.thumbnail_selection_anchor is not None:
            start_page = min(self.thumbnail_selection_anchor, page_number)
            end_page = max(self.thumbnail_selection_anchor, page_number)
            range_pages = set(range(start_page, end_page + 1))
            if page_number in self.thumbnail_selected_pages:
                self.thumbnail_selected_pages.difference_update(range_pages)
            else:
                self.thumbnail_selected_pages.update(range_pages)
        else:
            if page_number in self.thumbnail_selected_pages:
                self.thumbnail_selected_pages.remove(page_number)
            else:
                self.thumbnail_selected_pages.add(page_number)
        self.thumbnail_selection_anchor = page_number
        self._refresh_ui_state()
        self._sync_selection_status()

    def _on_range_changed(self, *_args) -> None:
        self.completed_state_message = ""
        expression = self.range_var.get()
        if not expression.strip():
            self.range_selected_pages.clear()
            self.range_error_message = ""
            self.range_error_var.set("")
            self._refresh_ui_state()
            self._sync_selection_status()
            return

        try:
            pages = parse_range_expression(expression, self.page_count or None)
            self.range_selected_pages = pages
            self.range_error_message = ""
            self.range_error_var.set("")
        except RangeParseError as error:
            self.range_selected_pages.clear()
            self.range_error_message = str(error)
            self.range_error_var.set(str(error))

        self._refresh_ui_state()
        if self.range_error_message:
            self._set_status(AppState.ERROR, self.range_error_message)
        else:
            self._sync_selection_status()

    def _display_selected_pages(self) -> list[int]:
        if self.thumbnail_selected_pages:
            return sorted(self.thumbnail_selected_pages)
        return sorted(self.range_selected_pages)

    def _effective_pages(self) -> list[int]:
        if self.range_error_message:
            return []
        return self._display_selected_pages()

    def _refresh_ui_state(self) -> None:
        self.update_selection_ui()
        self._refresh_save_dir_text()

        refresh_enabled = bool(self.current_pdf_path and not self.is_processing)
        add_enabled = not self.is_processing
        choose_dir_enabled = not self.is_processing
        range_enabled = not self.is_processing

        self.update_action_buttons()
        self.refresh_button.configure(state="normal" if refresh_enabled else "disabled")
        self.add_pdf_button.configure(state="normal" if add_enabled else "disabled")
        self.choose_save_dir_button.configure(state="normal" if choose_dir_enabled else "disabled")
        self.range_entry.configure(state="normal" if range_enabled else "disabled")

        if self.current_status.state == AppState.ERROR and not self.range_error_message and self.document_ready:
            return

        if self.is_processing:
            return

        if self.range_error_message:
            self._set_status(AppState.ERROR, self.range_error_message)
            return

        if self.completed_state_message:
            self._set_status(AppState.COMPLETE, self.completed_state_message)
            return

        if not self.current_pdf_path:
            self.set_status("status_idle")
            return

        if not self.document_ready:
            self.set_status("status_loading")
            return

        if self.selected_pages:
            self.set_status("status_selecting")
            return

        self.set_status("status_ready")

    def _refresh_save_dir_text(self) -> None:
        if self.save_dir:
            self.save_dir_var.set(UI_TEXT["label_save_dir_value"].format(path=self.save_dir))
        else:
            self.save_dir_var.set(UI_TEXT["label_save_dir_default"])

    def _poll_worker_events(self) -> None:
        while True:
            try:
                event = self.queue_events.get_nowait()
            except queue.Empty:
                break
            self._handle_worker_event(event)
        self.after(50, self._poll_worker_events)

    def _handle_worker_event(self, event) -> None:
        event_type = event[0]

        if event_type == "pdf_loaded":
            _, generation, path, page_count = event
            if generation != self.render_generation:
                return
            self.pdf_path = path
            self.current_pdf_path = path
            self.current_pdf_name = Path(path).name
            self.page_count = page_count
            self.total_pages = page_count
            self._reset_thumbnail_generation_state()
            self.thumbnail_selected_pages = {page for page in self.thumbnail_selected_pages if page <= page_count}
            self.total_pages_var.set(UI_TEXT["label_total_pages_value"].format(count=page_count))
            if not self.save_dir or not self.save_dir_is_manual:
                self.save_dir = str(Path(path).parent)
                self.save_dir_is_manual = False
            self.thumbnail_viewport.set_document(page_count)
            self.after(0, self._enqueue_thumbnail_generation_batch, generation, 0, 24)
            self.completed_state_message = ""
            if self.range_var.get().strip():
                self._on_range_changed()
            else:
                self.range_selected_pages.clear()
                self.range_error_message = ""
                self.range_error_var.set("")
            self._refresh_ui_state()
            return

        if event_type == "pdf_load_failed":
            _, generation, error_message = event
            if generation != self.render_generation:
                return
            self.current_pdf_path = None
            self.current_pdf_name = ""
            self.total_pages = 0
            self.pdf_path = None
            self.page_count = 0
            self._reset_thumbnail_generation_state()
            self.selected_pages.clear()
            self.total_pages_var.set("")
            self.file_name_var.set("")
            self.thumbnail_viewport.reset(UI_TEXT["label_thumbnail_empty"])
            self.completed_state_message = ""
            self._set_status(AppState.ERROR, error_message)
            self._refresh_ui_state()
            return

        if event_type == "thumbnail_ready":
            _, generation, page_index, image = event
            if generation != self.render_generation:
                return
            self.queued_thumbnail_pages.discard(page_index)
            if page_index not in self.generated_thumbnail_pages:
                self.generated_thumbnail_pages.add(page_index)
                self.generated_thumbnail_count = len(self.generated_thumbnail_pages)
                if self.generated_thumbnail_count == self.page_count and self.page_count > 0:
                    self.thumbnail_generation_complete = True
            photo = ImageTk.PhotoImage(image)
            self.thumbnail_viewport.set_thumbnail(page_index, photo)
            self._update_document_ready_state()
            return

        if event_type == "thumbnail_failed":
            _, generation, page_index, error_message = event
            if generation != self.render_generation:
                return
            self.queued_thumbnail_pages.discard(page_index)
            self.thumbnail_viewport.requested_pages.discard(page_index)
            self._set_status(AppState.ERROR, error_message)
            return

        if event_type == "extract_complete":
            _, mode, output_dir, page_count = event
            self.is_processing = False
            self.completed_state_message = (
                UI_TEXT["message_complete_merged"].format(count=page_count)
                if mode == "merged"
                else UI_TEXT["message_complete_single"].format(count=page_count)
            )
            self._refresh_ui_state()
            dialog_message = UI_TEXT["dialog_complete_with_path"].format(
                summary=self.completed_state_message,
                path=output_dir,
            )
            messagebox.showinfo(UI_TEXT["message_complete_title"], dialog_message)
            try:
                open_directory(output_dir)
            except Exception:
                self._set_status(AppState.ERROR, UI_TEXT["message_error_open_folder"])
            self._refresh_ui_state()
            return

        if event_type == "extract_failed":
            _, _mode, error_message = event
            self.is_processing = False
            self.completed_state_message = ""
            self._set_status(AppState.ERROR, error_message)
            self._refresh_ui_state()

    def _start_extract(self, mode: str) -> None:
        if self.is_processing:
            return
        if self.range_error_message:
            self._set_status(AppState.ERROR, self.range_error_message)
            return
        if not self.pdf_path:
            self._set_status(AppState.ERROR, UI_TEXT["message_error_no_pdf"])
            return
        if not self.save_dir:
            self._set_status(AppState.ERROR, UI_TEXT["message_error_save_dir"])
            return

        pages = self._effective_pages()
        if not pages:
            self._set_status(AppState.ERROR, UI_TEXT["message_error_selection"])
            return

        self.is_processing = True
        self.completed_state_message = ""
        self._set_status(AppState.PROCESSING, UI_TEXT["message_processing_extract"])
        self._refresh_ui_state()

        worker = threading.Thread(
            target=self._extract_worker,
            args=(mode, list(pages), self.pdf_path, self.save_dir),
            daemon=True,
        )
        worker.start()

    def _extract_worker(self, mode: str, pages: list[int], pdf_path: str, save_dir: str) -> None:
        try:
            source_path = Path(pdf_path)
            output_dir = Path(save_dir)
            output_dir.mkdir(parents=True, exist_ok=True)
            reader = PdfReader(pdf_path)

            if mode == "merged":
                writer = PdfWriter()
                for page_number in pages:
                    writer.add_page(reader.pages[page_number - 1])
                merged_name = f"{source_path.stem}{UI_TEXT['output_merged_suffix']}.pdf"
                merged_path = make_available_path(output_dir / merged_name)
                with merged_path.open("wb") as handle:
                    writer.write(handle)
            else:
                for page_number in pages:
                    writer = PdfWriter()
                    writer.add_page(reader.pages[page_number - 1])
                    single_name = f"{source_path.stem}{UI_TEXT['output_single_suffix']}_{page_number:03d}.pdf"
                    single_path = make_available_path(output_dir / single_name)
                    with single_path.open("wb") as handle:
                        writer.write(handle)

            self.queue_events.put(("extract_complete", mode, str(output_dir), len(pages)))
        except Exception as error:
            self.queue_events.put(("extract_failed", mode, build_save_error_message(error)))

    def _on_drop(self, event) -> None:
        try:
            files = list(self.tk.splitlist(event.data))
        except tk.TclError:
            self._set_status(AppState.ERROR, build_drop_error_message(event.data))
            return

        if len(files) != 1:
            self._set_status(AppState.ERROR, UI_TEXT["label_drop_multiple_error"])
            return

        file_path = Path(files[0])
        if file_path.suffix.lower() != ".pdf":
            self._set_status(AppState.ERROR, UI_TEXT["label_drop_file_error"])
            return

        self._load_pdf_async(str(file_path))


def main() -> None:
    app = DakePdfSplitSelectApp()
    app.mainloop()


if __name__ == "__main__":
    main()
