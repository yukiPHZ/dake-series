# -*- coding: utf-8 -*-
import json
import os
import queue
import subprocess
import sys
import tempfile
import threading
import webbrowser
import tkinter as tk
from dataclasses import dataclass
from tkinter import filedialog, messagebox, ttk

try:
    import ctypes
except Exception:
    ctypes = None

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except Exception:
    DND_FILES = None
    TkinterDnD = None
    HAS_DND = False


APP_NAME = "PDF結合"
WINDOW_TITLE = "PDF結合"
CONFIG_NAME = "dake_pdf_merge_config.json"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"
COMMON_ICON_RELATIVE = os.path.join("..", "..", "02_assets", "dake_icon.ico")
COMMON_ICON_FILENAME = "dake_icon.ico"
LAUNCHER_LP_URL = "https://dakeapp.com/launcher/"

BG = "#F6F7F9"
CARD = "#FFFFFF"
TEXT = "#1E2430"
SUBTEXT = "#667085"
ACCENT = "#2F6FED"
ACCENT_HOVER = "#2458BF"
BORDER = "#E6EAF0"
PREVIEW_BG = "#F6F8FC"
PREVIEW_BORDER = "#C9D3E3"
SUCCESS = "#12B76A"
DISABLED_BG = "#E8ECF3"
DISABLED_FG = "#98A2B3"
ERROR = "#D92D20"
FOOTER_TEXT = "#AAB2BD"
FONT_CANDIDATES = ["BIZ UDPGothic", "Yu Gothic UI", "Meiryo"]
THUMBNAIL_WORKER_COUNT = 3
DRAG_AUTOSCROLL_EDGE = 64
DRAG_AUTOSCROLL_INTERVAL_MS = 60
PREFERRED_WINDOW_WIDTH = 1180
PREFERRED_WINDOW_HEIGHT = 760
PREFERRED_MIN_WIDTH = 900
PREFERRED_MIN_HEIGHT = 620
NARROW_LAYOUT_THRESHOLD = 1080
TASK_IDLE = "idle"
TASK_PROCESSING = "processing"
TASK_SAVING = "saving"

_PDF_READER = None
_PDF_WRITER = None
_FITZ = None
_FITZ_LOAD_ATTEMPTED = False
_PDF_IMPORT_LOCK = threading.Lock()

UI_TEXT = {
    "main_title": "PDFを結合する",
    "main_description": "複数のPDFを追加して、そのまま1つにまとめます。",
    "button_add": "PDFを追加",
    "button_select_folder": "保存先を選ぶ",
    "button_refresh": "リフレッシュ",
    "button_cancel": "キャンセル",
    "button_execute": "結合して保存",
    "button_move_up": "↑",
    "button_move_down": "↓",
    "button_delete": "削除",
    "drag_hint": "ドラッグして順番を入れ替えできます",
    "label_save_folder": "保存先",
    "label_page_count_unknown": "ページ数を読み込み中",
    "label_page_suffix": "ページ",
    "label_loading_thumbnail": "サムネイル\n読み込み中",
    "empty_title": "PDFを追加してください",
    "empty_title_drop": "PDFをドロップしてください",
    "empty_subtitle": "ドラッグ＆ドロップ または クリックして追加",
    "status_loading": "読み込み中",
    "status_processing": "処理中",
    "status_saving": "保存中",
    "status_ready": "準備完了",
    "status_idle": "未選択",
    "status_canceling": "キャンセル中",
    "status_canceled": "キャンセル完了",
    "status_complete": "保存完了",
    "status_error": "エラー",
    "status_reordered": "並び順を変更しました",
    "count_added": "{count}件追加済み",
    "detail_ready": "結合して保存できます",
    "detail_none": "PDFを追加してください",
    "detail_processing": "PDFを順番どおりに処理しています",
    "detail_saving": "保存ファイルを書き出しています",
    "detail_finalizing": "保存ファイルを確定しています",
    "detail_save_done": "保存フォルダを開きます",
    "detail_cancel": "",
    "detail_error": "処理中に問題が発生しました",
    "detail_file_error": "PDFの処理中に問題が発生しました",
    "msg_refresh_blocked": "処理中はリフレッシュできません",
    "msg_no_files": "PDFを追加してください",
    "msg_save_folder_error_title": "保存先エラー",
    "msg_save_folder_error": "保存先フォルダを準備できませんでした。",
    "msg_save_done": "結合して保存が完了しました",
    "msg_atomic_save_failed": "保存ファイルを作成できませんでした。",
    "dialog_close_title": "処理中です",
    "dialog_close_processing": "処理を中止して閉じますか？",
    "dialog_close_saving": "保存中です。完了までお待ちください。",
    "filetype_pdf": "PDFファイル",
    "footer_left": "シンプルそれDAKEシリーズ / 止まらない、迷わない、すぐ終わる。",
    "footer_link_1": "戸建買取査定",
    "footer_link_2": "Instagram",
    "footer_separator": " ｜ ",
    "footer_copyright": COPYRIGHT,
    "link_other_tools": "他のDAKEツール",
    "cli_error_not_enough_inputs": "入力PDFが2つ未満です",
    "cli_error_file_not_found": "入力ファイルが見つかりません",
    "cli_error_not_pdf": "PDF以外のファイルが含まれています",
    "cli_error_output_failed": "出力先を準備できませんでした",
    "cli_error_merge_failed": "PDF結合に失敗しました",
}

MAIN_TITLE = UI_TEXT["main_title"]
MAIN_SUBTITLE = UI_TEXT["main_description"]

BUTTON_ADD = UI_TEXT["button_add"]
BUTTON_FOLDER = UI_TEXT["button_select_folder"]
BUTTON_REFRESH = UI_TEXT["button_refresh"]
BUTTON_CANCEL = UI_TEXT["button_cancel"]
BUTTON_MERGE = UI_TEXT["button_execute"]
BUTTON_MOVE_UP = UI_TEXT["button_move_up"]
BUTTON_MOVE_DOWN = UI_TEXT["button_move_down"]
BUTTON_DELETE = UI_TEXT["button_delete"]

ROW2_GUIDE = UI_TEXT["drag_hint"]

LABEL_SAVE_FOLDER = UI_TEXT["label_save_folder"]
LABEL_PAGE_COUNT_UNKNOWN = UI_TEXT["label_page_count_unknown"]
LABEL_PAGE_SUFFIX = UI_TEXT["label_page_suffix"]
LABEL_LOADING_THUMBNAIL = UI_TEXT["label_loading_thumbnail"]

EMPTY_TITLE_DEFAULT = UI_TEXT["empty_title"]
EMPTY_TITLE_DROP = UI_TEXT["empty_title_drop"]
EMPTY_SUBTITLE = UI_TEXT["empty_subtitle"]

STATUS_LOADING = UI_TEXT["status_loading"]
STATUS_PROCESSING = UI_TEXT["status_processing"]
STATUS_SAVING = UI_TEXT["status_saving"]
STATUS_READY = UI_TEXT["status_ready"]
STATUS_NONE = UI_TEXT["status_idle"]
STATUS_CANCELING = UI_TEXT["status_canceling"]
STATUS_CANCELED = UI_TEXT["status_canceled"]
STATUS_SAVE_DONE = UI_TEXT["status_complete"]
STATUS_ERROR = UI_TEXT["status_error"]

DETAIL_READY = UI_TEXT["detail_ready"]
DETAIL_NONE = UI_TEXT["detail_none"]
DETAIL_PROCESSING = UI_TEXT["detail_processing"]
DETAIL_SAVING = UI_TEXT["detail_saving"]
DETAIL_SAVE_DONE = UI_TEXT["detail_save_done"]
DETAIL_CANCEL = UI_TEXT["detail_cancel"]
DETAIL_ERROR = UI_TEXT["detail_error"]
DETAIL_FILE_ERROR = UI_TEXT["detail_file_error"]

MSG_REFRESH_BLOCKED = UI_TEXT["msg_refresh_blocked"]
MSG_NO_FILES = UI_TEXT["msg_no_files"]
MSG_SAVE_FOLDER_ERROR_TITLE = UI_TEXT["msg_save_folder_error_title"]
MSG_SAVE_FOLDER_ERROR = UI_TEXT["msg_save_folder_error"]
MSG_SAVE_DONE = UI_TEXT["msg_save_done"]


@dataclass(frozen=True)
class WindowGeometry:
    width: int
    height: int
    x: int
    y: int
    min_width: int
    min_height: int


def calculate_initial_window_geometry(screen_width: int, screen_height: int) -> WindowGeometry:
    screen_width = max(1, int(screen_width))
    screen_height = max(1, int(screen_height))
    horizontal_margin = min(80, max(24, screen_width // 20))
    vertical_margin = min(120, max(40, screen_height // 10))
    available_width = max(1, screen_width - horizontal_margin)
    available_height = max(1, screen_height - vertical_margin)
    width = min(PREFERRED_WINDOW_WIDTH, available_width)
    height = min(PREFERRED_WINDOW_HEIGHT, available_height)
    x = max(0, (screen_width - width) // 2)
    y = max(0, (screen_height - height) // 2)
    return WindowGeometry(
        width=width,
        height=height,
        x=x,
        y=y,
        min_width=min(PREFERRED_MIN_WIDTH, width),
        min_height=min(PREFERRED_MIN_HEIGHT, height),
    )


def app_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def resource_dir() -> str:
    if getattr(sys, "frozen", False):
        return getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def make_root():
    if HAS_DND:
        return TkinterDnD.Tk()
    return tk.Tk()


def detect_font_name(root: tk.Misc):
    families = set(root.tk.call("font", "families"))
    for name in FONT_CANDIDATES:
        if name in families:
            return name
    return "TkDefaultFont"


FONT_NAME = None


def ensure_font_name(root: tk.Misc):
    global FONT_NAME
    if FONT_NAME is None:
        FONT_NAME = detect_font_name(root)
    return FONT_NAME


def get_pdf_reader_writer():
    global _PDF_READER, _PDF_WRITER
    if _PDF_READER is None or _PDF_WRITER is None:
        with _PDF_IMPORT_LOCK:
            if _PDF_READER is None or _PDF_WRITER is None:
                from pypdf import PdfReader, PdfWriter

                _PDF_READER = PdfReader
                _PDF_WRITER = PdfWriter
    return _PDF_READER, _PDF_WRITER


def get_fitz():
    global _FITZ, _FITZ_LOAD_ATTEMPTED
    if not _FITZ_LOAD_ATTEMPTED:
        with _PDF_IMPORT_LOCK:
            if not _FITZ_LOAD_ATTEMPTED:
                try:
                    import fitz  # PyMuPDF
                except Exception:
                    fitz = None
                _FITZ = fitz
                _FITZ_LOAD_ATTEMPTED = True
    return _FITZ


def icon_ico_path() -> str:
    if getattr(sys, "frozen", False):
        return os.path.join(getattr(sys, "_MEIPASS", os.path.dirname(sys.executable)), COMMON_ICON_FILENAME)
    return os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), COMMON_ICON_RELATIVE))


def apply_window_icon(window: tk.Misc):
    try:
        ico = icon_ico_path()
        if os.path.exists(ico):
            try:
                window.iconbitmap(ico)
            except Exception:
                pass
            try:
                window.iconbitmap(default=ico)
            except Exception:
                pass
            try:
                window.wm_iconbitmap(ico)
            except Exception:
                pass
    except Exception:
        pass


def set_windows_app_id():
    if not sys.platform.startswith("win") or ctypes is None:
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Shimarisu.DakePDFMerge")
    except Exception:
        pass


def config_path() -> str:
    return os.path.join(app_dir(), CONFIG_NAME)


def default_downloads() -> str:
    return os.path.join(os.path.expanduser("~"), "Downloads")


def load_config() -> dict:
    try:
        with open(config_path(), "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_config(data: dict) -> None:
    try:
        with open(config_path(), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


class CliError(Exception):
    pass


class AtomicSaveError(Exception):
    pass


def write_cli_error(message: str) -> None:
    try:
        sys.stderr.write(f"{message}\n")
        sys.stderr.flush()
    except Exception:
        pass


def parse_shimarisu_cli_args(argv: list[str]) -> tuple[list[str], str | None, bool]:
    inputs: list[str] = []
    output: str | None = None
    silent = False
    i = 0

    while i < len(argv):
        arg = argv[i]
        if arg == "--from-shimarisu":
            i += 1
            continue
        if arg == "--silent":
            silent = True
            i += 1
            continue
        if arg == "--output":
            i += 1
            if i >= len(argv) or argv[i].startswith("--"):
                raise CliError(UI_TEXT["cli_error_output_failed"])
            output = argv[i]
            i += 1
            continue
        if arg == "--inputs":
            i += 1
            while i < len(argv) and not argv[i].startswith("--"):
                inputs.append(argv[i])
                i += 1
            continue
        raise CliError(UI_TEXT["cli_error_merge_failed"])

    return inputs, output, silent


def validate_cli_inputs(inputs: list[str]) -> list[str]:
    if len(inputs) < 2:
        raise CliError(UI_TEXT["cli_error_not_enough_inputs"])

    validated: list[str] = []
    for path in inputs:
        full_path = os.path.abspath(path)
        if not os.path.isfile(full_path):
            raise CliError(UI_TEXT["cli_error_file_not_found"])
        if os.path.splitext(full_path)[1].lower() != ".pdf":
            raise CliError(UI_TEXT["cli_error_not_pdf"])
        validated.append(full_path)
    return validated


def unique_output_path(path: str, protected_paths: list[str]) -> str:
    protected = {os.path.normcase(os.path.abspath(p)) for p in protected_paths}
    output = os.path.abspath(path)
    stem, ext = os.path.splitext(output)
    n = 1

    while os.path.exists(output) or os.path.normcase(output) in protected:
        output = f"{stem}_{n:02d}{ext}"
        n += 1

    return output


def _commit_temp_file(temp_path: str, final_path: str) -> None:
    if os.name == "nt":
        os.rename(temp_path, final_path)
        return
    os.link(temp_path, final_path)
    os.remove(temp_path)


def write_pdf_atomically(
    writer,
    final_path: str,
    protected_paths: list[str],
    *,
    commit_func=None,
    before_commit=None,
) -> str:
    final_path = os.path.abspath(final_path)
    folder = os.path.dirname(final_path) or os.getcwd()
    os.makedirs(folder, exist_ok=True)
    protected = [os.path.abspath(path) for path in protected_paths]
    candidate = unique_output_path(final_path, protected)
    temp_path: str | None = None
    commit = commit_func or _commit_temp_file

    try:
        with tempfile.NamedTemporaryFile(
            mode="w+b",
            dir=folder,
            prefix=".dake_pdf_merge_",
            suffix=".tmp",
            delete=False,
        ) as stream:
            temp_path = stream.name
            writer.write(stream)
            stream.flush()
            os.fsync(stream.fileno())

        if before_commit is not None:
            before_commit()
        candidate = unique_output_path(candidate, protected)
        commit(temp_path, candidate)
        temp_path = None
        return candidate
    except Exception as exc:
        if temp_path:
            try:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
            except Exception:
                pass
        raise AtomicSaveError(UI_TEXT["msg_atomic_save_failed"]) from exc


def make_cli_output_path(inputs: list[str], output_arg: str | None) -> str:
    from datetime import datetime

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"merged_{timestamp}.pdf"

    if output_arg:
        output = os.path.abspath(output_arg)
        if os.path.splitext(output)[1].lower() == ".pdf":
            folder = os.path.dirname(output) or os.getcwd()
            os.makedirs(folder, exist_ok=True)
            return unique_output_path(output, inputs)

        os.makedirs(output, exist_ok=True)
        return unique_output_path(os.path.join(output, filename), inputs)

    folder = os.path.dirname(inputs[0]) or os.getcwd()
    return unique_output_path(os.path.join(folder, filename), inputs)


def merge_pdfs_for_cli(inputs: list[str], output: str) -> str:
    try:
        PdfReader, PdfWriter = get_pdf_reader_writer()
        writer = PdfWriter()
        for path in inputs:
            reader = PdfReader(path)
            for page in reader.pages:
                writer.add_page(page)

        return write_pdf_atomically(writer, output, inputs)
    except CliError:
        raise
    except Exception as e:
        raise CliError(UI_TEXT["cli_error_merge_failed"]) from e


def run_shimarisu_cli(argv: list[str]) -> int:
    try:
        inputs, output_arg, _silent = parse_shimarisu_cli_args(argv)
        validated_inputs = validate_cli_inputs(inputs)
        output = make_cli_output_path(validated_inputs, output_arg)
        merge_pdfs_for_cli(validated_inputs, output)
        return 0
    except CliError as e:
        write_cli_error(str(e))
        return 1
    except Exception:
        write_cli_error(UI_TEXT["cli_error_merge_failed"])
        return 1


def shorten_path(path: str, max_len: int = 56) -> str:
    if len(path) <= max_len:
        return path
    drive, rest = os.path.splitdrive(path)
    parts = rest.strip("\\/").split(os.sep)
    if len(parts) <= 2:
        return path[: max_len - 1] + "..."
    return f"{drive}\\...\\{parts[-2]}\\{parts[-1]}"


def open_folder(path: str) -> None:
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


def open_web_link(url: str) -> None:
    try:
        webbrowser.open(url)
    except Exception:
        pass


def bind_footer_link(label: tk.Label, url: str) -> None:
    def _on_enter(_event):
        label.configure(fg=ACCENT, cursor="hand2")

    def _on_leave(_event):
        label.configure(fg=SUBTEXT, cursor="")

    def _on_click(_event):
        open_web_link(url)

    label.bind("<Enter>", _on_enter)
    label.bind("<Leave>", _on_leave)
    label.bind("<Button-1>", _on_click)


def extract_drop_paths(raw: str) -> list[str]:
    try:
        parts = list(tk.Tcl().splitlist(raw))
    except Exception:
        parts = raw.split()
    cleaned = []
    for part in parts:
        part = part.strip().strip("{}")
        if part:
            cleaned.append(part)
    return cleaned


def format_card_filename(name: str, line_chars: int = 18, max_lines: int = 2) -> str:
    text = name.strip()
    max_chars = line_chars * max_lines
    if len(text) > max_chars:
        text = text[: max_chars - 3] + "..."
    if len(text) <= line_chars:
        return text
    lines = [text[i:i + line_chars] for i in range(0, len(text), line_chars)]
    if len(lines) > max_lines:
        lines = lines[:max_lines]
        lines[-1] = lines[-1][: max(0, line_chars - 3)] + "..."
    return "\n".join(lines[:max_lines])


class ModernButton(tk.Label):
    def __init__(self, parent, text, command, *, primary=False, width=12):
        self.command = command
        self.primary = primary
        self.enabled = True
        self.normal_bg = ACCENT if primary else "#FFFFFF"
        self.normal_fg = "#FFFFFF" if primary else TEXT
        self.hover_bg = ACCENT_HOVER if primary else "#F2F4F7"
        super().__init__(
            parent,
            text=text,
            bg=self.normal_bg,
            fg=self.normal_fg,
            padx=16,
            pady=9,
            cursor="hand2",
            bd=1,
            relief="solid",
            width=width,
            font=(FONT_NAME, 10, "bold"),
        )
        self.bind("<Button-1>", self._on_click)
        self.bind("<Enter>", self._hover_in)
        self.bind("<Leave>", self._hover_out)

    def _on_click(self, _event):
        if self.enabled:
            self.command()

    def _hover_in(self, _event):
        if self.enabled:
            self.configure(bg=self.hover_bg)

    def _hover_out(self, _event):
        self._apply_visual()

    def set_enabled(self, enabled: bool):
        self.enabled = enabled
        self._apply_visual()

    def _apply_visual(self):
        if self.enabled:
            self.configure(bg=self.normal_bg, fg=self.normal_fg, cursor="hand2")
        else:
            self.configure(bg=DISABLED_BG, fg=DISABLED_FG, cursor="arrow")


class MergeFileCard(tk.Frame):
    def __init__(self, parent, app, path: str, index: int):
        super().__init__(parent, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        self.app = app
        self.path = path
        self.index = index
        self.thumb_image = None
        self.top = None
        self.number_label = None
        self.thumb_frame = None
        self.thumb_label = None
        self.name_label = None
        self.meta = None
        self.buttons = []
        self.build_ui()
        self.bind_drag_handlers()

    def build_ui(self):
        self.top = tk.Frame(self, bg=CARD)
        self.top.pack(fill="x", padx=12, pady=(10, 6))

        self.number_label = tk.Label(
            self.top,
            text=f"{self.index + 1:02d}",
            font=(FONT_NAME, 11, "bold"),
            bg=CARD,
            fg=ACCENT,
        )
        self.number_label.pack(side="left", padx=(0, 8))

        btns = tk.Frame(self.top, bg=CARD)
        btns.pack(side="right")

        actions = [
            (BUTTON_MOVE_UP, lambda: self.app.move_file(self.index, -1)),
            (BUTTON_MOVE_DOWN, lambda: self.app.move_file(self.index, 1)),
            (BUTTON_DELETE, lambda: self.app.remove_file(self.index)),
        ]
        for label, cmd in actions:
            button = tk.Label(
                btns,
                text=label,
                bg="#FFFFFF",
                fg=TEXT,
                padx=8,
                pady=5,
                cursor="hand2",
                bd=1,
                relief="solid",
                font=(FONT_NAME, 9, "bold"),
            )
            button.pack(side="left", padx=3)
            button.bind("<Button-1>", lambda _e, c=cmd: c())
            button.bind("<Enter>", lambda _e, w=button: w.configure(bg="#F2F4F7"))
            button.bind("<Leave>", lambda _e, w=button: w.configure(bg="#FFFFFF"))
            self.buttons.append(button)

        self.thumb_frame = tk.Frame(
            self,
            bg=PREVIEW_BG,
            highlightthickness=1,
            highlightbackground=PREVIEW_BORDER,
        )
        self.thumb_frame.pack(padx=12, pady=(0, 8))

        self.thumb_label = tk.Label(
            self.thumb_frame,
            text=LABEL_LOADING_THUMBNAIL,
            bg="#FFFFFF",
            fg=SUBTEXT,
            width=18,
            height=14,
            font=(FONT_NAME, 9),
        )
        self.thumb_label.pack(padx=1, pady=1)

        self.name_label = tk.Label(
            self,
            text=format_card_filename(os.path.basename(self.path)),
            bg=CARD,
            fg=TEXT,
            font=(FONT_NAME, 10, "bold"),
            wraplength=156,
            justify="left",
            anchor="nw",
            height=2,
        )
        self.name_label.pack(fill="x", padx=12)

        self.meta = tk.Label(
            self,
            text=LABEL_PAGE_COUNT_UNKNOWN,
            bg=CARD,
            fg=SUBTEXT,
            font=(FONT_NAME, 8),
            anchor="w",
        )
        self.meta.pack(fill="x", padx=12, pady=(3, 10))

    def set_page_count(self, count: int | None):
        self.meta.configure(
            text=f"{count}{LABEL_PAGE_SUFFIX}" if count is not None else LABEL_PAGE_COUNT_UNKNOWN
        )

    def set_thumbnail(self, img):
        self.thumb_image = img
        self.thumb_label.configure(image=img, text="", width=160, height=220, bg="#FFFFFF")

    def update_index(self, index: int):
        self.index = index
        self.number_label.configure(text=f"{index + 1:02d}")

    def update_visual(self):
        self.configure(bg=CARD, highlightbackground=BORDER)
        self.top.configure(bg=CARD)
        self.number_label.configure(bg=CARD, fg=ACCENT)
        self.thumb_frame.configure(highlightbackground=PREVIEW_BORDER)
        self.name_label.configure(bg=CARD)
        self.meta.configure(bg=CARD)
        for button in self.buttons:
            button.configure(bg="#FFFFFF")

    def set_drag_visual(self, *, source: bool = False, target: bool = False):
        if target:
            self.configure(highlightbackground=ACCENT, highlightthickness=2)
        elif source:
            self.configure(highlightbackground=PREVIEW_BORDER, highlightthickness=2)
        else:
            self.configure(highlightbackground=BORDER, highlightthickness=1)

    def bind_drag_handlers(self):
        drag_widgets = (
            self,
            self.top,
            self.number_label,
            self.thumb_frame,
            self.thumb_label,
            self.name_label,
            self.meta,
        )
        for widget in drag_widgets:
            widget.bind("<ButtonPress-1>", self._on_drag_start, add="+")
            widget.bind("<B1-Motion>", self._on_drag_motion, add="+")
            widget.bind("<ButtonRelease-1>", self._on_drag_release, add="+")

    def _on_drag_start(self, event):
        self.app.begin_card_drag(self.index, event)

    def _on_drag_motion(self, event):
        self.app.update_card_drag(event)

    def _on_drag_release(self, event):
        self.app.finish_card_drag(event)


class DAKEPDFMergeApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.initial_geometry = calculate_initial_window_geometry(
            self.root.winfo_screenwidth(),
            self.root.winfo_screenheight(),
        )
        self.root.geometry(
            f"{self.initial_geometry.width}x{self.initial_geometry.height}"
            f"+{self.initial_geometry.x}+{self.initial_geometry.y}"
        )
        self.root.minsize(self.initial_geometry.min_width, self.initial_geometry.min_height)
        self.root.configure(bg=BG)

        self.cfg = load_config()
        self.save_folder = self.cfg.get("last_folder", default_downloads())
        self.files: list[str] = []
        self.page_count_cache: dict[str, int | None] = {}
        self.thumbnail_cache: dict[str, bytes | None] = {}
        self.merge_card_by_path: dict[str, MergeFileCard] = {}
        self.thumbnail_job_queue: queue.Queue = queue.Queue()
        self.thumbnail_queue: queue.Queue = queue.Queue()
        self.thumbnail_jobs_pending: set[str] = set()
        self._thumbnail_workers_started = False
        self.ui_queue: queue.Queue = queue.Queue()
        self.worker_running = False
        self.task_state = TASK_IDLE
        self.cancel_requested = False
        self.close_requested = False
        self.merge_thread: threading.Thread | None = None
        self._status_anim_job = None
        self._status_anim_base = STATUS_PROCESSING
        self._status_anim_dots = 0
        self._status_anim_active = False
        self._complete_reset_job = None
        self.card_ui_loading = False
        self.pending_card_paths: set[str] = set()
        self._card_ui_finish_scheduled = False

        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")
        except Exception:
            pass
        self.style.configure(
            "Horizontal.TProgressbar",
            troughcolor="#E9EEF7",
            bordercolor="#E9EEF7",
            lightcolor=ACCENT,
            darkcolor=ACCENT,
            background=ACCENT,
            thickness=16,
        )

        self.action_buttons: list[ModernButton] = []
        self.list_area = None
        self.empty_outer = None
        self.empty_panel = None
        self.empty_content = None
        self.empty_title_label = None
        self.empty_subtitle_label = None
        self.empty_start_label = None
        self.empty_drop_label = None
        self.empty_visible = False
        self.empty_drop_hover = False
        self.count_label = None
        self.drag_source_index = None
        self.drag_target_index = None
        self.drag_start_xy = None
        self.drag_started = False
        self.drag_pointer_xy = None
        self.drag_autoscroll_direction = 0
        self._drag_autoscroll_job = None
        self.layout_mode: str | None = None
        self.layout_switch_count = 0
        self.top_card = None
        self.footer = None
        self.footer_left = None
        self.footer_right = None

        self.build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.bind("<Configure>", self.on_root_configure, add="+")
        self.apply_layout_mode(
            "narrow" if self.initial_geometry.width < NARROW_LAYOUT_THRESHOLD else "wide"
        )
        self.root.after(120, self.process_thumbnail_queue)
        self.root.after(60, self.process_ui_queue)

    def build_ui(self):
        shell = tk.Frame(self.root, bg=BG)
        shell.pack(fill="both", expand=True)

        main = tk.Frame(shell, bg=BG)
        main.pack(fill="both", expand=True, padx=20, pady=(18, 8))

        title = tk.Frame(main, bg=BG)
        title.pack(fill="x", pady=(2, 10))

        tk.Label(
            title,
            text=MAIN_TITLE,
            font=(FONT_NAME, 20, "bold"),
            bg=BG,
            fg=TEXT,
        ).pack(side="left", anchor="w")

        tk.Label(
            title,
            text=MAIN_SUBTITLE,
            font=(FONT_NAME, 10),
            bg=BG,
            fg=SUBTEXT,
        ).pack(side="left", padx=(12, 0), pady=(6, 0))

        self.top_card = tk.Frame(
            main,
            bg="#FFFFFF",
            highlightthickness=1,
            highlightbackground=BORDER,
        )
        self.top_card.pack(fill="x", pady=(0, 12))

        self.add_button = ModernButton(
            self.top_card,
            BUTTON_ADD,
            self.add_files,
            primary=True,
            width=10,
        )

        self.folder_button = ModernButton(
            self.top_card,
            BUTTON_FOLDER,
            self.choose_folder,
            width=12,
        )

        self.refresh_button = ModernButton(
            self.top_card,
            BUTTON_REFRESH,
            self.reset_merge,
            width=10,
        )

        self.count_label = tk.Label(
            self.top_card,
            text="",
            font=(FONT_NAME, 9, "bold"),
            bg="#FFFFFF",
            fg=TEXT,
        )
        self.action_buttons.extend([self.add_button, self.folder_button, self.refresh_button])

        self.folder_short_label = tk.Label(
            self.top_card,
            text=f"{LABEL_SAVE_FOLDER}: {shorten_path(self.save_folder)}",
            font=(FONT_NAME, 9),
            bg="#FFFFFF",
            fg=SUBTEXT,
        )
        self.drag_hint_label = tk.Label(
            self.top_card,
            text=ROW2_GUIDE,
            font=(FONT_NAME, 9),
            bg="#FFFFFF",
            fg=SUBTEXT,
        )

        body = tk.Frame(main, bg=BG)
        body.pack(fill="both", expand=True)

        self.list_area = tk.Frame(body, bg=BG, highlightthickness=1, highlightbackground=BORDER)
        self.list_area.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(self.list_area, bg=BG, highlightthickness=0)
        self.scrollbar = tk.Scrollbar(self.list_area, orient="vertical", command=self.canvas.yview)
        self.cards_wrap = tk.Frame(self.canvas, bg=BG)
        self.cards_wrap.bind("<Configure>", lambda _e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas_window = self.canvas.create_window((0, 0), window=self.cards_wrap, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.bind("<Configure>", self.on_canvas_resize)
        self.bind_mousewheel(self.canvas)
        self.bind_mousewheel(self.cards_wrap)
        self.build_empty_state()

        bottom = tk.Frame(main, bg=BG)
        bottom.pack(fill="x", pady=(12, 0))

        left = tk.Frame(bottom, bg=BG)
        left.pack(side="left", fill="x", expand=True, anchor="n")

        self.progress = ttk.Progressbar(left, orient="horizontal", mode="determinate")
        self.progress.pack(anchor="nw", fill="x", expand=True)

        self.progress_label = tk.Label(
            left,
            text=STATUS_NONE,
            font=(FONT_NAME, 10, "bold"),
            bg=BG,
            fg=ACCENT,
        )
        self.progress_label.pack(anchor="nw", pady=(6, 0))

        right = tk.Frame(bottom, bg=BG)
        right.pack(side="right", anchor="n", padx=(14, 0))

        self.cancel_button = ModernButton(right, BUTTON_CANCEL, self.cancel_task, width=10)
        self.cancel_button.pack(side="left", padx=(0, 10))

        self.merge_button = ModernButton(
            right,
            BUTTON_MERGE,
            self.merge_files,
            primary=True,
            width=14,
        )
        self.merge_button.pack(side="left")
        self.action_buttons.append(self.merge_button)

        launcher_link_row = tk.Frame(main, bg=BG)
        launcher_link_row.pack(fill="x", pady=(6, 0))
        launcher_link = tk.Label(
            launcher_link_row,
            text=UI_TEXT["link_other_tools"],
            font=(FONT_NAME, 9),
            bg=BG,
            fg=SUBTEXT,
        )
        launcher_link.pack(side="right")
        bind_footer_link(launcher_link, LAUNCHER_LP_URL)

        self.footer = tk.Frame(shell, bg=BG)
        self.footer.pack(fill="x", padx=24, pady=(0, 10))

        self.footer_left = tk.Frame(self.footer, bg=BG)
        self.footer_right = tk.Frame(self.footer, bg=BG)

        tk.Label(
            self.footer_left,
            text=UI_TEXT["footer_left"],
            font=(FONT_NAME, 9),
            bg=BG,
            fg=SUBTEXT,
        ).pack(side="left")

        footer_assessment = tk.Label(
            self.footer_right,
            text=UI_TEXT["footer_link_1"],
            font=(FONT_NAME, 9),
            bg=BG,
            fg=SUBTEXT,
        )
        footer_assessment.pack(side="left")
        bind_footer_link(
            footer_assessment,
            "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
        )

        tk.Label(
            self.footer_right,
            text=UI_TEXT["footer_separator"],
            font=(FONT_NAME, 9),
            bg=BG,
            fg=SUBTEXT,
        ).pack(side="left")

        footer_instagram = tk.Label(
            self.footer_right,
            text=UI_TEXT["footer_link_2"],
            font=(FONT_NAME, 9),
            bg=BG,
            fg=SUBTEXT,
        )
        footer_instagram.pack(side="left")
        bind_footer_link(
            footer_instagram,
            "https://instagram.com/kikuta.shimarisu_fudosan",
        )

        tk.Label(
            self.footer_right,
            text=f"{UI_TEXT['footer_separator']}{UI_TEXT['footer_copyright']}",
            font=(FONT_NAME, 9),
            bg=BG,
            fg=SUBTEXT,
        ).pack(side="left")

        if HAS_DND:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind("<<DropEnter>>", self.on_drop_enter)
            self.root.dnd_bind("<<DropLeave>>", self.on_drop_leave)
            self.root.dnd_bind("<<Drop>>", self.on_drop)

        self.refresh_merge_cards()
        self.set_processing_state(False)

    def on_root_configure(self, event):
        if event.widget is not self.root:
            return
        mode = "narrow" if event.width < NARROW_LAYOUT_THRESHOLD else "wide"
        self.apply_layout_mode(mode)

    def apply_layout_mode(self, mode: str):
        if mode not in {"wide", "narrow"} or mode == self.layout_mode:
            return
        self.layout_mode = mode
        self.layout_switch_count += 1
        self._layout_top_card(mode)
        self._layout_footer(mode)
        if self.folder_short_label is not None:
            self.refresh_status()

    def _layout_top_card(self, mode: str):
        widgets = (
            self.add_button,
            self.folder_button,
            self.refresh_button,
            self.count_label,
            self.folder_short_label,
            self.drag_hint_label,
        )
        for widget in widgets:
            widget.grid_forget()
        for column in range(5):
            self.top_card.grid_columnconfigure(column, weight=0)
        self.top_card.grid_columnconfigure(4, weight=1)

        self.add_button.grid(row=0, column=0, padx=(14, 0), pady=(12, 8), sticky="w")
        self.folder_button.grid(row=0, column=1, padx=8, pady=(12, 8), sticky="w")
        self.refresh_button.grid(row=0, column=2, padx=(0, 8), pady=(12, 8), sticky="w")
        self.count_label.grid(row=0, column=3, padx=(8, 0), pady=(12, 8), sticky="w")

        if mode == "wide":
            self.folder_short_label.grid(
                row=0,
                column=4,
                padx=(16, 14),
                pady=(12, 8),
                sticky="e",
            )
            self.drag_hint_label.grid(
                row=1,
                column=0,
                columnspan=5,
                padx=14,
                pady=(0, 12),
                sticky="w",
            )
        else:
            self.folder_short_label.grid(
                row=1,
                column=0,
                columnspan=5,
                padx=14,
                pady=(0, 6),
                sticky="w",
            )
            self.drag_hint_label.grid(
                row=2,
                column=0,
                columnspan=5,
                padx=14,
                pady=(0, 12),
                sticky="w",
            )

    def _layout_footer(self, mode: str):
        self.footer_left.grid_forget()
        self.footer_right.grid_forget()
        self.footer.grid_columnconfigure(0, weight=1)
        self.footer.grid_columnconfigure(1, weight=0)
        if mode == "wide":
            self.footer_left.grid(row=0, column=0, sticky="w")
            self.footer_right.grid(row=0, column=1, sticky="e")
        else:
            self.footer.grid_columnconfigure(1, weight=0)
            self.footer_left.grid(row=0, column=0, columnspan=2, pady=(0, 4))
            self.footer_right.grid(row=1, column=0, columnspan=2)

    def bind_mousewheel(self, widget):
        def _on_mousewheel(event):
            delta = -1 * int(event.delta / 120) if event.delta else 0
            if delta:
                self.canvas.yview_scroll(delta, "units")

        widget.bind("<Enter>", lambda _e: self.root.bind_all("<MouseWheel>", _on_mousewheel))
        widget.bind("<Leave>", lambda _e: self.root.unbind_all("<MouseWheel>"))

    def cancel_complete_reset(self):
        if self._complete_reset_job is not None:
            try:
                self.root.after_cancel(self._complete_reset_job)
            except Exception:
                pass
            self._complete_reset_job = None

    def schedule_complete_reset(self):
        self.cancel_complete_reset()
        self._complete_reset_job = self.root.after(1400, self.restore_after_complete)

    def restore_after_complete(self):
        self._complete_reset_job = None
        self.refresh_status()
        self.refresh_bottom_status()

    def set_top_status(self, text: str, *, color: str = TEXT):
        self.count_label.configure(text=text, fg=color)

    def set_bottom_status(self, title: str, detail: str, *, color: str = ACCENT):
        self.progress_label.configure(text=title, fg=color)

    def refresh_bottom_status(self):
        self.cancel_complete_reset()
        if self.worker_running:
            return

        if self.card_ui_loading:
            self.set_bottom_status(STATUS_LOADING, "", color=ACCENT)
            if not self._status_anim_active or self._status_anim_base != STATUS_LOADING:
                self.start_status_animation(STATUS_LOADING)
            return

        self.stop_status_animation()

        if self.files:
            self.progress["value"] = 0
            self.set_bottom_status(STATUS_READY, DETAIL_READY, color=ACCENT)
        else:
            self.progress["value"] = 0
            self.set_bottom_status(STATUS_NONE, DETAIL_NONE, color=ACCENT)

    def finish_card_ui_loading(self):
        self._card_ui_finish_scheduled = False
        self.card_ui_loading = False
        self.pending_card_paths.clear()
        self.refresh_bottom_status()

    def sync_empty_state_view(self):
        if self.empty_panel is None:
            return

        self.empty_panel.configure(bg="#FFFFFF", highlightbackground=PREVIEW_BORDER)

        self.empty_title_label.configure(
            text=EMPTY_TITLE_DROP if self.empty_drop_hover and not self.files else EMPTY_TITLE_DEFAULT,
            bg="#FFFFFF",
            fg=TEXT,
        )
        self.empty_subtitle_label.configure(
            text=EMPTY_SUBTITLE,
            bg="#FFFFFF",
            fg=SUBTEXT,
        )

        if self.empty_visible:
            self.empty_outer.lift()

    def build_empty_state(self):
        if self.empty_outer is not None:
            return

        self.empty_outer = tk.Frame(self.list_area, bg=BG, highlightthickness=0)
        self.empty_outer.pack_propagate(False)

        self.empty_panel = tk.Frame(
            self.empty_outer,
            bg="#FFFFFF",
            highlightthickness=1,
            highlightbackground=PREVIEW_BORDER,
        )
        self.empty_panel.pack(fill="both", expand=True)
        self.empty_panel.pack_propagate(False)

        self.empty_content = tk.Frame(self.empty_panel, bg="#FFFFFF")
        self.empty_content.place(relx=0.5, rely=0.5, anchor="center")

        self.empty_title_label = tk.Label(
            self.empty_content,
            text=EMPTY_TITLE_DEFAULT,
            font=(FONT_NAME, 13),
            bg="#FFFFFF",
            fg=TEXT,
        )
        self.empty_title_label.pack()

        self.empty_subtitle_label = tk.Label(
            self.empty_content,
            text=EMPTY_SUBTITLE,
            font=(FONT_NAME, 10),
            bg="#FFFFFF",
            fg=SUBTEXT,
        )
        self.empty_subtitle_label.pack(pady=(8, 0))

        for widget in (
            self.empty_outer,
            self.empty_panel,
            self.empty_content,
            self.empty_title_label,
            self.empty_subtitle_label,
        ):
            widget.bind("<Button-1>", self.on_empty_state_click)
            widget.configure(cursor="hand2")

        self.empty_start_label = None
        self.empty_drop_label = None
        self.sync_empty_state_view()
        self.hide_empty_state()

    def show_empty_state(self):
        if self.empty_outer is None:
            self.build_empty_state()

        self.empty_drop_hover = False
        self.empty_outer.place(relx=0.5, rely=0.5, anchor="center", width=520, height=190)
        self.empty_outer.lift()
        self.empty_visible = True
        self.sync_empty_state_view()

    def hide_empty_state(self):
        if self.empty_outer is None:
            return
        self.empty_drop_hover = False
        self.empty_outer.place_forget()
        self.empty_visible = False

    def on_drop_enter(self, event):
        if self.empty_visible and not self.files:
            self.empty_drop_hover = True
            self.sync_empty_state_view()
        return getattr(event, "action", None)

    def on_drop_leave(self, event):
        if self.empty_visible:
            self.empty_drop_hover = False
            self.sync_empty_state_view()
        return getattr(event, "action", None)

    def on_empty_state_click(self, _event=None):
        if self.worker_running or self.files or not self.empty_visible:
            return
        self.add_files()

    def start_status_animation(self, base_text: str):
        self.stop_status_animation()
        self._status_anim_active = True
        self._status_anim_base = base_text
        self._status_anim_dots = 0
        self._run_status_animation()

    def _run_status_animation(self):
        if not self._status_anim_active:
            return
        dots = "." * (self._status_anim_dots % 4)
        self.progress_label.configure(text=f"{self._status_anim_base}{dots}", fg=ACCENT)
        self._status_anim_dots += 1
        self._status_anim_job = self.root.after(800, self._run_status_animation)

    def stop_status_animation(self):
        self._status_anim_active = False
        if self._status_anim_job is not None:
            try:
                self.root.after_cancel(self._status_anim_job)
            except Exception:
                pass
            self._status_anim_job = None

    def set_task_state(self, state: str):
        if state not in {TASK_IDLE, TASK_PROCESSING, TASK_SAVING}:
            raise ValueError(f"Unknown task state: {state}")
        self.task_state = state
        self.worker_running = state != TASK_IDLE
        if self.worker_running:
            self.cancel_complete_reset()
        for button in self.action_buttons:
            button.set_enabled(state == TASK_IDLE)
        self.cancel_button.set_enabled(state == TASK_PROCESSING)
        if state == TASK_IDLE:
            self.stop_status_animation()
            self.progress_label.configure(fg=ACCENT)

    def set_processing_state(self, processing: bool):
        self.set_task_state(TASK_PROCESSING if processing else TASK_IDLE)

    def choose_folder(self):
        if self.worker_running:
            return
        path = filedialog.askdirectory(initialdir=self.save_folder or default_downloads())
        if path:
            self.save_folder = path
            self.cfg["last_folder"] = path
            save_config(self.cfg)
            self.refresh_status()

    def add_files(self):
        if self.worker_running:
            return
        paths = filedialog.askopenfilenames(filetypes=[(UI_TEXT["filetype_pdf"], "*.pdf")])
        self.add_pdf_paths(paths)

    def add_pdf_paths(self, paths):
        new_paths = []
        known_paths = {os.path.normcase(os.path.normpath(path)) for path in self.files}
        for raw in paths:
            path = os.path.abspath(str(raw))
            if not path.lower().endswith(".pdf"):
                continue
            if not os.path.isfile(path):
                continue
            path_key = os.path.normcase(os.path.normpath(path))
            if path_key in known_paths:
                continue
            known_paths.add(path_key)
            new_paths.append(path)

        if new_paths:
            self.files.extend(new_paths)
            self.card_ui_loading = True
            self.pending_card_paths.update(new_paths)
            self._card_ui_finish_scheduled = False
            self.set_bottom_status(STATUS_LOADING, "", color=ACCENT)
            self.start_status_animation(STATUS_LOADING)
            self.refresh_merge_cards()
        else:
            self.refresh_bottom_status()

    def on_drop(self, event):
        if self.worker_running:
            return
        if self.empty_visible:
            self.empty_drop_hover = False
            self.sync_empty_state_view()
        paths = extract_drop_paths(event.data)
        self.add_pdf_paths(paths)

    def reset_merge(self):
        if self.worker_running:
            messagebox.showinfo(APP_NAME, MSG_REFRESH_BLOCKED)
            return

        self.cancel_complete_reset()
        self.stop_status_animation()
        self.clear_card_drag_state()
        self.files.clear()
        self.page_count_cache.clear()
        self.thumbnail_cache.clear()
        self.pending_card_paths.clear()
        self.progress["value"] = 0
        self.refresh_merge_cards()

    def remove_file(self, index: int):
        if self.worker_running:
            return
        if 0 <= index < len(self.files):
            path = self.files.pop(index)
            self.page_count_cache.pop(path, None)
            self.thumbnail_cache.pop(path, None)
            self.pending_card_paths.discard(path)
            self.refresh_merge_cards()

    def move_file(self, index: int, delta: int):
        if self.worker_running:
            return
        new_index = index + delta
        if 0 <= index < len(self.files) and 0 <= new_index < len(self.files):
            self.files[index], self.files[new_index] = self.files[new_index], self.files[index]
            self.reflow_cards()
            self.refresh_status()
            self.refresh_bottom_status()

    def begin_card_drag(self, index: int, event):
        if self.worker_running or not (0 <= index < len(self.files)):
            return
        self.stop_drag_autoscroll()
        self.drag_source_index = index
        self.drag_target_index = index
        self.drag_start_xy = (event.x_root, event.y_root)
        self.drag_pointer_xy = (event.x_root, event.y_root)
        self.drag_started = False

    def update_card_drag(self, event):
        if self.drag_source_index is None or self.worker_running:
            return

        if not self.drag_started and self.drag_start_xy is not None:
            start_x, start_y = self.drag_start_xy
            if abs(event.x_root - start_x) < 6 and abs(event.y_root - start_y) < 6:
                return
            self.drag_started = True

        self.drag_pointer_xy = (event.x_root, event.y_root)
        self.update_drag_autoscroll(event.x_root, event.y_root)
        target_index = self.find_card_index_at(event.x_root, event.y_root)
        if target_index is not None:
            self.drag_target_index = target_index
        self.update_card_drag_visuals()

    def finish_card_drag(self, event):
        if self.drag_source_index is None:
            return

        source_index = self.drag_source_index
        target_index = self.find_card_index_at(event.x_root, event.y_root)
        if target_index is None:
            target_index = self.drag_target_index

        was_dragged = self.drag_started
        self.clear_card_drag_state()

        if was_dragged and target_index is not None and target_index != source_index:
            self.reorder_file(source_index, target_index)

    def find_card_index_at(self, root_x: int, root_y: int) -> int | None:
        if not self.files or self.list_area is None:
            return None

        area_x = self.list_area.winfo_rootx()
        area_y = self.list_area.winfo_rooty()
        area_w = self.list_area.winfo_width()
        area_h = self.list_area.winfo_height()
        margin = 40
        if (
            root_x < area_x - margin
            or root_x > area_x + area_w + margin
            or root_y < area_y - margin
            or root_y > area_y + area_h + margin
        ):
            return None

        nearest_index = None
        nearest_distance = None
        for path in self.files:
            card = self.merge_card_by_path.get(path)
            if card is None:
                continue

            x1 = card.winfo_rootx()
            y1 = card.winfo_rooty()
            x2 = x1 + card.winfo_width()
            y2 = y1 + card.winfo_height()
            if x1 <= root_x <= x2 and y1 <= root_y <= y2:
                return card.index

            center_x = x1 + card.winfo_width() / 2
            center_y = y1 + card.winfo_height() / 2
            distance = (root_x - center_x) ** 2 + (root_y - center_y) ** 2
            if nearest_distance is None or distance < nearest_distance:
                nearest_distance = distance
                nearest_index = card.index

        return nearest_index

    def update_card_drag_visuals(self):
        for path in self.files:
            card = self.merge_card_by_path.get(path)
            if card is None:
                continue
            card.set_drag_visual(
                source=card.index == self.drag_source_index,
                target=card.index == self.drag_target_index,
            )

    def clear_card_drag_state(self):
        self.stop_drag_autoscroll()
        self.drag_source_index = None
        self.drag_target_index = None
        self.drag_start_xy = None
        self.drag_pointer_xy = None
        self.drag_started = False
        for card in self.merge_card_by_path.values():
            card.set_drag_visual()

    def update_drag_autoscroll(self, root_x: int, root_y: int):
        if not self.drag_started or self.drag_source_index is None:
            self.stop_drag_autoscroll()
            return

        canvas_x = self.canvas.winfo_rootx()
        canvas_y = self.canvas.winfo_rooty()
        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()
        margin = 24
        direction = 0
        if canvas_x - margin <= root_x <= canvas_x + canvas_w + margin:
            if canvas_y - margin <= root_y <= canvas_y + DRAG_AUTOSCROLL_EDGE:
                direction = -1
            elif canvas_y + canvas_h - DRAG_AUTOSCROLL_EDGE <= root_y <= canvas_y + canvas_h + margin:
                direction = 1

        if direction == 0:
            self.stop_drag_autoscroll()
            return

        self.drag_autoscroll_direction = direction
        if self._drag_autoscroll_job is None:
            self._drag_autoscroll_job = self.root.after(
                DRAG_AUTOSCROLL_INTERVAL_MS,
                self.run_drag_autoscroll,
            )

    def run_drag_autoscroll(self):
        self._drag_autoscroll_job = None
        if (
            not self.drag_started
            or self.drag_source_index is None
            or self.drag_autoscroll_direction == 0
        ):
            return

        before = self.canvas.yview()
        self.canvas.yview_scroll(self.drag_autoscroll_direction, "units")
        after = self.canvas.yview()
        if before == after:
            self.stop_drag_autoscroll()
            return

        if self.drag_pointer_xy is not None:
            root_x, root_y = self.drag_pointer_xy
            target_index = self.find_card_index_at(root_x, root_y)
            if target_index is not None:
                self.drag_target_index = target_index
                self.update_card_drag_visuals()

        self._drag_autoscroll_job = self.root.after(
            DRAG_AUTOSCROLL_INTERVAL_MS,
            self.run_drag_autoscroll,
        )

    def stop_drag_autoscroll(self):
        self.drag_autoscroll_direction = 0
        if self._drag_autoscroll_job is not None:
            try:
                self.root.after_cancel(self._drag_autoscroll_job)
            except Exception:
                pass
            self._drag_autoscroll_job = None

    def reorder_file(self, source_index: int, target_index: int):
        if self.worker_running:
            return
        if not (0 <= source_index < len(self.files)):
            return

        target_index = max(0, min(target_index, len(self.files) - 1))
        path = self.files.pop(source_index)
        self.files.insert(target_index, path)
        self.reflow_cards()
        self.refresh_status()
        self.set_bottom_status(UI_TEXT["status_reordered"], "", color=ACCENT)
        self.schedule_complete_reset()

    def get_page_count(self, path: str) -> int | None:
        if path in self.page_count_cache:
            return self.page_count_cache[path]
        try:
            PdfReader, _PdfWriter = get_pdf_reader_writer()
            count = len(PdfReader(path).pages)
            return count
        except Exception:
            return None

    def start_thumbnail_workers(self):
        if self._thumbnail_workers_started:
            return
        self._thumbnail_workers_started = True
        for index in range(THUMBNAIL_WORKER_COUNT):
            threading.Thread(
                target=self.thumbnail_worker,
                name=f"pdf-preview-{index + 1}",
                daemon=True,
            ).start()

    def thumbnail_worker(self):
        while True:
            path = self.thumbnail_job_queue.get()
            try:
                try:
                    payload, page_count = self.load_pdf_preview(path)
                except Exception:
                    payload, page_count = None, None
                self.thumbnail_queue.put((path, payload, page_count))
            finally:
                self.thumbnail_job_queue.task_done()

    def load_pdf_preview(self, path: str) -> tuple[bytes | None, int | None]:
        payload = None
        page_count = None
        fitz_module = get_fitz()

        if fitz_module is not None:
            try:
                with fitz_module.open(path) as doc:
                    page_count = len(doc)
                    if page_count:
                        page = doc.load_page(0)
                        rect = page.rect
                        target_w, target_h = 160, 220
                        scale = min(target_w / rect.width, target_h / rect.height)
                        pix = page.get_pixmap(
                            matrix=fitz_module.Matrix(scale, scale),
                            alpha=False,
                        )
                        payload = pix.tobytes("ppm")
            except Exception:
                payload = None
                page_count = None

        if page_count is None:
            page_count = self.get_page_count(path)

        return payload, page_count

    def queue_thumbnail_job(self, path: str):
        if path in self.thumbnail_cache and path in self.page_count_cache:
            return
        if path in self.thumbnail_jobs_pending:
            return

        self.start_thumbnail_workers()
        self.thumbnail_jobs_pending.add(path)
        self.thumbnail_job_queue.put(path)

    def apply_cached_pdf_info(self, card: MergeFileCard, path: str):
        if path in self.page_count_cache:
            card.set_page_count(self.page_count_cache[path])
        payload = self.thumbnail_cache.get(path)
        if payload:
            card.set_thumbnail(tk.PhotoImage(data=payload))

    def process_thumbnail_queue(self):
        try:
            while True:
                path, payload, page_count = self.thumbnail_queue.get_nowait()
                self.thumbnail_jobs_pending.discard(path)
                if path not in self.files:
                    self.pending_card_paths.discard(path)
                    continue

                self.thumbnail_cache[path] = payload
                self.page_count_cache[path] = page_count
                card = self.merge_card_by_path.get(path)
                if card is None:
                    self.pending_card_paths.discard(path)
                    continue
                card.set_page_count(page_count)
                if payload:
                    image = tk.PhotoImage(data=payload)
                    card.set_thumbnail(image)
                self.pending_card_paths.discard(path)
        except queue.Empty:
            pass

        if self.card_ui_loading and self.files and not self.pending_card_paths and not self._card_ui_finish_scheduled:
            self._card_ui_finish_scheduled = True
            self.root.after_idle(self.finish_card_ui_loading)

        self.root.after(120, self.process_thumbnail_queue)

    def enqueue_ui_call(self, callback, *args, **kwargs):
        self.ui_queue.put((callback, args, kwargs))

    def process_ui_queue(self):
        try:
            while True:
                callback, args, kwargs = self.ui_queue.get_nowait()
                callback(*args, **kwargs)
        except queue.Empty:
            pass

        self.root.after(60, self.process_ui_queue)

    def refresh_status(self):
        count = len(self.files)
        if count:
            self.set_top_status(UI_TEXT["count_added"].format(count=count))
        else:
            self.set_top_status("")
        max_path_len = 42 if self.layout_mode == "narrow" else 56
        self.folder_short_label.configure(
            text=f"{LABEL_SAVE_FOLDER}: {shorten_path(self.save_folder, max_path_len)}"
        )

    def refresh_merge_cards(self):
        active_paths = set(self.files)
        for path in list(self.merge_card_by_path):
            if path in active_paths:
                continue
            card = self.merge_card_by_path.pop(path)
            card.destroy()
            self.page_count_cache.pop(path, None)
            self.thumbnail_cache.pop(path, None)
            self.pending_card_paths.discard(path)

        self.refresh_status()
        self.refresh_bottom_status()

        if not self.files:
            self.card_ui_loading = False
            self.pending_card_paths.clear()
            self._card_ui_finish_scheduled = False
            self.show_empty_state()
            self.refresh_bottom_status()
            return

        self.hide_empty_state()

        for i, path in enumerate(self.files):
            card = self.merge_card_by_path.get(path)
            if card is not None:
                continue
            card = MergeFileCard(self.cards_wrap, self, path, i)
            self.merge_card_by_path[path] = card
            card.update_visual()
            self.apply_cached_pdf_info(card, path)
            if path in self.thumbnail_cache and path in self.page_count_cache:
                self.pending_card_paths.discard(path)
            self.queue_thumbnail_job(path)

        self.reflow_cards()

        if self.card_ui_loading and not self.pending_card_paths and not self._card_ui_finish_scheduled:
            self._card_ui_finish_scheduled = True
            self.root.after_idle(self.finish_card_ui_loading)

    def on_canvas_resize(self, _event=None):
        try:
            width = max(self.canvas.winfo_width() - 2, 200)
            self.canvas.itemconfigure(self.canvas_window, width=width)
        except Exception:
            pass
        if self.empty_visible and self.empty_outer is not None:
            self.empty_outer.lift()
        self.reflow_cards()

    def reflow_cards(self):
        cards = [self.merge_card_by_path[path] for path in self.files if path in self.merge_card_by_path]
        if not cards:
            return
        width = max(self.canvas.winfo_width() - 24, 320)
        card_outer = 210
        cols = max(1, width // card_outer)
        while cols > 1 and ((cols * card_outer) + 12) > width:
            cols -= 1
        used = cols * card_outer
        extra = max(0, width - used)
        pad_x = max(4, extra // (cols * 2 + 2) if cols else 4)
        for i, card in enumerate(cards):
            card.update_index(i)
            card.grid(row=i // cols, column=i % cols, padx=pad_x, pady=8, sticky="n")

    def make_output_name(self) -> str:
        from datetime import datetime

        base = os.path.splitext(os.path.basename(self.files[0]))[0]
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        return f"{base}_{ts}_{len(self.files)}files_dake.pdf"

    def merge_files(self):
        if self.worker_running:
            return
        if not self.files:
            messagebox.showwarning(APP_NAME, MSG_NO_FILES)
            return

        folder = self.save_folder or default_downloads()
        try:
            os.makedirs(folder, exist_ok=True)
        except Exception as e:
            messagebox.showerror(MSG_SAVE_FOLDER_ERROR_TITLE, f"{MSG_SAVE_FOLDER_ERROR}\n{e}")
            return

        self.cfg["last_folder"] = folder
        save_config(self.cfg)

        base_name = self.make_output_name()
        output = unique_output_path(os.path.join(folder, base_name), self.files)

        self.cancel_requested = False
        self.close_requested = False
        self.set_task_state(TASK_PROCESSING)
        self.progress["value"] = 0
        self.set_bottom_status(STATUS_PROCESSING, DETAIL_PROCESSING, color=ACCENT)
        self.start_status_animation(STATUS_PROCESSING)
        files_snapshot = list(self.files)
        page_counts = [self.page_count_cache.get(path) for path in files_snapshot]
        self.merge_thread = threading.Thread(
            target=self._merge_worker,
            args=(output, files_snapshot, page_counts),
            daemon=False,
        )
        self.merge_thread.start()

    def _merge_worker(self, output: str, files: list[str], page_counts: list[int | None]):
        try:
            PdfReader, PdfWriter = get_pdf_reader_writer()
            writer = PdfWriter()
            total_files = len(files)
            known_page_counts = [
                count for count in page_counts if isinstance(count, int) and count >= 0
            ]
            use_page_progress = (
                len(known_page_counts) == len(page_counts)
                and sum(known_page_counts) > 0
            )
            total_pages = sum(known_page_counts) if use_page_progress else 0
            processed_pages = 0
            last_progress = -1

            for i, path in enumerate(files, start=1):
                if self.cancel_requested:
                    self.enqueue_ui_call(self.finish_cancel)
                    return
                try:
                    reader = PdfReader(path)
                    for page in reader.pages:
                        if self.cancel_requested:
                            self.enqueue_ui_call(self.finish_cancel)
                            return
                        writer.add_page(page)
                        processed_pages += 1
                        if use_page_progress:
                            progress = min(99, int(processed_pages / total_pages * 100))
                            if progress != last_progress:
                                last_progress = progress
                                self.enqueue_ui_call(self.update_progress, progress, "")
                except Exception as e:
                    self.enqueue_ui_call(self.handle_file_error, path, e)
                    return

                if not use_page_progress:
                    progress = min(99, int(i / total_files * 100))
                    self.enqueue_ui_call(self.update_progress, progress, "")

            if self.cancel_requested:
                self.enqueue_ui_call(self.finish_cancel)
                return

            self.task_state = TASK_SAVING
            self.enqueue_ui_call(self.begin_saving)
            final_output = write_pdf_atomically(
                writer,
                output,
                files,
                before_commit=lambda: self.enqueue_ui_call(self.show_finalizing),
            )

            self.enqueue_ui_call(self.finish_success, final_output)
        except Exception as e:
            self.enqueue_ui_call(self.finish_error, e)

    def update_progress(self, value: int, text: str):
        self.progress["value"] = value

    def cancel_task(self):
        if self.task_state == TASK_PROCESSING:
            self.cancel_requested = True
            self.set_bottom_status(STATUS_CANCELING, DETAIL_CANCEL, color=ACCENT)
            self.start_status_animation(STATUS_CANCELING)

    def begin_saving(self):
        self.set_task_state(TASK_SAVING)
        self.set_bottom_status(STATUS_SAVING, DETAIL_SAVING, color=ACCENT)
        self.start_status_animation(STATUS_SAVING)

    def show_finalizing(self):
        if self.task_state != TASK_SAVING:
            return
        self.stop_status_animation()
        self.set_bottom_status(UI_TEXT["detail_finalizing"], "", color=ACCENT)

    def finish_cancel(self):
        self.set_task_state(TASK_IDLE)
        self.progress["value"] = 0
        if self.close_requested:
            self._close_when_merge_worker_stops()
            return
        self.set_bottom_status(STATUS_CANCELED, DETAIL_CANCEL, color=ACCENT)
        self.schedule_complete_reset()

    def finish_success(self, output: str):
        self.set_task_state(TASK_IDLE)
        self.progress["value"] = 100
        self.set_bottom_status(STATUS_SAVE_DONE, DETAIL_SAVE_DONE, color=SUCCESS)
        self.root.lift()
        self.root.focus_force()
        messagebox.showinfo(APP_NAME, MSG_SAVE_DONE, parent=self.root)
        open_folder(os.path.dirname(output))
        self.schedule_complete_reset()

    def finish_error(self, error: Exception):
        self.set_task_state(TASK_IDLE)
        self.progress["value"] = 0
        self.refresh_status()
        self.set_bottom_status(STATUS_ERROR, DETAIL_ERROR, color=ERROR)
        if self.close_requested:
            self._close_when_merge_worker_stops()
            return
        if isinstance(error, AtomicSaveError):
            message = UI_TEXT["msg_atomic_save_failed"]
        else:
            message = f"{DETAIL_ERROR}\n{error}"
        messagebox.showerror(STATUS_ERROR, message, parent=self.root)

    def handle_file_error(self, path: str, error: Exception):
        self.set_task_state(TASK_IDLE)
        self.progress["value"] = 0
        self.refresh_status()
        self.set_bottom_status(STATUS_ERROR, DETAIL_FILE_ERROR, color=ERROR)

        if self.close_requested:
            self._close_when_merge_worker_stops()
            return

        messagebox.showerror(
            STATUS_ERROR,
            f"{DETAIL_FILE_ERROR}\n\n{path}\n\n{error}",
            parent=self.root,
        )

    def on_close(self):
        if self.task_state == TASK_SAVING:
            messagebox.showinfo(
                UI_TEXT["dialog_close_title"],
                UI_TEXT["dialog_close_saving"],
                parent=self.root,
            )
            return
        if self.task_state == TASK_PROCESSING:
            should_close = messagebox.askyesno(
                UI_TEXT["dialog_close_title"],
                UI_TEXT["dialog_close_processing"],
                parent=self.root,
            )
            if not should_close:
                return
            self.close_requested = True
            self.cancel_requested = True
            self.cancel_button.set_enabled(False)
            self.set_bottom_status(STATUS_CANCELING, DETAIL_CANCEL, color=ACCENT)
            self.start_status_animation(STATUS_CANCELING)
            return
        self._destroy_root()

    def _close_when_merge_worker_stops(self):
        if self.merge_thread is not None and self.merge_thread.is_alive():
            self.root.after(20, self._close_when_merge_worker_stops)
            return
        self._destroy_root()

    def _destroy_root(self):
        self.stop_drag_autoscroll()
        self.stop_status_animation()
        self.cancel_complete_reset()
        self.root.destroy()


def main():
    if "--from-shimarisu" in sys.argv[1:]:
        raise SystemExit(run_shimarisu_cli(sys.argv[1:]))

    set_windows_app_id()
    root = make_root()
    apply_window_icon(root)
    ensure_font_name(root)
    DAKEPDFMergeApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

