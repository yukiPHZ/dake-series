# -*- coding: utf-8 -*-
from __future__ import annotations

import ctypes
import io
import json
import os
import queue
import re
import sys
import tempfile
import threading
import time
import webbrowser
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

import tkinter as tk
from tkinter import filedialog, font as tkfont, ttk

try:
    import fitz  # type: ignore
except Exception:
    fitz = None

try:
    from PIL import Image, ImageGrab, ImageTk  # type: ignore
except Exception:
    Image = None
    ImageGrab = None
    ImageTk = None

try:
    from pypdf import PdfReader, PdfWriter  # type: ignore
except Exception:
    PdfReader = None
    PdfWriter = None

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore

    DND_ENABLED = True
except Exception:
    DND_FILES = None
    TkinterDnD = None
    DND_ENABLED = False


APP_NAME = "DakePDF抽出"
WINDOW_TITLE = "PDF抽出"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "brand_series": "シンプルそれDAKEシリーズ",
    "header_subtitle": "止まらない、迷わない、すぐ終わる。",
    "main_title": "PDFからページを抽出する",
    "main_description": "1つのPDFから、必要なページを何度でも連続して抽出します。",
    "button_add": "PDFを追加",
    "button_refresh": "リフレッシュ",
    "button_select_folder": "保存先を変更",
    "button_extract": "抽出して保存",
    "button_clear_selection": "選択解除",
    "button_thumbnail_minus": "－",
    "button_thumbnail_plus": "＋",
    "extract_mode_title": "抽出方法",
    "extract_mode_combined": "選択ページを1つにまとめる",
    "extract_mode_each": "1ページずつ保存する",
    "thumbnail_size": "サムネイルサイズ",
    "selected_count": "{count}ページ選択中",
    "selected_pages": "選択ページ：{pages}",
    "selected_pages_empty": "選択ページ：なし",
    "selected_pages_more": "ほか{count}ページ",
    "extracted": "抽出済み",
    "extracted_mark": "✓",
    "selected_badge": "選択中",
    "status_idle": "PDFを追加してください",
    "status_loading": "PDFを読み込み中",
    "status_rendering": "サムネイルを作成中",
    "status_ready": "ページを選択してください",
    "status_extracting": "抽出中",
    "status_saved": "{count}ページを保存しました",
    "status_error": "保存できませんでした",
    "status_busy_dots": [".", "..", "..."],
    "status_loaded": "{count}ページを読み込みました",
    "status_rendering_progress": "{current}/{total}ページのサムネイルを作成中",
    "status_extracting_progress": "{current}/{total}ページを保存中",
    "status_no_selection": "抽出するページを選択してください。",
    "status_no_pdf": "先にPDFを追加してください。",
    "status_selection_cleared": "選択を解除しました。",
    "status_folder_changed": "保存先を変更しました。",
    "status_size_changed": "サムネイルサイズを変更しました。",
    "status_drop_ready": "PDFを選ぶか、ここへドロップしてください。",
    "file_label_empty": "PDF：未選択",
    "file_label_value": "PDF：{name}",
    "page_count_empty": "総ページ数：0ページ",
    "page_count_value": "総ページ数：{count}ページ",
    "save_dir_empty": "保存先：未設定",
    "save_dir_value": "保存先：{path}",
    "thumbnail_empty_title": "PDFを追加してください",
    "thumbnail_empty_detail": "ファイル選択またはドラッグ＆ドロップで読み込みます。",
    "thumbnail_loading": "読み込み中",
    "thumbnail_error": "表示できません",
    "thumbnail_page": "P.{page}",
    "dialog_pdf_title": "PDFを選択",
    "dialog_pdf_filter": "PDFファイル",
    "dialog_folder_title": "保存先を選択",
    "error_dependency_missing": "PDF処理に必要なライブラリが見つかりません。requirements.txt をインストールしてから起動してください。",
    "error_pdf_one_file": "PDFは1ファイルだけ指定してください。",
    "error_pdf_file": "PDFファイルを指定してください。",
    "error_pdf_not_found": "PDFファイルが見つかりません。場所を確認して、もう一度選んでください。",
    "error_pdf_no_permission": "PDFを開けませんでした。ファイルの権限や使用中でないかを確認してください。",
    "error_pdf_encrypted": "暗号化されたPDFです。パスワード保護を解除したPDFでお試しください。",
    "error_pdf_broken": "PDFを読み込めませんでした。破損していないか確認してください。",
    "error_no_pages": "PDFにページが見つかりませんでした。",
    "error_save_dir_missing": "保存先フォルダが見つかりません。保存先を変更してください。",
    "error_save_dir_denied": "保存先へ書き込めません。権限のあるフォルダへ変更してください。",
    "error_save_failed": "PDFを保存できませんでした。保存先とファイル名を確認してください。",
    "error_extract_exception": "抽出中に問題が起きました。選択ページと保存先を確認してください。",
    "error_detail": "詳細：{detail}",
    "error_drop_detail": "ドロップされた内容を読み取れませんでした。PDFファイルを1つだけ指定してください。",
    "error_window_closed": "ウインドウを閉じたため処理を終了しました。",
    "page_suffix": "ページ",
    "page_range_separator": "～",
    "page_list_separator": "・",
    "file_range_separator": "-",
    "file_page_separator": "_",
    "file_page_prefix": "p",
    "file_collision_suffix": "_{number}",
    "fallback_output_name": "pdf_extract",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_tagline": "止まらない、迷わない、すぐ終わる。",
    "footer_link_1": "戸建買取査定",
    "footer_link_2": "Instagram",
    "footer_separator": " ｜ ",
    "footer_copyright": "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta",
    "launch_check_ok": "DakePDF抽出 launch-check OK",
    "process_check_ok": "DakePDF抽出 process-check OK",
    "screenshot_check_ok": "DakePDF抽出 screenshot OK",
    "self_check_ok": "DakePDF抽出 self-check OK",
    "footer_link_check_ok": "DakePDF抽出 footer-link-check OK",
    "layout_check_ok": "DakePDF抽出 layout-check OK",
    "sample_pdf_name": "千葉市中央区_契約書類.pdf",
    "sample_pdf_title": "Sample Contract Pages",
    "sample_pdf_page": "Page {page}",
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
    "success": "#12B76A",
    "success_bg": "#EAFBF3",
    "danger": "#D92D20",
    "danger_bg": "#FDECEC",
    "soft": "#EEF2F7",
    "white": "#FFFFFF",
    "link": "#58677D",
    "disabled": "#D8DEE8",
}

LINKS = {
    "assessment": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "instagram": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

APP_USER_MODEL_ID = "ShimarisuFudosan.DAKE.PDFExtract"
INTERNAL_FOLDER_NAME = "DAKE_PDF_Extract"
EXE_NAME = "DakePDF_Extract.exe"
CONFIG_FILE_NAME = "DakePDF_Extract_config.json"
COMMON_ICON_RELATIVE = Path("..") / ".." / "02_assets" / "dake_icon.ico"
COMMON_ICON_FILENAME = "dake_icon.ico"

WINDOW_SIZE = "1080x760"
WINDOW_MIN_SIZE = (820, 640)
FONT_CANDIDATES = ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo", "Segoe UI")
THUMBNAIL_MIN = 100
THUMBNAIL_MAX = 300
THUMBNAIL_DEFAULT = 156
THUMBNAIL_STEP = 20
RENDER_CACHE_WIDTH = 360
CANVAS_PAD_X = 18
CANVAS_PAD_Y = 18
CARD_GAP_X = 16
CARD_GAP_Y = 18
CARD_LABEL_HEIGHT = 54
QUEUE_POLL_MS = 50
SHIFT_MASK = 0x0001
CONTROL_MASK = 0x0004
MODE_COMBINED = "combined"
MODE_EACH = "each"


class UserFacingError(Exception):
    pass


@dataclass
class PageItem:
    page_index: int
    thumbnail_error: bool = False


@dataclass
class ExtractResult:
    pages: list[int]
    files: list[Path]


def set_app_user_model_id() -> None:
    if not sys.platform.startswith("win"):
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)
    except Exception:
        pass


def make_root() -> tk.Tk:
    if DND_ENABLED and TkinterDnD is not None:
        return TkinterDnD.Tk()
    return tk.Tk()


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def icon_candidates() -> list[Path]:
    source_dir = Path(__file__).resolve().parent
    base_dir = app_dir()
    candidates = [
        source_dir / COMMON_ICON_RELATIVE,
        base_dir / COMMON_ICON_RELATIVE,
        base_dir.parent / COMMON_ICON_RELATIVE,
        base_dir.parent.parent / COMMON_ICON_RELATIVE,
        base_dir.parent.parent.parent / "02_assets" / COMMON_ICON_FILENAME,
        base_dir / COMMON_ICON_FILENAME,
    ]
    if getattr(sys, "frozen", False):
        candidates.insert(0, Path(getattr(sys, "_MEIPASS", base_dir)) / COMMON_ICON_FILENAME)
    return candidates


def apply_window_icon(root: tk.Tk) -> bool:
    for candidate in icon_candidates():
        try:
            path = candidate.resolve()
        except OSError:
            path = candidate
        if not path.exists():
            continue
        applied = False
        try:
            root.iconbitmap(str(path))
            applied = True
        except tk.TclError:
            pass
        try:
            photo = tk.PhotoImage(file=str(path), master=root)
            root.iconphoto(True, photo)
            root._dake_icon_photo = photo  # type: ignore[attr-defined]
            applied = True
        except tk.TclError:
            if Image is not None and ImageTk is not None:
                try:
                    with Image.open(path) as image:
                        image = image.convert("RGBA")
                        photo = ImageTk.PhotoImage(image, master=root)
                    root.iconphoto(True, photo)
                    root._dake_icon_photo = photo  # type: ignore[attr-defined]
                    applied = True
                except Exception:
                    pass
        if applied:
            return True
    return False


def choose_font_family(root: tk.Tk) -> str:
    try:
        available = set(tkfont.families(root))
    except Exception:
        return "TkDefaultFont"
    for family in FONT_CANDIDATES:
        if family in available:
            return family
    return "TkDefaultFont"


def shorten_path(path: Path, max_len: int = 74) -> str:
    text = str(path)
    if len(text) <= max_len:
        return text
    drive = path.drive
    tail = Path(*path.parts[-2:]) if len(path.parts) >= 2 else path.name
    return f"{drive}\\...\\{tail}" if drive else f"...\\{tail}"


def sanitize_filename(name: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", name).strip(" .")
    return cleaned or UI_TEXT["fallback_output_name"]


def page_digits(total_pages: int) -> int:
    return max(3, len(str(max(total_pages, 1))))


def page_ranges(pages: list[int]) -> list[tuple[int, int]]:
    ordered = sorted(set(pages))
    if not ordered:
        return []
    ranges: list[tuple[int, int]] = []
    start = previous = ordered[0]
    for page in ordered[1:]:
        if page == previous + 1:
            previous = page
            continue
        ranges.append((start, previous))
        start = previous = page
    ranges.append((start, previous))
    return ranges


def page_token(pages: list[int], total_pages: int) -> str:
    digits = page_digits(total_pages)
    range_separator = UI_TEXT["file_range_separator"]
    item_separator = UI_TEXT["file_page_separator"]
    prefix = UI_TEXT["file_page_prefix"]
    tokens: list[str] = []
    for start, end in page_ranges(pages):
        if start == end:
            tokens.append(f"{start:0{digits}d}")
        else:
            tokens.append(f"{start:0{digits}d}{range_separator}{end:0{digits}d}")
    return f"{prefix}{item_separator.join(tokens)}"


def display_page_ranges(pages: list[int], max_items: int = 6) -> str:
    ranges = page_ranges(pages)
    parts: list[str] = []
    used_pages = 0
    for start, end in ranges:
        span = end - start + 1
        if used_pages >= max_items:
            break
        if start == end:
            parts.append(str(start))
        else:
            parts.append(f"{start}{UI_TEXT['page_range_separator']}{end}")
        used_pages += span
    remaining = len(set(pages)) - used_pages
    if remaining > 0:
        parts.append(UI_TEXT["selected_pages_more"].format(count=remaining))
    return (
        UI_TEXT["page_list_separator"].join(parts)
        + UI_TEXT["page_suffix"]
    )


def unique_path(target_path: Path) -> Path:
    if not target_path.exists():
        return target_path
    for number in range(2, 10000):
        suffix = UI_TEXT["file_collision_suffix"].format(number=number)
        candidate = target_path.with_name(f"{target_path.stem}{suffix}{target_path.suffix}")
        if not candidate.exists():
            return candidate
    raise UserFacingError(UI_TEXT["error_save_failed"])


def validate_output_dir(output_dir: Path) -> None:
    if not output_dir.exists() or not output_dir.is_dir():
        raise UserFacingError(UI_TEXT["error_save_dir_missing"])
    temp_path: Path | None = None
    try:
        handle, temp_name = tempfile.mkstemp(prefix=".dake_write_test_", dir=str(output_dir))
        os.close(handle)
        temp_path = Path(temp_name)
    except OSError:
        raise UserFacingError(UI_TEXT["error_save_dir_denied"])
    finally:
        if temp_path is not None:
            try:
                temp_path.unlink(missing_ok=True)
            except OSError:
                pass


def format_error_detail(exc: BaseException) -> str:
    text = str(exc).strip().replace("\n", " ")
    if not text:
        return ""
    return UI_TEXT["error_detail"].format(detail=text)


def build_pdf_error(exc: BaseException) -> str:
    if isinstance(exc, PermissionError):
        return UI_TEXT["error_pdf_no_permission"]
    detail = str(exc).lower()
    if "encrypted" in detail or "password" in detail or "decrypt" in detail:
        return UI_TEXT["error_pdf_encrypted"]
    return UI_TEXT["error_pdf_broken"]


def open_pdf_reader(source_pdf: Path) -> Any:
    if PdfReader is None:
        raise UserFacingError(UI_TEXT["error_dependency_missing"])
    try:
        reader = PdfReader(str(source_pdf))
    except PermissionError:
        raise UserFacingError(UI_TEXT["error_pdf_no_permission"])
    except Exception as exc:
        raise UserFacingError(build_pdf_error(exc))
    if getattr(reader, "is_encrypted", False):
        try:
            reader.decrypt("")
        except Exception:
            pass
        if getattr(reader, "is_encrypted", False):
            raise UserFacingError(UI_TEXT["error_pdf_encrypted"])
    if len(reader.pages) < 1:
        raise UserFacingError(UI_TEXT["error_no_pages"])
    return reader


def write_pdf_atomic(reader: Any, output_path: Path, pages: list[int]) -> None:
    if PdfWriter is None:
        raise UserFacingError(UI_TEXT["error_dependency_missing"])
    temp_path: Path | None = None
    try:
        writer = PdfWriter()
        for page_number in pages:
            writer.add_page(reader.pages[page_number - 1])
        handle, temp_name = tempfile.mkstemp(
            prefix=f".{output_path.stem}_",
            suffix=".tmp.pdf",
            dir=str(output_path.parent),
        )
        os.close(handle)
        temp_path = Path(temp_name)
        with temp_path.open("wb") as file:
            writer.write(file)
        os.replace(str(temp_path), str(output_path))
        temp_path = None
    except UserFacingError:
        raise
    except PermissionError:
        raise UserFacingError(UI_TEXT["error_save_dir_denied"])
    except Exception as exc:
        detail = format_error_detail(exc)
        message = UI_TEXT["error_save_failed"]
        raise UserFacingError(f"{message} {detail}".strip())
    finally:
        if temp_path is not None:
            try:
                temp_path.unlink(missing_ok=True)
            except OSError:
                pass


def extract_pages_to_files(
    source_pdf: Path,
    output_dir: Path,
    pages: list[int],
    mode: str,
    progress: Callable[[int, int], None] | None = None,
) -> ExtractResult:
    if not source_pdf.exists():
        raise UserFacingError(UI_TEXT["error_pdf_not_found"])
    if not pages:
        raise UserFacingError(UI_TEXT["status_no_selection"])
    validate_output_dir(output_dir)
    reader = open_pdf_reader(source_pdf)
    total_pages = len(reader.pages)
    clean_pages = sorted(set(pages))
    for page_number in clean_pages:
        if page_number < 1 or page_number > total_pages:
            raise UserFacingError(UI_TEXT["error_extract_exception"])

    source_stem = sanitize_filename(source_pdf.stem)
    output_files: list[Path] = []
    if mode == MODE_EACH and len(clean_pages) > 1:
        total = len(clean_pages)
        for current, page_number in enumerate(clean_pages, start=1):
            token = page_token([page_number], total_pages)
            output_path = unique_path(output_dir / f"{source_stem}_{token}.pdf")
            write_pdf_atomic(reader, output_path, [page_number])
            output_files.append(output_path)
            if progress is not None:
                progress(current, total)
        return ExtractResult(pages=clean_pages, files=output_files)

    token = page_token(clean_pages, total_pages)
    output_path = unique_path(output_dir / f"{source_stem}_{token}.pdf")
    write_pdf_atomic(reader, output_path, clean_pages)
    output_files.append(output_path)
    if progress is not None:
        progress(len(clean_pages), len(clean_pages))
    return ExtractResult(pages=clean_pages, files=output_files)


class ConfigStore:
    def __init__(self) -> None:
        self.path = app_dir() / CONFIG_FILE_NAME

    def load(self) -> dict[str, Any]:
        if not self.path.exists():
            return {}
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
        except Exception:
            return {}
        return data if isinstance(data, dict) else {}

    def save(self, values: dict[str, Any]) -> None:
        current = self.load()
        current.update(values)
        try:
            self.path.write_text(
                json.dumps(current, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )
        except OSError:
            pass

    def load_thumbnail_size(self) -> int:
        value = self.load().get("thumbnail_size", THUMBNAIL_DEFAULT)
        try:
            number = int(value)
        except (TypeError, ValueError):
            number = THUMBNAIL_DEFAULT
        return max(THUMBNAIL_MIN, min(THUMBNAIL_MAX, number))

    def load_extract_mode(self) -> str:
        value = str(self.load().get("extract_mode", MODE_COMBINED))
        return value if value in {MODE_COMBINED, MODE_EACH} else MODE_COMBINED


class DakePdfExtractApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(*WINDOW_MIN_SIZE)
        self.root.configure(bg=THEME["background"])
        self.root.protocol("WM_DELETE_WINDOW", self.close)
        self.window_icon_ok = apply_window_icon(self.root)

        self.font_family = choose_font_family(root)
        self.config_store = ConfigStore()
        self.queue: queue.Queue[tuple[Any, ...]] = queue.Queue()
        self.load_id = 0
        self.load_failure_source_pdf: Path | None = None
        self.extract_id = 0
        self.closed = False
        self.stop_event = threading.Event()

        self.source_pdf: Path | None = None
        self.output_dir: Path | None = None
        self.output_dir_manual = False
        self.page_count = 0
        self.pages: list[PageItem] = []
        self.selected_pages: set[int] = set()
        self.extracted_pages: set[int] = set()
        self.last_clicked_page: int | None = None
        self.thumbnail_images: dict[int, Any] = {}
        self.photo_cache: dict[tuple[int, int], Any] = {}
        self.card_bounds: dict[int, tuple[int, int, int, int]] = {}
        self.redraw_job: str | None = None
        self.save_config_job: str | None = None
        self.busy_status_key: str | None = None
        self.busy_dot_index = 0
        self.is_loading = False
        self.is_extracting = False
        self.footer_compact: bool | None = None

        self.file_var = tk.StringVar(value=UI_TEXT["file_label_empty"])
        self.page_count_var = tk.StringVar(value=UI_TEXT["page_count_empty"])
        self.save_dir_var = tk.StringVar(value=UI_TEXT["save_dir_empty"])
        self.status_var = tk.StringVar(value=UI_TEXT["status_idle"])
        self.detail_var = tk.StringVar(value=UI_TEXT["status_drop_ready"])
        self.selected_count_var = tk.StringVar(value=UI_TEXT["selected_count"].format(count=0))
        self.selected_pages_var = tk.StringVar(value=UI_TEXT["selected_pages_empty"])
        self.extract_mode_var = tk.StringVar(value=self.config_store.load_extract_mode())
        self.thumbnail_size_var = tk.IntVar(value=self.config_store.load_thumbnail_size())

        self._build_styles()
        self._build_ui()
        self._register_events()
        self._register_drop_targets()
        self._update_selection_ui()
        self._update_action_buttons()
        self._render_canvas(False)
        self.root.after(QUEUE_POLL_MS, self._poll_queue)
        self.root.after(360, self._animate_busy_status)

    def _build_styles(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(
            "Dake.Horizontal.TScale",
            troughcolor=THEME["soft"],
            background=THEME["background"],
            bordercolor=THEME["border"],
            lightcolor=THEME["border"],
            darkcolor=THEME["border"],
        )
        style.configure(
            "Dake.TRadiobutton",
            background=THEME["card"],
            foreground=THEME["text"],
            font=(self.font_family, 9),
        )

    def _build_ui(self) -> None:
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=1)

        top = tk.Frame(self.root, bg=THEME["background"])
        top.grid(row=0, column=0, sticky="ew", padx=24, pady=(18, 10))
        top.grid_columnconfigure(0, weight=1)

        header = tk.Frame(top, bg=THEME["background"])
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(0, weight=1)
        header.grid_columnconfigure(1, weight=0)

        title_area = tk.Frame(header, bg=THEME["background"])
        title_area.grid(row=0, column=0, sticky="w")
        tk.Label(
            title_area,
            text=UI_TEXT["brand_series"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 9, "bold"),
        ).pack(anchor="w")
        tk.Label(
            title_area,
            text=UI_TEXT["main_title"],
            bg=THEME["background"],
            fg=THEME["text"],
            font=(self.font_family, 22, "bold"),
        ).pack(anchor="w", pady=(2, 0))
        tk.Label(
            title_area,
            text=UI_TEXT["main_description"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 10),
        ).pack(anchor="w", pady=(4, 0))

        header_buttons = tk.Frame(header, bg=THEME["background"])
        header_buttons.grid(row=0, column=1, sticky="e", padx=(16, 0))
        self.add_button = self._make_button(header_buttons, UI_TEXT["button_add"], self.choose_pdf, "primary")
        self.add_button.pack(side="left")
        self.refresh_button = self._make_button(
            header_buttons,
            UI_TEXT["button_refresh"],
            self.refresh_pdf,
            "secondary",
        )
        self.refresh_button.pack(side="left", padx=(8, 0))
        self.folder_button = self._make_button(
            header_buttons,
            UI_TEXT["button_select_folder"],
            self.choose_output_dir,
            "secondary",
        )
        self.folder_button.pack(side="left", padx=(8, 0))

        meta = tk.Frame(top, bg=THEME["card"], highlightbackground=THEME["border"], highlightthickness=1)
        meta.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        meta.grid_columnconfigure(0, weight=1)
        meta.grid_columnconfigure(1, weight=0)
        info = tk.Frame(meta, bg=THEME["card"])
        info.grid(row=0, column=0, sticky="ew", padx=14, pady=10)
        tk.Label(info, textvariable=self.file_var, bg=THEME["card"], fg=THEME["text"], font=(self.font_family, 9, "bold")).pack(side="left")
        tk.Label(info, textvariable=self.page_count_var, bg=THEME["card"], fg=THEME["muted"], font=(self.font_family, 9)).pack(side="left", padx=(18, 0))
        tk.Label(info, textvariable=self.save_dir_var, bg=THEME["card"], fg=THEME["muted"], font=(self.font_family, 9)).pack(side="left", padx=(18, 0))

        size_area = tk.Frame(meta, bg=THEME["card"])
        size_area.grid(row=0, column=1, sticky="e", padx=14, pady=8)
        tk.Label(
            size_area,
            text=UI_TEXT["thumbnail_size"],
            bg=THEME["card"],
            fg=THEME["muted"],
            font=(self.font_family, 9),
        ).pack(side="left", padx=(0, 8))
        self.minus_button = self._make_square_button(size_area, UI_TEXT["button_thumbnail_minus"], self.decrease_thumbnail_size)
        self.minus_button.pack(side="left")
        self.size_scale = tk.Scale(
            size_area,
            from_=THUMBNAIL_MIN,
            to=THUMBNAIL_MAX,
            orient="horizontal",
            showvalue=False,
            length=150,
            resolution=1,
            variable=self.thumbnail_size_var,
            command=self.on_thumbnail_size_changed,
            bg=THEME["card"],
            troughcolor=THEME["soft"],
            activebackground=THEME["accent"],
            highlightthickness=0,
        )
        self.size_scale.pack(side="left", padx=8)
        self.plus_button = self._make_square_button(size_area, UI_TEXT["button_thumbnail_plus"], self.increase_thumbnail_size)
        self.plus_button.pack(side="left")

        canvas_frame = tk.Frame(self.root, bg=THEME["card"], highlightbackground=THEME["border"], highlightthickness=1)
        canvas_frame.grid(row=1, column=0, sticky="nsew", padx=24, pady=(10, 8))
        canvas_frame.grid_columnconfigure(0, weight=1)
        canvas_frame.grid_rowconfigure(0, weight=1)
        self.canvas = tk.Canvas(
            canvas_frame,
            bg=THEME["card"],
            bd=0,
            highlightthickness=0,
            xscrollincrement=1,
            yscrollincrement=1,
        )
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        action_bar = tk.Frame(self.root, bg=THEME["card"], highlightbackground=THEME["border"], highlightthickness=1)
        action_bar.grid(row=2, column=0, sticky="ew", padx=24, pady=(0, 8))
        action_bar.grid_columnconfigure(0, weight=1)
        action_bar.grid_columnconfigure(1, weight=0)
        action_bar.grid_columnconfigure(2, weight=1)

        selection_info = tk.Frame(action_bar, bg=THEME["card"])
        selection_info.grid(row=0, column=0, sticky="w", padx=14, pady=10)
        tk.Label(
            selection_info,
            textvariable=self.selected_count_var,
            bg=THEME["card"],
            fg=THEME["text"],
            font=(self.font_family, 11, "bold"),
        ).pack(anchor="w")
        tk.Label(
            selection_info,
            textvariable=self.selected_pages_var,
            bg=THEME["card"],
            fg=THEME["muted"],
            font=(self.font_family, 9),
        ).pack(anchor="w", pady=(2, 0))

        mode_area = tk.Frame(action_bar, bg=THEME["card"])
        mode_area.grid(row=0, column=1, sticky="", padx=12, pady=10)
        tk.Label(
            mode_area,
            text=UI_TEXT["extract_mode_title"],
            bg=THEME["card"],
            fg=THEME["muted"],
            font=(self.font_family, 9, "bold"),
        ).pack(anchor="w")
        radios = tk.Frame(mode_area, bg=THEME["card"])
        radios.pack(anchor="w", pady=(4, 0))
        for value, label in (
            (MODE_COMBINED, UI_TEXT["extract_mode_combined"]),
            (MODE_EACH, UI_TEXT["extract_mode_each"]),
        ):
            tk.Radiobutton(
                radios,
                text=label,
                variable=self.extract_mode_var,
                value=value,
                command=self.on_extract_mode_changed,
                bg=THEME["card"],
                fg=THEME["text"],
                activebackground=THEME["card"],
                activeforeground=THEME["accent"],
                selectcolor=THEME["card"],
                font=(self.font_family, 9),
            ).pack(side="left", padx=(0, 12))

        action_buttons = tk.Frame(action_bar, bg=THEME["card"])
        action_buttons.grid(row=0, column=2, sticky="e", padx=14, pady=10)
        self.extract_button = self._make_button(
            action_buttons,
            UI_TEXT["button_extract"],
            self.start_extract,
            "primary",
        )
        self.extract_button.pack(side="left")
        self.clear_button = self._make_button(
            action_buttons,
            UI_TEXT["button_clear_selection"],
            self.clear_selection,
            "secondary",
        )
        self.clear_button.pack(side="left", padx=(8, 0))

        status = tk.Frame(self.root, bg=THEME["background"])
        status.grid(row=3, column=0, sticky="ew", padx=24, pady=(0, 8))
        self.status_badge = tk.Label(
            status,
            textvariable=self.status_var,
            bg=THEME["soft"],
            fg=THEME["muted"],
            font=(self.font_family, 9, "bold"),
            padx=10,
            pady=4,
        )
        self.status_badge.pack(side="left")
        tk.Label(
            status,
            textvariable=self.detail_var,
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 9),
        ).pack(side="left", padx=(10, 0))

        self.footer = tk.Frame(self.root, bg=THEME["background"])
        self.footer.grid(row=4, column=0, sticky="ew", padx=24, pady=(0, 14))
        self.footer_left = tk.Frame(self.footer, bg=THEME["background"])
        self._make_footer_text(self.footer_left, UI_TEXT["footer_left"], True)
        self._make_footer_text(self.footer_left, UI_TEXT["footer_separator"])
        self._make_footer_text(self.footer_left, UI_TEXT["footer_tagline"])
        self.footer_right = tk.Frame(self.footer, bg=THEME["background"])
        self._make_footer_link(self.footer_right, UI_TEXT["footer_link_1"], LINKS["assessment"])
        self._make_footer_text(self.footer_right, UI_TEXT["footer_separator"])
        self._make_footer_link(self.footer_right, UI_TEXT["footer_link_2"], LINKS["instagram"])
        self._make_footer_text(self.footer_right, UI_TEXT["footer_separator"])
        self._make_footer_text(self.footer_right, UI_TEXT["footer_copyright"])
        self._update_footer_layout()

    def _make_button(self, parent: tk.Misc, label: str, command: Callable[[], None], variant: str) -> tk.Button:
        primary = variant == "primary"
        button = tk.Button(
            parent,
            text=label,
            command=command,
            bg=THEME["accent"] if primary else THEME["white"],
            fg=THEME["white"] if primary else THEME["text"],
            activebackground=THEME["accent_hover"] if primary else THEME["selection_bg"],
            activeforeground=THEME["white"] if primary else THEME["text"],
            disabledforeground=THEME["muted"],
            font=(self.font_family, 10, "bold"),
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=THEME["accent"] if primary else THEME["border"],
            highlightcolor=THEME["accent"] if primary else THEME["border"],
            padx=18,
            pady=9,
            cursor="hand2",
        )
        return button

    def _make_square_button(self, parent: tk.Misc, label: str, command: Callable[[], None]) -> tk.Button:
        return tk.Button(
            parent,
            text=label,
            command=command,
            width=2,
            height=1,
            bg=THEME["white"],
            fg=THEME["text"],
            activebackground=THEME["selection_bg"],
            activeforeground=THEME["text"],
            font=(self.font_family, 10, "bold"),
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=THEME["border"],
            highlightcolor=THEME["accent"],
            cursor="hand2",
        )

    def _make_footer_text(self, parent: tk.Frame, label: str, bold: bool = False) -> None:
        tk.Label(
            parent,
            text=label,
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 8, "bold" if bold else "normal"),
        ).pack(side="left")

    def _make_footer_link(self, parent: tk.Frame, label: str, url: str) -> None:
        link = tk.Label(
            parent,
            text=label,
            bg=THEME["background"],
            fg=THEME["link"],
            font=(self.font_family, 8, "bold"),
            cursor="hand2",
        )
        link.pack(side="left")
        link.bind("<Button-1>", lambda _event: self._open_url(url))
        link.bind("<Enter>", lambda _event: link.configure(fg=THEME["accent"]))
        link.bind("<Leave>", lambda _event: link.configure(fg=THEME["link"]))

    def _register_events(self) -> None:
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind("<Button-1>", self._on_canvas_click)
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.root.bind("<Configure>", self._on_root_configure, add="+")
        self.root.bind_all("<Return>", self._on_enter_key)
        self.root.bind_all("<Escape>", self._on_escape_key)

    def _register_drop_targets(self) -> None:
        if not DND_ENABLED or DND_FILES is None:
            return
        for target in (self.root, self.canvas):
            if not hasattr(target, "drop_target_register"):
                continue
            try:
                target.drop_target_register(DND_FILES)  # type: ignore[attr-defined]
                target.dnd_bind("<<Drop>>", self._on_drop)  # type: ignore[attr-defined]
            except Exception:
                pass

    def _on_root_configure(self, event: tk.Event) -> None:
        if event.widget == self.root:
            self._update_footer_layout()

    def _update_footer_layout(self) -> None:
        compact = self.root.winfo_width() < 900
        if compact == self.footer_compact:
            return
        self.footer_compact = compact
        self.footer_left.pack_forget()
        self.footer_right.pack_forget()
        if compact:
            self.footer_left.pack(anchor="center", pady=(0, 3))
            self.footer_right.pack(anchor="center")
            return
        self.footer_left.pack(side="left")
        self.footer_right.pack(side="right")

    def _open_url(self, url: str) -> None:
        try:
            webbrowser.open(url, new=2)
        except Exception:
            pass

    def _set_status(self, status_key: str, detail: str | None = None, tone: str = "neutral") -> None:
        self.busy_status_key = status_key if status_key in {"status_loading", "status_rendering", "status_extracting"} else None
        self.busy_dot_index = 0
        self.status_var.set(UI_TEXT[status_key])
        if detail is not None:
            self.detail_var.set(detail)
        if tone == "success":
            self.status_badge.configure(bg=THEME["success_bg"], fg=THEME["success"])
        elif tone == "error":
            self.status_badge.configure(bg=THEME["danger_bg"], fg=THEME["danger"])
        elif tone == "active":
            self.status_badge.configure(bg=THEME["selection_bg"], fg=THEME["accent"])
        else:
            self.status_badge.configure(bg=THEME["soft"], fg=THEME["muted"])

    def _animate_busy_status(self) -> None:
        if self.closed:
            return
        if self.busy_status_key is not None:
            dots = UI_TEXT["status_busy_dots"]
            self.status_var.set(f"{UI_TEXT[self.busy_status_key]}{dots[self.busy_dot_index % len(dots)]}")
            self.busy_dot_index += 1
        self.root.after(360, self._animate_busy_status)

    def _on_canvas_configure(self, _event: tk.Event) -> None:
        self._schedule_redraw(True)

    def _on_canvas_click(self, event: tk.Event) -> None:
        page_number = self._page_from_xy(event.x, event.y)
        if page_number is None:
            return
        self.toggle_page(page_number, bool(event.state & SHIFT_MASK))

    def _on_mousewheel(self, event: tk.Event) -> str | None:
        if event.state & CONTROL_MASK:
            if event.delta > 0:
                self.increase_thumbnail_size()
            else:
                self.decrease_thumbnail_size()
            return "break"
        units = -1 * int(event.delta / 120) if event.delta else 0
        self.canvas.yview_scroll(units, "units")
        return "break"

    def _on_enter_key(self, _event: tk.Event) -> str:
        self.start_extract()
        return "break"

    def _on_escape_key(self, _event: tk.Event) -> str:
        self.clear_selection()
        return "break"

    def choose_pdf(self) -> None:
        filename = filedialog.askopenfilename(
            title=UI_TEXT["dialog_pdf_title"],
            filetypes=[(UI_TEXT["dialog_pdf_filter"], "*.pdf")],
        )
        if filename:
            self.load_pdf(Path(filename))

    def choose_output_dir(self) -> None:
        selected = filedialog.askdirectory(title=UI_TEXT["dialog_folder_title"])
        if not selected:
            return
        self.output_dir = Path(selected)
        self.output_dir_manual = True
        self._refresh_file_info()
        self.config_store.save({"last_output_dir": str(self.output_dir)})
        self._set_status("status_ready", UI_TEXT["status_folder_changed"], "success")

    def refresh_pdf(self) -> None:
        if self.source_pdf is None or self.is_loading or self.is_extracting:
            return
        self.load_pdf(self.source_pdf, keep_output_dir=True, reset_extracted=True)

    def _on_drop(self, event: tk.Event) -> None:
        try:
            paths = self._drop_paths(str(event.data))  # type: ignore[attr-defined]
        except Exception:
            self._set_status("status_idle", UI_TEXT["error_drop_detail"], "error")
            return
        if len(paths) != 1:
            self._set_status("status_idle", UI_TEXT["error_pdf_one_file"], "error")
            return
        self.load_pdf(paths[0])

    def _drop_paths(self, data: str) -> list[Path]:
        try:
            parts = self.root.tk.splitlist(data)
        except tk.TclError:
            parts = [data]
        return [Path(part) for part in parts if str(part).strip()]

    def load_pdf(
        self,
        pdf_path: Path,
        *,
        keep_output_dir: bool = False,
        reset_extracted: bool | None = None,
    ) -> None:
        if self.is_loading:
            return
        if fitz is None or Image is None or ImageTk is None:
            self._set_status("status_error", UI_TEXT["error_dependency_missing"], "error")
            return
        if pdf_path.suffix.lower() != ".pdf":
            self._set_status("status_idle", UI_TEXT["error_pdf_file"], "error")
            return
        if not pdf_path.exists():
            self._set_status("status_idle", UI_TEXT["error_pdf_not_found"], "error")
            return

        same_pdf = self.source_pdf is not None and self._same_path(self.source_pdf, pdf_path)
        previous_output_dir = self.output_dir
        previous_output_dir_manual = self.output_dir_manual
        if reset_extracted is None:
            reset_extracted = not same_pdf
        self.load_id += 1
        current_load_id = self.load_id
        self.load_failure_source_pdf = pdf_path if keep_output_dir else None
        self.source_pdf = pdf_path
        if keep_output_dir and previous_output_dir is not None:
            self.output_dir = previous_output_dir
            self.output_dir_manual = previous_output_dir_manual
        elif not same_pdf:
            self.output_dir = pdf_path.parent
            self.output_dir_manual = False
        if reset_extracted:
            self.extracted_pages.clear()
        self.page_count = 0
        self.pages = []
        self.selected_pages.clear()
        self.thumbnail_images.clear()
        self.photo_cache.clear()
        self.card_bounds.clear()
        self.last_clicked_page = None
        self.is_loading = True
        self._refresh_file_info()
        self._update_selection_ui()
        self._update_action_buttons()
        self._set_status("status_loading", UI_TEXT["status_loading"], "active")
        self.canvas.yview_moveto(0)
        self._render_canvas(False)

        thread = threading.Thread(
            target=self._load_pdf_worker,
            args=(current_load_id, pdf_path),
            daemon=True,
        )
        thread.start()

    def _same_path(self, left: Path, right: Path) -> bool:
        try:
            return left.resolve() == right.resolve()
        except OSError:
            return left == right

    def _load_pdf_worker(self, current_load_id: int, pdf_path: Path) -> None:
        try:
            if fitz is None:
                raise UserFacingError(UI_TEXT["error_dependency_missing"])
            with fitz.open(str(pdf_path)) as document:  # type: ignore[union-attr]
                if getattr(document, "needs_pass", False):
                    raise UserFacingError(UI_TEXT["error_pdf_encrypted"])
                page_count = int(document.page_count)
                if page_count < 1:
                    raise UserFacingError(UI_TEXT["error_no_pages"])
                self.queue.put(("load_ready", current_load_id, pdf_path, page_count))
                for page_index in range(page_count):
                    if self.closed or self.stop_event.is_set() or current_load_id != self.load_id:
                        return
                    try:
                        page = document.load_page(page_index)
                        rect = page.rect
                        zoom = min(RENDER_CACHE_WIDTH / max(rect.width, 1), RENDER_CACHE_WIDTH * 1.45 / max(rect.height, 1))
                        zoom = max(0.18, min(zoom, 2.0))
                        pixmap = page.get_pixmap(  # type: ignore[union-attr]
                            matrix=fitz.Matrix(zoom, zoom),
                            alpha=False,
                            colorspace=fitz.csRGB,
                        )
                        data = pixmap.tobytes("png")
                        self.queue.put(("thumbnail_ready", current_load_id, page_index, data))
                    except Exception:
                        self.queue.put(("thumbnail_failed", current_load_id, page_index))
                    if page_index == 0 or (page_index + 1) % 5 == 0:
                        self.queue.put(("thumbnail_progress", current_load_id, page_index + 1, page_count))
                self.queue.put(("thumbnail_done", current_load_id))
        except UserFacingError as exc:
            self.queue.put(("load_failed", current_load_id, str(exc)))
        except PermissionError:
            self.queue.put(("load_failed", current_load_id, UI_TEXT["error_pdf_no_permission"]))
        except Exception as exc:
            self.queue.put(("load_failed", current_load_id, build_pdf_error(exc)))

    def _poll_queue(self) -> None:
        if self.closed:
            return
        handled = 0
        while handled < 28:
            try:
                message = self.queue.get_nowait()
            except queue.Empty:
                break
            self._handle_queue_message(message)
            handled += 1
        self.root.after(QUEUE_POLL_MS, self._poll_queue)

    def _handle_queue_message(self, message: tuple[Any, ...]) -> None:
        kind = message[0]
        if kind in {"load_ready", "thumbnail_ready", "thumbnail_failed", "thumbnail_progress", "thumbnail_done", "load_failed"}:
            current_load_id = message[1]
            if current_load_id != self.load_id:
                return

        if kind == "load_ready":
            _kind, _load_id, pdf_path, page_count = message
            self.page_count = int(page_count)
            self.pages = [PageItem(page_index=index) for index in range(self.page_count)]
            self.source_pdf = Path(pdf_path)
            self._refresh_file_info()
            self._set_status("status_rendering", UI_TEXT["status_loaded"].format(count=self.page_count), "active")
            self._update_action_buttons()
            self._render_canvas(False)
            return

        if kind == "thumbnail_ready":
            _kind, _load_id, page_index, data = message
            self._apply_thumbnail(int(page_index), data)
            return

        if kind == "thumbnail_failed":
            _kind, _load_id, page_index = message
            if 0 <= int(page_index) < len(self.pages):
                self.pages[int(page_index)].thumbnail_error = True
            self._schedule_redraw(True)
            return

        if kind == "thumbnail_progress":
            _kind, _load_id, current, total = message
            self._set_status(
                "status_rendering",
                UI_TEXT["status_rendering_progress"].format(current=current, total=total),
                "active",
            )
            return

        if kind == "thumbnail_done":
            self.is_loading = False
            self.load_failure_source_pdf = None
            self._set_status("status_ready", UI_TEXT["status_ready"], "neutral")
            self._update_action_buttons()
            return

        if kind == "load_failed":
            _kind, _load_id, message_text = message
            self.is_loading = False
            self.source_pdf = self.load_failure_source_pdf
            self.load_failure_source_pdf = None
            self.page_count = 0
            self.pages = []
            self.selected_pages.clear()
            self.thumbnail_images.clear()
            self.photo_cache.clear()
            self._refresh_file_info()
            self._update_selection_ui()
            self._update_action_buttons()
            self._set_status("status_idle", str(message_text), "error")
            self._render_canvas(False)
            return

        if kind == "extract_progress":
            _kind, current_extract_id, current, total = message
            if current_extract_id != self.extract_id:
                return
            self._set_status(
                "status_extracting",
                UI_TEXT["status_extracting_progress"].format(current=current, total=total),
                "active",
            )
            return

        if kind == "extract_done":
            _kind, current_extract_id, result = message
            if current_extract_id != self.extract_id:
                return
            self.is_extracting = False
            self.extracted_pages.update(result.pages)
            saved_count = len(result.pages)
            self.selected_pages.clear()
            self.last_clicked_page = None
            self._update_selection_ui()
            self._update_action_buttons()
            self._set_status(
                "status_saved",
                UI_TEXT["status_saved"].format(count=saved_count),
                "success",
            )
            self.status_var.set(UI_TEXT["status_saved"].format(count=saved_count))
            self._schedule_redraw(True)
            return

        if kind == "extract_failed":
            _kind, current_extract_id, message_text = message
            if current_extract_id != self.extract_id:
                return
            self.is_extracting = False
            self._update_action_buttons()
            self._set_status("status_error", str(message_text), "error")

    def _apply_thumbnail(self, page_index: int, data: bytes) -> None:
        if Image is None or ImageTk is None:
            return
        try:
            image = Image.open(io.BytesIO(data)).convert("RGB")
            self.thumbnail_images[page_index] = image.copy()
            if 0 <= page_index < len(self.pages):
                self.pages[page_index].thumbnail_error = False
            keys_to_remove = [key for key in self.photo_cache if key[0] == page_index]
            for key in keys_to_remove:
                self.photo_cache.pop(key, None)
        except Exception:
            if 0 <= page_index < len(self.pages):
                self.pages[page_index].thumbnail_error = True
        self._schedule_redraw(True)

    def _refresh_file_info(self) -> None:
        if self.source_pdf is None:
            self.file_var.set(UI_TEXT["file_label_empty"])
        else:
            self.file_var.set(UI_TEXT["file_label_value"].format(name=self.source_pdf.name))
        if self.page_count:
            self.page_count_var.set(UI_TEXT["page_count_value"].format(count=self.page_count))
        else:
            self.page_count_var.set(UI_TEXT["page_count_empty"])
        if self.output_dir is None:
            self.save_dir_var.set(UI_TEXT["save_dir_empty"])
        else:
            self.save_dir_var.set(UI_TEXT["save_dir_value"].format(path=shorten_path(self.output_dir)))

    def _update_selection_ui(self) -> None:
        selected = sorted(self.selected_pages)
        count = len(selected)
        self.selected_count_var.set(UI_TEXT["selected_count"].format(count=count))
        if selected:
            self.selected_pages_var.set(UI_TEXT["selected_pages"].format(pages=display_page_ranges(selected)))
        else:
            self.selected_pages_var.set(UI_TEXT["selected_pages_empty"])

    def _update_action_buttons(self) -> None:
        busy = self.is_loading or self.is_extracting
        can_extract = self.source_pdf is not None and self.page_count > 0 and bool(self.selected_pages) and not busy
        self.extract_button.configure(state="normal" if can_extract else "disabled")
        self.clear_button.configure(state="normal" if self.selected_pages and not busy else "disabled")
        self.refresh_button.configure(
            state="normal" if self.source_pdf is not None and not busy else "disabled"
        )
        self.folder_button.configure(state="normal" if not busy else "disabled")
        self.add_button.configure(state="normal" if not busy else "disabled")

    def toggle_page(self, page_number: int, shift_pressed: bool = False) -> None:
        if page_number < 1 or page_number > self.page_count:
            return
        if shift_pressed and self.last_clicked_page is not None:
            start = min(self.last_clicked_page, page_number)
            end = max(self.last_clicked_page, page_number)
            self.selected_pages.update(range(start, end + 1))
        else:
            if page_number in self.selected_pages:
                self.selected_pages.remove(page_number)
            else:
                self.selected_pages.add(page_number)
        self.last_clicked_page = page_number
        self._update_selection_ui()
        self._update_action_buttons()
        self._set_status("status_ready", UI_TEXT["status_ready"], "neutral")
        self._schedule_redraw(True)

    def select_pages(self, pages: set[int]) -> None:
        self.selected_pages = {page for page in pages if 1 <= page <= self.page_count}
        self.last_clicked_page = max(self.selected_pages) if self.selected_pages else None
        self._update_selection_ui()
        self._update_action_buttons()
        self._schedule_redraw(True)

    def clear_selection(self) -> None:
        if not self.selected_pages:
            return
        self.selected_pages.clear()
        self.last_clicked_page = None
        self._update_selection_ui()
        self._update_action_buttons()
        self._set_status("status_ready", UI_TEXT["status_selection_cleared"], "neutral")
        self._schedule_redraw(True)

    def start_extract(self) -> None:
        if self.is_extracting:
            return
        if self.source_pdf is None or self.page_count < 1:
            self._set_status("status_idle", UI_TEXT["status_no_pdf"], "error")
            return
        if not self.selected_pages:
            self._set_status("status_ready", UI_TEXT["status_no_selection"], "error")
            return
        if self.output_dir is None:
            self._set_status("status_error", UI_TEXT["error_save_dir_missing"], "error")
            return
        pages = sorted(self.selected_pages)
        mode = self.extract_mode_var.get()
        self.extract_id += 1
        current_extract_id = self.extract_id
        self.is_extracting = True
        self.config_store.save({"extract_mode": mode, "last_output_dir": str(self.output_dir)})
        self._update_action_buttons()
        self._set_status("status_extracting", UI_TEXT["status_extracting"], "active")
        thread = threading.Thread(
            target=self._extract_worker,
            args=(current_extract_id, self.source_pdf, self.output_dir, pages, mode),
            daemon=True,
        )
        thread.start()

    def _extract_worker(self, current_extract_id: int, source_pdf: Path, output_dir: Path, pages: list[int], mode: str) -> None:
        def progress(current: int, total: int) -> None:
            self.queue.put(("extract_progress", current_extract_id, current, total))

        try:
            result = extract_pages_to_files(source_pdf, output_dir, pages, mode, progress)
            self.queue.put(("extract_done", current_extract_id, result))
        except UserFacingError as exc:
            self.queue.put(("extract_failed", current_extract_id, str(exc)))
        except Exception as exc:
            detail = format_error_detail(exc)
            self.queue.put(("extract_failed", current_extract_id, f"{UI_TEXT['error_extract_exception']} {detail}".strip()))

    def on_extract_mode_changed(self) -> None:
        self.config_store.save({"extract_mode": self.extract_mode_var.get()})
        self._update_action_buttons()

    def on_thumbnail_size_changed(self, _value: str | None = None) -> None:
        self.photo_cache.clear()
        self._schedule_redraw(True)
        if self.save_config_job is not None:
            try:
                self.root.after_cancel(self.save_config_job)
            except tk.TclError:
                pass
        self.save_config_job = self.root.after(240, self._save_thumbnail_size)

    def _save_thumbnail_size(self) -> None:
        self.save_config_job = None
        self.config_store.save({"thumbnail_size": int(self.thumbnail_size_var.get())})

    def increase_thumbnail_size(self) -> None:
        current = int(self.thumbnail_size_var.get())
        self.thumbnail_size_var.set(min(THUMBNAIL_MAX, current + THUMBNAIL_STEP))
        self.on_thumbnail_size_changed()

    def decrease_thumbnail_size(self) -> None:
        current = int(self.thumbnail_size_var.get())
        self.thumbnail_size_var.set(max(THUMBNAIL_MIN, current - THUMBNAIL_STEP))
        self.on_thumbnail_size_changed()

    def _schedule_redraw(self, preserve_scroll: bool = True) -> None:
        if self.redraw_job is not None:
            return
        self.redraw_job = self.root.after(16, lambda: self._render_canvas(preserve_scroll))

    def _render_canvas(self, preserve_scroll: bool = True) -> None:
        self.redraw_job = None
        scroll_top = self.canvas.yview()[0] if preserve_scroll else 0.0
        self.canvas.delete("all")
        self.card_bounds.clear()
        width = max(240, self.canvas.winfo_width())
        thumb_width = int(self.thumbnail_size_var.get())
        thumb_height = int(thumb_width * 1.42)
        card_width = thumb_width + 28
        card_height = thumb_height + CARD_LABEL_HEIGHT
        columns = max(1, (width - CANVAS_PAD_X * 2 + CARD_GAP_X) // (card_width + CARD_GAP_X))

        if self.page_count < 1:
            self._draw_empty_state(width)
            self.canvas.configure(scrollregion=(0, 0, width, max(360, self.canvas.winfo_height())))
            return

        rows = (self.page_count + columns - 1) // columns
        content_width = CANVAS_PAD_X * 2 + columns * card_width + (columns - 1) * CARD_GAP_X
        content_height = CANVAS_PAD_Y * 2 + rows * card_height + (rows - 1) * CARD_GAP_Y

        for page_index in range(self.page_count):
            row = page_index // columns
            column = page_index % columns
            x = CANVAS_PAD_X + column * (card_width + CARD_GAP_X)
            y = CANVAS_PAD_Y + row * (card_height + CARD_GAP_Y)
            self._draw_page_card(page_index, x, y, card_width, card_height, thumb_width, thumb_height)

        self.canvas.configure(scrollregion=(0, 0, max(width, content_width), max(content_height, self.canvas.winfo_height())))
        if preserve_scroll:
            self.canvas.yview_moveto(scroll_top)

    def _draw_empty_state(self, width: int) -> None:
        center_x = width // 2
        center_y = max(150, self.canvas.winfo_height() // 2 - 20)
        self.canvas.create_text(
            center_x,
            center_y,
            text=UI_TEXT["thumbnail_empty_title"],
            fill=THEME["text"],
            font=(self.font_family, 18, "bold"),
        )
        self.canvas.create_text(
            center_x,
            center_y + 34,
            text=UI_TEXT["thumbnail_empty_detail"],
            fill=THEME["muted"],
            font=(self.font_family, 10),
        )

    def _draw_page_card(
        self,
        page_index: int,
        x: int,
        y: int,
        card_width: int,
        card_height: int,
        thumb_width: int,
        thumb_height: int,
    ) -> None:
        page_number = page_index + 1
        selected = page_number in self.selected_pages
        extracted = page_number in self.extracted_pages
        fill = THEME["selection_bg"] if selected else THEME["card"]
        outline = THEME["selection_border"] if selected else THEME["border"]
        border_width = 2 if selected else 1
        self.canvas.create_rectangle(
            x,
            y,
            x + card_width,
            y + card_height,
            fill=fill,
            outline=outline,
            width=border_width,
        )
        image_x = x + (card_width - thumb_width) // 2
        image_y = y + 14
        self.canvas.create_rectangle(
            image_x,
            image_y,
            image_x + thumb_width,
            image_y + thumb_height,
            fill=THEME["background"],
            outline=THEME["border"],
        )
        photo = self._scaled_photo(page_index, thumb_width, thumb_height)
        if photo is not None:
            self.canvas.create_image(
                image_x + thumb_width // 2,
                image_y + thumb_height // 2,
                image=photo,
                anchor="center",
            )
        else:
            text = UI_TEXT["thumbnail_error"] if self._page_has_error(page_index) else UI_TEXT["thumbnail_loading"]
            self.canvas.create_text(
                image_x + thumb_width // 2,
                image_y + thumb_height // 2,
                text=text,
                fill=THEME["muted"],
                font=(self.font_family, 9),
            )
        if extracted and not selected:
            self.canvas.create_rectangle(
                image_x,
                image_y,
                image_x + thumb_width,
                image_y + thumb_height,
                fill=THEME["white"],
                outline="",
                stipple="gray12",
            )
        label_y = image_y + thumb_height + 16
        self.canvas.create_text(
            x + 14,
            label_y,
            text=UI_TEXT["thumbnail_page"].format(page=page_number),
            fill=THEME["text"],
            font=(self.font_family, 10, "bold"),
            anchor="w",
        )
        if selected:
            self.canvas.create_text(
                x + card_width - 14,
                label_y,
                text=UI_TEXT["selected_badge"],
                fill=THEME["accent"],
                font=(self.font_family, 8, "bold"),
                anchor="e",
            )
        elif extracted:
            badge_x = x + card_width - 30
            badge_y = y + 14
            self.canvas.create_oval(
                badge_x,
                badge_y,
                badge_x + 20,
                badge_y + 20,
                fill=THEME["success"],
                outline=THEME["success"],
            )
            self.canvas.create_text(
                badge_x + 10,
                badge_y + 9,
                text=UI_TEXT["extracted_mark"],
                fill=THEME["white"],
                font=(self.font_family, 10, "bold"),
            )
            if thumb_width >= 145:
                self.canvas.create_text(
                    x + card_width - 14,
                    label_y,
                    text=UI_TEXT["extracted"],
                    fill=THEME["success"],
                    font=(self.font_family, 8, "bold"),
                    anchor="e",
                )
        self.card_bounds[page_number] = (x, y, x + card_width, y + card_height)

    def _page_has_error(self, page_index: int) -> bool:
        return 0 <= page_index < len(self.pages) and self.pages[page_index].thumbnail_error

    def _scaled_photo(self, page_index: int, thumb_width: int, thumb_height: int) -> Any | None:
        if Image is None or ImageTk is None:
            return None
        image = self.thumbnail_images.get(page_index)
        if image is None:
            return None
        key = (page_index, thumb_width)
        cached = self.photo_cache.get(key)
        if cached is not None:
            return cached
        try:
            resized = image.copy()
            resample = Image.Resampling.LANCZOS if hasattr(Image, "Resampling") else Image.LANCZOS
            resized.thumbnail((thumb_width, thumb_height), resample)
            photo = ImageTk.PhotoImage(resized, master=self.root)
        except Exception:
            return None
        self.photo_cache[key] = photo
        return photo

    def _page_from_xy(self, x: int, y: int) -> int | None:
        canvas_x = int(self.canvas.canvasx(x))
        canvas_y = int(self.canvas.canvasy(y))
        for page_number, (left, top, right, bottom) in self.card_bounds.items():
            if left <= canvas_x <= right and top <= canvas_y <= bottom:
                return page_number
        return None

    def close(self) -> None:
        self.closed = True
        self.stop_event.set()
        self.load_id += 1
        self.extract_id += 1
        try:
            self.config_store.save(
                {
                    "thumbnail_size": int(self.thumbnail_size_var.get()),
                    "extract_mode": self.extract_mode_var.get(),
                }
            )
        except Exception:
            pass
        self.root.destroy()


def create_sample_pdf(path: Path, page_count: int = 8) -> None:
    if fitz is None:
        raise UserFacingError(UI_TEXT["error_dependency_missing"])
    document = fitz.open()  # type: ignore[union-attr]
    for page_number in range(1, page_count + 1):
        page = document.new_page(width=595, height=842)
        page.insert_text((70, 86), UI_TEXT["sample_pdf_title"], fontsize=24)
        page.insert_text((70, 132), UI_TEXT["sample_pdf_page"].format(page=page_number), fontsize=18)
        page.draw_rect((70, 180, 525, 740), color=(0.18, 0.43, 0.93), width=1)
        for line in range(9):
            y = 220 + line * 46
            page.draw_line((100, y), (495, y), color=(0.45, 0.55, 0.65), width=0.7)
    document.save(str(path))
    document.close()


def run_launch_check() -> int:
    missing = []
    if fitz is None:
        missing.append("fitz")
    if Image is None or ImageTk is None:
        missing.append("Pillow")
    if PdfReader is None or PdfWriter is None:
        missing.append("pypdf")
    if missing:
        print(",".join(missing), file=sys.stderr)
        return 1
    print(UI_TEXT["launch_check_ok"])
    print(f"icon={any(path.exists() for path in icon_candidates())}")
    return 0


def run_process_check() -> int:
    with tempfile.TemporaryDirectory(prefix="dake_pdf_extract_check_") as temp_dir:
        temp = Path(temp_dir)
        source = temp / UI_TEXT["sample_pdf_name"]
        output = temp / "out"
        output.mkdir()
        create_sample_pdf(source, 8)
        combined = extract_pages_to_files(source, output, [2, 3, 4], MODE_COMBINED)
        if len(combined.files) != 1 or not combined.files[0].name.endswith("_p002-004.pdf"):
            return 1
        sparse = extract_pages_to_files(source, output, [2, 4, 7], MODE_COMBINED)
        if len(sparse.files) != 1 or not sparse.files[0].name.endswith("_p002_004_007.pdf"):
            return 1
        each = extract_pages_to_files(source, output, [2, 4, 7], MODE_EACH)
        if len(each.files) != 3:
            return 1
        collision = extract_pages_to_files(source, output, [2, 3, 4], MODE_COMBINED)
        if not collision.files[0].name.endswith("_p002-004_2.pdf"):
            return 1
    print(UI_TEXT["process_check_ok"])
    return 0


def run_self_check() -> int:
    with tempfile.TemporaryDirectory(prefix="dake_pdf_extract_self_") as temp_dir:
        temp = Path(temp_dir)
        source = temp / UI_TEXT["sample_pdf_name"]
        create_sample_pdf(source, 8)
        root = make_root()
        root.withdraw()
        app = DakePdfExtractApp(root)
        app.refresh_pdf()
        if app.load_id != 0:
            app.close()
            return 1
        app.load_pdf(source)
        deadline = time.time() + 10
        while time.time() < deadline:
            root.update()
            if not app.is_loading and app.page_count == 8 and len(app.thumbnail_images) >= 8:
                break
            time.sleep(0.05)
        if app.is_loading or app.page_count != 8:
            app.close()
            return 1
        updated_page_count = 6
        create_sample_pdf(source, updated_page_count)
        initial_output_dir = app.output_dir
        initial_thumbnail_size = int(app.thumbnail_size_var.get())
        app.extracted_pages.update({1, 5})
        app.select_pages({2, 3})
        app.refresh_pdf()
        refresh_load_id = app.load_id
        app.refresh_pdf()
        if app.load_id != refresh_load_id:
            app.close()
            return 1
        deadline = time.time() + 10
        while time.time() < deadline:
            root.update()
            if not app.is_loading and app.page_count == updated_page_count and len(app.thumbnail_images) >= updated_page_count:
                break
            time.sleep(0.05)
        if (
            app.is_loading
            or app.page_count != updated_page_count
            or app.selected_pages
            or app.extracted_pages
            or app.output_dir != initial_output_dir
            or int(app.thumbnail_size_var.get()) != initial_thumbnail_size
        ):
            app.close()
            return 1
        app.toggle_page(2)
        app.toggle_page(4, True)
        if app.selected_pages != {2, 3, 4}:
            app.close()
            return 1
        app.clear_selection()
        app.select_pages({2, 4, 6})
        app.output_dir = temp
        app.start_extract()
        deadline = time.time() + 10
        while time.time() < deadline:
            root.update()
            if not app.is_extracting:
                break
            time.sleep(0.05)
        if app.selected_pages or not {2, 4, 6}.issubset(app.extracted_pages):
            app.close()
            return 1
        for page in range(1, 6):
            app.select_pages({page})
            app.start_extract()
            deadline = time.time() + 10
            while time.time() < deadline:
                root.update()
                if not app.is_extracting:
                    break
                time.sleep(0.05)
            if app.selected_pages:
                app.close()
                return 1
        app.close()
    print(UI_TEXT["self_check_ok"])
    return 0


def iter_widgets(widget: tk.Misc) -> list[tk.Misc]:
    widgets = [widget]
    for child in widget.winfo_children():
        widgets.extend(iter_widgets(child))
    return widgets


def run_footer_link_check() -> int:
    opened: list[str] = []
    original_open = webbrowser.open

    def fake_open(url: str, new: int = 0, autoraise: bool = True) -> bool:
        opened.append(url)
        return True

    webbrowser.open = fake_open
    root = make_root()
    app = DakePdfExtractApp(root)
    try:
        root.deiconify()
        root.lift()
        root.update()
    except tk.TclError:
        pass
    root.update()
    try:
        targets = {
            UI_TEXT["footer_link_1"]: LINKS["assessment"],
            UI_TEXT["footer_link_2"]: LINKS["instagram"],
        }
        for label_text in targets:
            for widget in iter_widgets(root):
                try:
                    if str(widget.cget("text")) == label_text:
                        widget.event_generate("<Button-1>", x=2, y=2)
                        root.update()
                        break
                except tk.TclError:
                    continue
    finally:
        app.close()
        webbrowser.open = original_open
    expected = {LINKS["assessment"], LINKS["instagram"]}
    if not expected.issubset(set(opened)):
        print(f"opened={opened}", file=sys.stderr)
        return 1
    print(UI_TEXT["footer_link_check_ok"])
    return 0


def run_layout_check() -> int:
    root = make_root()
    app = DakePdfExtractApp(root)
    try:
        expected_texts = {
            UI_TEXT["footer_left"],
            UI_TEXT["footer_tagline"],
            UI_TEXT["footer_link_1"],
            UI_TEXT["footer_link_2"],
            UI_TEXT["footer_copyright"],
        }
        actual_texts: set[str] = set()
        separator_count = 0
        for widget in iter_widgets(root):
            try:
                text = str(widget.cget("text"))
            except tk.TclError:
                continue
            if text in expected_texts:
                actual_texts.add(text)
            if text == UI_TEXT["footer_separator"]:
                separator_count += 1
        if actual_texts != expected_texts or separator_count < 3:
            return 1
        root.geometry("820x700")
        root.update()
        app._update_footer_layout()
        compact_ok = app.footer_compact is True
        root.geometry("1080x760")
        root.update()
        app._update_footer_layout()
        wide_ok = app.footer_compact is False
        icon_ok = app.window_icon_ok
    finally:
        app.close()
    if not (compact_ok and wide_ok and icon_ok):
        return 1
    print(UI_TEXT["layout_check_ok"])
    return 0


def run_demo_screenshot(output_path: Path) -> int:
    if ImageGrab is None:
        print(UI_TEXT["error_dependency_missing"], file=sys.stderr)
        return 1
    with tempfile.TemporaryDirectory(prefix="dake_pdf_extract_screen_") as temp_dir:
        temp = Path(temp_dir)
        source = temp / UI_TEXT["sample_pdf_name"]
        create_sample_pdf(source, 8)
        root = make_root()
        app = DakePdfExtractApp(root)
        app.root.geometry(WINDOW_SIZE)
        try:
            root.deiconify()
            root.lift()
            root.attributes("-topmost", True)
            root.focus_force()
            root.update()
        except tk.TclError:
            pass
        app.load_pdf(source)
        deadline = time.time() + 12
        while time.time() < deadline:
            root.update()
            if app.page_count == 8 and len(app.thumbnail_images) >= 8:
                break
            time.sleep(0.05)
        app.extracted_pages.update({1, 5})
        app.select_pages({2, 3, 4})
        app.thumbnail_size_var.set(170)
        app.on_thumbnail_size_changed()
        for _ in range(10):
            try:
                root.lift()
                root.focus_force()
            except tk.TclError:
                pass
            root.update()
            time.sleep(0.05)
        try:
            root.attributes("-topmost", True)
            root.update()
            time.sleep(0.2)
        except tk.TclError:
            pass
        x = root.winfo_rootx()
        y = root.winfo_rooty()
        width = root.winfo_width()
        height = root.winfo_height()
        image = ImageGrab.grab(bbox=(x, y, x + width, y + height))
        output_path.parent.mkdir(parents=True, exist_ok=True)
        image.save(output_path)
        if output_path.name == "screenshot.webp":
            rgb = image.convert("RGB")
            rgb.save(output_path.with_suffix(".jpg"), quality=92)
            booth = Image.new("RGB", (1200, 630), (246, 247, 249))
            thumbnail = rgb.copy()
            thumbnail.thumbnail((1200, 630))
            paste_x = (1200 - thumbnail.width) // 2
            paste_y = (630 - thumbnail.height) // 2
            booth.paste(thumbnail, (paste_x, paste_y))
            booth.save(output_path.parent / "booth_thumbnail.jpg", quality=92)
        app.close()
    print(UI_TEXT["screenshot_check_ok"])
    return 0


def main(argv: list[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    if "--launch-check" in args:
        return run_launch_check()
    if "--process-check" in args:
        return run_process_check()
    if "--self-check" in args:
        return run_self_check()
    if "--footer-link-check" in args:
        return run_footer_link_check()
    if "--layout-check" in args:
        return run_layout_check()
    if "--demo-screenshot" in args:
        index = args.index("--demo-screenshot")
        try:
            output_path = Path(args[index + 1])
        except IndexError:
            return 1
        return run_demo_screenshot(output_path)

    set_app_user_model_id()
    root = make_root()
    DakePdfExtractApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
