# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import io
import math
import os
import queue
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import webbrowser
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, ttk

from PIL import Image, ImageOps, ImageTk, UnidentifiedImageError

try:
    import pillow_heif

    pillow_heif.register_heif_opener()
    HEIF_READY = True
except Exception:
    HEIF_READY = False

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas as pdf_canvas

    REPORTLAB_READY = True
except Exception:
    A4 = (595.275590551, 841.88976378)
    ImageReader = None
    pdf_canvas = None
    REPORTLAB_READY = False

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    ROOT_CLASS = TkinterDnD.Tk
    DND_READY = True
except Exception:
    DND_FILES = ""
    ROOT_CLASS = tk.Tk
    DND_READY = False


APP_NAME = "Dake画像まとめPDF"
WINDOW_TITLE = "画像まとめPDF"
EXE_NAME = "DakeImage_BatchPDF.exe"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "brand_series": "シンプルそれDAKEシリーズ",
    "footer_phrase": "止まらない、迷わない、すぐ終わる。",
    "main_title": "画像をまとめてPDFにする",
    "main_description": "複数の画像を並べて、1つのPDFにまとめます。",
    "add_title": "画像を追加",
    "empty_title": "画像を追加してください",
    "empty_subtitle": "ドラッグ＆ドロップ または クリックして追加",
    "empty_subtitle_click_only": "クリックして追加",
    "supported_formats": "JPG / JPEG / PNG / BMP / WEBP / HEIC / HEIF",
    "list_title": "追加画像一覧",
    "layout_title": "PDFの配置",
    "layout_one": "1ページに1枚",
    "layout_four": "1ページに4枚",
    "button_add": "画像を追加",
    "button_up": "上へ",
    "button_down": "下へ",
    "button_delete": "選択画像を削除",
    "button_clear": "すべてクリア",
    "button_execute": "PDFにして保存",
    "button_cancel": "キャンセル",
    "status_idle": "未選択",
    "status_loading": "読み込み中",
    "status_ready": "準備完了",
    "status_processing": "処理中",
    "status_saving": "保存中",
    "status_complete": "保存完了",
    "status_cancelled": "キャンセルしました",
    "status_error": "エラー",
    "status_idle_detail": "画像を追加してください。",
    "status_add_prepare": "画像を準備中 {current} / {total}",
    "status_add_complete": "{added}件追加しました。",
    "status_add_complete_with_skipped": "{added}件追加 / {skipped}件スキップしました。",
    "status_add_none": "追加できる画像がありませんでした。",
    "status_add_duplicate": "同じ画像は追加済みです。",
    "status_reordered": "順番を変更しました。",
    "status_deleted": "選択画像を削除しました。",
    "status_cleared": "すべてクリアしました。",
    "status_prepare_image": "画像を準備中 {current} / {total}",
    "status_create_pdf": "PDFを作成中 {current} / {total}ページ",
    "status_saving_detail": "保存中",
    "status_complete_detail": "保存先フォルダを開きます。",
    "status_cancel_request": "キャンセル中です。安全なところで止めています。",
    "dialog_select_title": "画像を選択してください",
    "dialog_save_title": "PDFとして保存",
    "dialog_complete_title": "保存が完了しました",
    "dialog_error_title": "処理できませんでした",
    "dialog_warning_title": "確認してください",
    "dialog_cancel_title": "キャンセルしました",
    "message_complete": "PDFを保存しました。\n\n保存先:\n{path}",
    "message_cancelled": "PDF作成をキャンセルしました。元画像は変更していません。",
    "message_no_images": "先に画像を追加してください。",
    "message_no_valid_images": "読み込める画像がありませんでした。",
    "message_reportlab_missing": "PDF作成に必要なライブラリが見つかりません。requirements.txt の内容をインストールしてください。",
    "message_heic_missing": "HEIC / HEIFを読むためのライブラリが見つかりません。requirements.txt の内容をインストールしてください。",
    "message_image_open_failed": "{name} を読み込めませんでした。画像ファイルを確認してから、もう一度お試しください。",
    "message_image_broken": "{name} を読み込めませんでした。ファイルが壊れている可能性があります。",
    "message_save_folder_invalid": "保存先フォルダが見つかりません。別の保存先を選んでください。",
    "message_save_failed": "PDFを保存できませんでした。保存先やファイル名を確認してから、もう一度お試しください。",
    "message_memory": "画像が大きすぎて処理できませんでした。枚数を減らすか、画像サイズを小さくしてお試しください。",
    "message_unknown_error": "処理中に問題が起きました。画像と保存先を確認してから、もう一度お試しください。",
    "message_open_folder_failed": "保存先フォルダを開けませんでした。手動で確認してください。",
    "thumbnail_loading": "準備中",
    "thumbnail_error": "表示できません",
    "file_type_images": "画像ファイル",
    "file_type_all": "すべてのファイル",
    "output_name_one": "Dake_画像まとめ_1枚配置_{timestamp}.pdf",
    "output_name_four": "Dake_画像まとめ_4枚配置_{timestamp}.pdf",
    "output_name_date_format": "%Y%m%d_%H%M%S",
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
    "success": "#12B76A",
    "success_bg": "#EAFBF3",
    "danger": "#D92D20",
    "danger_bg": "#FDECEC",
    "white": "#FFFFFF",
    "disabled": "#D0D5DD",
}

FOOTER_URLS = {
    "footer_link_1": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "footer_link_2": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

SUPPORTED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".webp", ".heic", ".heif"}
IMAGE_FILE_PATTERN = "*.jpg *.jpeg *.png *.bmp *.webp *.heic *.heif"
LAYOUT_ONE = "one"
LAYOUT_FOUR = "four"
WINDOW_SIZE = "980x760"
WINDOW_MIN_SIZE = (880, 660)
POLL_INTERVAL_MS = 60
THUMBNAIL_SIZE = (96, 68)
THUMBNAIL_SOURCE_MAX_EDGE = 1400
PDF_IMAGE_MAX_EDGE = 2600
LIST_PAD_X = 14
LIST_PAD_Y = 14
ROW_HEIGHT = 86
ROW_GAP = 10
OUTER_MARGIN_PT = 34
GUTTER_PT = 18
RESAMPLE = Image.Resampling.LANCZOS if hasattr(Image, "Resampling") else Image.LANCZOS
ROTATION_AREA_EPSILON = 0.0001
MAX_UNIQUE_PATH_ATTEMPTS = 10000


@dataclass
class ImageItem:
    path: Path
    path_key: str
    photo: ImageTk.PhotoImage | None
    pixel_size: tuple[int, int]


@dataclass(frozen=True)
class LoadedImage:
    path: Path
    path_key: str
    thumbnail_png: bytes
    pixel_size: tuple[int, int]


@dataclass(frozen=True)
class SlotPlacement:
    rotate_clockwise: bool
    draw_width: float
    draw_height: float


class ImageReadError(RuntimeError):
    def __init__(self, code: str, path: Path, original: Exception | None = None) -> None:
        super().__init__(code)
        self.code = code
        self.path = path
        self.original = original


class CancelledError(RuntimeError):
    pass


class FlatButton(tk.Button):
    def __init__(
        self,
        parent: tk.Misc,
        *,
        text_key: str,
        command: Any,
        font: tuple[str, int] | tuple[str, int, str],
        primary: bool = False,
    ) -> None:
        self.primary = primary
        self.normal_bg = THEME["accent"] if primary else THEME["card"]
        self.hover_bg = THEME["accent_hover"] if primary else THEME["soft"]
        self.normal_fg = THEME["white"] if primary else THEME["text"]
        super().__init__(
            parent,
            text=UI_TEXT[text_key],
            command=command,
            bg=self.normal_bg,
            fg=self.normal_fg,
            activebackground=self.hover_bg,
            activeforeground=self.normal_fg,
            disabledforeground=THEME["white"] if primary else THEME["muted"],
            relief="solid",
            bd=0 if primary else 1,
            highlightthickness=1 if not primary else 0,
            highlightbackground=THEME["border"],
            padx=14,
            pady=8,
            cursor="hand2",
            font=font,
        )
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)

    def _on_enter(self, _event: tk.Event) -> None:
        if str(self["state"]) == tk.DISABLED:
            return
        self.configure(bg=self.hover_bg)

    def _on_leave(self, _event: tk.Event) -> None:
        if str(self["state"]) == tk.DISABLED:
            return
        self.configure(bg=self.normal_bg)

    def set_enabled(self, enabled: bool) -> None:
        if enabled:
            self.configure(state=tk.NORMAL, bg=self.normal_bg, fg=self.normal_fg, cursor="hand2")
        else:
            self.configure(state=tk.DISABLED, bg=THEME["disabled"], cursor="arrow")


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def icon_candidates() -> list[Path]:
    base_dir = app_dir()
    source_dir = Path(__file__).resolve().parent
    return [
        source_dir / ".." / ".." / "02_assets" / "dake_icon.ico",
        base_dir / ".." / ".." / "02_assets" / "dake_icon.ico",
        base_dir / ".." / ".." / ".." / "02_assets" / "dake_icon.ico",
        Path(getattr(sys, "_MEIPASS", base_dir)) / "dake_icon.ico",
        base_dir / "dake_icon.ico",
    ]


def choose_font_family(root: tk.Tk) -> str:
    try:
        families = set(tkfont.families(root))
    except Exception:
        return "TkDefaultFont"
    for family in ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo"):
        if family in families:
            return family
    return "TkDefaultFont"


def normalize_path(path: Path) -> Path:
    try:
        return path.expanduser().resolve()
    except Exception:
        return path.expanduser().absolute()


def path_key(path: Path) -> str:
    return str(normalize_path(path)).casefold()


def is_supported_image(path: Path) -> bool:
    return path.suffix.lower() in SUPPORTED_EXTENSIONS


def default_output_dir() -> Path:
    downloads = Path.home() / "Downloads"
    return downloads if downloads.exists() else Path.home()


def default_output_filename(layout: str, now: datetime | None = None) -> str:
    timestamp = (now or datetime.now()).strftime(UI_TEXT["output_name_date_format"])
    template_key = "output_name_four" if layout == LAYOUT_FOUR else "output_name_one"
    return UI_TEXT[template_key].format(timestamp=timestamp)


def ensure_pdf_suffix(path: Path) -> Path:
    if path.suffix.lower() == ".pdf":
        return path
    return path.with_suffix(".pdf")


def ensure_unique_path(path: Path) -> Path:
    path = ensure_pdf_suffix(path)
    if not path.exists():
        return path

    stem = path.stem
    suffix = path.suffix
    for index in range(2, MAX_UNIQUE_PATH_ATTEMPTS + 1):
        candidate = path.with_name(f"{stem}_{index}{suffix}")
        if not candidate.exists():
            return candidate

    raise RuntimeError(UI_TEXT["message_save_failed"])


def commit_temp_to_unique_path(temp_path: Path, output_path: Path) -> Path:
    for _attempt in range(MAX_UNIQUE_PATH_ATTEMPTS):
        candidate = ensure_unique_path(output_path)
        try:
            os.rename(temp_path, candidate)
        except OSError:
            if candidate.exists():
                continue
            raise
        return candidate
    raise RuntimeError(UI_TEXT["message_save_failed"])


def open_folder(path: Path) -> bool:
    try:
        if os.name == "nt":
            os.startfile(str(path))  # type: ignore[attr-defined]
            return True
        if sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
            return True
        subprocess.Popen(["xdg-open", str(path)])
        return True
    except Exception:
        return False


def write_error_log(context: str, exc: BaseException) -> None:
    try:
        log_dir = app_dir() / "logs"
        log_dir.mkdir(parents=True, exist_ok=True)
        log_path = log_dir / "image_batchpdf_error.log"
        text = [
            time.strftime("%Y-%m-%d %H:%M:%S"),
            context,
            "".join(traceback.format_exception(type(exc), exc, exc.__traceback__)),
            "",
        ]
        with log_path.open("a", encoding="utf-8") as handle:
            handle.write("\n".join(text))
    except Exception:
        pass


def normalize_image(image: Image.Image) -> Image.Image:
    if image.mode == "RGB" and "transparency" not in image.info:
        return image.copy()

    if "A" in image.getbands() or image.mode == "P" and "transparency" in image.info:
        rgba_image = image.convert("RGBA")
        background = Image.new("RGB", rgba_image.size, "white")
        background.paste(rgba_image, mask=rgba_image.getchannel("A"))
        rgba_image.close()
        return background

    return image.convert("RGB")


def load_display_image(source_path: Path, max_edge: int) -> Image.Image:
    if source_path.suffix.lower() in {".heic", ".heif"} and not HEIF_READY:
        raise ImageReadError("message_heic_missing", source_path)

    try:
        with Image.open(source_path) as raw_image:
            raw_image.load()
            transposed = ImageOps.exif_transpose(raw_image)
            try:
                image = normalize_image(transposed)
            finally:
                if transposed is not raw_image:
                    transposed.close()
    except MemoryError as exc:
        raise ImageReadError("message_memory", source_path, exc) from exc
    except UnidentifiedImageError as exc:
        raise ImageReadError("message_image_broken", source_path, exc) from exc
    except OSError as exc:
        raise ImageReadError("message_image_open_failed", source_path, exc) from exc
    except ImageReadError:
        raise
    except Exception as exc:
        raise ImageReadError("message_image_open_failed", source_path, exc) from exc

    if max(image.size) > max_edge:
        image.thumbnail((max_edge, max_edge), RESAMPLE)
    return image


def make_thumbnail_png(source_path: Path) -> tuple[bytes, tuple[int, int]]:
    image = load_display_image(source_path, THUMBNAIL_SOURCE_MAX_EDGE)
    try:
        original_size = image.size
        image.thumbnail(THUMBNAIL_SIZE, RESAMPLE)
        thumbnail = Image.new("RGB", THUMBNAIL_SIZE, "white")
        x = (THUMBNAIL_SIZE[0] - image.width) // 2
        y = (THUMBNAIL_SIZE[1] - image.height) // 2
        thumbnail.paste(image, (x, y))
        buffer = io.BytesIO()
        thumbnail.save(buffer, format="PNG")
        return buffer.getvalue(), original_size
    finally:
        image.close()


def fit_rect(source_size: tuple[int, int], box_width: float, box_height: float) -> tuple[float, float]:
    image_width, image_height = source_size
    if image_width <= 0 or image_height <= 0:
        return 1.0, 1.0
    scale = min(box_width / image_width, box_height / image_height)
    return image_width * scale, image_height * scale


def choose_best_orientation(source_size: tuple[int, int], box_width: float, box_height: float) -> SlotPlacement:
    image_width, image_height = source_size
    draw_width_0, draw_height_0 = fit_rect((image_width, image_height), box_width, box_height)
    area_0 = draw_width_0 * draw_height_0

    draw_width_90, draw_height_90 = fit_rect((image_height, image_width), box_width, box_height)
    area_90 = draw_width_90 * draw_height_90

    if area_90 > area_0 * (1.0 + ROTATION_AREA_EPSILON):
        return SlotPlacement(True, draw_width_90, draw_height_90)
    return SlotPlacement(False, draw_width_0, draw_height_0)


def rotate_clockwise_90(image: Image.Image) -> Image.Image:
    if hasattr(Image, "Transpose"):
        return image.transpose(Image.Transpose.ROTATE_270)
    return image.rotate(-90, expand=True)


def prepare_image_for_slot(image: Image.Image, box_width: float, box_height: float) -> tuple[Image.Image, SlotPlacement]:
    placement = choose_best_orientation(image.size, box_width, box_height)
    if placement.rotate_clockwise:
        return rotate_clockwise_90(image), placement
    return image, placement


def draw_image_centered(pdf: Any, source_path: Path, x: float, y: float, width: float, height: float) -> None:
    image = load_display_image(source_path, PDF_IMAGE_MAX_EDGE)
    if ImageReader is None:
        image.close()
        raise RuntimeError(UI_TEXT["message_reportlab_missing"])
    slot_image: Image.Image | None = None
    try:
        slot_image, placement = prepare_image_for_slot(image, width, height)
        reader = ImageReader(slot_image)
        draw_x = x + (width - placement.draw_width) / 2
        draw_y = y + (height - placement.draw_height) / 2
        pdf.drawImage(
            reader,
            draw_x,
            draw_y,
            placement.draw_width,
            placement.draw_height,
            preserveAspectRatio=True,
            mask="auto",
        )
    finally:
        if slot_image is not None and slot_image is not image:
            slot_image.close()
        image.close()


def generate_pdf(
    source_paths: list[Path],
    output_path: Path,
    layout: str,
    cancel_event: threading.Event,
    event_queue: queue.Queue[tuple[Any, ...]] | None = None,
) -> Path:
    if not REPORTLAB_READY or pdf_canvas is None:
        raise RuntimeError(UI_TEXT["message_reportlab_missing"])
    if not source_paths:
        raise RuntimeError(UI_TEXT["message_no_images"])

    output_path = ensure_pdf_suffix(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    handle, temp_name = tempfile.mkstemp(
        prefix=f".{output_path.stem}_",
        suffix=".tmp.pdf",
        dir=str(output_path.parent),
    )
    os.close(handle)
    temp_path = Path(temp_name)

    try:
        page_width, page_height = A4
        pdf = pdf_canvas.Canvas(str(temp_path), pagesize=A4)
        total_images = len(source_paths)
        total_pages = total_images if layout == LAYOUT_ONE else math.ceil(total_images / 4)

        def publish(kind: str, current: int, total: int) -> None:
            if event_queue is not None:
                event_queue.put(("pdf_progress", kind, current, total))

        if layout == LAYOUT_ONE:
            box_x = OUTER_MARGIN_PT
            box_y = OUTER_MARGIN_PT
            box_width = page_width - OUTER_MARGIN_PT * 2
            box_height = page_height - OUTER_MARGIN_PT * 2
            for index, source_path in enumerate(source_paths, start=1):
                if cancel_event.is_set():
                    raise CancelledError()
                publish("prepare", index, total_images)
                draw_image_centered(pdf, source_path, box_x, box_y, box_width, box_height)
                publish("create", index, total_pages)
                pdf.showPage()
        else:
            cell_width = (page_width - OUTER_MARGIN_PT * 2 - GUTTER_PT) / 2
            cell_height = (page_height - OUTER_MARGIN_PT * 2 - GUTTER_PT) / 2
            cells = [
                (OUTER_MARGIN_PT, OUTER_MARGIN_PT + cell_height + GUTTER_PT),
                (OUTER_MARGIN_PT + cell_width + GUTTER_PT, OUTER_MARGIN_PT + cell_height + GUTTER_PT),
                (OUTER_MARGIN_PT, OUTER_MARGIN_PT),
                (OUTER_MARGIN_PT + cell_width + GUTTER_PT, OUTER_MARGIN_PT),
            ]
            image_index = 0
            for page_index in range(total_pages):
                if cancel_event.is_set():
                    raise CancelledError()
                page_sources = source_paths[page_index * 4 : page_index * 4 + 4]
                for cell_index, source_path in enumerate(page_sources):
                    if cancel_event.is_set():
                        raise CancelledError()
                    image_index += 1
                    publish("prepare", image_index, total_images)
                    cell_x, cell_y = cells[cell_index]
                    draw_image_centered(pdf, source_path, cell_x, cell_y, cell_width, cell_height)
                publish("create", page_index + 1, total_pages)
                pdf.showPage()

        if cancel_event.is_set():
            raise CancelledError()
        if event_queue is not None:
            event_queue.put(("pdf_progress", "saving", 1, 1))
        pdf.save()
        return commit_temp_to_unique_path(temp_path, output_path)
    except CancelledError:
        raise
    except ImageReadError:
        raise
    except MemoryError as exc:
        raise ImageReadError("message_memory", output_path, exc) from exc
    except Exception:
        raise
    finally:
        try:
            temp_path.unlink(missing_ok=True)
        except Exception:
            pass


class DakeImageBatchPdfApp:
    def __init__(self) -> None:
        self.root = ROOT_CLASS()
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(*WINDOW_MIN_SIZE)
        self.root.configure(bg=THEME["background"])
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        self.font_family = choose_font_family(self.root)
        self.fonts = {
            "title": (self.font_family, 20, "bold"),
            "description": (self.font_family, 10),
            "section": (self.font_family, 11, "bold"),
            "body": (self.font_family, 10),
            "small": (self.font_family, 9),
            "footer": (self.font_family, 8),
            "button": (self.font_family, 10, "bold"),
            "row_title": (self.font_family, 10, "bold"),
        }

        self.items: list[ImageItem] = []
        self.item_keys: set[str] = set()
        self.selected_index: int | None = None
        self.event_queue: queue.Queue[tuple[Any, ...]] = queue.Queue()
        self.is_loading = False
        self.is_generating = False
        self.cancel_event: threading.Event | None = None
        self.close_after_worker = False
        self.footer_stacked: bool | None = None

        self.layout_var = tk.StringVar(value=LAYOUT_ONE)
        self.status_var = tk.StringVar(value=UI_TEXT["status_idle"])
        self.detail_var = tk.StringVar(value=UI_TEXT["status_idle_detail"])

        self.apply_window_icon()
        self.configure_styles()
        self.build_ui()
        self.register_drop_targets()
        self.render_image_list()
        self.update_buttons()
        self.root.after(POLL_INTERVAL_MS, self.poll_events)

    def apply_window_icon(self) -> None:
        for candidate in icon_candidates():
            try:
                resolved = candidate.resolve()
                if resolved.exists():
                    self.root.iconbitmap(str(resolved))
                    return
            except Exception:
                continue

    def configure_styles(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(
            "Dake.Vertical.TScrollbar",
            troughcolor=THEME["background"],
            background="#C9D2E0",
            bordercolor=THEME["background"],
            arrowcolor=THEME["muted"],
            relief="flat",
            width=14,
        )

    def build_ui(self) -> None:
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(2, weight=1)

        self.build_header()
        self.build_add_area()
        self.build_list_area()
        self.build_layout_area()
        self.build_status_and_actions()
        self.build_footer()

    def build_header(self) -> None:
        header = tk.Frame(self.root, bg=THEME["background"])
        header.grid(row=0, column=0, sticky="ew", padx=24, pady=(20, 12))
        header.grid_columnconfigure(0, weight=1)
        tk.Label(
            header,
            text=UI_TEXT["main_title"],
            font=self.fonts["title"],
            fg=THEME["text"],
            bg=THEME["background"],
            anchor="w",
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            header,
            text=UI_TEXT["main_description"],
            font=self.fonts["description"],
            fg=THEME["muted"],
            bg=THEME["background"],
            anchor="w",
        ).grid(row=1, column=0, sticky="w", pady=(5, 0))

    def build_add_area(self) -> None:
        self.add_area = tk.Frame(
            self.root,
            bg=THEME["card"],
            highlightthickness=1,
            highlightbackground=THEME["border"],
            highlightcolor=THEME["accent"],
            cursor="hand2",
        )
        self.add_area.grid(row=1, column=0, sticky="ew", padx=24, pady=(0, 12))
        self.add_area.grid_columnconfigure(0, weight=1)
        self.add_area.bind("<Button-1>", self.choose_images)

        add_inner = tk.Frame(self.add_area, bg=THEME["card"], cursor="hand2")
        add_inner.grid(row=0, column=0, sticky="ew", padx=20, pady=16)
        add_inner.grid_columnconfigure(0, weight=1)
        add_inner.bind("<Button-1>", self.choose_images)
        tk.Label(
            add_inner,
            text=UI_TEXT["empty_title"],
            font=(self.font_family, 13, "bold"),
            fg=THEME["text"],
            bg=THEME["card"],
            cursor="hand2",
        ).grid(row=0, column=0, sticky="w")
        subtitle = UI_TEXT["empty_subtitle"] if DND_READY else UI_TEXT["empty_subtitle_click_only"]
        tk.Label(
            add_inner,
            text=subtitle,
            font=self.fonts["body"],
            fg=THEME["muted"],
            bg=THEME["card"],
            cursor="hand2",
        ).grid(row=1, column=0, sticky="w", pady=(5, 0))
        tk.Label(
            add_inner,
            text=UI_TEXT["supported_formats"],
            font=self.fonts["small"],
            fg=THEME["muted"],
            bg=THEME["card"],
            cursor="hand2",
        ).grid(row=0, column=1, rowspan=2, sticky="e", padx=(16, 0))

    def build_list_area(self) -> None:
        panel = tk.Frame(
            self.root,
            bg=THEME["card"],
            highlightthickness=1,
            highlightbackground=THEME["border"],
            highlightcolor=THEME["border"],
        )
        panel.grid(row=2, column=0, sticky="nsew", padx=24, pady=(0, 12))
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(1, weight=1)

        tk.Label(
            panel,
            text=UI_TEXT["list_title"],
            font=self.fonts["section"],
            fg=THEME["text"],
            bg=THEME["card"],
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=16, pady=(13, 8))

        canvas_frame = tk.Frame(panel, bg=THEME["card"])
        canvas_frame.grid(row=1, column=0, sticky="nsew", padx=16, pady=(0, 14))
        canvas_frame.grid_columnconfigure(0, weight=1)
        canvas_frame.grid_rowconfigure(0, weight=1)

        self.list_canvas = tk.Canvas(
            canvas_frame,
            bg=THEME["background"],
            highlightthickness=1,
            highlightbackground=THEME["border"],
            highlightcolor=THEME["selection_border"],
            bd=0,
        )
        self.list_canvas.grid(row=0, column=0, sticky="nsew")
        self.list_scrollbar = ttk.Scrollbar(
            canvas_frame,
            orient="vertical",
            command=self.list_canvas.yview,
            style="Dake.Vertical.TScrollbar",
        )
        self.list_scrollbar.grid(row=0, column=1, sticky="ns")
        self.list_canvas.configure(yscrollcommand=self.list_scrollbar.set)
        self.list_canvas.bind("<Configure>", lambda _event: self.render_image_list())
        self.list_canvas.bind("<Button-1>", self.on_canvas_click)
        self.list_canvas.bind("<MouseWheel>", self.on_mouse_wheel)
        self.list_canvas.bind("<Enter>", lambda _event: self.list_canvas.focus_set())
        self.root.bind("<Delete>", lambda _event: self.delete_selected(), add="+")

    def build_layout_area(self) -> None:
        frame = tk.Frame(self.root, bg=THEME["background"])
        frame.grid(row=3, column=0, sticky="ew", padx=24, pady=(0, 10))
        frame.grid_columnconfigure(1, weight=1)

        tk.Label(
            frame,
            text=UI_TEXT["layout_title"],
            font=self.fonts["section"],
            fg=THEME["text"],
            bg=THEME["background"],
        ).grid(row=0, column=0, sticky="w", padx=(0, 16))

        options = tk.Frame(frame, bg=THEME["background"])
        options.grid(row=0, column=1, sticky="w")
        for column, (value, text_key) in enumerate(((LAYOUT_ONE, "layout_one"), (LAYOUT_FOUR, "layout_four"))):
            tk.Radiobutton(
                options,
                text=UI_TEXT[text_key],
                value=value,
                variable=self.layout_var,
                bg=THEME["background"],
                fg=THEME["text"],
                selectcolor=THEME["card"],
                activebackground=THEME["background"],
                activeforeground=THEME["text"],
                font=self.fonts["body"],
                command=self.update_buttons,
            ).grid(row=0, column=column, sticky="w", padx=(0, 18))

    def build_status_and_actions(self) -> None:
        frame = tk.Frame(self.root, bg=THEME["background"])
        frame.grid(row=4, column=0, sticky="ew", padx=24, pady=(0, 12))
        frame.grid_columnconfigure(0, weight=1)

        status = tk.Frame(frame, bg=THEME["background"])
        status.grid(row=0, column=0, sticky="ew")
        status.grid_columnconfigure(1, weight=1)

        self.status_badge = tk.Label(
            status,
            textvariable=self.status_var,
            font=(self.font_family, 9, "bold"),
            fg=THEME["muted"],
            bg=THEME["soft"],
            padx=12,
            pady=5,
        )
        self.status_badge.grid(row=0, column=0, sticky="w")
        tk.Label(
            status,
            textvariable=self.detail_var,
            font=self.fonts["small"],
            fg=THEME["muted"],
            bg=THEME["background"],
            anchor="w",
        ).grid(row=0, column=1, sticky="ew", padx=(12, 0))

        actions = tk.Frame(frame, bg=THEME["background"])
        actions.grid(row=1, column=0, sticky="ew", pady=(11, 0))
        actions.grid_columnconfigure(0, weight=1)

        left = tk.Frame(actions, bg=THEME["background"])
        left.grid(row=0, column=0, sticky="w")
        right = tk.Frame(actions, bg=THEME["background"])
        right.grid(row=0, column=1, sticky="e")

        self.add_button = FlatButton(left, text_key="button_add", command=self.choose_images, font=self.fonts["button"])
        self.add_button.pack(side="left", padx=(0, 8))
        self.up_button = FlatButton(left, text_key="button_up", command=self.move_selected_up, font=self.fonts["button"])
        self.up_button.pack(side="left", padx=(0, 8))
        self.down_button = FlatButton(left, text_key="button_down", command=self.move_selected_down, font=self.fonts["button"])
        self.down_button.pack(side="left", padx=(0, 8))
        self.delete_button = FlatButton(left, text_key="button_delete", command=self.delete_selected, font=self.fonts["button"])
        self.delete_button.pack(side="left", padx=(0, 8))
        self.clear_button = FlatButton(left, text_key="button_clear", command=self.clear_all, font=self.fonts["button"])
        self.clear_button.pack(side="left")

        self.cancel_button = FlatButton(right, text_key="button_cancel", command=self.cancel_pdf, font=self.fonts["button"])
        self.cancel_button.pack(side="left", padx=(0, 8))
        self.execute_button = FlatButton(
            right,
            text_key="button_execute",
            command=self.save_pdf,
            font=self.fonts["button"],
            primary=True,
        )
        self.execute_button.pack(side="left")

    def build_footer(self) -> None:
        self.footer = tk.Frame(self.root, bg=THEME["background"])
        self.footer.grid(row=5, column=0, sticky="ew", padx=24, pady=(0, 16))
        self.footer.grid_columnconfigure(0, weight=1)
        self.footer.grid_columnconfigure(1, weight=1)

        self.footer_left = tk.Frame(self.footer, bg=THEME["background"])
        self.footer_right = tk.Frame(self.footer, bg=THEME["background"])

        for key in ("footer_left", "footer_separator", "footer_phrase"):
            tk.Label(
                self.footer_left,
                text=UI_TEXT[key],
                font=self.fonts["footer"],
                fg=THEME["muted"],
                bg=THEME["background"],
            ).pack(side="left")

        self.make_footer_link(self.footer_right, "footer_link_1").pack(side="left")
        self.make_footer_text(self.footer_right, "footer_separator").pack(side="left")
        self.make_footer_link(self.footer_right, "footer_link_2").pack(side="left")
        self.make_footer_text(self.footer_right, "footer_separator").pack(side="left")
        self.make_footer_text(self.footer_right, "footer_copyright").pack(side="left")

        self.root.bind("<Configure>", lambda event: self.layout_footer(event.width), add="+")
        self.layout_footer(self.root.winfo_width())

    def make_footer_text(self, parent: tk.Misc, text_key: str) -> tk.Label:
        return tk.Label(
            parent,
            text=UI_TEXT[text_key],
            font=self.fonts["footer"],
            fg=THEME["muted"],
            bg=THEME["background"],
        )

    def make_footer_link(self, parent: tk.Misc, text_key: str) -> tk.Label:
        label = self.make_footer_text(parent, text_key)
        label.configure(cursor="hand2")
        label.bind("<Button-1>", lambda _event, key=text_key: webbrowser.open(FOOTER_URLS[key], new=2))
        label.bind("<Enter>", lambda _event: label.configure(fg=THEME["accent"]))
        label.bind("<Leave>", lambda _event: label.configure(fg=THEME["muted"]))
        return label

    def layout_footer(self, width: int) -> None:
        stacked = width < 900
        if self.footer_stacked == stacked:
            return
        self.footer_stacked = stacked
        self.footer_left.grid_forget()
        self.footer_right.grid_forget()
        if stacked:
            self.footer_left.grid(row=0, column=0, columnspan=2, sticky="", pady=(0, 3))
            self.footer_right.grid(row=1, column=0, columnspan=2, sticky="")
            return
        self.footer_left.grid(row=0, column=0, sticky="w")
        self.footer_right.grid(row=0, column=1, sticky="e")

    def register_drop_targets(self) -> None:
        if not DND_READY or not DND_FILES:
            return
        widgets = [self.add_area, self.list_canvas]
        for widget in widgets:
            try:
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", self.on_drop)
            except Exception:
                continue

    def choose_images(self, _event: tk.Event | None = None) -> None:
        if self.is_busy():
            return
        selected = filedialog.askopenfilenames(
            parent=self.root,
            title=UI_TEXT["dialog_select_title"],
            filetypes=[
                (UI_TEXT["file_type_images"], IMAGE_FILE_PATTERN),
                (UI_TEXT["file_type_all"], "*.*"),
            ],
        )
        if selected:
            self.start_add_files([Path(path) for path in selected])

    def on_drop(self, event: tk.Event) -> str:
        if self.is_busy():
            return "break"
        self.start_add_files(parse_drop_paths(self.root, getattr(event, "data", "")))
        return "copy"

    def start_add_files(self, raw_paths: Iterable[Path]) -> None:
        candidates: list[Path] = []
        skipped = 0
        seen_in_batch: set[str] = set()
        for raw_path in raw_paths:
            path = normalize_path(raw_path)
            key = path_key(path)
            if key in self.item_keys or key in seen_in_batch:
                skipped += 1
                continue
            if not path.exists() or not path.is_file() or not is_supported_image(path):
                skipped += 1
                continue
            seen_in_batch.add(key)
            candidates.append(path)

        if not candidates:
            if skipped:
                self.set_status(UI_TEXT["status_ready"], UI_TEXT["status_add_none"], "error")
            return

        self.is_loading = True
        self.added_in_current_batch = 0
        self.skipped_in_current_batch = skipped
        self.set_status(
            UI_TEXT["status_loading"],
            UI_TEXT["status_add_prepare"].format(current=0, total=len(candidates)),
            "working",
        )
        self.update_buttons()
        thread = threading.Thread(target=self.load_images_worker, args=(candidates,), daemon=True)
        thread.start()

    def load_images_worker(self, paths: list[Path]) -> None:
        total = len(paths)
        for index, source_path in enumerate(paths, start=1):
            self.event_queue.put(("add_progress", index, total))
            try:
                thumbnail_png, pixel_size = make_thumbnail_png(source_path)
                self.event_queue.put(
                    ("add_item", LoadedImage(source_path, path_key(source_path), thumbnail_png, pixel_size))
                )
            except ImageReadError as exc:
                self.event_queue.put(("add_failed", exc.code, source_path.name, exc))
            except Exception as exc:
                self.event_queue.put(("add_failed", "message_image_open_failed", source_path.name, exc))
        self.event_queue.put(("add_done", total))

    def poll_events(self) -> None:
        try:
            while True:
                event = self.event_queue.get_nowait()
                self.handle_event(event)
        except queue.Empty:
            pass
        if self.root.winfo_exists():
            self.root.after(POLL_INTERVAL_MS, self.poll_events)

    def handle_event(self, event: tuple[Any, ...]) -> None:
        kind = event[0]
        if kind == "add_progress":
            _kind, current, total = event
            self.set_status(
                UI_TEXT["status_loading"],
                UI_TEXT["status_add_prepare"].format(current=current, total=total),
                "working",
            )
            return
        if kind == "add_item":
            _kind, loaded = event
            self.append_loaded_image(loaded)
            return
        if kind == "add_failed":
            _kind, code, name, exc = event
            self.skipped_in_current_batch += 1
            write_error_log(f"add_failed:{name}:{code}", exc)
            return
        if kind == "add_done":
            self.is_loading = False
            added = self.added_in_current_batch
            skipped = self.skipped_in_current_batch
            if added:
                detail = (
                    UI_TEXT["status_add_complete_with_skipped"].format(added=added, skipped=skipped)
                    if skipped
                    else UI_TEXT["status_add_complete"].format(added=added)
                )
                self.set_status(UI_TEXT["status_ready"], detail, "success")
            else:
                self.set_status(UI_TEXT["status_error"], UI_TEXT["message_no_valid_images"], "error")
            self.render_image_list()
            self.update_buttons()
            return
        if kind == "pdf_progress":
            _kind, progress_kind, current, total = event
            if progress_kind == "prepare":
                detail = UI_TEXT["status_prepare_image"].format(current=current, total=total)
                self.set_status(UI_TEXT["status_processing"], detail, "working")
            elif progress_kind == "create":
                detail = UI_TEXT["status_create_pdf"].format(current=current, total=total)
                self.set_status(UI_TEXT["status_processing"], detail, "working")
            else:
                self.set_status(UI_TEXT["status_saving"], UI_TEXT["status_saving_detail"], "working")
            return
        if kind == "pdf_done":
            _kind, output_path = event
            self.finish_pdf_success(Path(output_path))
            return
        if kind == "pdf_cancelled":
            self.finish_pdf_cancelled()
            return
        if kind == "pdf_error":
            _kind, message_key, filename, exc = event
            self.finish_pdf_error(str(message_key), str(filename), exc)

    def append_loaded_image(self, loaded: LoadedImage) -> None:
        if loaded.path_key in self.item_keys:
            self.skipped_in_current_batch += 1
            return
        photo = self.photo_from_png(loaded.thumbnail_png)
        self.items.append(ImageItem(loaded.path, loaded.path_key, photo, loaded.pixel_size))
        self.item_keys.add(loaded.path_key)
        self.selected_index = len(self.items) - 1
        self.added_in_current_batch += 1
        self.render_image_list()

    def photo_from_png(self, data: bytes) -> ImageTk.PhotoImage | None:
        try:
            with Image.open(io.BytesIO(data)) as image:
                return ImageTk.PhotoImage(image)
        except Exception:
            return None

    def render_image_list(self) -> None:
        canvas = self.list_canvas
        canvas.delete("all")
        width = max(canvas.winfo_width(), 420)
        height = max(canvas.winfo_height(), 260)

        if not self.items:
            canvas.configure(scrollregion=(0, 0, width, height))
            center_x = width // 2
            center_y = height // 2
            canvas.create_text(
                center_x,
                center_y - 26,
                text=UI_TEXT["empty_title"],
                fill=THEME["text"],
                font=(self.font_family, 14, "bold"),
                tags=("empty",),
            )
            subtitle = UI_TEXT["empty_subtitle"] if DND_READY else UI_TEXT["empty_subtitle_click_only"]
            canvas.create_text(
                center_x,
                center_y + 2,
                text=subtitle,
                fill=THEME["muted"],
                font=self.fonts["body"],
                tags=("empty",),
            )
            canvas.create_text(
                center_x,
                center_y + 30,
                text=UI_TEXT["supported_formats"],
                fill=THEME["muted"],
                font=self.fonts["small"],
                tags=("empty",),
            )
            return

        total_height = LIST_PAD_Y * 2 + len(self.items) * ROW_HEIGHT + max(0, len(self.items) - 1) * ROW_GAP
        canvas.configure(scrollregion=(0, 0, width, max(height, total_height)))

        row_width = max(260, width - LIST_PAD_X * 2)
        for index, item in enumerate(self.items):
            y = LIST_PAD_Y + index * (ROW_HEIGHT + ROW_GAP)
            self.draw_item_row(index, item, LIST_PAD_X, y, row_width)

    def draw_item_row(self, index: int, item: ImageItem, x: int, y: int, width: int) -> None:
        selected = index == self.selected_index
        fill = THEME["selection_bg"] if selected else THEME["card"]
        outline = THEME["selection_border"] if selected else THEME["border"]
        self.list_canvas.create_rectangle(
            x,
            y,
            x + width,
            y + ROW_HEIGHT,
            fill=fill,
            outline=outline,
            width=2 if selected else 1,
            tags=("row", f"row_{index}"),
        )
        badge_x = x + 14
        badge_y = y + 24
        self.list_canvas.create_rectangle(
            badge_x,
            badge_y,
            badge_x + 34,
            badge_y + 34,
            fill=THEME["accent"],
            outline=THEME["accent"],
            tags=("row", f"row_{index}"),
        )
        self.list_canvas.create_text(
            badge_x + 17,
            badge_y + 17,
            text=str(index + 1),
            fill=THEME["white"],
            font=(self.font_family, 10, "bold"),
            tags=("row", f"row_{index}"),
        )
        thumb_x = x + 66
        thumb_y = y + (ROW_HEIGHT - THUMBNAIL_SIZE[1]) // 2
        self.list_canvas.create_rectangle(
            thumb_x,
            thumb_y,
            thumb_x + THUMBNAIL_SIZE[0],
            thumb_y + THUMBNAIL_SIZE[1],
            fill=THEME["white"],
            outline=THEME["border"],
            tags=("row", f"row_{index}"),
        )
        if item.photo is not None:
            self.list_canvas.create_image(
                thumb_x + THUMBNAIL_SIZE[0] // 2,
                thumb_y + THUMBNAIL_SIZE[1] // 2,
                image=item.photo,
                anchor="center",
                tags=("row", f"row_{index}"),
            )
        else:
            self.list_canvas.create_text(
                thumb_x + THUMBNAIL_SIZE[0] // 2,
                thumb_y + THUMBNAIL_SIZE[1] // 2,
                text=UI_TEXT["thumbnail_error"],
                fill=THEME["muted"],
                font=self.fonts["small"],
                tags=("row", f"row_{index}"),
            )

        text_x = thumb_x + THUMBNAIL_SIZE[0] + 16
        available_width = max(120, width - (text_x - x) - 18)
        self.list_canvas.create_text(
            text_x,
            y + 29,
            text=item.path.name,
            fill=THEME["text"],
            font=self.fonts["row_title"],
            anchor="w",
            width=available_width,
            tags=("row", f"row_{index}"),
        )
        size_text = f"{item.pixel_size[0]} x {item.pixel_size[1]}"
        self.list_canvas.create_text(
            text_x,
            y + 57,
            text=size_text,
            fill=THEME["muted"],
            font=self.fonts["small"],
            anchor="w",
            tags=("row", f"row_{index}"),
        )

    def on_canvas_click(self, event: tk.Event) -> None:
        if not self.items:
            self.choose_images()
            return
        y = self.list_canvas.canvasy(event.y)
        index = int((y - LIST_PAD_Y) // (ROW_HEIGHT + ROW_GAP))
        row_top = LIST_PAD_Y + index * (ROW_HEIGHT + ROW_GAP)
        if 0 <= index < len(self.items) and row_top <= y <= row_top + ROW_HEIGHT:
            self.selected_index = index
            self.render_image_list()
            self.update_buttons()

    def on_mouse_wheel(self, event: tk.Event) -> None:
        if not self.items:
            return
        delta = -1 if event.delta > 0 else 1
        self.list_canvas.yview_scroll(delta * 3, "units")

    def move_selected_up(self) -> None:
        if self.is_busy() or self.selected_index is None or self.selected_index <= 0:
            return
        index = self.selected_index
        self.items[index - 1], self.items[index] = self.items[index], self.items[index - 1]
        self.selected_index = index - 1
        self.render_image_list()
        self.set_status(UI_TEXT["status_ready"], UI_TEXT["status_reordered"], "success")
        self.update_buttons()

    def move_selected_down(self) -> None:
        if self.is_busy() or self.selected_index is None or self.selected_index >= len(self.items) - 1:
            return
        index = self.selected_index
        self.items[index + 1], self.items[index] = self.items[index], self.items[index + 1]
        self.selected_index = index + 1
        self.render_image_list()
        self.set_status(UI_TEXT["status_ready"], UI_TEXT["status_reordered"], "success")
        self.update_buttons()

    def delete_selected(self) -> None:
        if self.is_busy() or self.selected_index is None:
            return
        item = self.items.pop(self.selected_index)
        self.item_keys.discard(item.path_key)
        if self.items:
            self.selected_index = min(self.selected_index, len(self.items) - 1)
        else:
            self.selected_index = None
        self.render_image_list()
        self.set_status(UI_TEXT["status_ready"], UI_TEXT["status_deleted"], "success")
        self.update_buttons()

    def clear_all(self) -> None:
        if self.is_busy() or not self.items:
            return
        self.items.clear()
        self.item_keys.clear()
        self.selected_index = None
        self.render_image_list()
        self.set_status(UI_TEXT["status_idle"], UI_TEXT["status_cleared"], "success")
        self.update_buttons()

    def save_pdf(self) -> None:
        if self.is_busy():
            return
        if not self.items:
            messagebox.showwarning(UI_TEXT["dialog_warning_title"], UI_TEXT["message_no_images"])
            self.set_status(UI_TEXT["status_error"], UI_TEXT["message_no_images"], "error")
            return
        if not REPORTLAB_READY:
            self.show_error(UI_TEXT["message_reportlab_missing"])
            return

        selected = filedialog.asksaveasfilename(
            parent=self.root,
            title=UI_TEXT["dialog_save_title"],
            initialdir=str(default_output_dir()),
            initialfile=default_output_filename(self.layout_var.get()),
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf"), (UI_TEXT["file_type_all"], "*.*")],
            confirmoverwrite=False,
        )
        if not selected:
            return

        output_path = ensure_pdf_suffix(Path(selected))
        if not output_path.parent.exists() or not output_path.parent.is_dir():
            self.show_error(UI_TEXT["message_save_folder_invalid"])
            return

        source_paths = [item.path for item in self.items]
        layout = self.layout_var.get()
        self.is_generating = True
        self.cancel_event = threading.Event()
        self.close_after_worker = False
        self.set_status(UI_TEXT["status_processing"], UI_TEXT["status_prepare_image"].format(current=0, total=len(source_paths)), "working")
        self.update_buttons()
        thread = threading.Thread(
            target=self.pdf_worker,
            args=(source_paths, output_path, layout, self.cancel_event),
            daemon=True,
        )
        thread.start()

    def pdf_worker(
        self,
        source_paths: list[Path],
        output_path: Path,
        layout: str,
        cancel_event: threading.Event,
    ) -> None:
        try:
            result_path = generate_pdf(source_paths, output_path, layout, cancel_event, self.event_queue)
            self.event_queue.put(("pdf_done", str(result_path)))
        except CancelledError:
            self.event_queue.put(("pdf_cancelled",))
        except ImageReadError as exc:
            self.event_queue.put(("pdf_error", exc.code, exc.path.name, exc))
        except MemoryError as exc:
            self.event_queue.put(("pdf_error", "message_memory", "", exc))
        except PermissionError as exc:
            self.event_queue.put(("pdf_error", "message_save_failed", "", exc))
        except OSError as exc:
            self.event_queue.put(("pdf_error", "message_save_failed", "", exc))
        except Exception as exc:
            self.event_queue.put(("pdf_error", "message_unknown_error", "", exc))

    def cancel_pdf(self) -> None:
        if not self.is_generating or self.cancel_event is None:
            return
        self.cancel_event.set()
        self.set_status(UI_TEXT["status_processing"], UI_TEXT["status_cancel_request"], "working")
        self.cancel_button.set_enabled(False)

    def finish_pdf_success(self, output_path: Path) -> None:
        self.is_generating = False
        self.cancel_event = None
        self.set_status(UI_TEXT["status_complete"], UI_TEXT["status_complete_detail"], "success")
        self.update_buttons()
        messagebox.showinfo(UI_TEXT["dialog_complete_title"], UI_TEXT["message_complete"].format(path=output_path))
        if not open_folder(output_path.parent):
            messagebox.showwarning(UI_TEXT["dialog_warning_title"], UI_TEXT["message_open_folder_failed"])
        if self.close_after_worker:
            self.root.destroy()

    def finish_pdf_cancelled(self) -> None:
        self.is_generating = False
        self.cancel_event = None
        self.set_status(UI_TEXT["status_cancelled"], UI_TEXT["message_cancelled"], "success")
        self.update_buttons()
        messagebox.showinfo(UI_TEXT["dialog_cancel_title"], UI_TEXT["message_cancelled"])
        if self.close_after_worker:
            self.root.destroy()

    def finish_pdf_error(self, message_key: str, filename: str, exc: BaseException) -> None:
        self.is_generating = False
        self.cancel_event = None
        write_error_log(f"pdf_error:{message_key}:{filename}", exc)
        message_template = UI_TEXT.get(message_key, UI_TEXT["message_unknown_error"])
        try:
            message = message_template.format(name=filename)
        except Exception:
            message = message_template
        self.set_status(UI_TEXT["status_error"], message, "error")
        self.update_buttons()
        messagebox.showerror(UI_TEXT["dialog_error_title"], message)
        if self.close_after_worker:
            self.root.destroy()

    def show_error(self, message: str) -> None:
        self.set_status(UI_TEXT["status_error"], message, "error")
        messagebox.showerror(UI_TEXT["dialog_error_title"], message)

    def set_status(self, label: str, detail: str, level: str) -> None:
        self.status_var.set(label)
        self.detail_var.set(detail)
        if level == "working":
            bg, fg = THEME["selection_bg"], THEME["accent"]
        elif level == "success":
            bg, fg = THEME["success_bg"], THEME["success"]
        elif level == "error":
            bg, fg = THEME["danger_bg"], THEME["danger"]
        else:
            bg, fg = THEME["soft"], THEME["muted"]
        self.status_badge.configure(bg=bg, fg=fg)

    def update_buttons(self) -> None:
        busy = self.is_busy()
        has_items = bool(self.items)
        has_selection = self.selected_index is not None
        self.add_button.set_enabled(not busy)
        self.up_button.set_enabled(not busy and has_selection and self.selected_index is not None and self.selected_index > 0)
        self.down_button.set_enabled(
            not busy and has_selection and self.selected_index is not None and self.selected_index < len(self.items) - 1
        )
        self.delete_button.set_enabled(not busy and has_selection)
        self.clear_button.set_enabled(not busy and has_items)
        self.execute_button.set_enabled(not busy and has_items)
        self.cancel_button.set_enabled(self.is_generating and self.cancel_event is not None and not self.cancel_event.is_set())

    def is_busy(self) -> bool:
        return self.is_loading or self.is_generating

    def on_close(self) -> None:
        if self.is_generating and self.cancel_event is not None:
            self.close_after_worker = True
            self.cancel_pdf()
            return
        self.root.destroy()

    def run(self) -> None:
        self.root.mainloop()


def parse_drop_paths(root: tk.Tk, raw_data: str) -> list[Path]:
    if not raw_data:
        return []
    try:
        items = root.tk.splitlist(raw_data)
    except tk.TclError:
        items = [raw_data]
    paths: list[Path] = []
    for item in items:
        cleaned = str(item).strip()
        if cleaned.startswith("{") and cleaned.endswith("}"):
            cleaned = cleaned[1:-1]
        if cleaned:
            paths.append(Path(cleaned))
    return paths


def run_launch_check() -> int:
    if not REPORTLAB_READY:
        return 1
    if not HEIF_READY:
        return 2
    return 0


def run_cli(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--from-shimarisu", action="store_true")
    parser.add_argument("--inputs", nargs="+")
    parser.add_argument("--output")
    parser.add_argument("--layout", choices=[LAYOUT_ONE, LAYOUT_FOUR], default=LAYOUT_ONE)
    try:
        args, unknown = parser.parse_known_args(argv)
        if unknown or not args.inputs or not args.output:
            return 1
        inputs = [normalize_path(Path(raw_path)) for raw_path in args.inputs]
        if any(not path.exists() or not path.is_file() or not is_supported_image(path) for path in inputs):
            return 1
        cancel_event = threading.Event()
        generate_pdf(inputs, normalize_path(Path(args.output)), args.layout, cancel_event, None)
        return 0
    except Exception as exc:
        write_error_log("cli", exc)
        return 1


def main(argv: list[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    if "--launch-check" in args:
        return run_launch_check()
    if "--from-shimarisu" in args:
        return run_cli(args)
    app = DakeImageBatchPdfApp()
    app.run()
    return 0


if __name__ == "__main__":
    sys.exit(main())
