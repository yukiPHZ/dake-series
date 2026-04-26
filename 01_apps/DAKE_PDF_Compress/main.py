# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import queue
import shutil
import sys
import tempfile
import threading
import webbrowser
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, ttk

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    DND_ENABLED = True
except Exception:
    DND_FILES = None
    TkinterDnD = None
    DND_ENABLED = False


APP_NAME = "DakePDF圧縮"
WINDOW_TITLE = "DakePDF圧縮"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "brand_series": "シンプルそれDAKEシリーズ",
    "main_title": "PDFを圧縮する",
    "main_description": "PDFを追加して、ファイルサイズを軽くします。",
    "drop_title": "PDFをドロップしてください",
    "drop_subtitle": "クリックしてPDFを選ぶこともできます",
    "drop_title_selected": "PDFが追加されました",
    "drop_subtitle_selected": "このPDFを圧縮して保存できます",
    "button_select": "PDFを選ぶ",
    "button_execute": "圧縮して保存",
    "button_clear": "クリア",
    "status_idle": "PDF未選択",
    "status_ready": "圧縮できます",
    "status_processing": "圧縮中...",
    "status_complete": "圧縮が完了しました",
    "status_error": "エラー",
    "status_low_reduction": "圧縮効果は小さめです",
    "label_file_name": "ファイル名",
    "label_original_size": "元サイズ",
    "label_save_name": "保存予定ファイル名",
    "label_save_folder": "保存先",
    "label_compressed_size": "圧縮後サイズ",
    "label_reduction_rate": "削減率",
    "value_empty": "未選択",
    "value_not_yet": "未処理",
    "dialog_select_title": "PDFを選択",
    "dialog_complete_title": "圧縮完了",
    "dialog_error_title": "確認してください",
    "dialog_filetype_pdf": "PDFファイル",
    "dialog_filetype_all": "すべてのファイル",
    "message_complete": "PDFの圧縮が完了しました。",
    "message_complete_detail": "保存先フォルダを開きます。",
    "message_low_reduction": "このPDFはあまり圧縮できませんでした。元のPDF構造上、削減幅が小さい可能性があります。",
    "error_not_pdf": "PDFファイルを追加してください。",
    "error_multiple_files": "PDFは1つだけ追加してください。",
    "error_read_failed": "PDFを読み込めませんでした。",
    "error_encrypted": "暗号化されたPDFは処理できません。",
    "error_save_failed": "PDFを保存できませんでした。",
    "error_output_missing": "圧縮後ファイルが作成されませんでした。",
    "error_file_in_use": "ファイルが使用中の可能性があります。PDFを閉じてからもう一度お試しください。",
    "error_dependency_missing": "PDF処理に必要なライブラリが見つかりません。requirements.txt をインストールしてください。",
    "error_no_file": "先にPDFを追加してください。",
    "error_unknown": "処理中に問題が発生しました。",
    "detail_suffix": "詳細: {detail}",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_link_1": "戸建買取査定",
    "footer_link_2": "Instagram",
    "footer_separator": " ｜ ",
    "footer_copyright": COPYRIGHT,
}

COLORS = {
    "base_bg": "#F6F7F9",
    "card_bg": "#FFFFFF",
    "text": "#1E2430",
    "muted": "#667085",
    "border": "#E6EAF0",
    "accent": "#2F6FED",
    "accent_hover": "#2458BF",
    "selection_bg": "#EAF2FF",
    "success": "#12B76A",
    "success_bg": "#E8FFF3",
    "error": "#B42318",
    "error_bg": "#FEE4E2",
    "warning": "#B54708",
    "warning_bg": "#FFFAEB",
    "disabled": "#D8DEE8",
    "white": "#FFFFFF",
}

LINK_URLS = {
    "footer_link_1": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "footer_link_2": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

FONT_CANDIDATES = ["BIZ UDPGothic", "Yu Gothic UI", "Meiryo"]
COMMON_ICON_RELATIVE = Path("..") / ".." / "02_assets" / "dake_icon.ico"
COMMON_ICON_FILENAME = "dake_icon.ico"
WINDOW_SIZE = "860x620"
WINDOW_MIN_SIZE = (780, 560)
QUEUE_POLL_INTERVAL_MS = 80
LOW_REDUCTION_THRESHOLD = 1.0


class CompressError(Exception):
    def __init__(self, message_key: str, detail: str | None = None) -> None:
        super().__init__(detail or message_key)
        self.message_key = message_key
        self.detail = detail


@dataclass
class PdfResult:
    output_path: Path
    original_size: int
    compressed_size: int
    reduction_rate: float
    low_reduction: bool


def make_root() -> tk.Tk:
    if DND_ENABLED and TkinterDnD is not None:
        return TkinterDnD.Tk()
    return tk.Tk()


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def resource_icon_path() -> Path:
    if getattr(sys, "frozen", False):
        bundled = Path(getattr(sys, "_MEIPASS", app_dir())) / COMMON_ICON_FILENAME
        if bundled.exists():
            return bundled
        return (Path(sys.executable).resolve().parent / COMMON_ICON_RELATIVE).resolve()
    return (Path(__file__).resolve().parent / COMMON_ICON_RELATIVE).resolve()


def apply_window_icon(window: tk.Misc) -> None:
    try:
        icon_path = resource_icon_path()
        if icon_path.exists():
            window.iconbitmap(str(icon_path))
    except Exception:
        pass


def choose_font_family(root: tk.Tk) -> str:
    available = set(tkfont.families(root))
    for family in FONT_CANDIDATES:
        if family in available:
            return family
    return "TkDefaultFont"


def format_bytes(size: int) -> str:
    if size >= 1024 * 1024:
        return f"{size / (1024 * 1024):.1f} MB"
    if size >= 1024:
        return f"{size / 1024:.0f} KB"
    return f"{size} B"


def truncate_middle(text: str, max_chars: int = 64) -> str:
    if len(text) <= max_chars:
        return text
    front = max_chars // 2
    back = max_chars - front - 3
    return f"{text[:front]}...{text[-back:]}"


def unique_output_path(source_path: Path) -> Path:
    base = source_path.with_name(f"{source_path.stem}_compressed.pdf")
    if not base.exists() and base.resolve() != source_path.resolve():
        return base

    counter = 2
    while True:
        candidate = source_path.with_name(f"{source_path.stem}_compressed_{counter}.pdf")
        if not candidate.exists() and candidate.resolve() != source_path.resolve():
            return candidate
        counter += 1


def validate_pdf(path: Path) -> None:
    if fitz is None:
        raise CompressError("error_dependency_missing")
    if not path.exists() or not path.is_file():
        raise CompressError("error_read_failed")
    if path.suffix.lower() != ".pdf":
        raise CompressError("error_not_pdf")

    doc = None
    try:
        doc = fitz.open(str(path))
        if getattr(doc, "needs_pass", False):
            raise CompressError("error_encrypted")
        if doc.page_count < 1:
            raise CompressError("error_read_failed")
    except CompressError:
        raise
    except PermissionError as exc:
        raise CompressError("error_file_in_use", str(exc)) from exc
    except Exception as exc:
        raise CompressError("error_read_failed", str(exc)) from exc
    finally:
        if doc is not None:
            try:
                doc.close()
            except Exception:
                pass


def rewrite_images_if_supported(doc: Any) -> None:
    rewrite_images = getattr(doc, "rewrite_images", None)
    if rewrite_images is None:
        return

    try:
        rewrite_images(
            dpi_threshold=220,
            dpi_target=150,
            quality=82,
            lossy=True,
            lossless=True,
            bitonal=False,
            color=True,
            gray=True,
        )
    except TypeError:
        try:
            rewrite_images(dpi_threshold=220, dpi_target=150, quality=82)
        except Exception:
            pass
    except Exception:
        pass


def save_optimized_pdf(doc: Any, output_path: Path) -> None:
    try:
        doc.ez_save(
            str(output_path),
            garbage=4,
            clean=True,
            deflate=True,
            deflate_images=True,
            deflate_fonts=True,
        )
    except AttributeError:
        doc.save(str(output_path), garbage=4, clean=True, deflate=True)
    except TypeError:
        doc.save(str(output_path), garbage=4, clean=True, deflate=True)


def verify_created_pdf(path: Path) -> None:
    if not path.exists() or path.stat().st_size <= 0:
        raise CompressError("error_output_missing")

    doc = None
    try:
        doc = fitz.open(str(path))
        if doc.page_count < 1:
            raise CompressError("error_output_missing")
    except CompressError:
        raise
    except Exception as exc:
        raise CompressError("error_output_missing", str(exc)) from exc
    finally:
        if doc is not None:
            try:
                doc.close()
            except Exception:
                pass


def compress_pdf(source_path: Path) -> PdfResult:
    if fitz is None:
        raise CompressError("error_dependency_missing")

    validate_pdf(source_path)
    original_size = source_path.stat().st_size
    output_path = unique_output_path(source_path)
    temp_handle, temp_name = tempfile.mkstemp(
        prefix=".dake_pdf_compress_",
        suffix=".pdf",
        dir=str(source_path.parent),
    )
    os.close(temp_handle)
    temp_path = Path(temp_name)
    temp_path.unlink(missing_ok=True)

    doc = None
    try:
        doc = fitz.open(str(source_path))
        if getattr(doc, "needs_pass", False):
            raise CompressError("error_encrypted")
        rewrite_images_if_supported(doc)
        save_optimized_pdf(doc, temp_path)
    except CompressError:
        raise
    except PermissionError as exc:
        raise CompressError("error_file_in_use", str(exc)) from exc
    except Exception as exc:
        raise CompressError("error_save_failed", str(exc)) from exc
    finally:
        if doc is not None:
            try:
                doc.close()
            except Exception:
                pass

    try:
        verify_created_pdf(temp_path)
        compressed_size = temp_path.stat().st_size
        if compressed_size >= original_size:
            temp_path.unlink(missing_ok=True)
            shutil.copy2(source_path, output_path)
        else:
            temp_path.replace(output_path)
    except PermissionError as exc:
        raise CompressError("error_file_in_use", str(exc)) from exc
    except CompressError:
        raise
    except Exception as exc:
        raise CompressError("error_save_failed", str(exc)) from exc
    finally:
        if temp_path.exists():
            try:
                temp_path.unlink()
            except Exception:
                pass

    if not output_path.exists():
        raise CompressError("error_output_missing")

    compressed_size = output_path.stat().st_size
    if original_size <= 0:
        reduction_rate = 0.0
    else:
        reduction_rate = max(0.0, (1 - (compressed_size / original_size)) * 100)
    low_reduction = reduction_rate < LOW_REDUCTION_THRESHOLD

    return PdfResult(
        output_path=output_path,
        original_size=original_size,
        compressed_size=compressed_size,
        reduction_rate=reduction_rate,
        low_reduction=low_reduction,
    )


class DakePdfCompressApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(*WINDOW_MIN_SIZE)
        self.root.configure(bg=COLORS["base_bg"])
        apply_window_icon(self.root)

        self.font_family = choose_font_family(root)
        self.selected_pdf: Path | None = None
        self.is_processing = False
        self.event_queue: queue.Queue[tuple[str, Any]] = queue.Queue()

        self.status_var = tk.StringVar(value=UI_TEXT["status_idle"])
        self.drop_title_var = tk.StringVar(value=UI_TEXT["drop_title"])
        self.drop_subtitle_var = tk.StringVar(value=UI_TEXT["drop_subtitle"])
        self.file_name_var = tk.StringVar(value=UI_TEXT["value_empty"])
        self.original_size_var = tk.StringVar(value=UI_TEXT["value_empty"])
        self.save_name_var = tk.StringVar(value=UI_TEXT["value_empty"])
        self.save_folder_var = tk.StringVar(value=UI_TEXT["value_empty"])
        self.compressed_size_var = tk.StringVar(value=UI_TEXT["value_not_yet"])
        self.reduction_rate_var = tk.StringVar(value=UI_TEXT["value_not_yet"])
        self.notice_var = tk.StringVar(value="")

        self.setup_style()
        self.build_ui()
        self.setup_drop_targets()
        self.root.after(QUEUE_POLL_INTERVAL_MS, self.poll_queue)

    def setup_style(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(
            "Dake.Horizontal.TProgressbar",
            background=COLORS["accent"],
            troughcolor=COLORS["border"],
            bordercolor=COLORS["border"],
            lightcolor=COLORS["accent"],
            darkcolor=COLORS["accent"],
        )

    def make_label(self, parent: tk.Misc, **kwargs: Any) -> tk.Label:
        options = {
            "bg": kwargs.pop("bg", COLORS["card_bg"]),
            "fg": kwargs.pop("fg", COLORS["text"]),
            "font": kwargs.pop("font", (self.font_family, 10)),
        }
        options.update(kwargs)
        return tk.Label(parent, **options)

    def build_ui(self) -> None:
        self.container = tk.Frame(self.root, bg=COLORS["base_bg"])
        self.container.pack(fill=tk.BOTH, expand=True, padx=28, pady=22)

        header = tk.Frame(self.container, bg=COLORS["base_bg"])
        header.pack(fill=tk.X)
        self.make_label(
            header,
            text=UI_TEXT["brand_series"],
            bg=COLORS["base_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 9),
        ).pack(anchor=tk.W)
        self.make_label(
            header,
            text=UI_TEXT["main_title"],
            bg=COLORS["base_bg"],
            font=(self.font_family, 24, "bold"),
        ).pack(anchor=tk.W, pady=(6, 0))
        self.make_label(
            header,
            text=UI_TEXT["main_description"],
            bg=COLORS["base_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 11),
        ).pack(anchor=tk.W, pady=(4, 0))

        self.card = tk.Frame(
            self.container,
            bg=COLORS["card_bg"],
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            highlightcolor=COLORS["border"],
        )
        self.card.pack(fill=tk.BOTH, expand=True, pady=(18, 14))

        self.drop_area = tk.Frame(
            self.card,
            bg=COLORS["selection_bg"],
            highlightthickness=1,
            highlightbackground=COLORS["accent"],
            highlightcolor=COLORS["accent"],
            cursor="hand2",
        )
        self.drop_area.pack(fill=tk.X, padx=22, pady=(22, 18), ipady=24)
        self.drop_area.bind("<Button-1>", self.select_pdf_dialog)

        self.make_label(
            self.drop_area,
            textvariable=self.drop_title_var,
            bg=COLORS["selection_bg"],
            fg=COLORS["text"],
            font=(self.font_family, 16, "bold"),
            cursor="hand2",
        ).pack()
        self.make_label(
            self.drop_area,
            textvariable=self.drop_subtitle_var,
            bg=COLORS["selection_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 10),
            cursor="hand2",
        ).pack(pady=(8, 0))

        info = tk.Frame(self.card, bg=COLORS["card_bg"])
        info.pack(fill=tk.X, padx=22)
        for column in range(2):
            info.grid_columnconfigure(column, weight=1, uniform="info")

        self.add_info_row(info, 0, 0, UI_TEXT["label_file_name"], self.file_name_var)
        self.add_info_row(info, 0, 1, UI_TEXT["label_original_size"], self.original_size_var)
        self.add_info_row(info, 1, 0, UI_TEXT["label_save_name"], self.save_name_var)
        self.add_info_row(info, 1, 1, UI_TEXT["label_save_folder"], self.save_folder_var)
        self.add_info_row(info, 2, 0, UI_TEXT["label_compressed_size"], self.compressed_size_var)
        self.add_info_row(info, 2, 1, UI_TEXT["label_reduction_rate"], self.reduction_rate_var)

        self.notice_label = self.make_label(
            self.card,
            textvariable=self.notice_var,
            fg=COLORS["warning"],
            font=(self.font_family, 10),
            wraplength=720,
            justify=tk.LEFT,
        )
        self.notice_label.pack(anchor=tk.W, fill=tk.X, padx=24, pady=(14, 0))

        action_row = tk.Frame(self.card, bg=COLORS["card_bg"])
        action_row.pack(fill=tk.X, padx=22, pady=(20, 18))

        self.select_button = tk.Button(
            action_row,
            text=UI_TEXT["button_select"],
            command=self.select_pdf_dialog,
            bg=COLORS["white"],
            fg=COLORS["text"],
            activebackground=COLORS["selection_bg"],
            activeforeground=COLORS["text"],
            relief=tk.FLAT,
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            font=(self.font_family, 10, "bold"),
            padx=18,
            pady=10,
            cursor="hand2",
        )
        self.select_button.pack(side=tk.LEFT)

        self.clear_button = tk.Button(
            action_row,
            text=UI_TEXT["button_clear"],
            command=self.clear_selection,
            bg=COLORS["white"],
            fg=COLORS["muted"],
            activebackground=COLORS["selection_bg"],
            relief=tk.FLAT,
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            font=(self.font_family, 10, "bold"),
            padx=18,
            pady=10,
            cursor="hand2",
        )
        self.clear_button.pack(side=tk.LEFT, padx=(10, 0))

        self.execute_button = tk.Button(
            action_row,
            text=UI_TEXT["button_execute"],
            command=self.start_compression,
            bg=COLORS["accent"],
            fg=COLORS["white"],
            activebackground=COLORS["accent_hover"],
            activeforeground=COLORS["white"],
            disabledforeground=COLORS["white"],
            relief=tk.FLAT,
            font=(self.font_family, 11, "bold"),
            padx=26,
            pady=11,
            cursor="hand2",
        )
        self.execute_button.pack(side=tk.RIGHT)

        status_row = tk.Frame(self.card, bg=COLORS["card_bg"])
        status_row.pack(fill=tk.X, padx=22, pady=(0, 22))
        self.status_badge = self.make_label(
            status_row,
            textvariable=self.status_var,
            bg=COLORS["base_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 10, "bold"),
            padx=12,
            pady=7,
        )
        self.status_badge.pack(side=tk.LEFT)
        self.progress = ttk.Progressbar(
            status_row,
            mode="indeterminate",
            style="Dake.Horizontal.TProgressbar",
            length=180,
        )
        self.progress.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(16, 0))

        self.footer = tk.Frame(self.container, bg=COLORS["base_bg"])
        self.footer.pack(fill=tk.X)
        self.make_label(
            self.footer,
            text=UI_TEXT["footer_left"],
            bg=COLORS["base_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 9),
        ).pack(side=tk.LEFT)

        footer_right = tk.Frame(self.footer, bg=COLORS["base_bg"])
        footer_right.pack(side=tk.RIGHT)
        self.add_footer_link(footer_right, "footer_link_1")
        self.make_label(
            footer_right,
            text=UI_TEXT["footer_separator"],
            bg=COLORS["base_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 9),
        ).pack(side=tk.LEFT)
        self.add_footer_link(footer_right, "footer_link_2")
        self.make_label(
            footer_right,
            text=UI_TEXT["footer_separator"],
            bg=COLORS["base_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 9),
        ).pack(side=tk.LEFT)
        self.make_label(
            footer_right,
            text=UI_TEXT["footer_copyright"],
            bg=COLORS["base_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 9),
        ).pack(side=tk.LEFT)

        self.update_action_state()

    def add_info_row(
        self,
        parent: tk.Frame,
        row: int,
        column: int,
        label_text: str,
        value_var: tk.StringVar,
    ) -> None:
        frame = tk.Frame(parent, bg=COLORS["card_bg"])
        frame.grid(row=row, column=column, sticky="ew", padx=(0, 20), pady=8)
        self.make_label(
            frame,
            text=label_text,
            fg=COLORS["muted"],
            font=(self.font_family, 9),
        ).pack(anchor=tk.W)
        self.make_label(
            frame,
            textvariable=value_var,
            fg=COLORS["text"],
            font=(self.font_family, 11, "bold"),
            wraplength=330,
            justify=tk.LEFT,
        ).pack(anchor=tk.W, pady=(4, 0))

    def add_footer_link(self, parent: tk.Frame, key: str) -> None:
        label = self.make_label(
            parent,
            text=UI_TEXT[key],
            bg=COLORS["base_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 9, "underline"),
            cursor="hand2",
        )
        label.pack(side=tk.LEFT)
        label.bind("<Button-1>", lambda _event, url=LINK_URLS[key]: webbrowser.open(url))

    def setup_drop_targets(self) -> None:
        if not DND_ENABLED or DND_FILES is None:
            return
        widgets = [
            self.root,
            self.container,
            self.card,
            self.drop_area,
        ]
        for widget in widgets:
            try:
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", self.handle_drop)
            except Exception:
                pass

    def select_pdf_dialog(self, _event: tk.Event | None = None) -> None:
        if self.is_processing:
            return
        selected = filedialog.askopenfilename(
            title=UI_TEXT["dialog_select_title"],
            filetypes=[
                (UI_TEXT["dialog_filetype_pdf"], "*.pdf"),
                (UI_TEXT["dialog_filetype_all"], "*.*"),
            ],
        )
        if selected:
            self.load_pdf(Path(selected))

    def handle_drop(self, event: tk.Event) -> None:
        if self.is_processing:
            return
        try:
            paths = [Path(value) for value in self.root.tk.splitlist(event.data)]  # type: ignore[attr-defined]
        except Exception:
            self.show_error("error_read_failed")
            return

        files = [path for path in paths if path.is_file()]
        if len(files) != 1:
            self.show_error("error_multiple_files")
            return
        self.load_pdf(files[0])

    def load_pdf(self, path: Path) -> None:
        try:
            validate_pdf(path)
        except CompressError as exc:
            self.show_error(exc.message_key, exc.detail)
            return

        output_path = unique_output_path(path)
        self.selected_pdf = path
        self.file_name_var.set(truncate_middle(path.name, 58))
        self.original_size_var.set(format_bytes(path.stat().st_size))
        self.save_name_var.set(truncate_middle(output_path.name, 58))
        self.save_folder_var.set(truncate_middle(str(path.parent), 70))
        self.compressed_size_var.set(UI_TEXT["value_not_yet"])
        self.reduction_rate_var.set(UI_TEXT["value_not_yet"])
        self.notice_var.set("")
        self.drop_title_var.set(UI_TEXT["drop_title_selected"])
        self.drop_subtitle_var.set(UI_TEXT["drop_subtitle_selected"])
        self.set_status("status_ready", "ready")
        self.update_action_state()

    def clear_selection(self) -> None:
        if self.is_processing:
            return
        self.selected_pdf = None
        self.file_name_var.set(UI_TEXT["value_empty"])
        self.original_size_var.set(UI_TEXT["value_empty"])
        self.save_name_var.set(UI_TEXT["value_empty"])
        self.save_folder_var.set(UI_TEXT["value_empty"])
        self.compressed_size_var.set(UI_TEXT["value_not_yet"])
        self.reduction_rate_var.set(UI_TEXT["value_not_yet"])
        self.notice_var.set("")
        self.drop_title_var.set(UI_TEXT["drop_title"])
        self.drop_subtitle_var.set(UI_TEXT["drop_subtitle"])
        self.set_status("status_idle", "idle")
        self.update_action_state()

    def start_compression(self) -> None:
        if self.is_processing:
            return
        if self.selected_pdf is None:
            self.show_error("error_no_file")
            return

        source_path = self.selected_pdf
        self.is_processing = True
        self.notice_var.set("")
        self.set_status("status_processing", "processing")
        self.update_action_state()
        self.progress.start(10)

        worker = threading.Thread(target=self.compress_worker, args=(source_path,), daemon=True)
        worker.start()

    def compress_worker(self, source_path: Path) -> None:
        try:
            result = compress_pdf(source_path)
            self.event_queue.put(("success", result))
        except CompressError as exc:
            self.event_queue.put(("error", exc))
        except Exception as exc:
            self.event_queue.put(("error", CompressError("error_unknown", str(exc))))

    def poll_queue(self) -> None:
        try:
            while True:
                event = self.event_queue.get_nowait()
                self.handle_queue_event(event)
        except queue.Empty:
            pass
        self.root.after(QUEUE_POLL_INTERVAL_MS, self.poll_queue)

    def handle_queue_event(self, event: tuple[str, Any]) -> None:
        event_type, payload = event
        if event_type == "success":
            self.handle_success(payload)
        elif event_type == "error":
            self.handle_worker_error(payload)

    def handle_success(self, result: PdfResult) -> None:
        self.is_processing = False
        self.progress.stop()
        self.compressed_size_var.set(format_bytes(result.compressed_size))
        self.reduction_rate_var.set(f"{result.reduction_rate:.1f}%")
        self.save_name_var.set(truncate_middle(result.output_path.name, 58))
        self.save_folder_var.set(truncate_middle(str(result.output_path.parent), 70))

        if result.low_reduction:
            self.notice_var.set(UI_TEXT["message_low_reduction"])
            self.set_status("status_low_reduction", "warning")
            message = f"{UI_TEXT['message_complete']}\n\n{UI_TEXT['message_low_reduction']}\n\n{UI_TEXT['message_complete_detail']}"
        else:
            self.notice_var.set("")
            self.set_status("status_complete", "success")
            message = f"{UI_TEXT['message_complete']}\n\n{UI_TEXT['message_complete_detail']}"

        self.update_action_state()
        messagebox.showinfo(UI_TEXT["dialog_complete_title"], message)
        self.open_output_folder(result.output_path.parent)

    def handle_worker_error(self, exc: CompressError) -> None:
        self.is_processing = False
        self.progress.stop()
        self.set_status("status_error", "error")
        self.update_action_state()
        self.show_error(exc.message_key, exc.detail)

    def open_output_folder(self, folder: Path) -> None:
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(folder))  # type: ignore[attr-defined]
            else:
                webbrowser.open(folder.as_uri())
        except Exception:
            pass

    def show_error(self, message_key: str, detail: str | None = None) -> None:
        message = UI_TEXT.get(message_key, UI_TEXT["error_unknown"])
        if detail:
            message = f"{message}\n\n{UI_TEXT['detail_suffix'].format(detail=detail)}"
        self.notice_var.set(message)
        self.set_status("status_error", "error")
        messagebox.showwarning(UI_TEXT["dialog_error_title"], message)

    def set_status(self, key: str, state: str) -> None:
        self.status_var.set(UI_TEXT[key])
        palette = {
            "idle": (COLORS["base_bg"], COLORS["muted"]),
            "ready": (COLORS["selection_bg"], COLORS["accent"]),
            "processing": (COLORS["selection_bg"], COLORS["accent"]),
            "success": (COLORS["success_bg"], COLORS["success"]),
            "warning": (COLORS["warning_bg"], COLORS["warning"]),
            "error": (COLORS["error_bg"], COLORS["error"]),
        }
        bg, fg = palette.get(state, palette["idle"])
        self.status_badge.configure(bg=bg, fg=fg)

    def update_action_state(self) -> None:
        has_pdf = self.selected_pdf is not None
        if self.is_processing:
            self.execute_button.configure(state=tk.DISABLED, bg=COLORS["disabled"], cursor="arrow")
            self.select_button.configure(state=tk.DISABLED, cursor="arrow")
            self.clear_button.configure(state=tk.DISABLED, cursor="arrow")
            return

        self.select_button.configure(state=tk.NORMAL, cursor="hand2")
        self.clear_button.configure(state=tk.NORMAL if has_pdf else tk.DISABLED, cursor="hand2" if has_pdf else "arrow")
        self.execute_button.configure(
            state=tk.NORMAL if has_pdf else tk.DISABLED,
            bg=COLORS["accent"] if has_pdf else COLORS["disabled"],
            cursor="hand2" if has_pdf else "arrow",
        )


def main() -> None:
    root = make_root()
    DakePdfCompressApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
