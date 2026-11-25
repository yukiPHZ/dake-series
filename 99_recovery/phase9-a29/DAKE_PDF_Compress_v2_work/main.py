# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import queue
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import webbrowser
from dataclasses import dataclass
from pathlib import Path
from typing import Any

tk = None
filedialog = None
tkfont = None
messagebox = None
ttk = None
fitz = None
DND_FILES = None
TkinterDnD = None
DND_ENABLED = False
FITZ_IMPORT_ATTEMPTED = False
DND_IMPORT_ATTEMPTED = False


APP_NAME = "DakePDF圧縮"
WINDOW_TITLE = "DakePDF圧縮"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "brand_series": "シンプルそれDAKEシリーズ",
    "header_subtitle": "止まらない、迷わない、すぐ終わる。",
    "main_title": "PDFを圧縮する",
    "main_description": "PDFを追加して、画質を保ちながらしっかり軽くします。",
    "drop_title": "PDFをドロップしてください",
    "drop_subtitle": "クリックしてPDFを選ぶこともできます",
    "drop_title_selected": "PDFが追加されました",
    "drop_subtitle_selected": "このPDFを圧縮して保存できます",
    "button_select": "PDFを選ぶ",
    "button_execute": "圧縮して保存",
    "button_clear": "クリア",
    "status_idle": "PDF未選択",
    "status_ready": "圧縮できます",
    "status_processing_base": "圧縮中",
    "status_processing": "圧縮中...",
    "status_processing_dots": ["圧縮中.", "圧縮中..", "圧縮中..."],
    "status_phrase_1": "Simple",
    "status_phrase_2": "Simple, fast",
    "status_phrase_3": "Simple, fast, for real work.",
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
    "message_low_reduction": "このPDFはあまり圧縮できませんでした。すでに圧縮済み、またはPDF構造上、削減幅が小さい可能性があります。",
    "message_fallback_used": "Ghostscriptが見つからない、またはうまく処理できなかったため、内蔵の圧縮処理で保存しました。",
    "error_not_pdf": "PDFファイルを追加してください。",
    "error_multiple_files": "PDFは1つだけ追加してください。",
    "error_read_failed": "PDFを読み込めませんでした。",
    "error_encrypted": "暗号化されたPDFは処理できません。",
    "error_save_failed": "PDFを保存できませんでした。",
    "error_output_missing": "圧縮後ファイルが作成されませんでした。",
    "error_file_in_use": "ファイルが使用中の可能性があります。PDFを閉じてからもう一度お試しください。",
    "error_dependency_missing": "PDF処理に必要なライブラリが見つかりません。requirements.txt をインストールしてください。",
    "error_no_file": "先にPDFを追加してください。",
    "error_no_reduction": "このPDFは圧縮効果がありませんでした。すでに圧縮済み、またはPDF構造上、削減幅が小さい可能性があります。",
    "error_ghostscript_failed": "しっかり圧縮を実行できませんでした。内蔵の圧縮処理に切り替えます。",
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
WINDOW_SIZE = "860x740"
WINDOW_MIN_SIZE = (760, 720)
QUEUE_POLL_INTERVAL_MS = 80
LOW_REDUCTION_THRESHOLD = 5.0
FOOTER_NARROW_WIDTH = 900
STATUS_ANIMATION_INTERVAL_MS = 450
STATUS_PHRASE_DELAY_SECONDS = 1.6
GHOSTSCRIPT_PDF_SETTINGS = "/ebook"
GHOSTSCRIPT_LIGHTER_PDF_SETTINGS = "/screen"
GHOSTSCRIPT_TIMEOUT_SECONDS = 300
DEBUG_LOG_ENV = "DAKE_PDF_COMPRESS_DEBUG"

CLI_HELP_TEXT = """DakePDF_Compress CLI
Usage:
  DakePDF_Compress.exe --from-shimarisu --inputs "A.pdf" ["B.pdf" ...]

Options:
  --from-shimarisu   Run without GUI for SHIMARISU.
  --inputs           One or more PDF files.
  --help-cli         Show this help and exit.

Output:
  Saves next to each source PDF as *_compressed.pdf.
  Uses Ghostscript first when available, then built-in fallback.
  Prints output PDF path on success.
"""

CLI_ERROR_TEXT = {
    "missing_inputs": "No input PDF.",
    "not_pdf": "PDF only.",
    "not_found": "File not found.",
    "dependency": "Missing PDF library.",
    "read_failed": "Cannot read PDF.",
    "encrypted": "Encrypted PDF.",
    "save_failed": "Save failed.",
    "output_missing": "Output missing.",
    "file_in_use": "File in use.",
    "no_reduction": "No size reduction.",
    "unknown": "Compression failed.",
}


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
    engine: str
    used_fallback: bool = False
    ghostscript_path: str | None = None


def get_fitz() -> Any:
    global fitz, FITZ_IMPORT_ATTEMPTED
    if not FITZ_IMPORT_ATTEMPTED:
        try:
            import fitz as fitz_module  # PyMuPDF

            fitz = fitz_module
        except Exception:
            fitz = None
        FITZ_IMPORT_ATTEMPTED = True
    return fitz


def ensure_tkinter() -> None:
    global filedialog, messagebox, tk, tkfont, ttk
    if tk is not None:
        return
    import tkinter as tk_module
    from tkinter import filedialog as filedialog_module
    from tkinter import font as tkfont_module
    from tkinter import messagebox as messagebox_module
    from tkinter import ttk as ttk_module

    tk = tk_module
    filedialog = filedialog_module
    tkfont = tkfont_module
    messagebox = messagebox_module
    ttk = ttk_module


def ensure_dnd() -> None:
    global DND_ENABLED, DND_FILES, DND_IMPORT_ATTEMPTED, TkinterDnD
    if DND_IMPORT_ATTEMPTED:
        return

    try:
        from tkinterdnd2 import DND_FILES as dnd_files
        from tkinterdnd2 import TkinterDnD as tkinter_dnd

        DND_FILES = dnd_files
        TkinterDnD = tkinter_dnd
        DND_ENABLED = True
    except Exception:
        DND_FILES = None
        TkinterDnD = None
        DND_ENABLED = False
    DND_IMPORT_ATTEMPTED = True


def make_root() -> tk.Tk:
    ensure_tkinter()
    ensure_dnd()
    if DND_ENABLED and TkinterDnD is not None:
        return TkinterDnD.Tk()
    return tk.Tk()


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def resource_icon_path() -> Path:
    candidates: list[Path] = []
    if getattr(sys, "frozen", False):
        bundled = Path(getattr(sys, "_MEIPASS", app_dir())) / COMMON_ICON_FILENAME
        exe_path = Path(sys.executable).resolve()
        candidates.extend(
            [
                bundled,
                (exe_path.parent / COMMON_ICON_RELATIVE).resolve(),
                (exe_path.parents[3] / "02_assets" / COMMON_ICON_FILENAME).resolve()
                if len(exe_path.parents) > 3
                else bundled,
            ]
        )
    else:
        source_path = Path(__file__).resolve()
        candidates.extend(
            [
                (source_path.parents[2] / "02_assets" / COMMON_ICON_FILENAME).resolve(),
                (source_path.parent / COMMON_ICON_RELATIVE).resolve(),
            ]
        )

    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[-1]


def apply_window_icon(window: tk.Misc) -> None:
    try:
        icon_path = resource_icon_path()
        if icon_path.exists():
            window.iconbitmap(str(icon_path))
    except Exception:
        pass


def choose_font_family(root: tk.Tk) -> str:
    ensure_tkinter()
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


def debug_log(message: str) -> None:
    if os.environ.get(DEBUG_LOG_ENV) == "1":
        print(f"[DakePDF_Compress] {message}")


def make_temp_pdf_path(source_path: Path) -> Path:
    temp_handle, temp_name = tempfile.mkstemp(
        prefix=".dake_pdf_compress_",
        suffix=".pdf",
        dir=str(source_path.parent),
    )
    os.close(temp_handle)
    temp_path = Path(temp_name)
    temp_path.unlink(missing_ok=True)
    return temp_path


def find_ghostscript() -> Path | None:
    for command_name in ("gswin64c.exe", "gswin32c.exe"):
        found = shutil.which(command_name)
        if found:
            return Path(found)

    search_roots = [
        (Path("C:/Program Files/gs"), "gswin64c.exe"),
        (Path("C:/Program Files (x86)/gs"), "gswin32c.exe"),
    ]
    for root, executable_name in search_roots:
        if not root.exists():
            continue
        candidates = sorted(root.glob(f"*/bin/{executable_name}"), reverse=True)
        for candidate in candidates:
            if candidate.exists():
                return candidate
    return None


def validate_pdf(path: Path) -> None:
    pdf_lib = get_fitz()
    if pdf_lib is None:
        raise CompressError("error_dependency_missing")
    if not path.exists() or not path.is_file():
        raise CompressError("error_read_failed")
    if path.suffix.lower() != ".pdf":
        raise CompressError("error_not_pdf")

    doc = None
    try:
        doc = pdf_lib.open(str(path))
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
        doc = get_fitz().open(str(path))
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


def save_pymupdf_compressed(source_path: Path, output_path: Path) -> None:
    pdf_lib = get_fitz()
    if pdf_lib is None:
        raise CompressError("error_dependency_missing")

    doc = None
    try:
        doc = pdf_lib.open(str(source_path))
        if getattr(doc, "needs_pass", False):
            raise CompressError("error_encrypted")
        rewrite_images_if_supported(doc)
        save_optimized_pdf(doc, output_path)
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


def run_ghostscript_compression(
    source_path: Path,
    output_path: Path,
    ghostscript_path: Path,
    pdf_settings: str = GHOSTSCRIPT_PDF_SETTINGS,
) -> None:
    command = [
        str(ghostscript_path),
        "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.4",
        f"-dPDFSETTINGS={pdf_settings}",
        "-dNOPAUSE",
        "-dQUIET",
        "-dBATCH",
        f"-sOutputFile={output_path}",
        str(source_path),
    ]
    creationflags = 0
    if sys.platform.startswith("win") and hasattr(subprocess, "CREATE_NO_WINDOW"):
        creationflags = subprocess.CREATE_NO_WINDOW

    try:
        completed = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=GHOSTSCRIPT_TIMEOUT_SECONDS,
            creationflags=creationflags,
            check=False,
        )
    except PermissionError as exc:
        raise CompressError("error_file_in_use", str(exc)) from exc
    except subprocess.TimeoutExpired as exc:
        raise CompressError("error_ghostscript_failed", str(exc)) from exc
    except Exception as exc:
        raise CompressError("error_ghostscript_failed", str(exc)) from exc

    if completed.returncode != 0:
        detail = (completed.stderr or completed.stdout or "").strip()
        raise CompressError("error_ghostscript_failed", detail[:160] if detail else None)


def build_pdf_result(
    output_path: Path,
    original_size: int,
    engine: str,
    used_fallback: bool,
    ghostscript_path: Path | None,
) -> PdfResult:
    compressed_size = output_path.stat().st_size
    reduction_rate = 0.0
    if original_size > 0:
        reduction_rate = max(0.0, (1 - (compressed_size / original_size)) * 100)

    debug_log(f"engine={engine}")
    debug_log(f"ghostscript_path={ghostscript_path if ghostscript_path else 'not_found'}")
    debug_log(f"original_size={original_size}")
    debug_log(f"compressed_size={compressed_size}")
    debug_log(f"reduction_rate={reduction_rate:.1f}")
    debug_log(f"fallback={used_fallback}")

    return PdfResult(
        output_path=output_path,
        original_size=original_size,
        compressed_size=compressed_size,
        reduction_rate=reduction_rate,
        low_reduction=reduction_rate < LOW_REDUCTION_THRESHOLD,
        engine=engine,
        used_fallback=used_fallback,
        ghostscript_path=str(ghostscript_path) if ghostscript_path else None,
    )


def accept_smaller_pdf(
    temp_path: Path,
    output_path: Path,
    original_size: int,
) -> bool:
    verify_created_pdf(temp_path)
    if temp_path.stat().st_size >= original_size:
        temp_path.unlink(missing_ok=True)
        return False
    temp_path.replace(output_path)
    return True


def compress_with_fallback(
    source_path: Path,
    output_path: Path,
    original_size: int,
    ghostscript_path: Path | None,
    used_fallback: bool,
) -> PdfResult:
    temp_path = make_temp_pdf_path(source_path)
    try:
        save_pymupdf_compressed(source_path, temp_path)
        if not accept_smaller_pdf(temp_path, output_path, original_size):
            raise CompressError("error_no_reduction")
        return build_pdf_result(
            output_path=output_path,
            original_size=original_size,
            engine="pymupdf",
            used_fallback=used_fallback,
            ghostscript_path=ghostscript_path,
        )
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


def compress_pdf(source_path: Path) -> PdfResult:
    if get_fitz() is None:
        raise CompressError("error_dependency_missing")

    validate_pdf(source_path)
    original_size = source_path.stat().st_size
    output_path = unique_output_path(source_path)
    ghostscript_path = find_ghostscript()

    if ghostscript_path is not None:
        temp_path = make_temp_pdf_path(source_path)
        try:
            run_ghostscript_compression(source_path, temp_path, ghostscript_path)
            if accept_smaller_pdf(temp_path, output_path, original_size):
                return build_pdf_result(
                    output_path=output_path,
                    original_size=original_size,
                    engine="ghostscript",
                    used_fallback=False,
                    ghostscript_path=ghostscript_path,
                )
        except CompressError as exc:
            debug_log(f"ghostscript_failed={exc.message_key}")
        except Exception as exc:
            debug_log(f"ghostscript_failed={type(exc).__name__}")
        finally:
            if temp_path.exists():
                try:
                    temp_path.unlink()
                except Exception:
                    pass

    return compress_with_fallback(
        source_path=source_path,
        output_path=output_path,
        original_size=original_size,
        ghostscript_path=ghostscript_path,
        used_fallback=True,
    )


def cli_write_error(message: str) -> None:
    print(message, file=sys.stderr)


def cli_error_for_exception(exc: CompressError) -> str:
    mapping = {
        "error_not_pdf": "not_pdf",
        "error_read_failed": "read_failed",
        "error_encrypted": "encrypted",
        "error_save_failed": "save_failed",
        "error_output_missing": "output_missing",
        "error_file_in_use": "file_in_use",
        "error_dependency_missing": "dependency",
        "error_no_reduction": "no_reduction",
    }
    return CLI_ERROR_TEXT.get(mapping.get(exc.message_key, "unknown"), CLI_ERROR_TEXT["unknown"])


def collect_cli_inputs(argv: list[str]) -> list[Path]:
    inputs: list[Path] = []
    index = 0
    while index < len(argv):
        if argv[index] != "--inputs":
            index += 1
            continue

        index += 1
        while index < len(argv) and not argv[index].startswith("--"):
            value = argv[index].strip()
            if value:
                inputs.append(Path(value))
            index += 1
    return inputs


def run_cli(argv: list[str]) -> int | None:
    if "--help-cli" in argv:
        print(CLI_HELP_TEXT)
        return 0

    if "--from-shimarisu" not in argv:
        return None

    inputs = collect_cli_inputs(argv)
    if not inputs:
        cli_write_error(CLI_ERROR_TEXT["missing_inputs"])
        return 1

    output_paths: list[Path] = []
    try:
        for source_path in inputs:
            if source_path.suffix.lower() != ".pdf":
                cli_write_error(CLI_ERROR_TEXT["not_pdf"])
                return 1
            if not source_path.exists() or not source_path.is_file():
                cli_write_error(CLI_ERROR_TEXT["not_found"])
                return 1
            result = compress_pdf(source_path)
            output_paths.append(result.output_path)
    except CompressError as exc:
        cli_write_error(cli_error_for_exception(exc))
        return 1
    except Exception:
        cli_write_error(CLI_ERROR_TEXT["unknown"])
        return 1

    for output_path in output_paths:
        print(str(output_path))
    return 0


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
        self.status_animation_after_id: str | None = None
        self.status_animation_index = 0
        self.status_animation_started_at = 0.0
        self.footer_mode: str | None = None

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
        self.root.bind("<Configure>", self.handle_root_configure)

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
            text=UI_TEXT["main_title"],
            bg=COLORS["base_bg"],
            font=(self.font_family, 22, "bold"),
        ).pack(anchor=tk.W)
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
        self.update_footer_layout()

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

    def footer_thought_text(self) -> str:
        return f"{UI_TEXT['footer_left']}{UI_TEXT['footer_separator']}{UI_TEXT['header_subtitle']}"

    def clear_footer(self) -> None:
        for child in self.footer.winfo_children():
            child.destroy()

    def add_footer_text(self, parent: tk.Frame, text: str) -> tk.Label:
        label = self.make_label(
            parent,
            text=text,
            bg=COLORS["base_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 9),
        )
        label.pack(side=tk.LEFT)
        return label

    def add_footer_link(self, parent: tk.Frame, key: str) -> None:
        label = self.make_label(
            parent,
            text=UI_TEXT[key],
            bg=COLORS["base_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 9),
            cursor="hand2",
        )
        label.pack(side=tk.LEFT)
        label.bind("<Button-1>", lambda _event, url=LINK_URLS[key]: webbrowser.open(url))
        label.bind("<Enter>", lambda _event, widget=label: widget.configure(fg=COLORS["accent"]))
        label.bind("<Leave>", lambda _event, widget=label: widget.configure(fg=COLORS["muted"]))

    def add_footer_link_line(self, parent: tk.Frame) -> None:
        self.add_footer_link(parent, "footer_link_1")
        self.add_footer_text(parent, UI_TEXT["footer_separator"])
        self.add_footer_link(parent, "footer_link_2")
        self.add_footer_text(parent, UI_TEXT["footer_separator"])
        self.add_footer_text(parent, UI_TEXT["footer_copyright"])

    def update_footer_layout(self, width: int | None = None) -> None:
        if width is None:
            width = self.root.winfo_width()
        mode = "narrow" if width < FOOTER_NARROW_WIDTH else "wide"
        if mode == self.footer_mode:
            return

        self.footer_mode = mode
        self.clear_footer()

        if mode == "wide":
            left = tk.Frame(self.footer, bg=COLORS["base_bg"])
            left.pack(side=tk.LEFT)
            self.add_footer_text(left, self.footer_thought_text())

            right = tk.Frame(self.footer, bg=COLORS["base_bg"])
            right.pack(side=tk.RIGHT)
            self.add_footer_link_line(right)
            return

        thought_line = tk.Frame(self.footer, bg=COLORS["base_bg"])
        thought_line.pack(anchor=tk.CENTER)
        self.add_footer_text(thought_line, self.footer_thought_text())

        link_line = tk.Frame(self.footer, bg=COLORS["base_bg"])
        link_line.pack(anchor=tk.CENTER, pady=(4, 0))
        self.add_footer_link_line(link_line)

    def handle_root_configure(self, event: tk.Event) -> None:
        if event.widget == self.root:
            self.update_footer_layout(event.width)

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
        self.start_status_animation()
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
        self.stop_status_animation()
        self.compressed_size_var.set(format_bytes(result.compressed_size))
        self.reduction_rate_var.set(f"{result.reduction_rate:.1f}%")
        self.save_name_var.set(truncate_middle(result.output_path.name, 58))
        self.save_folder_var.set(truncate_middle(str(result.output_path.parent), 70))

        notices = []
        if result.used_fallback:
            notices.append(UI_TEXT["message_fallback_used"])
        if result.low_reduction:
            notices.append(UI_TEXT["message_low_reduction"])
        self.notice_var.set("\n".join(notices))

        if result.low_reduction:
            self.set_status("status_low_reduction", "warning")
            message = f"{UI_TEXT['message_complete']}\n\n{self.notice_var.get()}\n\n{UI_TEXT['message_complete_detail']}"
            dialog = messagebox.showwarning
        else:
            self.set_status("status_complete", "success")
            if self.notice_var.get():
                message = f"{UI_TEXT['message_complete']}\n\n{self.notice_var.get()}\n\n{UI_TEXT['message_complete_detail']}"
            else:
                message = f"{UI_TEXT['message_complete']}\n\n{UI_TEXT['message_complete_detail']}"
            dialog = messagebox.showinfo

        self.update_action_state()
        dialog(UI_TEXT["dialog_complete_title"], message)
        self.open_output_folder(result.output_path.parent)

    def handle_worker_error(self, exc: CompressError) -> None:
        self.is_processing = False
        self.progress.stop()
        self.stop_status_animation()
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

    def start_status_animation(self) -> None:
        self.stop_status_animation()
        self.status_animation_started_at = time.monotonic()
        self.status_animation_index = 0
        self.animate_processing_status()

    def stop_status_animation(self) -> None:
        if self.status_animation_after_id is not None:
            try:
                self.root.after_cancel(self.status_animation_after_id)
            except Exception:
                pass
            self.status_animation_after_id = None

    def animate_processing_status(self) -> None:
        if not self.is_processing:
            return

        elapsed = time.monotonic() - self.status_animation_started_at
        sequence = list(UI_TEXT["status_processing_dots"])
        if elapsed >= STATUS_PHRASE_DELAY_SECONDS:
            sequence.extend(
                [
                    UI_TEXT["status_phrase_1"],
                    UI_TEXT["status_phrase_2"],
                    UI_TEXT["status_phrase_3"],
                ]
            )
        self.status_var.set(sequence[self.status_animation_index % len(sequence)])
        self.status_animation_index += 1
        self.status_animation_after_id = self.root.after(
            STATUS_ANIMATION_INTERVAL_MS,
            self.animate_processing_status,
        )

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
    cli_exit_code = run_cli(sys.argv[1:])
    if cli_exit_code is not None:
        raise SystemExit(cli_exit_code)

    root = make_root()
    DakePdfCompressApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
