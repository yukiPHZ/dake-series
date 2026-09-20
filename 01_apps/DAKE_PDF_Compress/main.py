# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import multiprocessing as mp
import shutil
import subprocess
import sys
import time
import webbrowser
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

CORE_DIR = Path(__file__).resolve().parents[2] / "00_core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

from dake_quality_engine import run_launch_check, safe_load_json_config, safe_run
from dake_quality_engine.logging import write_debug_log

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
    "status_checking": "PDFを確認中...",
    "status_analyzing": "PDFを解析中",
    "status_images": "画像を最適化中",
    "status_optimizing": "PDFを整えています",
    "status_verifying": "仕上がりを確認中",
    "status_saving": "保存中",
    "status_closing": "処理を終了しています",
    "result_sizes": "{before} → {after}",
    "result_reduction": "{rate:.1f}%軽くなりました",
    "status_idle": "PDF未選択",
    "status_ready": "圧縮できます",
    "status_processing": "圧縮中...",
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
    "message_complete_detail": "圧縮したPDFを同じフォルダに保存しました。",
    "message_low_reduction": "このPDFは、すでにかなり軽いようです。",
    "error_not_pdf": "PDFファイルを追加してください。",
    "error_multiple_files": "PDFは1つだけ追加してください。",
    "error_read_failed": "PDFを読み込めませんでした。",
    "error_encrypted": "暗号化されたPDFは処理できません。",
    "error_save_failed": "PDFを保存できませんでした。",
    "error_output_missing": "圧縮後ファイルが作成されませんでした。",
    "error_file_in_use": "ファイルが使用中の可能性があります。PDFを閉じてからもう一度お試しください。",
    "error_dependency_missing": "PDF処理の準備ができません。アプリを入れ直してください。",
    "error_no_file": "先にPDFを追加してください。",
    "error_no_reduction": "このPDFは、すでにかなり軽いようです。元PDFをご利用ください。",
    "error_unknown": "処理中に問題が発生しました。",
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
STATUS_PHRASE_DELAY_SECONDS = 8.0
DEBUG_LOG_ENV = "DAKE_PDF_COMPRESS_DEBUG"
CONFIG_FILENAME = "dake_pdf_compress_config.json"

CLI_HELP_TEXT = """DakePDF_Compress CLI
Usage:
  DakePDF_Compress.exe --from-shimarisu --inputs "A.pdf" ["B.pdf" ...]

Options:
  --from-shimarisu   Run without GUI for SHIMARISU.
  --inputs           One or more PDF files.
  --help-cli         Show this help and exit.

Output:
  Saves next to each source PDF as *_compressed.pdf.
  Automatically selects a verified compression result.
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
    def __init__(self, message_key: str, detail: str | None = None,
                 context: dict[str, object] | None = None) -> None:
        super().__init__(detail or message_key)
        self.message_key = message_key
        self.detail = detail
        self.context = context or {}


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
    diagnostics: dict = field(default_factory=dict)


def get_fitz() -> Any:
    global fitz, FITZ_IMPORT_ATTEMPTED
    if not FITZ_IMPORT_ATTEMPTED:
        try:
            import pymupdf as fitz_module

            fitz_module.TOOLS.mupdf_display_errors(False)
            fitz_module.TOOLS.mupdf_display_warnings(False)
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


def app_log_dir() -> Path:
    return app_dir() / "logs"


def app_config_path() -> Path:
    return app_dir() / CONFIG_FILENAME


def load_optional_config() -> dict[str, Any]:
    return safe_load_json_config(app_config_path(), default={})


def write_app_debug_log(
    message: str,
    *,
    exc: BaseException | None = None,
    context: dict[str, object] | None = None,
) -> None:
    path = write_debug_log(message, log_dir=app_log_dir(), exc=exc, context=context)
    if path is not None:
        return
    # Keep app diagnostics available even if the shared logger cannot write.
    try:
        from datetime import datetime
        import traceback
        log_path = app_log_dir() / f"{datetime.now():%Y-%m-%d}.log"
        log_path.parent.mkdir(parents=True, exist_ok=True)
        lines = [f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {message}"]
        for key, value in (context or {}).items():
            lines.append(f"{key}={value}")
        if exc is not None:
            lines.append("".join(traceback.format_exception(exc)))
        with log_path.open("a", encoding="utf-8") as file:
            file.write("\n".join(lines) + "\n")
    except Exception:
        pass


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
    except Exception as exc:
        write_app_debug_log("window icon setup skipped", exc=exc)
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
        write_app_debug_log(message)


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
        if not doc.is_pdf or doc.page_count < 1:
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


def compress_pdf(source_path: Path, phase=lambda key: None) -> PdfResult:
    from adaptive import compress, library
    try:
        library()
    except (ImportError, RuntimeError) as exc:
        raise CompressError("error_dependency_missing", str(exc)) from exc
    validate_pdf(source_path)
    gs = find_ghostscript()
    try:
        selected, profile, candidates = compress(source_path, gs, phase)
        diagnostics = {"profile": profile, "candidates": candidates,
                       "ghostscript": str(gs) if gs else None}
        write_app_debug_log("adaptive compression", context=diagnostics)
        if selected is None:
            key = "error_no_reduction" if any(c.get("checks") == "passed" for c in candidates) else "error_unknown"
            failure = next((c for c in candidates if c.get("rejected")), {})
            raise CompressError(key, failure.get("exception_message"), context={
                "stage": failure.get("stage", "candidate_selection"),
                "page_index": failure.get("page_index"),
                "candidate": failure.get("candidate"),
                "exception_type": failure.get("exception_type", "CompressError"),
                "exception_message": failure.get("exception_message", key),
            })
        rate = (1 - selected["size"] / profile["bytes"]) * 100
        return PdfResult(Path(selected["path"]), profile["bytes"], selected["size"],
                         rate, rate < LOW_REDUCTION_THRESHOLD, selected["name"],
                         ghostscript_path=str(gs) if gs else None, diagnostics=diagnostics)
    except CompressError:
        raise
    except PermissionError as exc:
        raise CompressError("error_file_in_use", str(exc)) from exc
    except Exception as exc:
        context = {
            "stage": getattr(exc, "stage", "compression"),
            "page_index": getattr(exc, "page_index", None),
            "candidate": getattr(exc, "candidate", None),
            "exception_type": getattr(exc, "exception_type", type(exc).__name__),
            "exception_message": getattr(exc, "exception_message", str(exc)),
        }
        write_app_debug_log("compression failed", exc=exc, context=context)
        raise CompressError("error_unknown", str(exc), context=context) from exc


def pdf_worker(connection):
    """Exactly one GUI worker; all PDF access stays outside the Tk process."""
    try:
        while True:
            command, generation, path = connection.recv()
            if command == "stop":
                return
            stage = "validation" if command == "check" else "compression"
            try:
                if command == "check":
                    validate_pdf(path)
                    payload = (path, path.stat().st_size, unique_output_path(path))
                    connection.send(("checked", (generation, payload)))
                else:
                    def report_phase(key):
                        nonlocal stage
                        stage = key
                        connection.send(("phase", key))
                    result = compress_pdf(path, report_phase)
                    connection.send(("success", result))
            except Exception as exc:
                cause = exc.__cause__ or exc
                context = dict(getattr(exc, "context", {}) or {})
                context.setdefault("stage", stage)
                context.setdefault("page_index", None)
                context.setdefault("candidate", None)
                context.setdefault("exception_type", type(cause).__name__)
                context.setdefault("exception_message", str(cause))
                write_app_debug_log("PDF worker failed", exc=exc, context=context)
                key = exc.message_key if isinstance(exc, CompressError) else "error_unknown"
                detail = exc.detail if isinstance(exc, CompressError) else str(exc)
                connection.send(("check_error" if command == "check" else "error",
                                 (generation, key, detail, context)))
    except (EOFError, BrokenPipeError, OSError):
        pass
    finally:
        connection.close()


def cli_write_error(message: str) -> None:
    print(message, file=sys.stderr)


def restore_cli_streams() -> None:
    """Windowed PyInstaller clears sys.stdout; retain inherited SHIMARISU pipes."""
    if os.name != "nt":
        return
    import ctypes
    import msvcrt
    get_handle = ctypes.windll.kernel32.GetStdHandle
    get_handle.argtypes = [ctypes.c_ulong]
    get_handle.restype = ctypes.c_void_p
    for name, code in (("stdout", -11), ("stderr", -12)):
        if getattr(sys, name) is not None:
            if hasattr(getattr(sys, name), "reconfigure"):
                getattr(sys, name).reconfigure(encoding="utf-8")
            continue
        handle = get_handle(code & 0xffffffff)
        if handle and handle != ctypes.c_void_p(-1).value:
            fd = msvcrt.open_osfhandle(handle, os.O_WRONLY | os.O_BINARY)
            setattr(sys, name, os.fdopen(fd, "w", encoding="utf-8", buffering=1))


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
    if "--help-cli" in argv or "--from-shimarisu" in argv:
        restore_cli_streams()
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
            result_run = safe_run(compress_pdf, source_path, title=APP_NAME, log_dir=str(app_log_dir()))
            if not result_run.ok or result_run.value is None:
                if isinstance(result_run.error, CompressError):
                    cli_write_error(cli_error_for_exception(result_run.error))
                else:
                    cli_write_error(CLI_ERROR_TEXT["unknown"])
                return 1
            result = result_run.value
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
        self.is_checking = False
        self.generation = 0
        self.pending_check = None
        self.worker_process = None
        self.worker_connection = None
        self.worker_busy = False
        self.worker_started = 0.0
        self.closing = False
        self.phase_key = "status_processing"
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
        self.poll_after_id = self.root.after(QUEUE_POLL_INTERVAL_MS, self.poll_queue)
        self.root.bind("<Configure>", self.handle_root_configure)
        self.root.protocol("WM_DELETE_WINDOW", self.close_app)
        # Developer-only readiness evidence: no effect unless explicitly enabled.
        self.root.after_idle(self.record_ready)

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
            mode="determinate",
            style="Dake.Horizontal.TProgressbar",
            length=180,
        )
        self.progress.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(16, 0))

        self.footer = tk.Frame(self.container, bg=COLORS["base_bg"])
        self.footer.pack(fill=tk.X)
        self.update_footer_layout()

        for widget in self.drop_area.winfo_children():
            widget.bind("<Button-1>", self.select_pdf_dialog)
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
            font=(self.font_family, 14 if value_var is self.reduction_rate_var else 11, "bold"),
            wraplength=290,
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

        if len(paths) != 1:
            self.show_error("error_multiple_files")
            return
        self.load_pdf(paths[0])

    def load_pdf(self, path: Path) -> None:
        if self.is_processing or self.closing:
            return
        self.generation += 1
        self.selected_pdf = None
        self.is_checking = True
        self.file_name_var.set(truncate_middle(path.name, 58))
        self.original_size_var.set(UI_TEXT["value_not_yet"])
        self.save_name_var.set(UI_TEXT["value_not_yet"])
        self.save_folder_var.set(truncate_middle(str(path.parent), 70))
        self.compressed_size_var.set(UI_TEXT["value_not_yet"])
        self.reduction_rate_var.set(UI_TEXT["value_not_yet"])
        self.notice_var.set("")
        self.drop_title_var.set(UI_TEXT["status_checking"])
        self.drop_subtitle_var.set(truncate_middle(path.name, 58))
        self.set_status("status_checking", "processing")
        self.start_progress()
        self.update_action_state()
        self.pending_check = ("check", self.generation, path)
        # Paint first, then create / dispatch to the worker.
        self.root.after_idle(self.dispatch_pending)

    def ensure_worker(self) -> None:
        if self.worker_process is not None and self.worker_process.is_alive():
            return
        ctx = mp.get_context("spawn")
        parent, child = ctx.Pipe()
        self.worker_connection = parent
        self.worker_process = ctx.Process(target=pdf_worker, args=(child,))
        self.worker_process.start()
        child.close()

    def dispatch_pending(self) -> None:
        if self.closing or self.worker_busy or self.pending_check is None:
            return
        try:
            self.ensure_worker()
            self.worker_connection.send(self.pending_check)
            self.pending_check = None
            self.worker_busy = True
            self.worker_started = time.monotonic()
        except Exception as exc:
            self.is_checking = False
            self.worker_busy = False
            self.show_error("error_unknown", str(exc))
            self.update_action_state()

    def apply_checked(self, payload) -> None:
        path, size, output_path = payload
        self.is_checking = False
        self.stop_progress()
        self.selected_pdf = path
        self.original_size_var.set(format_bytes(size))
        self.save_name_var.set(truncate_middle(output_path.name, 58))
        self.drop_title_var.set(UI_TEXT["drop_title_selected"])
        self.drop_subtitle_var.set(UI_TEXT["drop_subtitle_selected"])
        self.set_status("status_ready", "ready")
        self.update_action_state()

    def clear_selection(self) -> None:
        if self.is_processing:
            return
        self.generation += 1
        self.pending_check = None
        self.is_checking = False
        self.stop_progress()
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
        self.phase_key = "status_analyzing"
        self.notice_var.set("")
        self.set_status("status_processing", "processing")
        self.start_status_animation()
        self.update_action_state()
        self.start_progress()

        try:
            self.ensure_worker()
            self.worker_connection.send(("compress", self.generation, source_path))
            self.worker_busy = True
            self.worker_started = time.monotonic()
        except Exception as exc:
            self.handle_worker_error(CompressError("error_unknown", str(exc)))

    def poll_queue(self) -> None:
        if self.closing:
            return
        try:
            if self.worker_connection is not None:
                while self.worker_connection.poll():
                    event = self.worker_connection.recv()
                    if event[0] != "phase":
                        self.worker_busy = False
                    self.handle_queue_event(event)
                    if self.closing:
                        return
            if self.worker_busy and self.worker_process is not None:
                limit = 160 if self.is_processing else 30
                if not self.worker_process.is_alive() or time.monotonic() - self.worker_started > limit:
                    raise RuntimeError("PDF worker stopped or exceeded time budget")
        except (EOFError, OSError, RuntimeError) as exc:
            self.stop_worker()
            self.is_checking = False
            self.handle_worker_error(CompressError("error_unknown", str(exc)))
        self.dispatch_pending()
        self.poll_after_id = self.root.after(QUEUE_POLL_INTERVAL_MS, self.poll_queue)

    def stop_worker(self) -> None:
        if self.worker_process is not None:
            if self.worker_process.is_alive():
                self.worker_process.terminate()
            self.worker_process.join(0.2)
            if not self.worker_process.is_alive():
                self.worker_process.close()
            self.worker_process = None
        if self.worker_connection is not None:
            self.worker_connection.close()
            self.worker_connection = None
        self.worker_busy = False

    def close_app(self) -> None:
        if self.is_processing:
            # Let the already-requested atomic save finish. Window stays responsive.
            self.set_status("status_closing", "processing")
            self.close_after_processing = True
            return
        self.closing = True
        self.generation += 1
        self.pending_check = None
        self.stop_status_animation()
        self.root.after_cancel(self.poll_after_id)
        self.stop_worker()
        self.root.destroy()

    def record_ready(self) -> None:
        path = os.environ.get("DAKE_STARTUP_READY_FILE")
        if path:
            if not self.root.winfo_viewable():
                self.root.after(10, self.record_ready)
                return
            try:
                Path(path).write_text(str(time.perf_counter_ns()), encoding="ascii")
                if os.environ.get("DAKE_STARTUP_PROBE") == "1":
                    self.root.after(50, self.close_app)
            except OSError as exc:
                write_app_debug_log("startup measurement unavailable", exc=exc)

    def handle_queue_event(self, event: tuple[str, Any]) -> None:
        event_type, payload = event
        if event_type == "checked":
            generation, value = payload
            if generation == self.generation:
                self.apply_checked(value)
        elif event_type == "check_error":
            generation, key, detail, _context = payload
            if generation == self.generation:
                self.is_checking = False
                self.stop_progress()
                self.drop_title_var.set(UI_TEXT["status_error"])
                self.show_error(key, detail)
                self.update_action_state()
        elif event_type == "phase":
            self.phase_key = payload
            self.set_status(payload, "processing")
        elif event_type == "success":
            self.handle_success(payload)
        elif event_type == "error":
            _generation, key, detail, context = payload
            self.handle_worker_error(CompressError(key, detail, context))

    def handle_success(self, result: PdfResult) -> None:
        self.is_processing = False
        self.stop_progress()
        self.stop_status_animation()
        self.compressed_size_var.set(format_bytes(result.compressed_size))
        summary = UI_TEXT["result_reduction"].format(rate=result.reduction_rate)
        self.reduction_rate_var.set(summary)
        self.drop_title_var.set(UI_TEXT["result_sizes"].format(
            before=format_bytes(result.original_size), after=format_bytes(result.compressed_size)))
        self.drop_subtitle_var.set(summary)
        self.save_name_var.set(truncate_middle(result.output_path.name, 58))
        self.save_folder_var.set(truncate_middle(str(result.output_path.parent), 70))
        self.notice_var.set(UI_TEXT["message_low_reduction"] if result.low_reduction else "")
        self.set_status("status_low_reduction" if result.low_reduction else "status_complete",
                        "warning" if result.low_reduction else "success")
        message = UI_TEXT["result_reduction"].format(rate=result.reduction_rate)
        if result.low_reduction:
            message += "\n\n" + UI_TEXT["message_low_reduction"]
        message += "\n\n" + UI_TEXT["message_complete_detail"]
        dialog = messagebox.showwarning if result.low_reduction else messagebox.showinfo
        self.update_action_state()
        dialog(UI_TEXT["dialog_complete_title"], message)
        self.open_output_folder(result.output_path.parent)
        if getattr(self, "close_after_processing", False):
            self.close_app()

    def handle_worker_error(self, exc: CompressError) -> None:
        self.is_processing = False
        self.stop_progress()
        self.stop_status_animation()
        self.set_status("status_error", "error")
        self.update_action_state()
        self.show_error(exc.message_key, exc.detail)
        if getattr(self, "close_after_processing", False):
            self.close_app()

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
            write_app_debug_log("user-visible error", context={"key": message_key, "detail": detail})
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

    def start_progress(self) -> None:
        self.progress.configure(mode="indeterminate")
        self.progress.start(10)

    def stop_progress(self) -> None:
        self.progress.stop()
        self.progress.configure(mode="determinate", value=0)

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
        self.status_var.set(UI_TEXT[self.phase_key] + "." * (self.status_animation_index % 3 + 1))
        if elapsed >= STATUS_PHRASE_DELAY_SECONDS:
            phrases = ["status_phrase_1", "status_phrase_2", "status_phrase_3"]
            self.notice_var.set(UI_TEXT[phrases[min(2, int((elapsed - STATUS_PHRASE_DELAY_SECONDS) / 4))]])
        self.status_animation_index += 1
        self.status_animation_after_id = self.root.after(
            STATUS_ANIMATION_INTERVAL_MS,
            self.animate_processing_status,
        )

    def update_action_state(self) -> None:
        has_pdf = self.selected_pdf is not None and not self.is_checking
        if self.is_processing:
            self.execute_button.configure(state=tk.DISABLED, bg=COLORS["disabled"], cursor="arrow")
            self.select_button.configure(state=tk.DISABLED, cursor="arrow")
            self.clear_button.configure(state=tk.DISABLED, cursor="arrow")
            return

        self.select_button.configure(state=tk.NORMAL, cursor="hand2")
        self.clear_button.configure(state=tk.NORMAL if has_pdf or self.is_checking else tk.DISABLED, cursor="hand2" if has_pdf or self.is_checking else "arrow")
        self.execute_button.configure(
            state=tk.NORMAL if has_pdf else tk.DISABLED,
            bg=COLORS["accent"] if has_pdf else COLORS["disabled"],
            cursor="hand2" if has_pdf else "arrow",
        )


def check_runtime_dependencies() -> None:
    if get_fitz() is None:
        raise CompressError("error_dependency_missing")
    from adaptive import library
    library()


def create_launch_check_window() -> Any:
    root = make_root()
    root.withdraw()
    root.title(WINDOW_TITLE)
    apply_window_icon(root)
    return root


def run_app() -> None:
    load_optional_config()
    root = make_root()
    DakePdfCompressApp(root)
    root.mainloop()


def main(argv: list[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    cli_exit_code = run_cli(args)
    if cli_exit_code is not None:
        return cli_exit_code

    if "--launch-check" in args:
        return run_launch_check(
            checks=(load_optional_config, check_runtime_dependencies),
            create_window=create_launch_check_window,
            log_dir=str(app_log_dir()),
        )

    result = safe_run(run_app, title=APP_NAME, log_dir=str(app_log_dir()), show_error=True)
    return 0 if result.ok else 1


if __name__ == "__main__":
    mp.freeze_support()
    raise SystemExit(main())
