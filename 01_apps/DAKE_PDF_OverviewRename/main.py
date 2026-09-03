# -*- coding: utf-8 -*-
from __future__ import annotations

import ctypes
import queue
import sys
import threading
import time
import tkinter as tk
import webbrowser
from collections import deque
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, font as tkfont, messagebox

from rename_core import (
    FileSnapshot,
    RenamePlan,
    RenameRequest,
    RenameTransactionError,
    RenameValidationError,
    UndoRecord,
    normalize_requested_stem,
    rename_batch,
    undo_rename,
)


APP_NAME = "DakePDF俯瞰名前変更"
WINDOW_TITLE = "DakePDF俯瞰名前変更"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"
APP_USER_MODEL_ID = "Shimarisu.DakePDFOverviewRename"

UI_TEXT = {
    "main_title": "PDFを見ながら名前を変える",
    "main_description": "フォルダ内のPDFをサムネイルで一覧表示し、その場で名前を変更します。",
    "select_folder": "フォルダを選ぶ",
    "refresh": "リフレッシュ",
    "reload": "再読み込み",
    "folder_unselected": "フォルダ未選択",
    "view_size": "表示サイズ",
    "size_small": "小",
    "size_normal": "標準",
    "size_large": "大",
    "undo": "変更を元に戻す",
    "apply": "名前変更を反映 {count}",
    "page_count": "{count}ページ",
    "page_unknown": "ページ数 -",
    "thumbnail_loading": "プレビュー読み込み中",
    "thumbnail_error": "プレビューできません",
    "empty_title": "PDFがありません",
    "empty_hint": "フォルダ直下のPDFを表示します。",
    "status_empty": "0件 ｜ サムネイル 0 / 0 ｜ 0件の変更待ち",
    "status_summary": "{total}件 ｜ サムネイル {ready} / {total} ｜ {pending}件の変更待ち",
    "status_scanning": "PDFを確認しています…",
    "status_cards": "{total}件のカードを準備しています…",
    "status_renaming": "名前を変更しています…",
    "status_undoing": "変更を元に戻しています…",
    "status_complete": "{count}件の名前を変更しました。",
    "status_undo_complete": "{count}件を元の名前へ戻しました。",
    "dialog_folder": "PDFフォルダを選択",
    "discard_title": "未反映の変更",
    "discard_message": "入力中の名前を破棄してもよいですか？",
    "confirm_title": "名前変更の確認",
    "confirm_message": "{count}件のファイル名を変更します。よろしいですか？",
    "undo_confirm_title": "元に戻す確認",
    "undo_confirm_message": "直前に変更した{count}件を元の名前へ戻します。よろしいですか？",
    "error_title": "エラー",
    "error_scan": "フォルダを読み込めませんでした。\n{detail}",
    "error_dependency": "PDFプレビューに必要な pypdfium2 と Pillow を読み込めませんでした。",
    "error_rename": "名前を変更できませんでした。\n{detail}",
    "error_rollback": "名前変更に失敗し、完全には元へ戻せませんでした。対象フォルダを確認してください。\n{detail}",
    "error_validation": "変更を開始できません。入力内容とフォルダの状態を確認してください。\n\n{detail}",
    "pdf_suffix_hint": "末尾の .pdf は不要です。反映時は拡張子を1つにします。",
    "validation_line": "・{file}: {reason}",
    "validation_more": "ほか {count}件",
    "reason_empty": "新しい名前が空欄です",
    "reason_forbidden_character": "Windowsで使えない文字が含まれています",
    "reason_control_character": "制御文字が含まれています",
    "reason_trailing_dot": "末尾がピリオドです",
    "reason_trailing_space": "末尾が空白です",
    "reason_reserved_name": "Windowsの予約名です",
    "reason_too_long": "ファイル名が長すぎます",
    "reason_duplicate_destination": "変更後の名前が重複します",
    "reason_destination_exists": "変更対象外の同名ファイルがあります",
    "reason_source_missing": "元ファイルが見つかりません",
    "reason_source_changed": "読み込み後に元ファイルが変更されています",
    "reason_invalid_source": "対象外のファイルです",
    "reason_folder_unreadable": "フォルダを確認できません",
    "reason_no_files": "変更対象がありません",
    "reason_nothing_to_undo": "元に戻せる変更がありません",
    "reason_plan_changed": "確認後にフォルダの状態が変わりました",
    "reason_unknown": "入力またはフォルダの状態に問題があります",
    "preview_title": "1ページ目のプレビュー",
    "preview_loading": "大きいプレビューを読み込んでいます…",
    "preview_error": "プレビューできませんでした。",
    "footer_brand": "シンプルそれDAKEシリーズ / 止まらない、迷わない、すぐ終わる。",
    "footer_link_1": "戸建買取査定",
    "footer_link_2": "Instagram",
    "footer_separator": " ｜ ",
    "footer_copyright": COPYRIGHT,
}

LINK_URLS = {
    "footer_link_1": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "footer_link_2": "https://instagram.com/kikuta.shimarisu_fudosan",
}

THEME = {
    "background": "#F6F7F9",
    "card": "#FFFFFF",
    "text": "#1E2430",
    "muted": "#667085",
    "border": "#E6EAF0",
    "accent": "#2F6FED",
    "accent_hover": "#2458BF",
    "pending": "#EAF2FF",
    "pending_border": "#7AA7FF",
    "success": "#12B76A",
    "error": "#D92D20",
    "placeholder": "#EEF2F7",
}

FONT_CANDIDATES = ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo")
WINDOW_SIZE = "1180x780"
WINDOW_MIN_SIZE = (900, 620)
CARD_BATCH_SIZE = 24
POLL_MS = 45
THUMB_WORKERS = 3
THUMB_RENDER_BOX = (270, 350)
PREVIEW_RENDER_BOX = (850, 1050)
SIZE_CONFIG = {
    "small": (190, 160, 205),
    "normal": (240, 210, 270),
    "large": (300, 270, 350),
}

_PDFIUM = None
_PIL = None
_IMPORT_LOCK = threading.Lock()
_PDFIUM_LOCK = threading.Lock()


def set_windows_app_id() -> None:
    if not sys.platform.startswith("win"):
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)
    except Exception:
        pass


def icon_candidates() -> list[Path]:
    base = Path(__file__).resolve().parent
    candidates = [base / ".." / ".." / "02_assets" / "dake_icon.ico"]
    if getattr(sys, "frozen", False):
        runtime = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
        candidates.insert(0, runtime / "dake_icon.ico")
    return candidates


def apply_window_icon(window: tk.Misc) -> None:
    for candidate in icon_candidates():
        try:
            resolved = candidate.resolve()
            if not resolved.exists():
                continue
            window.iconbitmap(str(resolved))
            try:
                window.iconbitmap(default=str(resolved))
            except Exception:
                pass
            return
        except Exception:
            continue


def detect_font(root: tk.Misc) -> str:
    try:
        available = set(tkfont.families(root))
    except Exception:
        available = set()
    for candidate in FONT_CANDIDATES:
        if candidate in available:
            return candidate
    return "TkDefaultFont"


def load_preview_dependencies():
    global _PDFIUM, _PIL
    if _PDFIUM is not None and _PIL is not None:
        return _PDFIUM, _PIL
    with _IMPORT_LOCK:
        if _PDFIUM is None:
            import pypdfium2 as pdfium

            _PDFIUM = pdfium
        if _PIL is None:
            from PIL import Image, ImageTk

            _PIL = (Image, ImageTk)
    return _PDFIUM, _PIL


def scan_pdf_folder(folder: Path) -> list[FileSnapshot]:
    paths = sorted(
        (
            path
            for path in folder.iterdir()
            if path.is_file() and path.suffix.casefold() == ".pdf"
        ),
        key=lambda path: (path.name.casefold(), path.name),
    )
    snapshots: list[FileSnapshot] = []
    for path in paths:
        try:
            snapshots.append(FileSnapshot.capture(path))
        except OSError:
            continue
    return snapshots


def render_first_page(path: Path, box: tuple[int, int]):
    pdfium, pil_modules = load_preview_dependencies()
    Image, _ = pil_modules
    with _PDFIUM_LOCK:
        document = pdfium.PdfDocument(str(path))
        page = None
        bitmap = None
        try:
            page_count = len(document)
            if page_count < 1:
                raise ValueError("zero pages")
            page = document[0]
            width, height = page.get_size()
            scale = max(0.25, min(box[0] / max(width, 1), box[1] / max(height, 1)))
            bitmap = page.render(scale=scale)
            image = bitmap.to_pil().convert("RGB").copy()
        finally:
            if bitmap is not None:
                try:
                    bitmap.close()
                except Exception:
                    pass
            if page is not None:
                try:
                    page.close()
                except Exception:
                    pass
            try:
                document.close()
            except Exception:
                pass
    image.thumbnail(box, Image.Resampling.LANCZOS)
    return image, page_count


@dataclass(frozen=True)
class RenderRequest:
    generation: int
    kind: str
    identifier: int
    snapshot: FileSnapshot
    box: tuple[int, int]


@dataclass(frozen=True)
class RenderResult:
    request: RenderRequest
    image: object | None
    page_count: int | None
    error: str | None


class RenderPool:
    def __init__(self, worker_count: int = THUMB_WORKERS):
        self.results: queue.Queue[RenderResult] = queue.Queue()
        self._condition = threading.Condition()
        self._pending: deque[RenderRequest] = deque()
        self._active_folders: dict[Path, int] = {}
        self._generation = 0
        self._stopping = False
        self._threads = [
            threading.Thread(target=self._run, name=f"overview-thumb-{index}", daemon=True)
            for index in range(worker_count)
        ]
        for thread in self._threads:
            thread.start()

    def replace(self, generation: int, requests: list[RenderRequest]) -> None:
        with self._condition:
            self._generation = generation
            self._pending.clear()
            self._pending.extend(requests)
            self._condition.notify_all()

    def cancel(self, generation: int) -> None:
        with self._condition:
            self._generation = generation
            self._pending.clear()
            self._condition.notify_all()

    def wait_idle(self, folder: Path, timeout: float = 30.0) -> bool:
        deadline = time.monotonic() + timeout
        resolved = folder.resolve()
        with self._condition:
            while self._active_folders.get(resolved, 0):
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    return False
                self._condition.wait(remaining)
        return True

    def shutdown(self) -> None:
        with self._condition:
            self._stopping = True
            self._pending.clear()
            self._generation += 1
            self._condition.notify_all()
        for thread in self._threads:
            thread.join(timeout=1.5)

    def _run(self) -> None:
        while True:
            with self._condition:
                while not self._pending and not self._stopping:
                    self._condition.wait()
                if self._stopping:
                    return
                request = self._pending.popleft()
                if request.generation != self._generation:
                    continue
                folder = request.snapshot.path.parent.resolve()
                self._active_folders[folder] = self._active_folders.get(folder, 0) + 1
            try:
                image, page_count = render_first_page(request.snapshot.path, request.box)
                result = RenderResult(request, image, page_count, None)
            except Exception as exc:
                result = RenderResult(request, None, None, str(exc))
            finally:
                with self._condition:
                    count = self._active_folders.get(folder, 1) - 1
                    if count:
                        self._active_folders[folder] = count
                    else:
                        self._active_folders.pop(folder, None)
                    self._condition.notify_all()
            self.results.put(result)


class LatestPreviewWorker:
    def __init__(self):
        self.results: queue.Queue[RenderResult] = queue.Queue()
        self._condition = threading.Condition()
        self._pending: RenderRequest | None = None
        self._generation = 0
        self._active_folder: Path | None = None
        self._stopping = False
        self._thread = threading.Thread(target=self._run, name="overview-preview", daemon=True)
        self._thread.start()

    def request(self, request: RenderRequest) -> None:
        with self._condition:
            self._generation = request.generation
            self._pending = request
            self._condition.notify_all()

    def cancel(self, generation: int) -> None:
        with self._condition:
            self._generation = generation
            self._pending = None
            self._condition.notify_all()

    def wait_idle(self, folder: Path, timeout: float = 30.0) -> bool:
        deadline = time.monotonic() + timeout
        resolved = folder.resolve()
        with self._condition:
            while self._active_folder == resolved:
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    return False
                self._condition.wait(remaining)
        return True

    def shutdown(self) -> None:
        with self._condition:
            self._stopping = True
            self._pending = None
            self._generation += 1
            self._condition.notify_all()
        self._thread.join(timeout=1.5)

    def _run(self) -> None:
        while True:
            with self._condition:
                while self._pending is None and not self._stopping:
                    self._condition.wait()
                if self._stopping:
                    return
                request = self._pending
                self._pending = None
                if request.generation != self._generation:
                    continue
                self._active_folder = request.snapshot.path.parent.resolve()
            try:
                image, page_count = render_first_page(request.snapshot.path, request.box)
                result = RenderResult(request, image, page_count, None)
            except Exception as exc:
                result = RenderResult(request, None, None, str(exc))
            finally:
                with self._condition:
                    self._active_folder = None
                    self._condition.notify_all()
            self.results.put(result)


class LatestFolderScanner:
    def __init__(self):
        self.results: queue.Queue[tuple[int, Path, list[FileSnapshot] | None, str | None]] = queue.Queue()
        self._condition = threading.Condition()
        self._pending: tuple[int, Path] | None = None
        self._token = 0
        self._stopping = False
        self._thread = threading.Thread(target=self._run, name="overview-scan", daemon=True)
        self._thread.start()

    def request(self, token: int, folder: Path) -> None:
        with self._condition:
            self._token = token
            self._pending = (token, folder)
            self._condition.notify_all()

    def cancel(self, token: int) -> None:
        with self._condition:
            self._token = token
            self._pending = None
            self._condition.notify_all()

    def shutdown(self) -> None:
        with self._condition:
            self._stopping = True
            self._pending = None
            self._token += 1
            self._condition.notify_all()
        self._thread.join(timeout=1.5)

    def _run(self) -> None:
        while True:
            with self._condition:
                while self._pending is None and not self._stopping:
                    self._condition.wait()
                if self._stopping:
                    return
                token, folder = self._pending
                self._pending = None
            try:
                snapshots = scan_pdf_folder(folder)
                result = (token, folder, snapshots, None)
            except Exception as exc:
                result = (token, folder, None, str(exc))
            self.results.put(result)


@dataclass
class CardState:
    identifier: int
    snapshot: FileSnapshot
    original_name: str
    variable: tk.StringVar
    frame: tk.Frame
    body: tk.Frame
    image_label: tk.Label
    page_label: tk.Label
    name_label: tk.Label
    entry: tk.Entry
    suffix_label: tk.Label
    hint_label: tk.Label
    base_image: object | None = None
    photo: object | None = None
    rendered: bool = False

    @property
    def pending(self) -> bool:
        return normalize_requested_stem(self.variable.get()) != Path(self.original_name).stem


class OverviewRenameApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(*WINDOW_MIN_SIZE)
        self.root.configure(bg=THEME["background"])
        apply_window_icon(self.root)
        self.font = detect_font(root)
        self.folder: Path | None = None
        self.cards: list[CardState] = []
        self.undo_record: UndoRecord | None = None
        self.generation = 0
        self.preview_generation = 0
        self.scan_token = 0
        self.busy = False
        self.closing = False
        self.rendered_count = 0
        self._layout_after: str | None = None
        self._poll_after: str | None = None
        self._current_columns = 0
        self._preview_window: tk.Toplevel | None = None
        self._preview_label: tk.Label | None = None
        self._preview_photo = None
        self._toolbar_stacked: bool | None = None
        self._footer_stacked: bool | None = None
        self.render_pool = RenderPool()
        self.preview_worker = LatestPreviewWorker()
        self.scanner = LatestFolderScanner()
        self.operation_results: queue.Queue[tuple[str, object, object | None]] = queue.Queue()
        self.size_var = tk.StringVar(value="normal")
        self.path_var = tk.StringVar(value=UI_TEXT["folder_unselected"])
        self.status_var = tk.StringVar(value=UI_TEXT["status_empty"])
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self._poll_after = self.root.after(POLL_MS, self._poll_results)

    def _build_ui(self) -> None:
        shell = tk.Frame(self.root, bg=THEME["background"])
        shell.pack(fill="both", expand=True)

        self.header = tk.Frame(shell, bg=THEME["background"], padx=24, pady=16)
        self.header.pack(fill="x")
        self.title_label = tk.Label(
            self.header, text=UI_TEXT["main_title"], bg=THEME["background"], fg=THEME["text"],
            font=(self.font, 20, "bold"), anchor="w",
        )
        self.title_label.pack(side="left")
        self.description_label = tk.Label(
            self.header, text=UI_TEXT["main_description"], bg=THEME["background"], fg=THEME["muted"],
            font=(self.font, 10), anchor="w", padx=18,
        )
        self.description_label.pack(side="left", fill="x", expand=True)

        self.toolbar = tk.Frame(shell, bg=THEME["card"], padx=18, pady=12, highlightthickness=1, highlightbackground=THEME["border"])
        self.toolbar.pack(fill="x", padx=24)
        self.select_button = self._button(self.toolbar, UI_TEXT["select_folder"], self.choose_folder)
        self.select_button.grid(row=0, column=0, padx=(0, 8))
        self.refresh_button = self._button(self.toolbar, UI_TEXT["refresh"], self.refresh)
        self.refresh_button.grid(row=0, column=1, padx=(0, 8))
        self.reload_button = self._button(self.toolbar, UI_TEXT["reload"], self.reload)
        self.reload_button.grid(row=0, column=2, padx=(0, 12))
        self.path_label = tk.Label(self.toolbar, textvariable=self.path_var, bg=THEME["card"], fg=THEME["muted"], font=(self.font, 9), anchor="w")
        self.path_label.grid(row=0, column=3, sticky="ew")
        self.size_controls = tk.Frame(self.toolbar, bg=THEME["card"])
        self.size_controls.grid(row=0, column=4, padx=12)
        tk.Label(self.size_controls, text=UI_TEXT["view_size"], bg=THEME["card"], fg=THEME["muted"], font=(self.font, 9)).pack(side="left", padx=(0, 4))
        for value, key in (("small", "size_small"), ("normal", "size_normal"), ("large", "size_large")):
            tk.Radiobutton(
                self.size_controls, text=UI_TEXT[key], value=value, variable=self.size_var, command=self.change_size,
                bg=THEME["card"], fg=THEME["text"], activebackground=THEME["card"], selectcolor=THEME["pending"],
                font=(self.font, 9), indicatoron=False, padx=7, pady=4, relief="flat",
            ).pack(side="left")
        self.undo_button = self._button(self.toolbar, UI_TEXT["undo"], self.undo, secondary=True)
        self.undo_button.grid(row=0, column=5, padx=(0, 8))
        self.apply_button = self._button(self.toolbar, UI_TEXT["apply"].format(count=0), self.apply, primary=True)
        self.apply_button.grid(row=0, column=6)
        self.toolbar.grid_columnconfigure(3, weight=1)

        status = tk.Label(shell, textvariable=self.status_var, bg=THEME["background"], fg=THEME["muted"], font=(self.font, 9), anchor="w")
        status.pack(fill="x", padx=26, pady=(9, 7))

        viewport = tk.Frame(shell, bg=THEME["background"])
        viewport.pack(fill="both", expand=True, padx=(24, 17))
        self.canvas = tk.Canvas(viewport, bg=THEME["background"], highlightthickness=0)
        scrollbar = tk.Scrollbar(viewport, orient="vertical", command=self._on_scrollbar)
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.cards_frame = tk.Frame(self.canvas, bg=THEME["background"])
        self.cards_window = self.canvas.create_window((0, 0), window=self.cards_frame, anchor="nw")
        self.cards_frame.bind("<Configure>", self._update_scrollregion)
        self.canvas.bind("<Configure>", self._on_canvas_resize)
        self.root.bind("<MouseWheel>", self._route_mousewheel, add="+")

        self.footer = tk.Frame(shell, bg=THEME["card"], padx=24, pady=10, highlightthickness=1, highlightbackground=THEME["border"])
        self.footer.pack(fill="x", side="bottom")
        self.footer_brand = tk.Label(self.footer, text=UI_TEXT["footer_brand"], bg=THEME["card"], fg=THEME["muted"], font=(self.font, 9), anchor="w")
        self.footer_brand.pack(side="left")
        self.footer_right = tk.Frame(self.footer, bg=THEME["card"])
        self.footer_right.pack(side="right")
        self._link(self.footer_right, "footer_link_1").pack(side="left")
        tk.Label(self.footer_right, text=UI_TEXT["footer_separator"], bg=THEME["card"], fg=THEME["muted"], font=(self.font, 8)).pack(side="left")
        self._link(self.footer_right, "footer_link_2").pack(side="left")
        tk.Label(self.footer_right, text=UI_TEXT["footer_separator"] + UI_TEXT["footer_copyright"], bg=THEME["card"], fg=THEME["muted"], font=(self.font, 8)).pack(side="left")
        self.root.bind("<Configure>", self._responsive_layout, add="+")
        self._sync_controls()

    def _button(self, parent, text: str, command, primary: bool = False, secondary: bool = False) -> tk.Button:
        bg = THEME["accent"] if primary else (THEME["placeholder"] if secondary else THEME["card"])
        fg = "#FFFFFF" if primary else THEME["text"]
        return tk.Button(
            parent, text=text, command=command, bg=bg, fg=fg, activebackground=THEME["accent_hover"] if primary else THEME["pending"],
            activeforeground="#FFFFFF" if primary else THEME["text"], font=(self.font, 9, "bold"),
            relief="flat", bd=0, padx=12, pady=7, cursor="hand2",
        )

    def _link(self, parent, key: str) -> tk.Label:
        label = tk.Label(parent, text=UI_TEXT[key], bg=THEME["card"], fg=THEME["muted"], font=(self.font, 8, "underline"), cursor="hand2")
        label.bind("<Button-1>", lambda _event: webbrowser.open(LINK_URLS[key]))
        label.bind("<Enter>", lambda _event: label.configure(fg=THEME["accent"]))
        label.bind("<Leave>", lambda _event: label.configure(fg=THEME["muted"]))
        return label

    def _responsive_layout(self, event=None) -> None:
        if event is not None and event.widget is not self.root:
            return
        self._responsive_toolbar()
        self._responsive_footer()

    def _responsive_toolbar(self) -> None:
        available = max(self.toolbar.winfo_width() - 36, 1)
        fixed = (
            self.select_button.winfo_reqwidth()
            + self.refresh_button.winfo_reqwidth()
            + self.reload_button.winfo_reqwidth()
            + self.size_controls.winfo_reqwidth()
            + self.undo_button.winfo_reqwidth()
            + self.apply_button.winfo_reqwidth()
            + 280
        )
        stacked = fixed > available
        if stacked == self._toolbar_stacked:
            return
        self._toolbar_stacked = stacked
        for widget in (self.select_button, self.refresh_button, self.reload_button, self.path_label, self.size_controls, self.undo_button, self.apply_button):
            widget.grid_forget()
        if stacked:
            self.select_button.grid(row=0, column=0, padx=(0, 8), pady=(0, 8))
            self.refresh_button.grid(row=0, column=1, padx=(0, 8), pady=(0, 8))
            self.reload_button.grid(row=0, column=2, padx=(0, 12), pady=(0, 8))
            self.path_label.grid(row=0, column=3, columnspan=3, sticky="ew", pady=(0, 8))
            self.size_controls.grid(row=1, column=0, columnspan=4, sticky="w")
            self.undo_button.grid(row=1, column=4, padx=(8, 8))
            self.apply_button.grid(row=1, column=5)
        else:
            self.select_button.grid(row=0, column=0, padx=(0, 8))
            self.refresh_button.grid(row=0, column=1, padx=(0, 8))
            self.reload_button.grid(row=0, column=2, padx=(0, 12))
            self.path_label.grid(row=0, column=3, sticky="ew")
            self.size_controls.grid(row=0, column=4, padx=12)
            self.undo_button.grid(row=0, column=5, padx=(0, 8))
            self.apply_button.grid(row=0, column=6)

    def _responsive_footer(self) -> None:
        available = max(self.root.winfo_width() - 48, 1)
        stacked = self.footer_brand.winfo_reqwidth() + self.footer_right.winfo_reqwidth() > available
        if stacked == self._footer_stacked:
            return
        self._footer_stacked = stacked
        self.footer_brand.pack_forget()
        self.footer_right.pack_forget()
        if stacked:
            self.footer_brand.configure(anchor="center")
            self.footer_brand.pack(side="top", fill="x")
            self.footer_right.pack(side="top", pady=(4, 0))
        else:
            self.footer_brand.configure(anchor="w")
            self.footer_brand.pack(side="left")
            self.footer_right.pack(side="right")

    def _route_mousewheel(self, event) -> str | None:
        try:
            if event.widget.winfo_toplevel() is not self.root:
                return None
        except (AttributeError, tk.TclError):
            return None
        return self._on_mousewheel(event)

    def _on_mousewheel(self, event) -> str:
        self.canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")
        self.root.after_idle(self._reprioritize_unrendered)
        return "break"

    def _on_scrollbar(self, *arguments) -> None:
        self.canvas.yview(*arguments)
        self.root.after_idle(self._reprioritize_unrendered)

    def _update_scrollregion(self, _event=None) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_resize(self, event) -> None:
        self.canvas.itemconfigure(self.cards_window, width=event.width)
        if self._layout_after is not None:
            self.root.after_cancel(self._layout_after)
        self._layout_after = self.root.after(120, self._layout_cards)

    def _layout_cards(self) -> None:
        self._layout_after = None
        width = max(self.canvas.winfo_width(), WINDOW_MIN_SIZE[0] - 60)
        card_width = SIZE_CONFIG[self.size_var.get()][0]
        columns = max(1, width // (card_width + 14))
        if columns == self._current_columns and self.cards:
            return
        self._current_columns = columns
        for index, card in enumerate(self.cards):
            card.frame.grid(row=index // columns, column=index % columns, padx=7, pady=7, sticky="n")
        for column in range(columns):
            self.cards_frame.grid_columnconfigure(column, weight=1)
        self._update_scrollregion()

    def _clear_cards(self) -> None:
        for child in self.cards_frame.winfo_children():
            child.destroy()
        self.cards.clear()
        self.rendered_count = 0
        self._current_columns = 0
        self._update_scrollregion()

    def has_pending(self) -> bool:
        return any(card.pending for card in self.cards)

    def _confirm_discard(self) -> bool:
        if not self.has_pending():
            return True
        return messagebox.askyesno(UI_TEXT["discard_title"], UI_TEXT["discard_message"], parent=self.root)

    def choose_folder(self) -> None:
        if self.busy or not self._confirm_discard():
            return
        selected = filedialog.askdirectory(title=UI_TEXT["dialog_folder"], parent=self.root)
        if selected:
            self._start_load(Path(selected))

    def refresh(self) -> None:
        if self.busy or self.folder is None or not self._confirm_discard():
            return
        self._reset_to_initial()

    def reload(self) -> None:
        if self.busy or self.folder is None or not self._confirm_discard():
            return
        folder = self.folder
        self._close_preview()
        self._start_load(folder)
        self.canvas.yview_moveto(0.0)

    def _reset_to_initial(self) -> None:
        self.scan_token += 1
        self.generation += 1
        self.scanner.cancel(self.scan_token)
        self.render_pool.cancel(self.generation)
        self._close_preview()
        self.undo_record = None
        self.folder = None
        self._clear_cards()
        self.canvas.yview_moveto(0.0)
        self.path_var.set(UI_TEXT["folder_unselected"])
        self.status_var.set(UI_TEXT["status_empty"])
        self._sync_controls()

    def _start_load(self, folder: Path) -> None:
        self.scan_token += 1
        self.generation += 1
        self.preview_generation += 1
        self.render_pool.cancel(self.generation)
        self.preview_worker.cancel(self.preview_generation)
        self.undo_record = None
        self.folder = folder.resolve()
        self.path_var.set(str(self.folder))
        self._clear_cards()
        self.status_var.set(UI_TEXT["status_scanning"])
        self.scanner.request(self.scan_token, self.folder)
        self._sync_controls()

    def _accept_scan(self, token: int, folder: Path, snapshots: list[FileSnapshot] | None, error: str | None) -> None:
        if token != self.scan_token or self.folder != folder.resolve():
            return
        if error is not None or snapshots is None:
            self.status_var.set(UI_TEXT["status_empty"])
            messagebox.showerror(UI_TEXT["error_title"], UI_TEXT["error_scan"].format(detail=error or ""), parent=self.root)
            return
        self.status_var.set(UI_TEXT["status_cards"].format(total=len(snapshots)))
        self._create_cards_batch(snapshots, 0, token)

    def _create_cards_batch(self, snapshots: list[FileSnapshot], start: int, token: int) -> None:
        if self.closing or token != self.scan_token:
            return
        end = min(start + CARD_BATCH_SIZE, len(snapshots))
        for snapshot in snapshots[start:end]:
            self._create_card(snapshot)
        self._layout_cards()
        self.root.update_idletasks()
        if end < len(snapshots):
            self.root.after(1, self._create_cards_batch, snapshots, end, token)
        else:
            self._finish_card_creation()

    def _create_card(self, snapshot: FileSnapshot) -> None:
        identifier = len(self.cards)
        width, image_width, image_height = SIZE_CONFIG[self.size_var.get()]
        frame = tk.Frame(self.cards_frame, bg=THEME["border"], padx=1, pady=1)
        body = tk.Frame(frame, bg=THEME["card"], width=width, height=image_height + 132, padx=10, pady=10)
        body.pack(fill="both", expand=True)
        body.pack_propagate(False)
        image_label = tk.Label(
            body, text=UI_TEXT["thumbnail_loading"], bg=THEME["placeholder"], fg=THEME["muted"],
            font=(self.font, 9), width=max(1, image_width // 9), height=max(1, image_height // 18), cursor="hand2",
        )
        image_label.pack(fill="x")
        page_label = tk.Label(body, text=UI_TEXT["page_unknown"], bg=THEME["card"], fg=THEME["muted"], font=(self.font, 8), anchor="w")
        page_label.pack(fill="x", pady=(8, 2))
        name_label = tk.Label(body, text=snapshot.path.name, bg=THEME["card"], fg=THEME["text"], font=(self.font, 9, "bold"), anchor="w")
        name_label.pack(fill="x", pady=(0, 5))
        edit_row = tk.Frame(body, bg=THEME["card"])
        edit_row.pack(fill="x")
        variable = tk.StringVar(value=snapshot.path.stem)
        entry = tk.Entry(edit_row, textvariable=variable, font=(self.font, 9), relief="solid", bd=1, highlightthickness=1, highlightbackground=THEME["border"], highlightcolor=THEME["accent"])
        entry.pack(side="left", fill="x", expand=True)
        suffix_label = tk.Label(edit_row, text=".pdf", bg=THEME["card"], fg=THEME["muted"], font=(self.font, 9), padx=3)
        suffix_label.pack(side="left")
        hint_label = tk.Label(
            body, text="", bg=THEME["card"], fg=THEME["accent"], font=(self.font, 7),
            anchor="w", justify="left", wraplength=max(width - 24, 80),
        )
        hint_label.pack(fill="x", pady=(3, 0))
        card = CardState(identifier, snapshot, snapshot.path.name, variable, frame, body, image_label, page_label, name_label, entry, suffix_label, hint_label)
        variable.trace_add("write", lambda *_args, current=card: self._on_name_changed(current))
        image_label.bind("<Button-1>", lambda _event, current=card: self.show_preview(current))
        self.cards.append(card)

    def _finish_card_creation(self) -> None:
        if not self.cards:
            self._show_empty()
            self._sync_status()
            return
        self._layout_cards()
        self.root.update_idletasks()
        self._reprioritize_unrendered()
        self._sync_status()

    def _show_empty(self) -> None:
        holder = tk.Frame(self.cards_frame, bg=THEME["background"], pady=80)
        holder.grid(row=0, column=0, sticky="ew")
        tk.Label(holder, text=UI_TEXT["empty_title"], bg=THEME["background"], fg=THEME["text"], font=(self.font, 16, "bold")).pack()
        tk.Label(holder, text=UI_TEXT["empty_hint"], bg=THEME["background"], fg=THEME["muted"], font=(self.font, 10)).pack(pady=8)

    def _on_name_changed(self, card: CardState) -> None:
        entered_pdf = card.variable.get().casefold().endswith(".pdf")
        card.hint_label.configure(text=UI_TEXT["pdf_suffix_hint"] if entered_pdf else "")
        self._style_card(card)
        self._sync_status()

    def _style_card(self, card: CardState) -> None:
        pending = card.pending
        border = THEME["pending_border"] if pending else THEME["border"]
        background = THEME["pending"] if pending else THEME["card"]
        card.frame.configure(bg=border)
        card.body.configure(bg=background)
        for widget in (card.page_label, card.name_label, card.suffix_label, card.hint_label):
            widget.configure(bg=background)
        card.entry.master.configure(bg=background)

    def _sync_status(self) -> None:
        self.status_var.set(UI_TEXT["status_summary"].format(total=len(self.cards), ready=self.rendered_count, pending=sum(card.pending for card in self.cards)))
        self._sync_controls()

    def _sync_controls(self) -> None:
        pending = sum(card.pending for card in self.cards)
        self.apply_button.configure(text=UI_TEXT["apply"].format(count=pending), state="normal" if pending and not self.busy else "disabled")
        self.undo_button.configure(state="normal" if self.undo_record is not None and not self.busy else "disabled")
        self.refresh_button.configure(state="normal" if self.folder is not None and not self.busy else "disabled")
        self.reload_button.configure(state="normal" if self.folder is not None and not self.busy else "disabled")
        self.select_button.configure(state="disabled" if self.busy else "normal")
        for card in self.cards:
            card.entry.configure(state="disabled" if self.busy else "normal")

    def change_size(self) -> None:
        if not self.cards:
            return
        scroll = self.canvas.yview()[0]
        width, image_width, image_height = SIZE_CONFIG[self.size_var.get()]
        for card in self.cards:
            card.body.configure(width=width, height=image_height + 132)
            card.hint_label.configure(wraplength=max(width - 24, 80))
            card.image_label.configure(width=max(1, image_width // 9), height=max(1, image_height // 18))
            if card.base_image is not None:
                self._apply_card_image(card)
        self._current_columns = 0
        self._layout_cards()
        self.root.update_idletasks()
        self.canvas.yview_moveto(scroll)
        self.root.after_idle(self._reprioritize_unrendered)

    def _apply_card_image(self, card: CardState) -> None:
        if card.base_image is None:
            return
        _, pil_modules = load_preview_dependencies()
        Image, ImageTk = pil_modules
        _, image_width, image_height = SIZE_CONFIG[self.size_var.get()]
        image = card.base_image.copy()
        image.thumbnail((image_width, image_height), Image.Resampling.LANCZOS)
        card.photo = ImageTk.PhotoImage(image)
        card.image_label.configure(image=card.photo, text="", width=image_width, height=image_height)

    def _accept_thumbnail(self, result: RenderResult) -> None:
        request = result.request
        if request.generation != self.generation or request.identifier >= len(self.cards):
            return
        card = self.cards[request.identifier]
        if card.snapshot.path != request.snapshot.path or card.rendered:
            return
        card.rendered = True
        self.rendered_count += 1
        if result.error is not None or result.image is None:
            card.image_label.configure(text=UI_TEXT["thumbnail_error"], image="", fg=THEME["error"])
        else:
            card.base_image = result.image
            card.page_label.configure(text=UI_TEXT["page_count"].format(count=result.page_count))
            self._apply_card_image(card)
        self._sync_status()

    def show_preview(self, card: CardState) -> None:
        if self.busy:
            return
        self.preview_generation += 1
        if self._preview_window is None or not self._preview_window.winfo_exists():
            self._preview_window = tk.Toplevel(self.root)
            self._preview_window.title(UI_TEXT["preview_title"])
            self._preview_window.geometry("900x700")
            self._preview_window.configure(bg=THEME["background"])
            apply_window_icon(self._preview_window)
            self._preview_label = tk.Label(self._preview_window, bg=THEME["background"], fg=THEME["muted"], font=(self.font, 10))
            self._preview_label.pack(fill="both", expand=True, padx=16, pady=16)
            self._preview_window.protocol("WM_DELETE_WINDOW", self._close_preview)
        self._preview_window.deiconify()
        self._preview_window.lift()
        if self._preview_label is not None:
            self._preview_label.configure(text=UI_TEXT["preview_loading"], image="")
        self.preview_worker.request(RenderRequest(self.preview_generation, "preview", card.identifier, card.snapshot, PREVIEW_RENDER_BOX))

    def _close_preview(self) -> None:
        self.preview_generation += 1
        self.preview_worker.cancel(self.preview_generation)
        if self._preview_window is not None:
            self._preview_window.destroy()
        self._preview_window = None
        self._preview_label = None
        self._preview_photo = None

    def _accept_preview(self, result: RenderResult) -> None:
        if result.request.generation != self.preview_generation or self._preview_label is None:
            return
        if result.error is not None or result.image is None:
            self._preview_label.configure(text=UI_TEXT["preview_error"], image="", fg=THEME["error"])
            return
        _, pil_modules = load_preview_dependencies()
        _, ImageTk = pil_modules
        self._preview_photo = ImageTk.PhotoImage(result.image)
        self._preview_label.configure(text="", image=self._preview_photo)

    def _validation_message(self, error: RenameValidationError) -> str:
        lines: list[str] = []
        for issue in error.issues[:10]:
            reason = UI_TEXT.get(f"reason_{issue.code}", UI_TEXT["reason_unknown"])
            lines.append(UI_TEXT["validation_line"].format(file=issue.source_name or APP_NAME, reason=reason))
        if len(error.issues) > 10:
            lines.append(UI_TEXT["validation_more"].format(count=len(error.issues) - 10))
        return "\n".join(lines)

    def apply(self) -> None:
        if self.busy or self.folder is None:
            return
        requests = [RenameRequest(card.snapshot, card.variable.get()) for card in self.cards]
        try:
            from rename_core import build_rename_plan

            plan = build_rename_plan(requests)
        except RenameValidationError as exc:
            messagebox.showerror(UI_TEXT["error_title"], UI_TEXT["error_validation"].format(detail=self._validation_message(exc)), parent=self.root)
            return
        if not plan.entries:
            self._sync_status()
            return
        if not messagebox.askyesno(UI_TEXT["confirm_title"], UI_TEXT["confirm_message"].format(count=len(plan.entries)), parent=self.root):
            return
        self._begin_operation("rename", plan)

    def undo(self) -> None:
        if self.busy or self.undo_record is None:
            return
        record = self.undo_record
        if not messagebox.askyesno(UI_TEXT["undo_confirm_title"], UI_TEXT["undo_confirm_message"].format(count=len(record.entries)), parent=self.root):
            return
        self._begin_operation("undo", record)

    def _begin_operation(self, kind: str, payload: RenamePlan | UndoRecord) -> None:
        if self.folder is None:
            return
        self.busy = True
        self.generation += 1
        self.preview_generation += 1
        self.render_pool.cancel(self.generation)
        self.preview_worker.cancel(self.preview_generation)
        self.status_var.set(UI_TEXT["status_renaming"] if kind == "rename" else UI_TEXT["status_undoing"])
        self._sync_controls()
        folder = self.folder

        def run() -> None:
            try:
                if not self.render_pool.wait_idle(folder) or not self.preview_worker.wait_idle(folder):
                    raise TimeoutError("PDF renderer did not become idle")
                if kind == "rename":
                    plan = payload
                    assert isinstance(plan, RenamePlan)
                    _, record = rename_batch(
                        [RenameRequest(entry.snapshot, entry.destination.stem) for entry in plan.entries]
                    )
                    self.operation_results.put((kind, record, plan))
                else:
                    record = payload
                    assert isinstance(record, UndoRecord)
                    plan = undo_rename(record)
                    self.operation_results.put((kind, plan, record))
            except Exception as exc:
                self.operation_results.put(("error", exc, kind))

        threading.Thread(target=run, name=f"overview-{kind}", daemon=True).start()

    def _accept_operation(self, kind: str, value: object, context: object | None) -> None:
        self.busy = False
        if kind == "error":
            exc = value
            if isinstance(exc, RenameValidationError):
                detail = self._validation_message(exc)
                message = UI_TEXT["error_validation"].format(detail=detail)
            elif isinstance(exc, RenameTransactionError) and exc.rollback_errors:
                message = UI_TEXT["error_rollback"].format(detail=str(exc))
            else:
                message = UI_TEXT["error_rename"].format(detail=str(exc))
            messagebox.showerror(UI_TEXT["error_title"], message, parent=self.root)
            self._sync_status()
            return
        if kind == "rename":
            record = value
            plan = context
            assert isinstance(record, UndoRecord) and isinstance(plan, RenamePlan)
            self.undo_record = record
            renamed = {entry.original_path: entry.renamed_snapshot for entry in record.entries}
            for card in self.cards:
                snapshot = renamed.get(card.snapshot.path)
                if snapshot is None:
                    continue
                card.snapshot = snapshot
                card.original_name = snapshot.path.name
                card.variable.set(snapshot.path.stem)
                card.name_label.configure(text=snapshot.path.name)
                self._style_card(card)
            self.status_var.set(UI_TEXT["status_complete"].format(count=len(record.entries)))
        else:
            plan = value
            record = context
            assert isinstance(plan, RenamePlan) and isinstance(record, UndoRecord)
            restored = {entry.source: FileSnapshot.capture(entry.destination) for entry in plan.entries}
            for card in self.cards:
                snapshot = restored.get(card.snapshot.path)
                if snapshot is None:
                    continue
                card.snapshot = snapshot
                card.original_name = snapshot.path.name
                card.variable.set(snapshot.path.stem)
                card.name_label.configure(text=snapshot.path.name)
                self._style_card(card)
            self.undo_record = None
            self.status_var.set(UI_TEXT["status_undo_complete"].format(count=len(plan.entries)))
        self._reschedule_unrendered()
        self._sync_controls()

    def _reschedule_unrendered(self) -> None:
        self._reprioritize_unrendered()

    def _reprioritize_unrendered(self) -> None:
        if self.closing or self.busy:
            return
        top = self.canvas.canvasy(0)
        bottom = top + self.canvas.winfo_height()
        pending = [card for card in self.cards if not card.rendered]
        pending.sort(
            key=lambda card: (
                not (card.frame.winfo_y() + card.frame.winfo_height() >= top and card.frame.winfo_y() <= bottom),
                card.identifier,
            )
        )
        requests = [
            RenderRequest(self.generation, "thumbnail", card.identifier, card.snapshot, THUMB_RENDER_BOX)
            for card in pending
        ]
        self.render_pool.replace(self.generation, requests)

    def _poll_results(self) -> None:
        if self.closing:
            return
        for _ in range(32):
            try:
                token, folder, snapshots, error = self.scanner.results.get_nowait()
            except queue.Empty:
                break
            self._accept_scan(token, folder, snapshots, error)
        for _ in range(32):
            try:
                result = self.render_pool.results.get_nowait()
            except queue.Empty:
                break
            self._accept_thumbnail(result)
        for _ in range(8):
            try:
                result = self.preview_worker.results.get_nowait()
            except queue.Empty:
                break
            self._accept_preview(result)
        for _ in range(4):
            try:
                kind, value, context = self.operation_results.get_nowait()
            except queue.Empty:
                break
            self._accept_operation(kind, value, context)
        self._poll_after = self.root.after(POLL_MS, self._poll_results)

    def on_close(self) -> None:
        if self.busy:
            return
        if not self._confirm_discard():
            return
        self.closing = True
        if self._poll_after is not None:
            try:
                self.root.after_cancel(self._poll_after)
            except Exception:
                pass
        if self._layout_after is not None:
            try:
                self.root.after_cancel(self._layout_after)
            except Exception:
                pass
        self.scanner.shutdown()
        self.render_pool.shutdown()
        self.preview_worker.shutdown()
        self.root.destroy()


def main() -> None:
    set_windows_app_id()
    root = tk.Tk()
    OverviewRenameApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
