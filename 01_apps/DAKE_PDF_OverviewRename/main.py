# -*- coding: utf-8 -*-
from __future__ import annotations

import ctypes
import os
import queue
import sys
import threading
import uuid
import webbrowser
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox

try:
    from PIL import Image, ImageTk
except Exception:
    Image = None  # type: ignore[assignment]
    ImageTk = None  # type: ignore[assignment]

try:
    import pypdfium2 as pdfium
except Exception:
    pdfium = None  # type: ignore[assignment]


APP_NAME = "DakePDF俯瞰名前変更"
WINDOW_TITLE = "PDF俯瞰名前変更"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"
APP_USER_MODEL_ID = "Shimarisu.DakePDFOverviewRename"
COMMON_ICON_RELATIVE = Path("..") / ".." / "02_assets" / "dake_icon.ico"
COMMON_ICON_FILENAME = "dake_icon.ico"

UI_TEXT = {
    "main_title": "PDFを見ながら名前を変える",
    "main_description": "フォルダ内のPDFをサムネイルで一覧表示し、その場で名前を変更します。",
    "button_select_folder": "フォルダを選ぶ",
    "button_refresh": "リフレッシュ",
    "button_execute": "名前変更を反映",
    "button_undo": "変更を元に戻す",
    "folder_empty": "フォルダを選択してください",
    "folder_dialog_title": "PDFが入っているフォルダを選択",
    "empty_title": "PDFがありません",
    "empty_subtitle": "PDFが入っているフォルダを選択してください",
    "thumbnail_loading": "読み込み中",
    "thumbnail_unavailable": "プレビューできません",
    "page_count_loading": "ページ数確認中",
    "page_count_format": "{count}ページ",
    "status_idle": "未選択",
    "status_loading": "PDFを読み込み中",
    "status_rendering": "サムネイル生成中 {done} / {total}",
    "status_ready": "準備完了",
    "status_changes": "{count}件の変更待ち",
    "status_processing": "名前変更中",
    "status_complete": "名前変更完了",
    "status_error": "エラー",
    "status_undo_ready": "直前の変更を元に戻せます",
    "display_size_label": "表示サイズ",
    "display_size_small": "小",
    "display_size_standard": "標準",
    "display_size_large": "大",
    "suffix_pdf": ".pdf",
    "original_name_prefix": "元: ",
    "preview_title": "PDFプレビュー",
    "preview_loading": "プレビューを読み込み中",
    "preview_close": "閉じる",
    "preview_hint": "Escでも閉じられます",
    "dialog_error_title": "エラー",
    "dialog_warning_title": "確認",
    "dialog_complete_title": "完了",
    "dialog_undo_title": "元に戻す",
    "error_dependency": "PDFのプレビューに必要なライブラリを読み込めませんでした。pypdfium2 と Pillow を確認してください。",
    "error_folder_missing": "選択したフォルダを開けませんでした。",
    "error_invalid_empty": "ファイル名を空欄にはできません。",
    "error_invalid_chars": "ファイル名に使用できない文字が含まれています。",
    "error_invalid_tail": "ファイル名の末尾にピリオドまたは空白は使用できません。",
    "error_reserved_name": "Windowsで予約されている名前は使用できません。",
    "error_duplicate_name": "同じファイル名が複数あります。",
    "error_existing_name": "変更先と同じ名前のファイルが既にあります。",
    "error_rename_failed": "名前変更を完了できませんでした。可能な範囲で元の名前へ戻しました。",
    "error_undo_failed": "元の名前へ戻せませんでした。ファイルの状態を確認してください。",
    "confirm_rename": "{count}件のPDFの名前を変更します。\nよろしいですか？",
    "confirm_undo": "直前に変更した{count}件を元の名前へ戻します。\nよろしいですか？",
    "complete_rename": "{count}件の名前を変更しました。",
    "complete_undo": "{count}件を元の名前へ戻しました。",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_tagline": "止まらない、迷わない、すぐ終わる。",
    "footer_tagline_separator": " / ",
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
    "selection_bg": "#EAF2FF",
    "selection_border": "#7AA7FF",
    "success": "#12B76A",
    "error": "#D92D20",
    "disabled": "#98A2B3",
}

FONT_CANDIDATES = ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo")
WINDOW_SIZE = "1240x820"
WINDOW_MIN_SIZE = (980, 660)
FOOTER_BREAKPOINT = 1080
RENDER_WORKERS = 1
QUEUE_POLL_MS = 40
PDFIUM_LOCK = threading.Lock()
INVALID_FILENAME_CHARS = set('<>:"/\\|?*')
RESERVED_WINDOWS_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}
SIZE_PRESETS = {
    "small": {"thumb_w": 150, "thumb_h": 194, "card_w": 190},
    "standard": {"thumb_w": 190, "thumb_h": 246, "card_w": 232},
    "large": {"thumb_w": 245, "thumb_h": 317, "card_w": 286},
}


@dataclass
class PdfItem:
    path: Path
    original_stem: str
    name_var: tk.StringVar
    card: tk.Frame | None = None
    original_label: tk.Label | None = None
    thumb_label: tk.Label | None = None
    page_label: tk.Label | None = None
    entry: tk.Entry | None = None
    suffix_label: tk.Label | None = None
    photo: object | None = None
    page_count: int | None = None
    error: str | None = None

    @property
    def proposed_stem(self) -> str:
        return self.name_var.get().strip()

    @property
    def changed(self) -> bool:
        return self.proposed_stem != self.original_stem


@dataclass(frozen=True)
class RenderResult:
    token: int
    path: Path
    page_count: int | None
    image: object | None
    error: str | None


class RenameValidationError(Exception):
    pass


def set_windows_app_id() -> None:
    if not sys.platform.startswith("win"):
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)
    except Exception:
        pass


def detect_font_family(root: tk.Misc) -> str:
    try:
        families = set(tkfont.families(root))
    except Exception:
        families = set()
    for name in FONT_CANDIDATES:
        if name in families:
            return name
    return "TkDefaultFont"


def app_icon_path() -> Path | None:
    if getattr(sys, "frozen", False):
        meipass = Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
        candidate = meipass / COMMON_ICON_FILENAME
        return candidate if candidate.exists() else None
    candidate = (Path(__file__).resolve().parent / COMMON_ICON_RELATIVE).resolve()
    return candidate if candidate.exists() else None


def apply_window_icon(window: tk.Misc) -> None:
    icon = app_icon_path()
    if icon is None:
        return
    try:
        window.iconbitmap(str(icon))
    except Exception:
        pass
    try:
        window.iconbitmap(default=str(icon))
    except Exception:
        pass


def bind_link(label: tk.Label, url: str) -> None:
    label.configure(cursor="hand2")
    label.bind("<Button-1>", lambda _event: webbrowser.open(url))
    label.bind("<Enter>", lambda _event: label.configure(fg=THEME["accent"]))
    label.bind("<Leave>", lambda _event: label.configure(fg=THEME["muted"]))


def normalize_case_name(name: str) -> str:
    return os.path.normcase(name).casefold()


def validate_stem(stem: str) -> None:
    if not stem:
        raise RenameValidationError(UI_TEXT["error_invalid_empty"])
    if any(ch in INVALID_FILENAME_CHARS or ord(ch) < 32 for ch in stem):
        raise RenameValidationError(UI_TEXT["error_invalid_chars"])
    if stem.endswith(".") or stem.endswith(" "):
        raise RenameValidationError(UI_TEXT["error_invalid_tail"])
    if stem.upper() in RESERVED_WINDOWS_NAMES:
        raise RenameValidationError(UI_TEXT["error_reserved_name"])


def execute_two_phase_rename(changes: list[tuple[Path, Path]]) -> None:
    if not changes:
        return

    temp_for_source: dict[Path, Path] = {}
    moved_to_temp: list[tuple[Path, Path]] = []
    moved_to_final: list[tuple[Path, Path, Path]] = []

    try:
        for source, _dest in changes:
            temp = source.with_name(f".__dake_tmp_{uuid.uuid4().hex}.pdf")
            while temp.exists():
                temp = source.with_name(f".__dake_tmp_{uuid.uuid4().hex}.pdf")
            source.rename(temp)
            temp_for_source[source] = temp
            moved_to_temp.append((source, temp))

        for source, dest in changes:
            temp = temp_for_source[source]
            temp.rename(dest)
            moved_to_final.append((source, temp, dest))
    except Exception as exc:
        rollback_temp: dict[Path, Path] = {}
        for source, _temp, dest in reversed(moved_to_final):
            try:
                rt = dest.with_name(f".__dake_rollback_{uuid.uuid4().hex}.pdf")
                while rt.exists():
                    rt = dest.with_name(f".__dake_rollback_{uuid.uuid4().hex}.pdf")
                dest.rename(rt)
                rollback_temp[source] = rt
            except Exception:
                pass

        for source, temp in reversed(moved_to_temp):
            current = rollback_temp.get(source, temp)
            try:
                if current.exists() and not source.exists():
                    current.rename(source)
            except Exception:
                pass
        raise exc


def render_first_page(path: Path, width: int, height: int) -> tuple[int, object]:
    if pdfium is None or Image is None:
        raise RuntimeError(UI_TEXT["error_dependency"])

    # PDFiumはプロセス内の同時呼び出しが非対応のため、全レンダリングを直列化する。
    # UI自体はバックグラウンド処理により止めない。
    with PDFIUM_LOCK:
        document = pdfium.PdfDocument(str(path))
        try:
            page_count = len(document)
            if page_count <= 0:
                raise RuntimeError(UI_TEXT["thumbnail_unavailable"])
            page = document[0]
            try:
                page_width, page_height = page.get_size()
                scale = min(width / max(page_width, 1), height / max(page_height, 1))
                scale = max(scale, 0.2)
                bitmap = page.render(scale=scale)
                try:
                    pil_image = bitmap.to_pil().convert("RGB")
                finally:
                    try:
                        bitmap.close()
                    except Exception:
                        pass
            finally:
                try:
                    page.close()
                except Exception:
                    pass
        finally:
            try:
                document.close()
            except Exception:
                pass

    canvas = Image.new("RGB", (width, height), "white")
    pil_image.thumbnail((width, height))
    x = (width - pil_image.width) // 2
    y = (height - pil_image.height) // 2
    canvas.paste(pil_image, (x, y))
    return page_count, canvas


class DakePdfOverviewRenameApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(*WINDOW_MIN_SIZE)
        self.root.configure(bg=THEME["background"])
        apply_window_icon(self.root)

        self.font_name = detect_font_family(root)
        self.current_folder: Path | None = None
        self.items: list[PdfItem] = []
        self.item_by_path: dict[Path, PdfItem] = {}
        self.render_token = 0
        self.render_total = 0
        self.render_done = 0
        self.render_queue: queue.Queue[RenderResult] = queue.Queue()
        self.executor = ThreadPoolExecutor(max_workers=RENDER_WORKERS, thread_name_prefix="dake-pdf-thumb")
        self.size_mode = tk.StringVar(value="standard")
        self.status_var = tk.StringVar(value=UI_TEXT["status_idle"])
        self.folder_var = tk.StringVar(value=UI_TEXT["folder_empty"])
        self.undo_changes: list[tuple[Path, Path]] = []
        self.is_processing = False
        self._relayout_after_id: str | None = None

        self._build_ui()
        self.root.after(QUEUE_POLL_MS, self._poll_render_queue)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self) -> None:
        shell = tk.Frame(self.root, bg=THEME["background"])
        shell.pack(fill="both", expand=True)

        main = tk.Frame(shell, bg=THEME["background"])
        main.pack(fill="both", expand=True, padx=24, pady=(20, 8))

        header = tk.Frame(main, bg=THEME["background"])
        header.pack(fill="x", pady=(0, 14))

        tk.Label(
            header,
            text=UI_TEXT["main_title"],
            font=(self.font_name, 20, "bold"),
            bg=THEME["background"],
            fg=THEME["text"],
        ).pack(side="left", anchor="w")

        tk.Label(
            header,
            text=UI_TEXT["main_description"],
            font=(self.font_name, 10),
            bg=THEME["background"],
            fg=THEME["muted"],
        ).pack(side="left", anchor="w", padx=(14, 0), pady=(6, 0))

        toolbar = tk.Frame(main, bg=THEME["card"], highlightthickness=1, highlightbackground=THEME["border"])
        toolbar.pack(fill="x", pady=(0, 12))

        left_tools = tk.Frame(toolbar, bg=THEME["card"])
        left_tools.pack(side="left", fill="x", expand=True, padx=12, pady=10)

        select_button = self._secondary_button(left_tools, UI_TEXT["button_select_folder"], self.select_folder)
        select_button.pack(side="left")

        refresh_button = self._secondary_button(left_tools, UI_TEXT["button_refresh"], self.refresh_folder)
        refresh_button.pack(side="left", padx=(8, 0))

        folder_label = tk.Label(
            left_tools,
            textvariable=self.folder_var,
            font=(self.font_name, 9),
            bg=THEME["card"],
            fg=THEME["muted"],
            anchor="w",
        )
        folder_label.pack(side="left", fill="x", expand=True, padx=(12, 8))

        right_tools = tk.Frame(toolbar, bg=THEME["card"])
        right_tools.pack(side="right", padx=12, pady=10)

        tk.Label(
            right_tools,
            text=UI_TEXT["display_size_label"],
            font=(self.font_name, 9),
            bg=THEME["card"],
            fg=THEME["muted"],
        ).pack(side="left", padx=(0, 6))

        for value, text in (
            ("small", UI_TEXT["display_size_small"]),
            ("standard", UI_TEXT["display_size_standard"]),
            ("large", UI_TEXT["display_size_large"]),
        ):
            rb = tk.Radiobutton(
                right_tools,
                text=text,
                value=value,
                variable=self.size_mode,
                command=self._change_size_mode,
                indicatoron=False,
                font=(self.font_name, 9),
                bg=THEME["card"],
                fg=THEME["text"],
                activebackground=THEME["selection_bg"],
                activeforeground=THEME["text"],
                selectcolor=THEME["selection_bg"],
                relief="flat",
                bd=0,
                padx=8,
                pady=4,
                cursor="hand2",
            )
            rb.pack(side="left", padx=(0, 2))

        self.undo_button = self._secondary_button(right_tools, UI_TEXT["button_undo"], self.undo_last)
        self.undo_button.pack(side="left", padx=(10, 8))
        self.undo_button.configure(state="disabled")

        self.execute_button = self._primary_button(right_tools, UI_TEXT["button_execute"], self.apply_renames)
        self.execute_button.pack(side="left")
        self.execute_button.configure(state="disabled")

        content_border = tk.Frame(main, bg=THEME["border"])
        content_border.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(
            content_border,
            bg=THEME["background"],
            highlightthickness=0,
            bd=0,
        )
        self.scrollbar = tk.Scrollbar(content_border, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True, padx=(1, 0), pady=1)
        self.scrollbar.pack(side="right", fill="y", pady=1, padx=(0, 1))

        self.cards_frame = tk.Frame(self.canvas, bg=THEME["background"])
        self.canvas_window = self.canvas.create_window((0, 0), window=self.cards_frame, anchor="nw")

        self.cards_frame.bind("<Configure>", self._sync_scrollregion)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        self.empty_frame = tk.Frame(self.cards_frame, bg=THEME["background"])
        tk.Label(
            self.empty_frame,
            text=UI_TEXT["empty_title"],
            font=(self.font_name, 15, "bold"),
            bg=THEME["background"],
            fg=THEME["text"],
        ).pack(pady=(80, 6))
        tk.Label(
            self.empty_frame,
            text=UI_TEXT["empty_subtitle"],
            font=(self.font_name, 10),
            bg=THEME["background"],
            fg=THEME["muted"],
        ).pack()
        self.empty_frame.grid(row=0, column=0, sticky="nsew")

        status_row = tk.Frame(main, bg=THEME["background"])
        status_row.pack(fill="x", pady=(8, 0))
        tk.Label(
            status_row,
            textvariable=self.status_var,
            font=(self.font_name, 9),
            bg=THEME["background"],
            fg=THEME["muted"],
        ).pack(side="left")

        self.footer = tk.Frame(shell, bg=THEME["background"])
        self.footer.pack(fill="x", padx=24, pady=(4, 10))
        self._build_footer_wide()
        self.root.bind("<Configure>", self._on_root_resize)

    def _primary_button(self, parent: tk.Misc, text: str, command) -> tk.Button:
        return tk.Button(
            parent,
            text=text,
            command=command,
            font=(self.font_name, 9, "bold"),
            bg=THEME["accent"],
            fg="white",
            activebackground=THEME["accent_hover"],
            activeforeground="white",
            disabledforeground="#D0D5DD",
            relief="flat",
            bd=0,
            padx=14,
            pady=7,
            cursor="hand2",
        )

    def _secondary_button(self, parent: tk.Misc, text: str, command) -> tk.Button:
        return tk.Button(
            parent,
            text=text,
            command=command,
            font=(self.font_name, 9),
            bg=THEME["card"],
            fg=THEME["text"],
            activebackground=THEME["selection_bg"],
            activeforeground=THEME["text"],
            relief="solid",
            bd=1,
            padx=12,
            pady=6,
            cursor="hand2",
        )

    def _build_footer_wide(self) -> None:
        for child in self.footer.winfo_children():
            child.destroy()

        left = tk.Frame(self.footer, bg=THEME["background"])
        left.pack(side="left")
        right = tk.Frame(self.footer, bg=THEME["background"])
        right.pack(side="right")

        footer_text = (
            UI_TEXT["footer_left"]
            + UI_TEXT["footer_tagline_separator"]
            + UI_TEXT["footer_tagline"]
        )
        tk.Label(
            left,
            text=footer_text,
            font=(self.font_name, 9),
            bg=THEME["background"],
            fg=THEME["muted"],
        ).pack(side="left")

        self._append_footer_links(right)

    def _build_footer_narrow(self) -> None:
        for child in self.footer.winfo_children():
            child.destroy()

        footer_text = (
            UI_TEXT["footer_left"]
            + UI_TEXT["footer_tagline_separator"]
            + UI_TEXT["footer_tagline"]
        )
        row1 = tk.Frame(self.footer, bg=THEME["background"])
        row1.pack(anchor="center")
        tk.Label(
            row1,
            text=footer_text,
            font=(self.font_name, 9),
            bg=THEME["background"],
            fg=THEME["muted"],
        ).pack()

        row2 = tk.Frame(self.footer, bg=THEME["background"])
        row2.pack(anchor="center", pady=(3, 0))
        self._append_footer_links(row2)

    def _append_footer_links(self, parent: tk.Misc) -> None:
        link1 = tk.Label(
            parent,
            text=UI_TEXT["footer_link_1"],
            font=(self.font_name, 9),
            bg=THEME["background"],
            fg=THEME["muted"],
        )
        link1.pack(side="left")
        bind_link(link1, LINK_URLS["footer_link_1"])

        tk.Label(
            parent,
            text=UI_TEXT["footer_separator"],
            font=(self.font_name, 9),
            bg=THEME["background"],
            fg=THEME["muted"],
        ).pack(side="left")

        link2 = tk.Label(
            parent,
            text=UI_TEXT["footer_link_2"],
            font=(self.font_name, 9),
            bg=THEME["background"],
            fg=THEME["muted"],
        )
        link2.pack(side="left")
        bind_link(link2, LINK_URLS["footer_link_2"])

        tk.Label(
            parent,
            text=f"{UI_TEXT['footer_separator']}{UI_TEXT['footer_copyright']}",
            font=(self.font_name, 9),
            bg=THEME["background"],
            fg=THEME["muted"],
        ).pack(side="left")

    def _on_root_resize(self, event: tk.Event) -> None:
        if event.widget is not self.root:
            return
        if event.width < FOOTER_BREAKPOINT:
            if getattr(self, "_footer_mode", None) != "narrow":
                self._footer_mode = "narrow"
                self._build_footer_narrow()
        else:
            if getattr(self, "_footer_mode", None) != "wide":
                self._footer_mode = "wide"
                self._build_footer_wide()

    def _sync_scrollregion(self, _event=None) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event: tk.Event) -> None:
        self.canvas.itemconfigure(self.canvas_window, width=event.width)
        if self._relayout_after_id is not None:
            try:
                self.root.after_cancel(self._relayout_after_id)
            except Exception:
                pass
        self._relayout_after_id = self.root.after(80, self._layout_cards)

    def _on_mousewheel(self, event: tk.Event) -> None:
        if not self.canvas.winfo_ismapped():
            return
        try:
            delta = int(-1 * (event.delta / 120))
            self.canvas.yview_scroll(delta, "units")
        except Exception:
            pass

    def select_folder(self) -> None:
        if self.is_processing:
            return
        folder = filedialog.askdirectory(title=UI_TEXT["folder_dialog_title"])
        if not folder:
            return
        self.load_folder(Path(folder))

    def refresh_folder(self) -> None:
        if self.is_processing or self.current_folder is None:
            return
        self.load_folder(self.current_folder)

    def load_folder(self, folder: Path) -> None:
        if not folder.exists() or not folder.is_dir():
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["error_folder_missing"])
            return

        self.render_token += 1
        token = self.render_token
        self.current_folder = folder
        self.folder_var.set(str(folder))
        self.status_var.set(UI_TEXT["status_loading"])
        self.undo_changes = []
        self.undo_button.configure(state="disabled")
        self._clear_cards()

        try:
            paths = sorted(
                (p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"),
                key=lambda p: p.name.casefold(),
            )
        except Exception:
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["error_folder_missing"])
            self.status_var.set(UI_TEXT["status_error"])
            return

        if not paths:
            self.empty_frame.grid(row=0, column=0, sticky="nsew")
            self.status_var.set(UI_TEXT["empty_title"])
            self._update_actions()
            return

        self.empty_frame.grid_remove()
        for path in paths:
            var = tk.StringVar(value=path.stem)
            item = PdfItem(path=path, original_stem=path.stem, name_var=var)
            var.trace_add("write", lambda *_args, current=item: self._on_name_change(current))
            self.items.append(item)
            self.item_by_path[path] = item
            self._create_card(item)

        self._layout_cards()
        self.render_total = len(self.items)
        self.render_done = 0
        self.status_var.set(UI_TEXT["status_rendering"].format(done=0, total=self.render_total))

        preset = SIZE_PRESETS[self.size_mode.get()]
        for item in self.items:
            self.executor.submit(
                self._render_worker,
                token,
                item.path,
                preset["thumb_w"],
                preset["thumb_h"],
            )
        self._update_actions()

    def _clear_cards(self) -> None:
        self.items.clear()
        self.item_by_path.clear()
        for child in self.cards_frame.winfo_children():
            if child is not self.empty_frame:
                child.destroy()

    def _create_card(self, item: PdfItem) -> None:
        preset = SIZE_PRESETS[self.size_mode.get()]
        card = tk.Frame(
            self.cards_frame,
            bg=THEME["card"],
            highlightthickness=1,
            highlightbackground=THEME["border"],
            width=preset["card_w"],
        )
        item.card = card

        thumb_wrap = tk.Frame(card, bg=THEME["card"])
        thumb_wrap.pack(fill="x", padx=12, pady=(12, 8))

        thumb = tk.Label(
            thumb_wrap,
            text=UI_TEXT["thumbnail_loading"],
            font=(self.font_name, 9),
            width=max(1, preset["thumb_w"] // 9),
            height=max(1, preset["thumb_h"] // 18),
            bg="white",
            fg=THEME["muted"],
            relief="solid",
            bd=1,
            cursor="hand2",
        )
        thumb.pack(anchor="center")
        thumb.bind("<Button-1>", lambda _event, current=item: self.open_preview(current))
        item.thumb_label = thumb

        page_label = tk.Label(
            card,
            text=UI_TEXT["page_count_loading"],
            font=(self.font_name, 8),
            bg=THEME["card"],
            fg=THEME["muted"],
            anchor="w",
        )
        page_label.pack(fill="x", padx=12)
        item.page_label = page_label

        original_label = tk.Label(
            card,
            text=f"{UI_TEXT['original_name_prefix']}{item.path.name}",
            font=(self.font_name, 8),
            bg=THEME["card"],
            fg=THEME["muted"],
            anchor="w",
        )
        original_label.pack(fill="x", padx=12, pady=(4, 5))
        item.original_label = original_label

        name_row = tk.Frame(card, bg=THEME["card"])
        name_row.pack(fill="x", padx=12, pady=(0, 12))

        entry = tk.Entry(
            name_row,
            textvariable=item.name_var,
            font=(self.font_name, 10),
            bg="white",
            fg=THEME["text"],
            insertbackground=THEME["text"],
            relief="solid",
            bd=1,
        )
        entry.pack(side="left", fill="x", expand=True, ipady=5)
        item.entry = entry

        suffix = tk.Label(
            name_row,
            text=UI_TEXT["suffix_pdf"],
            font=(self.font_name, 9),
            bg=THEME["card"],
            fg=THEME["muted"],
        )
        suffix.pack(side="left", padx=(4, 0))
        item.suffix_label = suffix

    def _layout_cards(self) -> None:
        self._relayout_after_id = None
        if not self.items:
            return
        preset = SIZE_PRESETS[self.size_mode.get()]
        available = max(self.canvas.winfo_width(), WINDOW_MIN_SIZE[0] - 48)
        gap = 12
        columns = max(1, available // (preset["card_w"] + gap))

        for idx, item in enumerate(self.items):
            if item.card is None:
                continue
            row, col = divmod(idx, columns)
            item.card.configure(width=preset["card_w"])
            item.card.grid(row=row, column=col, padx=(0 if col == 0 else gap, 0), pady=(0, gap), sticky="nw")

        self.cards_frame.update_idletasks()
        self._sync_scrollregion()

    def _change_size_mode(self) -> None:
        if not self.items:
            return
        folder = self.current_folder
        if folder is not None:
            pending = {item.path: item.name_var.get() for item in self.items}
            self.render_token += 1
            token = self.render_token
            for child in self.cards_frame.winfo_children():
                if child is not self.empty_frame:
                    child.destroy()
            for item in self.items:
                item.card = None
                item.thumb_label = None
                item.page_label = None
                item.entry = None
                item.suffix_label = None
                item.photo = None
                item.name_var.set(pending[item.path])
                self._create_card(item)
            self._layout_cards()
            self.render_done = 0
            self.render_total = len(self.items)
            preset = SIZE_PRESETS[self.size_mode.get()]
            for item in self.items:
                self.executor.submit(
                    self._render_worker,
                    token,
                    item.path,
                    preset["thumb_w"],
                    preset["thumb_h"],
                )
            self.status_var.set(UI_TEXT["status_rendering"].format(done=0, total=self.render_total))

    def _render_worker(self, token: int, path: Path, width: int, height: int) -> None:
        try:
            count, image = render_first_page(path, width, height)
            result = RenderResult(token, path, count, image, None)
        except Exception as exc:
            result = RenderResult(token, path, None, None, str(exc))
        self.render_queue.put(result)

    def _poll_render_queue(self) -> None:
        try:
            while True:
                result = self.render_queue.get_nowait()
                if result.token != self.render_token:
                    continue
                item = self.item_by_path.get(result.path)
                if item is None:
                    continue
                self.render_done += 1
                item.page_count = result.page_count
                item.error = result.error

                if item.page_label is not None:
                    if result.page_count is not None:
                        item.page_label.configure(text=UI_TEXT["page_count_format"].format(count=result.page_count))
                    else:
                        item.page_label.configure(text=UI_TEXT["thumbnail_unavailable"], fg=THEME["error"])

                if item.thumb_label is not None:
                    if result.image is not None and ImageTk is not None:
                        photo = ImageTk.PhotoImage(result.image)
                        item.photo = photo
                        item.thumb_label.configure(image=photo, text="", width=0, height=0)
                    else:
                        item.thumb_label.configure(text=UI_TEXT["thumbnail_unavailable"])

                if self.render_done >= self.render_total:
                    self._set_ready_status()
                else:
                    self.status_var.set(
                        UI_TEXT["status_rendering"].format(done=self.render_done, total=self.render_total)
                    )
        except queue.Empty:
            pass
        finally:
            self.root.after(QUEUE_POLL_MS, self._poll_render_queue)

    def _on_name_change(self, item: PdfItem) -> None:
        self._update_card_style(item)
        self._update_actions()
        if self.render_total and self.render_done < self.render_total:
            return
        self._set_ready_status()

    def _update_card_style(self, item: PdfItem) -> None:
        if item.card is None:
            return
        changed = item.changed
        bg = THEME["selection_bg"] if changed else THEME["card"]
        border = THEME["selection_border"] if changed else THEME["border"]
        item.card.configure(bg=bg, highlightbackground=border)
        for widget in (item.original_label, item.page_label, item.suffix_label):
            if widget is not None:
                widget.configure(bg=bg)
        if item.entry is not None:
            try:
                item.entry.master.configure(bg=bg)
            except Exception:
                pass
        if item.thumb_label is not None:
            try:
                item.thumb_label.master.configure(bg=bg)
            except Exception:
                pass

    def _changed_items(self) -> list[PdfItem]:
        return [item for item in self.items if item.changed]

    def _update_actions(self) -> None:
        count = len(self._changed_items())
        if count > 0 and not self.is_processing:
            self.execute_button.configure(state="normal", text=f"{UI_TEXT['button_execute']} {count}")
        else:
            self.execute_button.configure(state="disabled", text=UI_TEXT["button_execute"])
        if self.undo_changes and not self.is_processing:
            self.undo_button.configure(state="normal")
        else:
            self.undo_button.configure(state="disabled")

    def _set_ready_status(self) -> None:
        count = len(self._changed_items())
        if count:
            self.status_var.set(UI_TEXT["status_changes"].format(count=count))
        elif self.undo_changes:
            self.status_var.set(UI_TEXT["status_undo_ready"])
        elif self.items:
            self.status_var.set(UI_TEXT["status_ready"])
        else:
            self.status_var.set(UI_TEXT["status_idle"])

    def _validate_changes(self, items: list[PdfItem]) -> list[tuple[Path, Path]]:
        if self.current_folder is None:
            return []

        for item in items:
            validate_stem(item.proposed_stem)

        destinations: list[tuple[Path, Path]] = []
        target_keys: set[str] = set()
        source_keys = {normalize_case_name(item.path.name) for item in items}

        try:
            existing_names = {normalize_case_name(p.name): p for p in self.current_folder.iterdir() if p.is_file()}
        except Exception as exc:
            raise RenameValidationError(UI_TEXT["error_folder_missing"]) from exc

        for item in items:
            dest = self.current_folder / f"{item.proposed_stem}.pdf"
            key = normalize_case_name(dest.name)
            if key in target_keys:
                raise RenameValidationError(f"{UI_TEXT['error_duplicate_name']}\n{dest.name}")
            target_keys.add(key)

            existing = existing_names.get(key)
            if existing is not None and key not in source_keys:
                raise RenameValidationError(f"{UI_TEXT['error_existing_name']}\n{existing.name}")
            destinations.append((item.path, dest))

        return destinations

    def apply_renames(self) -> None:
        if self.is_processing:
            return
        items = self._changed_items()
        if not items:
            return

        try:
            changes = self._validate_changes(items)
        except RenameValidationError as exc:
            messagebox.showerror(UI_TEXT["dialog_error_title"], str(exc))
            return

        if not messagebox.askyesno(
            UI_TEXT["dialog_warning_title"],
            UI_TEXT["confirm_rename"].format(count=len(changes)),
        ):
            return

        self.is_processing = True
        self.status_var.set(UI_TEXT["status_processing"])
        self._update_actions()
        self.root.update_idletasks()

        try:
            execute_two_phase_rename(changes)
        except Exception:
            self.is_processing = False
            self.status_var.set(UI_TEXT["status_error"])
            self._update_actions()
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["error_rename_failed"])
            return

        self.undo_changes = [(dest, source) for source, dest in changes]
        self.is_processing = False
        count = len(changes)
        messagebox.showinfo(
            UI_TEXT["dialog_complete_title"],
            UI_TEXT["complete_rename"].format(count=count),
        )
        if self.current_folder is not None:
            self.load_folder_after_rename(self.current_folder, self.undo_changes)

    def load_folder_after_rename(self, folder: Path, undo_changes: list[tuple[Path, Path]]) -> None:
        saved_undo = list(undo_changes)
        self.load_folder(folder)
        self.undo_changes = saved_undo
        self._update_actions()
        self.status_var.set(UI_TEXT["status_undo_ready"])

    def undo_last(self) -> None:
        if self.is_processing or not self.undo_changes:
            return
        changes = list(self.undo_changes)

        for source, dest in changes:
            if not source.exists():
                messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["error_undo_failed"])
                return
            if dest.exists() and normalize_case_name(dest.name) != normalize_case_name(source.name):
                messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["error_undo_failed"])
                return

        if not messagebox.askyesno(
            UI_TEXT["dialog_undo_title"],
            UI_TEXT["confirm_undo"].format(count=len(changes)),
        ):
            return

        self.is_processing = True
        self.status_var.set(UI_TEXT["status_processing"])
        self._update_actions()
        self.root.update_idletasks()

        try:
            execute_two_phase_rename(changes)
        except Exception:
            self.is_processing = False
            self.status_var.set(UI_TEXT["status_error"])
            self._update_actions()
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["error_undo_failed"])
            return

        count = len(changes)
        self.undo_changes = []
        self.is_processing = False
        messagebox.showinfo(
            UI_TEXT["dialog_complete_title"],
            UI_TEXT["complete_undo"].format(count=count),
        )
        if self.current_folder is not None:
            self.load_folder(self.current_folder)

    def open_preview(self, item: PdfItem) -> None:
        if pdfium is None or Image is None or ImageTk is None:
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["error_dependency"])
            return

        preview = tk.Toplevel(self.root)
        preview.title(UI_TEXT["preview_title"])
        preview.geometry("780x760")
        preview.minsize(560, 500)
        preview.configure(bg=THEME["background"])
        apply_window_icon(preview)
        preview.transient(self.root)

        top = tk.Frame(preview, bg=THEME["background"])
        top.pack(fill="x", padx=18, pady=(14, 8))
        tk.Label(
            top,
            text=item.path.name,
            font=(self.font_name, 11, "bold"),
            bg=THEME["background"],
            fg=THEME["text"],
            anchor="w",
        ).pack(side="left", fill="x", expand=True)
        close_btn = self._secondary_button(top, UI_TEXT["preview_close"], preview.destroy)
        close_btn.pack(side="right")

        image_label = tk.Label(
            preview,
            text=UI_TEXT["preview_loading"],
            font=(self.font_name, 10),
            bg=THEME["card"],
            fg=THEME["muted"],
            relief="solid",
            bd=1,
        )
        image_label.pack(fill="both", expand=True, padx=18, pady=(0, 8))

        tk.Label(
            preview,
            text=UI_TEXT["preview_hint"],
            font=(self.font_name, 8),
            bg=THEME["background"],
            fg=THEME["muted"],
        ).pack(pady=(0, 10))
        preview.bind("<Escape>", lambda _event: preview.destroy())

        def worker() -> None:
            try:
                _count, image = render_first_page(item.path, 700, 620)
            except Exception:
                image = None

            def apply() -> None:
                if not preview.winfo_exists():
                    return
                if image is None:
                    image_label.configure(text=UI_TEXT["thumbnail_unavailable"], fg=THEME["error"])
                    return
                photo = ImageTk.PhotoImage(image)
                image_label.configure(image=photo, text="")
                image_label.image = photo

            try:
                self.root.after(0, apply)
            except Exception:
                pass

        threading.Thread(target=worker, daemon=True).start()

    def _on_close(self) -> None:
        try:
            self.executor.shutdown(wait=False, cancel_futures=True)
        except Exception:
            pass
        self.root.destroy()


def main() -> None:
    set_windows_app_id()
    root = tk.Tk()
    DakePdfOverviewRenameApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
