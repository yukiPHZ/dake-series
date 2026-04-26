# -*- coding: utf-8 -*-
import json
import os
import queue
import subprocess
import sys
import threading
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import ctypes
except Exception:
    ctypes = None

from pypdf import PdfReader, PdfWriter

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None

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

UI_TEXT = {
    "status_reordered": "並び順を変更しました",
    "drag_hint": "ドラッグして順番を入れ替えできます",
}

MAIN_TITLE = "PDFを結合する"
MAIN_SUBTITLE = "複数のPDFを追加して、そのまま1つにまとめます。"
FOOTER_SERIES_TEXT = "シンプルそれDAKEシリーズ"
FOOTER_SERIES_COPY = " / 止まらない、迷わない、すぐ終わる。"

BUTTON_ADD = "PDFを追加"
BUTTON_FOLDER = "保存先を選ぶ"
BUTTON_REFRESH = "リフレッシュ"
BUTTON_CANCEL = "キャンセル"
BUTTON_MERGE = "結合して保存"

ROW2_GUIDE = UI_TEXT["drag_hint"]

LABEL_SAVE_FOLDER = "保存先"
LABEL_PAGE_COUNT_UNKNOWN = "ページ数を読み込み中"
LABEL_PAGE_SUFFIX = "ページ"
LABEL_LOADING_THUMBNAIL = "サムネイル\n読み込み中"

EMPTY_TITLE_DEFAULT = "PDFを追加してください"
EMPTY_TITLE_DROP = "PDFをドロップしてください"
EMPTY_SUBTITLE = "ドラッグ＆ドロップ または クリックして追加"

STATUS_LOADING = "読み込み中"
STATUS_PROCESSING = "処理中"
STATUS_SAVING = "保存中"
STATUS_READY = "準備完了"
STATUS_NONE = "未選択"
STATUS_CANCELING = "キャンセル中"
STATUS_CANCELED = "キャンセル完了"
STATUS_SAVE_DONE = "保存完了"
STATUS_ERROR = "エラー"

DETAIL_READY = "結合して保存できます"
DETAIL_NONE = "PDFを追加してください"
DETAIL_PROCESSING = "PDFを順番どおりに処理しています"
DETAIL_SAVING = "保存ファイルを書き出しています"
DETAIL_SAVE_DONE = "保存フォルダを開きます"
DETAIL_CANCEL = ""
DETAIL_ERROR = "処理中に問題が発生しました"
DETAIL_FILE_ERROR = "PDFの処理中に問題が発生しました"

MSG_REFRESH_BLOCKED = "処理中はリフレッシュできません"
MSG_NO_FILES = "PDFを追加してください"
MSG_SAVE_FOLDER_ERROR_TITLE = "保存先エラー"
MSG_SAVE_FOLDER_ERROR = "保存先フォルダを準備できませんでした。"
MSG_SAVE_DONE = "結合して保存が完了しました"


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


def detect_font_name():
    root = tk.Tk()
    root.withdraw()
    try:
        families = set(root.tk.call("font", "families"))
    finally:
        root.destroy()
    for name in FONT_CANDIDATES:
        if name in families:
            return name
    return "TkDefaultFont"


FONT_NAME = detect_font_name()


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
            ("↑", lambda: self.app.move_file(self.index, -1)),
            ("↓", lambda: self.app.move_file(self.index, 1)),
            ("削除", lambda: self.app.remove_file(self.index)),
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
        self.root.geometry("1480x900")
        self.root.minsize(1180, 760)
        self.root.configure(bg=BG)
        apply_window_icon(self.root)

        self.cfg = load_config()
        self.save_folder = self.cfg.get("last_folder", default_downloads())
        self.files: list[str] = []
        self.page_count_cache: dict[str, int] = {}
        self.merge_card_by_path: dict[str, MergeFileCard] = {}
        self.thumbnail_queue: queue.Queue = queue.Queue()
        self.ui_queue: queue.Queue = queue.Queue()
        self.worker_running = False
        self.cancel_requested = False
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

        self.build_ui()
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

        top_card = tk.Frame(main, bg="#FFFFFF", highlightthickness=1, highlightbackground=BORDER)
        top_card.pack(fill="x", pady=(0, 12))

        row1 = tk.Frame(top_card, bg="#FFFFFF")
        row1.pack(fill="x", padx=14, pady=(12, 8))

        self.add_button = ModernButton(row1, BUTTON_ADD, self.add_files, primary=True, width=10)
        self.add_button.pack(side="left")

        self.folder_button = ModernButton(row1, BUTTON_FOLDER, self.choose_folder, width=12)
        self.folder_button.pack(side="left", padx=8)

        self.refresh_button = ModernButton(row1, BUTTON_REFRESH, self.reset_merge, width=10)
        self.refresh_button.pack(side="left", padx=8)

        self.count_label = tk.Label(
            row1,
            text="",
            font=(FONT_NAME, 9, "bold"),
            bg="#FFFFFF",
            fg=TEXT,
        )
        self.count_label.pack(side="left", padx=(16, 0))

        self.action_buttons.extend([self.add_button, self.folder_button, self.refresh_button])

        right_info = tk.Frame(row1, bg="#FFFFFF")
        right_info.pack(side="right")

        self.folder_short_label = tk.Label(
            right_info,
            text=f"{LABEL_SAVE_FOLDER}: {shorten_path(self.save_folder)}",
            font=(FONT_NAME, 9),
            bg="#FFFFFF",
            fg=SUBTEXT,
        )
        self.folder_short_label.pack(side="right")

        row2 = tk.Frame(top_card, bg="#FFFFFF")
        row2.pack(fill="x", padx=14, pady=(0, 12))

        tk.Label(
            row2,
            text=ROW2_GUIDE,
            font=(FONT_NAME, 9),
            bg="#FFFFFF",
            fg=SUBTEXT,
        ).pack(anchor="w")

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

        footer = tk.Frame(shell, bg=BG)
        footer.pack(fill="x", padx=24, pady=(0, 10))

        footer_left = tk.Frame(footer, bg=BG)
        footer_left.pack(side="left")

        footer_right = tk.Frame(footer, bg=BG)
        footer_right.pack(side="right")

        tk.Label(
            footer_left,
            text=f"{FOOTER_SERIES_TEXT}{FOOTER_SERIES_COPY}",
            font=(FONT_NAME, 9),
            bg=BG,
            fg=SUBTEXT,
        ).pack(side="left")

        footer_assessment = tk.Label(
            footer_right,
            text="戸建買取査定",
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
            footer_right,
            text=" ｜ ",
            font=(FONT_NAME, 9),
            bg=BG,
            fg=SUBTEXT,
        ).pack(side="left")

        footer_instagram = tk.Label(
            footer_right,
            text="Instagram",
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
            footer_right,
            text=f" ｜ {COPYRIGHT}",
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

    def set_processing_state(self, processing: bool):
        if processing:
            self.cancel_complete_reset()
        for button in self.action_buttons:
            button.set_enabled(not processing)
        self.cancel_button.set_enabled(processing)
        if not processing:
            self.stop_status_animation()
            self.progress_label.configure(fg=ACCENT)

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
        paths = filedialog.askopenfilenames(filetypes=[("PDFファイル", "*.pdf")])
        self.add_pdf_paths(paths)

    def add_pdf_paths(self, paths):
        if paths:
            self.card_ui_loading = True
            self.pending_card_paths.clear()
            self._card_ui_finish_scheduled = False
            self.set_bottom_status(STATUS_LOADING, "", color=ACCENT)
            self.start_status_animation(STATUS_LOADING)

        added = 0
        for raw in paths:
            path = os.path.abspath(str(raw))
            if not path.lower().endswith(".pdf"):
                continue
            if not os.path.isfile(path):
                continue
            if path in self.files:
                continue
            self.files.append(path)
            added += 1

        if added:
            self.refresh_merge_cards()
        else:
            self.card_ui_loading = False
            self.pending_card_paths.clear()
            self._card_ui_finish_scheduled = False
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
        self.files.clear()
        self.page_count_cache.clear()
        self.progress["value"] = 0
        self.refresh_merge_cards()

    def remove_file(self, index: int):
        if self.worker_running:
            return
        if 0 <= index < len(self.files):
            self.files.pop(index)
            self.refresh_merge_cards()

    def move_file(self, index: int, delta: int):
        if self.worker_running:
            return
        new_index = index + delta
        if 0 <= index < len(self.files) and 0 <= new_index < len(self.files):
            self.files[index], self.files[new_index] = self.files[new_index], self.files[index]
            self.refresh_merge_cards()

    def begin_card_drag(self, index: int, event):
        if self.worker_running or not (0 <= index < len(self.files)):
            return
        self.drag_source_index = index
        self.drag_target_index = index
        self.drag_start_xy = (event.x_root, event.y_root)
        self.drag_started = False

    def update_card_drag(self, event):
        if self.drag_source_index is None or self.worker_running:
            return

        if not self.drag_started and self.drag_start_xy is not None:
            start_x, start_y = self.drag_start_xy
            if abs(event.x_root - start_x) < 6 and abs(event.y_root - start_y) < 6:
                return
            self.drag_started = True

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
        self.drag_source_index = None
        self.drag_target_index = None
        self.drag_start_xy = None
        self.drag_started = False
        for card in self.merge_card_by_path.values():
            card.set_drag_visual()

    def reorder_file(self, source_index: int, target_index: int):
        if self.worker_running:
            return
        if not (0 <= source_index < len(self.files)):
            return

        target_index = max(0, min(target_index, len(self.files) - 1))
        path = self.files.pop(source_index)
        self.files.insert(target_index, path)
        self.refresh_merge_cards()
        self.set_bottom_status(UI_TEXT["status_reordered"], "", color=ACCENT)
        self.schedule_complete_reset()

    def get_page_count(self, path: str) -> int | None:
        if path in self.page_count_cache:
            return self.page_count_cache[path]
        try:
            count = len(PdfReader(path).pages)
            self.page_count_cache[path] = count
            return count
        except Exception:
            return None

    def queue_thumbnail_job(self, path: str):
        def worker():
            payload = None
            page_count = self.get_page_count(path)
            if fitz is not None:
                try:
                    doc = fitz.open(path)
                    page_count = len(doc)
                    page = doc.load_page(0)
                    rect = page.rect
                    target_w, target_h = 160, 220
                    scale = min(target_w / rect.width, target_h / rect.height)
                    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
                    payload = pix.tobytes("ppm")
                    doc.close()
                except Exception:
                    payload = None
            self.thumbnail_queue.put((path, payload, page_count))

        threading.Thread(target=worker, daemon=True).start()

    def process_thumbnail_queue(self):
        try:
            while True:
                path, payload, page_count = self.thumbnail_queue.get_nowait()
                card = self.merge_card_by_path.get(path)
                if card is None:
                    self.pending_card_paths.discard(path)
                    continue
                if page_count is not None:
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
            self.set_top_status(f"{count}件追加済み")
        else:
            self.set_top_status("")
        self.folder_short_label.configure(text=f"{LABEL_SAVE_FOLDER}: {shorten_path(self.save_folder)}")

    def refresh_merge_cards(self):
        for child in self.cards_wrap.winfo_children():
            child.destroy()
        self.merge_card_by_path = {}

        if self.card_ui_loading and self.files:
            self.pending_card_paths = set(self.files)
            self._card_ui_finish_scheduled = False
        else:
            self.pending_card_paths.clear()
            self._card_ui_finish_scheduled = False

        self.refresh_status()
        self.refresh_bottom_status()

        if not self.files:
            self.card_ui_loading = False
            self.pending_card_paths.clear()
            self._card_ui_finish_scheduled = False
            self.show_empty_state()
            return

        self.hide_empty_state()

        width = max(self.canvas.winfo_width() - 24, 320)
        card_outer = 210
        cols = max(1, width // card_outer)
        while cols > 1 and ((cols * card_outer) + 12) > width:
            cols -= 1
        used = cols * card_outer
        extra = max(0, width - used)
        pad_x = max(4, extra // (cols * 2 + 2) if cols else 4)

        for i, path in enumerate(self.files):
            card = MergeFileCard(self.cards_wrap, self, path, i)
            self.merge_card_by_path[path] = card
            card.grid(row=i // cols, column=i % cols, padx=pad_x, pady=8, sticky="n")
            card.update_visual()
            self.queue_thumbnail_job(path)

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
        children = list(self.cards_wrap.winfo_children())
        if not children:
            return
        width = max(self.canvas.winfo_width() - 24, 320)
        card_outer = 210
        cols = max(1, width // card_outer)
        while cols > 1 and ((cols * card_outer) + 12) > width:
            cols -= 1
        used = cols * card_outer
        extra = max(0, width - used)
        pad_x = max(4, extra // (cols * 2 + 2) if cols else 4)
        for i, child in enumerate(children):
            child.grid_forget()
            child.grid(row=i // cols, column=i % cols, padx=pad_x, pady=8, sticky="n")

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
        output = os.path.join(folder, base_name)
        if os.path.exists(output):
            stem, ext = os.path.splitext(base_name)
            n = 1
            while True:
                candidate = os.path.join(folder, f"{stem}_{n:02d}{ext}")
                if not os.path.exists(candidate):
                    output = candidate
                    break
                n += 1

        self.cancel_requested = False
        self.worker_running = True
        self.set_processing_state(True)
        self.progress["value"] = 0
        self.set_bottom_status(STATUS_PROCESSING, DETAIL_PROCESSING, color=ACCENT)
        self.start_status_animation(STATUS_PROCESSING)
        threading.Thread(target=self._merge_worker, args=(output,), daemon=True).start()

    def _merge_worker(self, output: str):
        try:
            writer = PdfWriter()
            total = len(self.files)

            for i, path in enumerate(self.files, start=1):
                if self.cancel_requested:
                    self.enqueue_ui_call(self.finish_cancel)
                    return
                try:
                    reader = PdfReader(path)
                    for page in reader.pages:
                        writer.add_page(page)
                except Exception as e:
                    self.enqueue_ui_call(self.handle_file_error, path, e)
                    return

                progress = int(i / total * 100)
                self.enqueue_ui_call(self.update_progress, progress, "")

            self.enqueue_ui_call(
                self.set_bottom_status,
                STATUS_SAVING,
                DETAIL_SAVING,
                color=ACCENT,
            )
            self.enqueue_ui_call(self.start_status_animation, STATUS_SAVING)

            with open(output, "wb") as f:
                writer.write(f)

            self.enqueue_ui_call(self.finish_success, output)
        except Exception as e:
            self.enqueue_ui_call(self.finish_error, e)

    def update_progress(self, value: int, text: str):
        self.progress["value"] = value

    def cancel_task(self):
        if self.worker_running:
            self.cancel_requested = True
            self.set_bottom_status(STATUS_CANCELING, DETAIL_CANCEL, color=ACCENT)
            self.start_status_animation(STATUS_CANCELING)

    def finish_cancel(self):
        self.worker_running = False
        self.set_processing_state(False)
        self.progress["value"] = 0
        self.set_bottom_status(STATUS_CANCELED, DETAIL_CANCEL, color=ACCENT)
        self.schedule_complete_reset()

    def finish_success(self, output: str):
        self.worker_running = False
        self.set_processing_state(False)
        self.progress["value"] = 100
        self.set_bottom_status(STATUS_SAVE_DONE, DETAIL_SAVE_DONE, color=SUCCESS)
        self.root.lift()
        self.root.focus_force()
        messagebox.showinfo(APP_NAME, MSG_SAVE_DONE, parent=self.root)
        open_folder(os.path.dirname(output))
        self.schedule_complete_reset()

    def finish_error(self, error: Exception):
        self.worker_running = False
        self.set_processing_state(False)
        self.progress["value"] = 0
        self.refresh_status()
        self.set_bottom_status(STATUS_ERROR, DETAIL_ERROR, color=ERROR)
        messagebox.showerror(
            STATUS_ERROR,
            f"{DETAIL_ERROR}\n{error}",
            parent=self.root,
        )

    def handle_file_error(self, path: str, error: Exception):
        self.worker_running = False
        self.set_processing_state(False)
        self.progress["value"] = 0
        self.refresh_status()
        self.set_bottom_status(STATUS_ERROR, DETAIL_FILE_ERROR, color=ERROR)

        messagebox.showerror(
            STATUS_ERROR,
            f"{DETAIL_FILE_ERROR}\n\n{path}\n\n{error}",
            parent=self.root,
        )


def main():
    set_windows_app_id()
    root = make_root()
    apply_window_icon(root)
    DAKEPDFMergeApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

