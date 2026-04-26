# -*- coding: utf-8 -*-
from __future__ import annotations

import io
import json
import os
import queue
import re
import subprocess
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
    import fitz  # type: ignore
except Exception:
    fitz = None

try:
    from PIL import Image, ImageTk  # type: ignore
except Exception:
    Image = None
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


APP_NAME = "DakePDFページ並べ替え"
WINDOW_TITLE = "PDFページ並べ替え"
INTERNAL_FOLDER_NAME = "DAKE_PDF_Reorder"
EXE_NAME = "DakePDF_Reorder.exe"
CONFIG_FILE_NAME = "dake_pdf_reorder_config.json"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"
COMMON_ICON_RELATIVE = Path("..") / ".." / "02_assets" / "dake_icon.ico"
COMMON_ICON_FILENAME = "dake_icon.ico"

UI_TEXT = {
    "main_title": "PDFのページ順を変える",
    "main_description": "PDFを読み込み、ページを並べ替えて保存します。",
    "empty_title": "PDFを追加してください",
    "empty_subtitle_dnd": "ここにPDFをドラッグ＆ドロップ、または下のボタンから追加",
    "empty_subtitle_button": "下のボタンからPDFを追加",
    "button_add_pdf": "PDFを追加",
    "button_choose_folder": "保存先を選ぶ",
    "button_save": "並べ替えて保存",
    "status_idle": "未読込",
    "status_loading": "読み込み中",
    "status_ready": "準備完了",
    "status_rendering": "サムネイル準備中",
    "status_saving": "保存中",
    "status_complete": "保存完了",
    "status_error": "エラー",
    "status_drop_ready": "PDFを1ファイル追加してください。",
    "status_loading_detail": "PDFを読み込んでいます。",
    "status_rendering_detail": "{count}ページを読み込みました。サムネイルを準備しています。",
    "status_ready_detail": "{count}ページを並べ替えできます。",
    "status_drag_detail": "ページ順を変更しました。",
    "status_saving_detail": "新しいPDFを書き出しています。",
    "status_saving_progress": "{current}/{total}ページを書き出しています。",
    "status_complete_detail": "保存先フォルダを開きます。",
    "save_folder_label": "保存先: {path}",
    "file_label_none": "PDF: 未選択",
    "file_label_value": "PDF: {name} / {count}ページ",
    "thumbnail_loading": "読み込み中",
    "thumbnail_error": "表示できません",
    "thumbnail_order_label": "{number}",
    "thumbnail_source_label": "元P.{page}",
    "dialog_pdf_title": "PDFを選択",
    "dialog_output_title": "保存先を選択",
    "dialog_pdf_filter": "PDFファイル",
    "dialog_error_title": "処理できませんでした",
    "dialog_warning_title": "確認してください",
    "dialog_complete_title": "保存が完了しました",
    "dialog_overwrite_title": "上書き確認",
    "dialog_overwrite_message": "同名ファイルが存在します。上書きしますか？\n\n{path}",
    "message_select_pdf": "PDFファイルを選択してください。",
    "message_pdf_one_file": "PDFは1ファイルだけ指定してください。",
    "message_pdf_not_found": "PDFファイルが見つかりませんでした。",
    "message_pdf_invalid": "PDFを読み込めませんでした。このPDFは処理できない可能性があります。",
    "message_page_failed": "ページの取得に失敗しました。",
    "message_no_pages": "ページが見つかりませんでした。",
    "message_no_pdf": "先にPDFを追加してください。",
    "message_save_folder_missing": "保存先フォルダを選択してください。",
    "message_save_folder_invalid": "保存先フォルダが見つかりませんでした。",
    "message_save_failed": "保存できませんでした。",
    "message_source_overwrite_blocked": "元PDFと同じ場所には保存できません。別の保存先を選んでください。",
    "message_dependency_missing": "必要なライブラリが見つかりませんでした。requirements.txt をインストールしてください。",
    "message_complete": "並べ替えたPDFを保存しました。\n\n保存先:\n{path}",
    "message_open_folder_failed": "保存先フォルダを開けませんでした。手動で確認してください。",
    "message_unknown_error": "原因を特定できませんでした。",
    "message_detail": "詳細: {detail}",
    "output_suffix": "_reordered",
    "footer_series": "シンプルそれDAKEシリーズ",
    "footer_assessment": "戸建買取査定",
    "footer_instagram": "Instagram",
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
    "danger": "#D92D20",
    "success": "#12B76A",
    "white": "#FFFFFF",
    "disabled": "#D8DEE8",
}

LINKS = {
    "assessment": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "instagram": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

WINDOW_SIZE = "980x740"
WINDOW_MIN_SIZE = (900, 660)
CANVAS_PAD_X = 18
CANVAS_PAD_Y = 18
CARD_WIDTH = 154
CARD_HEIGHT = 224
CARD_GAP_X = 16
CARD_GAP_Y = 18
THUMB_BOX_WIDTH = 122
THUMB_BOX_HEIGHT = 158
POLL_INTERVAL_MS = 60


@dataclass
class PageItem:
    original_index: int
    photo: Any = None
    thumbnail_error: bool = False


def make_root() -> tk.Tk:
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
        exe_dir = Path(sys.executable).resolve().parent
        candidates.extend(
            [
                bundled,
                (exe_dir / COMMON_ICON_RELATIVE).resolve(),
                (exe_dir.parent / COMMON_ICON_RELATIVE).resolve(),
            ]
        )
    else:
        candidates.append((Path(__file__).resolve().parent / COMMON_ICON_RELATIVE).resolve())
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


def default_output_dir() -> Path:
    downloads = Path.home() / "Downloads"
    return downloads if downloads.exists() else Path.home()


def shorten_path(path: Path, max_len: int = 68) -> str:
    text = str(path)
    if len(text) <= max_len:
        return text
    drive = path.drive
    parts = path.parts
    tail = Path(*parts[-2:]) if len(parts) >= 2 else path.name
    return f"{drive}\\...\\{tail}" if drive else f"...\\{tail}"


def sanitize_filename(name: str) -> str:
    safe = re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", name).strip(" .")
    return safe or "output"


def format_exception(exc: Exception) -> str:
    detail = str(exc).strip().replace("\n", " ")
    if not detail:
        return UI_TEXT["message_unknown_error"]
    return UI_TEXT["message_detail"].format(detail=detail)


def open_folder(path: Path) -> bool:
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
        return True
    except Exception:
        return False


def open_url(url: str) -> None:
    try:
        webbrowser.open_new_tab(url)
    except Exception:
        pass


class ConfigStore:
    def __init__(self) -> None:
        self.path = app_dir() / CONFIG_FILE_NAME

    def load(self) -> dict[str, str]:
        if not self.path.exists():
            data = {"last_output_dir": str(default_output_dir())}
            self.save(data)
            return data
        try:
            raw = self.path.read_text(encoding="utf-8")
            data = json.loads(raw)
            if isinstance(data, dict):
                return {str(k): str(v) for k, v in data.items()}
        except Exception:
            pass
        data = {"last_output_dir": str(default_output_dir())}
        self.save(data)
        return data

    def save(self, data: dict[str, str]) -> None:
        try:
            self.path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    def load_output_dir(self) -> Path:
        data = self.load()
        candidate = Path(data.get("last_output_dir", ""))
        if candidate.exists() and candidate.is_dir():
            return candidate
        return default_output_dir()

    def save_output_dir(self, output_dir: Path) -> None:
        self.save({"last_output_dir": str(output_dir)})


class PdfReorderApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(*WINDOW_MIN_SIZE)
        self.root.configure(bg=THEME["background"])
        apply_window_icon(self.root)

        self.font_family = self.pick_font_family()
        self.config_store = ConfigStore()
        self.output_dir = self.config_store.load_output_dir()

        self.source_pdf: Path | None = None
        self.pages: list[PageItem] = []
        self.queue: queue.Queue[tuple[Any, ...]] = queue.Queue()
        self.load_id = 0
        self.is_busy = False
        self.drag_original_index: int | None = None
        self.drag_insert_index: int | None = None
        self.card_image_items: dict[int, int] = {}
        self.status_var = tk.StringVar(value=UI_TEXT["status_idle"])
        self.detail_var = tk.StringVar(value=UI_TEXT["status_drop_ready"])
        self.file_var = tk.StringVar(value=UI_TEXT["file_label_none"])
        self.folder_var = tk.StringVar(value=UI_TEXT["save_folder_label"].format(path=shorten_path(self.output_dir)))

        self.setup_styles()
        self.build_ui()
        self.setup_drop_target()
        self.update_buttons()
        self.render_pages()
        self.root.after(POLL_INTERVAL_MS, self.poll_queue)

    def pick_font_family(self) -> str:
        try:
            families = set(tkfont.families(self.root))
        except Exception:
            return "TkDefaultFont"
        for family in ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo"):
            if family in families:
                return family
        return "TkDefaultFont"

    def setup_styles(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
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
        self.root.grid_rowconfigure(1, weight=1)

        header = tk.Frame(self.root, bg=THEME["background"])
        header.grid(row=0, column=0, sticky="ew", padx=24, pady=(22, 12))
        header.grid_columnconfigure(0, weight=1)

        title = tk.Label(
            header,
            text=UI_TEXT["main_title"],
            bg=THEME["background"],
            fg=THEME["text"],
            font=(self.font_family, 20, "bold"),
            anchor="w",
        )
        title.grid(row=0, column=0, sticky="w")

        description = tk.Label(
            header,
            text=UI_TEXT["main_description"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 10),
            anchor="w",
        )
        description.grid(row=1, column=0, sticky="w", pady=(6, 0))

        panel = tk.Frame(
            self.root,
            bg=THEME["card"],
            highlightthickness=1,
            highlightbackground=THEME["border"],
            highlightcolor=THEME["border"],
        )
        panel.grid(row=1, column=0, sticky="nsew", padx=24, pady=(0, 14))
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(1, weight=1)

        info = tk.Frame(panel, bg=THEME["card"])
        info.grid(row=0, column=0, sticky="ew", padx=18, pady=(14, 10))
        info.grid_columnconfigure(0, weight=1)

        file_label = tk.Label(
            info,
            textvariable=self.file_var,
            bg=THEME["card"],
            fg=THEME["text"],
            font=(self.font_family, 10, "bold"),
            anchor="w",
        )
        file_label.grid(row=0, column=0, sticky="ew")

        folder_label = tk.Label(
            info,
            textvariable=self.folder_var,
            bg=THEME["card"],
            fg=THEME["muted"],
            font=(self.font_family, 9),
            anchor="e",
        )
        folder_label.grid(row=0, column=1, sticky="e", padx=(16, 0))

        canvas_wrap = tk.Frame(panel, bg=THEME["card"])
        canvas_wrap.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 16))
        canvas_wrap.grid_columnconfigure(0, weight=1)
        canvas_wrap.grid_rowconfigure(0, weight=1)

        self.canvas = tk.Canvas(
            canvas_wrap,
            bg=THEME["background"],
            highlightthickness=1,
            highlightbackground=THEME["border"],
            highlightcolor=THEME["selection_border"],
            bd=0,
            relief="flat",
        )
        self.canvas.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(
            canvas_wrap,
            orient="vertical",
            command=self.canvas.yview,
            style="Dake.Vertical.TScrollbar",
        )
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.bind("<Configure>", self.on_canvas_configure)
        self.canvas.bind("<ButtonPress-1>", self.on_canvas_press)
        self.canvas.bind("<B1-Motion>", self.on_canvas_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_canvas_release)
        self.canvas.bind("<MouseWheel>", self.on_mouse_wheel)
        self.canvas.bind("<Enter>", lambda _event: self.canvas.focus_set())

        controls = tk.Frame(self.root, bg=THEME["background"])
        controls.grid(row=2, column=0, sticky="ew", padx=24, pady=(0, 14))
        controls.grid_columnconfigure(0, weight=1)

        status_box = tk.Frame(controls, bg=THEME["background"])
        status_box.grid(row=0, column=0, sticky="ew")
        status_box.grid_columnconfigure(1, weight=1)

        self.status_badge = tk.Label(
            status_box,
            textvariable=self.status_var,
            bg=THEME["soft"],
            fg=THEME["muted"],
            font=(self.font_family, 9, "bold"),
            padx=12,
            pady=5,
        )
        self.status_badge.grid(row=0, column=0, sticky="w")

        detail_label = tk.Label(
            status_box,
            textvariable=self.detail_var,
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 9),
            anchor="w",
        )
        detail_label.grid(row=0, column=1, sticky="ew", padx=(12, 0))

        button_bar = tk.Frame(controls, bg=THEME["background"])
        button_bar.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        button_bar.grid_columnconfigure(0, weight=1)

        self.add_button = self.make_button(
            button_bar,
            UI_TEXT["button_add_pdf"],
            self.choose_pdf,
            variant="secondary",
        )
        self.add_button.grid(row=0, column=1, sticky="e", padx=(0, 10))

        self.folder_button = self.make_button(
            button_bar,
            UI_TEXT["button_choose_folder"],
            self.choose_output_dir,
            variant="secondary",
        )
        self.folder_button.grid(row=0, column=2, sticky="e", padx=(0, 10))

        self.save_button = self.make_button(
            button_bar,
            UI_TEXT["button_save"],
            self.save_reordered_pdf,
            variant="primary",
        )
        self.save_button.grid(row=0, column=3, sticky="e")

        footer = tk.Frame(self.root, bg=THEME["background"])
        footer.grid(row=3, column=0, sticky="ew", padx=24, pady=(0, 18))
        footer.grid_columnconfigure(1, weight=1)

        footer_left = tk.Label(
            footer,
            text=UI_TEXT["footer_series"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 8),
            anchor="w",
        )
        footer_left.grid(row=0, column=0, sticky="w")

        footer_right = tk.Frame(footer, bg=THEME["background"])
        footer_right.grid(row=0, column=1, sticky="e")

        self.make_footer_link(footer_right, UI_TEXT["footer_assessment"], LINKS["assessment"]).pack(side="left")
        tk.Label(
            footer_right,
            text="  /  ",
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 8),
        ).pack(side="left")
        self.make_footer_link(footer_right, UI_TEXT["footer_instagram"], LINKS["instagram"]).pack(side="left")
        tk.Label(
            footer_right,
            text=f"  /  {UI_TEXT['footer_copyright']}",
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 8),
        ).pack(side="left")

    def make_button(self, parent: tk.Misc, text: str, command: Any, variant: str) -> tk.Button:
        is_primary = variant == "primary"
        bg = THEME["accent"] if is_primary else THEME["white"]
        fg = THEME["white"] if is_primary else THEME["text"]
        active_bg = THEME["accent_hover"] if is_primary else THEME["selection_bg"]
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg=fg,
            activebackground=active_bg,
            activeforeground=fg,
            disabledforeground=THEME["muted"],
            font=(self.font_family, 10, "bold"),
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=THEME["accent"] if is_primary else THEME["border"],
            highlightcolor=THEME["accent"] if is_primary else THEME["border"],
            padx=18,
            pady=9,
            cursor="hand2",
        )

    def make_footer_link(self, parent: tk.Misc, text: str, url: str) -> tk.Label:
        label = tk.Label(
            parent,
            text=text,
            bg=THEME["background"],
            fg=THEME["accent"],
            font=(self.font_family, 8, "underline"),
            cursor="hand2",
        )
        label.bind("<Button-1>", lambda _event: open_url(url))
        return label

    def setup_drop_target(self) -> None:
        if not DND_ENABLED or DND_FILES is None:
            return
        for widget in (self.root, self.canvas):
            try:
                widget.drop_target_register(DND_FILES)  # type: ignore[attr-defined]
                widget.dnd_bind("<<Drop>>", self.on_drop)  # type: ignore[attr-defined]
            except Exception:
                pass

    def set_status(self, status_key: str, detail: str | None = None) -> None:
        self.status_var.set(UI_TEXT[status_key])
        if detail is not None:
            self.detail_var.set(detail)
        if status_key in ("status_ready", "status_rendering"):
            self.status_badge.configure(bg=THEME["selection_bg"], fg=THEME["accent"])
        elif status_key == "status_saving":
            self.status_badge.configure(bg=THEME["selection_bg"], fg=THEME["accent"])
        elif status_key == "status_complete":
            self.status_badge.configure(bg="#EAFBF3", fg=THEME["success"])
        elif status_key == "status_error":
            self.status_badge.configure(bg="#FDECEC", fg=THEME["danger"])
        else:
            self.status_badge.configure(bg=THEME["soft"], fg=THEME["muted"])

    def update_buttons(self) -> None:
        normal = "normal"
        disabled = "disabled"
        self.add_button.configure(state=disabled if self.is_busy else normal)
        self.folder_button.configure(state=disabled if self.is_busy else normal)
        can_save = self.source_pdf is not None and bool(self.pages) and not self.is_busy
        self.save_button.configure(state=normal if can_save else disabled)

    def on_canvas_configure(self, _event: tk.Event) -> None:
        self.render_pages()

    def on_mouse_wheel(self, event: tk.Event) -> None:
        if self.pages:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def choose_pdf(self) -> None:
        if self.is_busy:
            return
        path = filedialog.askopenfilename(
            title=UI_TEXT["dialog_pdf_title"],
            filetypes=[(UI_TEXT["dialog_pdf_filter"], "*.pdf")],
        )
        if path:
            self.load_pdf(Path(path))

    def choose_output_dir(self) -> None:
        if self.is_busy:
            return
        selected = filedialog.askdirectory(
            title=UI_TEXT["dialog_output_title"],
            initialdir=str(self.output_dir),
            mustexist=True,
        )
        if selected:
            self.output_dir = Path(selected)
            self.config_store.save_output_dir(self.output_dir)
            self.folder_var.set(UI_TEXT["save_folder_label"].format(path=shorten_path(self.output_dir)))

    def on_drop(self, event: tk.Event) -> None:
        if self.is_busy:
            return
        try:
            paths = [Path(value) for value in self.root.tk.splitlist(event.data)]  # type: ignore[attr-defined]
        except Exception:
            self.show_warning(UI_TEXT["message_select_pdf"])
            return
        if len(paths) != 1:
            self.show_warning(UI_TEXT["message_pdf_one_file"])
            return
        self.load_pdf(paths[0])

    def load_pdf(self, pdf_path: Path) -> None:
        if pdf_path.suffix.lower() != ".pdf":
            self.show_warning(UI_TEXT["message_select_pdf"])
            return
        if not pdf_path.exists():
            self.show_error(UI_TEXT["message_pdf_not_found"])
            return
        if fitz is None or Image is None or ImageTk is None:
            self.show_error(UI_TEXT["message_dependency_missing"])
            return

        self.load_id += 1
        current_load_id = self.load_id
        self.source_pdf = pdf_path
        self.pages = []
        self.card_image_items.clear()
        self.file_var.set(UI_TEXT["file_label_none"])
        self.is_busy = True
        self.drag_original_index = None
        self.drag_insert_index = None
        self.set_status("status_loading", UI_TEXT["status_loading_detail"])
        self.update_buttons()
        self.render_pages()

        thread = threading.Thread(
            target=self.load_pdf_worker,
            args=(current_load_id, pdf_path),
            daemon=True,
        )
        thread.start()

    def load_pdf_worker(self, current_load_id: int, pdf_path: Path) -> None:
        try:
            with fitz.open(str(pdf_path)) as document:  # type: ignore[union-attr]
                page_count = document.page_count
                if page_count < 1:
                    raise ValueError(UI_TEXT["message_no_pages"])
                self.queue.put(("load_ready", current_load_id, pdf_path, page_count))
                for index in range(page_count):
                    if current_load_id != self.load_id:
                        return
                    try:
                        page = document.load_page(index)
                        rect = page.rect
                        zoom = min(THUMB_BOX_WIDTH / rect.width, THUMB_BOX_HEIGHT / rect.height, 0.55)
                        zoom = max(zoom, 0.12)
                        pixmap = page.get_pixmap(
                            matrix=fitz.Matrix(zoom, zoom),  # type: ignore[union-attr]
                            alpha=False,
                            colorspace=fitz.csRGB,  # type: ignore[union-attr]
                        )
                        data = pixmap.tobytes("png")
                        self.queue.put(("thumbnail", current_load_id, index, data))
                    except Exception:
                        self.queue.put(("thumbnail_error", current_load_id, index))
                self.queue.put(("thumbnail_done", current_load_id))
        except Exception as exc:
            self.queue.put(("load_error", current_load_id, exc))

    def poll_queue(self) -> None:
        while True:
            try:
                message = self.queue.get_nowait()
            except queue.Empty:
                break
            self.handle_queue_message(message)
        self.root.after(POLL_INTERVAL_MS, self.poll_queue)

    def handle_queue_message(self, message: tuple[Any, ...]) -> None:
        kind = message[0]
        if kind in {"load_ready", "thumbnail", "thumbnail_error", "thumbnail_done", "load_error"}:
            current_load_id = message[1]
            if current_load_id != self.load_id:
                return

        if kind == "load_ready":
            _kind, _load_id, pdf_path, page_count = message
            self.source_pdf = pdf_path
            self.pages = [PageItem(original_index=index) for index in range(page_count)]
            self.file_var.set(UI_TEXT["file_label_value"].format(name=pdf_path.name, count=page_count))
            self.is_busy = False
            self.set_status(
                "status_rendering",
                UI_TEXT["status_rendering_detail"].format(count=page_count),
            )
            self.update_buttons()
            self.render_pages()
            return

        if kind == "thumbnail":
            _kind, _load_id, page_index, data = message
            self.apply_thumbnail(page_index, data)
            return

        if kind == "thumbnail_error":
            _kind, _load_id, page_index = message
            self.mark_thumbnail_error(page_index)
            return

        if kind == "thumbnail_done":
            if self.source_pdf is not None:
                self.set_status(
                    "status_ready",
                    UI_TEXT["status_ready_detail"].format(count=len(self.pages)),
                )
            return

        if kind == "load_error":
            _kind, _load_id, exc = message
            self.is_busy = False
            self.pages = []
            self.source_pdf = None
            self.file_var.set(UI_TEXT["file_label_none"])
            self.set_status("status_error", UI_TEXT["message_pdf_invalid"])
            self.update_buttons()
            self.render_pages()
            self.show_error(f"{UI_TEXT['message_pdf_invalid']}\n{format_exception(exc)}")
            return

        if kind == "save_progress":
            _kind, current, total = message
            self.set_status("status_saving", UI_TEXT["status_saving_progress"].format(current=current, total=total))
            return

        if kind == "save_done":
            _kind, output_path = message
            self.is_busy = False
            self.set_status("status_complete", UI_TEXT["status_complete_detail"])
            self.update_buttons()
            messagebox.showinfo(
                UI_TEXT["dialog_complete_title"],
                UI_TEXT["message_complete"].format(path=output_path),
            )
            if not open_folder(Path(output_path).parent):
                self.show_warning(UI_TEXT["message_open_folder_failed"])
            return

        if kind == "save_error":
            _kind, exc = message
            self.is_busy = False
            self.set_status("status_error", UI_TEXT["message_save_failed"])
            self.update_buttons()
            self.show_error(f"{UI_TEXT['message_save_failed']}\n{format_exception(exc)}")

    def apply_thumbnail(self, page_index: int, data: bytes) -> None:
        page = self.find_page_by_original_index(page_index)
        if page is None or Image is None or ImageTk is None:
            return
        try:
            image = Image.open(io.BytesIO(data))
            image.thumbnail((THUMB_BOX_WIDTH, THUMB_BOX_HEIGHT))
            page.photo = ImageTk.PhotoImage(image)
            page.thumbnail_error = False
        except Exception:
            page.thumbnail_error = True
            page.photo = None
        self.update_card_thumbnail(page.original_index)

    def mark_thumbnail_error(self, page_index: int) -> None:
        page = self.find_page_by_original_index(page_index)
        if page is None:
            return
        page.thumbnail_error = True
        page.photo = None
        self.update_card_thumbnail(page.original_index)

    def find_page_by_original_index(self, original_index: int) -> PageItem | None:
        for page in self.pages:
            if page.original_index == original_index:
                return page
        return None

    def update_card_thumbnail(self, original_index: int) -> None:
        page = self.find_page_by_original_index(original_index)
        if page is None:
            return
        image_item = self.card_image_items.get(original_index)
        if image_item is None:
            self.render_pages()
            return
        self.canvas.delete(f"placeholder_{original_index}")
        if page.photo is not None:
            self.canvas.itemconfigure(image_item, image=page.photo)
        else:
            self.render_pages()

    def render_pages(self) -> None:
        self.canvas.delete("page_card")
        self.canvas.delete("empty_state")
        self.canvas.delete("insert_marker")
        self.card_image_items.clear()

        width = max(self.canvas.winfo_width(), 400)
        height = max(self.canvas.winfo_height(), 260)
        if not self.pages:
            self.canvas.configure(scrollregion=(0, 0, width, height))
            title_y = height // 2 - 20
            subtitle = UI_TEXT["empty_subtitle_dnd"] if DND_ENABLED else UI_TEXT["empty_subtitle_button"]
            self.canvas.create_text(
                width // 2,
                title_y,
                text=UI_TEXT["empty_title"],
                fill=THEME["text"],
                font=(self.font_family, 14, "bold"),
                tags=("empty_state",),
            )
            self.canvas.create_text(
                width // 2,
                title_y + 32,
                text=subtitle,
                fill=THEME["muted"],
                font=(self.font_family, 10),
                tags=("empty_state",),
            )
            return

        columns = self.column_count(width)
        rows = (len(self.pages) + columns - 1) // columns
        total_height = CANVAS_PAD_Y * 2 + rows * CARD_HEIGHT + max(0, rows - 1) * CARD_GAP_Y
        self.canvas.configure(scrollregion=(0, 0, width, max(height, total_height)))

        for display_index, page in enumerate(self.pages):
            x, y = self.card_position(display_index, columns)
            self.draw_page_card(page, display_index, x, y)

        if self.drag_insert_index is not None:
            self.draw_insert_marker(self.drag_insert_index)

    def draw_page_card(self, page: PageItem, display_index: int, x: int, y: int) -> None:
        original_index = page.original_index
        tag = f"page_{original_index}"
        is_dragging = original_index == self.drag_original_index
        fill = THEME["selection_bg"] if is_dragging else THEME["card"]
        outline = THEME["selection_border"] if is_dragging else THEME["border"]
        width = 2 if is_dragging else 1

        self.canvas.create_rectangle(
            x,
            y,
            x + CARD_WIDTH,
            y + CARD_HEIGHT,
            fill=fill,
            outline=outline,
            width=width,
            tags=("page_card", tag),
        )
        self.canvas.create_rectangle(
            x + 12,
            y + 12,
            x + 12 + THUMB_BOX_WIDTH,
            y + 12 + THUMB_BOX_HEIGHT,
            fill=THEME["background"],
            outline=THEME["border"],
            width=1,
            tags=("page_card", tag),
        )
        image_x = x + 12 + THUMB_BOX_WIDTH // 2
        image_y = y + 12 + THUMB_BOX_HEIGHT // 2
        if page.photo is not None:
            image_item = self.canvas.create_image(
                image_x,
                image_y,
                image=page.photo,
                anchor="center",
                tags=("page_card", tag),
            )
        else:
            image_item = self.canvas.create_image(
                image_x,
                image_y,
                anchor="center",
                tags=("page_card", tag),
            )
        self.card_image_items[original_index] = image_item
        if page.photo is None:
            placeholder_text = UI_TEXT["thumbnail_error"] if page.thumbnail_error else UI_TEXT["thumbnail_loading"]
            self.canvas.create_text(
                image_x,
                image_y,
                text=placeholder_text,
                fill=THEME["muted"],
                font=(self.font_family, 9),
                tags=("page_card", tag, f"placeholder_{original_index}"),
            )

        badge_size = 30
        self.canvas.create_rectangle(
            x + 12,
            y + 12,
            x + 12 + badge_size,
            y + 12 + badge_size,
            fill=THEME["accent"],
            outline=THEME["accent"],
            tags=("page_card", tag),
        )
        self.canvas.create_text(
            x + 12 + badge_size // 2,
            y + 12 + badge_size // 2,
            text=UI_TEXT["thumbnail_order_label"].format(number=display_index + 1),
            fill=THEME["white"],
            font=(self.font_family, 10, "bold"),
            tags=("page_card", tag),
        )
        self.canvas.create_text(
            x + CARD_WIDTH // 2,
            y + CARD_HEIGHT - 26,
            text=UI_TEXT["thumbnail_source_label"].format(page=original_index + 1),
            fill=THEME["muted"],
            font=(self.font_family, 9, "bold"),
            tags=("page_card", tag),
        )

    def column_count(self, width: int) -> int:
        available = max(width - CANVAS_PAD_X * 2, CARD_WIDTH)
        return max(1, (available + CARD_GAP_X) // (CARD_WIDTH + CARD_GAP_X))

    def card_position(self, index: int, columns: int | None = None) -> tuple[int, int]:
        if columns is None:
            columns = self.column_count(max(self.canvas.winfo_width(), 400))
        row = index // columns
        col = index % columns
        x = CANVAS_PAD_X + col * (CARD_WIDTH + CARD_GAP_X)
        y = CANVAS_PAD_Y + row * (CARD_HEIGHT + CARD_GAP_Y)
        return x, y

    def page_index_from_event(self, event: tk.Event) -> int | None:
        current = self.canvas.find_withtag("current")
        if not current:
            return None
        tags = self.canvas.gettags(current[0])
        for tag in tags:
            if tag.startswith("page_"):
                try:
                    original_index = int(tag.split("_", 1)[1])
                except ValueError:
                    return None
                for index, page in enumerate(self.pages):
                    if page.original_index == original_index:
                        return index
        return None

    def on_canvas_press(self, event: tk.Event) -> None:
        if self.is_busy or not self.pages:
            return
        index = self.page_index_from_event(event)
        if index is None:
            return
        self.drag_original_index = self.pages[index].original_index
        self.drag_insert_index = index
        self.render_pages()

    def on_canvas_drag(self, event: tk.Event) -> None:
        if self.drag_original_index is None:
            return
        x = int(self.canvas.canvasx(event.x))
        y = int(self.canvas.canvasy(event.y))
        insert_index = self.insert_index_from_xy(x, y)
        if insert_index != self.drag_insert_index:
            self.drag_insert_index = insert_index
            self.render_pages()
        self.autoscroll(event.y)

    def on_canvas_release(self, _event: tk.Event) -> None:
        if self.drag_original_index is None:
            return
        original_index = self.drag_original_index
        insert_index = self.drag_insert_index
        self.drag_original_index = None
        self.drag_insert_index = None

        old_index = None
        for index, page in enumerate(self.pages):
            if page.original_index == original_index:
                old_index = index
                break
        if old_index is None or insert_index is None:
            self.render_pages()
            return

        insert_index = max(0, min(insert_index, len(self.pages)))
        page = self.pages.pop(old_index)
        if insert_index > old_index:
            insert_index -= 1
        insert_index = max(0, min(insert_index, len(self.pages)))
        self.pages.insert(insert_index, page)
        self.set_status("status_ready", UI_TEXT["status_drag_detail"])
        self.render_pages()

    def insert_index_from_xy(self, x: int, y: int) -> int:
        columns = self.column_count(max(self.canvas.winfo_width(), 400))
        row_height = CARD_HEIGHT + CARD_GAP_Y
        col_width = CARD_WIDTH + CARD_GAP_X
        row = max(0, (y - CANVAS_PAD_Y) // row_height)
        raw_col = max(0, (x - CANVAS_PAD_X) // col_width)
        col = min(columns - 1, raw_col)
        within_col = (x - CANVAS_PAD_X) - col * col_width
        if within_col > CARD_WIDTH // 2:
            col += 1
        index = row * columns + col
        return max(0, min(index, len(self.pages)))

    def draw_insert_marker(self, index: int) -> None:
        columns = self.column_count(max(self.canvas.winfo_width(), 400))
        x, y = self.card_position(index, columns)
        if index >= len(self.pages) and len(self.pages) > 0:
            last_x, last_y = self.card_position(len(self.pages) - 1, columns)
            last_col = (len(self.pages) - 1) % columns
            if last_col < columns - 1:
                x = last_x + CARD_WIDTH + CARD_GAP_X // 2
                y = last_y
            else:
                x = CANVAS_PAD_X
                y = last_y + CARD_HEIGHT + CARD_GAP_Y
        self.canvas.create_line(
            x,
            y,
            x,
            y + CARD_HEIGHT,
            fill=THEME["accent"],
            width=3,
            tags=("insert_marker",),
        )

    def autoscroll(self, pointer_y: int) -> None:
        height = self.canvas.winfo_height()
        if pointer_y < 32:
            self.canvas.yview_scroll(-1, "units")
        elif pointer_y > height - 32:
            self.canvas.yview_scroll(1, "units")

    def save_reordered_pdf(self) -> None:
        if self.source_pdf is None or not self.pages:
            self.show_warning(UI_TEXT["message_no_pdf"])
            return
        if PdfReader is None or PdfWriter is None:
            self.show_error(UI_TEXT["message_dependency_missing"])
            return
        if not self.output_dir.exists() or not self.output_dir.is_dir():
            self.show_warning(UI_TEXT["message_save_folder_invalid"])
            return

        output_name = f"{sanitize_filename(self.source_pdf.stem)}{UI_TEXT['output_suffix']}.pdf"
        output_path = self.output_dir / output_name
        try:
            if self.source_pdf.resolve() == output_path.resolve():
                self.show_error(UI_TEXT["message_source_overwrite_blocked"])
                return
        except Exception:
            pass

        if output_path.exists():
            overwrite = messagebox.askyesno(
                UI_TEXT["dialog_overwrite_title"],
                UI_TEXT["dialog_overwrite_message"].format(path=output_path),
            )
            if not overwrite:
                return

        self.config_store.save_output_dir(self.output_dir)
        order = [page.original_index for page in self.pages]
        self.is_busy = True
        self.set_status("status_saving", UI_TEXT["status_saving_detail"])
        self.update_buttons()
        thread = threading.Thread(
            target=self.save_pdf_worker,
            args=(self.source_pdf, output_path, order),
            daemon=True,
        )
        thread.start()

    def save_pdf_worker(self, source_pdf: Path, output_path: Path, order: list[int]) -> None:
        temp_path: Path | None = None
        try:
            reader = PdfReader(str(source_pdf))  # type: ignore[misc]
            if getattr(reader, "is_encrypted", False):
                try:
                    reader.decrypt("")
                except Exception:
                    pass
                if getattr(reader, "is_encrypted", False):
                    raise ValueError(UI_TEXT["message_pdf_invalid"])

            writer = PdfWriter()  # type: ignore[misc]
            total = len(order)
            for current, page_index in enumerate(order, start=1):
                writer.add_page(reader.pages[page_index])
                if current == 1 or current == total or current % 10 == 0:
                    self.queue.put(("save_progress", current, total))

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
            self.queue.put(("save_done", str(output_path)))
        except Exception as exc:
            self.queue.put(("save_error", exc))
        finally:
            if temp_path is not None:
                try:
                    temp_path.unlink(missing_ok=True)
                except Exception:
                    pass

    def show_warning(self, message: str) -> None:
        messagebox.showwarning(UI_TEXT["dialog_warning_title"], message)

    def show_error(self, message: str) -> None:
        messagebox.showerror(UI_TEXT["dialog_error_title"], message)


def main() -> None:
    root = make_root()
    PdfReorderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
