# -*- coding: utf-8 -*-
from __future__ import annotations

import io
import os
import queue
import re
import subprocess
import sys
import threading
import time
import webbrowser
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox

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
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore

    DND_ENABLED = True
except Exception:
    DND_FILES = None
    TkinterDnD = None
    DND_ENABLED = False


APP_NAME = "DakePDFトリミング"
WINDOW_TITLE = "DakePDFトリミング"
DISPLAY_NAME = "PDFトリミング"
INTERNAL_FOLDER_NAME = "DAKE_PDF_Crop"
EXE_NAME = "DakePDF_Crop.exe"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"
COMMON_ICON_RELATIVE = Path("..") / ".." / "02_assets" / "dake_icon.ico"
COMMON_ICON_FILENAME = "dake_icon.ico"

UI_TEXT = {
    "brand_series": "シンプルそれDAKEシリーズ",
    "main_title": "PDFをトリミングする",
    "main_description": "ドラッグで残したい範囲を選び、そのまま保存します。",
    "button_select_pdf": "PDFを選ぶ",
    "button_refresh": "リフレッシュ",
    "button_reset": "範囲をリセット",
    "button_execute": "この範囲で保存",
    "button_prev_page": "前へ",
    "button_next_page": "次へ",
    "page_indicator": "{current} / {total}",
    "empty_title": "PDFを追加してください",
    "empty_title_drop": "PDFをドロップしてください",
    "empty_subtitle": "ドラッグ＆ドロップ または クリックして追加",
    "status_idle": "未選択",
    "status_loading": "読み込み中",
    "status_ready": "準備完了",
    "status_saving": "保存中",
    "status_complete": "保存完了",
    "status_error": "エラー",
    "status_loading_1": "読み込み中.",
    "status_loading_2": "読み込み中..",
    "status_loading_3": "読み込み中...",
    "status_saving_1": "保存中.",
    "status_saving_2": "保存中..",
    "status_saving_3": "保存中...",
    "status_phrase_1": "Simple",
    "status_phrase_2": "Simple, fast",
    "status_phrase_3": "Simple, fast, for real work.",
    "status_file_ready": "{name} / {count}ページ",
    "status_page_changed": "ページを切り替えました",
    "status_area_selected": "範囲を指定しました",
    "status_area_reset": "範囲をリセットしました",
    "status_saved_detail": "{name} を保存しました",
    "canvas_loading": "プレビューを準備しています",
    "canvas_no_preview": "プレビューを表示できません",
    "dialog_select_pdf_title": "PDFを選択",
    "dialog_complete_title": "保存完了",
    "dialog_error_title": "エラー",
    "dialog_warning_title": "確認してください",
    "dialog_complete_message": "トリミングしたPDFを保存しました。\n\n保存先:\n{path}",
    "filetype_pdf": "PDFファイル",
    "error_not_pdf": "PDFファイルを選択してください。",
    "error_no_area": "トリミング範囲を指定してください。",
    "error_area_too_small": "もう少し大きい範囲を指定してください。",
    "error_load_failed": "PDFを読み込めませんでした。",
    "error_save_failed": "保存に失敗しました。",
    "error_dependency_missing": "必要なライブラリが見つかりませんでした。requirements.txt をインストールしてください。",
    "error_no_pages": "PDFにページが見つかりませんでした。",
    "error_open_folder_failed": "保存フォルダを開けませんでした。手動で確認してください。",
    "error_unknown_detail": "詳細: {detail}",
    "output_suffix": "_crop",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_subtitle": "止まらない、迷わない、すぐ終わる。",
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
    "success": "#12B76A",
    "danger": "#D92D20",
    "button_disabled": "#D8DEE8",
    "soft": "#EEF2F7",
    "white": "#FFFFFF",
}

LINKS = {
    "link_1": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "link_2": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

WINDOW_SIZE = "920x680"
WINDOW_MIN_SIZE = (820, 600)
CANVAS_MARGIN = 28
EMPTY_CARD_WIDTH = 420
EMPTY_CARD_HEIGHT = 168
MIN_SELECTION_CANVAS = 18
MIN_SELECTION_RATIO = 0.015
QUEUE_POLL_MS = 80
STATUS_ANIMATION_MS = 420
STATUS_PHRASE_DELAY_SECONDS = 1.7


@dataclass(frozen=True)
class PdfPreviewPayload:
    source_path: Path
    page_count: int
    page_index: int
    page_rect: tuple[float, float, float, float]
    image_data: bytes


def make_root() -> tk.Tk:
    if DND_ENABLED and TkinterDnD is not None:
        try:
            return TkinterDnD.Tk()
        except Exception:
            pass
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


def sanitize_filename(stem: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", stem).strip(" .")
    return cleaned or "pdf"


def unique_output_path(source_pdf: Path, page_index: int) -> Path:
    suffix = UI_TEXT["output_suffix"]
    parent = source_pdf.parent
    base = sanitize_filename(source_pdf.stem)
    page_label = f"p{page_index + 1}"
    candidate = parent / f"{base}_{page_label}{suffix}.pdf"
    if not candidate.exists():
        return candidate
    for number in range(2, 1000):
        numbered = parent / f"{base}_{page_label}{suffix}_{number}.pdf"
        if not numbered.exists():
            return numbered
    timestamp = int(time.time())
    return parent / f"{base}_{page_label}{suffix}_{timestamp}.pdf"


def format_exception(exc: Exception) -> str:
    detail = str(exc).strip().replace("\n", " ")
    if not detail:
        return ""
    return UI_TEXT["error_unknown_detail"].format(detail=detail)


def crop_single_page_with_ratios(
    source_pdf: Path,
    output_pdf: Path,
    page_index: int,
    ratios: tuple[float, float, float, float],
) -> int:
    if fitz is None:
        raise RuntimeError(UI_TEXT["error_dependency_missing"])

    x0_ratio, y0_ratio, x1_ratio, y1_ratio = ratios
    temp_path = output_pdf.with_name(f".{output_pdf.stem}.tmp.pdf")
    try:
        with fitz.open(str(source_pdf)) as source_document:  # type: ignore[union-attr]
            if source_document.page_count < 1:
                raise ValueError(UI_TEXT["error_no_pages"])
            if page_index < 0 or page_index >= source_document.page_count:
                raise ValueError(UI_TEXT["error_save_failed"])

            output_document = fitz.open()  # type: ignore[union-attr]
            try:
                output_document.insert_pdf(source_document, from_page=page_index, to_page=page_index)
                page = output_document.load_page(0)
                base_rect = page.mediabox
                if base_rect.is_empty or base_rect.width <= 0 or base_rect.height <= 0:
                    base_rect = page.rect
                if base_rect.is_empty or base_rect.width <= 0 or base_rect.height <= 0:
                    raise ValueError(UI_TEXT["error_save_failed"])

                crop_rect = fitz.Rect(
                    base_rect.x0 + base_rect.width * x0_ratio,
                    base_rect.y0 + base_rect.height * y0_ratio,
                    base_rect.x0 + base_rect.width * x1_ratio,
                    base_rect.y0 + base_rect.height * y1_ratio,
                )
                crop_rect = crop_rect & base_rect
                if crop_rect.is_empty or crop_rect.width <= 1 or crop_rect.height <= 1:
                    raise ValueError(UI_TEXT["error_save_failed"])
                page.set_cropbox(crop_rect)
                output_document.save(str(temp_path), garbage=4, deflate=True, clean=True)
            finally:
                output_document.close()
        os.replace(str(temp_path), str(output_pdf))
        return 1
    finally:
        try:
            if temp_path.exists():
                temp_path.unlink()
        except Exception:
            pass


class PdfCropApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(*WINDOW_MIN_SIZE)
        self.root.configure(bg=THEME["background"])
        apply_window_icon(self.root)

        self.font_family = self.pick_font_family()
        self.queue: queue.Queue[tuple[Any, ...]] = queue.Queue()
        self.load_id = 0
        self.busy_mode: str | None = None
        self.busy_started_at = 0.0
        self.status_tick = 0

        self.source_pdf: Path | None = None
        self.page_count = 0
        self.current_page_index = 0
        self.pending_ready_status_key = "status_ready"
        self.page_rect: tuple[float, float, float, float] | None = None
        self.preview_source: Any | None = None
        self.preview_photo: Any | None = None
        self.preview_origin = (0.0, 0.0)
        self.preview_size = (0.0, 0.0)
        self.preview_scale = 1.0
        self.selection_pdf: tuple[float, float, float, float] | None = None
        self.drag_start_pdf: tuple[float, float] | None = None
        self.drag_current_pdf: tuple[float, float] | None = None

        self.status_var = tk.StringVar(value=UI_TEXT["status_idle"])
        self.detail_var = tk.StringVar(value="")
        self.page_var = tk.StringVar(value="")

        self.build_ui()
        self.setup_drop_target()
        self.update_buttons()
        self.draw_empty_state()
        self.root.after(QUEUE_POLL_MS, self.poll_queue)
        self.root.after(STATUS_ANIMATION_MS, self.animate_status)

    def pick_font_family(self) -> str:
        try:
            families = set(tkfont.families(self.root))
        except Exception:
            return "TkDefaultFont"
        for family in ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo"):
            if family in families:
                return family
        return "TkDefaultFont"

    def build_ui(self) -> None:
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=1)

        header = tk.Frame(self.root, bg=THEME["background"])
        header.grid(row=0, column=0, sticky="ew", padx=28, pady=(24, 16))
        header.grid_columnconfigure(0, weight=1)

        header_text = tk.Frame(header, bg=THEME["background"])
        header_text.grid(row=0, column=0, sticky="w")

        tk.Label(
            header_text,
            text=UI_TEXT["main_title"],
            bg=THEME["background"],
            fg=THEME["text"],
            font=(self.font_family, 18, "bold"),
            anchor="w",
        ).pack(side="left")

        tk.Label(
            header_text,
            text=UI_TEXT["main_description"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 10),
            anchor="w",
        ).pack(side="left", padx=(18, 0), pady=(3, 0))

        panel = tk.Frame(
            self.root,
            bg=THEME["card"],
            highlightthickness=1,
            highlightbackground=THEME["border"],
            highlightcolor=THEME["border"],
        )
        panel.grid(row=1, column=0, sticky="nsew", padx=28, pady=(0, 16))
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(0, weight=1)

        self.canvas = tk.Canvas(
            panel,
            bg=THEME["background"],
            bd=0,
            highlightthickness=0,
        )
        self.canvas.grid(row=0, column=0, sticky="nsew", padx=16, pady=16)
        self.canvas.bind("<Configure>", self.on_canvas_configure)
        self.canvas.bind("<ButtonPress-1>", self.on_drag_start)
        self.canvas.bind("<B1-Motion>", self.on_drag_move)
        self.canvas.bind("<ButtonRelease-1>", self.on_drag_end)
        self.canvas.bind("<Motion>", self.on_canvas_motion)
        self.canvas.bind("<Leave>", self.on_canvas_leave)
        self.canvas.bind("<Double-Button-1>", lambda _event: self.choose_pdf())

        controls = tk.Frame(self.root, bg=THEME["background"])
        controls.grid(row=2, column=0, sticky="ew", padx=28, pady=(0, 16))
        controls.grid_columnconfigure(0, weight=1)

        status_area = tk.Frame(controls, bg=THEME["background"])
        status_area.grid(row=0, column=0, sticky="w")

        self.status_label = tk.Label(
            status_area,
            textvariable=self.status_var,
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 10, "bold"),
            anchor="w",
        )
        self.status_label.pack(side="left")

        self.detail_label = tk.Label(
            status_area,
            textvariable=self.detail_var,
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 9),
            anchor="w",
        )
        self.detail_label.pack(side="left", padx=(12, 0))

        page_area = tk.Frame(controls, bg=THEME["background"])
        page_area.grid(row=0, column=1, sticky="e", padx=(16, 12))

        self.prev_button = self.make_button(
            page_area,
            UI_TEXT["button_prev_page"],
            lambda: self.change_page(-1),
            variant="secondary",
        )
        self.prev_button.pack(side="left")

        self.page_label = tk.Label(
            page_area,
            textvariable=self.page_var,
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 10, "bold"),
            anchor="center",
            width=8,
        )
        self.page_label.pack(side="left", padx=8)

        self.next_button = self.make_button(
            page_area,
            UI_TEXT["button_next_page"],
            lambda: self.change_page(1),
            variant="secondary",
        )
        self.next_button.pack(side="left")

        button_area = tk.Frame(controls, bg=THEME["background"])
        button_area.grid(row=0, column=2, sticky="e")

        self.select_button = self.make_button(
            button_area,
            UI_TEXT["button_select_pdf"],
            self.choose_pdf,
            variant="secondary",
        )
        self.select_button.pack(side="left", padx=(0, 10))

        self.refresh_button = self.make_button(
            button_area,
            UI_TEXT["button_refresh"],
            self.refresh_pdf,
            variant="secondary",
        )
        self.refresh_button.pack(side="left", padx=(0, 10))

        self.reset_button = self.make_button(
            button_area,
            UI_TEXT["button_reset"],
            self.reset_selection,
            variant="secondary",
        )
        self.reset_button.pack(side="left", padx=(0, 10))

        self.save_button = self.make_button(
            button_area,
            UI_TEXT["button_execute"],
            self.save_crop,
            variant="primary",
        )
        self.save_button.pack(side="left")

        footer = tk.Frame(self.root, bg=THEME["background"])
        footer.grid(row=3, column=0, sticky="ew", padx=28, pady=(0, 18))
        footer.grid_columnconfigure(1, weight=1)

        left_block = tk.Frame(footer, bg=THEME["background"])
        left_block.grid(row=0, column=0, sticky="w")
        tk.Label(
            left_block,
            text=f"{UI_TEXT['footer_left']} / {UI_TEXT['footer_subtitle']}",
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 8),
            anchor="w",
        ).pack(side="left")

        right_block = tk.Frame(footer, bg=THEME["background"])
        right_block.grid(row=0, column=1, sticky="e")
        self.make_footer_link(right_block, UI_TEXT["footer_link_1"], LINKS["link_1"]).pack(side="left")
        self.make_footer_text(right_block, UI_TEXT["footer_separator"]).pack(side="left")
        self.make_footer_link(right_block, UI_TEXT["footer_link_2"], LINKS["link_2"]).pack(side="left")
        self.make_footer_text(right_block, UI_TEXT["footer_separator"]).pack(side="left")
        self.make_footer_text(right_block, UI_TEXT["footer_copyright"]).pack(side="left")

    def make_button(self, parent: tk.Misc, text: str, command: Any, variant: str) -> tk.Button:
        is_primary = variant == "primary"
        bg = THEME["accent"] if is_primary else THEME["white"]
        fg = THEME["white"] if is_primary else THEME["text"]
        active_bg = THEME["accent_hover"] if is_primary else THEME["selection_bg"]
        button = tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg=fg,
            activebackground=active_bg,
            activeforeground=fg,
            disabledforeground=THEME["muted"],
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=THEME["accent"] if is_primary else THEME["border"],
            highlightcolor=THEME["accent"] if is_primary else THEME["border"],
            padx=18,
            pady=8,
            cursor="hand2",
            font=(self.font_family, 10, "bold"),
        )
        button.bind("<Enter>", lambda _event, target=button, primary=is_primary: self.on_button_hover(target, primary, True))
        button.bind("<Leave>", lambda _event, target=button, primary=is_primary: self.on_button_hover(target, primary, False))
        return button

    def on_button_hover(self, button: tk.Button, is_primary: bool, hover: bool) -> None:
        if str(button.cget("state")) == "disabled":
            return
        if is_primary:
            button.configure(bg=THEME["accent_hover"] if hover else THEME["accent"])
        else:
            button.configure(bg=THEME["selection_bg"] if hover else THEME["white"])

    def make_footer_text(self, parent: tk.Misc, text: str) -> tk.Label:
        return tk.Label(
            parent,
            text=text,
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 8),
        )

    def make_footer_link(self, parent: tk.Misc, text: str, url: str) -> tk.Label:
        label = tk.Label(
            parent,
            text=text,
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 8),
            cursor="hand2",
        )
        label.bind("<Button-1>", lambda _event: open_url(url))
        label.bind("<Enter>", lambda _event: label.configure(fg=THEME["accent"]))
        label.bind("<Leave>", lambda _event: label.configure(fg=THEME["muted"]))
        return label

    def setup_drop_target(self) -> None:
        if not DND_ENABLED or DND_FILES is None:
            return
        for target in (self.root, self.canvas):
            try:
                target.drop_target_register(DND_FILES)  # type: ignore[attr-defined]
                target.dnd_bind("<<Drop>>", self.on_drop)  # type: ignore[attr-defined]
            except Exception:
                pass

    def update_canvas_cursor(self, x: float | None = None, y: float | None = None) -> None:
        if self.drag_start_pdf is not None:
            self.canvas.configure(cursor="crosshair")
            return
        if self.busy_mode or self.preview_source is None or self.page_rect is None:
            self.canvas.configure(cursor="")
            return
        if x is not None and y is not None and self.point_inside_preview(x, y):
            self.canvas.configure(cursor="crosshair")
        else:
            self.canvas.configure(cursor="")

    def update_canvas_cursor_at_pointer(self) -> None:
        try:
            x = self.canvas.winfo_pointerx() - self.canvas.winfo_rootx()
            y = self.canvas.winfo_pointery() - self.canvas.winfo_rooty()
        except Exception:
            self.update_canvas_cursor()
            return
        self.update_canvas_cursor(x, y)

    def on_canvas_motion(self, event: tk.Event) -> None:
        self.update_canvas_cursor(event.x, event.y)

    def on_canvas_leave(self, _event: tk.Event) -> None:
        if self.drag_start_pdf is None:
            self.canvas.configure(cursor="")

    def update_buttons(self) -> None:
        is_busy = self.busy_mode is not None
        has_pdf = self.source_pdf is not None and self.preview_source is not None
        has_area = self.is_valid_selection()
        self.set_button_enabled(self.select_button, not is_busy)
        self.set_button_enabled(self.refresh_button, not is_busy and self.source_pdf is not None)
        self.set_button_enabled(self.reset_button, not is_busy and has_pdf and self.selection_pdf is not None)
        self.set_button_enabled(self.save_button, not is_busy and has_pdf and has_area)
        self.set_button_enabled(self.prev_button, not is_busy and has_pdf and self.current_page_index > 0)
        self.set_button_enabled(
            self.next_button,
            not is_busy and has_pdf and self.current_page_index < self.page_count - 1,
        )
        self.update_canvas_cursor()

    def set_button_enabled(self, button: tk.Button, enabled: bool) -> None:
        if enabled:
            button.configure(state="normal", cursor="hand2")
            if button is self.save_button:
                button.configure(bg=THEME["accent"], fg=THEME["white"], activebackground=THEME["accent_hover"])
            else:
                button.configure(bg=THEME["white"], fg=THEME["text"], activebackground=THEME["selection_bg"])
        else:
            button.configure(
                state="disabled",
                cursor="arrow",
                bg=THEME["button_disabled"],
                fg=THEME["muted"],
                activebackground=THEME["button_disabled"],
            )

    def set_status(self, key: str, detail: str = "") -> None:
        self.busy_mode = None
        self.status_var.set(UI_TEXT[key])
        self.detail_var.set(detail)
        color = {
            "status_complete": THEME["success"],
            "status_error": THEME["danger"],
            "status_ready": THEME["text"],
            "status_page_changed": THEME["text"],
        }.get(key, THEME["muted"])
        self.status_label.configure(fg=color)

    def start_busy_status(self, mode: str) -> None:
        self.busy_mode = mode
        self.busy_started_at = time.monotonic()
        self.status_tick = 0
        self.animate_status(force=True)
        self.update_buttons()

    def animate_status(self, force: bool = False) -> None:
        if self.busy_mode:
            elapsed = time.monotonic() - self.busy_started_at
            if elapsed >= STATUS_PHRASE_DELAY_SECONDS and self.status_tick % 8 in (4, 5, 6):
                phrase_key = ("status_phrase_1", "status_phrase_2", "status_phrase_3")[self.status_tick % 3]
                self.status_var.set(UI_TEXT[phrase_key])
            else:
                prefix = "status_loading" if self.busy_mode == "loading" else "status_saving"
                dot_key = f"{prefix}_{(self.status_tick % 3) + 1}"
                self.status_var.set(UI_TEXT[dot_key])
            self.status_label.configure(fg=THEME["muted"])
            self.status_tick += 1
        if not force:
            self.root.after(STATUS_ANIMATION_MS, self.animate_status)

    def choose_pdf(self) -> None:
        if self.busy_mode:
            return
        selected = filedialog.askopenfilename(
            title=UI_TEXT["dialog_select_pdf_title"],
            filetypes=((UI_TEXT["filetype_pdf"], "*.pdf"),),
        )
        if selected:
            self.load_pdf(Path(selected), page_index=0, ready_status_key="status_ready")

    def on_drop(self, event: Any) -> None:
        if self.busy_mode:
            return
        paths = [Path(item) for item in self.root.tk.splitlist(getattr(event, "data", ""))]
        pdf_paths = [path for path in paths if path.suffix.lower() == ".pdf" and path.is_file()]
        if not pdf_paths:
            self.show_warning(UI_TEXT["error_not_pdf"])
            return
        self.load_pdf(pdf_paths[0], page_index=0, ready_status_key="status_ready")

    def refresh_pdf(self) -> None:
        if self.busy_mode:
            return
        if self.source_pdf is None:
            self.set_status("status_error", UI_TEXT["error_not_pdf"])
            self.update_buttons()
            return
        self.load_pdf(self.source_pdf, page_index=0, ready_status_key="status_ready")

    def change_page(self, delta: int) -> None:
        if self.busy_mode or self.source_pdf is None or self.page_count < 1:
            return
        next_index = self.current_page_index + delta
        if next_index < 0 or next_index >= self.page_count:
            return
        self.load_pdf(self.source_pdf, page_index=next_index, ready_status_key="status_page_changed")

    def load_pdf(self, path: Path, page_index: int = 0, ready_status_key: str = "status_ready") -> None:
        if path.suffix.lower() != ".pdf" or not path.is_file():
            self.show_warning(UI_TEXT["error_not_pdf"])
            return
        if fitz is None or Image is None or ImageTk is None:
            self.show_error(UI_TEXT["error_dependency_missing"])
            return

        self.load_id += 1
        current_load_id = self.load_id
        self.source_pdf = path
        self.page_count = 0
        self.current_page_index = max(0, page_index)
        self.pending_ready_status_key = ready_status_key
        self.page_rect = None
        self.preview_source = None
        self.preview_photo = None
        self.selection_pdf = None
        self.drag_start_pdf = None
        self.drag_current_pdf = None
        self.detail_var.set("")
        self.page_var.set("")
        self.start_busy_status("loading")
        self.draw_loading_state()

        canvas_width = max(self.canvas.winfo_width() - CANVAS_MARGIN * 2, 640)
        canvas_height = max(self.canvas.winfo_height() - CANVAS_MARGIN * 2, 420)
        worker = threading.Thread(
            target=self.load_pdf_worker,
            args=(current_load_id, path, self.current_page_index, canvas_width, canvas_height),
            daemon=True,
        )
        worker.start()

    def load_pdf_worker(
        self,
        current_load_id: int,
        path: Path,
        page_index: int,
        target_width: int,
        target_height: int,
    ) -> None:
        try:
            if fitz is None:
                raise RuntimeError(UI_TEXT["error_dependency_missing"])
            with fitz.open(str(path)) as document:  # type: ignore[union-attr]
                if document.page_count < 1:
                    raise ValueError(UI_TEXT["error_no_pages"])
                page_index = min(max(page_index, 0), document.page_count - 1)
                page = document.load_page(page_index)
                rect = page.rect
                if rect.width <= 0 or rect.height <= 0:
                    raise ValueError(UI_TEXT["error_load_failed"])
                scale = min(target_width / rect.width, target_height / rect.height, 2.0)
                scale = max(scale, 0.2)
                pixmap = page.get_pixmap(
                    matrix=fitz.Matrix(scale, scale),  # type: ignore[union-attr]
                    alpha=False,
                    colorspace=fitz.csRGB,  # type: ignore[union-attr]
                )
                payload = PdfPreviewPayload(
                    source_path=path,
                    page_count=document.page_count,
                    page_index=page_index,
                    page_rect=(float(rect.x0), float(rect.y0), float(rect.x1), float(rect.y1)),
                    image_data=pixmap.tobytes("png"),
                )
            self.queue.put(("load_done", current_load_id, payload))
        except Exception as exc:
            self.queue.put(("load_error", current_load_id, exc))

    def poll_queue(self) -> None:
        while True:
            try:
                message = self.queue.get_nowait()
            except queue.Empty:
                break
            self.handle_queue_message(message)
        self.root.after(QUEUE_POLL_MS, self.poll_queue)

    def handle_queue_message(self, message: tuple[Any, ...]) -> None:
        kind = message[0]
        if kind in {"load_done", "load_error"}:
            current_load_id = message[1]
            if current_load_id != self.load_id:
                return

        if kind == "load_done":
            _kind, _load_id, payload = message
            self.apply_preview(payload)
            return

        if kind == "load_error":
            _kind, _load_id, exc = message
            self.source_pdf = None
            self.page_count = 0
            self.current_page_index = 0
            self.preview_source = None
            self.page_rect = None
            self.page_var.set("")
            self.set_status("status_error", UI_TEXT["error_load_failed"])
            self.draw_empty_state()
            self.update_buttons()
            message = UI_TEXT["error_load_failed"]
            detail = format_exception(exc)
            self.show_error(f"{message}\n{detail}" if detail else message)
            return

        if kind == "save_done":
            _kind, output_path = message
            path = Path(output_path)
            self.set_status("status_complete", UI_TEXT["status_saved_detail"].format(name=path.name))
            self.update_buttons()
            messagebox.showinfo(
                UI_TEXT["dialog_complete_title"],
                UI_TEXT["dialog_complete_message"].format(path=path),
                parent=self.root,
            )
            if not open_folder(path.parent):
                self.show_warning(UI_TEXT["error_open_folder_failed"])
            return

        if kind == "save_error":
            _kind, exc = message
            self.set_status("status_error", UI_TEXT["error_save_failed"])
            self.update_buttons()
            message = UI_TEXT["error_save_failed"]
            detail = format_exception(exc)
            self.show_error(f"{message}\n{detail}" if detail else message)

    def apply_preview(self, payload: PdfPreviewPayload) -> None:
        if Image is None:
            self.show_error(UI_TEXT["error_dependency_missing"])
            return
        try:
            self.preview_source = Image.open(io.BytesIO(payload.image_data)).convert("RGB")
        except Exception as exc:
            self.handle_queue_message(("load_error", self.load_id, exc))
            return
        self.source_pdf = payload.source_path
        self.page_count = payload.page_count
        self.current_page_index = payload.page_index
        self.page_rect = payload.page_rect
        self.selection_pdf = None
        self.drag_start_pdf = None
        self.drag_current_pdf = None
        self.page_var.set(
            UI_TEXT["page_indicator"].format(current=payload.page_index + 1, total=payload.page_count)
        )
        self.set_status(
            self.pending_ready_status_key,
            UI_TEXT["status_file_ready"].format(name=payload.source_path.name, count=payload.page_count),
        )
        self.render_preview()
        self.update_buttons()

    def on_canvas_configure(self, _event: tk.Event) -> None:
        if self.preview_source is not None:
            self.render_preview()
        elif self.busy_mode == "loading":
            self.draw_loading_state()
        else:
            self.draw_empty_state()

    def draw_empty_state(self) -> None:
        self.canvas.configure(cursor="")
        self.canvas.delete("all")
        width = max(self.canvas.winfo_width(), 480)
        height = max(self.canvas.winfo_height(), 300)
        card_width = min(EMPTY_CARD_WIDTH, width - 48)
        card_height = EMPTY_CARD_HEIGHT
        x0 = (width - card_width) / 2
        y0 = (height - card_height) / 2
        x1 = x0 + card_width
        y1 = y0 + card_height
        self.canvas.create_rectangle(x0, y0, x1, y1, fill=THEME["card"], outline=THEME["border"], width=1)
        title = UI_TEXT["empty_title_drop"] if DND_ENABLED else UI_TEXT["empty_title"]
        self.canvas.create_text(
            width / 2,
            y0 + 56,
            text=title,
            fill=THEME["text"],
            font=(self.font_family, 16, "bold"),
        )
        self.canvas.create_text(
            width / 2,
            y0 + 94,
            text=UI_TEXT["empty_subtitle"],
            fill=THEME["muted"],
            font=(self.font_family, 10),
        )

    def draw_loading_state(self) -> None:
        self.canvas.configure(cursor="")
        self.canvas.delete("all")
        width = max(self.canvas.winfo_width(), 480)
        height = max(self.canvas.winfo_height(), 300)
        self.canvas.create_text(
            width / 2,
            height / 2,
            text=UI_TEXT["canvas_loading"],
            fill=THEME["muted"],
            font=(self.font_family, 11),
        )

    def render_preview(self) -> None:
        if self.preview_source is None or self.page_rect is None or ImageTk is None:
            self.draw_empty_state()
            return
        self.canvas.delete("all")
        canvas_width = max(self.canvas.winfo_width(), 480)
        canvas_height = max(self.canvas.winfo_height(), 300)
        page_width = self.page_rect[2] - self.page_rect[0]
        page_height = self.page_rect[3] - self.page_rect[1]
        if page_width <= 0 or page_height <= 0:
            self.canvas.create_text(
                canvas_width / 2,
                canvas_height / 2,
                text=UI_TEXT["canvas_no_preview"],
                fill=THEME["muted"],
                font=(self.font_family, 11),
            )
            return

        available_width = max(canvas_width - CANVAS_MARGIN * 2, 120)
        available_height = max(canvas_height - CANVAS_MARGIN * 2, 120)
        display_scale = min(available_width / page_width, available_height / page_height, 1.0)
        display_scale = max(display_scale, 0.05)
        display_width = max(1, int(page_width * display_scale))
        display_height = max(1, int(page_height * display_scale))
        origin_x = (canvas_width - display_width) / 2
        origin_y = (canvas_height - display_height) / 2

        resized = self.preview_source.resize((display_width, display_height), Image.Resampling.LANCZOS)
        self.preview_photo = ImageTk.PhotoImage(resized)
        self.preview_origin = (origin_x, origin_y)
        self.preview_size = (float(display_width), float(display_height))
        self.preview_scale = display_scale
        self.canvas.create_rectangle(
            origin_x - 1,
            origin_y - 1,
            origin_x + display_width + 1,
            origin_y + display_height + 1,
            fill=THEME["white"],
            outline=THEME["border"],
        )
        self.canvas.create_image(origin_x, origin_y, image=self.preview_photo, anchor="nw")
        self.draw_selection()
        self.update_canvas_cursor_at_pointer()

    def draw_selection(self) -> None:
        self.canvas.delete("selection")
        rect = self.current_selection_pdf()
        if rect is None or self.page_rect is None:
            return
        x0, y0 = self.pdf_to_canvas(rect[0], rect[1])
        x1, y1 = self.pdf_to_canvas(rect[2], rect[3])
        self.canvas.create_rectangle(
            x0,
            y0,
            x1,
            y1,
            fill=THEME["selection_bg"],
            stipple="gray25",
            outline=THEME["selection_border"],
            width=2,
            tags=("selection",),
        )

    def current_selection_pdf(self) -> tuple[float, float, float, float] | None:
        if self.drag_start_pdf and self.drag_current_pdf:
            return self.normalize_rect((*self.drag_start_pdf, *self.drag_current_pdf))
        return self.selection_pdf

    def normalize_rect(self, values: tuple[float, float, float, float]) -> tuple[float, float, float, float]:
        x0, y0, x1, y1 = values
        return min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1)

    def canvas_to_pdf(self, x: float, y: float) -> tuple[float, float]:
        if self.page_rect is None:
            return 0.0, 0.0
        origin_x, origin_y = self.preview_origin
        width, height = self.preview_size
        clamped_x = min(max(x, origin_x), origin_x + width)
        clamped_y = min(max(y, origin_y), origin_y + height)
        pdf_x = self.page_rect[0] + (clamped_x - origin_x) / self.preview_scale
        pdf_y = self.page_rect[1] + (clamped_y - origin_y) / self.preview_scale
        return pdf_x, pdf_y

    def pdf_to_canvas(self, pdf_x: float, pdf_y: float) -> tuple[float, float]:
        if self.page_rect is None:
            return 0.0, 0.0
        origin_x, origin_y = self.preview_origin
        x = origin_x + (pdf_x - self.page_rect[0]) * self.preview_scale
        y = origin_y + (pdf_y - self.page_rect[1]) * self.preview_scale
        return x, y

    def point_inside_preview(self, x: float, y: float) -> bool:
        origin_x, origin_y = self.preview_origin
        width, height = self.preview_size
        return origin_x <= x <= origin_x + width and origin_y <= y <= origin_y + height

    def on_drag_start(self, event: tk.Event) -> None:
        if self.busy_mode:
            return
        if self.preview_source is None or self.page_rect is None:
            self.choose_pdf()
            return
        if not self.point_inside_preview(event.x, event.y):
            return
        point = self.canvas_to_pdf(event.x, event.y)
        self.canvas.configure(cursor="crosshair")
        self.drag_start_pdf = point
        self.drag_current_pdf = point
        self.selection_pdf = None
        self.draw_selection()
        self.update_buttons()
        self.update_canvas_cursor(event.x, event.y)

    def on_drag_move(self, event: tk.Event) -> None:
        if self.drag_start_pdf is None:
            return
        self.drag_current_pdf = self.canvas_to_pdf(event.x, event.y)
        self.draw_selection()

    def on_drag_end(self, event: tk.Event) -> None:
        if self.drag_start_pdf is None:
            return
        self.drag_current_pdf = self.canvas_to_pdf(event.x, event.y)
        rect = self.current_selection_pdf()
        self.drag_start_pdf = None
        self.drag_current_pdf = None
        self.selection_pdf = rect
        if not self.is_valid_selection():
            self.selection_pdf = None
            self.set_status("status_error", UI_TEXT["error_area_too_small"])
        else:
            self.set_status("status_ready", UI_TEXT["status_area_selected"])
        self.draw_selection()
        self.update_buttons()

    def reset_selection(self) -> None:
        if self.busy_mode:
            return
        self.selection_pdf = None
        self.drag_start_pdf = None
        self.drag_current_pdf = None
        self.draw_selection()
        if self.source_pdf is not None:
            self.set_status("status_ready", UI_TEXT["status_area_reset"])
        else:
            self.set_status("status_idle")
        self.update_buttons()

    def is_valid_selection(self) -> bool:
        if self.selection_pdf is None or self.page_rect is None:
            return False
        x0, y0, x1, y1 = self.selection_pdf
        page_width = self.page_rect[2] - self.page_rect[0]
        page_height = self.page_rect[3] - self.page_rect[1]
        min_width = max(MIN_SELECTION_CANVAS / self.preview_scale, page_width * MIN_SELECTION_RATIO)
        min_height = max(MIN_SELECTION_CANVAS / self.preview_scale, page_height * MIN_SELECTION_RATIO)
        return (x1 - x0) >= min_width and (y1 - y0) >= min_height

    def selection_ratios(self) -> tuple[float, float, float, float]:
        if self.selection_pdf is None or self.page_rect is None:
            raise ValueError(UI_TEXT["error_no_area"])
        x0, y0, x1, y1 = self.selection_pdf
        page_x0, page_y0, page_x1, page_y1 = self.page_rect
        page_width = page_x1 - page_x0
        page_height = page_y1 - page_y0
        if page_width <= 0 or page_height <= 0:
            raise ValueError(UI_TEXT["error_save_failed"])
        ratios = (
            (x0 - page_x0) / page_width,
            (y0 - page_y0) / page_height,
            (x1 - page_x0) / page_width,
            (y1 - page_y0) / page_height,
        )
        return tuple(min(max(value, 0.0), 1.0) for value in ratios)  # type: ignore[return-value]

    def save_crop(self) -> None:
        if self.busy_mode:
            return
        if self.source_pdf is None:
            self.show_warning(UI_TEXT["error_not_pdf"])
            return
        if not self.is_valid_selection():
            self.show_warning(UI_TEXT["error_no_area"])
            return
        try:
            ratios = self.selection_ratios()
        except Exception as exc:
            self.show_error(str(exc))
            return
        output_path = unique_output_path(self.source_pdf, self.current_page_index)
        self.start_busy_status("saving")
        worker = threading.Thread(
            target=self.save_crop_worker,
            args=(self.source_pdf, output_path, self.current_page_index, ratios),
            daemon=True,
        )
        worker.start()

    def save_crop_worker(
        self,
        source_pdf: Path,
        output_path: Path,
        page_index: int,
        ratios: tuple[float, float, float, float],
    ) -> None:
        try:
            crop_single_page_with_ratios(source_pdf, output_path, page_index, ratios)
            self.queue.put(("save_done", str(output_path)))
        except Exception as exc:
            self.queue.put(("save_error", exc))

    def show_warning(self, message: str) -> None:
        self.set_status("status_error", message)
        messagebox.showwarning(UI_TEXT["dialog_warning_title"], message, parent=self.root)

    def show_error(self, message: str) -> None:
        self.set_status("status_error", message)
        messagebox.showerror(UI_TEXT["dialog_error_title"], message, parent=self.root)


def main() -> None:
    root = make_root()
    PdfCropApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
