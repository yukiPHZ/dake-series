# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import queue
import re
import threading
import urllib.parse
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, ttk

import fitz
from PIL import Image, ImageTk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    ROOT_CLASS = TkinterDnD.Tk
    DND_ENABLED = True
except Exception:
    DND_FILES = None
    ROOT_CLASS = tk.Tk
    DND_ENABLED = False


APP_NAME = "Dake確認印"
WINDOW_TITLE = "Dake確認印"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "brand_series": "シンプルそれDAKEシリーズ",
    "main_title": "PDFに確認印を押す",
    "main_description": "PDFを選んで、押したい場所をクリックします。",
    "button_add_pdf": "PDFを選ぶ",
    "button_save": "確認印を押して保存",
    "button_clear": "やり直す",
    "button_prev": "‹",
    "button_next": "›",
    "label_name": "名字",
    "label_date": "日付",
    "label_size": "印影サイズ",
    "label_file": "選択中",
    "size_small": "小",
    "size_medium": "中",
    "size_large": "大",
    "empty_title": "PDFを追加してください",
    "empty_subtitle": "ドラッグ＆ドロップ または クリックして追加",
    "drop_title": "PDFをドロップしてください",
    "preview_loading": "プレビューを読み込んでいます",
    "status_idle": "未選択",
    "status_loading": "読み込み中",
    "status_ready": "押したい場所をクリックしてください",
    "status_stamp_ready": "確認印の位置を指定しました",
    "status_processing": "保存中",
    "status_complete": "保存完了",
    "status_error": "エラー",
    "notice": "この印影は社内確認・作業記録用です。電子署名・電子契約・本人性証明を目的とするものではありません。",
    "error_pdf_only": "PDFファイルを選択してください。",
    "error_no_pdf": "PDFが選択されていません。",
    "error_no_position": "確認印を押す位置をクリックしてください。",
    "error_name_empty": "名字を入力してください。",
    "error_date_invalid": "日付は8桁の数字で入力してください。",
    "error_save_failed": "保存できませんでした。",
    "error_load_failed": "PDFを読み込めませんでした。",
    "error_no_pages": "PDFにページがありません。",
    "complete_title": "保存完了",
    "complete_message": "確認印付きPDFを保存しました。",
    "complete_message_detail": "確認印付きPDFを保存しました。\n\n保存先：{path}",
    "dialog_error_title": "エラー",
    "dialog_select_pdf_title": "PDFを選択",
    "dialog_pdf_filter_label": "PDFファイル",
    "dialog_all_filter_label": "すべてのファイル",
    "page_format": "{current} / {total}",
    "file_none": "未選択",
    "file_name_template": "{name}",
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
    "success": "#12B76A",
    "success_bg": "#EAFBF3",
    "error": "#D92D20",
    "error_bg": "#FDECEC",
    "soft": "#EEF2F7",
    "stamp": "#C62828",
}

LINKS = {
    "footer_link_1": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "footer_link_2": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

STATUS_COLORS = {
    "status_idle": ("soft", "muted"),
    "status_loading": ("selection_bg", "accent"),
    "status_ready": ("selection_bg", "accent"),
    "status_stamp_ready": ("success_bg", "success"),
    "status_processing": ("selection_bg", "accent"),
    "status_complete": ("success_bg", "success"),
    "status_error": ("error_bg", "error"),
}

SIZE_OPTIONS = {
    "small": ("size_small", 60),
    "medium": ("size_medium", 80),
    "large": ("size_large", 100),
}

OUTPUT_SUFFIX = "_確認印"
DEFAULT_NAME = "菊田"
CJK_FONT_NAME = "japan"
DATE_FONT_NAME = "helv"
WINDOW_SIZE = "1080x760"
WINDOW_MIN_SIZE = (960, 680)
RENDER_SCALE = 2.0
PDF_POINT_PER_UI_PIXEL = 0.75
QUEUE_POLL_INTERVAL_MS = 80
STAMP_RGB = (198 / 255, 40 / 255, 40 / 255)
STAMP_STROKE_OPACITY = 0.9
STAMP_TEXT_OPACITY = 0.92


@dataclass(frozen=True)
class StampPosition:
    page_index: int
    x: float
    y: float


@dataclass(frozen=True)
class RenderResult:
    token: int
    page_index: int
    page_count: int
    page_width: float
    page_height: float
    image: Image.Image
    error: str | None = None


def get_common_icon_path() -> Path:
    return (Path(__file__).resolve().parent / ".." / ".." / "02_assets" / "dake_icon.ico").resolve()


def choose_font_family(root: tk.Tk) -> str:
    preferred = ["BIZ UDPGothic", "Yu Gothic UI", "Meiryo"]
    available = set(tkfont.families(root))
    for family in preferred:
        if family in available:
            return family
    return "TkDefaultFont"


def parse_dropped_files(root: tk.Tk, raw_data: str) -> list[Path]:
    paths: list[Path] = []
    for raw_item in root.tk.splitlist(raw_data):
        value = raw_item.strip().strip("{}")
        if value.startswith("file:"):
            parsed = urllib.parse.urlparse(value)
            value = urllib.parse.unquote(parsed.path)
            if re.match(r"^/[A-Za-z]:/", value):
                value = value[1:]
        if value:
            paths.append(Path(value))
    return paths


def validate_pdf_path(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() == ".pdf"


def build_output_path(source_path: Path) -> Path:
    destination = source_path.with_name(f"{source_path.stem}{OUTPUT_SUFFIX}.pdf")
    if destination.resolve() != source_path.resolve() and not destination.exists():
        return destination

    index = 2
    while True:
        candidate = source_path.with_name(f"{source_path.stem}{OUTPUT_SUFFIX}_{index}.pdf")
        if candidate.resolve() != source_path.resolve() and not candidate.exists():
            return candidate
        index += 1


def fit_font_size(
    text: str,
    diameter: float,
    start_size: float,
    max_width: float,
    minimum: float,
    fontname: str = CJK_FONT_NAME,
) -> float:
    try:
        font = fitz.Font(fontname)
        size = start_size
        while size > minimum and font.text_length(text, fontsize=size) > max_width:
            size -= 0.5
        return max(size, minimum)
    except Exception:
        return max(min(start_size, diameter * 0.28), minimum)


def insert_centered_stamp_text(
    page: fitz.Page,
    text: str,
    center_x: float,
    baseline_y: float,
    fontsize: float,
    fontname: str,
) -> None:
    try:
        font = fitz.Font(fontname)
        text_width = font.text_length(text, fontsize=fontsize)
    except Exception:
        text_width = len(text) * fontsize * 0.55
    page.insert_text(
        (center_x - text_width / 2, baseline_y),
        text,
        fontsize=fontsize,
        fontname=fontname,
        color=STAMP_RGB,
        overlay=True,
        fill_opacity=STAMP_TEXT_OPACITY,
    )


def clamp_stamp_center(page_width: float, page_height: float, x: float, y: float, diameter: float) -> tuple[float, float]:
    half = diameter / 2
    if page_width > diameter:
        x = min(max(x, half), page_width - half)
    if page_height > diameter:
        y = min(max(y, half), page_height - half)
    return x, y


def draw_stamp_on_page(page: fitz.Page, center_x: float, center_y: float, diameter: float, name: str, date_text: str) -> None:
    half = diameter / 2
    rect = fitz.Rect(center_x - half, center_y - half, center_x + half, center_y + half)
    border_width = max(1.2, diameter * 0.035)
    page.draw_oval(
        rect,
        color=STAMP_RGB,
        width=border_width,
        overlay=True,
        stroke_opacity=STAMP_STROKE_OPACITY,
    )

    # Future extension point: place the date along the outer circle if the stamp style is expanded.
    name_size = fit_font_size(name, diameter, diameter * 0.28, diameter * 0.68, diameter * 0.16)
    date_size = fit_font_size(date_text, diameter, diameter * 0.12, diameter * 0.70, diameter * 0.085, DATE_FONT_NAME)

    name_baseline_y = center_y - diameter * 0.08 + name_size * 0.35
    date_baseline_y = center_y + diameter * 0.26 + date_size * 0.35
    insert_centered_stamp_text(page, name, center_x, name_baseline_y, name_size, CJK_FONT_NAME)
    insert_centered_stamp_text(page, date_text, center_x, date_baseline_y, date_size, DATE_FONT_NAME)


def save_stamped_pdf(
    source_path: Path,
    output_path: Path,
    stamp_position: StampPosition,
    name: str,
    date_text: str,
    diameter: float,
) -> None:
    with fitz.open(source_path) as document:
        if document.page_count < 1:
            raise ValueError(UI_TEXT["error_no_pages"])
        page_index = min(max(stamp_position.page_index, 0), document.page_count - 1)
        page = document.load_page(page_index)
        center_x, center_y = clamp_stamp_center(page.rect.width, page.rect.height, stamp_position.x, stamp_position.y, diameter)
        draw_stamp_on_page(page, center_x, center_y, diameter, name, date_text)
        document.save(output_path, garbage=4, deflate=True)


class CheckStampApp:
    def __init__(self) -> None:
        self.root = ROOT_CLASS()
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(*WINDOW_MIN_SIZE)
        self.root.configure(bg=THEME["background"])

        icon_path = get_common_icon_path()
        if icon_path.exists():
            try:
                self.root.iconbitmap(str(icon_path))
            except Exception:
                pass

        self.font_family = choose_font_family(self.root)
        self.render_queue: queue.Queue[RenderResult | tuple[str, Path | Exception]] = queue.Queue()
        self.render_token = 0
        self.busy = False

        self.pdf_path: Path | None = None
        self.page_count = 0
        self.current_page_index = 0
        self.page_width = 0.0
        self.page_height = 0.0
        self.stamp_position: StampPosition | None = None

        self.preview_source_image: Image.Image | None = None
        self.preview_photo: ImageTk.PhotoImage | None = None
        self.preview_origin = (0, 0)
        self.preview_display_size = (0, 0)

        self.name_var = tk.StringVar(value=DEFAULT_NAME)
        self.date_var = tk.StringVar(value=datetime.now().strftime("%Y%m%d"))
        self.size_var = tk.StringVar(value="medium")
        self.status_var = tk.StringVar(value=UI_TEXT["status_idle"])
        self.file_var = tk.StringVar(value=UI_TEXT["file_none"])
        self.page_var = tk.StringVar(value=UI_TEXT["page_format"].format(current=0, total=0))

        self.configure_styles()
        self.build_ui()
        self.enable_drop(self.root)
        self.root.after(QUEUE_POLL_INTERVAL_MS, self.poll_queue)

    def configure_styles(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TEntry", padding=(8, 6), font=(self.font_family, 11))
        style.configure(
            "TRadiobutton",
            background=THEME["card"],
            foreground=THEME["text"],
            font=(self.font_family, 10),
        )

    def build_ui(self) -> None:
        outer = tk.Frame(self.root, bg=THEME["background"])
        outer.pack(fill="both", expand=True, padx=20, pady=18)

        self.build_header(outer)

        content = tk.Frame(outer, bg=THEME["background"])
        content.pack(fill="both", expand=True, pady=(14, 0))
        content.grid_columnconfigure(0, weight=1)
        content.grid_columnconfigure(1, minsize=300)
        content.grid_rowconfigure(0, weight=1)

        self.build_preview_card(content)
        self.build_control_card(content)
        self.build_footer(outer)

    def build_header(self, parent: tk.Frame) -> None:
        header = tk.Frame(parent, bg=THEME["background"])
        header.pack(fill="x")

        tk.Label(
            header,
            text=UI_TEXT["brand_series"],
            bg=THEME["background"],
            fg=THEME["accent"],
            font=(self.font_family, 10, "bold"),
        ).pack(anchor="w")
        tk.Label(
            header,
            text=UI_TEXT["main_title"],
            bg=THEME["background"],
            fg=THEME["text"],
            font=(self.font_family, 22, "bold"),
        ).pack(anchor="w", pady=(4, 0))
        tk.Label(
            header,
            text=UI_TEXT["main_description"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 11),
        ).pack(anchor="w", pady=(4, 0))

    def build_preview_card(self, parent: tk.Frame) -> None:
        card = self.create_card(parent)
        card.grid(row=0, column=0, sticky="nsew", padx=(0, 14))
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(1, weight=1)

        toolbar = tk.Frame(card, bg=THEME["card"])
        toolbar.grid(row=0, column=0, sticky="ew", padx=14, pady=(12, 8))
        toolbar.grid_columnconfigure(1, weight=1)

        self.add_button = self.create_button(toolbar, UI_TEXT["button_add_pdf"], self.select_pdf, primary=False)
        self.add_button.grid(row=0, column=0, sticky="w")

        file_label = tk.Label(
            toolbar,
            textvariable=self.file_var,
            bg=THEME["card"],
            fg=THEME["muted"],
            font=(self.font_family, 10),
            anchor="w",
        )
        file_label.grid(row=0, column=1, sticky="ew", padx=12)

        page_nav = tk.Frame(toolbar, bg=THEME["card"])
        page_nav.grid(row=0, column=2, sticky="e")
        self.prev_button = self.create_square_button(page_nav, UI_TEXT["button_prev"], self.go_prev_page)
        self.prev_button.pack(side="left")
        self.page_label = tk.Label(
            page_nav,
            textvariable=self.page_var,
            bg=THEME["card"],
            fg=THEME["muted"],
            font=(self.font_family, 10, "bold"),
            width=8,
        )
        self.page_label.pack(side="left", padx=6)
        self.next_button = self.create_square_button(page_nav, UI_TEXT["button_next"], self.go_next_page)
        self.next_button.pack(side="left")

        self.canvas = tk.Canvas(
            card,
            bg=THEME["soft"],
            bd=0,
            highlightthickness=1,
            highlightbackground=THEME["border"],
            cursor="hand2",
        )
        self.canvas.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 14))
        self.canvas.bind("<Button-1>", self.handle_canvas_click)
        self.canvas.bind("<Configure>", self.handle_canvas_resize)
        self.enable_drop(self.canvas)
        self.draw_empty_preview()

    def build_control_card(self, parent: tk.Frame) -> None:
        card = self.create_card(parent)
        card.grid(row=0, column=1, sticky="nsew")
        card.grid_columnconfigure(0, weight=1)

        form = tk.Frame(card, bg=THEME["card"])
        form.grid(row=0, column=0, sticky="new", padx=18, pady=18)
        form.grid_columnconfigure(0, weight=1)

        self.create_field_label(form, UI_TEXT["label_name"]).grid(row=0, column=0, sticky="w")
        self.name_entry = ttk.Entry(form, textvariable=self.name_var)
        self.name_entry.grid(row=1, column=0, sticky="ew", pady=(6, 14))

        self.create_field_label(form, UI_TEXT["label_date"]).grid(row=2, column=0, sticky="w")
        self.date_entry = ttk.Entry(form, textvariable=self.date_var)
        self.date_entry.grid(row=3, column=0, sticky="ew", pady=(6, 14))

        self.create_field_label(form, UI_TEXT["label_size"]).grid(row=4, column=0, sticky="w")
        size_frame = tk.Frame(form, bg=THEME["card"])
        size_frame.grid(row=5, column=0, sticky="ew", pady=(8, 18))
        for value, (label_key, _) in SIZE_OPTIONS.items():
            rb = tk.Radiobutton(
                size_frame,
                text=UI_TEXT[label_key],
                value=value,
                variable=self.size_var,
                command=self.handle_size_change,
                bg=THEME["card"],
                fg=THEME["text"],
                selectcolor=THEME["selection_bg"],
                activebackground=THEME["card"],
                activeforeground=THEME["text"],
                font=(self.font_family, 10),
                padx=6,
            )
            rb.pack(side="left", padx=(0, 12))

        self.save_button = self.create_button(form, UI_TEXT["button_save"], self.start_save, primary=True)
        self.save_button.grid(row=6, column=0, sticky="ew", pady=(2, 10))

        self.clear_button = self.create_button(form, UI_TEXT["button_clear"], self.clear_stamp, primary=False)
        self.clear_button.grid(row=7, column=0, sticky="ew")

        status_area = tk.Frame(card, bg=THEME["card"])
        status_area.grid(row=1, column=0, sticky="new", padx=18, pady=(0, 14))
        self.status_badge = tk.Label(
            status_area,
            textvariable=self.status_var,
            bg=THEME["soft"],
            fg=THEME["muted"],
            font=(self.font_family, 10, "bold"),
            padx=12,
            pady=7,
        )
        self.status_badge.pack(anchor="w")

        notice = tk.Label(
            card,
            text=UI_TEXT["notice"],
            bg=THEME["card"],
            fg=THEME["muted"],
            font=(self.font_family, 9),
            justify="left",
            wraplength=250,
        )
        notice.grid(row=2, column=0, sticky="sew", padx=18, pady=(0, 18))

        card.grid_rowconfigure(3, weight=1)
        self.update_controls()

    def build_footer(self, parent: tk.Frame) -> None:
        footer = tk.Frame(parent, bg=THEME["background"])
        footer.pack(fill="x", pady=(12, 0))

        left = tk.Label(
            footer,
            text=UI_TEXT["footer_left"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 9),
        )
        left.pack(side="left")

        right = tk.Frame(footer, bg=THEME["background"])
        right.pack(side="right")
        self.create_footer_link(right, "footer_link_1")
        self.create_footer_text(right, "footer_separator")
        self.create_footer_link(right, "footer_link_2")
        self.create_footer_text(right, "footer_separator")
        self.create_footer_text(right, "footer_copyright")

    def create_card(self, parent: tk.Widget) -> tk.Frame:
        return tk.Frame(
            parent,
            bg=THEME["card"],
            highlightthickness=1,
            highlightbackground=THEME["border"],
            bd=0,
        )

    def create_button(self, parent: tk.Widget, text: str, command, primary: bool) -> tk.Button:
        bg = THEME["accent"] if primary else THEME["card"]
        fg = "#FFFFFF" if primary else THEME["text"]
        border = THEME["accent"] if primary else THEME["border"]
        active_bg = THEME["accent_hover"] if primary else THEME["selection_bg"]
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg=fg,
            activebackground=active_bg,
            activeforeground=fg,
            bd=0,
            highlightthickness=1,
            highlightbackground=border,
            font=(self.font_family, 10, "bold"),
            padx=12,
            pady=9,
            cursor="hand2",
        )

    def create_square_button(self, parent: tk.Widget, text: str, command) -> tk.Button:
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=THEME["card"],
            fg=THEME["text"],
            activebackground=THEME["selection_bg"],
            activeforeground=THEME["text"],
            bd=0,
            highlightthickness=1,
            highlightbackground=THEME["border"],
            font=(self.font_family, 13, "bold"),
            width=3,
            cursor="hand2",
        )

    def create_field_label(self, parent: tk.Widget, text: str) -> tk.Label:
        return tk.Label(
            parent,
            text=text,
            bg=THEME["card"],
            fg=THEME["text"],
            font=(self.font_family, 10, "bold"),
        )

    def create_footer_link(self, parent: tk.Widget, text_key: str) -> None:
        label = tk.Label(
            parent,
            text=UI_TEXT[text_key],
            bg=THEME["background"],
            fg=THEME["accent"],
            font=(self.font_family, 9, "underline"),
            cursor="hand2",
        )
        label.pack(side="left")
        label.bind("<Button-1>", lambda _event, key=text_key: self.open_link(key))

    def create_footer_text(self, parent: tk.Widget, text_key: str) -> None:
        tk.Label(
            parent,
            text=UI_TEXT[text_key],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 9),
        ).pack(side="left")

    def open_link(self, key: str) -> None:
        import webbrowser

        webbrowser.open(LINKS[key])

    def enable_drop(self, widget: tk.Widget) -> None:
        if not DND_ENABLED or DND_FILES is None:
            return
        try:
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", self.handle_drop)
        except Exception:
            pass

    def handle_drop(self, event) -> None:
        if self.busy:
            return
        for path in parse_dropped_files(self.root, event.data):
            if validate_pdf_path(path):
                self.load_pdf(path)
                return
        self.show_error(UI_TEXT["error_pdf_only"])

    def select_pdf(self) -> None:
        if self.busy:
            return
        selected = filedialog.askopenfilename(
            title=UI_TEXT["dialog_select_pdf_title"],
            filetypes=[
                (UI_TEXT["dialog_pdf_filter_label"], "*.pdf"),
                (UI_TEXT["dialog_all_filter_label"], "*.*"),
            ],
        )
        if selected:
            self.load_pdf(Path(selected))

    def load_pdf(self, path: Path) -> None:
        if not validate_pdf_path(path):
            self.show_error(UI_TEXT["error_pdf_only"])
            return

        self.pdf_path = path
        self.page_count = 0
        self.current_page_index = 0
        self.page_width = 0.0
        self.page_height = 0.0
        self.stamp_position = None
        self.preview_source_image = None
        self.file_var.set(UI_TEXT["file_name_template"].format(name=path.name))
        self.set_status("status_loading")
        self.draw_loading_preview()
        self.start_render()
        self.update_controls()

    def start_render(self) -> None:
        if self.pdf_path is None:
            return
        self.render_token += 1
        token = self.render_token
        pdf_path = self.pdf_path
        page_index = self.current_page_index

        def worker() -> None:
            try:
                with fitz.open(pdf_path) as document:
                    if document.page_count < 1:
                        raise ValueError(UI_TEXT["error_no_pages"])
                    safe_index = min(max(page_index, 0), document.page_count - 1)
                    page = document.load_page(safe_index)
                    matrix = fitz.Matrix(RENDER_SCALE, RENDER_SCALE)
                    pixmap = page.get_pixmap(matrix=matrix, alpha=False)
                    image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
                    result = RenderResult(
                        token=token,
                        page_index=safe_index,
                        page_count=document.page_count,
                        page_width=page.rect.width,
                        page_height=page.rect.height,
                        image=image,
                    )
            except Exception as exc:
                result = RenderResult(
                    token=token,
                    page_index=page_index,
                    page_count=0,
                    page_width=0.0,
                    page_height=0.0,
                    image=Image.new("RGB", (1, 1), "white"),
                    error=str(exc),
                )
            self.render_queue.put(result)

        threading.Thread(target=worker, daemon=True).start()

    def poll_queue(self) -> None:
        try:
            while True:
                item = self.render_queue.get_nowait()
                if isinstance(item, RenderResult):
                    self.handle_render_result(item)
                else:
                    marker, payload = item
                    if marker == "save_complete" and isinstance(payload, Path):
                        self.handle_save_complete(payload)
                    elif marker == "save_error" and isinstance(payload, Exception):
                        self.handle_save_error(payload)
        except queue.Empty:
            pass
        self.root.after(QUEUE_POLL_INTERVAL_MS, self.poll_queue)

    def handle_render_result(self, result: RenderResult) -> None:
        if result.token != self.render_token:
            return
        if result.error is not None:
            self.preview_source_image = None
            self.page_count = 0
            self.current_page_index = 0
            self.page_width = 0.0
            self.page_height = 0.0
            self.set_status("status_error")
            self.draw_empty_preview(UI_TEXT["error_load_failed"], result.error)
            self.update_controls()
            return

        self.preview_source_image = result.image
        self.current_page_index = result.page_index
        self.page_count = result.page_count
        self.page_width = result.page_width
        self.page_height = result.page_height
        self.set_status("status_ready")
        self.update_page_label()
        self.redraw_preview()
        self.update_controls()

    def handle_canvas_resize(self, _event) -> None:
        self.redraw_preview()

    def handle_canvas_click(self, event) -> None:
        if self.busy:
            return
        if self.pdf_path is None or self.preview_source_image is None:
            self.select_pdf()
            return
        page_point = self.canvas_to_page_point(event.x, event.y)
        if page_point is None:
            return
        diameter = self.current_stamp_diameter_points()
        x, y = clamp_stamp_center(self.page_width, self.page_height, page_point[0], page_point[1], diameter)
        self.stamp_position = StampPosition(self.current_page_index, x, y)
        self.set_status("status_stamp_ready")
        self.redraw_preview()
        self.update_controls()

    def canvas_to_page_point(self, canvas_x: int, canvas_y: int) -> tuple[float, float] | None:
        origin_x, origin_y = self.preview_origin
        display_width, display_height = self.preview_display_size
        if display_width <= 0 or display_height <= 0 or self.page_width <= 0 or self.page_height <= 0:
            return None
        local_x = canvas_x - origin_x
        local_y = canvas_y - origin_y
        if local_x < 0 or local_y < 0 or local_x > display_width or local_y > display_height:
            return None
        return (
            local_x / display_width * self.page_width,
            local_y / display_height * self.page_height,
        )

    def page_to_canvas_point(self, page_x: float, page_y: float) -> tuple[float, float]:
        origin_x, origin_y = self.preview_origin
        display_width, display_height = self.preview_display_size
        return (
            origin_x + page_x / self.page_width * display_width,
            origin_y + page_y / self.page_height * display_height,
        )

    def draw_empty_preview(self, title: str | None = None, subtitle: str | None = None) -> None:
        self.canvas.delete("all")
        width = max(1, self.canvas.winfo_width())
        height = max(1, self.canvas.winfo_height())
        title = title or UI_TEXT["empty_title"]
        subtitle = subtitle or UI_TEXT["empty_subtitle"]
        self.canvas.configure(bg=THEME["soft"])
        self.canvas.create_text(
            width / 2,
            height / 2 - 18,
            text=title,
            fill=THEME["text"],
            font=(self.font_family, 16, "bold"),
        )
        self.canvas.create_text(
            width / 2,
            height / 2 + 16,
            text=subtitle,
            fill=THEME["muted"],
            font=(self.font_family, 10),
        )

    def draw_loading_preview(self) -> None:
        self.draw_empty_preview(UI_TEXT["preview_loading"], UI_TEXT["drop_title"] if DND_ENABLED else UI_TEXT["empty_subtitle"])

    def redraw_preview(self) -> None:
        if self.preview_source_image is None:
            if self.pdf_path is None:
                self.draw_empty_preview()
            return
        canvas_width = max(1, self.canvas.winfo_width())
        canvas_height = max(1, self.canvas.winfo_height())
        padding = 18
        available_width = max(1, canvas_width - padding * 2)
        available_height = max(1, canvas_height - padding * 2)
        source_width, source_height = self.preview_source_image.size
        scale = min(available_width / source_width, available_height / source_height)
        display_width = max(1, int(source_width * scale))
        display_height = max(1, int(source_height * scale))
        resized = self.preview_source_image.resize((display_width, display_height), Image.Resampling.LANCZOS)
        self.preview_photo = ImageTk.PhotoImage(resized)

        origin_x = (canvas_width - display_width) // 2
        origin_y = (canvas_height - display_height) // 2
        self.preview_origin = (origin_x, origin_y)
        self.preview_display_size = (display_width, display_height)

        self.canvas.delete("all")
        self.canvas.configure(bg=THEME["soft"])
        self.canvas.create_image(origin_x, origin_y, anchor="nw", image=self.preview_photo)
        self.canvas.create_rectangle(
            origin_x,
            origin_y,
            origin_x + display_width,
            origin_y + display_height,
            outline=THEME["border"],
            width=1,
        )
        self.draw_stamp_overlay()

    def draw_stamp_overlay(self) -> None:
        if self.stamp_position is None:
            return
        if self.stamp_position.page_index != self.current_page_index:
            return
        if self.page_width <= 0 or self.page_height <= 0:
            return
        center_x, center_y = self.page_to_canvas_point(self.stamp_position.x, self.stamp_position.y)
        display_width, _display_height = self.preview_display_size
        diameter = self.current_stamp_diameter_points() / self.page_width * display_width
        half = diameter / 2
        stamp_color = THEME["stamp"]
        self.canvas.create_oval(
            center_x - half,
            center_y - half,
            center_x + half,
            center_y + half,
            outline=stamp_color,
            width=max(2, int(diameter * 0.035)),
        )
        name_size = max(8, int(diameter * 0.25))
        date_size = max(6, int(diameter * 0.11))
        self.canvas.create_text(
            center_x,
            center_y - diameter * 0.08,
            text=self.name_var.get().strip(),
            fill=stamp_color,
            font=(self.font_family, name_size, "bold"),
        )
        self.canvas.create_text(
            center_x,
            center_y + diameter * 0.24,
            text=self.date_var.get().strip(),
            fill=stamp_color,
            font=(self.font_family, date_size, "bold"),
        )

    def current_stamp_diameter_points(self) -> float:
        _label_key, ui_pixels = SIZE_OPTIONS.get(self.size_var.get(), SIZE_OPTIONS["medium"])
        return ui_pixels * PDF_POINT_PER_UI_PIXEL

    def handle_size_change(self) -> None:
        if self.stamp_position is not None and self.page_width > 0 and self.page_height > 0:
            diameter = self.current_stamp_diameter_points()
            x, y = clamp_stamp_center(self.page_width, self.page_height, self.stamp_position.x, self.stamp_position.y, diameter)
            self.stamp_position = StampPosition(self.stamp_position.page_index, x, y)
        self.redraw_preview()

    def go_prev_page(self) -> None:
        if self.busy or self.pdf_path is None or self.current_page_index <= 0:
            return
        self.current_page_index -= 1
        self.stamp_position = None
        self.set_status("status_loading")
        self.draw_loading_preview()
        self.start_render()
        self.update_controls()

    def go_next_page(self) -> None:
        if self.busy or self.pdf_path is None or self.current_page_index >= self.page_count - 1:
            return
        self.current_page_index += 1
        self.stamp_position = None
        self.set_status("status_loading")
        self.draw_loading_preview()
        self.start_render()
        self.update_controls()

    def clear_stamp(self) -> None:
        if self.busy:
            return
        self.stamp_position = None
        if self.pdf_path is None:
            self.file_var.set(UI_TEXT["file_none"])
            self.set_status("status_idle")
            self.draw_empty_preview()
        else:
            self.set_status("status_ready")
            self.redraw_preview()
        self.update_controls()

    def start_save(self) -> None:
        if self.busy:
            return
        if self.pdf_path is None:
            self.show_error(UI_TEXT["error_no_pdf"])
            return
        if self.stamp_position is None:
            self.show_error(UI_TEXT["error_no_position"])
            return
        name = self.name_var.get().strip()
        date_text = self.date_var.get().strip()
        if not name:
            self.show_error(UI_TEXT["error_name_empty"])
            return
        if not re.fullmatch(r"\d{8}", date_text):
            self.show_error(UI_TEXT["error_date_invalid"])
            return

        output_path = build_output_path(self.pdf_path)
        diameter = self.current_stamp_diameter_points()
        position = self.stamp_position
        self.busy = True
        self.set_status("status_processing")
        self.update_controls()

        def worker() -> None:
            try:
                save_stamped_pdf(self.pdf_path, output_path, position, name, date_text, diameter)
                self.render_queue.put(("save_complete", output_path))
            except Exception as exc:
                self.render_queue.put(("save_error", exc))

        threading.Thread(target=worker, daemon=True).start()

    def handle_save_complete(self, output_path: Path) -> None:
        self.busy = False
        self.set_status("status_complete")
        self.update_controls()
        messagebox.showinfo(
            UI_TEXT["complete_title"],
            UI_TEXT["complete_message_detail"].format(path=str(output_path)),
        )
        self.open_folder(output_path.parent)

    def handle_save_error(self, exc: Exception) -> None:
        self.busy = False
        self.set_status("status_error")
        self.update_controls()
        detail = str(exc).strip()
        message = UI_TEXT["error_save_failed"] if not detail else f"{UI_TEXT['error_save_failed']}\n\n{detail}"
        self.show_error(message)

    def update_page_label(self) -> None:
        if self.page_count < 1:
            self.page_var.set(UI_TEXT["page_format"].format(current=0, total=0))
        else:
            self.page_var.set(UI_TEXT["page_format"].format(current=self.current_page_index + 1, total=self.page_count))

    def update_controls(self) -> None:
        has_pdf = self.pdf_path is not None and self.preview_source_image is not None
        has_stamp = self.stamp_position is not None and has_pdf
        busy_state = tk.DISABLED if self.busy else tk.NORMAL
        self.add_button.configure(state=busy_state)
        self.name_entry.configure(state=busy_state)
        self.date_entry.configure(state=busy_state)
        self.clear_button.configure(state=tk.NORMAL if not self.busy and (self.pdf_path is not None or self.stamp_position is not None) else tk.DISABLED)
        self.save_button.configure(state=tk.NORMAL if not self.busy and has_stamp else tk.DISABLED)
        self.prev_button.configure(state=tk.NORMAL if not self.busy and has_pdf and self.current_page_index > 0 else tk.DISABLED)
        self.next_button.configure(state=tk.NORMAL if not self.busy and has_pdf and self.current_page_index < self.page_count - 1 else tk.DISABLED)
        self.update_page_label()

    def set_status(self, status_key: str) -> None:
        self.status_var.set(UI_TEXT[status_key])
        bg_key, fg_key = STATUS_COLORS.get(status_key, STATUS_COLORS["status_idle"])
        self.status_badge.configure(bg=THEME[bg_key], fg=THEME[fg_key])

    def show_error(self, message: str) -> None:
        self.set_status("status_error")
        messagebox.showerror(UI_TEXT["dialog_error_title"], message)

    def open_folder(self, folder: Path) -> None:
        try:
            os.startfile(str(folder))
        except Exception:
            pass

    def run(self) -> None:
        self.root.mainloop()


def main() -> None:
    app = CheckStampApp()
    app.run()


if __name__ == "__main__":
    main()
