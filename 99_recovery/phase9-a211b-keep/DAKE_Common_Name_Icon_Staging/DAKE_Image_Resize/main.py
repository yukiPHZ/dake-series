# -*- coding: utf-8 -*-
from __future__ import annotations

import io
import os
import queue
import threading
import webbrowser
from collections import OrderedDict
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, ttk

from PIL import Image, ImageFilter, ImageOps

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    ROOT_CLASS = TkinterDnD.Tk
    DND_ENABLED = True
except ImportError:
    DND_FILES = None
    ROOT_CLASS = tk.Tk
    DND_ENABLED = False


APP_NAME = "画像リサイズ"
WINDOW_TITLE = "画像リサイズ"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "main_title": "画像をリサイズする",
    "main_description": "スマホ写真を、軽くて使いやすいサイズに整えます。",
    "button_add": "画像を追加",
    "button_select_folder": "保存先を選ぶ",
    "button_refresh": "リフレッシュ",
    "button_execute": "リサイズして保存",
    "empty_title": "画像を追加してください",
    "empty_subtitle": "ドラッグ＆ドロップ または クリックして追加",
    "status_idle": "未選択",
    "status_loading": "読み込み中",
    "status_loading_dot": "読み込み中",
    "status_ready": "準備完了",
    "status_processing": "リサイズ中",
    "status_resizing": "リサイズ中",
    "status_saving": "保存中",
    "status_complete": "保存完了",
    "status_error": "エラー",
    "footer_left": "シンプルそれDAKEシリーズ / 止まらない、迷わない、すぐ終わる。",
    "footer_link_1": "戸建買取査定",
    "footer_link_2": "Instagram",
    "footer_separator": " ｜ ",
    "footer_copyright": COPYRIGHT,
    "dialog_complete_title": "保存完了",
    "dialog_complete_message": "画像のリサイズが完了しました。",
    "dialog_error_title": "エラー",
    "dialog_error_message": "処理できない画像がありました。",
    "dialog_no_files_title": "確認",
    "dialog_no_files_message": "画像がまだ追加されていません。",
    "dialog_busy_title": "処理中",
    "dialog_busy_message": "処理が終わるまでお待ちください。",
    "dialog_open_folder_note": "OK後に出力フォルダを開きます。",
    "complete_summary": "成功: {success}枚 / 失敗: {failed}枚",
    "failed_list_title": "処理できなかった画像",
    "failed_list_item": "・ {name}",
    "output_folder_name": "DakeImageResize_Output",
    "save_location_default": "保存先: 元画像フォルダ内",
    "save_location_selected": "保存先: 選択フォルダ内",
    "save_location_selected_detail": "保存先: 選択フォルダ内 / {folder}",
    "progress_template": "{current} / {total} 枚",
    "progress_idle": "0 / 0 枚",
    "row_initial_template": "{folder}　{size}",
    "row_result_template": "{folder}　{before} → {after}",
    "file_dialog_title": "画像を選択",
    "folder_dialog_title": "保存先を選択",
    "filetype_images": "画像ファイル",
    "filetype_all": "すべてのファイル",
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
    "idle_bg": "#F2F4F7",
    "scrollbar_thumb": "#CDD5DF",
    "scrollbar_thumb_hover": "#B7C3D4",
}

LINK_URLS = {
    "footer_link_1": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "footer_link_2": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

SUPPORTED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp"}
JPEG_QUALITY_STEPS = [95, 92, 90, 88, 86, 84, 82, 80, 78, 76, 74, 72, 70, 68, 66, 64, 62, 60, 58, 56, 54, 52, 50]
MAX_LONG_EDGE = 1600
TARGET_MIN_BYTES = 500 * 1024
TARGET_MAX_BYTES = 2 * 1024 * 1024
HARD_MAX_BYTES = 3 * 1024 * 1024
QUEUE_POLL_INTERVAL_MS = 100
STATUS_ANIMATION_INTERVAL_MS = 400
ROW_RENDER_BATCH_SIZE = 40
ROW_NAME_CHAR_LIMIT = 28
ROW_PATH_CHAR_LIMIT = 56
SAVE_PATH_CHAR_LIMIT = 68
STATUS_BADGE_WIDTH = 14
ROW_STATUS_WIDTH = 10
OUTPUT_FOLDER_NAME = UI_TEXT["output_folder_name"]
ANIMATED_STATUS_KEYS = {"status_loading_dot", "status_resizing", "status_saving"}
RESAMPLE = Image.Resampling.LANCZOS if hasattr(Image, "Resampling") else Image.LANCZOS
BASE_DIR = Path(__file__).resolve().parent
ICON_PATH = (BASE_DIR / "../../02_assets/dake_icon.ico").resolve()


def choose_font_family(root: tk.Tk) -> str:
    preferred = ["BIZ UDPGothic", "Yu Gothic UI", "Meiryo"]
    available = set(tkfont.families(root))
    for family in preferred:
        if family in available:
            return family
    return "TkDefaultFont"


def format_bytes(size: int) -> str:
    if size >= 1024 * 1024:
        return f"{size / (1024 * 1024):.2f} MB"
    if size >= 1024:
        return f"{size / 1024:.0f} KB"
    return f"{size} B"


def truncate_middle(text: str, max_chars: int) -> str:
    if len(text) <= max_chars:
        return text
    if max_chars <= 5:
        return text[:max_chars]

    front = max_chars // 2
    back = max_chars - front - 3
    return f"{text[:front]}...{text[-back:]}"


def normalize_image_for_jpeg(image: Image.Image) -> Image.Image:
    if image.mode == "RGB":
        return image

    if "A" in image.getbands():
        rgba_image = image.convert("RGBA")
        background = Image.new("RGB", rgba_image.size, (255, 255, 255))
        background.paste(rgba_image, mask=rgba_image.getchannel("A"))
        return background

    return image.convert("RGB")


def resize_image(image: Image.Image) -> tuple[Image.Image, bool]:
    width, height = image.size
    long_edge = max(width, height)
    if long_edge <= MAX_LONG_EDGE:
        return image, False

    scale = MAX_LONG_EDGE / float(long_edge)
    new_size = (max(1, round(width * scale)), max(1, round(height * scale)))
    return image.resize(new_size, RESAMPLE), True


def encode_jpeg(image: Image.Image) -> tuple[bytes, int, int]:
    working_image = image

    while True:
        best_under_cap: tuple[bytes, int, int] | None = None
        last_attempt: tuple[bytes, int, int] | None = None

        for quality in JPEG_QUALITY_STEPS:
            buffer = io.BytesIO()
            working_image.save(buffer, format="JPEG", quality=quality, optimize=True, progressive=True)
            data = buffer.getvalue()
            size = len(data)
            last_attempt = (data, quality, size)

            if best_under_cap is None and size <= HARD_MAX_BYTES:
                best_under_cap = (data, quality, size)

            if TARGET_MIN_BYTES <= size <= TARGET_MAX_BYTES:
                return data, quality, size

        if best_under_cap is not None:
            return best_under_cap

        if last_attempt is None:
            raise RuntimeError("JPEG encode failed")

        long_edge = max(working_image.size)
        if long_edge <= 960:
            return last_attempt

        scale = 0.95
        new_size = (
            max(1, round(working_image.size[0] * scale)),
            max(1, round(working_image.size[1] * scale)),
        )
        working_image = working_image.resize(new_size, RESAMPLE)


def resolve_output_folder_path(source_path: Path, selected_output_root: Path | None) -> Path:
    base_root = selected_output_root if selected_output_root is not None else source_path.parent
    return base_root / OUTPUT_FOLDER_NAME


def build_output_path(source_path: Path, destination_dir: Path) -> Path:
    base_name = f"{source_path.stem}_resizeDake"
    candidate = destination_dir / f"{base_name}.jpg"
    sequence = 2

    while candidate.exists():
        candidate = destination_dir / f"{base_name}_{sequence}.jpg"
        sequence += 1

    return candidate


def open_output_folder(folder_path: Path) -> None:
    if os.name == "nt":
        os.startfile(str(folder_path))
        return
    webbrowser.open(folder_path.as_uri())


class ImageResizeApp:
    def __init__(self) -> None:
        self.root = ROOT_CLASS()
        self.root.title(WINDOW_TITLE)
        self.root.geometry("980x700")
        self.root.minsize(900, 620)
        self.root.configure(bg=COLORS["base_bg"])

        self.font_family = choose_font_family(self.root)
        self.event_queue: queue.Queue[dict] = queue.Queue()
        self.files: OrderedDict[str, dict] = OrderedDict()
        self.file_widgets: dict[str, dict] = {}
        self.selected_output_root: Path | None = None
        self.loading_files = False
        self.processing = False
        self.status_animation_key: str | None = None
        self.status_animation_phase = 1
        self.status_animation_job: str | None = None

        self.output_var = tk.StringVar(value=UI_TEXT["save_location_default"])
        self.progress_var = tk.StringVar(value=UI_TEXT["progress_idle"])

        self.apply_window_icon()
        self.configure_styles()
        self.build_ui()
        self.update_output_label()
        self.update_empty_state()
        self.set_overall_status("status_idle")
        self.update_button_state()

        self.root.after(QUEUE_POLL_INTERVAL_MS, self.poll_queue)
        self.root.protocol("WM_DELETE_WINDOW", self.handle_close)

    def apply_window_icon(self) -> None:
        if ICON_PATH.exists():
            try:
                self.root.iconbitmap(str(ICON_PATH))
            except Exception:
                pass

    def configure_styles(self) -> None:
        style = ttk.Style()
        style.theme_use("clam")

        style.configure(
            "Primary.TButton",
            font=(self.font_family, 11, "bold"),
            padding=(18, 10),
            background=COLORS["accent"],
            foreground="#FFFFFF",
            borderwidth=0,
            relief="flat",
        )
        style.map(
            "Primary.TButton",
            background=[
                ("active", COLORS["accent_hover"]),
                ("disabled", COLORS["border"]),
            ],
            foreground=[("disabled", COLORS["muted"])],
        )

        style.configure(
            "Secondary.TButton",
            font=(self.font_family, 10),
            padding=(16, 10),
            background=COLORS["card_bg"],
            foreground=COLORS["text"],
            borderwidth=1,
            relief="solid",
        )
        style.map(
            "Secondary.TButton",
            background=[
                ("active", COLORS["selection_bg"]),
                ("disabled", COLORS["idle_bg"]),
            ],
            foreground=[("disabled", COLORS["muted"])],
            bordercolor=[
                ("!disabled", COLORS["border"]),
                ("disabled", COLORS["border"]),
            ],
        )

        style.configure(
            "Resize.Horizontal.TProgressbar",
            troughcolor=COLORS["idle_bg"],
            background=COLORS["accent"],
            bordercolor=COLORS["idle_bg"],
            lightcolor=COLORS["accent"],
            darkcolor=COLORS["accent"],
        )

        style.configure(
            "Resize.Vertical.TScrollbar",
            background=COLORS["scrollbar_thumb"],
            darkcolor=COLORS["scrollbar_thumb"],
            lightcolor=COLORS["scrollbar_thumb"],
            troughcolor=COLORS["idle_bg"],
            bordercolor=COLORS["idle_bg"],
            arrowcolor=COLORS["muted"],
            relief="flat",
            width=14,
        )
        style.map(
            "Resize.Vertical.TScrollbar",
            background=[("active", COLORS["scrollbar_thumb_hover"])],
            darkcolor=[("active", COLORS["scrollbar_thumb_hover"])],
            lightcolor=[("active", COLORS["scrollbar_thumb_hover"])],
        )

    def build_ui(self) -> None:
        outer = tk.Frame(self.root, bg=COLORS["base_bg"])
        outer.pack(fill="both", expand=True, padx=24, pady=(24, 18))

        self.build_header(outer)
        self.build_control_card(outer)
        self.build_list_card(outer)
        self.build_footer(outer)

    def build_header(self, parent: tk.Frame) -> None:
        header = tk.Frame(parent, bg=COLORS["base_bg"])
        header.pack(fill="x")

        text_row = tk.Frame(header, bg=COLORS["base_bg"])
        text_row.pack(fill="x")
        text_row.grid_columnconfigure(1, weight=1)
        text_row.bind("<Configure>", self.update_header_description_wrap)

        tk.Label(
            text_row,
            text=UI_TEXT["main_title"],
            bg=COLORS["base_bg"],
            fg=COLORS["text"],
            font=(self.font_family, 24, "bold"),
        ).grid(row=0, column=0, sticky="w")

        self.header_description_label = tk.Label(
            text_row,
            text=UI_TEXT["main_description"],
            bg=COLORS["base_bg"],
            fg=COLORS["muted"],
            justify="left",
            anchor="w",
            wraplength=520,
            font=(self.font_family, 11),
        )
        self.header_description_label.grid(row=0, column=1, sticky="ew", padx=(16, 0), pady=(7, 0))

    def update_header_description_wrap(self, event) -> None:
        wraplength = max(220, event.width - 320)
        self.header_description_label.configure(wraplength=wraplength)

    def build_control_card(self, parent: tk.Frame) -> None:
        card = tk.Frame(
            parent,
            bg=COLORS["card_bg"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
            bd=0,
        )
        card.pack(fill="x", pady=(18, 14))

        button_row = tk.Frame(card, bg=COLORS["card_bg"])
        button_row.pack(fill="x", padx=20, pady=(16, 10))

        self.add_button = ttk.Button(
            button_row,
            text=UI_TEXT["button_add"],
            style="Secondary.TButton",
            command=self.open_file_dialog,
        )
        self.add_button.pack(side="left")

        self.folder_button = ttk.Button(
            button_row,
            text=UI_TEXT["button_select_folder"],
            style="Secondary.TButton",
            command=self.select_output_folder,
        )
        self.folder_button.pack(side="left", padx=(10, 0))

        self.execute_button = ttk.Button(
            button_row,
            text=UI_TEXT["button_execute"],
            style="Primary.TButton",
            command=self.start_processing,
        )
        self.execute_button.pack(side="left", padx=(10, 0))

        self.refresh_button = ttk.Button(
            button_row,
            text=UI_TEXT["button_refresh"],
            style="Secondary.TButton",
            command=self.refresh,
        )
        self.refresh_button.pack(side="left", padx=(10, 0))

        tk.Label(
            card,
            textvariable=self.output_var,
            bg=COLORS["card_bg"],
            fg=COLORS["muted"],
            anchor="w",
            justify="left",
            font=(self.font_family, 10),
        ).pack(fill="x", padx=20, pady=(0, 16))

    def build_list_card(self, parent: tk.Frame) -> None:
        card = tk.Frame(
            parent,
            bg=COLORS["card_bg"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
            bd=0,
        )
        card.pack(fill="both", expand=True)

        list_area = tk.Frame(card, bg=COLORS["card_bg"])
        list_area.pack(fill="both", expand=True, padx=18, pady=(18, 12))

        self.list_canvas = tk.Canvas(
            list_area,
            bg=COLORS["card_bg"],
            highlightthickness=0,
            bd=0,
        )
        self.list_canvas.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(
            list_area,
            orient="vertical",
            command=self.list_canvas.yview,
            style="Resize.Vertical.TScrollbar",
        )
        scrollbar.pack(side="right", fill="y")
        self.list_canvas.configure(yscrollcommand=scrollbar.set)

        self.list_frame = tk.Frame(self.list_canvas, bg=COLORS["card_bg"])
        self.list_window = self.list_canvas.create_window((0, 0), window=self.list_frame, anchor="nw")

        self.list_frame.bind("<Configure>", self.update_scrollregion)
        self.list_canvas.bind("<Configure>", self.fit_list_width)

        self.empty_state = tk.Frame(
            list_area,
            bg=COLORS["card_bg"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
            bd=0,
        )

        empty_inner = tk.Frame(self.empty_state, bg=COLORS["card_bg"])
        empty_inner.place(relx=0.5, rely=0.5, anchor="center")

        empty_title = tk.Label(
            empty_inner,
            text=UI_TEXT["empty_title"],
            bg=COLORS["card_bg"],
            fg=COLORS["text"],
            font=(self.font_family, 18, "bold"),
        )
        empty_title.pack()

        empty_subtitle = tk.Label(
            empty_inner,
            text=UI_TEXT["empty_subtitle"],
            bg=COLORS["card_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 11),
        )
        empty_subtitle.pack(pady=(10, 0))

        for widget in (self.empty_state, empty_inner, empty_title, empty_subtitle):
            widget.bind("<Button-1>", lambda _event: self.open_file_dialog())

        self.register_drop_target(list_area)
        self.register_drop_target(self.list_canvas)
        self.register_drop_target(self.list_frame)
        self.register_drop_target(self.empty_state)

        status_row = tk.Frame(card, bg=COLORS["card_bg"])
        status_row.pack(fill="x", padx=18, pady=(0, 18))

        left_status = tk.Frame(status_row, bg=COLORS["card_bg"])
        left_status.pack(side="left")

        self.status_badge = tk.Label(
            left_status,
            text="",
            width=STATUS_BADGE_WIDTH,
            anchor="center",
            bg=COLORS["idle_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 10, "bold"),
            padx=10,
            pady=5,
        )
        self.status_badge.pack(side="left")

        self.progress_label = tk.Label(
            left_status,
            textvariable=self.progress_var,
            bg=COLORS["card_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 10),
        )
        self.progress_label.pack(side="left", padx=(12, 0))

        self.progressbar = ttk.Progressbar(
            status_row,
            mode="determinate",
            style="Resize.Horizontal.TProgressbar",
            length=280,
        )
        self.progressbar.pack(side="right", fill="x", expand=True)
        self.progressbar.configure(maximum=1, value=0)

    def build_footer(self, parent: tk.Frame) -> None:
        footer = tk.Frame(parent, bg=COLORS["base_bg"])
        footer.pack(fill="x", pady=(14, 0))

        tk.Label(
            footer,
            text=UI_TEXT["footer_left"],
            bg=COLORS["base_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 9),
        ).pack(side="left")

        right = tk.Frame(footer, bg=COLORS["base_bg"])
        right.pack(side="right")

        self.create_footer_link(right, "footer_link_1")
        self.create_footer_text(right, "footer_separator")
        self.create_footer_link(right, "footer_link_2")
        self.create_footer_text(right, "footer_separator")
        self.create_footer_text(right, "footer_copyright")

    def create_footer_link(self, parent: tk.Frame, text_key: str) -> None:
        label = tk.Label(
            parent,
            text=UI_TEXT[text_key],
            bg=COLORS["base_bg"],
            fg=COLORS["muted"],
            cursor="hand2",
            font=(self.font_family, 9),
        )
        label.pack(side="left")
        label.bind("<Button-1>", lambda _event, key=text_key: webbrowser.open(LINK_URLS[key]))
        label.bind("<Enter>", lambda _event, target=label: target.configure(fg=COLORS["accent"]))
        label.bind("<Leave>", lambda _event, target=label: target.configure(fg=COLORS["muted"]))

    def create_footer_text(self, parent: tk.Frame, text_key: str) -> None:
        tk.Label(
            parent,
            text=UI_TEXT[text_key],
            bg=COLORS["base_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 9),
        ).pack(side="left")

    def register_drop_target(self, widget: tk.Misc) -> None:
        if not DND_ENABLED:
            return

        widget.drop_target_register(DND_FILES)
        widget.dnd_bind("<<Drop>>", self.handle_drop)

    def handle_drop(self, event) -> None:
        self.begin_add_files(self.root.tk.splitlist(event.data))

    def open_file_dialog(self) -> None:
        if self.loading_files or self.processing:
            return

        filetypes = [
            (UI_TEXT["filetype_images"], "*.jpg *.jpeg *.png *.bmp"),
            (UI_TEXT["filetype_all"], "*.*"),
        ]
        selected = filedialog.askopenfilenames(
            title=UI_TEXT["file_dialog_title"],
            filetypes=filetypes,
            parent=self.root,
        )
        if selected:
            self.begin_add_files(selected)

    def select_output_folder(self) -> None:
        if self.loading_files or self.processing:
            return

        selected = filedialog.askdirectory(
            title=UI_TEXT["folder_dialog_title"],
            parent=self.root,
        )
        if selected:
            self.selected_output_root = Path(selected)
            self.update_output_label()
            self.refresh_row_details()

    def update_output_label(self) -> None:
        if self.selected_output_root is None:
            self.output_var.set(UI_TEXT["save_location_default"])
            return

        output_folder = resolve_output_folder_path(Path("."), self.selected_output_root)
        display_path = truncate_middle(str(output_folder), SAVE_PATH_CHAR_LIMIT)
        self.output_var.set(
            UI_TEXT["save_location_selected_detail"].format(folder=display_path)
        )

    def begin_add_files(self, raw_paths) -> None:
        if self.loading_files or self.processing:
            return

        incoming_paths = list(raw_paths)
        if not incoming_paths:
            return

        self.loading_files = True
        self.update_button_state()
        self.set_overall_status("status_loading_dot")

        worker = threading.Thread(
            target=self.load_files_worker,
            args=(incoming_paths, set(self.files.keys())),
            daemon=True,
        )
        worker.start()

    def load_files_worker(self, raw_paths: list[str], existing_keys: set[str]) -> None:
        loaded_records: list[dict] = []
        known_keys = set(existing_keys)

        for raw_path in raw_paths:
            try:
                path = Path(raw_path)
                if not path.exists() or path.is_dir():
                    continue
                if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
                    continue

                resolved = path.resolve()
                path_key = str(resolved)
                if path_key in known_keys:
                    continue

                loaded_records.append(
                    {
                        "path": resolved,
                        "path_key": path_key,
                        "input_size": resolved.stat().st_size,
                        "status_key": "status_ready",
                        "output_size": None,
                        "output_dir": None,
                    }
                )
                known_keys.add(path_key)
            except Exception:
                continue

        self.event_queue.put({"type": "files_loaded", "records": loaded_records})

    def render_loaded_file_rows(self, records: list[dict], start_index: int = 0) -> None:
        if not self.root.winfo_exists():
            return

        end_index = min(start_index + ROW_RENDER_BATCH_SIZE, len(records))
        for record in records[start_index:end_index]:
            self.files[record["path_key"]] = record
            self.create_file_row(record["path_key"])

        self.update_empty_state()

        if end_index < len(records):
            self.root.after(1, lambda: self.render_loaded_file_rows(records, end_index))
            return

        self.finish_loading_files()

    def finish_loading_files(self) -> None:
        self.loading_files = False
        self.update_button_state()
        self.progressbar.configure(maximum=max(1, len(self.files)), value=0)
        self.progress_var.set(UI_TEXT["progress_idle"])
        self.set_overall_status("status_ready" if self.files else "status_idle")
        self.update_empty_state()

    def create_file_row(self, path_key: str) -> None:
        record = self.files[path_key]
        path = record["path"]

        row = tk.Frame(
            self.list_frame,
            bg=COLORS["card_bg"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
            bd=0,
            padx=12,
            pady=9,
        )
        row.pack(fill="x", pady=(0, 8))
        row.grid_columnconfigure(1, weight=1)

        name_label = tk.Label(
            row,
            text=truncate_middle(path.name, ROW_NAME_CHAR_LIMIT),
            width=ROW_NAME_CHAR_LIMIT,
            bg=COLORS["card_bg"],
            fg=COLORS["text"],
            anchor="w",
            justify="left",
            font=(self.font_family, 10, "bold"),
        )
        name_label.grid(row=0, column=0, sticky="w")

        detail_var = tk.StringVar(value=self.build_row_detail(record))
        detail_label = tk.Label(
            row,
            textvariable=detail_var,
            bg=COLORS["card_bg"],
            fg=COLORS["muted"],
            anchor="w",
            justify="left",
            font=(self.font_family, 9),
        )
        detail_label.grid(row=0, column=1, sticky="ew", padx=(12, 10))

        status_label = tk.Label(
            row,
            text="",
            width=ROW_STATUS_WIDTH,
            anchor="center",
            bg=COLORS["idle_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 9, "bold"),
            padx=8,
            pady=4,
        )
        status_label.grid(row=0, column=2, sticky="e")

        self.file_widgets[path_key] = {
            "detail_var": detail_var,
            "status_label": status_label,
        }
        self.set_file_status(path_key, "status_ready")

    def build_row_detail(self, record: dict) -> str:
        output_dir = record.get("output_dir")
        if output_dir is None:
            output_dir = resolve_output_folder_path(record["path"], self.selected_output_root)

        folder_text = truncate_middle(str(output_dir), ROW_PATH_CHAR_LIMIT)
        before_text = format_bytes(record["input_size"])
        output_size = record.get("output_size")

        if output_size is not None and record.get("status_key") == "status_complete":
            return UI_TEXT["row_result_template"].format(
                folder=folder_text,
                before=before_text,
                after=format_bytes(output_size),
            )

        return UI_TEXT["row_initial_template"].format(folder=folder_text, size=before_text)

    def refresh_row_details(self) -> None:
        for path_key, record in self.files.items():
            widget_map = self.file_widgets.get(path_key)
            if widget_map:
                widget_map["detail_var"].set(self.build_row_detail(record))

    def set_file_status(self, path_key: str, status_key: str) -> None:
        record = self.files.get(path_key)
        widget_map = self.file_widgets.get(path_key)
        if record is None or widget_map is None:
            return

        record["status_key"] = status_key
        background, foreground = self.get_status_colors(status_key)
        widget_map["status_label"].configure(
            text=UI_TEXT[status_key],
            bg=background,
            fg=foreground,
        )

    def set_overall_status(self, status_key: str) -> None:
        self.stop_status_animation()
        if status_key in ANIMATED_STATUS_KEYS:
            self.status_animation_key = status_key
            self.status_animation_phase = 2
            self.render_status_badge(status_key, 1)
            self.status_animation_job = self.root.after(
                STATUS_ANIMATION_INTERVAL_MS,
                self.animate_status_badge,
            )
            return

        self.status_animation_key = None
        self.render_status_badge(status_key, None)

    def render_status_badge(self, status_key: str, dot_count: int | None) -> None:
        background, foreground = self.get_status_colors(status_key)
        text = UI_TEXT[status_key]
        if dot_count is not None:
            text = f"{text}{'.' * dot_count}"

        self.status_badge.configure(
            text=text,
            bg=background,
            fg=foreground,
        )

    def animate_status_badge(self) -> None:
        if self.status_animation_key is None or not self.root.winfo_exists():
            return

        self.render_status_badge(self.status_animation_key, self.status_animation_phase)
        self.status_animation_phase = 1 if self.status_animation_phase >= 3 else self.status_animation_phase + 1
        self.status_animation_job = self.root.after(
            STATUS_ANIMATION_INTERVAL_MS,
            self.animate_status_badge,
        )

    def stop_status_animation(self) -> None:
        if self.status_animation_job is not None:
            self.root.after_cancel(self.status_animation_job)
            self.status_animation_job = None

    def get_status_colors(self, status_key: str) -> tuple[str, str]:
        if status_key == "status_complete":
            return COLORS["success_bg"], COLORS["success"]
        if status_key == "status_error":
            return COLORS["error_bg"], COLORS["error"]
        if status_key in {
            "status_loading",
            "status_loading_dot",
            "status_ready",
            "status_processing",
            "status_resizing",
            "status_saving",
        }:
            return COLORS["selection_bg"], COLORS["accent"]
        return COLORS["idle_bg"], COLORS["muted"]

    def update_empty_state(self) -> None:
        if self.files:
            self.empty_state.place_forget()
            return

        self.empty_state.place(relx=0, rely=0, relwidth=1, relheight=1)

    def update_scrollregion(self, _event=None) -> None:
        self.list_canvas.configure(scrollregion=self.list_canvas.bbox("all"))

    def fit_list_width(self, event) -> None:
        self.list_canvas.itemconfigure(self.list_window, width=event.width)

    def update_button_state(self) -> None:
        state = "disabled" if self.loading_files or self.processing else "normal"
        for button in (
            self.add_button,
            self.folder_button,
            self.execute_button,
            self.refresh_button,
        ):
            button.configure(state=state)

    def refresh(self) -> None:
        if self.loading_files or self.processing:
            return

        self.selected_output_root = None
        self.update_output_label()
        self.files.clear()
        self.file_widgets.clear()

        for child in self.list_frame.winfo_children():
            child.destroy()

        self.progress_var.set(UI_TEXT["progress_idle"])
        self.progressbar.configure(value=0, maximum=1)
        self.set_overall_status("status_idle")
        self.update_empty_state()

    def start_processing(self) -> None:
        if self.processing or self.loading_files:
            messagebox.showinfo(
                UI_TEXT["dialog_busy_title"],
                UI_TEXT["dialog_busy_message"],
                parent=self.root,
            )
            return

        if not self.files:
            messagebox.showinfo(
                UI_TEXT["dialog_no_files_title"],
                UI_TEXT["dialog_no_files_message"],
                parent=self.root,
            )
            return

        self.processing = True
        self.update_button_state()
        total = len(self.files)
        self.progress_var.set(UI_TEXT["progress_template"].format(current=0, total=total))
        self.progressbar.configure(maximum=total, value=0)

        worker = threading.Thread(
            target=self.process_files_worker,
            args=(list(self.files.values()), self.selected_output_root),
            daemon=True,
        )
        worker.start()

    def process_files_worker(self, file_records: list[dict], output_root: Path | None) -> None:
        successes: list[dict] = []
        failures: list[dict] = []
        output_folders: list[str] = []
        total = len(file_records)

        for index, record in enumerate(file_records, start=1):
            source_path: Path = record["path"]
            path_key = record["path_key"]
            self.event_queue.put({"type": "overall_status", "status_key": "status_resizing"})
            self.event_queue.put({"type": "file_status", "path_key": path_key, "status_key": "status_processing"})

            try:
                processed_image, resized = self.prepare_image(source_path)
                if resized:
                    processed_image = processed_image.filter(
                        ImageFilter.UnsharpMask(radius=0.6, percent=90, threshold=2)
                    )

                self.event_queue.put({"type": "overall_status", "status_key": "status_saving"})
                self.event_queue.put({"type": "file_status", "path_key": path_key, "status_key": "status_saving"})

                target_dir = resolve_output_folder_path(source_path, output_root)
                target_dir.mkdir(parents=True, exist_ok=True)
                output_path = build_output_path(source_path, target_dir)
                jpeg_bytes, quality, output_size = encode_jpeg(processed_image)
                output_path.write_bytes(jpeg_bytes)

                if str(target_dir) not in output_folders:
                    output_folders.append(str(target_dir))

                successes.append(
                    {
                        "source_path": str(source_path),
                        "output_path": str(output_path),
                        "output_size": output_size,
                        "quality": quality,
                    }
                )

                self.event_queue.put(
                    {
                        "type": "file_done",
                        "path_key": path_key,
                        "status_key": "status_complete",
                        "output_dir": target_dir,
                        "output_size": output_size,
                    }
                )
            except Exception as exc:
                failures.append({"source_path": str(source_path), "error": str(exc)})
                self.event_queue.put(
                    {
                        "type": "file_done",
                        "path_key": path_key,
                        "status_key": "status_error",
                    }
                )

            self.event_queue.put({"type": "progress", "current": index, "total": total})

        final_status = "status_complete" if successes else "status_error"
        self.event_queue.put(
            {
                "type": "done",
                "status_key": final_status,
                "result": {
                    "successes": successes,
                    "failures": failures,
                    "output_folders": output_folders,
                },
            }
        )

    def prepare_image(self, source_path: Path) -> tuple[Image.Image, bool]:
        with Image.open(source_path) as image:
            image.load()
            transposed = ImageOps.exif_transpose(image)
            prepared = normalize_image_for_jpeg(transposed)
            resized_image, was_resized = resize_image(prepared)
            return resized_image, was_resized

    def poll_queue(self) -> None:
        try:
            while True:
                event = self.event_queue.get_nowait()
                self.handle_queue_event(event)
        except queue.Empty:
            pass
        finally:
            if self.root.winfo_exists():
                self.root.after(QUEUE_POLL_INTERVAL_MS, self.poll_queue)

    def handle_queue_event(self, event: dict) -> None:
        event_type = event["type"]

        if event_type == "files_loaded":
            self.render_loaded_file_rows(event["records"])
            return

        if event_type == "overall_status":
            self.set_overall_status(event["status_key"])
            return

        if event_type == "file_status":
            self.set_file_status(event["path_key"], event["status_key"])
            return

        if event_type == "file_done":
            record = self.files.get(event["path_key"])
            if record is not None:
                if "output_dir" in event:
                    record["output_dir"] = event["output_dir"]
                if "output_size" in event:
                    record["output_size"] = event["output_size"]
            self.set_file_status(event["path_key"], event["status_key"])
            widget_map = self.file_widgets.get(event["path_key"])
            if record is not None and widget_map is not None:
                widget_map["detail_var"].set(self.build_row_detail(record))
            return

        if event_type == "progress":
            current = event["current"]
            total = event["total"]
            self.progress_var.set(UI_TEXT["progress_template"].format(current=current, total=total))
            self.progressbar.configure(value=current)
            return

        if event_type == "done":
            self.processing = False
            self.update_button_state()
            self.set_overall_status(event["status_key"])
            self.show_result_dialog(event["result"])

    def show_result_dialog(self, result: dict) -> None:
        success_count = len(result["successes"])
        failure_count = len(result["failures"])
        title = UI_TEXT["dialog_complete_title"] if success_count else UI_TEXT["dialog_error_title"]
        base_message = UI_TEXT["dialog_complete_message"] if success_count else UI_TEXT["dialog_error_message"]

        message_lines = [
            base_message,
            "",
            UI_TEXT["complete_summary"].format(success=success_count, failed=failure_count),
        ]

        if failure_count:
            message_lines.extend(["", UI_TEXT["failed_list_title"]])
            message_lines.extend(
                UI_TEXT["failed_list_item"].format(name=Path(item["source_path"]).name)
                for item in result["failures"]
            )

        if result["output_folders"]:
            message_lines.extend(["", UI_TEXT["dialog_open_folder_note"]])

        message = "\n".join(message_lines)

        if failure_count and success_count:
            messagebox.showwarning(title, message, parent=self.root)
        elif failure_count:
            messagebox.showerror(title, message, parent=self.root)
        else:
            messagebox.showinfo(title, message, parent=self.root)

        for folder in result["output_folders"]:
            open_output_folder(Path(folder))

    def handle_close(self) -> None:
        if self.loading_files or self.processing:
            messagebox.showinfo(
                UI_TEXT["dialog_busy_title"],
                UI_TEXT["dialog_busy_message"],
                parent=self.root,
            )
            return

        self.stop_status_animation()
        self.root.destroy()

    def run(self) -> None:
        self.root.mainloop()


if __name__ == "__main__":
    ImageResizeApp().run()
