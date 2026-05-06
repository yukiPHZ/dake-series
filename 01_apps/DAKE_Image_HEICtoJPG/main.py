# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import queue
import sys
import threading
import time
import webbrowser
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import tkinter as tk
from tkinter import filedialog, font as tkfont

from PIL import Image, ImageOps, UnidentifiedImageError

try:
    import pillow_heif

    pillow_heif.register_heif_opener()
    HEIF_READY = True
except Exception:
    HEIF_READY = False

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    ROOT_CLASS = TkinterDnD.Tk
    DND_READY = True
except Exception:
    DND_FILES = ""
    ROOT_CLASS = tk.Tk
    DND_READY = False


APP_NAME = "HEIC→JPG変換"
WINDOW_TITLE = APP_NAME
EXE_NAME = "DakeHEIC_JPG.exe"
INTERNAL_NAME = "DakeHEIC_JPG"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"
BRAND_SERIES = "シンプルそれDAKEシリーズ"
MAIN_DESCRIPTION = "HEIC画像をドロップするだけでJPGに変換します。"

UI_TEXT = {
    "brand_series": BRAND_SERIES,
    "header_subtitle": MAIN_DESCRIPTION,
    "main_title": "HEICをJPGに変換する",
    "main_description": MAIN_DESCRIPTION,
    "empty_title_drop": "HEICをドロップしてください",
    "empty_subtitle": "ドラッグ＆ドロップ または クリックして選択",
    "empty_subtitle_click_only": "クリックして選択",
    "loading_title": "読み込んでいます",
    "loading_subtitle": "HEIC画像を確認しています",
    "processing_title": "変換しています",
    "saving_title": "保存しています",
    "complete_title": "完了",
    "complete_summary": "{success}件変換 / {skipped}件スキップ",
    "progress_template": "{current} / {total} 件",
    "status_idle": "HEICを待っています",
    "status_loading": "読み込み中",
    "status_processing": "処理中",
    "status_saving": "保存中",
    "status_complete": "完了",
    "status_error": "エラー",
    "status_phrase_1": "Simple",
    "status_phrase_2": "Simple, fast",
    "status_phrase_3": "Simple, fast, for real work.",
    "dialog_select_title": "HEIC画像を選択してください",
    "file_type_heic": "HEIC画像",
    "file_type_all": "すべてのファイル",
    "footer_left": BRAND_SERIES,
    "footer_subtitle": "止まらない、迷わない、すぐ終わる。",
    "footer_link_1": "戸建買取査定",
    "footer_link_2": "Instagram",
    "footer_separator": " ｜ ",
    "footer_copyright": COPYRIGHT,
}

FOOTER_URLS = {
    "footer_link_1": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "footer_link_2": "https://www.instagram.com/kikuta.shimarisu_fudosan",
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
}

SUPPORTED_EXTENSIONS = {".heic", ".heif"}
JPEG_QUALITY = 95
WINDOW_SIZE = "900x540"
WINDOW_MIN_WIDTH = 860
WINDOW_MIN_HEIGHT = 500
QUEUE_POLL_INTERVAL_MS = 80
STATUS_ANIMATION_INTERVAL_MS = 420
LONG_STATUS_DELAY_SECONDS = 1.8
ANIMATED_STATUS_KEYS = {"status_loading", "status_processing", "status_saving"}
BASE_DIR = Path(__file__).resolve().parent


@dataclass(frozen=True)
class WorkerEvent:
    kind: str
    payload: dict[str, object]


@dataclass
class ConversionResult:
    success: int = 0
    skipped: int = 0
    total: int = 0


class HeicJpgConverter:
    def convert(self, source_path: Path, before_save=None) -> Path:
        output_path = make_available_jpg_path(source_path)

        with Image.open(source_path) as source_image:
            source_image.load()
            transposed_image = ImageOps.exif_transpose(source_image)
            exif_data = transposed_image.info.get("exif")
            icc_profile = transposed_image.info.get("icc_profile")
            rgb_image = normalize_to_rgb(transposed_image)

            try:
                save_options: dict[str, object] = {
                    "format": "JPEG",
                    "quality": JPEG_QUALITY,
                }
                if exif_data:
                    save_options["exif"] = exif_data
                if icc_profile:
                    save_options["icc_profile"] = icc_profile

                if before_save is not None:
                    before_save(output_path)
                rgb_image.save(output_path, **save_options)
            finally:
                if rgb_image is not transposed_image:
                    rgb_image.close()
                if transposed_image is not source_image:
                    transposed_image.close()

        return output_path


class DakeHeicJpgApp:
    def __init__(self) -> None:
        self.root = ROOT_CLASS()
        self.converter = HeicJpgConverter()
        self.event_queue: queue.Queue[WorkerEvent] = queue.Queue()
        self.processing = False

        self.status_key: str | None = None
        self.status_started_at = 0.0
        self.status_tick = 0
        self.status_after_id: str | None = None

        self.font_family = "Yu Gothic UI"
        self.drop_area: tk.Frame | None = None
        self.drop_inner: tk.Frame | None = None
        self.drop_title_label: tk.Label | None = None
        self.drop_subtitle_label: tk.Label | None = None

        self.drop_title_var = tk.StringVar()
        self.drop_subtitle_var = tk.StringVar()
        self.status_var = tk.StringVar(value=UI_TEXT["status_idle"])
        self.progress_var = tk.StringVar(value="")

        self.configure_window()
        self.font_family = choose_font_family(self.root)
        self.build_ui()
        self.render_idle()
        self.root.after(QUEUE_POLL_INTERVAL_MS, self.poll_worker_events)

    def configure_window(self) -> None:
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(WINDOW_MIN_WIDTH, WINDOW_MIN_HEIGHT)
        self.root.resizable(False, False)
        self.root.configure(bg=COLORS["base_bg"])
        self.root.protocol("WM_DELETE_WINDOW", self.close)
        self.apply_window_icon()

    def apply_window_icon(self) -> None:
        try:
            icon_path = get_common_icon_path()
            if icon_path.exists():
                self.root.iconbitmap(default=str(icon_path))
        except Exception:
            pass

    def build_ui(self) -> None:
        outer = tk.Frame(self.root, bg=COLORS["base_bg"], padx=28, pady=24)
        outer.pack(fill="both", expand=True)

        header = tk.Frame(outer, bg=COLORS["base_bg"])
        header.pack(fill="x", pady=(0, 18))

        tk.Label(
            header,
            text=UI_TEXT["main_title"],
            bg=COLORS["base_bg"],
            fg=COLORS["text"],
            font=(self.font_family, 20, "bold"),
        ).pack(anchor="w")

        tk.Label(
            header,
            text=UI_TEXT["header_subtitle"],
            bg=COLORS["base_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 10),
            anchor="w",
            justify="left",
        ).pack(fill="x", pady=(5, 0))

        self.drop_area = tk.Frame(
            outer,
            bg=COLORS["card_bg"],
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            highlightcolor=COLORS["accent"],
            bd=0,
            cursor="hand2",
        )
        self.drop_area.pack(fill="both", expand=True)

        self.drop_inner = tk.Frame(self.drop_area, bg=COLORS["card_bg"], padx=32, pady=32)
        self.drop_inner.place(relx=0.5, rely=0.5, anchor="center")

        self.drop_title_label = tk.Label(
            self.drop_inner,
            textvariable=self.drop_title_var,
            bg=COLORS["card_bg"],
            fg=COLORS["text"],
            font=(self.font_family, 24, "bold"),
        )
        self.drop_title_label.pack()

        self.drop_subtitle_label = tk.Label(
            self.drop_inner,
            textvariable=self.drop_subtitle_var,
            bg=COLORS["card_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 12),
            pady=12,
        )
        self.drop_subtitle_label.pack()

        status_row = tk.Frame(outer, bg=COLORS["base_bg"])
        status_row.pack(fill="x", pady=(12, 0))

        tk.Label(
            status_row,
            textvariable=self.status_var,
            bg=COLORS["base_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 10, "bold"),
        ).pack(side="left")

        tk.Label(
            status_row,
            textvariable=self.progress_var,
            bg=COLORS["base_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 10),
            padx=12,
        ).pack(side="left")

        self.build_footer(outer)
        self.bind_input_targets()

    def build_footer(self, parent: tk.Frame) -> None:
        footer = tk.Frame(parent, bg=COLORS["base_bg"])
        footer.pack(fill="x", side="bottom", pady=(10, 0))

        footer_top = tk.Frame(footer, bg=COLORS["base_bg"])
        footer_top.pack(anchor="center")

        for key in ("footer_left", "footer_separator", "footer_subtitle"):
            tk.Label(
                footer_top,
                text=UI_TEXT[key],
                bg=COLORS["base_bg"],
                fg=COLORS["muted"],
                font=(self.font_family, 8),
            ).pack(side="left")

        footer_bottom = tk.Frame(footer, bg=COLORS["base_bg"])
        footer_bottom.pack(anchor="center", pady=(3, 0))

        self.create_footer_link(footer_bottom, "footer_link_1")
        self.create_footer_text(footer_bottom, "footer_separator")
        self.create_footer_link(footer_bottom, "footer_link_2")
        self.create_footer_text(footer_bottom, "footer_separator")
        self.create_footer_text(footer_bottom, "footer_copyright")

    def create_footer_text(self, parent: tk.Frame, text_key: str) -> None:
        tk.Label(
            parent,
            text=UI_TEXT[text_key],
            bg=COLORS["base_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 8),
        ).pack(side="left")

    def create_footer_link(self, parent: tk.Frame, text_key: str) -> None:
        label = tk.Label(
            parent,
            text=UI_TEXT[text_key],
            bg=COLORS["base_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 8),
            cursor="hand2",
        )
        label.pack(side="left")
        label.bind("<Button-1>", lambda _event, key=text_key: webbrowser.open_new(FOOTER_URLS[key]))
        label.bind("<Enter>", lambda _event: label.configure(fg=COLORS["accent"]))
        label.bind("<Leave>", lambda _event: label.configure(fg=COLORS["muted"]))

    def bind_input_targets(self) -> None:
        targets = [
            self.drop_area,
            self.drop_inner,
            self.drop_title_label,
            self.drop_subtitle_label,
        ]

        for widget in targets:
            if widget is None:
                continue
            widget.bind("<Button-1>", self.open_file_dialog)

        self.register_drop_targets([widget for widget in targets if widget is not None])

    def register_drop_targets(self, widgets: list[tk.Widget]) -> None:
        if not DND_READY or not DND_FILES:
            return

        for widget in widgets:
            try:
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", self.handle_drop)
            except Exception:
                continue

    def open_file_dialog(self, _event=None) -> None:
        if self.processing:
            return

        selected_paths = filedialog.askopenfilenames(
            title=UI_TEXT["dialog_select_title"],
            filetypes=[
                (UI_TEXT["file_type_heic"], "*.heic *.heif"),
                (UI_TEXT["file_type_all"], "*.*"),
            ],
            parent=self.root,
        )
        if selected_paths:
            self.start_conversion([Path(path) for path in selected_paths])

    def handle_drop(self, event) -> None:
        if self.processing:
            return

        self.start_conversion(parse_drop_paths(self.root, getattr(event, "data", "")))

    def start_conversion(self, paths: list[Path]) -> None:
        if self.processing or not paths:
            return

        self.processing = True
        self.render_stage("status_loading", UI_TEXT["loading_title"], UI_TEXT["loading_subtitle"], "")

        worker = threading.Thread(
            target=self.run_conversion_worker,
            args=(paths,),
            daemon=True,
        )
        worker.start()

    def run_conversion_worker(self, paths: list[Path]) -> None:
        self.publish("stage", {"status_key": "status_loading"})
        sources, skipped = collect_source_files(paths)
        result = ConversionResult(success=0, skipped=skipped, total=len(sources))

        for index, source_path in enumerate(sources, start=1):
            self.publish(
                "stage",
                {
                    "status_key": "status_processing",
                    "title": UI_TEXT["processing_title"],
                    "subtitle": source_path.name,
                    "current": index,
                    "total": result.total,
                },
            )

            try:
                output_path = self.converter.convert(
                    source_path,
                    before_save=lambda save_path, index=index, total=result.total: self.publish(
                        "stage",
                        {
                            "status_key": "status_saving",
                            "title": UI_TEXT["saving_title"],
                            "subtitle": save_path.name,
                            "current": index,
                            "total": total,
                        },
                    ),
                )
                result.success += 1
            except (OSError, ValueError, UnidentifiedImageError):
                result.skipped += 1
            except Exception:
                result.skipped += 1

        self.publish(
            "complete",
            {
                "success": result.success,
                "skipped": result.skipped,
                "total": result.total,
            },
        )

    def publish(self, kind: str, payload: dict[str, object] | None = None) -> None:
        self.event_queue.put(WorkerEvent(kind=kind, payload=payload or {}))

    def poll_worker_events(self) -> None:
        try:
            while True:
                event = self.event_queue.get_nowait()
                self.handle_worker_event(event)
        except queue.Empty:
            pass
        finally:
            if self.root.winfo_exists():
                self.root.after(QUEUE_POLL_INTERVAL_MS, self.poll_worker_events)

    def handle_worker_event(self, event: WorkerEvent) -> None:
        if event.kind == "stage":
            status_key = str(event.payload.get("status_key", "status_processing"))
            title = str(event.payload.get("title", UI_TEXT["loading_title"]))
            subtitle = str(event.payload.get("subtitle", UI_TEXT["loading_subtitle"]))
            current = event.payload.get("current")
            total = event.payload.get("total")
            progress_text = ""
            if isinstance(current, int) and isinstance(total, int) and total:
                progress_text = UI_TEXT["progress_template"].format(current=current, total=total)
            self.render_stage(status_key, title, subtitle, progress_text)
            return

        if event.kind == "complete":
            success = int(event.payload.get("success", 0))
            skipped = int(event.payload.get("skipped", 0))
            self.processing = False
            self.render_complete(success, skipped)

    def render_idle(self) -> None:
        self.stop_status_animation()
        subtitle_key = "empty_subtitle" if DND_READY else "empty_subtitle_click_only"
        self.drop_title_var.set(UI_TEXT["empty_title_drop"])
        self.drop_subtitle_var.set(UI_TEXT[subtitle_key])
        self.status_var.set(UI_TEXT["status_idle"])
        self.progress_var.set("")
        self.configure_drop_colors(COLORS["card_bg"], COLORS["text"])

    def render_stage(self, status_key: str, title: str, subtitle: str, progress_text: str) -> None:
        self.drop_title_var.set(title)
        self.drop_subtitle_var.set(subtitle)
        self.progress_var.set(progress_text)
        self.configure_drop_colors(COLORS["card_bg"], COLORS["text"])
        self.start_status_animation(status_key)

    def render_complete(self, success: int, skipped: int) -> None:
        self.stop_status_animation()
        summary = UI_TEXT["complete_summary"].format(success=success, skipped=skipped)
        self.drop_title_var.set(UI_TEXT["complete_title"])
        self.drop_subtitle_var.set(summary)
        self.status_var.set(UI_TEXT["status_complete"])
        self.progress_var.set(summary)
        self.configure_drop_colors(COLORS["success_bg"], COLORS["success"])

    def configure_drop_colors(self, background: str, foreground: str) -> None:
        if self.drop_area is not None:
            self.drop_area.configure(bg=background)
        if self.drop_inner is not None:
            self.drop_inner.configure(bg=background)
        if self.drop_title_label is not None:
            self.drop_title_label.configure(bg=background, fg=foreground)
        if self.drop_subtitle_label is not None:
            self.drop_subtitle_label.configure(bg=background)

    def start_status_animation(self, status_key: str) -> None:
        if status_key not in ANIMATED_STATUS_KEYS:
            self.stop_status_animation()
            self.status_var.set(UI_TEXT.get(status_key, ""))
            return

        if self.status_key != status_key:
            self.status_key = status_key
            self.status_started_at = time.monotonic()
            self.status_tick = 0

        if self.status_after_id is None:
            self.animate_status()

    def animate_status(self) -> None:
        if self.status_key not in ANIMATED_STATUS_KEYS:
            self.status_after_id = None
            return

        self.status_tick += 1
        elapsed = time.monotonic() - self.status_started_at
        self.status_var.set(self.build_status_text(self.status_key, self.status_tick, elapsed))
        self.status_after_id = self.root.after(STATUS_ANIMATION_INTERVAL_MS, self.animate_status)

    def build_status_text(self, status_key: str, tick: int, elapsed: float) -> str:
        if elapsed >= LONG_STATUS_DELAY_SECONDS:
            phrase_phase = tick % 12
            if phrase_phase == 6:
                return UI_TEXT["status_phrase_1"]
            if phrase_phase == 7:
                return UI_TEXT["status_phrase_2"]
            if phrase_phase == 8:
                return UI_TEXT["status_phrase_3"]

        dot_count = ((tick - 1) % 3) + 1
        return f"{UI_TEXT[status_key]}{'.' * dot_count}"

    def stop_status_animation(self) -> None:
        if self.status_after_id is not None:
            self.root.after_cancel(self.status_after_id)
            self.status_after_id = None
        self.status_key = None

    def close(self) -> None:
        self.stop_status_animation()
        self.root.destroy()

    def run(self) -> None:
        self.root.mainloop()


def choose_font_family(root: tk.Tk) -> str:
    preferred = ["BIZ UDPGothic", "Yu Gothic UI", "Meiryo"]
    available = set(tkfont.families(root))
    for family in preferred:
        if family in available:
            return family
    return "TkDefaultFont"


def get_common_icon_path() -> Path:
    base_candidates = [
        BASE_DIR,
        Path(sys.executable).resolve().parent,
        Path.cwd(),
    ]
    relative_candidates = [
        Path("..") / ".." / "02_assets" / "dake_icon.ico",
        Path("..") / ".." / ".." / "02_assets" / "dake_icon.ico",
    ]

    for base_path in base_candidates:
        for relative_path in relative_candidates:
            icon_path = (base_path / relative_path).resolve()
            if icon_path.exists():
                return icon_path

    return (BASE_DIR / ".." / ".." / "02_assets" / "dake_icon.ico").resolve()


def normalize_to_rgb(image: Image.Image) -> Image.Image:
    if image.mode == "RGB":
        return image

    if "A" in image.getbands():
        rgba_image = image.convert("RGBA")
        background = Image.new("RGB", rgba_image.size, (255, 255, 255))
        background.paste(rgba_image, mask=rgba_image.getchannel("A"))
        rgba_image.close()
        return background

    return image.convert("RGB")


def make_available_jpg_path(source_path: Path) -> Path:
    candidate = source_path.with_suffix(".jpg")
    if not candidate.exists():
        return candidate

    sequence = 2
    while True:
        numbered_candidate = source_path.with_name(f"{source_path.stem}_{sequence}.jpg")
        if not numbered_candidate.exists():
            return numbered_candidate
        sequence += 1


def collect_source_files(paths: Iterable[Path]) -> tuple[list[Path], int]:
    sources: list[Path] = []
    seen: set[str] = set()
    skipped = 0

    for raw_path in paths:
        try:
            path = raw_path.resolve()
        except Exception:
            skipped += 1
            continue

        if path.is_file():
            if add_if_supported(path, sources, seen):
                continue
            skipped += 1
            continue

        if path.is_dir():
            for child in iter_directory_files(path):
                if add_if_supported(child, sources, seen):
                    continue
                skipped += 1
            continue

        skipped += 1

    return sources, skipped


def iter_directory_files(folder_path: Path) -> Iterable[Path]:
    for root_path, _directory_names, file_names in os.walk(folder_path):
        for file_name in sorted(file_names, key=str.lower):
            yield Path(root_path) / file_name


def add_if_supported(path: Path, sources: list[Path], seen: set[str]) -> bool:
    if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
        return False

    key = str(path).lower()
    if key in seen:
        return False

    seen.add(key)
    sources.append(path)
    return True


def parse_drop_paths(root: tk.Tk, raw_data: str) -> list[Path]:
    if not raw_data:
        return []

    try:
        items = root.tk.splitlist(raw_data)
    except tk.TclError:
        items = [raw_data]

    paths: list[Path] = []
    for item in items:
        cleaned = str(item).strip()
        if cleaned.startswith("{") and cleaned.endswith("}"):
            cleaned = cleaned[1:-1]
        if cleaned:
            paths.append(Path(cleaned))
    return paths


def main() -> None:
    DakeHeicJpgApp().run()


if __name__ == "__main__":
    main()
