# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import queue
import re
import subprocess
import sys
import threading
import webbrowser
from dataclasses import dataclass
from pathlib import Path

import tkinter as tk
from tkinter import filedialog

from PIL import Image, ImageOps, UnidentifiedImageError

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    BASE_WINDOW = TkinterDnD.Tk
    DND_READY = True
except Exception:
    DND_FILES = ""
    BASE_WINDOW = tk.Tk
    DND_READY = False


APP_NAME = "画像→PDF"
WINDOW_TITLE = "画像→PDF"
EXE_NAME = "DakeImageToPDF.exe"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"
FONT_NAME = "Yu Gothic UI"
WINDOW_SIZE = "960x560"
WINDOW_MIN_WIDTH = 960
WINDOW_MIN_HEIGHT = 520
PDF_DPI = 150
A4_PORTRAIT_PX = (1240, 1754)
A4_LANDSCAPE_PX = (1754, 1240)
SAFE_MARGIN_MM = 8
SUPPORTED_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".webp"}
RESET_DELAY_MS = 4500

FOOTER_URLS = {
    "assessment": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "instagram": "https://instagram.com/kikuta.shimarisu_fudosan",
}

THEME = {
    "bg": "#F4F5F7",
    "panel": "#FFFFFF",
    "panel_soft": "#FAFBFC",
    "text": "#1F2937",
    "muted": "#6B7280",
    "border": "#D7DCE3",
    "accent": "#2563EB",
    "accent_soft": "#EAF2FF",
    "success": "#1F9D63",
    "error": "#C2413B",
    "drop_idle": "#FFFFFF",
    "drop_hover": "#F8FBFF",
    "drop_busy": "#F7F8FA",
}

UI_TEXT = {
    "brand_series": "シンプルそれDAKEシリーズ",
    "main_title": "画像をPDFにする",
    "main_description": "追加した画像を、そのままPDFに変換して保存します。",
    "drop_badge": "1枚だけ、即PDF",
    "drop_empty_title": "画像を追加してください",
    "drop_empty_subtitle": "ドラッグ＆ドロップ または クリックして追加",
    "drop_empty_subtitle_click_only": "クリックして追加",
    "drop_supported_hint": "対応: PNG / JPG / JPEG / BMP / WEBP",
    "drop_processing_hint": "終わると保存フォルダを開きます",
    "drop_reset_hint": "数秒後に初期画面へ戻ります",
    "status_idle": "未選択",
    "status_loading": "読み込み中",
    "status_converting": "PDF作成中",
    "status_saving": "保存中",
    "status_complete": "完了",
    "status_error": "エラー",
    "status_idle_detail": "画像1枚をA4のPDFに変換します",
    "status_loading_detail": "画像を読み込んでいます",
    "status_converting_detail": "A4に合わせてPDFを作成しています",
    "status_saving_detail": "PDFを保存しています",
    "status_complete_detail": "保存先を開きました",
    "status_error_detail": "クリックしてやり直せます",
    "loading_title": "読み込み中",
    "converting_title": "PDF作成中",
    "saving_title": "保存中",
    "complete_title": "PDFを保存しました",
    "complete_hint": "{folder}",
    "error_retry_hint": "クリックしてもう一度追加できます",
    "error_multiple_files": "画像は1枚だけ追加してください",
    "error_unsupported": "この形式の画像には対応していません",
    "error_open_failed": "画像を開けませんでした",
    "error_save_failed": "PDFの保存に失敗しました",
    "error_processing_failed": "PDFの作成に失敗しました",
    "error_busy": "処理が終わるまで少しだけお待ちください",
    "dialog_select_image_title": "画像を選択してください",
    "file_type_images": "画像ファイル",
    "file_type_all": "すべてのファイル",
    "save_folder_prefix": "保存先",
    "output_file_prefix": "出力ファイル",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_left_suffix": " / 止まらない、迷わない、すぐ終わる。",
    "footer_assessment": "戸建買取査定",
    "footer_separator": " ｜ ",
    "footer_instagram": "Instagram",
    "footer_copyright": COPYRIGHT,
}


@dataclass(frozen=True)
class WorkerEvent:
    kind: str
    payload: object | None = None


class AppError(Exception):
    def __init__(self, code: str, original: Exception | None = None):
        super().__init__(code)
        self.code = code
        self.original = original


class WorkerNotifier:
    def __init__(self) -> None:
        self.event_queue: queue.Queue[WorkerEvent] = queue.Queue()

    def publish(self, kind: str, payload: object | None = None) -> None:
        self.event_queue.put(WorkerEvent(kind=kind, payload=payload))


class ImagePdfService:
    def convert_to_pdf(self, source_path: Path, save_dir: Path, on_stage) -> Path:
        on_stage("loading", {"source_name": source_path.name})
        image = self._load_image(source_path)

        try:
            page_size = self._resolve_page_size(image.size)
            on_stage("converting", {"source_name": source_path.name})
            pdf_page = self._build_pdf_page(image, page_size)
        except AppError:
            image.close()
            raise
        except Exception as exc:
            image.close()
            raise AppError("error_processing_failed", exc) from exc

        output_path = make_available_path(
            save_dir / f"{make_safe_stem(source_path.stem)}_dake.pdf"
        )
        on_stage("saving", {"output_name": output_path.name})

        try:
            save_dir.mkdir(parents=True, exist_ok=True)
            pdf_page.save(output_path, "PDF", resolution=PDF_DPI)
            return output_path
        except Exception as exc:
            raise AppError("error_save_failed", exc) from exc
        finally:
            image.close()
            pdf_page.close()

    def _load_image(self, source_path: Path) -> Image.Image:
        try:
            with Image.open(source_path) as raw_image:
                prepared = ImageOps.exif_transpose(raw_image)
                if self._has_alpha(prepared):
                    alpha_image = prepared.convert("RGBA")
                    background = Image.new("RGBA", alpha_image.size, (255, 255, 255, 255))
                    merged = Image.alpha_composite(background, alpha_image)
                    alpha_image.close()
                    background.close()
                    return merged.convert("RGB")
                return prepared.convert("RGB")
        except UnidentifiedImageError as exc:
            raise AppError("error_unsupported", exc) from exc
        except OSError as exc:
            raise AppError("error_open_failed", exc) from exc

    def _resolve_page_size(self, image_size: tuple[int, int]) -> tuple[int, int]:
        width, height = image_size
        if width > height:
            return A4_LANDSCAPE_PX
        return A4_PORTRAIT_PX

    def _build_pdf_page(
        self,
        image: Image.Image,
        page_size: tuple[int, int],
    ) -> Image.Image:
        page_width, page_height = page_size
        margin_px = mm_to_px(SAFE_MARGIN_MM, PDF_DPI)
        content_width = max(1, page_width - margin_px * 2)
        content_height = max(1, page_height - margin_px * 2)
        image_width, image_height = image.size
        scale = min(content_width / image_width, content_height / image_height, 1.0)

        target_width = max(1, int(round(image_width * scale)))
        target_height = max(1, int(round(image_height * scale)))

        if scale < 1.0:
            fitted = image.resize(
                (target_width, target_height),
                Image.Resampling.LANCZOS,
            )
        else:
            fitted = image.copy()

        page = Image.new("RGB", (page_width, page_height), "white")
        offset_x = margin_px + (content_width - target_width) // 2
        offset_y = margin_px + (content_height - target_height) // 2
        page.paste(fitted, (offset_x, offset_y))
        fitted.close()
        return page

    def _has_alpha(self, image: Image.Image) -> bool:
        if image.mode in {"RGBA", "LA"}:
            return True
        return image.mode == "P" and "transparency" in image.info


class DakeImageToPdfApp:
    def __init__(self) -> None:
        self.root = BASE_WINDOW()
        self.notifier = WorkerNotifier()
        self.service = ImagePdfService()
        self.current_job_active = False
        self.drag_active = False
        self.reset_after_id: str | None = None
        self.icon_image: tk.PhotoImage | None = None

        self.drop_badge_var = tk.StringVar(value=UI_TEXT["drop_badge"])
        self.drop_title_var = tk.StringVar(value="")
        self.drop_subtitle_var = tk.StringVar(value="")
        self.drop_hint_var = tk.StringVar(value="")
        self.status_label_var = tk.StringVar(value="")
        self.status_detail_var = tk.StringVar(value="")

        self.drop_area: tk.Frame | None = None
        self.drop_inner: tk.Frame | None = None
        self.drop_badge_label: tk.Label | None = None
        self.drop_title_label: tk.Label | None = None
        self.drop_subtitle_label: tk.Label | None = None
        self.drop_hint_label: tk.Label | None = None
        self.status_label: tk.Label | None = None

        self._configure_window()
        self._build_ui()
        self._render_idle()
        self._poll_worker_events()

    def _configure_window(self) -> None:
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(WINDOW_MIN_WIDTH, WINDOW_MIN_HEIGHT)
        self.root.resizable(False, False)
        self.root.configure(bg=THEME["bg"])
        self.root.protocol("WM_DELETE_WINDOW", self._close)
        self._apply_window_icon()

    def _apply_window_icon(self) -> None:
        try:
            if getattr(sys, "frozen", False):
                return

            icon_path = get_common_icon_path().resolve()
            if icon_path.exists():
                self.root.iconbitmap(default=str(icon_path))
        except Exception:
            pass

    def _build_ui(self) -> None:
        container = tk.Frame(self.root, bg=THEME["bg"], padx=28, pady=24)
        container.pack(fill="both", expand=True)

        header = tk.Frame(container, bg=THEME["bg"])
        header.pack(fill="x", pady=(0, 16))

        header_title = tk.Label(
            header,
            text=UI_TEXT["main_title"],
            bg=THEME["bg"],
            fg=THEME["text"],
            font=(FONT_NAME, 18, "bold"),
        )
        header_title.pack(side="left")

        header_description = tk.Label(
            header,
            text=UI_TEXT["main_description"],
            bg=THEME["bg"],
            fg=THEME["muted"],
            font=(FONT_NAME, 10),
            padx=12,
        )
        header_description.pack(side="left")

        self.drop_area = tk.Frame(
            container,
            bg=THEME["drop_idle"],
            highlightthickness=1,
            highlightbackground=THEME["border"],
            highlightcolor=THEME["accent"],
            bd=0,
            cursor="hand2",
        )
        self.drop_area.pack(fill="both", expand=True)

        self.drop_inner = tk.Frame(self.drop_area, bg=THEME["drop_idle"], padx=32, pady=32)
        self.drop_inner.place(relx=0.5, rely=0.5, anchor="center")

        self.drop_badge_label = tk.Label(
            self.drop_inner,
            textvariable=self.drop_badge_var,
            bg=THEME["accent_soft"],
            fg=THEME["accent"],
            font=(FONT_NAME, 10, "bold"),
            padx=12,
            pady=6,
        )
        self.drop_badge_label.pack(pady=(0, 18))

        self.drop_title_label = tk.Label(
            self.drop_inner,
            textvariable=self.drop_title_var,
            bg=THEME["drop_idle"],
            fg=THEME["text"],
            font=(FONT_NAME, 24, "bold"),
        )
        self.drop_title_label.pack()

        self.drop_subtitle_label = tk.Label(
            self.drop_inner,
            textvariable=self.drop_subtitle_var,
            bg=THEME["drop_idle"],
            fg=THEME["muted"],
            font=(FONT_NAME, 12),
            pady=12,
        )
        self.drop_subtitle_label.pack()

        self.drop_hint_label = tk.Label(
            self.drop_inner,
            textvariable=self.drop_hint_var,
            bg=THEME["drop_idle"],
            fg=THEME["muted"],
            font=(FONT_NAME, 10),
        )
        self.drop_hint_label.pack()

        status_panel = tk.Frame(
            container,
            bg=THEME["panel"],
            highlightthickness=1,
            highlightbackground=THEME["border"],
            padx=18,
            pady=14,
        )
        status_panel.pack(fill="x", pady=(16, 0))

        self.status_label = tk.Label(
            status_panel,
            textvariable=self.status_label_var,
            bg=THEME["panel"],
            fg=THEME["text"],
            font=(FONT_NAME, 11, "bold"),
        )
        self.status_label.pack(anchor="w")

        status_detail = tk.Label(
            status_panel,
            textvariable=self.status_detail_var,
            bg=THEME["panel"],
            fg=THEME["muted"],
            font=(FONT_NAME, 10),
            pady=4,
        )
        status_detail.pack(anchor="w")

        footer = tk.Frame(container, bg=THEME["bg"])
        footer.pack(fill="x", pady=(12, 0))

        footer_left = tk.Frame(footer, bg=THEME["bg"])
        footer_left.pack(side="left")

        footer_left_label = tk.Label(
            footer_left,
            text=UI_TEXT["footer_left"],
            bg=THEME["bg"],
            fg=THEME["muted"],
            font=(FONT_NAME, 9),
        )
        footer_left_label.pack(side="left")

        footer_left_suffix = tk.Label(
            footer_left,
            text=UI_TEXT["footer_left_suffix"],
            bg=THEME["bg"],
            fg=THEME["muted"],
            font=(FONT_NAME, 9),
        )
        footer_left_suffix.pack(side="left")

        footer_right = tk.Frame(footer, bg=THEME["bg"])
        footer_right.pack(side="right")

        self._make_footer_label(footer_right, UI_TEXT["footer_assessment"], FOOTER_URLS["assessment"], True)
        self._make_footer_label(footer_right, UI_TEXT["footer_separator"])
        self._make_footer_label(footer_right, UI_TEXT["footer_instagram"], FOOTER_URLS["instagram"], True)
        self._make_footer_label(footer_right, UI_TEXT["footer_separator"])
        self._make_footer_label(footer_right, UI_TEXT["footer_copyright"])

        click_targets = [
            self.drop_area,
            self.drop_inner,
            self.drop_badge_label,
            self.drop_title_label,
            self.drop_subtitle_label,
            self.drop_hint_label,
        ]

        for widget in click_targets:
            widget.bind("<Button-1>", self._on_click_add)

        self._register_drop_targets(click_targets)

    def _make_footer_label(
        self,
        parent: tk.Widget,
        text: str,
        url: str | None = None,
        clickable: bool = False,
    ) -> None:
        label = tk.Label(
            parent,
            text=text,
            bg=THEME["bg"],
            fg=THEME["muted"] if not clickable else THEME["text"],
            font=(FONT_NAME, 9),
            cursor="hand2" if clickable else "arrow",
        )
        label.pack(side="left")
        if clickable and url:
            label.bind("<Button-1>", lambda _event, link=url: webbrowser.open_new(link))

    def _register_drop_targets(self, widgets: list[tk.Widget]) -> None:
        if not DND_READY or not DND_FILES:
            return

        for widget in widgets:
            try:
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", self._on_drop)
                widget.dnd_bind("<<DragEnter>>", self._on_drag_enter)
                widget.dnd_bind("<<DragLeave>>", self._on_drag_leave)
            except Exception:
                continue

    def _on_click_add(self, _event=None) -> None:
        if self.current_job_active:
            self._show_busy_message()
            return

        selected = filedialog.askopenfilename(
            title=UI_TEXT["dialog_select_image_title"],
            filetypes=[
                (
                    UI_TEXT["file_type_images"],
                    "*.png *.jpg *.jpeg *.bmp *.webp",
                ),
                (UI_TEXT["file_type_all"], "*.*"),
            ],
        )
        if selected:
            self._accept_paths([Path(selected)])

    def _on_drop(self, event) -> None:
        self.drag_active = False
        self._apply_drop_style()
        dropped_paths = self._parse_drop_paths(getattr(event, "data", ""))
        self._accept_paths(dropped_paths)

    def _on_drag_enter(self, _event) -> str:
        if not self.current_job_active:
            self.drag_active = True
            self._apply_drop_style()
        return "copy"

    def _on_drag_leave(self, _event) -> str:
        self.drag_active = False
        self._apply_drop_style()
        return "copy"

    def _accept_paths(self, paths: list[Path]) -> None:
        if self.current_job_active:
            self._show_busy_message()
            return

        self._cancel_reset()

        if len(paths) != 1:
            self._render_error(UI_TEXT["error_multiple_files"])
            return

        source_path = paths[0]
        if source_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            self._render_error(UI_TEXT["error_unsupported"])
            return

        if not source_path.exists() or not source_path.is_file():
            self._render_error(UI_TEXT["error_open_failed"])
            return

        self.current_job_active = True
        worker = threading.Thread(
            target=self._run_conversion,
            args=(source_path,),
            daemon=True,
        )
        worker.start()

    def _run_conversion(self, source_path: Path) -> None:
        try:
            output_path = self.service.convert_to_pdf(
                source_path=source_path,
                save_dir=default_downloads_path(),
                on_stage=self._publish_stage,
            )
            self.notifier.publish("complete", {"output_path": output_path})
        except AppError as exc:
            self.notifier.publish("error", {"code": exc.code})
        except Exception:
            self.notifier.publish("error", {"code": "error_processing_failed"})

    def _publish_stage(self, stage: str, payload: dict[str, object]) -> None:
        event_payload = {"stage": stage}
        event_payload.update(payload)
        self.notifier.publish("stage", event_payload)

    def _poll_worker_events(self) -> None:
        while True:
            try:
                event = self.notifier.event_queue.get_nowait()
            except queue.Empty:
                break
            self._handle_worker_event(event)

        self.root.after(80, self._poll_worker_events)

    def _handle_worker_event(self, event: WorkerEvent) -> None:
        payload = event.payload if isinstance(event.payload, dict) else {}

        if event.kind == "stage":
            self._render_stage(
                stage=str(payload.get("stage", "")),
                name=str(payload.get("source_name") or payload.get("output_name") or ""),
            )
            return

        if event.kind == "complete":
            output_path = payload.get("output_path")
            if isinstance(output_path, Path):
                self.current_job_active = False
                self._render_complete(output_path)
                self._open_folder(output_path.parent)
                self._schedule_reset()
            return

        if event.kind == "error":
            self.current_job_active = False
            message_key = str(payload.get("code", "error_processing_failed"))
            self._render_error(UI_TEXT.get(message_key, UI_TEXT["error_processing_failed"]))

    def _render_idle(self) -> None:
        self.current_job_active = False
        subtitle_key = (
            "drop_empty_subtitle" if DND_READY else "drop_empty_subtitle_click_only"
        )
        self._update_drop_content(
            title=UI_TEXT["drop_empty_title"],
            subtitle=UI_TEXT[subtitle_key],
            hint=UI_TEXT["drop_supported_hint"],
            title_color=THEME["text"],
            badge_color=THEME["accent"],
            badge_bg=THEME["accent_soft"],
        )
        self._update_status(
            label=UI_TEXT["status_idle"],
            detail=UI_TEXT["status_idle_detail"],
            color=THEME["text"],
        )
        self._apply_drop_style()

    def _render_stage(self, stage: str, name: str) -> None:
        if stage == "loading":
            self._update_drop_content(
                title=UI_TEXT["loading_title"],
                subtitle=name,
                hint=UI_TEXT["drop_processing_hint"],
                title_color=THEME["text"],
                badge_color=THEME["accent"],
                badge_bg=THEME["accent_soft"],
            )
            self._update_status(
                label=UI_TEXT["status_loading"],
                detail=UI_TEXT["status_loading_detail"],
                color=THEME["text"],
            )
        elif stage == "converting":
            self._update_drop_content(
                title=UI_TEXT["converting_title"],
                subtitle=name,
                hint=UI_TEXT["drop_processing_hint"],
                title_color=THEME["text"],
                badge_color=THEME["accent"],
                badge_bg=THEME["accent_soft"],
            )
            self._update_status(
                label=UI_TEXT["status_converting"],
                detail=UI_TEXT["status_converting_detail"],
                color=THEME["text"],
            )
        elif stage == "saving":
            self._update_drop_content(
                title=UI_TEXT["saving_title"],
                subtitle=name,
                hint=UI_TEXT["drop_processing_hint"],
                title_color=THEME["text"],
                badge_color=THEME["accent"],
                badge_bg=THEME["accent_soft"],
            )
            self._update_status(
                label=UI_TEXT["status_saving"],
                detail=UI_TEXT["status_saving_detail"],
                color=THEME["text"],
            )

        self._apply_drop_style()

    def _render_complete(self, output_path: Path) -> None:
        self._update_drop_content(
            title=UI_TEXT["complete_title"],
            subtitle=output_path.name,
            hint=UI_TEXT["complete_hint"].format(folder=output_path.parent),
            title_color=THEME["success"],
            badge_color=THEME["success"],
            badge_bg="#EAF8F1",
        )
        self._update_status(
            label=UI_TEXT["status_complete"],
            detail=f'{UI_TEXT["save_folder_prefix"]}: {output_path.parent}',
            color=THEME["success"],
        )
        self._apply_drop_style()

    def _render_error(self, message: str, auto_reset: bool = True) -> None:
        self.current_job_active = False
        self._update_drop_content(
            title=message,
            subtitle=UI_TEXT["error_retry_hint"],
            hint=UI_TEXT["drop_supported_hint"],
            title_color=THEME["error"],
            badge_color=THEME["error"],
            badge_bg="#FDECEC",
        )
        self._update_status(
            label=UI_TEXT["status_error"],
            detail=message,
            color=THEME["error"],
        )
        self._apply_drop_style()

        if auto_reset:
            self._schedule_reset()

    def _show_busy_message(self) -> None:
        self._update_status(
            label=self.status_label_var.get() or UI_TEXT["status_loading"],
            detail=UI_TEXT["error_busy"],
            color=THEME["text"],
        )

    def _update_drop_content(
        self,
        title: str,
        subtitle: str,
        hint: str,
        title_color: str,
        badge_color: str,
        badge_bg: str,
    ) -> None:
        self.drop_title_var.set(title)
        self.drop_subtitle_var.set(subtitle)
        self.drop_hint_var.set(hint)

        if self.drop_title_label is not None:
            self.drop_title_label.configure(fg=title_color)
        if self.drop_badge_label is not None:
            self.drop_badge_label.configure(bg=badge_bg, fg=badge_color)

    def _update_status(self, label: str, detail: str, color: str) -> None:
        self.status_label_var.set(label)
        self.status_detail_var.set(detail)
        if self.status_label is not None:
            self.status_label.configure(fg=color)

    def _apply_drop_style(self) -> None:
        if self.drop_area is None or self.drop_inner is None:
            return

        if self.current_job_active:
            bg_color = THEME["drop_busy"]
            border_color = THEME["border"]
        elif self.drag_active:
            bg_color = THEME["drop_hover"]
            border_color = THEME["accent"]
        else:
            bg_color = THEME["drop_idle"]
            border_color = THEME["border"]

        widgets = [
            self.drop_area,
            self.drop_inner,
            self.drop_title_label,
            self.drop_subtitle_label,
            self.drop_hint_label,
        ]

        self.drop_area.configure(bg=bg_color, highlightbackground=border_color)
        self.drop_inner.configure(bg=bg_color)
        for widget in widgets[2:]:
            if widget is not None:
                widget.configure(bg=bg_color)

    def _open_folder(self, folder_path: Path) -> None:
        try:
            if os.name == "nt":
                os.startfile(str(folder_path))
                return
            if sys.platform == "darwin":
                subprocess.run(["open", str(folder_path)], check=False)
                return
            subprocess.run(["xdg-open", str(folder_path)], check=False)
        except Exception:
            pass

    def _parse_drop_paths(self, raw_data: str) -> list[Path]:
        if not raw_data:
            return []

        try:
            items = list(self.root.tk.splitlist(raw_data))
        except tk.TclError:
            items = [raw_data]

        paths: list[Path] = []
        for item in items:
            cleaned = item.strip()
            if cleaned.startswith("{") and cleaned.endswith("}"):
                cleaned = cleaned[1:-1]
            if cleaned:
                paths.append(Path(cleaned))
        return paths

    def _schedule_reset(self) -> None:
        self._cancel_reset()
        self.reset_after_id = self.root.after(RESET_DELAY_MS, self._reset_to_idle)

    def _cancel_reset(self) -> None:
        if self.reset_after_id is not None:
            self.root.after_cancel(self.reset_after_id)
            self.reset_after_id = None

    def _reset_to_idle(self) -> None:
        self.reset_after_id = None
        self._render_idle()

    def _close(self) -> None:
        self._cancel_reset()
        self.root.destroy()

    def run(self) -> None:
        self.root.mainloop()


def default_downloads_path() -> Path:
    downloads = Path.home() / "Downloads"
    if downloads.exists():
        return downloads
    return Path.home()


def mm_to_px(mm: float, dpi: int) -> int:
    return max(1, int(round((mm / 25.4) * dpi)))


def make_available_path(target_path: Path) -> Path:
    if not target_path.exists():
        return target_path

    counter = 1
    while True:
        candidate = target_path.with_name(f"{target_path.stem}_{counter}{target_path.suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def make_safe_stem(stem: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*]+', "_", stem).strip(" ._")
    return cleaned or "image"


def resource_path(name: str) -> str:
    base_path = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return str(base_path / name)


def get_common_icon_path() -> Path:
    return Path(__file__).resolve().parent / ".." / ".." / "02_assets" / "dake_icon.ico"


def main() -> None:
    app = DakeImageToPdfApp()
    app.run()


if __name__ == "__main__":
    main()
