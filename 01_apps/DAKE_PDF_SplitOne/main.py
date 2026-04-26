# -*- coding: utf-8 -*-
import json
import os
import queue
import re
import sys
import threading
import tkinter as tk
import webbrowser
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox

from pypdf import PdfReader, PdfWriter

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    HAS_DND = True
except Exception:
    DND_FILES = None
    TkinterDnD = None
    HAS_DND = False


APP_NAME = "PDF分割One"
WINDOW_TITLE = "PDF分割One"
EXE_NAME = "DakePDF_Split_One.exe"
CONFIG_NAME = "dake_pdf_split_one_config.json"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"
FONT_NAME = "Yu Gothic UI"
WINDOW_WIDTH = 860
WINDOW_HEIGHT = 520
SAVE_FOLDER_TEXT_MAX = 72
SOURCE_FILE_TEXT_MAX = 36
OUTPUT_FOLDER_TEXT_MAX = 72
RESULT_PATH_TEXT_MAX = 104
PANEL_PATH_WRAP_LENGTH = 560
TOOLTIP_WRAP_LENGTH = 520
FOOTER_LINK_1_URL = (
    "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74"
)
FOOTER_LINK_2_URL = "https://instagram.com/kikuta.shimarisu_fudosan"

THEME = {
    "bg": "#F5F6F8",
    "panel": "#FFFFFF",
    "text": "#1F2937",
    "muted": "#6B7280",
    "border": "#E5E7EB",
    "accent": "#2B6CB0",
    "accent_soft": "#EDF4FF",
    "success": "#1F9D63",
    "error": "#C2413B",
    "button_bg": "#FFFFFF",
    "button_hover": "#F3F4F6",
    "button_disabled": "#ECEFF3",
}

UI_TEXT = {
    "brand_series": "シンプルそれDAKEシリーズ",
    "main_title": "PDFを1ページずつ分割する",
    "main_description": "追加したPDFを全ページ1枚ずつに分けて保存します。",
    "button_select_folder": "保存先を選ぶ",
    "button_refresh": "リフレッシュ",
    "empty_title": "PDFを追加してください",
    "empty_subtitle": "ドラッグ＆ドロップ または クリックして追加",
    "status_idle": "未選択",
    "status_loading": "読み込み中",
    "status_processing": "分割中",
    "status_complete": "完了",
    "status_error": "エラー",
    "progress_dots_1": "処理中.",
    "progress_dots_2": "処理中..",
    "progress_dots_3": "処理中...",
    "dialog_complete_title": "完了",
    "dialog_complete_message": "分割が完了しました。",
    "dialog_error_title": "エラー",
    "dialog_select_pdf_title": "PDFを選択してください",
    "dialog_select_folder_title": "保存先を選んでください",
    "pdf_file_type": "PDFファイル",
    "error_not_pdf": "PDFファイルのみ追加できます。",
    "error_multiple_files": "PDFは1ファイルだけ追加してください。",
    "error_open_failed": "PDFを開けませんでした。",
    "error_split_failed": "分割中にエラーが発生しました。",
    "idle_detail": "PDFを1ファイル追加すると自動で分割を始めます。",
    "loading_detail": "PDFの内容を確認しています。",
    "processing_detail": "{current}/{total} を分割しています: {name}",
    "processing_prepare_detail": "分割を始めています。",
    "complete_detail": "保存フォルダを開きます。",
    "error_detail": "もう一度PDFを追加してください。",
    "save_folder_prefix": "保存先",
    "output_folder_prefix": "出力フォルダ",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_left_suffix": " / 止まらない、迷わない、すぐ終わる。",
    "footer_link_1": "戸建買取査定",
    "footer_link_2": "Instagram",
    "footer_separator": " ｜ ",
    "footer_copyright": "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta",
}


@dataclass(frozen=True)
class WorkerEvent:
    kind: str
    payload: object | None = None


class PdfServiceError(Exception):
    def __init__(self, code: str, original: Exception | None = None):
        super().__init__(code)
        self.code = code
        self.original = original


class AppConfig:
    def __init__(self, path: Path):
        self.path = path
        self.data = self._load()

    def _load(self) -> dict:
        try:
            return json.loads(self.path.read_text(encoding="utf-8"))
        except Exception:
            return {}

    def save(self) -> None:
        try:
            self.path.write_text(
                json.dumps(self.data, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception:
            pass

    @property
    def last_save_folder(self) -> str:
        saved = self.data.get("last_save_folder", "")
        if saved and Path(saved).exists():
            return saved
        return default_downloads_path()

    @last_save_folder.setter
    def last_save_folder(self, value: str) -> None:
        self.data["last_save_folder"] = value
        self.save()


class WorkerNotifier:
    def __init__(self):
        self.event_queue: queue.Queue[WorkerEvent] = queue.Queue()

    def publish(self, kind: str, payload: object | None = None) -> None:
        self.event_queue.put(WorkerEvent(kind=kind, payload=payload))


class PdfSplitService:
    OUTPUT_PATTERN = re.compile(r"^p\d{3,}\.pdf$", re.IGNORECASE)

    def split_all_pages(
        self,
        source_pdf: Path,
        save_root: Path,
        on_loaded,
        on_progress,
    ) -> Path:
        try:
            reader = PdfReader(str(source_pdf))
            total_pages = len(reader.pages)
            if total_pages < 1:
                raise PdfServiceError("open_failed")
        except PdfServiceError:
            raise
        except Exception as exc:
            raise PdfServiceError("open_failed", exc) from exc

        try:
            output_folder = self._build_output_folder(source_pdf, save_root)
            output_folder.mkdir(parents=True, exist_ok=True)
            self._clear_previous_outputs(output_folder)
            on_loaded(total_pages, output_folder)

            digits = max(3, len(str(total_pages)))
            for index, page in enumerate(reader.pages, start=1):
                writer = PdfWriter()
                writer.add_page(page)
                output_path = output_folder / f"p{index:0{digits}d}.pdf"
                with output_path.open("wb") as stream:
                    writer.write(stream)
                on_progress(index, total_pages, output_path.name)

            return output_folder
        except PdfServiceError:
            raise
        except Exception as exc:
            raise PdfServiceError("split_failed", exc) from exc

    def _build_output_folder(self, source_pdf: Path, save_root: Path) -> Path:
        return save_root / f"{make_safe_stem(source_pdf.stem)}_split"

    def _clear_previous_outputs(self, output_folder: Path) -> None:
        for item in output_folder.iterdir():
            if item.is_file() and self.OUTPUT_PATTERN.fullmatch(item.name):
                item.unlink()


class StatusDots:
    def __init__(self):
        self.index = 0
        self.values = [
            UI_TEXT["progress_dots_1"],
            UI_TEXT["progress_dots_2"],
            UI_TEXT["progress_dots_3"],
        ]

    def next(self) -> str:
        value = self.values[self.index]
        self.index = (self.index + 1) % len(self.values)
        return value

    def reset(self) -> None:
        self.index = 0


class FlatButton(tk.Label):
    def __init__(self, parent, text: str, command, width: int):
        super().__init__(
            parent,
            text=text,
            width=width,
            padx=12,
            pady=8,
            bg=THEME["button_bg"],
            fg=THEME["text"],
            bd=1,
            relief="solid",
            cursor="hand2",
            font=(FONT_NAME, 10, "bold"),
            highlightthickness=0,
        )
        self.command = command
        self.disabled = False
        self.bind("<Button-1>", self._on_click)
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)

    def _on_click(self, _event) -> None:
        if not self.disabled:
            self.command()

    def _on_enter(self, _event) -> None:
        if not self.disabled:
            self.configure(bg=THEME["button_hover"])

    def _on_leave(self, _event) -> None:
        self.configure(
            bg=THEME["button_disabled"] if self.disabled else THEME["button_bg"]
        )

    def set_enabled(self, enabled: bool) -> None:
        self.disabled = not enabled
        self.configure(
            bg=THEME["button_bg"] if enabled else THEME["button_disabled"],
            fg=THEME["text"] if enabled else "#9CA3AF",
            cursor="hand2" if enabled else "arrow",
        )


class HoverTooltip:
    def __init__(self, widget):
        self.widget = widget
        self.text = ""
        self.tip_window: tk.Toplevel | None = None
        self.widget.bind("<Enter>", self._on_enter, add="+")
        self.widget.bind("<Leave>", self._on_leave, add="+")

    def update_text(self, text: str) -> None:
        self.text = text.strip()
        if not self.text:
            self.hide()

    def _on_enter(self, _event) -> None:
        if self.text:
            self.show()

    def _on_leave(self, _event) -> None:
        self.hide()

    def show(self) -> None:
        if self.tip_window or not self.text:
            return

        x = self.widget.winfo_rootx() + 12
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6

        self.tip_window = tk.Toplevel(self.widget)
        self.tip_window.wm_overrideredirect(True)
        self.tip_window.attributes("-topmost", True)
        self.tip_window.configure(bg=THEME["border"])
        self.tip_window.geometry(f"+{x}+{y}")

        label = tk.Label(
            self.tip_window,
            text=self.text,
            bg=THEME["panel"],
            fg=THEME["text"],
            font=(FONT_NAME, 9),
            justify="left",
            wraplength=TOOLTIP_WRAP_LENGTH,
            padx=8,
            pady=6,
            bd=1,
            relief="solid",
        )
        label.pack()

    def hide(self) -> None:
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None


class SplitController:
    def __init__(
        self,
        config: AppConfig,
        notifier: WorkerNotifier,
        service: PdfSplitService,
    ):
        self.config = config
        self.notifier = notifier
        self.service = service
        self.ui: SplitOneApp | None = None
        self.current_source: Path | None = None
        self.output_folder: Path | None = None
        self.save_folder = Path(self.config.last_save_folder)
        self.busy = False

    def attach_ui(self, ui) -> None:
        self.ui = ui
        self.ui.update_save_folder(self.save_folder)
        self.ui.show_idle()

    def choose_save_folder(self) -> None:
        if self.busy:
            return

        selected = filedialog.askdirectory(
            title=UI_TEXT["dialog_select_folder_title"],
            initialdir=str(self.save_folder),
            mustexist=True,
        )
        if not selected:
            return

        self.save_folder = Path(selected)
        self.config.last_save_folder = str(self.save_folder)
        if self.ui:
            self.ui.update_save_folder(self.save_folder)
            self.ui.update_status(
                state_key="idle",
                detail=UI_TEXT["idle_detail"],
            )

    def choose_pdf(self) -> None:
        if self.busy:
            return

        selected = filedialog.askopenfilename(
            title=UI_TEXT["dialog_select_pdf_title"],
            filetypes=[(UI_TEXT["pdf_file_type"], "*.pdf")],
        )
        if selected:
            self.start_from_path(Path(selected))

    def refresh(self) -> None:
        if self.busy:
            return

        self.current_source = None
        self.output_folder = None
        if self.ui:
            self.ui.show_idle()

    def handle_drop_data(self, raw_data: str) -> None:
        if self.busy or not self.ui:
            return

        paths = [Path(item.strip("{}")) for item in self.ui.root.tk.splitlist(raw_data)]
        files = [path for path in paths if path.is_file()]
        if len(files) != 1:
            self.ui.show_error(UI_TEXT["error_multiple_files"])
            return

        self.start_from_path(files[0])

    def start_from_path(self, path: Path) -> None:
        if self.busy or not self.ui:
            return

        if path.suffix.lower() != ".pdf":
            self.ui.show_error(UI_TEXT["error_not_pdf"])
            return

        self.busy = True
        self.current_source = path
        self.output_folder = self.service._build_output_folder(path, self.save_folder)
        self.ui.set_interaction_enabled(False)
        self.ui.show_file_context(path, self.output_folder)
        self.ui.update_status("loading", UI_TEXT["loading_detail"])

        worker = threading.Thread(
            target=self._run_split_job,
            args=(path, self.save_folder),
            daemon=True,
        )
        worker.start()

    def process_worker_events(self) -> None:
        while True:
            try:
                event = self.notifier.event_queue.get_nowait()
            except queue.Empty:
                return

            if not self.ui:
                continue

            if event.kind == "loaded":
                payload = event.payload or {}
                output_folder_text = str(payload.get("output_folder", "")).strip()
                if output_folder_text:
                    output_folder = Path(output_folder_text)
                    self.output_folder = output_folder
                self.ui.show_file_context(self.current_source, self.output_folder)
                self.ui.update_status("processing", UI_TEXT["processing_prepare_detail"])
            elif event.kind == "progress":
                payload = event.payload or {}
                detail = UI_TEXT["processing_detail"].format(
                    current=payload.get("current", 0),
                    total=payload.get("total", 0),
                    name=payload.get("name", ""),
                )
                self.ui.update_progress(detail)
            elif event.kind == "done":
                payload = event.payload or {}
                self.busy = False
                self.ui.set_interaction_enabled(True)
                self.ui.update_status("complete", UI_TEXT["complete_detail"])
                completed_folder_text = str(payload.get("output_folder", "")).strip()
                if completed_folder_text:
                    completed_folder = Path(completed_folder_text)
                    self.output_folder = completed_folder
                    self.ui.show_file_context(self.current_source, self.output_folder)
                self.ui.show_completion(self.output_folder)
            elif event.kind == "error":
                self.busy = False
                message = str(event.payload or UI_TEXT["error_split_failed"])
                self.ui.set_interaction_enabled(True)
                self.ui.show_error(message)

    def _run_split_job(self, source_path: Path, save_folder: Path) -> None:
        try:
            output_folder = self.service.split_all_pages(
                source_pdf=source_path,
                save_root=save_folder,
                on_loaded=self._on_loaded,
                on_progress=self._on_progress,
            )
            self.notifier.publish(
                "done",
                {"output_folder": str(output_folder)},
            )
        except PdfServiceError as exc:
            message = {
                "open_failed": UI_TEXT["error_open_failed"],
                "split_failed": UI_TEXT["error_split_failed"],
            }.get(exc.code, UI_TEXT["error_split_failed"])
            self.notifier.publish("error", message)

    def _on_loaded(self, total_pages: int, output_folder: Path) -> None:
        self.notifier.publish(
            "loaded",
            {
                "page_count": total_pages,
                "output_folder": str(output_folder),
            },
        )

    def _on_progress(self, current: int, total: int, name: str) -> None:
        self.notifier.publish(
            "progress",
            {"current": current, "total": total, "name": name},
        )


class SplitOneApp:
    def __init__(self, root: tk.Tk, controller: SplitController):
        self.root = root
        self.controller = controller
        self.status_dots = StatusDots()
        self.current_state = "idle"
        self.current_detail = UI_TEXT["idle_detail"]
        self.current_source: Path | None = None
        self.current_output: Path | None = None
        self._icon_image = None
        self.save_folder_tooltip: HoverTooltip | None = None
        self.panel_title_tooltip: HoverTooltip | None = None
        self.panel_subtitle_tooltip: HoverTooltip | None = None
        self.panel_meta_tooltip: HoverTooltip | None = None

        self.root.title(WINDOW_TITLE)
        self.root.configure(bg=THEME["bg"])
        self.root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
        self.root.resizable(False, False)
        self.root.minsize(WINDOW_WIDTH, WINDOW_HEIGHT)
        self.root.maxsize(WINDOW_WIDTH, WINDOW_HEIGHT)
        self._apply_icon()

        self._build_ui()
        self.controller.attach_ui(self)
        self._center_window()
        self._configure_drag_and_drop()
        self._poll_worker_queue()
        self._animate_busy_status()

    def _apply_icon(self) -> None:
        if getattr(sys, "frozen", False):
            return

        icon_ico = get_common_icon_path().resolve()
        if icon_ico.exists():
            try:
                self.root.iconbitmap(str(icon_ico))
            except Exception:
                pass

    def _build_ui(self) -> None:
        shell = tk.Frame(self.root, bg=THEME["bg"])
        shell.pack(fill="both", expand=True, padx=20, pady=18)

        top_bar = tk.Frame(shell, bg=THEME["bg"])
        top_bar.pack(fill="x")

        header_row = tk.Frame(top_bar, bg=THEME["bg"])
        header_row.pack(fill="x")

        self.main_title_label = tk.Label(
            header_row,
            text=UI_TEXT["main_title"],
            bg=THEME["bg"],
            fg=THEME["text"],
            font=(FONT_NAME, 12, "bold"),
            anchor="w",
        )
        self.main_title_label.pack(side="left")

        self.main_description_label = tk.Label(
            header_row,
            text=UI_TEXT["main_description"],
            bg=THEME["bg"],
            fg=THEME["muted"],
            font=(FONT_NAME, 9),
            anchor="w",
        )
        self.main_description_label.pack(side="left", padx=(12, 0))

        button_row = tk.Frame(top_bar, bg=THEME["bg"])
        button_row.pack(anchor="w", pady=(12, 0))

        self.folder_button = FlatButton(
            button_row,
            text=UI_TEXT["button_select_folder"],
            command=self.controller.choose_save_folder,
            width=14,
        )
        self.folder_button.pack(side="left")

        self.refresh_button = FlatButton(
            button_row,
            text=UI_TEXT["button_refresh"],
            command=self.controller.refresh,
            width=12,
        )
        self.refresh_button.pack(side="left", padx=(10, 0))

        self.save_folder_label = tk.Label(
            top_bar,
            text="",
            bg=THEME["bg"],
            fg=THEME["muted"],
            font=(FONT_NAME, 9),
            anchor="w",
        )
        self.save_folder_label.pack(fill="x", pady=(10, 0))

        self.drop_panel = tk.Frame(
            shell,
            bg=THEME["panel"],
            bd=1,
            relief="solid",
            cursor="hand2",
            highlightthickness=0,
        )
        self.drop_panel.pack(fill="both", expand=True, pady=(16, 16))

        self.drop_inner = tk.Frame(self.drop_panel, bg=THEME["panel"])
        self.drop_inner.place(relx=0.5, rely=0.5, anchor="center")

        self.panel_title = tk.Label(
            self.drop_inner,
            text="",
            bg=THEME["panel"],
            fg=THEME["text"],
            font=(FONT_NAME, 18, "bold"),
        )
        self.panel_title.pack()

        self.panel_subtitle = tk.Label(
            self.drop_inner,
            text="",
            bg=THEME["panel"],
            fg=THEME["muted"],
            font=(FONT_NAME, 11),
            wraplength=PANEL_PATH_WRAP_LENGTH,
            justify="left",
        )
        self.panel_subtitle.pack(pady=(10, 0))

        self.panel_meta = tk.Label(
            self.drop_inner,
            text="",
            bg=THEME["panel"],
            fg=THEME["muted"],
            font=(FONT_NAME, 9),
            wraplength=PANEL_PATH_WRAP_LENGTH,
            justify="left",
        )
        self.panel_meta.pack(pady=(10, 0))

        self.save_folder_tooltip = HoverTooltip(self.save_folder_label)
        self.panel_title_tooltip = HoverTooltip(self.panel_title)
        self.panel_subtitle_tooltip = HoverTooltip(self.panel_subtitle)
        self.panel_meta_tooltip = HoverTooltip(self.panel_meta)

        self._bind_drop_panel_click(self.drop_panel)
        self._bind_drop_panel_click(self.drop_inner)
        self._bind_drop_panel_click(self.panel_title)
        self._bind_drop_panel_click(self.panel_subtitle)
        self._bind_drop_panel_click(self.panel_meta)

        for widget in (self.drop_panel, self.drop_inner, self.panel_title, self.panel_subtitle, self.panel_meta):
            widget.bind("<Enter>", self._on_drop_panel_enter)
            widget.bind("<Leave>", self._on_drop_panel_leave)

        status_bar = tk.Frame(shell, bg=THEME["bg"])
        status_bar.pack(fill="x")

        self.status_label = tk.Label(
            status_bar,
            text="",
            bg=THEME["bg"],
            fg=THEME["text"],
            font=(FONT_NAME, 10, "bold"),
            anchor="w",
        )
        self.status_label.pack(side="left")

        self.detail_label = tk.Label(
            status_bar,
            text="",
            bg=THEME["bg"],
            fg=THEME["muted"],
            font=(FONT_NAME, 9),
            anchor="e",
        )
        self.detail_label.pack(side="right")

        footer = tk.Frame(shell, bg=THEME["bg"])
        footer.pack(fill="x", side="bottom", pady=(12, 0))

        left_block = tk.Frame(footer, bg=THEME["bg"])
        left_block.pack(side="left")

        tk.Label(
            left_block,
            text=UI_TEXT["footer_left"],
            bg=THEME["bg"],
            fg=THEME["muted"],
            font=(FONT_NAME, 8),
        ).pack(side="left")
        tk.Label(
            left_block,
            text=UI_TEXT["footer_left_suffix"],
            bg=THEME["bg"],
            fg=THEME["muted"],
            font=(FONT_NAME, 8),
        ).pack(side="left")

        right_block = tk.Frame(footer, bg=THEME["bg"])
        right_block.pack(side="right")

        self._create_footer_link(
            right_block,
            UI_TEXT["footer_link_1"],
            FOOTER_LINK_1_URL,
        ).pack(side="left")
        self._create_footer_text(right_block, UI_TEXT["footer_separator"]).pack(side="left")
        self._create_footer_link(
            right_block,
            UI_TEXT["footer_link_2"],
            FOOTER_LINK_2_URL,
        ).pack(side="left")
        self._create_footer_text(right_block, UI_TEXT["footer_separator"]).pack(side="left")
        self._create_footer_text(right_block, UI_TEXT["footer_copyright"]).pack(side="left")

    def _create_footer_text(self, parent, text: str) -> tk.Label:
        return tk.Label(
            parent,
            text=text,
            bg=THEME["bg"],
            fg=THEME["muted"],
            font=(FONT_NAME, 8),
        )

    def _create_footer_link(self, parent, text: str, url: str) -> tk.Label:
        label = tk.Label(
            parent,
            text=text,
            bg=THEME["bg"],
            fg=THEME["muted"],
            font=(FONT_NAME, 8),
            cursor="hand2",
        )
        label.bind("<Button-1>", lambda _event: webbrowser.open_new_tab(url))
        label.bind("<Enter>", lambda _event: label.configure(fg=THEME["accent"]))
        label.bind("<Leave>", lambda _event: label.configure(fg=THEME["muted"]))
        return label

    def _bind_drop_panel_click(self, widget) -> None:
        widget.bind("<Button-1>", lambda _event: self.controller.choose_pdf())

    def _on_drop_panel_enter(self, _event) -> None:
        if self.controller.busy:
            return
        self.drop_panel.configure(bg=THEME["accent_soft"], bd=1)
        self.drop_inner.configure(bg=THEME["accent_soft"])
        self.panel_title.configure(bg=THEME["accent_soft"])
        self.panel_subtitle.configure(bg=THEME["accent_soft"])
        self.panel_meta.configure(bg=THEME["accent_soft"])

    def _on_drop_panel_leave(self, _event) -> None:
        if self.controller.busy:
            return
        self._reset_drop_panel_colors()

    def _reset_drop_panel_colors(self) -> None:
        self.drop_panel.configure(bg=THEME["panel"])
        self.drop_inner.configure(bg=THEME["panel"])
        self.panel_title.configure(bg=THEME["panel"])
        self.panel_subtitle.configure(bg=THEME["panel"])
        self.panel_meta.configure(bg=THEME["panel"])

    def _configure_drag_and_drop(self) -> None:
        if not HAS_DND:
            return

        try:
            self.drop_panel.drop_target_register(DND_FILES)
            self.drop_panel.dnd_bind("<<Drop>>", self._on_drop)
        except Exception:
            pass

    def _on_drop(self, event) -> None:
        self.controller.handle_drop_data(event.data)

    def _poll_worker_queue(self) -> None:
        self.controller.process_worker_events()
        self.root.after(120, self._poll_worker_queue)

    def _animate_busy_status(self) -> None:
        if self.current_state in {"loading", "processing"}:
            animated = self.status_dots.next()
            self.detail_label.configure(
                text=f"{animated}  {self.current_detail}",
                fg=THEME["muted"],
            )
        self.root.after(450, self._animate_busy_status)

    def _center_window(self) -> None:
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() - WINDOW_WIDTH) // 2
        y = (self.root.winfo_screenheight() - WINDOW_HEIGHT) // 2
        self.root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}+{x}+{y}")

    def _set_label_text(
        self,
        label: tk.Label,
        tooltip: HoverTooltip | None,
        display_text: str,
        full_text: str,
    ) -> None:
        label.configure(text=display_text)
        if tooltip:
            tooltip.update_text(full_text if display_text != full_text else "")

    def show_idle(self) -> None:
        self.current_source = None
        self.current_output = None
        self.status_dots.reset()
        self.current_state = "idle"
        self.current_detail = UI_TEXT["idle_detail"]
        self._set_label_text(
            self.panel_title,
            self.panel_title_tooltip,
            UI_TEXT["empty_title"],
            UI_TEXT["empty_title"],
        )
        self._set_label_text(
            self.panel_subtitle,
            self.panel_subtitle_tooltip,
            UI_TEXT["empty_subtitle"],
            UI_TEXT["empty_subtitle"],
        )
        self._set_label_text(
            self.panel_meta,
            self.panel_meta_tooltip,
            "",
            "",
        )
        self._reset_drop_panel_colors()
        self.set_interaction_enabled(True)
        self.update_status("idle", UI_TEXT["idle_detail"])

    def show_file_context(self, source_path: Path | None, output_folder: Path | None) -> None:
        self.current_source = source_path
        self.current_output = output_folder
        if not source_path:
            self.show_idle()
            return

        source_name = source_path.name
        self._set_label_text(
            self.panel_title,
            self.panel_title_tooltip,
            shorten_file_name(source_name, SOURCE_FILE_TEXT_MAX),
            source_name,
        )
        if output_folder:
            output_folder_text = f"{UI_TEXT['output_folder_prefix']}: {output_folder.name}"
            self._set_label_text(
                self.panel_subtitle,
                self.panel_subtitle_tooltip,
                f"{UI_TEXT['output_folder_prefix']}: "
                f"{shorten_file_name(output_folder.name, OUTPUT_FOLDER_TEXT_MAX)}",
                output_folder_text,
            )
            self._set_label_text(
                self.panel_meta,
                self.panel_meta_tooltip,
                shorten_path(str(output_folder), RESULT_PATH_TEXT_MAX),
                str(output_folder),
            )
        else:
            self._set_label_text(
                self.panel_subtitle,
                self.panel_subtitle_tooltip,
                UI_TEXT["empty_subtitle"],
                UI_TEXT["empty_subtitle"],
            )
            self._set_label_text(
                self.panel_meta,
                self.panel_meta_tooltip,
                "",
                "",
            )

    def update_save_folder(self, folder: Path) -> None:
        full_text = f"{UI_TEXT['save_folder_prefix']}: {folder}"
        self._set_label_text(
            self.save_folder_label,
            self.save_folder_tooltip,
            f"{UI_TEXT['save_folder_prefix']}: "
            f"{shorten_path(str(folder), SAVE_FOLDER_TEXT_MAX)}",
            full_text,
        )

    def update_status(self, state_key: str, detail: str) -> None:
        self.current_state = state_key
        self.current_detail = detail
        self.status_dots.reset()

        color = {
            "idle": THEME["muted"],
            "loading": THEME["text"],
            "processing": THEME["text"],
            "complete": THEME["success"],
            "error": THEME["error"],
        }[state_key]

        state_label = {
            "idle": UI_TEXT["status_idle"],
            "loading": UI_TEXT["status_loading"],
            "processing": UI_TEXT["status_processing"],
            "complete": UI_TEXT["status_complete"],
            "error": UI_TEXT["status_error"],
        }[state_key]

        self.status_label.configure(text=state_label, fg=color)
        if state_key in {"loading", "processing"}:
            self.detail_label.configure(
                text=f"{UI_TEXT['progress_dots_1']}  {detail}",
                fg=THEME["muted"],
            )
        else:
            self.detail_label.configure(text=detail, fg=THEME["muted"])

    def update_progress(self, detail: str) -> None:
        self.current_state = "processing"
        self.current_detail = detail
        self.status_label.configure(text=UI_TEXT["status_processing"], fg=THEME["text"])
        self.detail_label.configure(text=detail, fg=THEME["muted"])

    def show_completion(self, output_folder: Path | None) -> None:
        messagebox.showinfo(
            UI_TEXT["dialog_complete_title"],
            UI_TEXT["dialog_complete_message"],
        )
        if output_folder:
            open_folder(output_folder)

    def show_error(self, message: str) -> None:
        self.update_status("error", message)
        messagebox.showerror(UI_TEXT["dialog_error_title"], message)

    def set_interaction_enabled(self, enabled: bool) -> None:
        self.folder_button.set_enabled(enabled)
        self.refresh_button.set_enabled(enabled)
        self.drop_panel.configure(cursor="hand2" if enabled else "arrow")


def resource_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def resource_path(name: str) -> Path:
    return resource_base_dir() / name


def get_common_icon_path() -> Path:
    return Path(__file__).resolve().parent / ".." / ".." / "02_assets" / "dake_icon.ico"


def default_downloads_path() -> str:
    downloads = Path.home() / "Downloads"
    return str(downloads if downloads.exists() else Path.home())


def make_safe_stem(name: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*]+', "_", name).strip(" .")
    return cleaned or "split_pdf"


def shorten_middle(text: str, max_length: int, keep_start: int, keep_end: int) -> str:
    if len(text) <= max_length or max_length <= 3:
        return text

    available = max_length - 3
    keep_start = max(1, min(keep_start, available - 1))
    keep_end = max(1, min(keep_end, available - keep_start))
    return f"{text[:keep_start]}...{text[-keep_end:]}"


def shorten_file_name(name: str, max_length: int) -> str:
    if len(name) <= max_length:
        return name

    suffix = Path(name).suffix
    keep_end = max(12, len(suffix) + 8) if suffix else 12
    keep_start = max(10, max_length - 3 - keep_end)
    return shorten_middle(name, max_length, keep_start, keep_end)


def shorten_path(path: str, max_length: int) -> str:
    normalized = path.replace("/", "\\")
    if len(normalized) <= max_length:
        return normalized

    drive, rest = os.path.splitdrive(normalized)
    parts = [part for part in rest.split("\\") if part]
    if not parts:
        return shorten_middle(normalized, max_length, 12, 16)

    root_prefix = f"{drive}\\" if drive else ""
    tail_parts = parts[-2:] if len(parts) >= 2 else parts[-1:]
    tail = "\\".join(tail_parts)
    candidate = f"{root_prefix}...\\{tail}"
    if len(candidate) <= max_length:
        return candidate

    shortened_tail = shorten_file_name(
        tail,
        max(12, max_length - len(root_prefix) - 4),
    )
    return f"{root_prefix}...\\{shortened_tail}"


def open_folder(path: Path) -> None:
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(path))
        else:
            webbrowser.open_new_tab(path.as_uri())
    except Exception:
        pass


def create_root() -> tk.Tk:
    if HAS_DND:
        return TkinterDnD.Tk()
    return tk.Tk()


def main() -> None:
    root = create_root()
    config = AppConfig(resource_base_dir() / CONFIG_NAME)
    notifier = WorkerNotifier()
    service = PdfSplitService()
    controller = SplitController(config=config, notifier=notifier, service=service)
    SplitOneApp(root=root, controller=controller)
    root.mainloop()


if __name__ == "__main__":
    main()
