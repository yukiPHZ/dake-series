from __future__ import annotations

import json
import os
import queue
import re
import shutil
import sys
import threading
import webbrowser
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkfont

import fitz

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore

    DND_ENABLED = True
except Exception:
    DND_FILES = None
    TkinterDnD = None
    DND_ENABLED = False

APP_NAME = "PDF→画像"
APP_NAME_EN = "DakePDF to Images"
WINDOW_TITLE = "PDF→画像"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"
INTERNAL_FOLDER_NAME = "DAKE_PDF_ToImages"
CONFIG_FILE_NAME = "dake_pdf_to_images_config.json"

UI_TEXT = {
    "main_title": "PDFを画像に変換する",
    "main_description": "PDFを追加すると、各ページを1枚ずつ画像として保存します。",
    "list_title": "追加したPDF",
    "empty_title": "PDFを追加してください",
    "empty_subtitle_dnd": "ドラッグ＆ドロップ または クリックして追加",
    "empty_subtitle_button": "クリックしてPDFを追加",
    "summary_empty": "PDFはまだ追加されていません。",
    "summary_ready": "{count}件のPDF / 合計 {pages} ページ",
    "summary_partial": "{count}件のPDF / 変換できるPDF {ready_count}件",
    "column_name": "PDFファイル",
    "column_pages": "ページ数",
    "column_state": "状態",
    "page_count_format": "{count}ページ",
    "page_count_unknown": "-",
    "list_state_ready": "準備完了",
    "list_state_error": "読み込み失敗",
    "button_add": "PDFを追加",
    "button_select_folder": "保存先を選ぶ",
    "button_refresh": "リフレッシュ",
    "button_execute": "画像に変換して保存",
    "output_prefix": "保存先:",
    "output_hint": "PDFごとに出力フォルダを自動作成します。",
    "status_idle": "未選択",
    "status_loading": "読み込み中",
    "status_ready": "準備完了",
    "status_processing": "処理中",
    "status_saving": "保存中",
    "status_complete": "保存完了",
    "status_error": "エラー",
    "status_detail_idle": "PDFを追加すると、ここに状態が表示されます。",
    "status_detail_loading": "PDFを読み込んでいます。",
    "status_detail_ready": "{count}件のPDFを保存できます。",
    "status_detail_ready_partial": "{ready_count}件は保存できます。{error_count}件は読み込めませんでした。",
    "status_detail_processing": "{current}/{total}件目を処理しています: {name}",
    "status_detail_saving": "{current}/{total}ページを書き出しています: {name}",
    "status_detail_complete": "{count}件のPDFを保存しました。",
    "status_detail_complete_partial": "{count}件のPDFを保存しました。{error_count}件は保存できませんでした。",
    "status_detail_error": "画像への変換を完了できませんでした。",
    "dialog_select_pdf_title": "PDFを選択",
    "dialog_select_output_title": "保存先を選ぶ",
    "dialog_complete_title": "変換が完了しました",
    "dialog_partial_title": "一部のPDFは保存できませんでした",
    "dialog_error_title": "変換できませんでした",
    "dialog_open_folder_error_title": "保存先を開けませんでした",
    "dialog_pdf_filter_label": "PDF",
    "message_no_ready_pdf": "保存できるPDFがありません。",
    "message_no_new_pdf": "新しく追加できるPDFがありませんでした。",
    "message_skip_non_pdf": "PDF以外のファイルを {count} 件スキップしました。",
    "message_skip_duplicate": "追加済みのPDFを {count} 件スキップしました。",
    "message_output_open_error": "保存先フォルダを開けませんでした。",
    "message_no_pages": "ページが見つかりませんでした。",
    "message_unknown_error": "原因を特定できませんでした。",
    "message_complete": "{success_count}件のPDFを画像に変換しました。\n保存先: {folder}",
    "message_complete_partial": "{success_count}件のPDFを画像に変換しました。\n保存先: {folder}\n\n保存できなかったPDF:\n{errors}",
    "message_complete_error": "画像への変換を完了できませんでした。\n\n対象:\n{errors}",
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
    "success_bg": "#EAFBF3",
    "error": "#D92D20",
    "error_bg": "#FDECEC",
    "soft": "#EEF2F7",
}

LINKS = {
    "assessment": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "instagram": "https://instagram.com/kikuta.shimarisu_fudosan",
}

STATUS_COLORS = {
    "status_idle": (THEME["soft"], THEME["muted"]),
    "status_loading": (THEME["selection_bg"], THEME["accent"]),
    "status_ready": (THEME["selection_bg"], THEME["accent"]),
    "status_processing": (THEME["selection_bg"], THEME["accent"]),
    "status_saving": (THEME["selection_bg"], THEME["accent"]),
    "status_complete": (THEME["success_bg"], THEME["success"]),
    "status_error": (THEME["error_bg"], THEME["error"]),
}

WINDOW_SIZE = "940x760"
WINDOW_MIN_SIZE = (860, 680)
RENDER_SCALE = 2.0
POLL_INTERVAL_MS = 80


@dataclass
class PdfItem:
    path: Path
    page_count: int | None = None
    error: str | None = None

    @property
    def ready(self) -> bool:
        return self.page_count is not None and self.page_count > 0 and self.error is None


def get_common_icon_path() -> Path:
    return Path(__file__).resolve().parent / ".." / ".." / "02_assets" / "dake_icon.ico"


def default_output_dir() -> Path:
    downloads = Path.home() / "Downloads"
    return downloads if downloads.exists() else Path.home()


def sanitize_name(name: str) -> str:
    value = re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", name).strip(" .")
    return value or "output"


def humanize_error(exc: Exception) -> str:
    message = str(exc).strip().replace("\n", " ")
    return message or UI_TEXT["message_unknown_error"]


class ConfigStore:
    def __init__(self) -> None:
        base_dir = Path(os.environ.get("LOCALAPPDATA", str(Path.home())))
        self.config_dir = base_dir / INTERNAL_FOLDER_NAME
        self.config_path = self.config_dir / CONFIG_FILE_NAME

    def load_output_dir(self) -> Path:
        fallback = default_output_dir()
        if not self.config_path.exists():
            return fallback
        try:
            data = json.loads(self.config_path.read_text(encoding="utf-8"))
        except Exception:
            return fallback
        raw_path = data.get("last_output_dir")
        if not raw_path:
            return fallback
        candidate = Path(raw_path)
        if candidate.exists() and candidate.is_dir():
            return candidate
        return fallback

    def save_output_dir(self, output_dir: Path) -> None:
        self.config_dir.mkdir(parents=True, exist_ok=True)
        payload = {"last_output_dir": str(output_dir)}
        self.config_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


class PdfImageService:
    def inspect(self, pdf_path: Path) -> int:
        with fitz.open(pdf_path) as document:
            if document.page_count < 1:
                raise ValueError(UI_TEXT["message_no_pages"])
            return document.page_count

    def create_output_dir(self, output_root: Path, pdf_path: Path) -> Path:
        base_name = sanitize_name(pdf_path.stem)
        candidate = output_root / base_name
        if not candidate.exists():
            candidate.mkdir(parents=True, exist_ok=False)
            return candidate

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        candidate = output_root / f"{base_name}_{timestamp}"
        if not candidate.exists():
            candidate.mkdir(parents=True, exist_ok=False)
            return candidate

        index = 1
        while True:
            candidate = output_root / f"{base_name}_{timestamp}_{index:02d}"
            if not candidate.exists():
                candidate.mkdir(parents=True, exist_ok=False)
                return candidate
            index += 1

    def convert(self, pdf_path: Path, output_root: Path, page_callback=None) -> tuple[Path, int]:
        with fitz.open(pdf_path) as document:
            page_count = document.page_count
            if page_count < 1:
                raise ValueError(UI_TEXT["message_no_pages"])
            output_dir = self.create_output_dir(output_root, pdf_path)
            digits = max(3, len(str(page_count)))
            try:
                for page_index in range(page_count):
                    page = document.load_page(page_index)
                    pixmap = page.get_pixmap(matrix=fitz.Matrix(RENDER_SCALE, RENDER_SCALE), alpha=False)
                    output_path = output_dir / f"page_{page_index + 1:0{digits}d}.png"
                    pixmap.save(output_path)
                    if page_callback is not None:
                        page_callback(page_index + 1, page_count)
            except Exception:
                shutil.rmtree(output_dir, ignore_errors=True)
                raise
            return output_dir, page_count


class App:
    def __init__(self) -> None:
        self.root = TkinterDnD.Tk() if DND_ENABLED else tk.Tk()
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(*WINDOW_MIN_SIZE)
        self.root.configure(bg=THEME["background"])

        self.font_family = self._pick_font_family()
        self.queue: queue.Queue = queue.Queue()
        self.config = ConfigStore()
        self.service = PdfImageService()
        self.items: list[PdfItem] = []
        self.busy = False
        self.output_dir = self.config.load_output_dir()
        self._build_styles()
        self._build_ui()
        self._apply_icon()
        self._bind_drop_targets()
        self._update_output_info()
        self._update_items([])
        self._update_status("status_idle", UI_TEXT["status_detail_idle"])
        self.root.after(POLL_INTERVAL_MS, self._poll_events)

    def _pick_font_family(self) -> str:
        available = set(tkfont.families(self.root))
        for family in ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo"):
            if family in available:
                return family
        return "TkDefaultFont"

    def _build_styles(self) -> None:
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure(
            "Secondary.TButton",
            font=(self.font_family, 10),
            padding=(14, 10),
            background=THEME["card"],
            foreground=THEME["text"],
            borderwidth=1,
            relief="solid",
        )
        style.map(
            "Secondary.TButton",
            background=[("active", THEME["selection_bg"]), ("disabled", THEME["soft"])],
            foreground=[("disabled", THEME["muted"])],
            bordercolor=[("disabled", THEME["border"])],
        )
        style.configure(
            "Primary.TButton",
            font=(self.font_family, 10, "bold"),
            padding=(16, 10),
            background=THEME["accent"],
            foreground="#FFFFFF",
            borderwidth=0,
        )
        style.map(
            "Primary.TButton",
            background=[("active", THEME["accent_hover"]), ("disabled", THEME["border"])],
            foreground=[("disabled", "#FFFFFF")],
        )
        style.configure(
            "Dake.Treeview",
            font=(self.font_family, 10),
            rowheight=34,
            background=THEME["card"],
            fieldbackground=THEME["card"],
            foreground=THEME["text"],
            borderwidth=0,
        )
        style.configure(
            "Dake.Treeview.Heading",
            font=(self.font_family, 10, "bold"),
            background=THEME["soft"],
            foreground=THEME["text"],
            relief="flat",
            padding=(8, 8),
        )
        style.map(
            "Dake.Treeview",
            background=[("selected", THEME["selection_bg"])],
            foreground=[("selected", THEME["text"])],
        )
        style.configure(
            "Dake.Vertical.TScrollbar",
            troughcolor=THEME["soft"],
            background=THEME["border"],
            bordercolor=THEME["soft"],
            arrowcolor=THEME["muted"],
        )
        style.configure(
            "Dake.Horizontal.TProgressbar",
            troughcolor=THEME["soft"],
            background=THEME["accent"],
            bordercolor=THEME["border"],
            lightcolor=THEME["accent"],
            darkcolor=THEME["accent"],
        )

    def _build_ui(self) -> None:
        container = tk.Frame(self.root, bg=THEME["background"])
        container.pack(fill="both", expand=True, padx=24, pady=20)

        header = tk.Frame(container, bg=THEME["background"])
        header.pack(fill="x")
        tk.Label(
            header,
            text=UI_TEXT["main_title"],
            bg=THEME["background"],
            fg=THEME["text"],
            font=(self.font_family, 22, "bold"),
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            header,
            text=UI_TEXT["main_description"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 11),
            anchor="w",
        ).pack(fill="x", pady=(8, 0))

        self.main_card = tk.Frame(
            container,
            bg=THEME["card"],
            highlightbackground=THEME["border"],
            highlightthickness=1,
        )
        self.main_card.pack(fill="both", expand=True, pady=(18, 0))

        card_header = tk.Frame(self.main_card, bg=THEME["card"])
        card_header.pack(fill="x", padx=18, pady=(18, 8))
        tk.Label(
            card_header,
            text=UI_TEXT["list_title"],
            bg=THEME["card"],
            fg=THEME["text"],
            font=(self.font_family, 12, "bold"),
            anchor="w",
        ).pack(fill="x")
        self.summary_label = tk.Label(
            card_header,
            text=UI_TEXT["summary_empty"],
            bg=THEME["card"],
            fg=THEME["muted"],
            font=(self.font_family, 10),
            anchor="w",
        )
        self.summary_label.pack(fill="x", pady=(6, 0))

        self.card_body = tk.Frame(self.main_card, bg=THEME["card"])
        self.card_body.pack(fill="both", expand=True, padx=18, pady=(0, 18))

        self.empty_state = tk.Frame(
            self.card_body,
            bg=THEME["card"],
            highlightbackground=THEME["border"],
            highlightthickness=1,
            cursor="hand2",
        )
        self.empty_state.pack(fill="both", expand=True)
        self.empty_state.bind("<Button-1>", lambda _event: self.choose_pdfs())

        empty_inner = tk.Frame(self.empty_state, bg=THEME["card"])
        empty_inner.place(relx=0.5, rely=0.5, anchor="center")
        empty_title = tk.Label(
            empty_inner,
            text=UI_TEXT["empty_title"],
            bg=THEME["card"],
            fg=THEME["text"],
            font=(self.font_family, 18, "bold"),
            cursor="hand2",
        )
        empty_title.pack()
        empty_title.bind("<Button-1>", lambda _event: self.choose_pdfs())
        empty_subtitle = UI_TEXT["empty_subtitle_dnd"] if DND_ENABLED else UI_TEXT["empty_subtitle_button"]
        empty_detail = tk.Label(
            empty_inner,
            text=empty_subtitle,
            bg=THEME["card"],
            fg=THEME["muted"],
            font=(self.font_family, 11),
            cursor="hand2",
        )
        empty_detail.pack(pady=(10, 0))
        empty_detail.bind("<Button-1>", lambda _event: self.choose_pdfs())

        self.table_frame = tk.Frame(self.card_body, bg=THEME["card"])
        self.table_frame.grid_columnconfigure(0, weight=1)
        self.table_frame.grid_rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            self.table_frame,
            columns=("name", "pages", "state"),
            show="headings",
            style="Dake.Treeview",
            selectmode="browse",
        )
        self.tree.heading("name", text=UI_TEXT["column_name"])
        self.tree.heading("pages", text=UI_TEXT["column_pages"])
        self.tree.heading("state", text=UI_TEXT["column_state"])
        self.tree.column("name", width=520, anchor="w", stretch=True)
        self.tree.column("pages", width=120, anchor="center", stretch=False)
        self.tree.column("state", width=180, anchor="center", stretch=False)
        self.tree.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(
            self.table_frame,
            orient="vertical",
            command=self.tree.yview,
            style="Dake.Vertical.TScrollbar",
        )
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)

        actions = tk.Frame(container, bg=THEME["background"])
        actions.pack(fill="x", pady=(14, 0))

        self.add_button = ttk.Button(actions, text=UI_TEXT["button_add"], style="Secondary.TButton", command=self.choose_pdfs)
        self.add_button.pack(side="left", padx=(0, 10))
        self.select_folder_button = ttk.Button(
            actions,
            text=UI_TEXT["button_select_folder"],
            style="Secondary.TButton",
            command=self.choose_output_dir,
        )
        self.select_folder_button.pack(side="left", padx=(0, 10))
        self.refresh_button = ttk.Button(
            actions,
            text=UI_TEXT["button_refresh"],
            style="Secondary.TButton",
            command=self.refresh_items,
        )
        self.refresh_button.pack(side="left", padx=(0, 10))
        self.execute_button = ttk.Button(
            actions,
            text=UI_TEXT["button_execute"],
            style="Primary.TButton",
            command=self.start_conversion,
        )
        self.execute_button.pack(side="right")

        self.output_frame = tk.Frame(container, bg=THEME["background"])
        self.output_frame.pack(fill="x", pady=(12, 0))
        self.output_path_label = tk.Label(
            self.output_frame,
            text="",
            bg=THEME["background"],
            fg=THEME["text"],
            font=(self.font_family, 10),
            anchor="w",
            justify="left",
            wraplength=860,
        )
        self.output_path_label.pack(fill="x")
        self.output_hint_label = tk.Label(
            self.output_frame,
            text=UI_TEXT["output_hint"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 9),
            anchor="w",
            justify="left",
        )
        self.output_hint_label.pack(fill="x", pady=(4, 0))

        status_card = tk.Frame(
            container,
            bg=THEME["card"],
            highlightbackground=THEME["border"],
            highlightthickness=1,
        )
        status_card.pack(fill="x", pady=(16, 0))
        status_top = tk.Frame(status_card, bg=THEME["card"])
        status_top.pack(fill="x", padx=16, pady=(14, 8))
        self.status_chip = tk.Label(
            status_top,
            text=UI_TEXT["status_idle"],
            bg=THEME["soft"],
            fg=THEME["muted"],
            font=(self.font_family, 10, "bold"),
            padx=10,
            pady=6,
        )
        self.status_chip.pack(side="left")
        self.status_detail_label = tk.Label(
            status_top,
            text=UI_TEXT["status_detail_idle"],
            bg=THEME["card"],
            fg=THEME["muted"],
            font=(self.font_family, 10),
            anchor="w",
            justify="left",
        )
        self.status_detail_label.pack(side="left", fill="x", expand=True, padx=(12, 0))
        self.progress = ttk.Progressbar(status_card, mode="determinate", style="Dake.Horizontal.TProgressbar")
        self.progress.pack(fill="x", padx=16, pady=(0, 14))

        footer = tk.Frame(container, bg=THEME["background"])
        footer.pack(fill="x", pady=(14, 0))

        footer_left = tk.Frame(footer, bg=THEME["background"])
        footer_left.pack(side="left", fill="x", expand=True)
        tk.Label(
            footer_left,
            text=UI_TEXT["footer_left"],
            bg=THEME["background"],
            fg=THEME["text"],
            font=(self.font_family, 9, "bold"),
            anchor="w",
        ).pack(anchor="w")
        tk.Label(
            footer_left,
            text=UI_TEXT["footer_subtitle"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 9),
            anchor="w",
        ).pack(anchor="w", pady=(4, 0))

        footer_right = tk.Frame(footer, bg=THEME["background"])
        footer_right.pack(side="right")
        link1 = tk.Label(
            footer_right,
            text=UI_TEXT["footer_link_1"],
            bg=THEME["background"],
            fg=THEME["accent"],
            font=(self.font_family, 9, "underline"),
            cursor="hand2",
        )
        link1.pack(side="left")
        link1.bind("<Button-1>", lambda _event: webbrowser.open(LINKS["assessment"], new=2))
        tk.Label(
            footer_right,
            text=UI_TEXT["footer_separator"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 9),
        ).pack(side="left")
        link2 = tk.Label(
            footer_right,
            text=UI_TEXT["footer_link_2"],
            bg=THEME["background"],
            fg=THEME["accent"],
            font=(self.font_family, 9, "underline"),
            cursor="hand2",
        )
        link2.pack(side="left")
        link2.bind("<Button-1>", lambda _event: webbrowser.open(LINKS["instagram"], new=2))
        tk.Label(
            footer_right,
            text=UI_TEXT["footer_separator"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 9),
        ).pack(side="left")
        tk.Label(
            footer_right,
            text=UI_TEXT["footer_copyright"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 9),
        ).pack(side="left")

    def _apply_icon(self) -> None:
        try:
            if getattr(sys, "frozen", False):
                return
            icon_path = get_common_icon_path().resolve()
            if icon_path.exists():
                self.root.iconbitmap(str(icon_path))
        except Exception:
            pass

    def _bind_drop_targets(self) -> None:
        if not DND_ENABLED:
            return
        for widget in (self.root, self.empty_state, self.main_card):
            try:
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", self._handle_drop)
            except Exception:
                pass

    def _poll_events(self) -> None:
        while True:
            try:
                event = self.queue.get_nowait()
            except queue.Empty:
                break
            self._handle_event(event)
        self.root.after(POLL_INTERVAL_MS, self._poll_events)

    def _handle_event(self, event: dict[str, object]) -> None:
        event_type = str(event["type"])
        if event_type == "busy":
            self.busy = bool(event["value"])
            self._update_buttons()
            return
        if event_type == "items":
            self._update_items(list(event["items"]))
            return
        if event_type == "status":
            self._update_status(
                str(event["status_key"]),
                str(event["detail"]),
                int(event.get("current", 0)),
                int(event.get("total", 1)),
            )
            return
        if event_type == "dialog":
            self._show_dialog(
                str(event["kind"]),
                str(event["title"]),
                str(event["message"]),
                bool(event.get("open_folder", False)),
            )

    def _post(self, event_type: str, **payload: object) -> None:
        self.queue.put({"type": event_type, **payload})

    def _show_dialog(self, kind: str, title: str, message: str, open_folder: bool) -> None:
        if kind == "info":
            messagebox.showinfo(title, message, parent=self.root)
        elif kind == "warning":
            messagebox.showwarning(title, message, parent=self.root)
        else:
            messagebox.showerror(title, message, parent=self.root)
        if open_folder:
            self.open_output_dir()

    def _update_items(self, items: list[PdfItem]) -> None:
        self.items = items
        for item_id in self.tree.get_children():
            self.tree.delete(item_id)

        if not self.items:
            self.summary_label.config(text=UI_TEXT["summary_empty"])
            self.table_frame.pack_forget()
            self.empty_state.pack(fill="both", expand=True)
        else:
            ready_count = len([item for item in self.items if item.ready])
            total_pages = sum(item.page_count or 0 for item in self.items if item.ready)
            if ready_count == len(self.items):
                summary = UI_TEXT["summary_ready"].format(count=len(self.items), pages=total_pages)
            else:
                summary = UI_TEXT["summary_partial"].format(count=len(self.items), ready_count=ready_count)
            self.summary_label.config(text=summary)
            self.empty_state.pack_forget()
            self.table_frame.pack(fill="both", expand=True)
            for item in self.items:
                page_text = UI_TEXT["page_count_format"].format(count=item.page_count) if item.page_count else UI_TEXT["page_count_unknown"]
                state_text = UI_TEXT["list_state_ready"] if item.ready else UI_TEXT["list_state_error"]
                if item.error:
                    state_text = item.error
                self.tree.insert("", "end", values=(item.path.name, page_text, state_text))

        self._update_buttons()

    def _update_buttons(self) -> None:
        disabled = tk.DISABLED if self.busy else tk.NORMAL
        ready_count = len([item for item in self.items if item.ready])
        self.add_button.config(state=disabled)
        self.select_folder_button.config(state=disabled)
        self.refresh_button.config(state=tk.DISABLED if self.busy or not self.items else tk.NORMAL)
        self.execute_button.config(state=tk.NORMAL if (not self.busy and ready_count > 0) else tk.DISABLED)

    def _update_output_info(self) -> None:
        self.output_path_label.config(text=f"{UI_TEXT['output_prefix']} {self.output_dir}")

    def _update_status(self, status_key: str, detail: str, current: int = 0, total: int = 1) -> None:
        bg, fg = STATUS_COLORS.get(status_key, STATUS_COLORS["status_idle"])
        self.status_chip.config(text=UI_TEXT[status_key], bg=bg, fg=fg)
        self.status_detail_label.config(text=detail)
        self.progress.configure(maximum=max(1, total), value=max(0, current))

    def _handle_drop(self, event) -> str:
        if self.busy or not DND_ENABLED:
            return "break"
        dropped = list(self.root.tk.splitlist(event.data))
        self.add_pdfs(dropped)
        return "break"

    def choose_pdfs(self) -> None:
        if self.busy:
            return
        selected = filedialog.askopenfilenames(
            title=UI_TEXT["dialog_select_pdf_title"],
            filetypes=[(UI_TEXT["dialog_pdf_filter_label"], "*.pdf")],
            parent=self.root,
        )
        if selected:
            self.add_pdfs(list(selected))

    def choose_output_dir(self) -> None:
        if self.busy:
            return
        selected = filedialog.askdirectory(
            title=UI_TEXT["dialog_select_output_title"],
            initialdir=str(self.output_dir),
            mustexist=True,
            parent=self.root,
        )
        if not selected:
            return
        self.output_dir = Path(selected)
        self.config.save_output_dir(self.output_dir)
        self._update_output_info()

    def refresh_items(self) -> None:
        if self.busy:
            return
        self._update_items([])
        self._update_status("status_idle", UI_TEXT["status_detail_idle"])

    def add_pdfs(self, raw_paths: list[str]) -> None:
        existing = {str(item.path).lower() for item in self.items}
        new_paths: list[Path] = []
        skipped_non_pdf = 0
        skipped_duplicate = 0

        for raw_path in raw_paths:
            candidate = Path(str(raw_path).strip().strip('"')).expanduser()
            if not candidate.exists() or candidate.is_dir() or candidate.suffix.lower() != ".pdf":
                skipped_non_pdf += 1
                continue
            resolved = candidate.resolve()
            key = str(resolved).lower()
            if key in existing:
                skipped_duplicate += 1
                continue
            existing.add(key)
            new_paths.append(resolved)

        if not new_paths:
            messages = [UI_TEXT["message_no_new_pdf"]]
            if skipped_non_pdf > 0:
                messages.append(UI_TEXT["message_skip_non_pdf"].format(count=skipped_non_pdf))
            if skipped_duplicate > 0:
                messages.append(UI_TEXT["message_skip_duplicate"].format(count=skipped_duplicate))
            messagebox.showwarning(APP_NAME, "\n".join(messages), parent=self.root)
            return

        self._post("busy", value=True)
        self._post("status", status_key="status_loading", detail=UI_TEXT["status_detail_loading"], current=0, total=1)
        worker = threading.Thread(
            target=self._load_worker,
            args=(new_paths, skipped_non_pdf, skipped_duplicate),
            daemon=True,
        )
        worker.start()

    def _load_worker(self, new_paths: list[Path], skipped_non_pdf: int, skipped_duplicate: int) -> None:
        items = list(self.items)
        for pdf_path in new_paths:
            try:
                items.append(PdfItem(path=pdf_path, page_count=self.service.inspect(pdf_path)))
            except Exception as exc:
                items.append(PdfItem(path=pdf_path, error=humanize_error(exc)))

        ready_count = len([item for item in items if item.ready])
        error_count = len([item for item in items if item.error])
        if ready_count > 0 and error_count == 0:
            detail = UI_TEXT["status_detail_ready"].format(count=ready_count)
            status_key = "status_ready"
        elif ready_count > 0:
            detail = UI_TEXT["status_detail_ready_partial"].format(ready_count=ready_count, error_count=error_count)
            status_key = "status_ready"
        else:
            detail = UI_TEXT["status_detail_error"]
            status_key = "status_error"

        if skipped_non_pdf > 0:
            detail = f"{detail} {UI_TEXT['message_skip_non_pdf'].format(count=skipped_non_pdf)}"
        if skipped_duplicate > 0:
            detail = f"{detail} {UI_TEXT['message_skip_duplicate'].format(count=skipped_duplicate)}"

        self._post("items", items=items)
        self._post("status", status_key=status_key, detail=detail, current=0, total=1)
        self._post("busy", value=False)

    def start_conversion(self) -> None:
        if self.busy:
            return
        ready_items = [item for item in self.items if item.ready]
        if not ready_items:
            messagebox.showwarning(APP_NAME, UI_TEXT["message_no_ready_pdf"], parent=self.root)
            return
        self._post("busy", value=True)
        total_pages = sum(item.page_count or 0 for item in ready_items)
        worker = threading.Thread(
            target=self._convert_worker,
            args=(ready_items, total_pages),
            daemon=True,
        )
        worker.start()

    def _convert_worker(self, ready_items: list[PdfItem], total_pages: int) -> None:
        failures: list[str] = []
        success_count = 0
        saved_pages = 0

        try:
            self.output_dir.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            self._post("status", status_key="status_error", detail=UI_TEXT["status_detail_error"], current=0, total=1)
            self._post(
                "dialog",
                kind="error",
                title=UI_TEXT["dialog_error_title"],
                message=humanize_error(exc),
                open_folder=False,
            )
            self._post("busy", value=False)
            return

        for index, item in enumerate(ready_items, start=1):
            self._post(
                "status",
                status_key="status_processing",
                detail=UI_TEXT["status_detail_processing"].format(
                    current=index,
                    total=len(ready_items),
                    name=item.path.name,
                ),
                current=saved_pages,
                total=total_pages,
            )
            try:
                self.service.convert(
                    item.path,
                    self.output_dir,
                    page_callback=lambda current_page, _page_total, name=item.path.name, base=saved_pages: self._post(
                        "status",
                        status_key="status_saving",
                        detail=UI_TEXT["status_detail_saving"].format(
                            current=base + current_page,
                            total=total_pages,
                            name=name,
                        ),
                        current=base + current_page,
                        total=total_pages,
                    ),
                )
                success_count += 1
                saved_pages += item.page_count or 0
            except Exception as exc:
                failures.append(f"{item.path.name}: {humanize_error(exc)}")

        if success_count > 0 and not failures:
            self._post(
                "status",
                status_key="status_complete",
                detail=UI_TEXT["status_detail_complete"].format(count=success_count),
                current=total_pages,
                total=total_pages,
            )
            self._post(
                "dialog",
                kind="info",
                title=UI_TEXT["dialog_complete_title"],
                message=UI_TEXT["message_complete"].format(
                    success_count=success_count,
                    folder=str(self.output_dir),
                ),
                open_folder=True,
            )
        elif success_count > 0:
            self._post(
                "status",
                status_key="status_complete",
                detail=UI_TEXT["status_detail_complete_partial"].format(
                    count=success_count,
                    error_count=len(failures),
                ),
                current=saved_pages,
                total=total_pages,
            )
            self._post(
                "dialog",
                kind="warning",
                title=UI_TEXT["dialog_partial_title"],
                message=UI_TEXT["message_complete_partial"].format(
                    success_count=success_count,
                    folder=str(self.output_dir),
                    errors="\n".join(failures),
                ),
                open_folder=True,
            )
        else:
            self._post(
                "status",
                status_key="status_error",
                detail=UI_TEXT["status_detail_error"],
                current=0,
                total=1,
            )
            self._post(
                "dialog",
                kind="error",
                title=UI_TEXT["dialog_error_title"],
                message=UI_TEXT["message_complete_error"].format(
                    errors="\n".join(failures) or UI_TEXT["message_unknown_error"]
                ),
                open_folder=False,
            )

        self._post("busy", value=False)

    def open_output_dir(self) -> None:
        try:
            os.startfile(str(self.output_dir))
        except Exception:
            messagebox.showwarning(
                UI_TEXT["dialog_open_folder_error_title"],
                UI_TEXT["message_output_open_error"],
                parent=self.root,
            )

    def run(self) -> None:
        self.root.mainloop()


def main() -> None:
    app = App()
    app.run()


if __name__ == "__main__":
    main()
