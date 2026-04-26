# -*- coding: utf-8 -*-
from __future__ import annotations

import base64
import ctypes
import os
import subprocess
import sys
import urllib.parse
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, font as tkfont, messagebox
import tkinter as tk

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None  # type: ignore[assignment]

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore

    DND_ENABLED = True
except Exception:
    DND_FILES = None
    TkinterDnD = None
    DND_ENABLED = False


APP_NAME = "聞くDAKE"
WINDOW_TITLE = "聞くDAKE"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "brand_series": "シンプルそれDAKEシリーズ",
    "main_title": "○ と ？ だけ付ける",
    "main_description": "確認してほしい場所を、○ と ？ だけで伝えます。",
    "button_select_pdf": "PDFを選ぶ",
    "button_save": "保存する",
    "button_undo": "1つ戻す",
    "button_prev": "前へ",
    "button_next": "次へ",
    "empty_title": "PDFを追加してください",
    "empty_subtitle": "ドラッグ＆ドロップ または PDFを選ぶ",
    "drop_title": "ここにPDFをドロップ",
    "status_idle": "未選択",
    "status_ready": "準備完了",
    "status_processing": "処理中",
    "status_saved": "保存しました",
    "status_error": "エラー",
    "status_loading": "読み込み中",
    "status_circle_added": "○を付けました",
    "status_question_added": "？を付けました",
    "status_undo": "1つ戻しました",
    "status_no_undo": "戻すマークがありません",
    "message_no_pdf": "PDFを選択してください。",
    "message_save_complete": "保存しました。",
    "message_save_failed": "保存できませんでした。",
    "message_invalid_file": "PDFファイルを選択してください。",
    "message_open_failed": "PDFを開けませんでした。",
    "message_dependency_missing": "PyMuPDF が見つかりません。pymupdf をインストールしてください。",
    "message_no_pages": "PDFのページを読み込めませんでした。",
    "message_same_file": "元PDFとは別名で保存してください。",
    "message_no_marks_confirm": "マークがありません。このまま保存しますか？",
    "message_no_folder_access": "保存先にアクセスできませんでした。",
    "dialog_open_title": "PDFを選ぶ",
    "dialog_save_title": "保存先を選ぶ",
    "dialog_error_title": "エラー",
    "dialog_complete_title": "保存完了",
    "dialog_confirm_title": "確認",
    "dialog_pdf_filter_label": "PDFファイル",
    "dialog_all_filter_label": "すべてのファイル",
    "page_format": "{current} / {total}",
    "save_suffix": "_確認マーク",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_link_1": "戸建買取査定",
    "footer_link_2": "Instagram",
    "footer_separator": " ｜ ",
    "footer_copyright": "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta",
}

THEME = {
    "background": "#F6F7F9",
    "panel": "#FFFFFF",
    "text": "#1E2430",
    "muted": "#667085",
    "border": "#E6EAF0",
    "soft": "#EEF2F7",
    "accent": "#2F6FED",
    "accent_hover": "#2458BF",
    "accent_text": "#FFFFFF",
    "danger": "#E11D48",
    "danger_soft": "#FFF1F3",
    "success": "#12B76A",
    "success_soft": "#EAFBF3",
}

LINKS = {
    "footer_link_1": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "footer_link_2": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

FONT_CANDIDATES = ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo")
WINDOW_SIZE = "1060x760"
WINDOW_MIN_SIZE = (880, 620)
APP_USER_MODEL_ID = "Shimarisu.DakePDFAskMark"

PAGE_MARGIN = 24
MIN_RENDER_SCALE = 0.25
MAX_RENDER_SCALE = 5.0
ZOOM_MIN = 0.35
ZOOM_MAX = 4.0
ZOOM_STEP = 1.12
MARK_RADIUS_PT = 18.0
MARK_LINE_WIDTH_PT = 3.2
QUESTION_FONT_SIZE_PT = 34.0
MARK_COLOR_HEX = "#E11D48"
MARK_COLOR_RGB = (225 / 255, 29 / 255, 72 / 255)


@dataclass(frozen=True)
class Mark:
    kind: str
    page_index: int
    x: float
    y: float


def make_root() -> tk.Tk:
    if DND_ENABLED and TkinterDnD is not None:
        try:
            return TkinterDnD.Tk()
        except Exception:
            pass
    return tk.Tk()


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


def icon_candidates() -> list[Path]:
    base = Path(__file__).resolve().parent
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        meipass = Path(getattr(sys, "_MEIPASS", exe_dir))
        return [
            exe_dir / ".." / ".." / "02_assets" / "dake_icon.ico",
            exe_dir / ".." / ".." / ".." / "02_assets" / "dake_icon.ico",
            meipass / "dake_icon.ico",
        ]
    return [
        base / ".." / ".." / "02_assets" / "dake_icon.ico",
        Path("..") / ".." / "02_assets" / "dake_icon.ico",
    ]


def apply_window_icon(root: tk.Tk) -> None:
    for candidate in icon_candidates():
        try:
            icon_path = candidate.resolve()
        except Exception:
            icon_path = candidate
        if not icon_path.exists():
            continue
        try:
            root.iconbitmap(str(icon_path))
            root.iconbitmap(default=str(icon_path))
            return
        except Exception:
            continue


def parse_dropped_files(root: tk.Tk, raw_data: str) -> list[Path]:
    paths: list[Path] = []
    for raw_item in root.tk.splitlist(raw_data):
        value = raw_item.strip().strip("{}")
        if value.startswith("file:"):
            parsed = urllib.parse.urlparse(value)
            value = urllib.parse.unquote(parsed.path)
            if value.startswith("/") and len(value) > 3 and value[2] == ":":
                value = value[1:]
        if value:
            paths.append(Path(value))
    return paths


def is_pdf_path(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() == ".pdf"


def humanize_error(exc: Exception) -> str:
    detail = str(exc).strip().replace("\n", " ")
    return detail or UI_TEXT["status_error"]


def open_folder(path: Path) -> None:
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception:
        pass


def clamp_mark_center(rect: "fitz.Rect", x: float, y: float) -> tuple[float, float]:
    radius = MARK_RADIUS_PT
    if rect.width > radius * 2:
        x = min(max(x, rect.x0 + radius), rect.x1 - radius)
    else:
        x = rect.x0 + rect.width / 2
    if rect.height > radius * 2:
        y = min(max(y, rect.y0 + radius), rect.y1 - radius)
    else:
        y = rect.y0 + rect.height / 2
    return x, y


def draw_pdf_question(page: "fitz.Page", x: float, y: float) -> None:
    font_size = QUESTION_FONT_SIZE_PT
    try:
        font = fitz.Font("helv")
        text_width = font.text_length("?", fontsize=font_size)
    except Exception:
        text_width = font_size * 0.55

    baseline = y + font_size * 0.36
    left = x - text_width / 2
    offsets = ((0.0, 0.0), (0.45, 0.0), (0.0, 0.45))
    for dx, dy in offsets:
        page.insert_text(
            fitz.Point(left + dx, baseline + dy),
            "?",
            fontsize=font_size,
            fontname="helv",
            color=MARK_COLOR_RGB,
            overlay=True,
        )


class AskMarkApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.font_family = detect_font_family(root)
        self.doc: object | None = None
        self.pdf_path: Path | None = None
        self.page_index = 0
        self.page_rect: object | None = None
        self.page_image: tk.PhotoImage | None = None
        self.image_x = PAGE_MARGIN
        self.image_y = PAGE_MARGIN
        self.image_width = 0
        self.image_height = 0
        self.render_scale = 1.0
        self.zoom = 1.0
        self.marks: list[Mark] = []
        self.busy = False
        self.resize_after_id: str | None = None

        self.status_var = tk.StringVar(value=UI_TEXT["status_idle"])
        self.page_var = tk.StringVar(value=UI_TEXT["page_format"].format(current=0, total=0))

        self._configure_root()
        self._build_ui()
        self._setup_bindings()
        self._show_empty_canvas()
        self._update_controls()

    def _configure_root(self) -> None:
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(*WINDOW_MIN_SIZE)
        self.root.configure(bg=THEME["background"])
        apply_window_icon(self.root)

    def _build_ui(self) -> None:
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=1)

        header = tk.Frame(self.root, bg=THEME["background"])
        header.grid(row=0, column=0, sticky="ew", padx=24, pady=(18, 10))
        header.grid_columnconfigure(0, weight=1)

        tk.Label(
            header,
            text=UI_TEXT["brand_series"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 10),
            anchor="w",
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            header,
            text=UI_TEXT["main_title"],
            bg=THEME["background"],
            fg=THEME["text"],
            font=(self.font_family, 22, "bold"),
            anchor="w",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))
        tk.Label(
            header,
            text=UI_TEXT["main_description"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 11),
            anchor="w",
        ).grid(row=2, column=0, sticky="w", pady=(4, 0))

        viewer = tk.Frame(self.root, bg=THEME["border"], bd=0)
        viewer.grid(row=1, column=0, sticky="nsew", padx=24, pady=(0, 10))
        viewer.grid_columnconfigure(0, weight=1)
        viewer.grid_rowconfigure(0, weight=1)

        self.canvas = tk.Canvas(
            viewer,
            bg=THEME["soft"],
            highlightthickness=0,
            bd=0,
            xscrollincrement=16,
            yscrollincrement=16,
        )
        self.canvas.grid(row=0, column=0, sticky="nsew")

        self.v_scroll = tk.Scrollbar(viewer, orient="vertical", command=self.canvas.yview)
        self.h_scroll = tk.Scrollbar(viewer, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=self.v_scroll.set, xscrollcommand=self.h_scroll.set)
        self.v_scroll.grid(row=0, column=1, sticky="ns")
        self.h_scroll.grid(row=1, column=0, sticky="ew")

        controls = tk.Frame(self.root, bg=THEME["background"])
        controls.grid(row=2, column=0, sticky="ew", padx=24, pady=(0, 12))
        controls.grid_columnconfigure(8, weight=1)

        self.select_button = self._make_button(controls, UI_TEXT["button_select_pdf"], self.select_pdf, primary=False)
        self.prev_button = self._make_button(controls, UI_TEXT["button_prev"], self.prev_page, primary=False)
        self.next_button = self._make_button(controls, UI_TEXT["button_next"], self.next_page, primary=False)
        self.undo_button = self._make_button(controls, UI_TEXT["button_undo"], self.undo, primary=False)
        self.save_button = self._make_button(controls, UI_TEXT["button_save"], self.save_pdf, primary=True)

        self.select_button.grid(row=0, column=0, padx=(0, 8), sticky="w")
        self.prev_button.grid(row=0, column=1, padx=(0, 6), sticky="w")
        tk.Label(
            controls,
            textvariable=self.page_var,
            bg=THEME["background"],
            fg=THEME["text"],
            font=(self.font_family, 11, "bold"),
            width=10,
            anchor="center",
        ).grid(row=0, column=2, padx=(0, 6), sticky="w")
        self.next_button.grid(row=0, column=3, padx=(0, 12), sticky="w")
        self.undo_button.grid(row=0, column=4, padx=(0, 8), sticky="w")
        self.save_button.grid(row=0, column=5, padx=(0, 16), sticky="w")
        tk.Label(
            controls,
            textvariable=self.status_var,
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 10),
            anchor="w",
        ).grid(row=0, column=8, sticky="ew")

        footer = tk.Frame(self.root, bg=THEME["background"])
        footer.grid(row=3, column=0, sticky="ew", padx=24, pady=(0, 12))
        footer.grid_columnconfigure(10, weight=1)

        tk.Label(
            footer,
            text=UI_TEXT["footer_left"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 9),
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            footer,
            text=UI_TEXT["footer_separator"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 9),
        ).grid(row=0, column=1, sticky="w")
        self._make_footer_link(footer, "footer_link_1").grid(row=0, column=2, sticky="w")
        tk.Label(
            footer,
            text=UI_TEXT["footer_separator"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 9),
        ).grid(row=0, column=3, sticky="w")
        self._make_footer_link(footer, "footer_link_2").grid(row=0, column=4, sticky="w")
        tk.Label(
            footer,
            text=UI_TEXT["footer_copyright"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 9),
            anchor="e",
        ).grid(row=0, column=10, sticky="e")

    def _make_button(self, parent: tk.Misc, text: str, command: object, primary: bool) -> tk.Button:
        bg = THEME["accent"] if primary else THEME["panel"]
        fg = THEME["accent_text"] if primary else THEME["text"]
        active_bg = THEME["accent_hover"] if primary else THEME["soft"]
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
            padx=16,
            pady=9,
            cursor="hand2",
            font=(self.font_family, 10, "bold"),
            highlightthickness=1,
            highlightbackground=THEME["border"],
            highlightcolor=THEME["border"],
        )
        return button

    def _make_footer_link(self, parent: tk.Misc, key: str) -> tk.Label:
        label = tk.Label(
            parent,
            text=UI_TEXT[key],
            bg=THEME["background"],
            fg=THEME["accent"],
            font=(self.font_family, 9),
            cursor="hand2",
        )
        label.bind("<Button-1>", lambda _event, link_key=key: self._open_link(link_key))
        return label

    def _setup_bindings(self) -> None:
        self.root.bind("<Control-z>", lambda _event: self.undo())
        self.root.bind("<Control-Z>", lambda _event: self.undo())
        self.canvas.bind("<Button-1>", self._on_left_click)
        self.canvas.bind("<Button-3>", self._on_right_click)
        self.canvas.bind("<Button-2>", self._on_right_click)
        self.canvas.bind("<Control-MouseWheel>", self._on_ctrl_mousewheel)
        self.canvas.bind("<Control-Button-4>", lambda _event: self._zoom_at_mouse(1))
        self.canvas.bind("<Control-Button-5>", lambda _event: self._zoom_at_mouse(-1))
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        if DND_ENABLED and DND_FILES is not None:
            for widget in (self.root, self.canvas):
                try:
                    widget.drop_target_register(DND_FILES)
                    widget.dnd_bind("<<Drop>>", self._on_drop)
                except Exception:
                    pass

    def _open_link(self, key: str) -> None:
        url = LINKS.get(key)
        if not url:
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(url)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", url])
            else:
                subprocess.Popen(["xdg-open", url])
        except Exception:
            pass

    def select_pdf(self) -> None:
        if self.busy:
            return
        filename = filedialog.askopenfilename(
            parent=self.root,
            title=UI_TEXT["dialog_open_title"],
            filetypes=[
                (UI_TEXT["dialog_pdf_filter_label"], "*.pdf"),
                (UI_TEXT["dialog_all_filter_label"], "*.*"),
            ],
        )
        if filename:
            self.open_pdf(Path(filename))

    def open_pdf(self, path: Path) -> None:
        if self.busy:
            return
        if fitz is None:
            self._show_error(UI_TEXT["message_dependency_missing"])
            return
        if not is_pdf_path(path):
            self._show_error(UI_TEXT["message_invalid_file"])
            return

        self._set_busy(True)
        self._set_status(UI_TEXT["status_loading"])
        try:
            new_doc = fitz.open(str(path))
            if new_doc.page_count < 1:
                new_doc.close()
                raise ValueError(UI_TEXT["message_no_pages"])
            self._close_document()
            self.doc = new_doc
            self.pdf_path = path
            self.page_index = 0
            self.zoom = 1.0
            self.marks.clear()
            self.render_page()
            self._set_status(UI_TEXT["status_ready"])
        except Exception as exc:
            detail = humanize_error(exc)
            message = f"{UI_TEXT['message_open_failed']}\n\n{detail}"
            messagebox.showerror(UI_TEXT["dialog_error_title"], message, parent=self.root)
            self._set_status(UI_TEXT["status_error"])
            self._show_empty_canvas()
        finally:
            self._set_busy(False)
            self._update_controls()

    def save_pdf(self) -> None:
        if self.busy:
            return
        if fitz is None:
            self._show_error(UI_TEXT["message_dependency_missing"])
            return
        if self.pdf_path is None or self.doc is None:
            self._show_error(UI_TEXT["message_no_pdf"])
            return
        if not self.marks:
            if not messagebox.askyesno(
                UI_TEXT["dialog_confirm_title"],
                UI_TEXT["message_no_marks_confirm"],
                parent=self.root,
            ):
                return

        initial_name = f"{self.pdf_path.stem}{UI_TEXT['save_suffix']}.pdf"
        filename = filedialog.asksaveasfilename(
            parent=self.root,
            title=UI_TEXT["dialog_save_title"],
            initialdir=str(self.pdf_path.parent),
            initialfile=initial_name,
            defaultextension=".pdf",
            filetypes=[(UI_TEXT["dialog_pdf_filter_label"], "*.pdf")],
        )
        if not filename:
            return

        output_path = Path(filename)
        try:
            if output_path.resolve() == self.pdf_path.resolve():
                self._show_error(UI_TEXT["message_same_file"])
                return
        except Exception:
            pass

        self._set_busy(True)
        self._set_status(UI_TEXT["status_processing"])
        self.root.update_idletasks()
        try:
            with fitz.open(str(self.pdf_path)) as output_doc:
                for mark in self.marks:
                    if mark.page_index < 0 or mark.page_index >= output_doc.page_count:
                        continue
                    page = output_doc.load_page(mark.page_index)
                    x, y = clamp_mark_center(page.rect, mark.x, mark.y)
                    if mark.kind == "circle":
                        rect = fitz.Rect(
                            x - MARK_RADIUS_PT,
                            y - MARK_RADIUS_PT,
                            x + MARK_RADIUS_PT,
                            y + MARK_RADIUS_PT,
                        )
                        page.draw_oval(
                            rect,
                            color=MARK_COLOR_RGB,
                            width=MARK_LINE_WIDTH_PT,
                            overlay=True,
                        )
                    elif mark.kind == "question":
                        draw_pdf_question(page, x, y)
                output_doc.save(str(output_path), garbage=4, deflate=True)

            self._set_status(UI_TEXT["status_saved"])
            messagebox.showinfo(
                UI_TEXT["dialog_complete_title"],
                UI_TEXT["message_save_complete"],
                parent=self.root,
            )
            open_folder(output_path.parent)
        except PermissionError:
            self._show_error(UI_TEXT["message_no_folder_access"])
        except Exception as exc:
            detail = humanize_error(exc)
            message = f"{UI_TEXT['message_save_failed']}\n\n{detail}"
            messagebox.showerror(UI_TEXT["dialog_error_title"], message, parent=self.root)
            self._set_status(UI_TEXT["status_error"])
        finally:
            self._set_busy(False)
            self._update_controls()

    def undo(self) -> None:
        if self.busy:
            return
        if not self.marks:
            self._set_status(UI_TEXT["status_no_undo"])
            return
        self.marks.pop()
        self._redraw_mark_overlay()
        self._set_status(UI_TEXT["status_undo"])
        self._update_controls()

    def prev_page(self) -> None:
        if self.doc is None or self.busy or self.page_index <= 0:
            return
        self.page_index -= 1
        self.render_page()
        self._set_status(UI_TEXT["status_ready"])
        self._update_controls()

    def next_page(self) -> None:
        if self.doc is None or self.busy:
            return
        page_count = getattr(self.doc, "page_count", 0)
        if self.page_index >= page_count - 1:
            return
        self.page_index += 1
        self.render_page()
        self._set_status(UI_TEXT["status_ready"])
        self._update_controls()

    def render_page(self) -> None:
        if fitz is None or self.doc is None:
            self._show_empty_canvas()
            return

        page = self.doc.load_page(self.page_index)
        self.page_rect = page.rect
        canvas_width = max(self.canvas.winfo_width(), 480)
        canvas_height = max(self.canvas.winfo_height(), 360)
        available_width = max(canvas_width - PAGE_MARGIN * 2, 240)
        available_height = max(canvas_height - PAGE_MARGIN * 2, 240)
        fit_scale = min(
            available_width / max(page.rect.width, 1.0),
            available_height / max(page.rect.height, 1.0),
        )
        self.render_scale = min(max(fit_scale * self.zoom, MIN_RENDER_SCALE), MAX_RENDER_SCALE)

        matrix = fitz.Matrix(self.render_scale, self.render_scale)
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        encoded = base64.b64encode(pix.tobytes("png")).decode("ascii")
        self.page_image = tk.PhotoImage(data=encoded, format="PNG")
        self.image_width = self.page_image.width()
        self.image_height = self.page_image.height()

        content_width = max(canvas_width, self.image_width + PAGE_MARGIN * 2)
        content_height = max(canvas_height, self.image_height + PAGE_MARGIN * 2)
        self.image_x = max(PAGE_MARGIN, (content_width - self.image_width) // 2)
        self.image_y = PAGE_MARGIN

        self.canvas.delete("all")
        self.canvas.create_rectangle(
            self.image_x + 3,
            self.image_y + 3,
            self.image_x + self.image_width + 3,
            self.image_y + self.image_height + 3,
            fill="#D8DEE8",
            outline="",
        )
        self.canvas.create_image(self.image_x, self.image_y, image=self.page_image, anchor="nw")
        self.canvas.create_rectangle(
            self.image_x,
            self.image_y,
            self.image_x + self.image_width,
            self.image_y + self.image_height,
            outline=THEME["border"],
            width=1,
        )
        self.canvas.configure(scrollregion=(0, 0, content_width, content_height))
        self._redraw_mark_overlay()
        self._update_page_label()

    def _show_empty_canvas(self) -> None:
        self.canvas.delete("all")
        width = max(self.canvas.winfo_width(), 600)
        height = max(self.canvas.winfo_height(), 360)
        self.canvas.configure(scrollregion=(0, 0, width, height))
        self.canvas.create_rectangle(
            20,
            20,
            width - 20,
            height - 20,
            fill=THEME["panel"],
            outline=THEME["border"],
            width=1,
            dash=(6, 5),
        )
        self.canvas.create_text(
            width / 2,
            height / 2 - 22,
            text=UI_TEXT["empty_title"],
            fill=THEME["text"],
            font=(self.font_family, 18, "bold"),
        )
        self.canvas.create_text(
            width / 2,
            height / 2 + 16,
            text=UI_TEXT["empty_subtitle"],
            fill=THEME["muted"],
            font=(self.font_family, 11),
        )
        self.page_var.set(UI_TEXT["page_format"].format(current=0, total=0))

    def _redraw_mark_overlay(self) -> None:
        self.canvas.delete("mark")
        if self.doc is None or self.page_rect is None:
            return
        for mark in self.marks:
            if mark.page_index != self.page_index:
                continue
            x, y = self._page_point_to_canvas(mark.x, mark.y)
            if mark.kind == "circle":
                radius = MARK_RADIUS_PT * self.render_scale
                line_width = max(2, int(MARK_LINE_WIDTH_PT * self.render_scale))
                self.canvas.create_oval(
                    x - radius,
                    y - radius,
                    x + radius,
                    y + radius,
                    outline=MARK_COLOR_HEX,
                    width=line_width,
                    tags=("mark",),
                )
            elif mark.kind == "question":
                font_size = max(16, int(QUESTION_FONT_SIZE_PT * self.render_scale))
                self.canvas.create_text(
                    x,
                    y,
                    text="?",
                    fill=MARK_COLOR_HEX,
                    font=(self.font_family, font_size, "bold"),
                    tags=("mark",),
                )

    def _event_to_page_point(self, event: tk.Event) -> tuple[float, float] | None:
        if self.doc is None or self.page_rect is None or self.image_width <= 0 or self.image_height <= 0:
            return None
        local_x = self.canvas.canvasx(event.x) - self.image_x
        local_y = self.canvas.canvasy(event.y) - self.image_y
        if not (0 <= local_x <= self.image_width and 0 <= local_y <= self.image_height):
            return None
        x = self.page_rect.x0 + local_x / self.render_scale
        y = self.page_rect.y0 + local_y / self.render_scale
        return clamp_mark_center(self.page_rect, x, y)

    def _page_point_to_canvas(self, x: float, y: float) -> tuple[float, float]:
        if self.page_rect is None:
            return self.image_x, self.image_y
        canvas_x = self.image_x + (x - self.page_rect.x0) * self.render_scale
        canvas_y = self.image_y + (y - self.page_rect.y0) * self.render_scale
        return canvas_x, canvas_y

    def _on_left_click(self, event: tk.Event) -> str:
        self._add_mark(event, "circle")
        return "break"

    def _on_right_click(self, event: tk.Event) -> str:
        self._add_mark(event, "question")
        return "break"

    def _add_mark(self, event: tk.Event, kind: str) -> None:
        if self.busy or self.doc is None:
            return
        point = self._event_to_page_point(event)
        if point is None:
            return
        self.marks.append(Mark(kind=kind, page_index=self.page_index, x=point[0], y=point[1]))
        self._redraw_mark_overlay()
        if kind == "circle":
            self._set_status(UI_TEXT["status_circle_added"])
        else:
            self._set_status(UI_TEXT["status_question_added"])
        self._update_controls()

    def _on_ctrl_mousewheel(self, event: tk.Event) -> str:
        direction = 1 if getattr(event, "delta", 0) > 0 else -1
        self._zoom_at_mouse(direction)
        return "break"

    def _zoom_at_mouse(self, direction: int) -> str:
        if self.doc is None or self.busy:
            return "break"
        if direction > 0:
            self.zoom = min(self.zoom * ZOOM_STEP, ZOOM_MAX)
        else:
            self.zoom = max(self.zoom / ZOOM_STEP, ZOOM_MIN)
        self.render_page()
        return "break"

    def _on_canvas_configure(self, _event: tk.Event) -> None:
        if self.resize_after_id is not None:
            try:
                self.root.after_cancel(self.resize_after_id)
            except Exception:
                pass
        self.resize_after_id = self.root.after(120, self._rerender_after_resize)

    def _rerender_after_resize(self) -> None:
        self.resize_after_id = None
        if self.doc is None:
            self._show_empty_canvas()
        else:
            self.render_page()

    def _on_drop(self, event: tk.Event) -> None:
        paths = parse_dropped_files(self.root, getattr(event, "data", ""))
        if not paths:
            return
        self.open_pdf(paths[0])

    def _set_status(self, text: str) -> None:
        self.status_var.set(text)

    def _show_error(self, message: str) -> None:
        messagebox.showerror(UI_TEXT["dialog_error_title"], message, parent=self.root)
        self._set_status(UI_TEXT["status_error"])

    def _set_busy(self, busy: bool) -> None:
        self.busy = busy
        self._update_controls()

    def _update_page_label(self) -> None:
        if self.doc is None:
            self.page_var.set(UI_TEXT["page_format"].format(current=0, total=0))
            return
        page_count = getattr(self.doc, "page_count", 0)
        self.page_var.set(UI_TEXT["page_format"].format(current=self.page_index + 1, total=page_count))

    def _update_controls(self) -> None:
        has_pdf = self.doc is not None and self.pdf_path is not None
        page_count = getattr(self.doc, "page_count", 0) if self.doc is not None else 0
        normal = tk.NORMAL if not self.busy else tk.DISABLED
        self.select_button.configure(state=normal)
        self.save_button.configure(state=tk.NORMAL if has_pdf and not self.busy else tk.DISABLED)
        self.undo_button.configure(state=tk.NORMAL if self.marks and not self.busy else tk.DISABLED)
        self.prev_button.configure(state=tk.NORMAL if has_pdf and not self.busy and self.page_index > 0 else tk.DISABLED)
        self.next_button.configure(
            state=tk.NORMAL if has_pdf and not self.busy and self.page_index < page_count - 1 else tk.DISABLED
        )
        self._update_page_label()

    def _close_document(self) -> None:
        if self.doc is not None:
            try:
                self.doc.close()
            except Exception:
                pass
        self.doc = None
        self.page_rect = None
        self.page_image = None

    def run(self) -> None:
        self.root.mainloop()
        self._close_document()


def main() -> None:
    set_windows_app_id()
    root = make_root()
    app = AskMarkApp(root)
    app.run()


if __name__ == "__main__":
    main()
