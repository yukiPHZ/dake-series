# -*- coding: utf-8 -*-
from __future__ import annotations

import ctypes
import json
import os
import re
import subprocess
import sys
import threading
import webbrowser
from dataclasses import dataclass
from datetime import date
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, ttk

try:
    from reportlab.lib import colors as pdf_colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.pdfgen import canvas as pdf_canvas

    REPORTLAB_AVAILABLE = True
    REPORTLAB_IMPORT_ERROR: Exception | None = None
except Exception as exc:
    A4 = None
    mm = 1
    pdf_colors = None
    pdfmetrics = None
    UnicodeCIDFont = None
    pdf_canvas = None
    REPORTLAB_AVAILABLE = False
    REPORTLAB_IMPORT_ERROR = exc


APP_NAME = "Dake書類送付状"
WINDOW_TITLE = "書類送付状"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "brand_series": "シンプルそれDAKEシリーズ",
    "main_title": "書類送付状を作る",
    "main_description": "相手方と送付内容を入力して、A4縦のPDFを作成します。",
    "section_recipient": "宛先",
    "section_items": "送付内容",
    "section_sender": "作成者情報",
    "section_save": "保存先",
    "label_date": "日付",
    "label_company": "会社名",
    "label_department": "部署名",
    "label_position": "役職",
    "label_name": "氏名",
    "label_honorific": "敬称",
    "label_zip": "郵便番号",
    "label_address": "住所",
    "label_tel": "電話番号",
    "label_fax": "FAX",
    "label_email": "メールアドレス",
    "label_document_name": "書類名",
    "label_copies": "部数",
    "label_note": "備考",
    "button_add_row": "行を追加",
    "button_select_folder": "保存先を選ぶ",
    "button_create_pdf": "PDFを作成",
    "status_idle": "入力してください",
    "status_ready": "準備完了",
    "status_processing": "PDF作成中",
    "status_complete": "PDFを作成しました",
    "status_error": "エラーが発生しました",
    "message_complete_title": "完了",
    "message_complete_body": "書類送付状PDFを作成しました。",
    "message_error_title": "エラー",
    "message_required_body": "宛先または送付内容を確認してください。",
    "message_busy_title": "処理中",
    "message_busy_body": "PDF作成が終わるまでお待ちください。",
    "message_open_folder_body": "OK後に保存先フォルダを開きます。",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_link_1": "戸建買取査定",
    "footer_link_2": "Instagram",
    "footer_separator": " ｜ ",
    "footer_copyright": COPYRIGHT,
    "honorific_company": "御中",
    "honorific_person": "様",
    "date_format": "{year}年{month}月{day}日",
    "date_placeholder": "例：2026年4月26日",
    "folder_dialog_title": "保存先フォルダを選択",
    "error_required_recipient": "宛先の会社名または氏名を入力してください。",
    "error_required_documents": "送付内容の書類名を1行以上入力してください。",
    "error_date_format": "日付は「2026年4月26日」の形式で入力してください。",
    "error_save_folder": "保存先フォルダが見つかりません。保存先を確認してください。",
    "error_reportlab_missing": "PDF作成に必要な reportlab が見つかりません。reportlab をインストールしてから実行してください。",
    "error_file_locked": "同名のPDFが開かれている可能性があります。PDFを閉じてからもう一度お試しください。",
    "error_pdf_failed": "PDF作成に失敗しました。入力内容と保存先を確認してください。",
    "filename_title": "書類送付状",
    "filename_recipient_fallback": "宛先",
    "documents_row_template": "{number}",
    "pdf_title": "書類送付状",
    "pdf_body": "下記の書類を送付いたします。ご査収のほどよろしくお願いいたします。",
    "pdf_table_document_name": "書類名",
    "pdf_table_copies": "部数",
    "pdf_table_note": "備考",
    "pdf_sender_tel": "TEL",
    "pdf_sender_fax": "FAX",
    "pdf_sender_email": "E-mail",
    "pdf_postal_mark": "〒",
    "pdf_closing": "以上",
    "empty_value": "",
    "save_folder_default_label": "Downloads",
}

COLORS = {
    "background": "#F6F7F9",
    "card": "#FFFFFF",
    "text": "#1E2430",
    "muted": "#667085",
    "border": "#E6EAF0",
    "accent": "#2F6FED",
    "accent_hover": "#2458BF",
    "disabled_bg": "#D6DEE8",
    "disabled_fg": "#98A2B3",
    "success": "#12B76A",
    "error": "#D92D20",
    "table_header": "#F2F4F7",
}

LINK_URLS = {
    "footer_link_1": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "footer_link_2": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

CONFIG_NAME = "document_cover_config.json"
WINDOW_APP_ID = "Dake.DocumentCover"
INITIAL_DOCUMENT_ROWS = 3
DATE_SEPARATED_PATTERN = re.compile(r"^\s*(\d{4})\D+(\d{1,2})\D+(\d{1,2})\D*\s*$")
DATE_DIGIT_PATTERN = re.compile(r"^\s*(\d{4})(\d{2})(\d{2})\s*$")
FILENAME_UNSAFE_PATTERN = re.compile(r'[\\/:*?"<>|]+')


@dataclass(frozen=True)
class DocumentRow:
    name: str
    copies: str
    note: str


@dataclass(frozen=True)
class CoverData:
    output_date: date
    recipient: dict[str, str]
    documents: list[DocumentRow]
    sender: dict[str, str]
    save_folder: Path


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def default_downloads() -> Path:
    downloads = Path.home() / "Downloads"
    return downloads if downloads.exists() else Path.home()


def config_path() -> Path:
    return app_dir() / CONFIG_NAME


def load_config() -> dict:
    try:
        with config_path().open("r", encoding="utf-8") as file:
            data = json.load(file)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_config(data: dict) -> None:
    try:
        with config_path().open("w", encoding="utf-8") as file:
            json.dump(data, file, ensure_ascii=False, indent=2)
    except Exception:
        pass


def format_japanese_date(value: date) -> str:
    return UI_TEXT["date_format"].format(year=value.year, month=value.month, day=value.day)


def parse_date_input(value: str) -> date:
    text = value.strip()
    match = DATE_SEPARATED_PATTERN.match(text) or DATE_DIGIT_PATTERN.match(text)
    if not match:
        raise ValueError(UI_TEXT["error_date_format"])
    year, month, day = (int(part) for part in match.groups())
    try:
        return date(year, month, day)
    except ValueError as exc:
        raise ValueError(UI_TEXT["error_date_format"]) from exc


def sanitize_filename_part(value: str) -> str:
    text = FILENAME_UNSAFE_PATTERN.sub("_", value.strip())
    text = re.sub(r"\s+", "_", text)
    return text.strip("._ ") or UI_TEXT["filename_recipient_fallback"]


def unique_pdf_path(folder: Path, filename: str) -> Path:
    candidate = folder / filename
    if not candidate.exists():
        return candidate
    stem = candidate.stem
    suffix = candidate.suffix
    index = 2
    while True:
        numbered = folder / f"{stem}_{index}{suffix}"
        if not numbered.exists():
            return numbered
        index += 1


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


def set_windows_app_id() -> None:
    if not sys.platform.startswith("win"):
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(WINDOW_APP_ID)
    except Exception:
        pass


def find_icon_path() -> Path | None:
    base = app_dir()
    candidates = [
        base / ".." / ".." / "02_assets" / "dake_icon.ico",
        base / ".." / ".." / ".." / "02_assets" / "dake_icon.ico",
        Path(__file__).resolve().parent / ".." / ".." / "02_assets" / "dake_icon.ico",
    ]
    for candidate in candidates:
        resolved = candidate.resolve()
        if resolved.exists():
            return resolved
    return None


def register_pdf_fonts() -> tuple[str, str]:
    if not REPORTLAB_AVAILABLE or pdfmetrics is None or UnicodeCIDFont is None:
        raise RuntimeError(UI_TEXT["error_reportlab_missing"])
    gothic = "HeiseiKakuGo-W5"
    mincho = "HeiseiMin-W3"
    for font_name in (gothic, mincho):
        try:
            pdfmetrics.getFont(font_name)
        except KeyError:
            pdfmetrics.registerFont(UnicodeCIDFont(font_name))
    return gothic, mincho


def wrap_pdf_text(text: str, max_width: float, font_name: str, font_size: int) -> list[str]:
    if pdfmetrics is None:
        return [text]
    lines: list[str] = []
    for raw_line in text.splitlines() or [text]:
        current = ""
        for char in raw_line:
            trial = current + char
            if current and pdfmetrics.stringWidth(trial, font_name, font_size) > max_width:
                lines.append(current)
                current = char
            else:
                current = trial
        lines.append(current)
    return lines or [UI_TEXT["empty_value"]]


def draw_wrapped(
    pdf: object,
    text: str,
    x: float,
    y: float,
    max_width: float,
    font_name: str,
    font_size: int,
    leading: float,
    max_lines: int | None = None,
) -> float:
    lines = wrap_pdf_text(text, max_width, font_name, font_size)
    if max_lines is not None:
        lines = lines[:max_lines]
    pdf.setFont(font_name, font_size)
    for line in lines:
        pdf.drawString(x, y, line)
        y -= leading
    return y


def draw_table_cell(
    pdf: object,
    text: str,
    x: float,
    y: float,
    width: float,
    height: float,
    font_name: str,
    font_size: int,
    align_center: bool = False,
) -> None:
    lines = wrap_pdf_text(text, width - 8, font_name, font_size)[:2]
    start_y = y + height - 12
    pdf.setFont(font_name, font_size)
    for line in lines:
        if align_center:
            pdf.drawCentredString(x + width / 2, start_y, line)
        else:
            pdf.drawString(x + 4, start_y, line)
        start_y -= 10


def build_recipient_lines(recipient: dict[str, str]) -> list[str]:
    company = recipient["company"]
    department = recipient["department"]
    position = recipient["position"]
    name = recipient["name"]
    honorific = recipient["honorific"]
    lines: list[str] = []
    if company:
        lines.append(company if (department or position or name) else f"{company} {honorific}")
    if department:
        lines.append(department if (position or name) else f"{department} {honorific}")
    if position and name:
        lines.append(f"{position}　{name} {honorific}")
    elif position:
        lines.append(f"{position} {honorific}")
    elif name:
        lines.append(f"{name} {honorific}")
    return lines


def build_sender_lines(sender: dict[str, str]) -> list[str]:
    lines: list[str] = []
    for key in ("company", "department", "name"):
        if sender[key]:
            lines.append(sender[key])
    if sender["zip"]:
        lines.append(f"{UI_TEXT['pdf_postal_mark']} {sender['zip']}")
    if sender["address"]:
        lines.append(sender["address"])
    contact_parts = []
    if sender["tel"]:
        contact_parts.append(f"{UI_TEXT['pdf_sender_tel']} {sender['tel']}")
    if sender["fax"]:
        contact_parts.append(f"{UI_TEXT['pdf_sender_fax']} {sender['fax']}")
    if contact_parts:
        lines.append(UI_TEXT["footer_separator"].join(contact_parts))
    if sender["email"]:
        lines.append(f"{UI_TEXT['pdf_sender_email']} {sender['email']}")
    return lines


def create_pdf(data: CoverData) -> Path:
    if not REPORTLAB_AVAILABLE or A4 is None or pdf_canvas is None or pdf_colors is None:
        detail = f" ({REPORTLAB_IMPORT_ERROR})" if REPORTLAB_IMPORT_ERROR else ""
        raise RuntimeError(UI_TEXT["error_reportlab_missing"] + detail)

    gothic, mincho = register_pdf_fonts()
    recipient_name = data.recipient["company"] or data.recipient["name"]
    filename_recipient = sanitize_filename_part(recipient_name)
    filename = f"{data.output_date:%Y%m%d}_{UI_TEXT['filename_title']}_{filename_recipient}.pdf"
    output_path = unique_pdf_path(data.save_folder, filename)

    width, height = A4
    margin_x = 22 * mm
    top_y = height - 24 * mm
    table_width = width - margin_x * 2
    col_widths = [78 * mm, 22 * mm, table_width - 100 * mm]
    row_height = 25 * mm / 2

    try:
        pdf = pdf_canvas.Canvas(str(output_path), pagesize=A4)
        pdf.setTitle(UI_TEXT["pdf_title"])
        pdf.setAuthor(APP_NAME)
        pdf.setFillColor(pdf_colors.HexColor(COLORS["text"]))

        pdf.setFont(gothic, 10)
        pdf.drawRightString(width - margin_x, top_y, format_japanese_date(data.output_date))

        recipient_y = top_y - 34
        for line in build_recipient_lines(data.recipient):
            recipient_y = draw_wrapped(pdf, line, margin_x, recipient_y, 72 * mm, gothic, 11, 16, max_lines=2)

        sender_x = width - margin_x - 70 * mm
        sender_y = top_y - 34
        for line in build_sender_lines(data.sender):
            sender_y = draw_wrapped(pdf, line, sender_x, sender_y, 70 * mm, gothic, 9, 13, max_lines=2)

        title_y = top_y - 132
        pdf.setFont(mincho, 21)
        pdf.drawCentredString(width / 2, title_y, UI_TEXT["pdf_title"])

        body_y = title_y - 42
        pdf.setFont(gothic, 11)
        pdf.drawString(margin_x, body_y, UI_TEXT["pdf_body"])

        table_x = margin_x
        table_top = body_y - 38
        headers = [
            UI_TEXT["pdf_table_document_name"],
            UI_TEXT["pdf_table_copies"],
            UI_TEXT["pdf_table_note"],
        ]

        pdf.setLineWidth(0.7)
        pdf.setStrokeColor(pdf_colors.HexColor(COLORS["border"]))
        pdf.setFillColor(pdf_colors.HexColor(COLORS["table_header"]))
        pdf.rect(table_x, table_top - row_height, table_width, row_height, stroke=1, fill=1)

        current_x = table_x
        pdf.setFillColor(pdf_colors.HexColor(COLORS["text"]))
        for index, header in enumerate(headers):
            pdf.line(current_x, table_top, current_x, table_top - row_height)
            draw_table_cell(pdf, header, current_x, table_top - row_height, col_widths[index], row_height, gothic, 10, True)
            current_x += col_widths[index]
        pdf.line(current_x, table_top, current_x, table_top - row_height)

        row_y = table_top - row_height
        for document in data.documents:
            row_y -= row_height
            pdf.setFillColor(pdf_colors.white)
            pdf.rect(table_x, row_y, table_width, row_height, stroke=1, fill=1)
            pdf.setFillColor(pdf_colors.HexColor(COLORS["text"]))
            current_x = table_x
            values = [document.name, document.copies, document.note]
            for index, value in enumerate(values):
                pdf.line(current_x, row_y + row_height, current_x, row_y)
                draw_table_cell(pdf, value, current_x, row_y, col_widths[index], row_height, gothic, 10, index == 1)
                current_x += col_widths[index]
            pdf.line(current_x, row_y + row_height, current_x, row_y)

        closing_y = max(row_y - 30, 40 * mm)
        pdf.setFont(gothic, 10)
        pdf.drawRightString(width - margin_x, closing_y, UI_TEXT["pdf_closing"])

        pdf.showPage()
        pdf.save()
    except PermissionError as exc:
        raise PermissionError(UI_TEXT["error_file_locked"]) from exc
    except OSError as exc:
        raise OSError(UI_TEXT["error_pdf_failed"]) from exc

    return output_path


class DocumentCoverApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.configure(bg=COLORS["background"])
        self.root.minsize(1080, 720)

        self.config = load_config()
        self.font_family = self.choose_font_family()
        self.fonts = {
            "title": (self.font_family, 22, "bold"),
            "description": (self.font_family, 11),
            "section": (self.font_family, 13, "bold"),
            "label": (self.font_family, 10, "bold"),
            "body": (self.font_family, 10),
            "small": (self.font_family, 9),
            "button": (self.font_family, 11, "bold"),
            "footer": (self.font_family, 9),
        }

        self.date_var = tk.StringVar(value=format_japanese_date(date.today()))
        self.recipient_vars = {
            "company": tk.StringVar(),
            "department": tk.StringVar(),
            "position": tk.StringVar(),
            "name": tk.StringVar(),
            "honorific": tk.StringVar(value=UI_TEXT["honorific_company"]),
        }
        self.sender_vars = {
            "company": tk.StringVar(),
            "department": tk.StringVar(),
            "name": tk.StringVar(),
            "zip": tk.StringVar(),
            "address": tk.StringVar(),
            "tel": tk.StringVar(),
            "fax": tk.StringVar(),
            "email": tk.StringVar(),
        }
        self.document_rows: list[dict[str, tk.StringVar]] = []
        self.items_card: tk.Frame | None = None
        self.add_row_button: tk.Button | None = None
        self.save_folder_var = tk.StringVar(value=self.initial_save_folder())
        self.status_var = tk.StringVar(value=UI_TEXT["status_idle"])
        self.create_button: tk.Button | None = None
        self.status_label: tk.Label | None = None
        self.is_processing = False

        self.load_sender_config()
        self.configure_style()
        self.apply_window_icon()
        self.build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.handle_close)

    def choose_font_family(self) -> str:
        available = set(tkfont.families(self.root))
        for candidate in ("BIZ UDPGothic", "BIZ UDPゴシック", "Yu Gothic UI", "Meiryo"):
            if candidate in available:
                return candidate
        return "TkDefaultFont"

    def configure_style(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(
            "Dake.TCombobox",
            fieldbackground=COLORS["card"],
            background=COLORS["card"],
            foreground=COLORS["text"],
            bordercolor=COLORS["border"],
            arrowcolor=COLORS["text"],
            padding=4,
        )

    def apply_window_icon(self) -> None:
        icon_path = find_icon_path()
        if icon_path is None:
            return
        try:
            self.root.iconbitmap(str(icon_path))
        except tk.TclError:
            pass

    def initial_save_folder(self) -> str:
        saved = self.config.get("last_save_folder")
        if isinstance(saved, str) and Path(saved).exists():
            return str(Path(saved))
        return str(default_downloads())

    def load_sender_config(self) -> None:
        sender = self.config.get("sender")
        if not isinstance(sender, dict):
            return
        for key, variable in self.sender_vars.items():
            value = sender.get(key)
            if isinstance(value, str):
                variable.set(value)

    def build_ui(self) -> None:
        outer = tk.Frame(self.root, bg=COLORS["background"])
        outer.pack(fill="both", expand=True, padx=26, pady=20)
        outer.grid_columnconfigure(0, weight=1)
        outer.grid_rowconfigure(1, weight=1)

        header = tk.Frame(outer, bg=COLORS["background"])
        header.grid(row=0, column=0, sticky="ew", pady=(0, 16))
        tk.Label(
            header,
            text=UI_TEXT["main_title"],
            font=self.fonts["title"],
            fg=COLORS["text"],
            bg=COLORS["background"],
        ).pack(anchor="w")
        tk.Label(
            header,
            text=UI_TEXT["main_description"],
            font=self.fonts["description"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
        ).pack(anchor="w", pady=(5, 0))

        content = tk.Frame(outer, bg=COLORS["background"])
        content.grid(row=1, column=0, sticky="nsew")
        for column in range(3):
            content.grid_columnconfigure(column, weight=1, uniform="content")
        content.grid_rowconfigure(0, weight=1)

        recipient_column = tk.Frame(content, bg=COLORS["background"])
        items_column = tk.Frame(content, bg=COLORS["background"])
        sender_column = tk.Frame(content, bg=COLORS["background"])
        recipient_column.grid(row=0, column=0, sticky="nsew", padx=(0, 9))
        items_column.grid(row=0, column=1, sticky="nsew", padx=9)
        sender_column.grid(row=0, column=2, sticky="nsew", padx=(9, 0))

        self.build_recipient_card(recipient_column)
        self.build_items_card(items_column)
        self.build_sender_card(sender_column)

        self.build_bottom_area(outer)

    def card(self, parent: tk.Widget, title: str) -> tk.Frame:
        frame = tk.Frame(
            parent,
            bg=COLORS["card"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
            bd=0,
        )
        frame.pack(fill="x", pady=(0, 14))
        inner = tk.Frame(frame, bg=COLORS["card"])
        inner.pack(fill="both", expand=True, padx=18, pady=16)
        tk.Label(
            inner,
            text=title,
            font=self.fonts["section"],
            fg=COLORS["text"],
            bg=COLORS["card"],
        ).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 12))
        inner.grid_columnconfigure(1, weight=1)
        inner.grid_columnconfigure(3, weight=1)
        return inner

    def build_recipient_card(self, parent: tk.Widget) -> None:
        card = self.card(parent, UI_TEXT["section_recipient"])
        self.add_labeled_entry(card, 1, 0, UI_TEXT["label_date"], self.date_var)
        self.add_labeled_entry(card, 2, 0, UI_TEXT["label_company"], self.recipient_vars["company"])
        self.add_labeled_entry(card, 3, 0, UI_TEXT["label_department"], self.recipient_vars["department"])
        self.add_labeled_entry(card, 4, 0, UI_TEXT["label_position"], self.recipient_vars["position"])
        self.add_labeled_entry(card, 5, 0, UI_TEXT["label_name"], self.recipient_vars["name"])

        tk.Label(
            card,
            text=UI_TEXT["label_honorific"],
            font=self.fonts["label"],
            fg=COLORS["text"],
            bg=COLORS["card"],
        ).grid(row=6, column=0, sticky="w", pady=(5, 5), padx=(0, 10))
        honorific_box = ttk.Combobox(
            card,
            textvariable=self.recipient_vars["honorific"],
            values=(UI_TEXT["honorific_company"], UI_TEXT["honorific_person"]),
            state="readonly",
            width=10,
            style="Dake.TCombobox",
            font=self.fonts["body"],
        )
        honorific_box.grid(row=6, column=1, sticky="w", pady=(5, 5))

    def build_items_card(self, parent: tk.Widget) -> None:
        card = self.card(parent, UI_TEXT["section_items"])
        self.items_card = card
        headers = [
            UI_TEXT["label_document_name"],
            UI_TEXT["label_copies"],
            UI_TEXT["label_note"],
        ]
        widths = [18, 6, 18]
        for column, header in enumerate(headers):
            tk.Label(
                card,
                text=header,
                font=self.fonts["label"],
                fg=COLORS["muted"],
                bg=COLORS["card"],
            ).grid(row=1, column=column + 1, sticky="w", padx=(0, 8), pady=(0, 6))
        for _ in range(INITIAL_DOCUMENT_ROWS):
            self.add_document_row()

        self.add_row_button = tk.Button(
            card,
            text=UI_TEXT["button_add_row"],
            command=self.add_document_row_from_ui,
            font=self.fonts["body"],
            fg=COLORS["text"],
            bg=COLORS["card"],
            activebackground=COLORS["background"],
            activeforeground=COLORS["text"],
            relief="solid",
            bd=1,
            highlightthickness=0,
            padx=12,
            pady=6,
            cursor="hand2",
        )
        self.place_add_row_button()
        for column, weight in ((1, 3), (2, 1), (3, 3)):
            card.grid_columnconfigure(column, weight=weight)

    def build_sender_card(self, parent: tk.Widget) -> None:
        card = self.card(parent, UI_TEXT["section_sender"])
        fields = [
            ("company", "label_company"),
            ("department", "label_department"),
            ("name", "label_name"),
            ("zip", "label_zip"),
            ("address", "label_address"),
            ("tel", "label_tel"),
            ("fax", "label_fax"),
            ("email", "label_email"),
        ]
        for row_index, (key, label_key) in enumerate(fields, start=1):
            self.add_labeled_entry(card, row_index, 0, UI_TEXT[label_key], self.sender_vars[key])

    def build_bottom_area(self, outer: tk.Frame) -> None:
        bottom = tk.Frame(outer, bg=COLORS["background"])
        bottom.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        bottom.grid_columnconfigure(0, weight=1)

        save_card = tk.Frame(
            bottom,
            bg=COLORS["card"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
            bd=0,
        )
        save_card.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        save_card.grid_columnconfigure(1, weight=1)

        tk.Label(
            save_card,
            text=UI_TEXT["section_save"],
            font=self.fonts["label"],
            fg=COLORS["text"],
            bg=COLORS["card"],
        ).grid(row=0, column=0, sticky="w", padx=(16, 10), pady=14)

        folder_entry = self.make_entry(save_card, self.save_folder_var)
        folder_entry.configure(state="readonly")
        folder_entry.grid(row=0, column=1, sticky="ew", pady=14)

        select_button = tk.Button(
            save_card,
            text=UI_TEXT["button_select_folder"],
            command=self.select_save_folder,
            font=self.fonts["body"],
            fg=COLORS["text"],
            bg=COLORS["card"],
            activebackground=COLORS["background"],
            activeforeground=COLORS["text"],
            relief="solid",
            bd=1,
            highlightthickness=0,
            padx=14,
            pady=7,
            cursor="hand2",
        )
        select_button.grid(row=0, column=2, padx=12, pady=14)

        self.create_button = tk.Button(
            save_card,
            text=UI_TEXT["button_create_pdf"],
            command=self.handle_create_pdf,
            font=self.fonts["button"],
            fg="#FFFFFF",
            bg=COLORS["accent"],
            activebackground=COLORS["accent_hover"],
            activeforeground="#FFFFFF",
            disabledforeground=COLORS["disabled_fg"],
            relief="flat",
            bd=0,
            highlightthickness=0,
            padx=24,
            pady=9,
            cursor="hand2",
        )
        self.create_button.grid(row=0, column=3, padx=(0, 16), pady=14)

        lower = tk.Frame(bottom, bg=COLORS["background"])
        lower.grid(row=1, column=0, sticky="ew")
        lower.grid_columnconfigure(0, weight=1)

        self.status_label = tk.Label(
            lower,
            textvariable=self.status_var,
            font=self.fonts["small"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
        )
        self.status_label.grid(row=0, column=0, sticky="w")

        footer = tk.Frame(lower, bg=COLORS["background"])
        footer.grid(row=0, column=1, sticky="e")
        self.add_footer_label(footer, UI_TEXT["footer_left"])
        self.add_footer_label(footer, UI_TEXT["footer_separator"])
        self.add_footer_link(footer, "footer_link_1")
        self.add_footer_label(footer, UI_TEXT["footer_separator"])
        self.add_footer_link(footer, "footer_link_2")
        self.add_footer_label(footer, UI_TEXT["footer_separator"])
        self.add_footer_label(footer, UI_TEXT["footer_copyright"])

    def add_labeled_entry(
        self,
        parent: tk.Widget,
        row: int,
        column: int,
        label: str,
        variable: tk.StringVar,
    ) -> tk.Entry:
        tk.Label(
            parent,
            text=label,
            font=self.fonts["label"],
            fg=COLORS["text"],
            bg=COLORS["card"],
        ).grid(row=row, column=column, sticky="w", pady=(5, 5), padx=(0, 10))
        entry = self.make_entry(parent, variable)
        entry.grid(row=row, column=column + 1, sticky="ew", pady=(5, 5))
        return entry

    def make_entry(self, parent: tk.Widget, variable: tk.StringVar, width: int = 24) -> tk.Entry:
        return tk.Entry(
            parent,
            textvariable=variable,
            font=self.fonts["body"],
            fg=COLORS["text"],
            bg="#FFFFFF",
            readonlybackground="#FFFFFF",
            relief="solid",
            bd=1,
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            highlightcolor=COLORS["accent"],
            insertbackground=COLORS["text"],
            width=width,
        )

    def add_document_row(self) -> None:
        if self.items_card is None:
            return
        index = len(self.document_rows) + 1
        row_vars = {
            "name": tk.StringVar(),
            "copies": tk.StringVar(),
            "note": tk.StringVar(),
        }
        self.document_rows.append(row_vars)
        tk.Label(
            self.items_card,
            text=UI_TEXT["documents_row_template"].format(number=index),
            font=self.fonts["small"],
            fg=COLORS["muted"],
            bg=COLORS["card"],
        ).grid(row=index + 1, column=0, sticky="e", padx=(0, 8), pady=4)
        widths = {"name": 18, "copies": 6, "note": 18}
        for column, key in enumerate(("name", "copies", "note"), start=1):
            entry = self.make_entry(self.items_card, row_vars[key], width=widths[key])
            entry.grid(row=index + 1, column=column, sticky="ew", padx=(0, 8), pady=4)
        self.place_add_row_button()

    def add_document_row_from_ui(self) -> None:
        self.add_document_row()
        self.set_status("status_ready")

    def place_add_row_button(self) -> None:
        if self.add_row_button is None:
            return
        self.add_row_button.grid_forget()
        self.add_row_button.grid(row=len(self.document_rows) + 2, column=1, sticky="w", pady=(10, 0))

    def add_footer_label(self, parent: tk.Widget, text: str) -> None:
        tk.Label(
            parent,
            text=text,
            font=self.fonts["footer"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
        ).pack(side="left")

    def add_footer_link(self, parent: tk.Widget, key: str) -> None:
        label = tk.Label(
            parent,
            text=UI_TEXT[key],
            font=self.fonts["footer"],
            fg=COLORS["accent"],
            bg=COLORS["background"],
            cursor="hand2",
        )
        label.pack(side="left")
        label.bind("<Button-1>", lambda _event, url=LINK_URLS[key]: webbrowser.open(url))

    def select_save_folder(self) -> None:
        selected = filedialog.askdirectory(
            title=UI_TEXT["folder_dialog_title"],
            initialdir=self.save_folder_var.get() or str(default_downloads()),
        )
        if not selected:
            return
        self.save_folder_var.set(selected)
        self.save_current_config()
        self.set_status("status_ready")

    def set_processing(self, processing: bool) -> None:
        self.is_processing = processing
        if self.create_button is None:
            return
        if processing:
            self.create_button.configure(state="disabled", bg=COLORS["disabled_bg"])
        else:
            self.create_button.configure(state="normal", bg=COLORS["accent"])
        self.root.update_idletasks()

    def set_status(self, key: str, is_error: bool = False) -> None:
        self.status_var.set(UI_TEXT[key])
        if self.status_label is not None:
            self.status_label.configure(fg=COLORS["error"] if is_error else COLORS["muted"])

    def collect_documents(self) -> list[DocumentRow]:
        rows: list[DocumentRow] = []
        for row_vars in self.document_rows:
            document = DocumentRow(
                name=row_vars["name"].get().strip(),
                copies=row_vars["copies"].get().strip(),
                note=row_vars["note"].get().strip(),
            )
            if document.name:
                rows.append(document)
        return rows

    def collect_data(self) -> CoverData:
        output_date = parse_date_input(self.date_var.get())
        recipient = {key: variable.get().strip() for key, variable in self.recipient_vars.items()}
        sender = {key: variable.get().strip() for key, variable in self.sender_vars.items()}
        documents = self.collect_documents()
        save_folder = Path(self.save_folder_var.get()).expanduser()

        errors = []
        if not (recipient["company"] or recipient["name"]):
            errors.append(UI_TEXT["error_required_recipient"])
        if not documents:
            errors.append(UI_TEXT["error_required_documents"])
        if not save_folder.exists() or not save_folder.is_dir():
            errors.append(UI_TEXT["error_save_folder"])
        if errors:
            raise ValueError("\n".join(errors))

        return CoverData(
            output_date=output_date,
            recipient=recipient,
            documents=documents,
            sender=sender,
            save_folder=save_folder,
        )

    def save_current_config(self) -> None:
        save_config(
            {
                "sender": {key: variable.get().strip() for key, variable in self.sender_vars.items()},
                "last_save_folder": self.save_folder_var.get(),
            }
        )

    def handle_create_pdf(self) -> None:
        if self.is_processing:
            messagebox.showinfo(UI_TEXT["message_busy_title"], UI_TEXT["message_busy_body"])
            return
        try:
            data = self.collect_data()
        except Exception as exc:
            self.set_status("status_error", is_error=True)
            messagebox.showerror(UI_TEXT["message_error_title"], str(exc) or UI_TEXT["message_required_body"])
            return

        self.save_current_config()
        self.set_processing(True)
        self.set_status("status_processing")
        worker = threading.Thread(target=self.create_pdf_worker, args=(data,), daemon=True)
        worker.start()

    def create_pdf_worker(self, data: CoverData) -> None:
        try:
            output_path = create_pdf(data)
        except Exception as exc:
            self.root.after(0, lambda error=exc: self.finish_create_pdf(None, error))
            return
        self.root.after(0, lambda path=output_path: self.finish_create_pdf(path, None))

    def finish_create_pdf(self, output_path: Path | None, error: Exception | None) -> None:
        self.set_processing(False)
        if error is not None:
            self.set_status("status_error", is_error=True)
            messagebox.showerror(UI_TEXT["message_error_title"], str(error) or UI_TEXT["error_pdf_failed"])
            return

        self.set_status("status_complete")
        message = f"{UI_TEXT['message_complete_body']}\n{UI_TEXT['message_open_folder_body']}"
        messagebox.showinfo(UI_TEXT["message_complete_title"], message)
        if output_path is not None:
            open_folder(output_path.parent)

    def handle_close(self) -> None:
        self.save_current_config()
        self.root.destroy()


def main() -> None:
    set_windows_app_id()
    root = tk.Tk()
    DocumentCoverApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
