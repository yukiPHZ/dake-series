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


APP_NAME = "DakeFAX送付状"
WINDOW_TITLE = "FAX送付状"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "brand_series": "シンプルそれDAKEシリーズ",
    "main_title": "FAX送付状を作る",
    "main_description": "送信先と内容を入力して、A4縦のFAX送付状PDFを作成します。",
    "section_recipient": "送信先",
    "section_fax_info": "FAX情報",
    "section_items": "送信内容",
    "section_message": "メッセージ",
    "section_sender": "送信者情報",
    "section_save": "保存先",
    "label_date": "日付",
    "label_company": "会社名",
    "label_department": "部署名",
    "label_position": "役職",
    "label_name": "氏名",
    "label_honorific": "敬称",
    "label_fax_number": "FAX番号",
    "label_tel": "電話番号",
    "label_subject": "件名",
    "label_total_pages": "送信枚数",
    "label_send_date": "送信日",
    "label_zip": "郵便番号",
    "label_address": "住所",
    "label_email": "メールアドレス",
    "label_item_name": "送付書類名",
    "label_item_pages": "枚数",
    "label_note": "備考",
    "button_add_row": "行を追加",
    "button_select_folder": "保存先を選ぶ",
    "button_create_pdf": "FAX送付状PDFを作成",
    "status_idle": "入力してください",
    "status_ready": "準備完了",
    "status_processing": "PDF作成中",
    "status_complete": "FAX送付状PDFを作成しました",
    "status_error": "入力内容を確認してください",
    "message_complete_title": "完了",
    "message_complete_body": "FAX送付状PDFを作成しました。",
    "message_error_title": "確認してください",
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
    "default_message": "下記のとおりFAXを送信いたします。ご確認のほどよろしくお願いいたします。",
    "date_format": "{year}年{month}月{day}日",
    "date_placeholder": "例：2026年4月27日",
    "folder_dialog_title": "保存先フォルダを選択",
    "error_required_recipient": "送信先の会社名または氏名を入力してください。",
    "error_required_items": "送信内容の送付書類名を1行以上入力してください。",
    "error_required_sender": "送信者の氏名を入力してください。",
    "error_date_format": "日付は「2026年4月27日」の形式で入力してください。",
    "error_send_date_format": "送信日は「2026年4月27日」の形式で入力してください。",
    "error_save_folder": "保存先フォルダが見つかりません。保存先を確認してください。",
    "error_reportlab_missing": "PDF作成に必要な reportlab が見つかりません。reportlab をインストールしてから実行してください。",
    "error_file_locked": "同名のPDFが開かれている可能性があります。PDFを閉じてからもう一度お試しください。",
    "error_pdf_failed": "PDF作成に失敗しました。入力内容と保存先を確認してください。",
    "filename_title": "FAX送付状",
    "filename_recipient_fallback": "送信先",
    "items_row_template": "{number}",
    "pdf_title": "FAX送付状",
    "pdf_recipient_heading": "送信先",
    "pdf_sender_heading": "送信者",
    "pdf_fax_info_heading": "FAX情報",
    "pdf_subject": "件名",
    "pdf_total_pages": "送信枚数",
    "pdf_send_date": "送信日",
    "pdf_table_item_name": "送付書類名",
    "pdf_table_pages": "枚数",
    "pdf_table_note": "備考",
    "pdf_tel": "TEL",
    "pdf_fax": "FAX",
    "pdf_email": "E-mail",
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

CONFIG_NAME = "fax_cover_config.json"
WINDOW_APP_ID = "Dake.FAXCover"
INITIAL_ITEM_ROWS = 3
DATE_SEPARATED_PATTERN = re.compile(r"^\s*(\d{4})\D+(\d{1,2})\D+(\d{1,2})\D*\s*$")
DATE_DIGIT_PATTERN = re.compile(r"^\s*(\d{4})(\d{2})(\d{2})\s*$")
FILENAME_UNSAFE_PATTERN = re.compile(r'[\\/:*?"<>|]+')


@dataclass(frozen=True)
class FaxItemRow:
    name: str
    pages: str
    note: str


@dataclass(frozen=True)
class FaxCoverData:
    output_date: date
    send_date: date
    recipient: dict[str, str]
    fax_info: dict[str, str]
    items: list[FaxItemRow]
    message: str
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


def parse_date_input(value: str, error_text: str) -> date:
    text = value.strip()
    match = DATE_SEPARATED_PATTERN.match(text) or DATE_DIGIT_PATTERN.match(text)
    if not match:
        raise ValueError(error_text)
    year, month, day = (int(part) for part in match.groups())
    try:
        return date(year, month, day)
    except ValueError as exc:
        raise ValueError(error_text) from exc


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
    canvas: object,
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
    canvas.setFont(font_name, font_size)
    for line in lines:
        canvas.drawString(x, y, line)
        y -= leading
    return y


def draw_label_value(
    canvas: object,
    label: str,
    value: str,
    x: float,
    y: float,
    label_width: float,
    max_width: float,
    font_name: str,
    label_font: str,
    font_size: int = 9,
) -> float:
    canvas.setFont(label_font, font_size)
    canvas.drawString(x, y, label)
    return draw_wrapped(canvas, value, x + label_width, y, max_width - label_width, font_name, font_size, 12, max_lines=2)


def draw_table_cell(
    canvas: object,
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
    canvas.setFont(font_name, font_size)
    for line in lines:
        if align_center:
            canvas.drawCentredString(x + width / 2, start_y, line)
        else:
            canvas.drawString(x + 4, start_y, line)
        start_y -= 11


def build_recipient_lines(recipient: dict[str, str]) -> list[str]:
    lines: list[str] = []
    company = recipient["company"]
    department = recipient["department"]
    position = recipient["position"]
    name = recipient["name"]
    honorific = recipient["honorific"]
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
    if recipient["fax"]:
        lines.append(f"{UI_TEXT['pdf_fax']} {recipient['fax']}")
    if recipient["tel"]:
        lines.append(f"{UI_TEXT['pdf_tel']} {recipient['tel']}")
    return lines


def build_sender_lines(sender: dict[str, str]) -> list[str]:
    lines: list[str] = []
    for key in ("company", "department", "name"):
        if sender[key]:
            lines.append(sender[key])
    if sender["postal_code"]:
        lines.append(f"{UI_TEXT['pdf_postal_mark']} {sender['postal_code']}")
    if sender["address"]:
        lines.append(sender["address"])
    if sender["tel"]:
        lines.append(f"{UI_TEXT['pdf_tel']} {sender['tel']}")
    if sender["fax"]:
        lines.append(f"{UI_TEXT['pdf_fax']} {sender['fax']}")
    if sender["email"]:
        lines.append(f"{UI_TEXT['pdf_email']} {sender['email']}")
    return lines


def create_pdf(data: FaxCoverData) -> Path:
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
    top_y = height - 22 * mm
    canvas = pdf_canvas.Canvas(str(output_path), pagesize=A4)
    canvas.setTitle(UI_TEXT["pdf_title"])
    canvas.setLineWidth(0.9)
    canvas.setStrokeColor(pdf_colors.black)
    canvas.setFillColor(pdf_colors.black)

    canvas.setFont(gothic, 10)
    canvas.drawRightString(width - margin_x, top_y, format_japanese_date(data.output_date))

    title_y = top_y - 36
    canvas.setFont(mincho, 22)
    canvas.drawCentredString(width / 2, title_y, UI_TEXT["pdf_title"])
    canvas.line(margin_x, title_y - 12, width - margin_x, title_y - 12)

    block_top = title_y - 42
    left_x = margin_x
    right_x = width / 2 + 8 * mm
    block_width = width / 2 - margin_x - 8 * mm

    canvas.setFont(gothic, 10)
    canvas.drawString(left_x, block_top, UI_TEXT["pdf_recipient_heading"])
    canvas.drawString(right_x, block_top, UI_TEXT["pdf_sender_heading"])
    canvas.line(left_x, block_top - 4, left_x + block_width, block_top - 4)
    canvas.line(right_x, block_top - 4, right_x + block_width, block_top - 4)

    y_left = block_top - 20
    for line in build_recipient_lines(data.recipient):
        font_size = 11 if line.startswith(UI_TEXT["pdf_fax"]) else 9
        y_left = draw_wrapped(canvas, line, left_x, y_left, block_width, gothic, font_size, 13, max_lines=2)

    y_right = block_top - 20
    for line in build_sender_lines(data.sender):
        y_right = draw_wrapped(canvas, line, right_x, y_right, block_width, gothic, 9, 12, max_lines=2)

    fax_info_top = min(y_left, y_right) - 16
    info_height = 50
    canvas.rect(margin_x, fax_info_top - info_height, width - margin_x * 2, info_height, stroke=1, fill=0)
    canvas.setFont(gothic, 10)
    canvas.drawString(margin_x + 8, fax_info_top - 16, UI_TEXT["pdf_fax_info_heading"])
    info_x = margin_x + 74
    info_y = fax_info_top - 16
    draw_label_value(canvas, UI_TEXT["pdf_subject"], data.fax_info["subject"], info_x, info_y, 48, 270, gothic, gothic, 10)
    draw_label_value(canvas, UI_TEXT["pdf_total_pages"], data.fax_info["total_pages"], info_x, info_y - 18, 48, 170, gothic, gothic, 10)
    draw_label_value(
        canvas,
        UI_TEXT["pdf_send_date"],
        format_japanese_date(data.send_date),
        info_x + 210,
        info_y - 18,
        48,
        190,
        gothic,
        gothic,
        10,
    )

    message_y = fax_info_top - info_height - 28
    message = data.message.strip() or UI_TEXT["default_message"]
    message_y = draw_wrapped(canvas, message, margin_x, message_y, width - margin_x * 2, gothic, 10, 15, max_lines=3)

    table_x = margin_x
    table_top = message_y - 18
    col_widths = [270, 70, width - margin_x * 2 - 340]
    row_height = 29
    headers = [
        UI_TEXT["pdf_table_item_name"],
        UI_TEXT["pdf_table_pages"],
        UI_TEXT["pdf_table_note"],
    ]
    table_width = sum(col_widths)

    canvas.setLineWidth(0.8)
    canvas.setFillColor(pdf_colors.HexColor(COLORS["table_header"]))
    canvas.rect(table_x, table_top - row_height, table_width, row_height, stroke=1, fill=1)
    canvas.setFillColor(pdf_colors.black)
    current_x = table_x
    for index, header in enumerate(headers):
        canvas.line(current_x, table_top, current_x, table_top - row_height)
        draw_table_cell(canvas, header, current_x, table_top - row_height, col_widths[index], row_height, gothic, 10, True)
        current_x += col_widths[index]
    canvas.line(current_x, table_top, current_x, table_top - row_height)

    row_y = table_top - row_height
    for item in data.items:
        row_y -= row_height
        canvas.rect(table_x, row_y, table_width, row_height, stroke=1, fill=0)
        current_x = table_x
        values = [item.name, item.pages, item.note]
        for index, value in enumerate(values):
            canvas.line(current_x, row_y + row_height, current_x, row_y)
            draw_table_cell(canvas, value, current_x, row_y, col_widths[index], row_height, gothic, 10, index == 1)
            current_x += col_widths[index]
        canvas.line(current_x, row_y + row_height, current_x, row_y)

    canvas.setFont(gothic, 10)
    canvas.drawRightString(width - margin_x, max(42 * mm, row_y - 28), UI_TEXT["pdf_closing"])
    canvas.showPage()
    canvas.save()
    return output_path


class FaxCoverApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.configure(bg=COLORS["background"])
        self.root.minsize(980, 760)

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

        today = format_japanese_date(date.today())
        self.date_var = tk.StringVar(value=today)
        self.recipient_vars = {
            "company": tk.StringVar(),
            "department": tk.StringVar(),
            "position": tk.StringVar(),
            "name": tk.StringVar(),
            "honorific": tk.StringVar(value=self.config.get("last_honorific") or UI_TEXT["honorific_company"]),
            "fax": tk.StringVar(),
            "tel": tk.StringVar(),
        }
        self.fax_info_vars = {
            "subject": tk.StringVar(),
            "total_pages": tk.StringVar(),
            "send_date": tk.StringVar(value=today),
        }
        self.sender_vars = {
            "company": tk.StringVar(),
            "department": tk.StringVar(),
            "name": tk.StringVar(),
            "postal_code": tk.StringVar(),
            "address": tk.StringVar(),
            "tel": tk.StringVar(),
            "fax": tk.StringVar(),
            "email": tk.StringVar(),
        }
        self.item_rows: list[dict[str, tk.StringVar]] = []
        self.save_folder_var = tk.StringVar(value=self.initial_save_folder())
        self.status_var = tk.StringVar(value=UI_TEXT["status_idle"])
        self.message_text: tk.Text | None = None
        self.items_frame: tk.Frame | None = None
        self.create_button: tk.Button | None = None
        self.status_label: tk.Label | None = None
        self.is_processing = False

        self.load_sender_config()
        self.configure_style()
        self.apply_window_icon()
        self.build_ui()

    def choose_font_family(self) -> str:
        available = set(tkfont.families(self.root))
        for candidate in ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo"):
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
        icon_path = (Path(__file__).resolve().parent / ".." / ".." / "02_assets" / "dake_icon.ico").resolve()
        if not icon_path.exists():
            return
        try:
            self.root.iconbitmap(str(icon_path))
        except tk.TclError:
            pass

    def initial_save_folder(self) -> str:
        saved = self.config.get("last_save_folder")
        if saved and Path(saved).exists():
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
        outer.pack(fill="both", expand=True, padx=24, pady=18)
        outer.grid_columnconfigure(0, weight=1)
        outer.grid_rowconfigure(1, weight=1)

        header = tk.Frame(outer, bg=COLORS["background"])
        header.grid(row=0, column=0, sticky="ew", pady=(0, 14))
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
        content.grid_columnconfigure(0, weight=1, uniform="content")
        content.grid_columnconfigure(1, weight=1, uniform="content")
        content.grid_rowconfigure(0, weight=1)

        left = tk.Frame(content, bg=COLORS["background"])
        right = tk.Frame(content, bg=COLORS["background"])
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        right.grid(row=0, column=1, sticky="nsew", padx=(10, 0))

        self.build_recipient_card(left)
        self.build_fax_info_card(left)
        self.build_items_card(left)
        self.build_message_card(left)
        self.build_sender_card(right)
        self.build_save_card(right)

        bottom = tk.Frame(outer, bg=COLORS["background"])
        bottom.grid(row=2, column=0, sticky="ew", pady=(14, 0))
        bottom.grid_columnconfigure(0, weight=1)
        bottom.grid_columnconfigure(1, weight=0)

        self.status_label = tk.Label(
            bottom,
            textvariable=self.status_var,
            font=self.fonts["small"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
        )
        self.status_label.grid(row=0, column=0, sticky="w")

        footer = tk.Frame(bottom, bg=COLORS["background"])
        footer.grid(row=0, column=1, sticky="e")
        self.add_footer_label(footer, UI_TEXT["footer_left"])
        self.add_footer_label(footer, UI_TEXT["footer_separator"])
        self.add_footer_link(footer, "footer_link_1")
        self.add_footer_label(footer, UI_TEXT["footer_separator"])
        self.add_footer_link(footer, "footer_link_2")
        self.add_footer_label(footer, UI_TEXT["footer_separator"])
        self.add_footer_label(footer, UI_TEXT["footer_copyright"])

    def card(self, parent: tk.Widget, title: str) -> tk.Frame:
        frame = tk.Frame(
            parent,
            bg=COLORS["card"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
            bd=0,
        )
        frame.pack(fill="x", pady=(0, 12))
        inner = tk.Frame(frame, bg=COLORS["card"])
        inner.pack(fill="both", expand=True, padx=16, pady=14)
        tk.Label(
            inner,
            text=title,
            font=self.fonts["section"],
            fg=COLORS["text"],
            bg=COLORS["card"],
        ).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 10))
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
        self.add_labeled_entry(card, 7, 0, UI_TEXT["label_fax_number"], self.recipient_vars["fax"])
        self.add_labeled_entry(card, 8, 0, UI_TEXT["label_tel"], self.recipient_vars["tel"])

    def build_fax_info_card(self, parent: tk.Widget) -> None:
        card = self.card(parent, UI_TEXT["section_fax_info"])
        self.add_labeled_entry(card, 1, 0, UI_TEXT["label_subject"], self.fax_info_vars["subject"])
        self.add_labeled_entry(card, 2, 0, UI_TEXT["label_total_pages"], self.fax_info_vars["total_pages"])
        self.add_labeled_entry(card, 3, 0, UI_TEXT["label_send_date"], self.fax_info_vars["send_date"])

    def build_items_card(self, parent: tk.Widget) -> None:
        card = self.card(parent, UI_TEXT["section_items"])
        headers = [
            UI_TEXT["label_item_name"],
            UI_TEXT["label_item_pages"],
            UI_TEXT["label_note"],
        ]
        for column, header in enumerate(headers):
            tk.Label(
                card,
                text=header,
                font=self.fonts["label"],
                fg=COLORS["muted"],
                bg=COLORS["card"],
            ).grid(row=1, column=column + 1, sticky="w", padx=(0, 8), pady=(0, 6))
        self.items_frame = card
        for _ in range(INITIAL_ITEM_ROWS):
            self.add_item_row()
        add_button = tk.Button(
            card,
            text=UI_TEXT["button_add_row"],
            command=self.add_item_row,
            font=self.fonts["body"],
            fg=COLORS["text"],
            bg=COLORS["card"],
            activebackground=COLORS["background"],
            activeforeground=COLORS["text"],
            relief="solid",
            bd=1,
            highlightthickness=0,
            padx=12,
            pady=5,
            cursor="hand2",
        )
        add_button.grid(row=99, column=1, sticky="w", pady=(8, 0))
        card.grid_columnconfigure(1, weight=3)
        card.grid_columnconfigure(2, weight=1)
        card.grid_columnconfigure(3, weight=3)

    def add_item_row(self) -> None:
        if self.items_frame is None:
            return
        row_vars = {"name": tk.StringVar(), "pages": tk.StringVar(), "note": tk.StringVar()}
        self.item_rows.append(row_vars)
        row_number = len(self.item_rows)
        row_index = row_number + 1
        tk.Label(
            self.items_frame,
            text=UI_TEXT["items_row_template"].format(number=row_number),
            font=self.fonts["small"],
            fg=COLORS["muted"],
            bg=COLORS["card"],
        ).grid(row=row_index, column=0, sticky="e", padx=(0, 8), pady=4)
        for column, key in enumerate(("name", "pages", "note")):
            entry = self.make_entry(self.items_frame, row_vars[key], width=(20 if column != 1 else 7))
            entry.grid(row=row_index, column=column + 1, sticky="ew", padx=(0, 8), pady=4)

    def build_message_card(self, parent: tk.Widget) -> None:
        card = self.card(parent, UI_TEXT["section_message"])
        self.message_text = tk.Text(
            card,
            height=3,
            wrap="word",
            font=self.fonts["body"],
            fg=COLORS["text"],
            bg="#FFFFFF",
            relief="solid",
            bd=1,
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            highlightcolor=COLORS["accent"],
        )
        self.message_text.grid(row=1, column=0, columnspan=4, sticky="ew")
        self.message_text.insert("1.0", UI_TEXT["default_message"])

    def build_sender_card(self, parent: tk.Widget) -> None:
        card = self.card(parent, UI_TEXT["section_sender"])
        fields = [
            ("company", "label_company"),
            ("department", "label_department"),
            ("name", "label_name"),
            ("postal_code", "label_zip"),
            ("address", "label_address"),
            ("tel", "label_tel"),
            ("fax", "label_fax_number"),
            ("email", "label_email"),
        ]
        for row_index, (key, label_key) in enumerate(fields, start=1):
            self.add_labeled_entry(card, row_index, 0, UI_TEXT[label_key], self.sender_vars[key])

    def build_save_card(self, parent: tk.Widget) -> None:
        card = self.card(parent, UI_TEXT["section_save"])
        folder_entry = self.make_entry(card, self.save_folder_var)
        folder_entry.configure(state="readonly")
        folder_entry.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(0, 10))

        select_button = tk.Button(
            card,
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
        select_button.grid(row=2, column=0, sticky="w")

        self.create_button = tk.Button(
            card,
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
            padx=22,
            pady=10,
            cursor="hand2",
        )
        self.create_button.grid(row=2, column=2, sticky="e")
        card.grid_columnconfigure(0, weight=1)
        card.grid_columnconfigure(1, weight=0)
        card.grid_columnconfigure(2, weight=0)

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

    def collect_items(self) -> list[FaxItemRow]:
        rows: list[FaxItemRow] = []
        for row_vars in self.item_rows:
            item = FaxItemRow(
                name=row_vars["name"].get().strip(),
                pages=row_vars["pages"].get().strip(),
                note=row_vars["note"].get().strip(),
            )
            if item.name:
                rows.append(item)
        return rows

    def collect_data(self) -> FaxCoverData:
        output_date = parse_date_input(self.date_var.get(), UI_TEXT["error_date_format"])
        send_date = parse_date_input(self.fax_info_vars["send_date"].get(), UI_TEXT["error_send_date_format"])
        recipient = {key: variable.get().strip() for key, variable in self.recipient_vars.items()}
        fax_info = {key: variable.get().strip() for key, variable in self.fax_info_vars.items()}
        sender = {key: variable.get().strip() for key, variable in self.sender_vars.items()}
        items = self.collect_items()
        save_folder = Path(self.save_folder_var.get()).expanduser()
        message = self.message_text.get("1.0", "end").strip() if self.message_text is not None else UI_TEXT["default_message"]

        errors = []
        if not (recipient["company"] or recipient["name"]):
            errors.append(UI_TEXT["error_required_recipient"])
        if not items:
            errors.append(UI_TEXT["error_required_items"])
        if not sender["name"]:
            errors.append(UI_TEXT["error_required_sender"])
        if not save_folder.exists() or not save_folder.is_dir():
            errors.append(UI_TEXT["error_save_folder"])
        if errors:
            raise ValueError("\n".join(errors))

        return FaxCoverData(
            output_date=output_date,
            send_date=send_date,
            recipient=recipient,
            fax_info=fax_info,
            items=items,
            message=message,
            sender=sender,
            save_folder=save_folder,
        )

    def save_current_config(self) -> None:
        save_config(
            {
                "sender": {key: variable.get().strip() for key, variable in self.sender_vars.items()},
                "last_save_folder": self.save_folder_var.get(),
                "last_honorific": self.recipient_vars["honorific"].get(),
            }
        )

    def handle_create_pdf(self) -> None:
        if self.is_processing:
            messagebox.showinfo(UI_TEXT["message_busy_title"], UI_TEXT["message_busy_body"])
            return
        try:
            data = self.collect_data()
            self.save_current_config()
        except Exception as exc:
            self.set_status("status_error", is_error=True)
            messagebox.showerror(UI_TEXT["message_error_title"], str(exc))
            return
        self.set_processing(True)
        self.set_status("status_processing")
        threading.Thread(target=self.create_pdf_worker, args=(data,), daemon=True).start()

    def create_pdf_worker(self, data: FaxCoverData) -> None:
        try:
            output_path = create_pdf(data)
        except PermissionError:
            self.root.after(0, self.handle_pdf_error, UI_TEXT["error_file_locked"])
        except Exception as exc:
            self.root.after(0, self.handle_pdf_error, str(exc) or UI_TEXT["error_pdf_failed"])
        else:
            self.root.after(0, self.handle_pdf_success, output_path)

    def handle_pdf_error(self, message: str) -> None:
        self.set_status("status_error", is_error=True)
        self.set_processing(False)
        messagebox.showerror(UI_TEXT["message_error_title"], message)

    def handle_pdf_success(self, output_path: Path) -> None:
        self.set_status("status_complete")
        self.set_processing(False)
        message = f"{UI_TEXT['message_complete_body']}\n{UI_TEXT['message_open_folder_body']}"
        messagebox.showinfo(UI_TEXT["message_complete_title"], message)
        open_folder(output_path.parent)


def main() -> None:
    set_windows_app_id()
    root = tk.Tk()
    FaxCoverApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
