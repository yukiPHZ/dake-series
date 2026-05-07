# -*- coding: utf-8 -*-
from __future__ import annotations

import ctypes
import json
import os
import queue
import re
import subprocess
import sys
import threading
import time
import webbrowser
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, ttk


APP_NAME = "Dake工程カレンダー"
WINDOW_TITLE = "工程カレンダー"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "display_name": "工程カレンダー",
    "main_title": "期間カレンダーを作る",
    "main_description": "指定した期間の日付枠だけを、A4縦のPDFに並べます。",
    "section_site": "現場情報",
    "section_period": "期間と保存先",
    "label_site_name": "現場名",
    "label_branch_name": "支店名",
    "label_staff_name": "担当者名",
    "label_phone": "電話番号",
    "label_start_date": "開始日",
    "label_end_date": "終了日",
    "label_save_folder": "保存先",
    "date_hint": "例：2025/11/29",
    "button_select_folder": "選択",
    "button_execute": "PDF作成",
    "status_ready": "入力してPDF作成を押してください。",
    "status_processing": "処理中",
    "status_processing_dots": ["処理中.", "処理中..", "処理中..."],
    "status_phrase_1": "Simple",
    "status_phrase_2": "Simple, fast",
    "status_phrase_3": "Simple, fast, for real work.",
    "status_complete": "PDFを保存しました。",
    "status_error": "PDF作成に失敗しました。",
    "dialog_select_save_dir": "保存先フォルダを選択",
    "dialog_error_title": "入力エラー",
    "dialog_saved_title": "保存完了",
    "dialog_saved_message": "期間カレンダーPDFを保存しました。\n\n{path}",
    "error_start_date": "開始日を正しく入力してください。",
    "error_end_date": "終了日を正しく入力してください。",
    "error_date_order": "終了日は開始日以降の日付にしてください。",
    "error_save_folder": "保存先フォルダを選択してください。",
    "error_reportlab_missing": "PDF生成に必要な reportlab が見つかりません。reportlab をインストールしてください。",
    "footer_left": "シンプルそれDAKEシリーズ ｜ 止まらない、迷わない、すぐ終わる。",
    "footer_link_1": "戸建買取査定",
    "footer_link_2": "Instagram",
    "footer_separator": " ｜ ",
    "footer_copyright": COPYRIGHT,
    "pdf_site_placeholder": "現場名未入力",
    "pdf_title_template": "{site}",
    "pdf_date_template": "{month}/{day}",
    "pdf_date_with_weekday_template": "{month}/{day}（{weekday}）",
    "pdf_full_date_template": "{year}年{month}月{day}日",
    "pdf_range_separator": "〜",
    "pdf_field_separator": "：",
    "pdf_target_period": "対象期間",
    "pdf_day_count": "日数",
    "pdf_day_count_template": "{count}日",
    "pdf_contact_template": "さくら都市{branch} 担当：{staff} {phone}",
    "pdf_weekday_names": ["月", "火", "水", "木", "金", "土", "日"],
    "pdf_weekday_headers": ["日", "月", "火", "水", "木", "金", "土"],
    "pdf_completion_label": "完工",
    "holiday_names": {
        "new_year": "元日",
        "coming_age": "成人の日",
        "foundation": "建国記念の日",
        "emperor": "天皇誕生日",
        "vernal": "春分の日",
        "showa": "昭和の日",
        "greenery": "みどりの日",
        "constitution": "憲法記念日",
        "children": "こどもの日",
        "marine": "海の日",
        "mountain": "山の日",
        "respect": "敬老の日",
        "autumnal": "秋分の日",
        "sports": "スポーツの日",
        "health_sports": "体育の日",
        "culture": "文化の日",
        "labor": "勤労感謝の日",
        "citizens": "国民の休日",
        "substitute": "振替休日",
    },
}

CONFIG_NAME = "dake_work_calendar_config.json"
COMMON_ICON_RELATIVE = Path("..") / ".." / "02_assets" / "dake_icon.ico"
COMMON_ICON_FILENAME = "dake_icon.ico"
FONT_CANDIDATES = ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo")
DATE_FORMATS = ("%Y/%m/%d", "%Y-%m-%d", "%Y.%m.%d", "%Y%m%d")
FOOTER_BREAKPOINT = 900
HEADER_BREAKPOINT = 920
DAYS_PER_ROW = 7
MAX_ROWS_PER_PAGE = 7

LINK_TARGETS = {
    "link_1": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "link_2": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

COLORS = {
    "background": "#F6F7F9",
    "panel": "#FFFFFF",
    "text": "#1E2430",
    "muted": "#667085",
    "border": "#E6EAF0",
    "accent": "#2F6FED",
    "accent_hover": "#2458BF",
    "accent_disabled": "#A8B8D8",
    "entry": "#FFFFFF",
    "footer": "#667085",
}

PDF_COLORS = {
    "text": "#1E2430",
    "muted": "#8A94A6",
    "border": "#D8DEE8",
    "subtle_bg": "#FAFBFC",
    "weekend_bg": "#F9FBFF",
    "holiday_bg": "#FFF8F8",
    "holiday": "#B94040",
    "saturday": "#2559A8",
    "accent": "#2F6FED",
}


@dataclass(frozen=True)
class PdfRequest:
    site_name: str
    branch_name: str
    staff_name: str
    phone: str
    start_date: date
    end_date: date
    save_folder: str


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def config_path() -> Path:
    return app_dir() / CONFIG_NAME


def default_save_folder() -> str:
    downloads = Path.home() / "Downloads"
    if downloads.exists():
        return str(downloads)
    return str(Path.home())


def load_config() -> dict:
    try:
        with config_path().open("r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            return data
    except Exception:
        pass
    return {}


def save_config(data: dict) -> None:
    try:
        with config_path().open("w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def find_icon_path() -> Path | None:
    base = app_dir()
    candidates = [
        base / COMMON_ICON_RELATIVE,
        base / ".." / ".." / ".." / "02_assets" / COMMON_ICON_FILENAME,
        Path(__file__).resolve().parent / COMMON_ICON_RELATIVE,
    ]
    if getattr(sys, "frozen", False):
        candidates.append(Path(getattr(sys, "_MEIPASS", base)) / COMMON_ICON_FILENAME)
    for candidate in candidates:
        resolved = candidate.resolve()
        if resolved.exists():
            return resolved
    return None


def apply_window_icon(window: tk.Tk) -> None:
    icon = find_icon_path()
    if icon is None:
        return
    try:
        window.iconbitmap(default=str(icon))
    except tk.TclError:
        try:
            window.iconbitmap(str(icon))
        except tk.TclError:
            pass


def set_windows_app_id() -> None:
    if not sys.platform.startswith("win"):
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Shimarisu.DakeWorkCalendar")
    except Exception:
        pass


def open_folder(path: str) -> None:
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


def parse_date_text(value: str) -> date | None:
    text = value.strip()
    if not text:
        return None
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    match = re.fullmatch(r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})", text)
    if match:
        try:
            return date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
        except ValueError:
            return None
    return None


def format_full_date(value: date) -> str:
    return UI_TEXT["pdf_full_date_template"].format(year=value.year, month=value.month, day=value.day)


def format_cell_date(value: date) -> str:
    weekday = UI_TEXT["pdf_weekday_names"][value.weekday()]
    return UI_TEXT["pdf_date_with_weekday_template"].format(month=value.month, day=value.day, weekday=weekday)


def sanitize_filename(value: str) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|]+', "_", value).strip()
    cleaned = re.sub(r"\s+", "_", cleaned)
    return cleaned[:60]


def color_hex(hex_value: str):
    from reportlab.lib import colors

    return colors.HexColor(hex_value)


def iter_period_days(start: date, end: date) -> list[date]:
    count = (end - start).days + 1
    return [start + timedelta(days=index) for index in range(count)]


def sunday_on_or_before(value: date) -> date:
    return value - timedelta(days=(value.weekday() + 1) % 7)


def saturday_on_or_after(value: date) -> date:
    return value + timedelta(days=(5 - value.weekday()) % 7)


def iter_period_weeks(start: date, end: date) -> list[list[date | None]]:
    weeks: list[list[date | None]] = []
    current = sunday_on_or_before(start)
    last = saturday_on_or_after(end)
    while current <= last:
        week: list[date | None] = []
        for offset in range(DAYS_PER_ROW):
            day = current + timedelta(days=offset)
            week.append(day if start <= day <= end else None)
        weeks.append(week)
        current += timedelta(days=DAYS_PER_ROW)
    return weeks


def nth_monday(year: int, month: int, nth: int) -> date:
    current = date(year, month, 1)
    days_until_monday = (0 - current.weekday()) % 7
    return current + timedelta(days=days_until_monday + 7 * (nth - 1))


def spring_equinox_day(year: int) -> int:
    return int(20.8431 + 0.242194 * (year - 1980) - ((year - 1980) // 4))


def autumn_equinox_day(year: int) -> int:
    return int(23.2488 + 0.242194 * (year - 1980) - ((year - 1980) // 4))


def fallback_japanese_holidays(years: set[int]) -> dict[date, str]:
    names = UI_TEXT["holiday_names"]
    holidays: dict[date, str] = {}

    for year in years:
        base: dict[date, str] = {}
        base[date(year, 1, 1)] = names["new_year"]
        base[nth_monday(year, 1, 2)] = names["coming_age"]
        base[date(year, 2, 11)] = names["foundation"]

        if year >= 2020:
            base[date(year, 2, 23)] = names["emperor"]
        elif 1989 <= year <= 2018:
            base[date(year, 12, 23)] = names["emperor"]

        base[date(year, 3, spring_equinox_day(year))] = names["vernal"]

        if year >= 2007:
            base[date(year, 4, 29)] = names["showa"]
            base[date(year, 5, 4)] = names["greenery"]
        elif 1989 <= year <= 2006:
            base[date(year, 4, 29)] = names["greenery"]

        base[date(year, 5, 3)] = names["constitution"]
        base[date(year, 5, 5)] = names["children"]

        if year == 2020:
            base[date(year, 7, 23)] = names["marine"]
            base[date(year, 7, 24)] = names["sports"]
            base[date(year, 8, 10)] = names["mountain"]
        elif year == 2021:
            base[date(year, 7, 22)] = names["marine"]
            base[date(year, 7, 23)] = names["sports"]
            base[date(year, 8, 8)] = names["mountain"]
        else:
            if year >= 2003:
                base[nth_monday(year, 7, 3)] = names["marine"]
            elif year >= 1996:
                base[date(year, 7, 20)] = names["marine"]
            if year >= 2016:
                base[date(year, 8, 11)] = names["mountain"]
            if year >= 2020:
                base[nth_monday(year, 10, 2)] = names["sports"]
            elif year >= 2000:
                base[nth_monday(year, 10, 2)] = names["health_sports"]
            else:
                base[date(year, 10, 10)] = names["health_sports"]

        if year >= 2003:
            base[nth_monday(year, 9, 3)] = names["respect"]
        else:
            base[date(year, 9, 15)] = names["respect"]

        base[date(year, 9, autumn_equinox_day(year))] = names["autumnal"]
        base[date(year, 11, 3)] = names["culture"]
        base[date(year, 11, 23)] = names["labor"]

        year_holidays = dict(base)
        current = date(year, 1, 2)
        while current <= date(year, 12, 30):
            if current not in year_holidays:
                previous_is_holiday = current - timedelta(days=1) in year_holidays
                next_is_holiday = current + timedelta(days=1) in year_holidays
                if previous_is_holiday and next_is_holiday:
                    year_holidays[current] = names["citizens"]
            current += timedelta(days=1)

        for holiday_date in sorted(base):
            if holiday_date.weekday() == 6:
                substitute = holiday_date + timedelta(days=1)
                while substitute in year_holidays:
                    substitute += timedelta(days=1)
                if substitute.year == year:
                    year_holidays[substitute] = names["substitute"]

        holidays.update(year_holidays)

    return holidays


def japanese_holidays(years: set[int]) -> dict[date, str]:
    try:
        import holidays as holidays_lib  # type: ignore

        try:
            holiday_map = holidays_lib.country_holidays("JP", years=sorted(years), language="ja")
        except TypeError:
            holiday_map = holidays_lib.country_holidays("JP", years=sorted(years))
        return {holiday_date: str(name) for holiday_date, name in holiday_map.items()}
    except Exception:
        return fallback_japanese_holidays(years)


def register_pdf_fonts() -> str:
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont

    font_name = "HeiseiKakuGo-W5"
    try:
        pdfmetrics.registerFont(UnicodeCIDFont(font_name))
    except Exception:
        pass
    return font_name


def make_output_path(request: PdfRequest) -> Path:
    folder = Path(request.save_folder)
    folder.mkdir(parents=True, exist_ok=True)
    site_part = sanitize_filename(request.site_name) or sanitize_filename(UI_TEXT["display_name"])
    span = f"{request.start_date:%Y%m%d}-{request.end_date:%Y%m%d}"
    output = folder / f"{UI_TEXT['display_name']}_{site_part}_{span}.pdf"
    if not output.exists():
        return output
    suffix = datetime.now().strftime("%H%M%S")
    return folder / f"{UI_TEXT['display_name']}_{site_part}_{span}_{suffix}.pdf"


def generate_pdf(request: PdfRequest) -> Path:
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfgen import canvas
    except Exception as exc:
        raise RuntimeError(UI_TEXT["error_reportlab_missing"]) from exc

    output_path = make_output_path(request)
    page_width, page_height = A4
    pdf = canvas.Canvas(str(output_path), pagesize=A4)
    font_name = register_pdf_fonts()

    days = iter_period_days(request.start_date, request.end_date)
    weeks = iter_period_weeks(request.start_date, request.end_date)
    chunks = [weeks[i : i + MAX_ROWS_PER_PAGE] for i in range(0, len(weeks), MAX_ROWS_PER_PAGE)]
    holiday_map = japanese_holidays({day.year for day in days})

    margin_x = 44
    header_y = page_height - 44
    grid_top = page_height - 136
    day_header_h = 22
    grid_bottom = 112
    grid_width = page_width - margin_x * 2
    cell_w = grid_width / DAYS_PER_ROW
    cell_h = (grid_top - day_header_h - grid_bottom) / MAX_ROWS_PER_PAGE

    site = request.site_name.strip() or UI_TEXT["pdf_site_placeholder"]
    title = UI_TEXT["pdf_title_template"].format(site=site)
    period_text = f"{format_full_date(request.start_date)}{UI_TEXT['pdf_range_separator']}{format_full_date(request.end_date)}"
    day_count_text = UI_TEXT["pdf_day_count_template"].format(count=len(days))
    contact_text = UI_TEXT["pdf_contact_template"].format(
        branch=request.branch_name.strip(),
        staff=request.staff_name.strip(),
        phone=request.phone.strip(),
    ).strip()

    for _page_index, chunk in enumerate(chunks, start=1):
        pdf.setFont(font_name, 17)
        pdf.setFillColor(color_hex(PDF_COLORS["text"]))
        pdf.drawString(margin_x, header_y, title)

        separator = UI_TEXT["pdf_field_separator"]
        info_y = header_y - 25
        pdf.setFont(font_name, 9.5)
        pdf.setFillColor(color_hex(PDF_COLORS["text"]))
        pdf.drawString(margin_x, info_y, f"{UI_TEXT['pdf_target_period']}{separator}{period_text}")
        pdf.drawRightString(page_width - margin_x, info_y, f"{UI_TEXT['pdf_day_count']}{separator}{day_count_text}")

        for col, header in enumerate(UI_TEXT["pdf_weekday_headers"]):
            x = margin_x + col * cell_w
            y = grid_top - day_header_h
            pdf.setFillColor(color_hex("#FFFFFF"))
            pdf.rect(x, y, cell_w, day_header_h, stroke=0, fill=1)
            pdf.setStrokeColor(color_hex(PDF_COLORS["border"]))
            pdf.rect(x, y, cell_w, day_header_h, stroke=1, fill=0)
            if col == 0:
                pdf.setFillColor(color_hex(PDF_COLORS["holiday"]))
            elif col == 6:
                pdf.setFillColor(color_hex(PDF_COLORS["saturday"]))
            else:
                pdf.setFillColor(color_hex(PDF_COLORS["text"]))
            pdf.setFont(font_name, 9.2)
            pdf.drawCentredString(x + cell_w / 2, y + 7, header)

        for row, week in enumerate(chunk):
            for col, day in enumerate(week):
                x = margin_x + col * cell_w
                y = grid_top - day_header_h - (row + 1) * cell_h
                if day is None:
                    pdf.setFillColor(color_hex(PDF_COLORS["subtle_bg"]))
                    pdf.rect(x, y, cell_w, cell_h, stroke=0, fill=1)
                    pdf.setStrokeColor(color_hex(PDF_COLORS["border"]))
                    pdf.rect(x, y, cell_w, cell_h, stroke=1, fill=0)
                    continue

                holiday_name = holiday_map.get(day)
                is_saturday = day.weekday() == 5
                is_sunday = day.weekday() == 6

                if holiday_name or is_sunday:
                    fill = PDF_COLORS["holiday_bg"]
                elif is_saturday:
                    fill = PDF_COLORS["weekend_bg"]
                else:
                    fill = "#FFFFFF"

                pdf.setFillColor(color_hex(fill))
                pdf.rect(x, y, cell_w, cell_h, stroke=0, fill=1)
                pdf.setStrokeColor(color_hex(PDF_COLORS["border"]))
                pdf.rect(x, y, cell_w, cell_h, stroke=1, fill=0)

                date_color = PDF_COLORS["text"]
                if holiday_name or is_sunday:
                    date_color = PDF_COLORS["holiday"]
                elif is_saturday:
                    date_color = PDF_COLORS["saturday"]

                pdf.setFillColor(color_hex(date_color))
                pdf.setFont(font_name, 10.5)
                text_y = y + cell_h - 17
                pdf.drawString(x + 7, text_y, format_cell_date(day))

                text_y -= 15
                if holiday_name:
                    pdf.setFont(font_name, 7.2)
                    pdf.setFillColor(color_hex(PDF_COLORS["holiday"]))
                    pdf.drawString(x + 7, text_y, holiday_name)
                    text_y -= 14

                if day == request.end_date:
                    pdf.setFont(font_name, 9.5)
                    pdf.setFillColor(color_hex(PDF_COLORS["accent"]))
                    pdf.drawString(x + 7, text_y, UI_TEXT["pdf_completion_label"])

        pdf.setStrokeColor(color_hex(PDF_COLORS["border"]))
        pdf.line(margin_x, 84, page_width - margin_x, 84)
        pdf.setFont(font_name, 9.4)
        pdf.setFillColor(color_hex(PDF_COLORS["text"]))
        pdf.drawRightString(page_width - margin_x, 65, contact_text)
        pdf.setFont(font_name, 7.4)
        pdf.setFillColor(color_hex(PDF_COLORS["muted"]))
        pdf.drawString(margin_x, 49, UI_TEXT["footer_left"])
        pdf.showPage()

    pdf.save()
    return output_path


class WorkCalendarApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.configure(bg=COLORS["background"])
        self.root.minsize(820, 560)
        apply_window_icon(self.root)

        self.config_data = load_config()
        today = date.today()
        self.site_name_var = tk.StringVar()
        self.branch_name_var = tk.StringVar(value=self.config_data.get("branch_name", ""))
        self.staff_name_var = tk.StringVar(value=self.config_data.get("staff_name", ""))
        self.phone_var = tk.StringVar(value=self.config_data.get("phone", ""))
        self.start_date_var = tk.StringVar(value=today.strftime("%Y/%m/%d"))
        self.end_date_var = tk.StringVar(value=(today + timedelta(days=45)).strftime("%Y/%m/%d"))
        self.save_folder_var = tk.StringVar(value=self.config_data.get("save_folder", default_save_folder()))
        self.status_var = tk.StringVar(value=UI_TEXT["status_ready"])

        self.font_family = self.choose_font_family()
        self.fonts = {
            "title": (self.font_family, 20, "bold"),
            "subtitle": (self.font_family, 10),
            "section": (self.font_family, 12, "bold"),
            "body": (self.font_family, 10),
            "small": (self.font_family, 9),
            "button": (self.font_family, 10, "bold"),
            "footer": (self.font_family, 8),
        }

        self.result_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.is_processing = False
        self.processing_started_at = 0.0
        self.status_tick = 0
        self.footer_mode = ""
        self.header_mode = ""

        self.build_ui()
        self.root.bind("<Configure>", self.on_root_configure)
        self.root.after(80, self.apply_responsive_layout)

    def choose_font_family(self) -> str:
        available = set(tkfont.families(self.root))
        for candidate in FONT_CANDIDATES:
            if candidate in available:
                return candidate
        return "TkDefaultFont"

    def build_ui(self) -> None:
        outer = tk.Frame(self.root, bg=COLORS["background"])
        outer.pack(fill="both", expand=True, padx=22, pady=(18, 0))

        self.canvas = tk.Canvas(outer, bg=COLORS["background"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=self.canvas.yview)
        self.content = tk.Frame(self.canvas, bg=COLORS["background"])
        window_id = self.canvas.create_window((0, 0), window=self.content, anchor="nw")

        def update_scrollregion(_event=None) -> None:
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        def update_width(event) -> None:
            self.canvas.itemconfigure(window_id, width=event.width)

        def on_mousewheel(event) -> None:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        self.content.bind("<Configure>", update_scrollregion)
        self.canvas.bind("<Configure>", update_width)
        self.canvas.bind("<Enter>", lambda _event: self.canvas.bind_all("<MouseWheel>", on_mousewheel))
        self.canvas.bind("<Leave>", lambda _event: self.canvas.unbind_all("<MouseWheel>"))
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.build_header()
        self.build_input_sections()
        self.build_actions()
        self.build_footer()

    def build_header(self) -> None:
        self.header = tk.Frame(self.content, bg=COLORS["background"])
        self.header.pack(fill="x", pady=(0, 14))
        self.header_title = tk.Label(
            self.header,
            text=UI_TEXT["main_title"],
            font=self.fonts["title"],
            fg=COLORS["text"],
            bg=COLORS["background"],
            anchor="w",
        )
        self.header_description = tk.Label(
            self.header,
            text=UI_TEXT["main_description"],
            font=self.fonts["subtitle"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
            anchor="w",
        )

    def panel(self, parent: tk.Misc, title: str) -> tk.Frame:
        wrapper = tk.Frame(
            parent,
            bg=COLORS["panel"],
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            highlightcolor=COLORS["border"],
        )
        wrapper.pack(fill="x", pady=(0, 12))
        tk.Label(
            wrapper,
            text=title,
            font=self.fonts["section"],
            fg=COLORS["text"],
            bg=COLORS["panel"],
            anchor="w",
        ).pack(fill="x", padx=18, pady=(14, 8))
        body = tk.Frame(wrapper, bg=COLORS["panel"])
        body.pack(fill="x", padx=18, pady=(0, 16))
        return body

    def build_input_sections(self) -> None:
        site_panel = self.panel(self.content, UI_TEXT["section_site"])
        for col in range(4):
            site_panel.columnconfigure(col, weight=1 if col in (1, 3) else 0)
        self.add_labeled_entry(site_panel, 0, 0, UI_TEXT["label_site_name"], self.site_name_var)
        self.add_labeled_entry(site_panel, 0, 2, UI_TEXT["label_branch_name"], self.branch_name_var)
        self.add_labeled_entry(site_panel, 1, 0, UI_TEXT["label_staff_name"], self.staff_name_var)
        self.add_labeled_entry(site_panel, 1, 2, UI_TEXT["label_phone"], self.phone_var)

        period_panel = self.panel(self.content, UI_TEXT["section_period"])
        for col in range(5):
            period_panel.columnconfigure(col, weight=1 if col in (1, 3) else 0)
        self.add_labeled_entry(period_panel, 0, 0, UI_TEXT["label_start_date"], self.start_date_var)
        self.add_labeled_entry(period_panel, 0, 2, UI_TEXT["label_end_date"], self.end_date_var)
        tk.Label(
            period_panel,
            text=UI_TEXT["date_hint"],
            font=self.fonts["small"],
            fg=COLORS["muted"],
            bg=COLORS["panel"],
            anchor="w",
        ).grid(row=1, column=1, columnspan=3, sticky="w", pady=(2, 8))

        tk.Label(
            period_panel,
            text=UI_TEXT["label_save_folder"],
            font=self.fonts["body"],
            fg=COLORS["text"],
            bg=COLORS["panel"],
            anchor="w",
        ).grid(row=2, column=0, sticky="w", padx=(0, 8), pady=6)
        tk.Entry(
            period_panel,
            textvariable=self.save_folder_var,
            font=self.fonts["body"],
            fg=COLORS["text"],
            bg=COLORS["entry"],
            relief="solid",
            bd=1,
        ).grid(row=2, column=1, columnspan=3, sticky="ew", pady=6)
        self.make_button(period_panel, UI_TEXT["button_select_folder"], self.choose_save_folder, compact=True).grid(
            row=2,
            column=4,
            sticky="e",
            padx=(10, 0),
            pady=6,
        )

    def build_actions(self) -> None:
        action_panel = tk.Frame(self.content, bg=COLORS["background"])
        action_panel.pack(fill="x", pady=(2, 18))
        self.progress = ttk.Progressbar(action_panel, mode="indeterminate", length=220)
        self.progress.pack(side="left", padx=(0, 12))
        tk.Label(
            action_panel,
            textvariable=self.status_var,
            font=self.fonts["body"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
        ).pack(side="left")
        self.create_button = self.make_button(action_panel, UI_TEXT["button_execute"], self.on_create_pdf)
        self.create_button.pack(side="right")

    def build_footer(self) -> None:
        self.footer = tk.Frame(self.root, bg=COLORS["background"])
        self.footer.pack(fill="x", padx=22, pady=(8, 12))

    def add_labeled_entry(
        self,
        parent: tk.Misc,
        row: int,
        column: int,
        label: str,
        variable: tk.StringVar,
    ) -> None:
        tk.Label(
            parent,
            text=label,
            font=self.fonts["body"],
            fg=COLORS["text"],
            bg=COLORS["panel"],
            anchor="w",
        ).grid(row=row, column=column, sticky="w", padx=(0, 8), pady=6)
        tk.Entry(
            parent,
            textvariable=variable,
            font=self.fonts["body"],
            fg=COLORS["text"],
            bg=COLORS["entry"],
            relief="solid",
            bd=1,
        ).grid(row=row, column=column + 1, sticky="ew", padx=(0, 18), pady=6)

    def make_button(self, parent: tk.Misc, text: str, command, compact: bool = False) -> tk.Button:
        return tk.Button(
            parent,
            text=text,
            command=command,
            font=self.fonts["button"],
            fg="#FFFFFF",
            bg=COLORS["accent"],
            activeforeground="#FFFFFF",
            activebackground=COLORS["accent_hover"],
            relief="flat",
            bd=0,
            padx=14 if compact else 24,
            pady=6 if compact else 10,
            cursor="hand2",
        )

    def link_label(self, parent: tk.Misc, text: str, url: str) -> tk.Label:
        label = tk.Label(
            parent,
            text=text,
            font=self.fonts["footer"],
            fg=COLORS["footer"],
            bg=COLORS["background"],
            cursor="hand2",
        )
        label.bind("<Button-1>", lambda _event: webbrowser.open(url))
        label.bind("<Enter>", lambda _event: label.configure(fg=COLORS["accent"]))
        label.bind("<Leave>", lambda _event: label.configure(fg=COLORS["footer"]))
        return label

    def plain_footer_label(self, parent: tk.Misc, text: str) -> tk.Label:
        return tk.Label(
            parent,
            text=text,
            font=self.fonts["footer"],
            fg=COLORS["footer"],
            bg=COLORS["background"],
        )

    def footer_link_row(self, parent: tk.Misc) -> tk.Frame:
        row = tk.Frame(parent, bg=COLORS["background"])
        self.link_label(row, UI_TEXT["footer_link_1"], LINK_TARGETS["link_1"]).pack(side="left")
        self.plain_footer_label(row, UI_TEXT["footer_separator"]).pack(side="left")
        self.link_label(row, UI_TEXT["footer_link_2"], LINK_TARGETS["link_2"]).pack(side="left")
        self.plain_footer_label(row, UI_TEXT["footer_separator"] + UI_TEXT["footer_copyright"]).pack(side="left")
        return row

    def on_root_configure(self, event) -> None:
        if event.widget is self.root:
            self.apply_responsive_layout()

    def apply_responsive_layout(self) -> None:
        width = max(self.root.winfo_width(), self.root.winfo_reqwidth())
        self.layout_header(width)
        self.layout_footer(width)

    def layout_header(self, width: int) -> None:
        mode = "wide" if width >= HEADER_BREAKPOINT else "narrow"
        if mode == self.header_mode:
            return
        self.header_mode = mode
        for widget in (self.header_title, self.header_description):
            widget.pack_forget()
        if mode == "wide":
            self.header_title.pack(side="left", anchor="w")
            self.header_description.pack(side="left", anchor="w", padx=(18, 0), pady=(6, 0))
        else:
            self.header_title.pack(anchor="w")
            self.header_description.pack(anchor="w", pady=(4, 0))

    def layout_footer(self, width: int) -> None:
        mode = "wide" if width >= FOOTER_BREAKPOINT else "narrow"
        if mode == self.footer_mode:
            return
        self.footer_mode = mode
        for child in self.footer.winfo_children():
            child.destroy()

        if mode == "wide":
            self.plain_footer_label(self.footer, UI_TEXT["footer_left"]).pack(side="left")
            self.footer_link_row(self.footer).pack(side="right")
        else:
            line_1 = tk.Frame(self.footer, bg=COLORS["background"])
            line_1.pack(anchor="center")
            self.plain_footer_label(line_1, UI_TEXT["footer_left"]).pack()
            line_2 = self.footer_link_row(self.footer)
            line_2.pack(anchor="center", pady=(2, 0))

    def choose_save_folder(self) -> None:
        initial_dir = self.save_folder_var.get().strip() or default_save_folder()
        selected = filedialog.askdirectory(
            title=UI_TEXT["dialog_select_save_dir"],
            initialdir=initial_dir if os.path.isdir(initial_dir) else default_save_folder(),
        )
        if selected:
            self.save_folder_var.set(selected)

    def collect_request(self) -> PdfRequest | None:
        start = parse_date_text(self.start_date_var.get())
        if start is None:
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["error_start_date"])
            return None

        end = parse_date_text(self.end_date_var.get())
        if end is None:
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["error_end_date"])
            return None

        if end < start:
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["error_date_order"])
            return None

        save_folder = self.save_folder_var.get().strip()
        if not save_folder:
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["error_save_folder"])
            return None

        return PdfRequest(
            site_name=self.site_name_var.get().strip(),
            branch_name=self.branch_name_var.get().strip(),
            staff_name=self.staff_name_var.get().strip(),
            phone=self.phone_var.get().strip(),
            start_date=start,
            end_date=end,
            save_folder=save_folder,
        )

    def on_create_pdf(self) -> None:
        if self.is_processing:
            return
        request = self.collect_request()
        if request is None:
            return

        self.save_current_config()
        self.is_processing = True
        self.processing_started_at = time.monotonic()
        self.status_tick = 0
        self.status_var.set(UI_TEXT["status_processing_dots"][0])
        self.create_button.configure(state="disabled", bg=COLORS["accent_disabled"])
        self.progress.start(12)

        worker = threading.Thread(target=self.worker_generate_pdf, args=(request,), daemon=True)
        worker.start()
        self.root.after(220, self.animate_processing_status)
        self.root.after(100, self.poll_worker)

    def animate_processing_status(self) -> None:
        if not self.is_processing:
            return
        dots = UI_TEXT["status_processing_dots"]
        phrases = (UI_TEXT["status_phrase_1"], UI_TEXT["status_phrase_2"], UI_TEXT["status_phrase_3"])
        elapsed = time.monotonic() - self.processing_started_at
        if elapsed >= 1.6 and self.status_tick % 9 in (5, 6, 7):
            phrase_index = (self.status_tick % 9) - 5
            self.status_var.set(phrases[phrase_index])
        else:
            self.status_var.set(dots[self.status_tick % len(dots)])
        self.status_tick += 1
        self.root.after(420, self.animate_processing_status)

    def worker_generate_pdf(self, request: PdfRequest) -> None:
        try:
            output_path = generate_pdf(request)
            self.result_queue.put(("ok", output_path))
        except Exception as exc:
            self.result_queue.put(("error", str(exc)))

    def poll_worker(self) -> None:
        try:
            status, payload = self.result_queue.get_nowait()
        except queue.Empty:
            self.root.after(100, self.poll_worker)
            return

        self.progress.stop()
        self.create_button.configure(state="normal", bg=COLORS["accent"])
        self.is_processing = False

        if status == "ok":
            output_path = Path(payload)
            self.status_var.set(UI_TEXT["status_complete"])
            messagebox.showinfo(
                UI_TEXT["dialog_saved_title"],
                UI_TEXT["dialog_saved_message"].format(path=str(output_path)),
            )
            open_folder(str(output_path.parent))
        else:
            self.status_var.set(UI_TEXT["status_error"])
            messagebox.showerror(UI_TEXT["dialog_error_title"], str(payload))

    def save_current_config(self) -> None:
        save_config(
            {
                "staff_name": self.staff_name_var.get().strip(),
                "phone": self.phone_var.get().strip(),
                "branch_name": self.branch_name_var.get().strip(),
                "save_folder": self.save_folder_var.get().strip(),
            }
        )


def main() -> None:
    set_windows_app_id()
    root = tk.Tk()
    WorkCalendarApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
