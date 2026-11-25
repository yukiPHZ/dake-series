# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import os
import re
import sys
import tempfile
import tkinter as tk
import webbrowser
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from tkinter import filedialog, font as tkfont
from tkinter import messagebox, ttk


APP_NAME = "Dakeマンション工程表"
WINDOW_TITLE = "Dakeマンション工程表"
COPYRIGHT = "© 2026 しまリス不動産 / Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "main_title": "マンション工程表を作る",
    "main_description": "管理会社提出用のA3横工程表を作成します。",
    "field_site_name": "現場名",
    "field_branch_name": "会社支店名",
    "field_person_name": "担当者名",
    "field_contact": "連絡先",
    "field_handover_date": "引渡し日",
    "field_start_date": "着工日",
    "field_finish_date": "完工日",
    "field_save_path": "保存先",
    "table_work_name": "工事名",
    "table_start_date": "開始日",
    "table_end_date": "終了日",
    "button_create_pdf": "PDFを作成",
    "button_auto_schedule": "日程を自動入力",
    "button_select_save_path": "保存先を選ぶ",
    "status_label": "状態",
    "status_ready": "入力してPDFを作成できます。",
    "status_auto_done": "日程を自動入力しました。",
    "status_pdf_done": "PDFを作成しました: {path}",
    "status_pdf_done_folder_open_failed": "PDFを作成しました。保存フォルダは手動で確認してください: {path}",
    "status_error": "入力内容を確認してください。",
    "dialog_error_title": "確認してください",
    "dialog_done_title": "PDFを作成しました",
    "dialog_done_message": "工程表PDFを作成しました。\n保存先フォルダを確認してください。\n\n{path}",
    "dialog_save_cancelled": "保存先が選ばれていません。",
    "error_date_format": "{label}は YYYY-MM-DD 形式で入力してください。",
    "error_date_order": "{start_label}は{end_label}以前の日付にしてください。",
    "error_missing_reportlab": "PDF作成に必要な reportlab が見つかりません。\nrequirements.txt の内容をインストールしてください。",
    "error_pdf_no_rows": "工程表の日付範囲に表示できる日付がありません。",
    "pdf_title": "工事工程表",
    "pdf_site_name": "現場名",
    "pdf_branch_name": "会社支店名",
    "pdf_person_name": "担当者名",
    "pdf_contact": "連絡先",
    "pdf_period": "工期",
    "pdf_work_name": "工事項目",
    "pdf_month_format": "{year}年{month}月",
    "pdf_weekdays": "月火水木金土日",
    "launch_check_ok": "Dakeマンション工程表 launch-check OK",
    "launch_check_pdf": "pdf={path}",
    "launch_check_page": "a3_landscape_one_page=OK",
    "launch_check_span": "calendar_days={days} start={start} finish={finish}",
    "launch_check_config": "config_save_restore=OK",
    "launch_check_folder": "open_output_folder=OK",
    "launch_check_footer": "pdf_submission_footer_removed=OK",
    "launch_check_workdays": "workdays={workdays} allocated_workdays={allocated}",
    "launch_check_weekend_bars": "weekend_work_bars=none",
    "launch_check_reserve": "reserve_workdays={reserve}",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_tagline": "止まらない、迷わない、すぐ終わる。",
    "footer_link_1": "戸建買取査定",
    "footer_link_2": "Instagram",
    "footer_separator": " ｜ ",
    "footer_copyright": COPYRIGHT,
}

THEME = {
    "background": "#F6F7F9",
    "panel": "#FFFFFF",
    "subtle": "#EEF2F7",
    "border": "#E2E8F0",
    "text": "#1E2430",
    "muted": "#667085",
    "accent": "#2F6FED",
    "accent_hover": "#2458BF",
    "success": "#118A4E",
    "success_bg": "#EAFBF3",
    "error": "#C92A2A",
    "error_bg": "#FDECEC",
    "link": "#58677D",
    "link_hover": "#2F6FED",
}

WINDOW_SIZE = "1040x760"
WINDOW_MIN_SIZE = (820, 680)
DATE_FORMAT = "%Y-%m-%d"
PDF_FILE_NAME = "mansion_schedule.pdf"
CONFIG_FILE_NAME = "mansion_schedule_config.json"
TOTAL_CALENDAR_DAYS = 45
TOTAL_INITIAL_WORKDAYS = 28
CONFIG_KEYS = ("branch_name", "person_name", "contact")
FOOTER_COMPACT_WIDTH = 900

LINKS = {
    "assessment": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "instagram": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

WORK_ITEMS: tuple[tuple[str, int], ...] = (
    ("養生", 1),
    ("残置物撤去", 2),
    ("電気工事", 1),
    ("水道設備解体", 2),
    ("建具工事", 1),
    ("畳工事", 1),
    ("木工事", 5),
    ("ユニットバス設置工事", 1),
    ("システムキッチン設置工事", 1),
    ("木工事", 1),
    ("クロス工事", 5),
    ("建具工事", 1),
    ("サッシ工事", 1),
    ("畳工事", 1),
    ("電気工事", 1),
    ("水道設備設置工事", 1),
    ("ハウスクリーニング", 2),
)

RESERVE_GAP_ANCHORS_BY_COUNT = {
    1: (10,),
    2: (8, 16),
    3: (6, 10, 16),
    4: (6, 8, 12, 16),
    5: (6, 8, 10, 12, 16),
}


@dataclass(frozen=True)
class ScheduleRow:
    name: str
    start_date: date
    end_date: date


@dataclass(frozen=True)
class ProjectInfo:
    site_name: str
    branch_name: str
    person_name: str
    contact: str
    handover_date: date
    start_date: date
    finish_date: date


def get_source_dir() -> Path:
    return Path(__file__).resolve().parent


def get_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return get_source_dir()


def get_common_icon_candidates() -> list[Path]:
    source_dir = get_source_dir()
    base_dir = get_base_dir()
    return [
        source_dir / ".." / ".." / "02_assets" / "dake_icon.ico",
        base_dir / ".." / ".." / "02_assets" / "dake_icon.ico",
        base_dir / ".." / ".." / ".." / "02_assets" / "dake_icon.ico",
    ]


def choose_font_family(root: tk.Tk) -> str:
    preferred = ["BIZ UDPGothic", "Yu Gothic UI", "Meiryo"]
    available = set(tkfont.families(root))
    for family in preferred:
        if family in available:
            return family
    return "TkDefaultFont"


def is_weekday(value: date) -> bool:
    return value.weekday() < 5


def days_from(start: date, count: int) -> list[date]:
    return [start + timedelta(days=index) for index in range(max(0, count))]


def days_between(start: date, finish: date) -> list[date]:
    days: list[date] = []
    current = start
    while current <= finish:
        days.append(current)
        current += timedelta(days=1)
    return days


def workdays_between(start: date, finish: date) -> list[date]:
    return [day for day in days_between(start, finish) if is_weekday(day)]


def parse_date(value: str, label: str) -> date:
    try:
        parsed = datetime.strptime(value.strip(), DATE_FORMAT).date()
    except ValueError as exc:
        raise ValueError(UI_TEXT["error_date_format"].format(label=label)) from exc
    return parsed


def format_date(value: date) -> str:
    return value.strftime(DATE_FORMAT)


def default_save_path(site_name: str = "") -> Path:
    clean_name = re.sub(r'[<>:"/\\|?*\r\n\t]+', "_", site_name.strip())
    clean_name = clean_name.strip(" .") or "mansion_schedule"
    return get_base_dir() / f"{clean_name}_工程表.pdf"


def get_config_path() -> Path:
    return get_base_dir() / CONFIG_FILE_NAME


def load_config(config_path: Path | None = None) -> dict[str, str]:
    path = config_path or get_config_path()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    if not isinstance(data, dict):
        return {}
    return {key: value for key in CONFIG_KEYS if isinstance((value := data.get(key)), str)}


def save_config(values: dict[str, str], config_path: Path | None = None) -> bool:
    path = config_path or get_config_path()
    data = {key: values.get(key, "").strip() for key in CONFIG_KEYS}
    try:
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        return False
    return True


def open_output_folder(output_path: Path, opener=None) -> bool:
    folder = output_path.parent.resolve()
    startfile = opener or getattr(os, "startfile", None)
    if startfile is None:
        return False
    try:
        startfile(str(folder))
    except OSError:
        return False
    return True


def compressed_work_durations(available_workdays: int) -> list[int]:
    desired = [duration for _name, duration in WORK_ITEMS]
    minimums: list[int] = []
    for duration in desired:
        if duration >= 8:
            minimums.append(6)
        elif duration >= 6:
            minimums.append(5)
        else:
            minimums.append(1)

    minimum_total = sum(minimums)
    if available_workdays < minimum_total:
        raise ValueError(UI_TEXT["error_pdf_no_rows"])
    if available_workdays >= sum(desired):
        return desired

    durations = minimums[:]
    extra_days = available_workdays - minimum_total
    priority_indexes = (
        [index for index, duration in enumerate(desired) if duration == 2]
        + [index for index, duration in enumerate(desired) if duration >= 8]
        + [index for index, duration in enumerate(desired) if duration >= 6]
    )
    while extra_days > 0:
        changed = False
        for index in priority_indexes:
            if extra_days <= 0:
                break
            if durations[index] < desired[index]:
                durations[index] += 1
                extra_days -= 1
                changed = True
        if not changed:
            break
    return durations


def reserve_gap_after_rows(reserve_workdays: int) -> dict[int, int]:
    if reserve_workdays <= 0:
        return {}

    anchor_count = min(reserve_workdays, max(RESERVE_GAP_ANCHORS_BY_COUNT))
    anchors = RESERVE_GAP_ANCHORS_BY_COUNT[anchor_count]
    gaps: dict[int, int] = {}
    for index in range(reserve_workdays):
        row_number = anchors[index % len(anchors)]
        gaps[row_number] = gaps.get(row_number, 0) + 1
    return gaps


def weekday_segments(start: date, finish: date, valid_days: set[date] | None = None) -> list[tuple[date, date]]:
    candidates = [
        day
        for day in days_between(start, finish)
        if is_weekday(day) and (valid_days is None or day in valid_days)
    ]
    if not candidates:
        return []

    segments: list[tuple[date, date]] = []
    segment_start = candidates[0]
    previous = candidates[0]
    for day in candidates[1:]:
        if (day - previous).days == 1:
            previous = day
            continue
        segments.append((segment_start, previous))
        segment_start = day
        previous = day
    segments.append((segment_start, previous))
    return segments


def build_auto_schedule(start_date: date) -> list[ScheduleRow]:
    schedule_days = days_from(start_date, TOTAL_CALENDAR_DAYS)
    workdays = [day for day in schedule_days if is_weekday(day)]
    durations = compressed_work_durations(len(workdays))
    reserve_workdays = max(0, len(workdays) - sum(durations))
    gap_after_rows = reserve_gap_after_rows(reserve_workdays)
    rows: list[ScheduleRow] = []
    index = 0
    for row_number, ((name, _template_duration), duration) in enumerate(zip(WORK_ITEMS, durations), start=1):
        start_index = index
        end_index = min(index + duration - 1, len(workdays) - 1)
        rows.append(ScheduleRow(name=name, start_date=workdays[start_index], end_date=workdays[end_index]))
        index += duration + gap_after_rows.get(row_number, 0)
    return rows


def build_project_from_handover(handover_date: date) -> tuple[date, date, list[ScheduleRow]]:
    start_date = handover_date + timedelta(days=3)
    finish_date = start_date + timedelta(days=TOTAL_CALENDAR_DAYS - 1)
    return start_date, finish_date, build_auto_schedule(start_date)


def register_pdf_fonts() -> tuple[str, str]:
    try:
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    except ModuleNotFoundError as exc:
        raise RuntimeError(UI_TEXT["error_missing_reportlab"]) from exc

    normal_font = "HeiseiKakuGo-W5"
    bold_font = "HeiseiKakuGo-W5"
    try:
        pdfmetrics.registerFont(UnicodeCIDFont(normal_font))
    except Exception:
        pass
    return normal_font, bold_font


def draw_fitted_text(canvas, text: str, x: float, y: float, max_width: float, font_name: str, max_size: float, min_size: float = 5.0) -> None:
    from reportlab.pdfbase import pdfmetrics

    size = max_size
    while size > min_size and pdfmetrics.stringWidth(text, font_name, size) > max_width:
        size -= 0.5
    canvas.setFont(font_name, size)
    canvas.drawString(x, y, text)


def draw_right_text(canvas, text: str, right: float, y: float, font_name: str, size: float) -> None:
    canvas.setFont(font_name, size)
    canvas.drawRightString(right, y, text)


def draw_pdf_header(canvas, info: ProjectInfo, font_name: str, bold_font: str, width: float, height: float) -> None:
    from reportlab.lib import colors

    margin_x = 28
    top = height - 30
    canvas.setFillColor(colors.HexColor("#1E2430"))
    canvas.setFont(bold_font, 20)
    canvas.drawString(margin_x, top, UI_TEXT["pdf_title"])

    meta_y = top - 26
    left_items = (
        (UI_TEXT["pdf_site_name"], info.site_name),
        (UI_TEXT["pdf_branch_name"], info.branch_name),
        (UI_TEXT["pdf_person_name"], info.person_name),
        (UI_TEXT["pdf_contact"], info.contact),
    )
    label_width = 52
    value_width = 290
    for index, (label, value) in enumerate(left_items):
        x = margin_x + (index % 2) * 380
        y = meta_y - (index // 2) * 18
        canvas.setFillColor(colors.HexColor("#667085"))
        canvas.setFont(font_name, 8)
        canvas.drawString(x, y, label)
        canvas.setFillColor(colors.HexColor("#1E2430"))
        draw_fitted_text(canvas, value or " ", x + label_width, y, value_width, font_name, 9)

    period = f"{format_date(info.start_date)} 〜 {format_date(info.finish_date)}"
    canvas.setFillColor(colors.HexColor("#667085"))
    canvas.setFont(font_name, 8)
    canvas.drawRightString(width - 210, meta_y, UI_TEXT["pdf_period"])
    canvas.setFillColor(colors.HexColor("#1E2430"))
    draw_right_text(canvas, period, width - margin_x, meta_y, font_name, 10)


def draw_schedule_pdf(output_path: Path, info: ProjectInfo, rows: list[ScheduleRow]) -> None:
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A3, landscape
        from reportlab.pdfgen import canvas
    except ModuleNotFoundError as exc:
        raise RuntimeError(UI_TEXT["error_missing_reportlab"]) from exc

    font_name, bold_font = register_pdf_fonts()
    page_size = landscape(A3)
    width, height = page_size
    output_path.parent.mkdir(parents=True, exist_ok=True)
    pdf = canvas.Canvas(str(output_path), pagesize=page_size)
    pdf.setTitle(UI_TEXT["pdf_title"])

    draw_pdf_header(pdf, info, font_name, bold_font, width, height)

    axis_days = days_between(info.start_date, info.finish_date)
    if not axis_days:
        raise ValueError(UI_TEXT["error_pdf_no_rows"])

    margin_x = 28
    chart_top = height - 118
    chart_bottom = 32
    chart_left = margin_x
    chart_right = width - margin_x
    label_width = 158
    month_height = 18
    date_height = 24
    row_count = len(rows)
    row_height = (chart_top - chart_bottom - month_height - date_height) / row_count
    axis_left = chart_left + label_width
    axis_width = chart_right - axis_left
    day_width = axis_width / len(axis_days)
    header_bottom = chart_top - month_height - date_height

    border = colors.HexColor("#D8DEE8")
    grid = colors.HexColor("#ECF0F5")
    header_bg = colors.HexColor("#F2F5F9")
    month_bg = colors.HexColor("#EAF2FF")
    weekend_bg = colors.HexColor("#F8FAFD")
    bar_color = colors.HexColor("#9CC6F4")
    bar_edge = colors.HexColor("#5E9DE1")
    text = colors.HexColor("#1E2430")
    muted = colors.HexColor("#667085")
    month_text = colors.HexColor("#334155")
    weekend_text = colors.HexColor("#7B8797")

    pdf.setStrokeColor(border)
    pdf.setLineWidth(0.8)
    pdf.rect(chart_left, chart_bottom, chart_right - chart_left, chart_top - chart_bottom, stroke=1, fill=0)
    pdf.setFillColor(header_bg)
    pdf.rect(chart_left, header_bottom, label_width, month_height + date_height, stroke=0, fill=1)
    pdf.rect(axis_left, chart_top - month_height, axis_width, month_height, stroke=0, fill=1)

    pdf.setFillColor(text)
    pdf.setFont(bold_font, 9)
    pdf.drawString(chart_left + 8, header_bottom + 14, UI_TEXT["pdf_work_name"])

    month_start = 0
    current_month = (axis_days[0].year, axis_days[0].month)
    for index, day in enumerate(axis_days + [axis_days[-1] + timedelta(days=35)]):
        month = (day.year, day.month)
        if month == current_month and index < len(axis_days):
            continue
        x = axis_left + month_start * day_width
        month_width = (index - month_start) * day_width
        pdf.setFillColor(month_bg)
        pdf.rect(x, chart_top - month_height, month_width, month_height, stroke=0, fill=1)
        pdf.setFillColor(month_text)
        pdf.setFont(bold_font, 8)
        month_label = UI_TEXT["pdf_month_format"].format(year=current_month[0], month=current_month[1])
        pdf.drawCentredString(x + month_width / 2, chart_top - 13, month_label)
        pdf.setStrokeColor(border)
        pdf.line(x, chart_bottom, x, chart_top)
        month_start = index
        current_month = month

    for index, day in enumerate(axis_days):
        if is_weekday(day):
            continue
        x = axis_left + index * day_width
        pdf.setFillColor(weekend_bg)
        pdf.rect(x, chart_bottom, day_width, chart_top - chart_bottom - month_height, stroke=0, fill=1)

    pdf.setStrokeColor(border)
    pdf.line(axis_left, chart_bottom, axis_left, chart_top)
    pdf.line(chart_left, header_bottom, chart_right, header_bottom)
    pdf.line(axis_left, chart_top - month_height, chart_right, chart_top - month_height)
    pdf.line(chart_left, chart_top - month_height - date_height, chart_right, chart_top - month_height - date_height)

    weekday_labels = UI_TEXT["pdf_weekdays"]
    for index, day in enumerate(axis_days):
        x = axis_left + index * day_width
        if index % 5 == 0:
            pdf.setStrokeColor(border)
            pdf.setLineWidth(0.5)
        else:
            pdf.setStrokeColor(grid)
            pdf.setLineWidth(0.25)
        pdf.line(x, chart_bottom, x, chart_top - month_height)
        pdf.setFillColor(weekend_text if not is_weekday(day) else muted)
        pdf.setFont(font_name, 5.5)
        pdf.drawCentredString(x + day_width / 2, header_bottom + 14, f"{day.month}/{day.day}")
        pdf.setFont(font_name, 4.8)
        pdf.drawCentredString(x + day_width / 2, header_bottom + 5, weekday_labels[day.weekday()])
    pdf.setStrokeColor(border)
    pdf.line(chart_right, chart_bottom, chart_right, chart_top)

    day_index = {day: index for index, day in enumerate(axis_days)}
    valid_days = set(day_index)
    for row_index, row in enumerate(rows):
        y_top = header_bottom - row_index * row_height
        y_bottom = y_top - row_height
        y_mid = y_bottom + row_height / 2
        pdf.setStrokeColor(grid)
        pdf.setLineWidth(0.35)
        pdf.line(chart_left, y_bottom, chart_right, y_bottom)
        pdf.setFillColor(text)
        draw_fitted_text(pdf, row.name, chart_left + 8, y_mid - 3.2, label_width - 14, font_name, 8)

        clipped_start = max(row.start_date, axis_days[0])
        clipped_end = min(row.end_date, axis_days[-1])
        if clipped_start > clipped_end:
            continue

        segments = weekday_segments(clipped_start, clipped_end, valid_days)
        if not segments:
            continue
        for first_day, last_day in segments:
            bar_x = axis_left + day_index[first_day] * day_width + 1.2
            bar_width = (day_index[last_day] - day_index[first_day] + 1) * day_width - 2.4
            bar_height = min(13, row_height - 6)
            bar_y = y_mid - bar_height / 2
            pdf.setFillColor(bar_color)
            pdf.setStrokeColor(bar_edge)
            pdf.roundRect(bar_x, bar_y, max(2, bar_width), bar_height, 3, stroke=1, fill=1)

    pdf.showPage()
    pdf.save()


def validate_pdf_a3_landscape_one_page(path: Path) -> None:
    content = path.read_bytes()
    page_count = len(re.findall(rb"/Type\s*/Page\b", content))
    if page_count != 1:
        raise RuntimeError(f"page_count={page_count}")
    media_box = re.search(rb"/MediaBox\s*\[\s*0\s+0\s+([0-9.]+)\s+([0-9.]+)\s*\]", content)
    if not media_box:
        raise RuntimeError("MediaBox not found")
    width = float(media_box.group(1))
    height = float(media_box.group(2))
    if width <= height or abs(width - 1190.551) > 2 or abs(height - 841.89) > 2:
        raise RuntimeError(f"unexpected_size={width}x{height}")


class MansionScheduleApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(*WINDOW_MIN_SIZE)
        self.root.configure(bg=THEME["background"])
        self.font_family = choose_font_family(self.root)
        self.root.option_add("*Font", (self.font_family, 10))

        today = date.today()
        start_date, finish_date, rows = build_project_from_handover(today)
        saved_config = load_config()
        self.project_vars = {
            "site_name": tk.StringVar(value=""),
            "branch_name": tk.StringVar(value=saved_config.get("branch_name", "")),
            "person_name": tk.StringVar(value=saved_config.get("person_name", "")),
            "contact": tk.StringVar(value=saved_config.get("contact", "")),
            "handover_date": tk.StringVar(value=format_date(today)),
            "start_date": tk.StringVar(value=format_date(start_date)),
            "finish_date": tk.StringVar(value=format_date(finish_date)),
            "save_path": tk.StringVar(value=str(default_save_path())),
        }
        self.row_vars: list[dict[str, tk.StringVar]] = []
        for row in rows:
            self.row_vars.append(
                {
                    "name": tk.StringVar(value=row.name),
                    "start_date": tk.StringVar(value=format_date(row.start_date)),
                    "end_date": tk.StringVar(value=format_date(row.end_date)),
                }
            )
        self.status_var = tk.StringVar(value=UI_TEXT["status_ready"])
        self.footer_compact: bool | None = None

        self._apply_window_icon()
        self._build_styles()
        self._build_ui()

    def run(self) -> None:
        self.root.mainloop()

    def _build_styles(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(
            "Primary.TButton",
            font=(self.font_family, 10, "bold"),
            padding=(18, 9),
            background=THEME["accent"],
            foreground="#FFFFFF",
            bordercolor=THEME["accent"],
            lightcolor=THEME["accent"],
            darkcolor=THEME["accent"],
        )
        style.map(
            "Primary.TButton",
            background=[("active", THEME["accent_hover"]), ("disabled", "#A9C0F7")],
            foreground=[("disabled", "#FFFFFF")],
        )
        style.configure(
            "Secondary.TButton",
            font=(self.font_family, 10, "bold"),
            padding=(14, 9),
            background=THEME["panel"],
            foreground=THEME["text"],
            bordercolor=THEME["border"],
            lightcolor=THEME["panel"],
            darkcolor=THEME["panel"],
        )
        style.map("Secondary.TButton", background=[("active", THEME["subtle"])])

    def _build_ui(self) -> None:
        container = tk.Frame(self.root, bg=THEME["background"])
        container.pack(fill="both", expand=True, padx=24, pady=20)

        title = tk.Label(
            container,
            text=UI_TEXT["main_title"],
            bg=THEME["background"],
            fg=THEME["text"],
            font=(self.font_family, 20, "bold"),
            anchor="w",
        )
        title.pack(fill="x")

        description = tk.Label(
            container,
            text=UI_TEXT["main_description"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 10),
            anchor="w",
        )
        description.pack(fill="x", pady=(6, 14))

        self._build_project_panel(container)
        self._build_schedule_panel(container)
        self._build_actions(container)
        self._build_footer(container)

    def _build_project_panel(self, parent: tk.Frame) -> None:
        panel = self._panel(parent)
        panel.pack(fill="x")

        fields = (
            ("field_site_name", "site_name"),
            ("field_branch_name", "branch_name"),
            ("field_person_name", "person_name"),
            ("field_contact", "contact"),
            ("field_handover_date", "handover_date"),
            ("field_start_date", "start_date"),
            ("field_finish_date", "finish_date"),
            ("field_save_path", "save_path"),
        )
        for column in range(4):
            panel.grid_columnconfigure(column * 2 + 1, weight=1)

        for index, (label_key, var_key) in enumerate(fields):
            row = index // 4
            column = (index % 4) * 2
            tk.Label(
                panel,
                text=UI_TEXT[label_key],
                bg=THEME["panel"],
                fg=THEME["muted"],
                font=(self.font_family, 9, "bold"),
                anchor="w",
            ).grid(row=row, column=column, sticky="w", padx=(16 if column == 0 else 10, 6), pady=8)
            entry = ttk.Entry(panel, textvariable=self.project_vars[var_key])
            entry.grid(row=row, column=column + 1, sticky="ew", padx=(0, 12), pady=8)
            if var_key == "handover_date":
                entry.bind("<FocusOut>", self._on_handover_changed)
                entry.bind("<Return>", self._on_handover_changed)

    def _build_schedule_panel(self, parent: tk.Frame) -> None:
        panel = self._panel(parent)
        panel.pack(fill="both", expand=True, pady=(14, 0))

        for column, weight in enumerate((3, 1, 1)):
            panel.grid_columnconfigure(column, weight=weight)

        headers = ("table_work_name", "table_start_date", "table_end_date")
        for column, key in enumerate(headers):
            tk.Label(
                panel,
                text=UI_TEXT[key],
                bg=THEME["subtle"],
                fg=THEME["text"],
                font=(self.font_family, 9, "bold"),
                anchor="w",
                padx=8,
                pady=6,
            ).grid(row=0, column=column, sticky="ew", padx=(16 if column == 0 else 0, 16 if column == 2 else 0), pady=(14, 4))

        for index, vars_for_row in enumerate(self.row_vars, start=1):
            name_entry = ttk.Entry(panel, textvariable=vars_for_row["name"])
            start_entry = ttk.Entry(panel, textvariable=vars_for_row["start_date"])
            end_entry = ttk.Entry(panel, textvariable=vars_for_row["end_date"])
            name_entry.grid(row=index, column=0, sticky="ew", padx=(16, 8), pady=3)
            start_entry.grid(row=index, column=1, sticky="ew", padx=(0, 8), pady=3)
            end_entry.grid(row=index, column=2, sticky="ew", padx=(0, 16), pady=3)

    def _build_actions(self, parent: tk.Frame) -> None:
        actions = tk.Frame(parent, bg=THEME["background"])
        actions.pack(fill="x", pady=(14, 0))

        ttk.Button(
            actions,
            text=UI_TEXT["button_auto_schedule"],
            style="Secondary.TButton",
            command=self.auto_fill_schedule,
        ).pack(side="left", padx=(0, 10))

        ttk.Button(
            actions,
            text=UI_TEXT["button_select_save_path"],
            style="Secondary.TButton",
            command=self.select_save_path,
        ).pack(side="left")

        ttk.Button(
            actions,
            text=UI_TEXT["button_create_pdf"],
            style="Primary.TButton",
            command=self.create_pdf,
        ).pack(side="right")

        status_wrap = tk.Frame(parent, bg=THEME["background"])
        status_wrap.pack(fill="x", pady=(10, 0))
        tk.Label(
            status_wrap,
            text=UI_TEXT["status_label"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 9, "bold"),
        ).pack(side="left", padx=(0, 8))
        self.status_badge = tk.Label(
            status_wrap,
            textvariable=self.status_var,
            bg=THEME["subtle"],
            fg=THEME["muted"],
            font=(self.font_family, 9, "bold"),
            padx=10,
            pady=4,
        )
        self.status_badge.pack(side="left")

    def _build_footer(self, parent: tk.Frame) -> None:
        self.footer = tk.Frame(parent, bg=THEME["background"])
        self.footer.pack(fill="x", pady=(12, 0))

        self.footer_left = tk.Frame(self.footer, bg=THEME["background"])
        self._make_footer_text(self.footer_left, UI_TEXT["footer_left"])
        self._make_footer_text(self.footer_left, UI_TEXT["footer_separator"])
        self._make_footer_text(self.footer_left, UI_TEXT["footer_tagline"])

        self.footer_right = tk.Frame(self.footer, bg=THEME["background"])
        self._make_footer_link(self.footer_right, UI_TEXT["footer_link_1"], LINKS["assessment"])
        self._make_footer_text(self.footer_right, UI_TEXT["footer_separator"])
        self._make_footer_link(self.footer_right, UI_TEXT["footer_link_2"], LINKS["instagram"])
        self._make_footer_text(self.footer_right, UI_TEXT["footer_separator"])
        self._make_footer_text(self.footer_right, UI_TEXT["footer_copyright"])

        self.root.bind("<Configure>", self._update_footer_layout, add="+")
        self._update_footer_layout()

    def _make_footer_link(self, parent: tk.Frame, label: str, url: str) -> None:
        link = tk.Label(
            parent,
            text=label,
            bg=THEME["background"],
            fg=THEME["link"],
            font=(self.font_family, 8, "bold"),
            cursor="hand2",
        )
        link.pack(side="left")
        link.bind("<Button-1>", lambda _event: webbrowser.open(url, new=2))
        link.bind("<Enter>", lambda _event: link.configure(fg=THEME["link_hover"]))
        link.bind("<Leave>", lambda _event: link.configure(fg=THEME["link"]))

    def _make_footer_text(self, parent: tk.Frame, label: str) -> None:
        text = tk.Label(
            parent,
            text=label,
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 8),
        )
        text.pack(side="left")

    def _update_footer_layout(self, _event=None) -> None:
        compact = self.root.winfo_width() < FOOTER_COMPACT_WIDTH
        if compact == self.footer_compact:
            return

        self.footer_compact = compact
        self.footer_left.pack_forget()
        self.footer_right.pack_forget()
        if compact:
            self.footer_left.pack(anchor="center", pady=(0, 2))
            self.footer_right.pack(anchor="center")
            return

        self.footer_left.pack(side="left", fill="x", expand=True)
        self.footer_right.pack(side="right")

    def _panel(self, parent: tk.Frame) -> tk.Frame:
        return tk.Frame(
            parent,
            bg=THEME["panel"],
            highlightbackground=THEME["border"],
            highlightthickness=1,
        )

    def _apply_window_icon(self) -> None:
        for icon_path in get_common_icon_candidates():
            try:
                resolved = icon_path.resolve()
            except OSError:
                resolved = icon_path
            if resolved.exists():
                try:
                    self.root.iconbitmap(str(resolved))
                except tk.TclError:
                    pass
                return

    def _on_handover_changed(self, _event=None) -> None:
        try:
            self.auto_fill_schedule(show_errors=False)
        except ValueError:
            return

    def _set_status(self, message: str, error: bool = False, success: bool = False) -> None:
        self.status_var.set(message)
        if error:
            self.status_badge.configure(bg=THEME["error_bg"], fg=THEME["error"])
        elif success:
            self.status_badge.configure(bg=THEME["success_bg"], fg=THEME["success"])
        else:
            self.status_badge.configure(bg=THEME["subtle"], fg=THEME["muted"])

    def auto_fill_schedule(self, show_errors: bool = True) -> None:
        try:
            handover_date = parse_date(self.project_vars["handover_date"].get(), UI_TEXT["field_handover_date"])
        except ValueError as exc:
            if show_errors:
                self._show_error(str(exc))
            raise

        start_date, finish_date, rows = build_project_from_handover(handover_date)
        self.project_vars["start_date"].set(format_date(start_date))
        self.project_vars["finish_date"].set(format_date(finish_date))
        if not self.project_vars["save_path"].get().strip():
            self.project_vars["save_path"].set(str(default_save_path(self.project_vars["site_name"].get())))

        for vars_for_row, row in zip(self.row_vars, rows):
            vars_for_row["name"].set(row.name)
            vars_for_row["start_date"].set(format_date(row.start_date))
            vars_for_row["end_date"].set(format_date(row.end_date))

        self._set_status(UI_TEXT["status_auto_done"], success=True)

    def select_save_path(self) -> None:
        initial_path = Path(self.project_vars["save_path"].get() or default_save_path(self.project_vars["site_name"].get()))
        filename = filedialog.asksaveasfilename(
            title=UI_TEXT["button_select_save_path"],
            initialdir=str(initial_path.parent if initial_path.parent.exists() else get_base_dir()),
            initialfile=initial_path.name,
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
        )
        if filename:
            self.project_vars["save_path"].set(filename)

    def collect_project_info(self) -> ProjectInfo:
        handover_date = parse_date(self.project_vars["handover_date"].get(), UI_TEXT["field_handover_date"])
        start_date = parse_date(self.project_vars["start_date"].get(), UI_TEXT["field_start_date"])
        finish_date = parse_date(self.project_vars["finish_date"].get(), UI_TEXT["field_finish_date"])
        if start_date > finish_date:
            raise ValueError(
                UI_TEXT["error_date_order"].format(
                    start_label=UI_TEXT["field_start_date"],
                    end_label=UI_TEXT["field_finish_date"],
                )
            )
        return ProjectInfo(
            site_name=self.project_vars["site_name"].get().strip(),
            branch_name=self.project_vars["branch_name"].get().strip(),
            person_name=self.project_vars["person_name"].get().strip(),
            contact=self.project_vars["contact"].get().strip(),
            handover_date=handover_date,
            start_date=start_date,
            finish_date=finish_date,
        )

    def collect_schedule_rows(self) -> list[ScheduleRow]:
        rows: list[ScheduleRow] = []
        for vars_for_row in self.row_vars:
            name = vars_for_row["name"].get().strip()
            start_date = parse_date(vars_for_row["start_date"].get(), UI_TEXT["table_start_date"])
            end_date = parse_date(vars_for_row["end_date"].get(), UI_TEXT["table_end_date"])
            if start_date > end_date:
                raise ValueError(
                    UI_TEXT["error_date_order"].format(
                        start_label=UI_TEXT["table_start_date"],
                        end_label=UI_TEXT["table_end_date"],
                    )
                )
            rows.append(ScheduleRow(name=name, start_date=start_date, end_date=end_date))
        return rows

    def create_pdf(self) -> None:
        save_path_value = self.project_vars["save_path"].get().strip()
        if not save_path_value:
            self._show_error(UI_TEXT["dialog_save_cancelled"])
            return
        output_path = Path(save_path_value)
        try:
            info = self.collect_project_info()
            rows = self.collect_schedule_rows()
            draw_schedule_pdf(output_path, info, rows)
        except Exception as exc:
            self._show_error(str(exc))
            return

        save_config(
            {
                "branch_name": self.project_vars["branch_name"].get(),
                "person_name": self.project_vars["person_name"].get(),
                "contact": self.project_vars["contact"].get(),
            }
        )
        folder_opened = open_output_folder(output_path)
        if folder_opened:
            self._set_status(UI_TEXT["status_pdf_done"].format(path=output_path), success=True)
        else:
            self._set_status(UI_TEXT["status_pdf_done_folder_open_failed"].format(path=output_path), success=True)
        messagebox.showinfo(UI_TEXT["dialog_done_title"], UI_TEXT["dialog_done_message"].format(path=output_path))

    def _show_error(self, message: str) -> None:
        self._set_status(UI_TEXT["status_error"], error=True)
        messagebox.showerror(UI_TEXT["dialog_error_title"], message)


def run_launch_check() -> int:
    handover_date = date(2026, 5, 29)
    start_date, finish_date, rows = build_project_from_handover(handover_date)
    if start_date != date(2026, 6, 1):
        raise RuntimeError("start date fixture failed")
    if finish_date != date(2026, 7, 15):
        raise RuntimeError("finish date fixture failed")
    if len(rows) != len(WORK_ITEMS):
        raise RuntimeError("work item fixture failed")
    if sum(duration for _name, duration in WORK_ITEMS) != TOTAL_INITIAL_WORKDAYS:
        raise RuntimeError("duration fixture failed")
    if rows[0].start_date != start_date or rows[-1].end_date != finish_date:
        raise RuntimeError("schedule range fixture failed")
    axis_days = days_between(start_date, finish_date)
    axis_workdays = workdays_between(start_date, finish_date)
    if len(axis_days) != TOTAL_CALENDAR_DAYS:
        raise RuntimeError("calendar day fixture failed")
    if not any(not is_weekday(day) for day in axis_days):
        raise RuntimeError("weekend fixture failed")
    if any(not is_weekday(row.start_date) or not is_weekday(row.end_date) for row in rows):
        raise RuntimeError("schedule weekday edge fixture failed")
    allocated_workdays = sum(len(workdays_between(row.start_date, row.end_date)) for row in rows)
    if allocated_workdays != TOTAL_INITIAL_WORKDAYS:
        raise RuntimeError("initial weekday allocation fixture failed")
    reserve_workdays = len(axis_workdays) - allocated_workdays
    if reserve_workdays < 0:
        raise RuntimeError("reserve weekday fixture failed")
    valid_days = set(axis_days)
    for row in rows:
        for segment_start, segment_finish in weekday_segments(row.start_date, row.end_date, valid_days):
            if any(not is_weekday(day) for day in days_between(segment_start, segment_finish)):
                raise RuntimeError("weekend bar fixture failed")

    info = ProjectInfo(
        site_name="テストマンション101",
        branch_name="テスト支店",
        person_name="山田太郎",
        contact="03-0000-0000",
        handover_date=handover_date,
        start_date=start_date,
        finish_date=finish_date,
    )

    with tempfile.TemporaryDirectory(dir=get_base_dir()) as temp_dir:
        output_path = Path(temp_dir) / PDF_FILE_NAME
        draw_schedule_pdf(output_path, info, rows)
        validate_pdf_a3_landscape_one_page(output_path)
        content = output_path.read_bytes()
        for forbidden in (b"Vibe-Coded", b"Yukihiko", b"pdf_note"):
            if forbidden in content:
                raise RuntimeError("submission footer fixture failed")

        config_path = Path(temp_dir) / CONFIG_FILE_NAME
        saved_values = {
            "branch_name": "テスト支店",
            "person_name": "山田太郎",
            "contact": "03-0000-0000",
        }
        if not save_config(saved_values, config_path):
            raise RuntimeError("config save fixture failed")
        if load_config(config_path) != saved_values:
            raise RuntimeError("config restore fixture failed")
        config_path.write_text("{", encoding="utf-8")
        if load_config(config_path):
            raise RuntimeError("broken config fixture failed")

        opened: list[str] = []
        if not open_output_folder(output_path, opener=opened.append):
            raise RuntimeError("open output folder fixture failed")
        if not opened or Path(opened[0]) != output_path.parent.resolve():
            raise RuntimeError("open output folder path fixture failed")

        print(UI_TEXT["launch_check_ok"])
        print(UI_TEXT["launch_check_pdf"].format(path=output_path))
        print(UI_TEXT["launch_check_page"])
        print(UI_TEXT["launch_check_span"].format(days=len(axis_days), start=format_date(start_date), finish=format_date(finish_date)))
        print(UI_TEXT["launch_check_workdays"].format(workdays=len(axis_workdays), allocated=allocated_workdays))
        print(UI_TEXT["launch_check_reserve"].format(reserve=reserve_workdays))
        print(UI_TEXT["launch_check_weekend_bars"])
        print(UI_TEXT["launch_check_config"])
        print(UI_TEXT["launch_check_folder"])
        print(UI_TEXT["launch_check_footer"])
    return 0


def main() -> None:
    if "--launch-check" in sys.argv:
        raise SystemExit(run_launch_check())
    app = MansionScheduleApp()
    app.run()


if __name__ == "__main__":
    main()
