# -*- coding: utf-8 -*-
from __future__ import annotations

import ctypes
import json
import os
import queue
import re
import subprocess
import sys
import tempfile
import threading
import time
import webbrowser
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, ttk


APP_NAME = "Dakeリフォーム進捗管理"
WINDOW_TITLE = "Dakeリフォーム進捗管理"
COPYRIGHT = "© 2026 しまりす不動産 / Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "main_title": "リフォーム進捗カレンダーを作る",
    "main_description": "進捗に合わせて工程を直し、A4縦1枚の最新版PDFにします。",
    "label_site_name": "現場名",
    "label_branch_name": "支店名",
    "label_staff_name": "担当者名",
    "label_phone": "電話番号",
    "label_start_date": "開始日",
    "label_end_date": "終了日",
    "label_save_folder": "保存先",
    "date_hint": "例：2026/05/29",
    "reference_day_template": "45日目：{date}",
    "reference_day_invalid": "45日目：-",
    "button_select_folder": "選択",
    "button_save_project": "状態保存",
    "button_create_pdf": "PDF作成",
    "button_reschedule": "日程再配置",
    "button_today": "今日へ移動",
    "button_up": "↑",
    "button_down": "↓",
    "section_work_list": "工程リスト",
    "section_preview": "カレンダープレビュー",
    "table_use": "使用",
    "table_work_name": "工程名",
    "table_start_date": "開始日",
    "table_end_date": "終了日",
    "table_status": "状態",
    "table_days": "日数",
    "table_order": "順番",
    "status_ready": "入力してPDF作成を押してください。",
    "status_processing": "PDF作成中...",
    "status_complete": "PDFを保存しました。",
    "status_error": "処理に失敗しました。",
    "status_rescheduled": "日程を再配置しました。",
    "status_moved_today": "開始日を今日へ移動しました。",
    "status_weekend_adjusted": "土日は工事日に含めないため、{from_date}から{to_date}へ移動しました。",
    "status_drag_moved": "{name}を{date}開始へ移動しました。",
    "status_drag_resized": "{name}の終了日を{date}へ変更しました。",
    "status_drag_start_resized": "{name}の開始日を{date}へ変更しました。",
    "status_drag_end_resized": "{name}の終了日を{date}へ変更しました。",
    "status_dragging_start": "開始日を変更中：{date}",
    "status_dragging_end": "終了日を変更中：{date}",
    "status_dragging_move": "工程を移動中：{date}へ移動",
    "status_row_moved": "工程の順番を変更しました。",
    "status_config_saved": "設定を保存しました。",
    "status_project_loaded": "前回保存した工程状態を読み込みました。",
    "status_project_saved": "工程状態を保存しました。",
    "status_project_save_failed": "工程状態を保存できませんでした。",
    "status_pdf_done_folder_open_failed": "PDFを保存しました。保存フォルダは手動で確認してください。",
    "dialog_error_title": "確認してください",
    "dialog_saved_title": "保存完了",
    "dialog_saved_message": "リフォーム進捗カレンダーPDFを保存しました。\n\n{path}",
    "dialog_select_save_dir": "保存先フォルダを選択",
    "error_start_date": "開始日を正しく入力してください。",
    "error_end_date": "終了日を正しく入力してください。",
    "error_date_order": "終了日は開始日以降の日付にしてください。",
    "error_date_too_long": "A4縦1枚で読める範囲を超えています。終了日は開始日から56日以内にしてください。",
    "error_save_folder": "保存先フォルダを選択してください。",
    "error_no_work": "PDFに出力できる工程がありません。",
    "error_reportlab_missing": "PDF生成に必要な reportlab が見つかりません。",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_tagline": "止まらない、迷わない、すぐ終わる。",
    "footer_link_1": "戸建買取査定",
    "footer_link_2": "Instagram",
    "footer_separator": " ｜ ",
    "footer_copyright": COPYRIGHT,
    "pdf_title": "リフォーム進捗カレンダー",
    "pdf_site_placeholder": "現場名未入力",
    "pdf_target_period": "対象期間",
    "pdf_day_count": "日数",
    "pdf_day_count_template": "{count}日",
    "pdf_contact_template": "さくら都市{branch} 担当：{staff} {phone}",
    "pdf_completion_label": "完工",
    "pdf_weekday_headers": ["日", "月", "火", "水", "木", "金", "土"],
    "pdf_weekday_names": ["月", "火", "水", "木", "金", "土", "日"],
    "pdf_full_date_template": "{year}年{month}月{day}日",
    "pdf_range_separator": "〜",
    "launch_check_ok": "Dakeリフォーム進捗管理 launch-check OK",
    "launch_check_pdf": "pdf={path}",
    "launch_check_page": "a4_portrait_one_page=OK",
    "launch_check_span": "calendar_days={days} start={start} finish={finish} weeks={weeks}",
    "launch_check_completion": "completion_cell=2026-07-13 完工",
    "launch_check_weekend": "weekend_background_and_no_work_items=OK",
    "launch_check_drag": "calendar_drag_move_and_resize_logic=OK",
    "launch_check_rows": "row_reorder_logic=OK",
    "launch_check_reschedule": "reschedule_done_items_fixed=OK",
    "launch_check_config": "config_save_restore=OK",
    "launch_check_project": "project_save_restore=OK",
    "launch_check_pdf_bands": "pdf_multi_day_bands=OK",
    "launch_check_reference": "reference_day_45=OK",
    "launch_check_limit": "max_56_day_limit=OK",
    "launch_check_folder": "open_output_folder=OK",
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

CONFIG_NAME = "reform_progress_config.json"
PROJECT_NAME = "reform_progress_project.json"
COMMON_ICON_RELATIVE = Path("..") / ".." / "02_assets" / "dake_icon.ico"
COMMON_ICON_FILENAME = "dake_icon.ico"
DATE_FORMATS = ("%Y/%m/%d", "%Y-%m-%d", "%Y.%m.%d", "%Y%m%d")
FONT_CANDIDATES = ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo")
REFERENCE_CALENDAR_DAYS = 45
DEFAULT_CALENDAR_DAYS = 45
MAX_CALENDAR_DAYS = 56
STATUS_OPTIONS = ("未着手", "作業中", "完了", "延期", "保留")
STATUS_SYMBOLS = {
    "未着手": "",
    "作業中": "●",
    "完了": "✓",
    "延期": "△",
    "保留": "…",
}
STATUS_COLORS = {
    "未着手": "#E8EDF5",
    "作業中": "#E8F2FF",
    "完了": "#EAF6EE",
    "延期": "#FFF3E8",
    "保留": "#F1EEF8",
}
WORK_TEMPLATES: tuple[tuple[str, int], ...] = (
    ("養生", 1),
    ("残置物撤去", 2),
    ("電気工事", 1),
    ("水道設備解体", 2),
    ("クロス工事（剥がし）", 2),
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
FREE_WORK_COUNT = 3
DAYS_PER_ROW = 7
FOOTER_COMPACT_WIDTH = 1280
LINKS = {
    "assessment": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "instagram": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

THEME = {
    "background": "#F6F7F9",
    "panel": "#FFFFFF",
    "subtle": "#EEF2F7",
    "border": "#D8DEE8",
    "grid": "#E5EAF2",
    "band_outline": "#D4DDEB",
    "drag_highlight": "#EAF2FF",
    "grip": "#6B7C93",
    "row_alt": "#FAFBFD",
    "text": "#1E2430",
    "muted": "#667085",
    "accent": "#2F6FED",
    "accent_hover": "#2458BF",
    "accent_disabled": "#9BAFD6",
    "success": "#118A4E",
    "error": "#C92A2A",
    "weekend_bg": "#FFF7F7",
    "outside_bg": "#F2F4F7",
    "weekend_text": "#C76A6A",
    "holiday_text": "#B94040",
    "link": "#58677D",
}


@dataclass
class WorkItem:
    selected: bool
    name: str
    start_date: date | None
    end_date: date | None
    status: str = STATUS_OPTIONS[0]
    planned_days: int = 1
    is_free: bool = False


@dataclass(frozen=True)
class PdfRequest:
    site_name: str
    branch_name: str
    staff_name: str
    phone: str
    start_date: date
    end_date: date
    save_folder: str
    work_items: tuple[WorkItem, ...]


@dataclass
class RowWidgets:
    frame: tk.Frame
    selected: tk.BooleanVar
    name: tk.StringVar
    start_date: tk.StringVar
    end_date: tk.StringVar
    status: tk.StringVar
    days: tk.StringVar


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def config_path() -> Path:
    return app_dir() / CONFIG_NAME


def project_path() -> Path:
    return app_dir() / PROJECT_NAME


def default_save_folder() -> str:
    downloads = Path.home() / "Downloads"
    if downloads.exists():
        return str(downloads)
    return str(Path.home())


def load_config(path: Path | None = None) -> dict:
    target = path or config_path()
    try:
        data = json.loads(target.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    if not isinstance(data, dict):
        return {}
    allowed = {"branch_name", "staff_name", "phone", "save_folder"}
    return {key: value for key, value in data.items() if key in allowed and isinstance(value, str)}


def save_config(values: dict[str, str], path: Path | None = None) -> bool:
    target = path or config_path()
    data = {
        "branch_name": values.get("branch_name", "").strip(),
        "staff_name": values.get("staff_name", "").strip(),
        "phone": values.get("phone", "").strip(),
        "save_folder": values.get("save_folder", "").strip(),
    }
    try:
        target.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        return False
    return True


def load_project(path: Path | None = None) -> dict:
    target = path or project_path()
    try:
        data = json.loads(target.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    if not isinstance(data, dict):
        return {}
    return data


def save_project(values: dict, path: Path | None = None) -> bool:
    target = path or project_path()
    try:
        target.write_text(json.dumps(values, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        return False
    return True


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
        try:
            resolved = candidate.resolve()
        except OSError:
            resolved = candidate
        if resolved.exists():
            return resolved
    return None


def apply_window_icon(window: tk.Tk) -> None:
    icon_path = find_icon_path()
    if icon_path is None:
        return
    try:
        window.iconbitmap(default=str(icon_path))
    except tk.TclError:
        try:
            window.iconbitmap(str(icon_path))
        except tk.TclError:
            pass


    try:
        icon_photo = tk.PhotoImage(file=str(icon_path))
        window.iconphoto(True, icon_photo)
        window._dake_icon_photo = icon_photo  # type: ignore[attr-defined]
    except tk.TclError:
        pass


def set_windows_app_id() -> None:
    if not sys.platform.startswith("win"):
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Shimarisu.DakeReformProgress")
    except Exception:
        pass


def open_output_folder(output_path: Path, opener=None) -> bool:
    folder = output_path.parent.resolve()
    if opener is not None:
        try:
            opener(str(folder))
        except OSError:
            return False
        return True
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(folder))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(folder)])
        else:
            subprocess.Popen(["xdg-open", str(folder)])
    except Exception:
        return False
    return True


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


def format_date(value: date | None) -> str:
    if value is None:
        return ""
    return value.strftime("%Y/%m/%d")


def format_full_date(value: date) -> str:
    return UI_TEXT["pdf_full_date_template"].format(year=value.year, month=value.month, day=value.day)


def format_short_date(value: date) -> str:
    return value.strftime("%Y/%m/%d")


def sanitize_filename(value: str) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|]+', "_", value).strip()
    cleaned = re.sub(r"\s+", "_", cleaned)
    return cleaned[:60]


def is_weekday(value: date) -> bool:
    return value.weekday() < 5


def snap_to_workday(value: date) -> tuple[date, bool]:
    if value.weekday() == 5:
        return value - timedelta(days=1), True
    if value.weekday() == 6:
        return value + timedelta(days=1), True
    return value, False


def next_workday_after(value: date) -> date:
    current = value + timedelta(days=1)
    while not is_weekday(current):
        current += timedelta(days=1)
    return current


def end_after_workdays(start: date, count: int) -> date:
    count = max(1, count)
    current, _adjusted = snap_to_workday(start)
    remaining = count - 1
    while remaining > 0:
        current += timedelta(days=1)
        if is_weekday(current):
            remaining -= 1
    return current


def days_between(start: date, end: date) -> list[date]:
    if end < start:
        return []
    return [start + timedelta(days=index) for index in range((end - start).days + 1)]


def calendar_day_count(start: date, end: date) -> int:
    if end < start:
        return 0
    return (end - start).days + 1


def reference_day_for_start(start: date) -> date:
    return start + timedelta(days=REFERENCE_CALENDAR_DAYS - 1)


def is_period_within_limit(start: date, end: date) -> bool:
    return calendar_day_count(start, end) <= MAX_CALENDAR_DAYS


def workdays_between(start: date, end: date) -> list[date]:
    return [day for day in days_between(start, end) if is_weekday(day)]


def item_duration(item: WorkItem) -> int:
    if item.start_date and item.end_date and item.end_date >= item.start_date:
        days = len(workdays_between(item.start_date, item.end_date))
        if days > 0:
            return days
    return max(1, item.planned_days)


def is_output_item(item: WorkItem) -> bool:
    if not item.selected:
        return False
    if not item.name.strip():
        return False
    if item.start_date is None or item.end_date is None:
        return False
    return item.end_date >= item.start_date


def copy_work_item(item: WorkItem) -> WorkItem:
    return WorkItem(
        selected=item.selected,
        name=item.name,
        start_date=item.start_date,
        end_date=item.end_date,
        status=item.status,
        planned_days=item.planned_days,
        is_free=item.is_free,
    )


def serialize_work_items(items: list[WorkItem] | tuple[WorkItem, ...]) -> list[dict[str, object]]:
    serialized: list[dict[str, object]] = []
    for index, item in enumerate(items):
        serialized.append(
            {
                "order": index + 1,
                "selected": bool(item.selected),
                "name": item.name.strip(),
                "start_date": format_date(item.start_date),
                "end_date": format_date(item.end_date),
                "status": item.status if item.status in STATUS_OPTIONS else STATUS_OPTIONS[0],
                "planned_days": max(1, int(item.planned_days)),
                "is_free": bool(item.is_free),
            }
        )
    return serialized


def deserialize_work_items(value: object) -> list[WorkItem] | None:
    if not isinstance(value, list):
        return None
    rows: list[tuple[int, int, WorkItem]] = []
    for position, entry in enumerate(value):
        if not isinstance(entry, dict):
            continue
        raw_order = entry.get("order", position + 1)
        try:
            order = int(raw_order)
        except (TypeError, ValueError):
            order = position + 1
        name = str(entry.get("name", "")).strip()
        start_date = parse_date_text(str(entry.get("start_date", "")))
        end_date = parse_date_text(str(entry.get("end_date", "")))
        status = str(entry.get("status", STATUS_OPTIONS[0]))
        if status not in STATUS_OPTIONS:
            status = STATUS_OPTIONS[0]
        try:
            planned_days = max(1, int(entry.get("planned_days", 1)))
        except (TypeError, ValueError):
            planned_days = 1
        item = WorkItem(
            selected=bool(entry.get("selected", False)),
            name=name,
            start_date=start_date,
            end_date=end_date,
            status=status,
            planned_days=planned_days,
            is_free=bool(entry.get("is_free", False)),
        )
        if item.start_date and item.end_date and item.end_date >= item.start_date:
            item.planned_days = item_duration(item)
        rows.append((order, position, item))
    if not rows:
        return None
    return [item for _order, _position, item in sorted(rows, key=lambda row: (row[0], row[1]))]


def build_project_data(
    site_name: str,
    branch_name: str,
    staff_name: str,
    phone: str,
    start_date_text: str,
    end_date_text: str,
    save_folder: str,
    work_items: list[WorkItem] | tuple[WorkItem, ...],
) -> dict[str, object]:
    return {
        "version": "1.1",
        "site_name": site_name.strip(),
        "branch_name": branch_name.strip(),
        "staff_name": staff_name.strip(),
        "phone": phone.strip(),
        "start_date": start_date_text.strip(),
        "end_date": end_date_text.strip(),
        "save_folder": save_folder.strip(),
        "work_items": serialize_work_items(work_items),
    }


def place_initial_work_items(start_date: date) -> list[WorkItem]:
    current, _adjusted = snap_to_workday(start_date)
    items: list[WorkItem] = []
    for name, duration in WORK_TEMPLATES:
        item_start = current
        item_end = end_after_workdays(item_start, duration)
        items.append(
            WorkItem(
                selected=True,
                name=name,
                start_date=item_start,
                end_date=item_end,
                status=STATUS_OPTIONS[0],
                planned_days=duration,
                is_free=False,
            )
        )
        current = next_workday_after(item_end)
    for _index in range(FREE_WORK_COUNT):
        items.append(
            WorkItem(
                selected=False,
                name="",
                start_date=None,
                end_date=None,
                status=STATUS_OPTIONS[0],
                planned_days=1,
                is_free=True,
            )
        )
    return items


def reschedule_work_items(items: list[WorkItem], start_date: date) -> list[WorkItem]:
    current, _adjusted = snap_to_workday(start_date)
    updated = [copy_work_item(item) for item in items]
    for item in updated:
        if not item.selected:
            continue
        duration = item_duration(item)
        if item.status == "完了" and item.start_date and item.end_date:
            if item.end_date >= current:
                current = next_workday_after(item.end_date)
            continue
        item.start_date = current
        item.end_date = end_after_workdays(current, duration)
        item.planned_days = duration
        current = next_workday_after(item.end_date)
    return updated


def move_item_to_start(item: WorkItem, start_date: date) -> tuple[WorkItem, bool, date]:
    snapped, adjusted = snap_to_workday(start_date)
    updated = copy_work_item(item)
    duration = item_duration(updated)
    updated.start_date = snapped
    updated.end_date = end_after_workdays(snapped, duration)
    updated.planned_days = duration
    return updated, adjusted, snapped


def resize_item_to_end(item: WorkItem, end_date: date) -> tuple[WorkItem, bool, date]:
    updated = copy_work_item(item)
    if updated.start_date is None:
        start, start_adjusted = snap_to_workday(end_date)
        updated.start_date = start
    snapped, adjusted = snap_to_workday(end_date)
    if updated.start_date and snapped < updated.start_date:
        snapped = updated.start_date
    updated.end_date = snapped
    updated.planned_days = item_duration(updated)
    return updated, adjusted, snapped


def resize_item_to_start(item: WorkItem, start_date: date) -> tuple[WorkItem, bool, date]:
    updated = copy_work_item(item)
    if updated.end_date is None:
        end, _end_adjusted = snap_to_workday(start_date)
        updated.end_date = end
    snapped, adjusted = snap_to_workday(start_date)
    if updated.end_date and snapped > updated.end_date:
        snapped = updated.end_date
    updated.start_date = snapped
    updated.planned_days = item_duration(updated)
    return updated, adjusted, snapped


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
    return current + timedelta(days=(0 - current.weekday()) % 7 + 7 * (nth - 1))


def spring_equinox_day(year: int) -> int:
    return int(20.8431 + 0.242194 * (year - 1980) - ((year - 1980) // 4))


def autumn_equinox_day(year: int) -> int:
    return int(23.2488 + 0.242194 * (year - 1980) - ((year - 1980) // 4))


def fallback_japanese_holidays(years: set[int]) -> dict[date, str]:
    names = UI_TEXT["holiday_names"]
    holidays: dict[date, str] = {}
    for year in years:
        base: dict[date, str] = {
            date(year, 1, 1): names["new_year"],
            nth_monday(year, 1, 2): names["coming_age"],
            date(year, 2, 11): names["foundation"],
            date(year, 3, spring_equinox_day(year)): names["vernal"],
            date(year, 5, 3): names["constitution"],
            date(year, 5, 5): names["children"],
            date(year, 9, autumn_equinox_day(year)): names["autumnal"],
            date(year, 11, 3): names["culture"],
            date(year, 11, 23): names["labor"],
        }
        if year >= 2020:
            base[date(year, 2, 23)] = names["emperor"]
        elif 1989 <= year <= 2018:
            base[date(year, 12, 23)] = names["emperor"]
        if year >= 2007:
            base[date(year, 4, 29)] = names["showa"]
            base[date(year, 5, 4)] = names["greenery"]
        elif 1989 <= year <= 2006:
            base[date(year, 4, 29)] = names["greenery"]
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

        year_holidays = dict(base)
        current = date(year, 1, 2)
        while current <= date(year, 12, 30):
            if current not in year_holidays and current - timedelta(days=1) in year_holidays and current + timedelta(days=1) in year_holidays:
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
    try:
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    except ModuleNotFoundError as exc:
        raise RuntimeError(UI_TEXT["error_reportlab_missing"]) from exc
    font_name = "HeiseiKakuGo-W5"
    try:
        pdfmetrics.registerFont(UnicodeCIDFont(font_name))
    except Exception:
        pass
    return font_name


def color_hex(hex_value: str):
    from reportlab.lib import colors

    return colors.HexColor(hex_value)


def ellipsize_text(text: str, max_width: float, font_name: str, size: float) -> str:
    from reportlab.pdfbase import pdfmetrics

    if pdfmetrics.stringWidth(text, font_name, size) <= max_width:
        return text
    marker = "…"
    for length in range(max(0, len(text) - 1), 0, -1):
        candidate = text[:length] + marker
        if pdfmetrics.stringWidth(candidate, font_name, size) <= max_width:
            return candidate
    return marker


def daily_entries(day: date, items: list[WorkItem] | tuple[WorkItem, ...]) -> list[str]:
    if not is_weekday(day):
        return []
    entries: list[str] = []
    for item in items:
        if not is_output_item(item):
            continue
        assert item.start_date is not None
        assert item.end_date is not None
        if item.start_date <= day <= item.end_date:
            symbol = STATUS_SYMBOLS.get(item.status, "")
            entries.append(f"{symbol}{short_work_name(item.name)}")
    return entries


def short_work_name(name: str) -> str:
    replacements = {
        "残置物撤去": "残置撤去",
        "電気工事": "電気",
        "水道設備解体": "水道解体",
        "クロス工事（剥がし）": "クロス剥がし",
        "ユニットバス設置工事": "UB設置",
        "システムキッチン設置工事": "キッチン設置",
        "水道設備設置工事": "水道設置",
        "ハウスクリーニング": "クリーニング",
    }
    return replacements.get(name.strip(), name.strip())


def visible_workdays_for_item(item: WorkItem, period_start: date | None = None, period_end: date | None = None) -> list[date]:
    if not is_output_item(item):
        return []
    assert item.start_date is not None
    assert item.end_date is not None
    start = max(item.start_date, period_start) if period_start else item.start_date
    end = min(item.end_date, period_end) if period_end else item.end_date
    return workdays_between(start, end)


def band_segments_for_item(item: WorkItem, period_start: date, period_end: date) -> list[list[date]]:
    days = visible_workdays_for_item(item, period_start, period_end)
    segments: list[list[date]] = []
    current: list[date] = []
    previous: date | None = None
    for day in days:
        if previous is not None and (day - previous).days != 1:
            segments.append(current)
            current = []
        current.append(day)
        previous = day
    if current:
        segments.append(current)
    return segments


def build_work_lanes(items: list[WorkItem] | tuple[WorkItem, ...]) -> dict[int, int]:
    lanes: list[set[date]] = []
    lane_by_index: dict[int, int] = {}
    for index, item in enumerate(items):
        days = set(visible_workdays_for_item(item))
        if not days:
            continue
        for lane_index, occupied in enumerate(lanes):
            if occupied.isdisjoint(days):
                occupied.update(days)
                lane_by_index[index] = lane_index
                break
        else:
            lanes.append(set(days))
            lane_by_index[index] = len(lanes) - 1
    return lane_by_index


def max_work_lanes(cell_h: float) -> int:
    return max(1, int((cell_h - 34) // 9))


def draw_pdf_band(pdf, x: float, y: float, width: float, height: float, fill_hex: str, left_round: bool, right_round: bool) -> None:
    radius = min(height / 2, 4)
    pdf.setFillColor(color_hex(fill_hex))
    if left_round or right_round:
        pdf.roundRect(x, y, width, height, radius, stroke=0, fill=1)
        if not left_round:
            pdf.rect(x, y, radius, height, stroke=0, fill=1)
        if not right_round:
            pdf.rect(x + width - radius, y, radius, height, stroke=0, fill=1)
    else:
        pdf.rect(x, y, width, height, stroke=0, fill=1)
    pdf.setStrokeColor(color_hex(THEME["band_outline"]))
    pdf.setLineWidth(0.25)
    if left_round and right_round:
        pdf.roundRect(x, y, width, height, radius, stroke=1, fill=0)
    else:
        pdf.rect(x, y, width, height, stroke=1, fill=0)


def draw_pdf_work_items(
    pdf,
    items: tuple[WorkItem, ...],
    date_boxes: dict[date, tuple[float, float, float, float]],
    cell_h: float,
    font_name: str,
    period_start: date,
    period_end: date,
) -> None:
    lane_by_index = build_work_lanes(items)
    overflow_days: set[date] = set()
    lane_step = 9.2
    band_h = 8.0
    for index, item in enumerate(items):
        days = visible_workdays_for_item(item, period_start, period_end)
        if not days:
            continue
        lane = lane_by_index.get(index, 0)
        label = f"{STATUS_SYMBOLS.get(item.status, '')}{short_work_name(item.name)}"
        workday_total = len(visible_workdays_for_item(item))
        if workday_total <= 1:
            day = days[0]
            box = date_boxes.get(day)
            if box is None:
                continue
            x0, y0, x1, _y1 = box
            if lane >= max_work_lanes(cell_h):
                overflow_days.add(day)
                continue
            text_y = y0 + cell_h - 24 - lane * lane_step
            pdf.setFont(font_name, 7.1)
            pdf.setFillColor(color_hex(THEME["text"]))
            pdf.drawString(x0 + 5, text_y, ellipsize_text(label, x1 - x0 - 10, font_name, 7.1))
            continue
        for segment in band_segments_for_item(item, period_start, period_end):
            boxes = [date_boxes[day] for day in segment if day in date_boxes]
            if not boxes:
                continue
            if lane >= min(max_work_lanes(box[3] - box[1]) for box in boxes):
                overflow_days.update(segment)
                continue
            first_box = boxes[0]
            last_box = boxes[-1]
            x0 = first_box[0] + 3
            x1 = last_box[2] - 3
            band_y = first_box[1] + cell_h - 27 - lane * lane_step
            fill = STATUS_COLORS.get(item.status, STATUS_COLORS["未着手"])
            draw_pdf_band(pdf, x0, band_y, max(4, x1 - x0), band_h, fill, True, True)
            pdf.setFont(font_name, 6.5)
            pdf.setFillColor(color_hex(THEME["text"]))
            label_width = max(8, min(x1 - x0 - 8, first_box[2] - first_box[0] - 8))
            pdf.drawString(x0 + 4, band_y + 2.1, ellipsize_text(label, label_width, font_name, 6.5))
    for day in overflow_days:
        box = date_boxes.get(day)
        if box is None:
            continue
        pdf.setFont(font_name, 7)
        pdf.setFillColor(color_hex(THEME["muted"]))
        pdf.drawString(box[0] + 5, box[1] + 9, "…")


def make_output_path(request: PdfRequest) -> Path:
    folder = Path(request.save_folder)
    folder.mkdir(parents=True, exist_ok=True)
    site_part = sanitize_filename(request.site_name) or "site"
    span = f"{request.start_date:%Y%m%d}-{request.end_date:%Y%m%d}"
    output = folder / f"リフォーム進捗カレンダー_{site_part}_{span}.pdf"
    if not output.exists():
        return output
    suffix = datetime.now().strftime("%H%M%S")
    return folder / f"リフォーム進捗カレンダー_{site_part}_{span}_{suffix}.pdf"


def generate_pdf(request: PdfRequest) -> Path:
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfgen import canvas
    except ModuleNotFoundError as exc:
        raise RuntimeError(UI_TEXT["error_reportlab_missing"]) from exc

    output_path = make_output_path(request)
    page_width, page_height = A4
    pdf = canvas.Canvas(str(output_path), pagesize=A4)
    pdf.setTitle(UI_TEXT["pdf_title"])
    font_name = register_pdf_fonts()

    period_days = days_between(request.start_date, request.end_date)
    weeks = iter_period_weeks(request.start_date, request.end_date)
    holiday_map = japanese_holidays({day.year for day in period_days})

    margin_x = 34
    header_top = page_height - 36
    grid_top = page_height - 132
    day_header_h = 20
    grid_bottom = 82
    footer_y = 36
    grid_width = page_width - margin_x * 2
    cell_w = grid_width / DAYS_PER_ROW
    cell_h = (grid_top - day_header_h - grid_bottom) / max(1, len(weeks))

    site = request.site_name.strip() or UI_TEXT["pdf_site_placeholder"]
    period_text = f"{format_full_date(request.start_date)}{UI_TEXT['pdf_range_separator']}{format_full_date(request.end_date)}"
    day_count_text = UI_TEXT["pdf_day_count_template"].format(count=len(period_days))
    contact_text = UI_TEXT["pdf_contact_template"].format(
        branch=request.branch_name.strip(),
        staff=request.staff_name.strip(),
        phone=request.phone.strip(),
    ).strip()
    date_boxes: dict[date, tuple[float, float, float, float]] = {}

    pdf.setFillColor(color_hex(THEME["text"]))
    pdf.setFont(font_name, 16)
    pdf.drawString(margin_x, header_top, UI_TEXT["pdf_title"])
    pdf.setFont(font_name, 11)
    pdf.drawString(margin_x, header_top - 22, site)
    pdf.setFont(font_name, 8.5)
    pdf.setFillColor(color_hex(THEME["muted"]))
    pdf.drawString(margin_x, header_top - 43, f"{UI_TEXT['pdf_target_period']}：{period_text}")
    pdf.drawRightString(page_width - margin_x, header_top - 43, f"{UI_TEXT['pdf_day_count']}：{day_count_text}")
    pdf.drawRightString(page_width - margin_x, header_top - 64, contact_text)

    for col, header in enumerate(UI_TEXT["pdf_weekday_headers"]):
        x = margin_x + col * cell_w
        y = grid_top - day_header_h
        pdf.setFillColor(color_hex("#FFFFFF"))
        pdf.rect(x, y, cell_w, day_header_h, stroke=0, fill=1)
        pdf.setStrokeColor(color_hex(THEME["border"]))
        pdf.rect(x, y, cell_w, day_header_h, stroke=1, fill=0)
        pdf.setFont(font_name, 9)
        pdf.setFillColor(color_hex(THEME["weekend_text"] if col in (0, 6) else THEME["text"]))
        pdf.drawCentredString(x + cell_w / 2, y + 6.5, header)

    for row, week in enumerate(weeks):
        for col, day in enumerate(week):
            x = margin_x + col * cell_w
            y = grid_top - day_header_h - (row + 1) * cell_h
            if day is None:
                pdf.setFillColor(color_hex(THEME["outside_bg"]))
                pdf.rect(x, y, cell_w, cell_h, stroke=0, fill=1)
                pdf.setStrokeColor(color_hex(THEME["border"]))
                pdf.rect(x, y, cell_w, cell_h, stroke=1, fill=0)
                continue

            is_weekend = day.weekday() >= 5
            holiday_name = holiday_map.get(day)
            pdf.setFillColor(color_hex(THEME["weekend_bg"] if is_weekend else "#FFFFFF"))
            pdf.rect(x, y, cell_w, cell_h, stroke=0, fill=1)
            pdf.setStrokeColor(color_hex(THEME["border"]))
            pdf.rect(x, y, cell_w, cell_h, stroke=1, fill=0)

            weekday = UI_TEXT["pdf_weekday_names"][day.weekday()]
            date_x = x + 5
            date_y = y + cell_h - 13
            date_boxes[day] = (x, y, x + cell_w, y + cell_h)
            pdf.setFont(font_name, 8.5)
            pdf.setFillColor(color_hex(THEME["text"]))
            pdf.drawString(date_x, date_y, f"{day.month}/{day.day}")
            pdf.setFillColor(color_hex(THEME["weekend_text"] if is_weekend else THEME["muted"]))
            pdf.drawString(date_x + 30, date_y, weekday)

            text_y = date_y - 11
            if holiday_name:
                pdf.setFont(font_name, 6.6)
                pdf.setFillColor(color_hex(THEME["holiday_text"]))
                pdf.drawString(date_x, text_y, ellipsize_text(holiday_name, cell_w - 10, font_name, 6.6))
                text_y -= 9


    draw_pdf_work_items(pdf, request.work_items, date_boxes, cell_h, font_name, request.start_date, request.end_date)
    completion_box = date_boxes.get(request.end_date)
    if completion_box:
        x, y, _x1, _y1 = completion_box
        pdf.setFont(font_name, 8.4)
        pdf.setFillColor(color_hex(THEME["accent"]))
        pdf.drawString(x + 5, y + 7, UI_TEXT["pdf_completion_label"])

    pdf.setStrokeColor(color_hex(THEME["border"]))
    pdf.line(margin_x, grid_bottom - 18, page_width - margin_x, grid_bottom - 18)
    pdf.setFillColor(color_hex(THEME["muted"]))
    pdf.setFont(font_name, 7.4)
    pdf.drawString(margin_x, footer_y, f"{UI_TEXT['footer_left']} ｜ {UI_TEXT['footer_tagline']}")
    pdf.drawRightString(page_width - margin_x, footer_y, UI_TEXT["footer_copyright"])
    pdf.showPage()
    pdf.save()
    return output_path


class ReformProgressApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.geometry("1200x780")
        self.root.minsize(1000, 680)
        self.root.configure(bg=THEME["background"])
        apply_window_icon(self.root)

        self.config_data = load_config()
        self.project_data = load_project()
        today = date.today()
        project_start = parse_date_text(str(self.project_data.get("start_date", ""))) if self.project_data else None
        start_default = project_start or today
        project_end = parse_date_text(str(self.project_data.get("end_date", ""))) if self.project_data else None
        end_default = project_end or (start_default + timedelta(days=DEFAULT_CALENDAR_DAYS - 1))

        def project_text(key: str, fallback: str = "") -> str:
            if self.project_data and isinstance(self.project_data.get(key), str):
                return str(self.project_data.get(key, ""))
            return fallback

        self.site_name_var = tk.StringVar(value=project_text("site_name"))
        self.branch_name_var = tk.StringVar(value=project_text("branch_name", self.config_data.get("branch_name", "")))
        self.staff_name_var = tk.StringVar(value=project_text("staff_name", self.config_data.get("staff_name", "")))
        self.phone_var = tk.StringVar(value=project_text("phone", self.config_data.get("phone", "")))
        self.start_date_var = tk.StringVar(value=format_date(start_default))
        self.end_date_var = tk.StringVar(value=format_date(end_default))
        self.reference_day_var = tk.StringVar()
        self.save_folder_var = tk.StringVar(value=project_text("save_folder", self.config_data.get("save_folder", default_save_folder())))
        self.status_var = tk.StringVar(value=UI_TEXT["status_project_loaded"] if self.project_data else UI_TEXT["status_ready"])

        project_items = deserialize_work_items(self.project_data.get("work_items")) if self.project_data else None
        self.work_items = project_items or place_initial_work_items(start_default)
        self.row_widgets: list[RowWidgets] = []
        self.preview_date_boxes: dict[date, tuple[float, float, float, float]] = {}
        self.preview_item_refs: dict[int, dict[str, object]] = {}
        self.drag_context: dict[str, object] | None = None
        self.row_drag_index: int | None = None

        self.result_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.is_processing = False
        self.processing_started = 0.0
        self.footer_compact: bool | None = None

        self.font_family = self.choose_font_family()
        self.fonts = {
            "title": (self.font_family, 18, "bold"),
            "subtitle": (self.font_family, 10),
            "section": (self.font_family, 10, "bold"),
            "body": (self.font_family, 9),
            "small": (self.font_family, 8),
            "button": (self.font_family, 9, "bold"),
            "footer": (self.font_family, 8),
        }

        self.build_styles()
        self.update_reference_day()
        self.build_ui()
        self.bind_global_traces()
        self.refresh_rows()
        self.root.after(100, self.draw_preview)

    def choose_font_family(self) -> str:
        available = set(tkfont.families(self.root))
        for candidate in FONT_CANDIDATES:
            if candidate in available:
                return candidate
        return "TkDefaultFont"

    def build_styles(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TEntry", font=self.fonts["body"], padding=4)
        style.configure("TCombobox", font=self.fonts["body"], padding=2)
        style.configure("TCheckbutton", background=THEME["panel"], focuscolor=THEME["panel"])
        style.configure("Work.TCheckbutton", background=THEME["panel"], padding=(6, 0), focuscolor=THEME["panel"])
        style.map("Work.TCheckbutton", background=[("active", THEME["row_alt"])], indicatorcolor=[("selected", THEME["accent"]), ("!selected", "#FFFFFF")])
        style.configure("Primary.TButton", font=self.fonts["button"], padding=(14, 8), background=THEME["accent"], foreground="#FFFFFF")
        style.map("Primary.TButton", background=[("active", THEME["accent_hover"]), ("disabled", THEME["accent_disabled"])])
        style.configure("Secondary.TButton", font=self.fonts["button"], padding=(10, 7), background=THEME["panel"], foreground=THEME["text"])
        style.map("Secondary.TButton", background=[("active", THEME["subtle"])])
        style.configure("Tiny.TButton", font=self.fonts["button"], padding=(6, 2), background=THEME["panel"], foreground=THEME["text"])
        style.configure("Vertical.TScrollbar", gripcount=0, background="#CBD5E1", troughcolor=THEME["subtle"], bordercolor=THEME["panel"], lightcolor="#CBD5E1", darkcolor="#CBD5E1", arrowsize=10, width=10)
        style.map("Vertical.TScrollbar", background=[("active", "#94A3B8")])
        style.configure("Dake.Horizontal.TProgressbar", troughcolor=THEME["subtle"], background=THEME["accent"], bordercolor=THEME["background"], lightcolor=THEME["accent"], darkcolor=THEME["accent"])

    def build_ui(self) -> None:
        self.main = tk.Frame(self.root, bg=THEME["background"])
        self.main.pack(fill="both", expand=True, padx=18, pady=(14, 0))

        self.build_header(self.main)
        self.build_top_fields(self.main)
        self.build_center(self.main)
        self.build_footer()

    def build_header(self, parent: tk.Frame) -> None:
        header = tk.Frame(parent, bg=THEME["background"])
        header.pack(fill="x", pady=(0, 8))
        tk.Label(header, text=UI_TEXT["main_title"], bg=THEME["background"], fg=THEME["text"], font=self.fonts["title"]).pack(side="left")
        tk.Label(header, text=UI_TEXT["main_description"], bg=THEME["background"], fg=THEME["muted"], font=self.fonts["subtitle"]).pack(side="left", padx=(16, 0), pady=(5, 0))

    def build_top_fields(self, parent: tk.Frame) -> None:
        panel = self.panel(parent)
        panel.pack(fill="x", pady=(0, 10))
        for col in range(10):
            panel.grid_columnconfigure(col, weight=1 if col in (1, 3, 5, 7) else 0)

        fields = (
            ("label_site_name", self.site_name_var, 0, 0),
            ("label_branch_name", self.branch_name_var, 0, 2),
            ("label_staff_name", self.staff_name_var, 0, 4),
            ("label_phone", self.phone_var, 0, 6),
            ("label_start_date", self.start_date_var, 1, 0),
            ("label_end_date", self.end_date_var, 1, 2),
        )
        for label_key, variable, row, col in fields:
            self.add_label(panel, UI_TEXT[label_key], row, col)
            ttk.Entry(panel, textvariable=variable, width=14).grid(row=row, column=col + 1, sticky="ew", padx=(0, 10), pady=5)

        tk.Label(panel, textvariable=self.reference_day_var, bg=THEME["panel"], fg=THEME["accent"], font=self.fonts["small"]).grid(row=1, column=4, columnspan=2, sticky="w", padx=(0, 8), pady=5)
        tk.Label(panel, text=UI_TEXT["date_hint"], bg=THEME["panel"], fg=THEME["muted"], font=self.fonts["small"]).grid(row=1, column=6, sticky="w", padx=(0, 8), pady=5)
        self.add_label(panel, UI_TEXT["label_save_folder"], 2, 0)
        ttk.Entry(panel, textvariable=self.save_folder_var).grid(row=2, column=1, columnspan=6, sticky="ew", padx=(0, 8), pady=5)
        ttk.Button(panel, text=UI_TEXT["button_select_folder"], style="Secondary.TButton", command=self.choose_save_folder).grid(row=2, column=7, sticky="ew", padx=(0, 8), pady=5)

        actions = tk.Frame(panel, bg=THEME["panel"])
        actions.grid(row=0, column=8, rowspan=3, columnspan=2, sticky="nse", padx=(8, 0))
        self.create_button = ttk.Button(actions, text=UI_TEXT["button_create_pdf"], style="Primary.TButton", command=self.on_create_pdf)
        self.create_button.pack(fill="x", pady=(0, 6))
        ttk.Button(actions, text=UI_TEXT["button_save_project"], style="Secondary.TButton", command=self.on_save_project).pack(fill="x", pady=(0, 6))
        ttk.Button(actions, text=UI_TEXT["button_reschedule"], style="Secondary.TButton", command=self.on_reschedule).pack(fill="x", pady=(0, 6))
        ttk.Button(actions, text=UI_TEXT["button_today"], style="Secondary.TButton", command=self.on_move_today).pack(fill="x")

    def build_center(self, parent: tk.Frame) -> None:
        center = tk.Frame(parent, bg=THEME["background"])
        center.pack(fill="both", expand=True)
        center.grid_columnconfigure(0, weight=6, uniform="center")
        center.grid_columnconfigure(1, weight=5, uniform="center")
        center.grid_rowconfigure(0, weight=1)

        left = self.panel(center)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        right = self.panel(center)
        right.grid(row=0, column=1, sticky="nsew")
        left.grid_columnconfigure(0, weight=1)
        left.grid_rowconfigure(1, weight=1)
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(1, weight=1)

        tk.Label(left, text=UI_TEXT["section_work_list"], bg=THEME["panel"], fg=THEME["text"], font=self.fonts["section"], anchor="w").grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 6))
        self.build_work_table(left)

        tk.Label(right, text=UI_TEXT["section_preview"], bg=THEME["panel"], fg=THEME["text"], font=self.fonts["section"], anchor="w").grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 6))
        preview_wrap = tk.Frame(right, bg=THEME["panel"])
        preview_wrap.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))
        preview_wrap.grid_rowconfigure(0, weight=1)
        preview_wrap.grid_columnconfigure(0, weight=1)
        self.preview_canvas = tk.Canvas(preview_wrap, bg=THEME["panel"], highlightthickness=0)
        self.preview_canvas.grid(row=0, column=0, sticky="nsew")
        self.preview_canvas.bind("<Configure>", lambda _event: self.draw_preview())
        self.preview_canvas.bind("<ButtonPress-1>", self.on_preview_press)
        self.preview_canvas.bind("<B1-Motion>", self.on_preview_motion)
        self.preview_canvas.bind("<ButtonRelease-1>", self.on_preview_release)
        self.preview_canvas.bind("<Motion>", self.on_preview_hover)
        self.preview_canvas.bind("<Leave>", lambda _event: self.preview_canvas.configure(cursor=""))

    def build_work_table(self, parent: tk.Frame) -> None:
        table_wrap = tk.Frame(parent, bg=THEME["panel"])
        table_wrap.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))
        table_wrap.grid_columnconfigure(0, weight=1)
        table_wrap.grid_rowconfigure(1, weight=1)

        header = tk.Frame(table_wrap, bg=THEME["subtle"])
        header.grid(row=0, column=0, sticky="ew")
        widths = (5, 22, 12, 12, 9, 5, 7)
        labels = ("table_use", "table_work_name", "table_start_date", "table_end_date", "table_status", "table_days", "table_order")
        for index, key in enumerate(labels):
            header.grid_columnconfigure(index, weight=1 if index == 1 else 0, minsize=widths[index] * 7)
            tk.Label(header, text=UI_TEXT[key], bg=THEME["subtle"], fg=THEME["text"], font=self.fonts["small"], anchor="w", padx=5, pady=4).grid(row=0, column=index, sticky="ew")

        self.rows_canvas = tk.Canvas(table_wrap, bg=THEME["panel"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(table_wrap, orient="vertical", command=self.rows_canvas.yview, style="Vertical.TScrollbar")
        self.rows_body = tk.Frame(self.rows_canvas, bg=THEME["panel"])
        self.rows_window = self.rows_canvas.create_window((0, 0), window=self.rows_body, anchor="nw")
        self.rows_body.bind("<Configure>", lambda _event: self.rows_canvas.configure(scrollregion=self.rows_canvas.bbox("all")))
        self.rows_canvas.bind("<Configure>", lambda event: self.rows_canvas.itemconfigure(self.rows_window, width=event.width))
        self.rows_canvas.configure(yscrollcommand=scrollbar.set)
        self.rows_canvas.grid(row=1, column=0, sticky="nsew")
        scrollbar.grid(row=1, column=1, sticky="ns")

    def build_footer(self) -> None:
        bottom = tk.Frame(self.root, bg=THEME["background"])
        bottom.pack(fill="x", padx=18, pady=(8, 10))
        self.progress = ttk.Progressbar(bottom, mode="indeterminate", length=130, style="Dake.Horizontal.TProgressbar")
        self.progress.pack(side="left", padx=(0, 10))
        tk.Label(bottom, textvariable=self.status_var, bg=THEME["background"], fg=THEME["muted"], font=self.fonts["body"], anchor="w").pack(side="left", fill="x", expand=True)
        self.footer = tk.Frame(bottom, bg=THEME["background"])
        self.footer.pack(side="right")
        self.root.bind("<Configure>", self.update_footer_layout, add="+")
        self.update_footer_layout()

    def panel(self, parent: tk.Misc) -> tk.Frame:
        return tk.Frame(parent, bg=THEME["panel"], highlightbackground=THEME["border"], highlightthickness=1)

    def add_label(self, parent: tk.Misc, text: str, row: int, col: int) -> None:
        tk.Label(parent, text=text, bg=THEME["panel"], fg=THEME["muted"], font=self.fonts["body"], anchor="w").grid(row=row, column=col, sticky="w", padx=(12 if col == 0 else 4, 6), pady=5)

    def update_reference_day(self) -> None:
        start = parse_date_text(self.start_date_var.get())
        if start is None:
            self.reference_day_var.set(UI_TEXT["reference_day_invalid"])
            return
        self.reference_day_var.set(UI_TEXT["reference_day_template"].format(date=format_short_date(reference_day_for_start(start))))

    def on_top_dates_changed(self) -> None:
        self.update_reference_day()
        self.draw_preview()

    def bind_global_traces(self) -> None:
        for variable in (self.start_date_var, self.end_date_var):
            variable.trace_add("write", lambda *_args: self.root.after_idle(self.on_top_dates_changed))

    def validate_period_limit(self, start: date, end: date, show_errors: bool) -> bool:
        if is_period_within_limit(start, end):
            return True
        if show_errors:
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["error_date_too_long"])
        return False

    def current_project_data(self) -> dict[str, object]:
        return build_project_data(
            site_name=self.site_name_var.get(),
            branch_name=self.branch_name_var.get(),
            staff_name=self.staff_name_var.get(),
            phone=self.phone_var.get(),
            start_date_text=self.start_date_var.get(),
            end_date_text=self.end_date_var.get(),
            save_folder=self.save_folder_var.get(),
            work_items=self.work_items,
        )

    def refresh_rows(self) -> None:
        for child in self.rows_body.winfo_children():
            child.destroy()
        self.row_widgets.clear()
        for index, item in enumerate(self.work_items):
            self.add_work_row(index, item)
        self.refresh_day_counts()
        self.draw_preview()

    def add_work_row(self, index: int, item: WorkItem) -> None:
        row_bg = THEME["panel"] if index % 2 == 0 else THEME["row_alt"]
        frame = tk.Frame(self.rows_body, bg=row_bg, highlightthickness=0)
        frame.grid(row=index, column=0, sticky="ew", pady=1)
        frame.grid_columnconfigure(1, weight=1)
        selected = tk.BooleanVar(value=item.selected)
        name = tk.StringVar(value=item.name)
        start = tk.StringVar(value=format_date(item.start_date))
        end = tk.StringVar(value=format_date(item.end_date))
        status = tk.StringVar(value=item.status)
        days = tk.StringVar()
        widgets = RowWidgets(frame=frame, selected=selected, name=name, start_date=start, end_date=end, status=status, days=days)
        self.row_widgets.append(widgets)

        ttk.Checkbutton(frame, variable=selected, command=self.on_rows_changed, style="Work.TCheckbutton").grid(row=0, column=0, padx=(6, 4), pady=4)
        name_entry = ttk.Entry(frame, textvariable=name, width=22)
        name_entry.grid(row=0, column=1, sticky="ew", padx=2, pady=4)
        start_entry = ttk.Entry(frame, textvariable=start, width=11)
        start_entry.grid(row=0, column=2, padx=2, pady=4)
        end_entry = ttk.Entry(frame, textvariable=end, width=11)
        end_entry.grid(row=0, column=3, padx=2, pady=4)
        status_combo = ttk.Combobox(frame, textvariable=status, values=STATUS_OPTIONS, width=8, state="readonly")
        status_combo.grid(row=0, column=4, padx=2, pady=4)
        tk.Label(frame, textvariable=days, bg=row_bg, fg=THEME["muted"], font=self.fonts["small"], width=4).grid(row=0, column=5, padx=2)
        buttons = tk.Frame(frame, bg=row_bg)
        buttons.grid(row=0, column=6, padx=(2, 6), pady=4)
        ttk.Button(buttons, text=UI_TEXT["button_up"], style="Tiny.TButton", width=2, command=lambda i=index: self.move_row(i, -1)).pack(side="left", padx=(0, 2))
        ttk.Button(buttons, text=UI_TEXT["button_down"], style="Tiny.TButton", width=2, command=lambda i=index: self.move_row(i, 1)).pack(side="left")

        for widget in (frame, name_entry, start_entry, end_entry):
            widget.bind("<ButtonPress-1>", lambda event, i=index: self.on_row_press(event, i), add="+")
            widget.bind("<ButtonRelease-1>", self.on_row_release, add="+")
        for entry in (name_entry, start_entry, end_entry):
            entry.bind("<FocusOut>", lambda _event: self.on_rows_changed())
            entry.bind("<Return>", lambda _event: self.on_rows_changed())
        status_combo.bind("<<ComboboxSelected>>", lambda _event: self.on_rows_changed())

    def refresh_day_counts(self) -> None:
        self.sync_items_from_rows(show_errors=False)
        for item, widgets in zip(self.work_items, self.row_widgets):
            widgets.days.set(str(item_duration(item)) if item.start_date and item.end_date else "")

    def sync_items_from_rows(self, show_errors: bool = False) -> bool:
        for index, widgets in enumerate(self.row_widgets):
            item = self.work_items[index]
            item.selected = bool(widgets.selected.get())
            item.name = widgets.name.get().strip()
            item.status = widgets.status.get() if widgets.status.get() in STATUS_OPTIONS else STATUS_OPTIONS[0]
            start = parse_date_text(widgets.start_date.get())
            end = parse_date_text(widgets.end_date.get())
            if widgets.start_date.get().strip() and start is None:
                if show_errors:
                    messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["error_start_date"])
                return False
            if widgets.end_date.get().strip() and end is None:
                if show_errors:
                    messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["error_end_date"])
                return False
            if start and end and end < start:
                if show_errors:
                    messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["error_date_order"])
                return False
            item.start_date = start
            item.end_date = end
            if start and end:
                item.planned_days = item_duration(item)
        return True

    def on_rows_changed(self) -> None:
        if not self.sync_items_from_rows(show_errors=False):
            return
        self.refresh_day_counts()
        self.draw_preview()

    def move_row(self, index: int, delta: int) -> None:
        self.sync_items_from_rows(show_errors=False)
        new_index = max(0, min(len(self.work_items) - 1, index + delta))
        if new_index == index:
            return
        self.work_items[index], self.work_items[new_index] = self.work_items[new_index], self.work_items[index]
        self.status_var.set(UI_TEXT["status_row_moved"])
        self.refresh_rows()

    def on_row_press(self, _event, index: int) -> None:
        self.row_drag_index = index

    def on_row_release(self, event) -> None:
        if self.row_drag_index is None:
            return
        self.sync_items_from_rows(show_errors=False)
        pointer_y = self.rows_body.winfo_pointery()
        destination = self.row_drag_index
        for index, widgets in enumerate(self.row_widgets):
            top = widgets.frame.winfo_rooty()
            bottom = top + max(1, widgets.frame.winfo_height())
            if top <= pointer_y <= bottom:
                destination = index
                break
        source = self.row_drag_index
        self.row_drag_index = None
        if destination == source:
            return
        item = self.work_items.pop(source)
        self.work_items.insert(destination, item)
        self.status_var.set(UI_TEXT["status_row_moved"])
        self.refresh_rows()

    def choose_save_folder(self) -> None:
        initial = self.save_folder_var.get().strip() or default_save_folder()
        selected = filedialog.askdirectory(
            title=UI_TEXT["dialog_select_save_dir"],
            initialdir=initial if os.path.isdir(initial) else default_save_folder(),
        )
        if selected:
            self.save_folder_var.set(selected)

    def collect_request(self) -> PdfRequest | None:
        if not self.sync_items_from_rows(show_errors=True):
            return None
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
        if not self.validate_period_limit(start, end, show_errors=True):
            return None
        save_folder = self.save_folder_var.get().strip()
        if not save_folder:
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["error_save_folder"])
            return None
        output_items = tuple(copy_work_item(item) for item in self.work_items if is_output_item(item))
        if not output_items:
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["error_no_work"])
            return None
        return PdfRequest(
            site_name=self.site_name_var.get().strip(),
            branch_name=self.branch_name_var.get().strip(),
            staff_name=self.staff_name_var.get().strip(),
            phone=self.phone_var.get().strip(),
            start_date=start,
            end_date=end,
            save_folder=save_folder,
            work_items=output_items,
        )

    def on_reschedule(self) -> None:
        start = parse_date_text(self.start_date_var.get())
        if start is None:
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["error_start_date"])
            return
        self.sync_items_from_rows(show_errors=False)
        self.work_items = reschedule_work_items(self.work_items, start)
        self.status_var.set(UI_TEXT["status_rescheduled"])
        self.refresh_rows()

    def on_move_today(self) -> None:
        today = date.today()
        self.start_date_var.set(format_date(today))
        self.end_date_var.set(format_date(today + timedelta(days=DEFAULT_CALENDAR_DAYS - 1)))
        self.work_items = reschedule_work_items(self.work_items, today)
        self.status_var.set(UI_TEXT["status_moved_today"])
        self.refresh_rows()

    def on_save_project(self) -> None:
        if not self.sync_items_from_rows(show_errors=True):
            return
        start = parse_date_text(self.start_date_var.get())
        if start is None:
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["error_start_date"])
            return
        end = parse_date_text(self.end_date_var.get())
        if end is None:
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["error_end_date"])
            return
        if end < start:
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["error_date_order"])
            return
        if not self.validate_period_limit(start, end, show_errors=True):
            return
        if save_project(self.current_project_data()):
            self.status_var.set(UI_TEXT["status_project_saved"])
        else:
            self.status_var.set(UI_TEXT["status_project_save_failed"])
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["status_project_save_failed"])

    def on_create_pdf(self) -> None:
        if self.is_processing:
            return
        request = self.collect_request()
        if request is None:
            return
        save_config(
            {
                "branch_name": self.branch_name_var.get(),
                "staff_name": self.staff_name_var.get(),
                "phone": self.phone_var.get(),
                "save_folder": self.save_folder_var.get(),
            }
        )
        self.is_processing = True
        self.processing_started = time.monotonic()
        self.create_button.configure(state="disabled")
        self.progress.start(12)
        self.status_var.set(UI_TEXT["status_processing"])
        threading.Thread(target=self.worker_create_pdf, args=(request,), daemon=True).start()
        self.root.after(120, self.poll_worker)

    def worker_create_pdf(self, request: PdfRequest) -> None:
        try:
            output_path = generate_pdf(request)
            self.result_queue.put(("ok", output_path))
        except Exception as exc:
            self.result_queue.put(("error", str(exc)))

    def poll_worker(self) -> None:
        try:
            status, payload = self.result_queue.get_nowait()
        except queue.Empty:
            if self.is_processing:
                elapsed = int((time.monotonic() - self.processing_started) * 2) % 4
                self.status_var.set(UI_TEXT["status_processing"] + "." * elapsed)
                self.root.after(140, self.poll_worker)
            return
        self.is_processing = False
        self.progress.stop()
        self.create_button.configure(state="normal")
        if status == "ok":
            output_path = Path(payload)
            if open_output_folder(output_path):
                self.status_var.set(UI_TEXT["status_complete"])
            else:
                self.status_var.set(UI_TEXT["status_pdf_done_folder_open_failed"])
            messagebox.showinfo(UI_TEXT["dialog_saved_title"], UI_TEXT["dialog_saved_message"].format(path=str(output_path)))
        else:
            self.status_var.set(UI_TEXT["status_error"])
            messagebox.showerror(UI_TEXT["dialog_error_title"], str(payload))

    def update_footer_layout(self, _event=None) -> None:
        compact = self.root.winfo_width() < FOOTER_COMPACT_WIDTH
        if compact == self.footer_compact:
            return
        self.footer_compact = compact
        for child in self.footer.winfo_children():
            child.destroy()
        left = tk.Frame(self.footer, bg=THEME["background"])
        self.make_footer_text(left, UI_TEXT["footer_left"])
        self.make_footer_text(left, UI_TEXT["footer_separator"])
        self.make_footer_text(left, UI_TEXT["footer_tagline"])
        right = tk.Frame(self.footer, bg=THEME["background"])
        self.make_footer_link(right, UI_TEXT["footer_link_1"], LINKS["assessment"])
        self.make_footer_text(right, UI_TEXT["footer_separator"])
        self.make_footer_link(right, UI_TEXT["footer_link_2"], LINKS["instagram"])
        self.make_footer_text(right, UI_TEXT["footer_separator"])
        self.make_footer_text(right, UI_TEXT["footer_copyright"])
        if compact:
            left.pack(anchor="e")
            right.pack(anchor="e")
        else:
            left.pack(side="left")
            right.pack(side="left", padx=(10, 0))

    def make_footer_text(self, parent: tk.Frame, text: str) -> None:
        tk.Label(parent, text=text, bg=THEME["background"], fg=THEME["muted"], font=self.fonts["footer"]).pack(side="left")

    def make_footer_link(self, parent: tk.Frame, text: str, url: str) -> None:
        label = tk.Label(parent, text=text, bg=THEME["background"], fg=THEME["link"], font=self.fonts["footer"], cursor="hand2")
        label.pack(side="left")
        label.bind("<Button-1>", lambda _event: webbrowser.open(url, new=2))
        label.bind("<Enter>", lambda _event: label.configure(fg=THEME["accent"]))
        label.bind("<Leave>", lambda _event: label.configure(fg=THEME["link"]))

    def draw_preview(self) -> None:
        if not hasattr(self, "preview_canvas"):
            return
        self.sync_items_from_rows(show_errors=False)
        canvas = self.preview_canvas
        canvas.delete("all")
        self.preview_date_boxes.clear()
        self.preview_item_refs.clear()
        start = parse_date_text(self.start_date_var.get())
        end = parse_date_text(self.end_date_var.get())
        width = max(200, canvas.winfo_width())
        height = max(200, canvas.winfo_height())
        margin = 18
        page_w = min(width - margin * 2, (height - margin * 2) * 0.707)
        page_h = page_w / 0.707
        if page_h > height - margin * 2:
            page_h = height - margin * 2
            page_w = page_h * 0.707
        x0 = (width - page_w) / 2
        y0 = margin
        x1 = x0 + page_w
        y1 = y0 + page_h
        canvas.create_rectangle(x0, y0, x1, y1, fill="#FFFFFF", outline=THEME["border"], width=1)
        canvas.create_text(x0 + 16, y0 + 18, text=UI_TEXT["pdf_title"], anchor="w", fill=THEME["text"], font=(self.font_family, 11, "bold"))
        if start is None or end is None or end < start:
            canvas.create_text((x0 + x1) / 2, (y0 + y1) / 2, text=UI_TEXT["error_date_order"], fill=THEME["muted"], font=self.fonts["body"])
            return
        if not is_period_within_limit(start, end):
            canvas.create_text((x0 + x1) / 2, (y0 + y1) / 2, text=UI_TEXT["error_date_too_long"], fill=THEME["error"], font=self.fonts["body"], width=page_w - 40)
            return
        weeks = iter_period_weeks(start, end)
        grid_x = x0 + 14
        grid_y = y0 + 58
        grid_w = page_w - 28
        grid_h = page_h - 98
        header_h = 20
        cell_w = grid_w / DAYS_PER_ROW
        cell_h = (grid_h - header_h) / max(1, len(weeks))
        for col, label in enumerate(UI_TEXT["pdf_weekday_headers"]):
            hx = grid_x + col * cell_w
            canvas.create_rectangle(hx, grid_y, hx + cell_w, grid_y + header_h, fill="#FFFFFF", outline=THEME["border"])
            canvas.create_text(hx + cell_w / 2, grid_y + header_h / 2, text=label, fill=THEME["weekend_text"] if col in (0, 6) else THEME["text"], font=self.fonts["small"])
        for row, week in enumerate(weeks):
            for col, day in enumerate(week):
                cx = grid_x + col * cell_w
                cy = grid_y + header_h + row * cell_h
                if day is None:
                    canvas.create_rectangle(cx, cy, cx + cell_w, cy + cell_h, fill=THEME["outside_bg"], outline=THEME["border"])
                    continue
                fill = THEME["weekend_bg"] if day.weekday() >= 5 else "#FFFFFF"
                canvas.create_rectangle(cx, cy, cx + cell_w, cy + cell_h, fill=fill, outline=THEME["border"])
                self.preview_date_boxes[day] = (cx, cy, cx + cell_w, cy + cell_h)
                canvas.create_text(cx + 4, cy + 10, text=f"{day.month}/{day.day}", anchor="w", fill=THEME["text"], font=self.fonts["small"])
                weekday = UI_TEXT["pdf_weekday_names"][day.weekday()]
                canvas.create_text(cx + 32, cy + 10, text=weekday, anchor="w", fill=THEME["weekend_text"] if day.weekday() >= 5 else THEME["muted"], font=self.fonts["small"])
                if day == end:
                    canvas.create_text(cx + 4, cy + cell_h - 10, text=UI_TEXT["pdf_completion_label"], anchor="w", fill=THEME["accent"], font=(self.font_family, 8, "bold"))
        self.draw_preview_items(start, end, cell_h)

    def draw_preview_items(self, start: date, end: date, cell_h: float) -> None:
        lane_by_index = build_work_lanes(self.work_items)
        lane_step = 15
        block_h = 13
        overflow_days: set[date] = set()
        for index, item in enumerate(self.work_items):
            days = visible_workdays_for_item(item, start, end)
            if not days:
                continue
            lane = lane_by_index.get(index, 0)
            label = f"{STATUS_SYMBOLS.get(item.status, '')}{short_work_name(item.name)}"
            segments = band_segments_for_item(item, start, end)
            for segment in segments:
                boxes = [self.preview_date_boxes[day] for day in segment if day in self.preview_date_boxes]
                if not boxes:
                    continue
                if lane >= max(1, int((cell_h - 30) // lane_step)):
                    overflow_days.update(segment)
                    continue
                first_box = boxes[0]
                last_box = boxes[-1]
                x0 = first_box[0] + 3
                x1 = last_box[2] - 3
                y0 = first_box[1] + 24 + lane * lane_step
                y1 = y0 + block_h
                fill = STATUS_COLORS.get(item.status, STATUS_COLORS["未着手"])
                line = self.preview_canvas.create_line(
                    x0 + block_h / 2,
                    y0 + block_h / 2,
                    x1 - block_h / 2,
                    y0 + block_h / 2,
                    width=block_h,
                    fill=fill,
                    capstyle="round",
                    tags=(f"work:{index}", "workblock"),
                )
                outline = self.preview_canvas.create_rectangle(x0, y0, x1, y1, outline=THEME["band_outline"], width=1, tags=(f"work:{index}", "workblock"))
                text = self.preview_canvas.create_text(x0 + 7, y0 + block_h / 2, text=label, anchor="w", fill=THEME["text"], font=(self.font_family, 7), tags=(f"work:{index}", "workblock"))
                left_grip = self.preview_canvas.create_line(x0 + 5, y0 + 3, x0 + 5, y1 - 3, fill=THEME["grip"], width=1, tags=(f"work:{index}", "workblock", "workgrip"))
                right_grip = self.preview_canvas.create_line(x1 - 5, y0 + 3, x1 - 5, y1 - 3, fill=THEME["grip"], width=1, tags=(f"work:{index}", "workblock", "workgrip"))
                self.register_preview_refs((line, outline, text, left_grip, right_grip), index, segment[0], (x0, y0, x1, y1))
        for day in overflow_days:
            box = self.preview_date_boxes.get(day)
            if box:
                self.preview_canvas.create_text(box[0] + 5, box[1] + cell_h - 18, text="…", anchor="w", fill=THEME["muted"], font=self.fonts["small"])

    def register_preview_refs(self, item_ids: tuple[int, ...], index: int, day: date, bbox: tuple[float, float, float, float]) -> None:
        for item_id in item_ids:
            self.preview_item_refs[item_id] = {"index": index, "day": day, "bbox": bbox}

    def preview_hit_mode(self, x: float, bbox: tuple[float, float, float, float]) -> str:
        left, _top, right, _bottom = bbox
        edge = min(12, max(7, (right - left) * 0.18))
        if x <= left + edge:
            return "resize_start"
        if x >= right - edge:
            return "resize_end"
        return "move"

    def ref_under_pointer(self) -> dict[str, object] | None:
        current = self.preview_canvas.find_withtag("current")
        if not current:
            return None
        return self.preview_item_refs.get(current[0])

    def on_preview_hover(self, event) -> None:
        if self.drag_context is not None:
            return
        ref = self.ref_under_pointer()
        if ref is None:
            self.preview_canvas.configure(cursor="")
            return
        bbox = ref.get("bbox")
        if not isinstance(bbox, tuple):
            self.preview_canvas.configure(cursor="")
            return
        mode = self.preview_hit_mode(event.x, bbox)
        self.preview_canvas.configure(cursor="sb_h_double_arrow" if mode in ("resize_start", "resize_end") else "fleur")

    def date_at_canvas_point(self, x: float, y: float) -> date | None:
        for day, (x0, y0, x1, y1) in self.preview_date_boxes.items():
            if x0 <= x <= x1 and y0 <= y <= y1:
                return day
        return None

    def on_preview_press(self, event) -> None:
        ref = self.ref_under_pointer()
        if ref is None:
            return
        index = ref.get("index")
        day = ref.get("day")
        bbox = ref.get("bbox")
        if not isinstance(index, int) or not isinstance(day, date) or not isinstance(bbox, tuple):
            return
        mode = self.preview_hit_mode(event.x, bbox)
        self.drag_context = {
            "index": index,
            "mode": mode,
            "origin_day": day,
            "target_day": day,
            "highlight": None,
        }
        self.preview_canvas.configure(cursor="sb_h_double_arrow" if mode in ("resize_start", "resize_end") else "fleur")

    def on_preview_motion(self, event) -> None:
        if self.drag_context is None:
            return
        target = self.date_at_canvas_point(event.x, event.y)
        if target is None:
            return
        self.drag_context["target_day"] = target
        highlight_id = self.drag_context.get("highlight")
        if highlight_id:
            self.preview_canvas.delete(highlight_id)
        box = self.preview_date_boxes.get(target)
        if not box:
            return
        x0, y0, x1, y1 = box
        self.drag_context["highlight"] = self.preview_canvas.create_rectangle(x0 + 2, y0 + 2, x1 - 2, y1 - 2, fill=THEME["drag_highlight"], outline=THEME["accent"], stipple="gray25", width=1)
        snapped, _adjusted = snap_to_workday(target)
        mode = str(self.drag_context.get("mode", "move"))
        if mode == "resize_start":
            self.status_var.set(UI_TEXT["status_dragging_start"].format(date=format_short_date(snapped)))
        elif mode == "resize_end":
            self.status_var.set(UI_TEXT["status_dragging_end"].format(date=format_short_date(snapped)))
        else:
            self.status_var.set(UI_TEXT["status_dragging_move"].format(date=format_short_date(snapped)))

    def on_preview_release(self, _event) -> None:
        if self.drag_context is None:
            return
        highlight_id = self.drag_context.get("highlight")
        if highlight_id:
            self.preview_canvas.delete(highlight_id)
        index = int(self.drag_context["index"])
        target_day = self.drag_context.get("target_day")
        mode = str(self.drag_context["mode"])
        self.drag_context = None
        self.preview_canvas.configure(cursor="")
        if not isinstance(target_day, date) or index >= len(self.work_items):
            return
        self.sync_items_from_rows(show_errors=False)
        original_day = target_day
        if mode == "resize_start":
            updated, adjusted, snapped = resize_item_to_start(self.work_items[index], target_day)
            self.work_items[index] = updated
            self.status_var.set(UI_TEXT["status_drag_start_resized"].format(name=updated.name, date=format_short_date(snapped)))
        elif mode == "resize_end":
            updated, adjusted, snapped = resize_item_to_end(self.work_items[index], target_day)
            self.work_items[index] = updated
            self.status_var.set(UI_TEXT["status_drag_end_resized"].format(name=updated.name, date=format_short_date(snapped)))
        else:
            updated, adjusted, snapped = move_item_to_start(self.work_items[index], target_day)
            self.work_items[index] = updated
            self.status_var.set(UI_TEXT["status_drag_moved"].format(name=updated.name, date=format_short_date(snapped)))
        if adjusted:
            self.status_var.set(UI_TEXT["status_weekend_adjusted"].format(from_date=format_short_date(original_day), to_date=format_short_date(snapped)))
        self.refresh_rows()


def build_launch_request(output_dir: Path) -> PdfRequest:
    start = date(2026, 5, 29)
    finish = date(2026, 7, 13)
    items = reschedule_work_items(place_initial_work_items(start), start)
    return PdfRequest(
        site_name="テストリフォーム",
        branch_name="東京支店",
        staff_name="山田太郎",
        phone="03-0000-0000",
        start_date=start,
        end_date=finish,
        save_folder=str(output_dir),
        work_items=tuple(item for item in items if is_output_item(item)),
    )


def run_launch_check() -> int:
    start = date(2026, 5, 29)
    finish = date(2026, 7, 13)
    days = days_between(start, finish)
    weeks = iter_period_weeks(start, finish)
    if len(days) != 46:
        raise RuntimeError("46 day fixture failed")
    if not any(finish in week for week in weeks):
        raise RuntimeError("completion cell fixture failed")
    items = reschedule_work_items(place_initial_work_items(start), start)
    if any(day.weekday() >= 5 for item in items if is_output_item(item) for day in workdays_between(item.start_date, item.end_date)):  # type: ignore[arg-type]
        raise RuntimeError("weekend allocation fixture failed")
    if any(daily_entries(day, items) for day in days if day.weekday() >= 5):
        raise RuntimeError("weekend display fixture failed")

    dragged, adjusted, snapped = move_item_to_start(items[0], date(2026, 6, 7))
    if not adjusted or snapped != date(2026, 6, 8) or dragged.start_date != date(2026, 6, 8):
        raise RuntimeError("drag move fixture failed")
    resized, resized_adjusted, resize_snapped = resize_item_to_end(items[1], date(2026, 6, 13))
    if not resized_adjusted or resize_snapped != date(2026, 6, 12) or resized.end_date != date(2026, 6, 12):
        raise RuntimeError("drag resize fixture failed")

    start_resize_item = WorkItem(True, "開始変更工程", date(2026, 6, 5), date(2026, 6, 10), STATUS_OPTIONS[0], 4)
    start_resized, start_resize_adjusted, start_resize_snapped = resize_item_to_start(start_resize_item, date(2026, 6, 7))
    if not start_resize_adjusted or start_resize_snapped != date(2026, 6, 8) or start_resized.start_date != date(2026, 6, 8):
        raise RuntimeError("drag start resize fixture failed")

    week_split_item = WorkItem(True, "週またぎ工程", date(2026, 6, 5), date(2026, 6, 9), STATUS_OPTIONS[1], 3)
    segments = band_segments_for_item(week_split_item, date(2026, 6, 1), date(2026, 6, 14))
    if segments != [[date(2026, 6, 5)], [date(2026, 6, 8), date(2026, 6, 9)]]:
        raise RuntimeError("pdf band segment fixture failed")

    if reference_day_for_start(date(2026, 5, 31)) != date(2026, 7, 14):
        raise RuntimeError("reference day fixture failed")
    if not is_period_within_limit(start, start + timedelta(days=55)) or is_period_within_limit(start, start + timedelta(days=56)):
        raise RuntimeError("max day limit fixture failed")

    reordered = items[:]
    moved = reordered.pop(2)
    reordered.insert(0, moved)
    if reordered[0].name != items[2].name:
        raise RuntimeError("row reorder fixture failed")

    fixed_items = reschedule_work_items(place_initial_work_items(start), start)
    fixed_items[0].status = "完了"
    fixed_items[0].start_date = date(2026, 6, 5)
    fixed_items[0].end_date = date(2026, 6, 5)
    reflowed = reschedule_work_items(fixed_items, start)
    if reflowed[0].start_date != date(2026, 6, 5) or reflowed[0].end_date != date(2026, 6, 5):
        raise RuntimeError("done item fixed fixture failed")

    with tempfile.TemporaryDirectory(dir=app_dir()) as temp_dir:
        output_dir = Path(temp_dir)
        request = build_launch_request(output_dir)
        output_path = generate_pdf(request)
        if not output_path.exists() or output_path.stat().st_size <= 0:
            raise RuntimeError("pdf output fixture failed")
        config_file = output_dir / CONFIG_NAME
        values = {
            "branch_name": "東京支店",
            "staff_name": "山田太郎",
            "phone": "03-0000-0000",
            "save_folder": str(output_dir),
        }
        if not save_config(values, config_file):
            raise RuntimeError("config save fixture failed")
        if load_config(config_file) != values:
            raise RuntimeError("config load fixture failed")
        project_file = output_dir / PROJECT_NAME
        project_values = build_project_data(
            site_name="保存テスト現場",
            branch_name="東京支店",
            staff_name="山田太郎",
            phone="03-0000-0000",
            start_date_text=format_date(start),
            end_date_text=format_date(finish),
            save_folder=str(output_dir),
            work_items=items[:3],
        )
        if not save_project(project_values, project_file):
            raise RuntimeError("project save fixture failed")
        loaded_project = load_project(project_file)
        restored_items = deserialize_work_items(loaded_project.get("work_items"))
        if loaded_project.get("site_name") != "保存テスト現場" or not restored_items or restored_items[0].name != items[0].name:
            raise RuntimeError("project load fixture failed")
        opened: list[str] = []
        if not open_output_folder(output_path, opener=opened.append):
            raise RuntimeError("open folder fixture failed")
        if not opened or Path(opened[0]) != output_path.parent.resolve():
            raise RuntimeError("open folder path fixture failed")

        print(UI_TEXT["launch_check_ok"])
        print(UI_TEXT["launch_check_pdf"].format(path=output_path))
        print(UI_TEXT["launch_check_page"])
        print(UI_TEXT["launch_check_span"].format(days=len(days), start=format_date(start), finish=format_date(finish), weeks=len(weeks)))
        print(UI_TEXT["launch_check_completion"])
        print(UI_TEXT["launch_check_weekend"])
        print(UI_TEXT["launch_check_drag"])
        print(UI_TEXT["launch_check_rows"])
        print(UI_TEXT["launch_check_reschedule"])
        print(UI_TEXT["launch_check_config"])
        print(UI_TEXT["launch_check_project"])
        print(UI_TEXT["launch_check_pdf_bands"])
        print(UI_TEXT["launch_check_reference"])
        print(UI_TEXT["launch_check_limit"])
        print(UI_TEXT["launch_check_folder"])
    return 0


def main() -> None:
    if "--launch-check" in sys.argv:
        raise SystemExit(run_launch_check())
    set_windows_app_id()
    root = tk.Tk()
    ReformProgressApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
