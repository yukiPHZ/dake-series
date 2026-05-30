# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import json
import os
import queue
import re
import sys
import threading
import time
import webbrowser
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tkinter import font as tkfont
import tkinter as tk
from tkinter import messagebox, ttk

try:
    import ctypes
except Exception:
    ctypes = None

try:
    from watchdog.events import FileSystemEventHandler
    from watchdog.observers import Observer
except Exception:
    FileSystemEventHandler = None
    Observer = None


APP_NAME = "DAKE Dashboard"
WINDOW_TITLE = "DAKE Dashboard"
APP_FOLDER_NAME = "DAKE_App_Dashboard"
README_NAME = "README.md"
RELEASE_BODY_NAME = "release_body.md"
BOOTH_PRODUCT_NAME = "booth_product.txt"
BOOTH_READY_NAME = "booth_ready"
SCREENSHOT_RELATIVE = Path("assets") / "screenshot.webp"
BOOTH_THUMBNAIL_RELATIVE = Path("assets") / "booth_thumbnail.jpg"
DIST_DIR_NAME = "dist"

UI_TEXT = {
    "header_title": "DAKE Dashboard",
    "header_subtitle": "README正本から現在地を見る",
    "last_loaded_waiting": "最終読込: 未読込",
    "last_loaded_value": "最終読込: {time}",
    "button_reload": "再読み込み",
    "button_open_folder": "フォルダを開く",
    "button_open_readme": "README.md を開く",
    "button_open_release": "Release URL を開く",
    "button_release_missing": "Release URL なし",
    "summary_total": "総アプリ数",
    "summary_implementation": "実装中",
    "summary_distribution_ready": "配布準備",
    "summary_released": "配布中",
    "summary_booth_ready": "BOOTH準備済み",
    "summary_needs_review": "要確認",
    "filter_all": "全部",
    "filter_implementation": "実装中",
    "filter_distribution_ready": "配布準備",
    "filter_released": "配布中",
    "filter_booth": "BOOTH",
    "filter_needs_review": "要確認",
    "search_label": "検索",
    "search_placeholder": "フォルダ名・表示名・説明を検索",
    "list_title": "アプリ一覧",
    "detail_title": "正本情報",
    "detail_empty": "アプリを選択すると、README正本と不足項目を確認できます。",
    "detail_meta_title": "DAKE_META",
    "detail_missing_title": "不足しているもの",
    "detail_next_title": "次に必要そうな作業",
    "detail_files_title": "実在チェック",
    "column_status": "状態",
    "column_folder": "フォルダ名",
    "column_display": "表示名",
    "column_release": "Release",
    "column_booth": "BOOTH素材",
    "column_screenshot": "screenshot",
    "column_exe": "exe",
    "column_updated": "最終更新",
    "status_implementation": "実装中",
    "status_distribution_ready": "配布準備",
    "status_released": "配布中",
    "status_booth_ready": "BOOTH準備済み",
    "status_needs_review": "要確認",
    "status_loading": "読み込み中",
    "status_ready": "正本を読み込みました",
    "status_error": "読み込みで確認が必要です",
    "status_launch_check_ok": "LAUNCH CHECK OK",
    "value_yes": "あり",
    "value_no": "なし",
    "value_partial": "一部",
    "value_unset": "未設定",
    "value_unknown": "不明",
    "value_none": "なし",
    "count_line": "{visible} / {total} 件を表示",
    "missing_none": "不足は見つかっていません。",
    "missing_readme": "README.md",
    "missing_dake_meta": "DAKE_META JSON",
    "missing_release_body": "release_body.md",
    "missing_screenshot": "assets/screenshot.webp",
    "missing_booth_thumbnail": "assets/booth_thumbnail.jpg",
    "missing_booth_product": "booth_product.txt",
    "missing_booth_ready": "booth_ready/",
    "missing_dist_exe": "dist/*.exe",
    "missing_release_url": "release_url",
    "meta_folder_name": "folder_name",
    "meta_display_name": "display_name",
    "meta_launcher_title": "launcher_title",
    "meta_launcher_description": "launcher_description",
    "meta_site_title": "site_title",
    "meta_site_description": "site_description",
    "meta_update_summary": "update_summary",
    "meta_exe_name": "exe_name",
    "meta_release_url": "release_url",
    "meta_screenshot_path": "screenshot_path",
    "meta_status": "status",
    "meta_show_in_launcher": "show_in_launcher",
    "meta_show_on_site": "show_on_site",
    "file_readme": "README.md",
    "file_release_body": "release_body.md",
    "file_screenshot": "assets/screenshot.webp",
    "file_booth_thumbnail": "assets/booth_thumbnail.jpg",
    "file_booth_product": "booth_product.txt",
    "file_booth_ready": "booth_ready/",
    "file_dist_exe": "dist/*.exe",
    "file_release_url": "release_url",
    "issue_readme_missing": "README.md が見つかりません。",
    "issue_readme_read_error": "README.md を読み取れませんでした: {error}",
    "issue_meta_missing": "DAKE_META が見つかりません。",
    "issue_meta_json_error": "DAKE_META JSON を解析できませんでした: {error}",
    "issue_meta_type_error": "DAKE_META がJSONオブジェクトではありません。",
    "issue_scan_error": "フォルダ確認でエラーが発生しました: {error}",
    "next_review_readme": "README.md と DAKE_META を確認する",
    "next_fix_meta": "DAKE_META JSON を正しい形式に直す",
    "next_build_exe": "build.bat で dist/*.exe を作る",
    "next_capture_screenshot": "assets/screenshot.webp を用意する",
    "next_prepare_release": "Release URL をREADME正本へ反映する",
    "next_prepare_booth": "BOOTH素材の不足を確認する",
    "next_release_check": "配布中情報と正本の差分を確認する",
    "next_booth_check": "BOOTH準備済み素材と正本の対応を確認する",
    "next_internal_no_release": "内部アプリとして release_url と show_on_site を空のまま維持する",
    "next_no_action": "現時点で明確な追加作業はありません。",
    "dialog_error_title": "エラー",
    "dialog_notice_title": "確認",
    "dialog_open_failed": "開けませんでした。\n\n{path}\n\n{error}",
    "dialog_missing_path": "対象が見つかりません。\n\n{path}",
    "dialog_release_missing": "Release URL が設定されていません。",
    "qpsc_card_title": "QPSC通知カード",
    "qpsc_card_subtitle": "README正本の動きを監視中",
    "qpsc_new_apps": "新規アプリ検出",
    "qpsc_distribution_ready": "配布準備",
    "qpsc_needs_review": "要確認",
    "qpsc_ship_line": "正式出荷ライン到達",
    "watch_status_watchdog": "watchdog監視: ON",
    "watch_status_polling": "watchdog未導入 / 30秒自動読込: ON",
    "watch_status_error": "watchdog監視: 起動できませんでした",
    "notification_reloaded": "{folder} を再読込しました",
    "footer_note": "一般公開・BOOTH登録・dakeapp.com掲載を行わない内部端末",
}

THEME = {
    "bg": "#070A10",
    "panel": "#0D1422",
    "panel_alt": "#101827",
    "panel_soft": "#121B2C",
    "border": "#243046",
    "border_active": "#415C94",
    "text": "#E8EDF7",
    "muted": "#A5B1C5",
    "quiet": "#667085",
    "accent": "#6B8CFF",
    "accent_hover": "#7FA0FF",
    "accent_soft": "#182645",
    "purple": "#9A8CFF",
    "success": "#74D7A3",
    "warning": "#FFD18A",
    "danger": "#FF8C9A",
    "review": "#F7A6FF",
    "selection": "#1C2B49",
    "input": "#090E18",
}

STATUS_THEME = {
    "implementation": ("#1B2333", THEME["muted"]),
    "distribution_ready": ("#202033", THEME["warning"]),
    "released": ("#152A25", THEME["success"]),
    "booth_ready": ("#1D213B", THEME["purple"]),
    "needs_review": ("#301B34", THEME["review"]),
}

FONT_CANDIDATES = ["BIZ UDPGothic", "Yu Gothic UI", "Meiryo", "MS Gothic"]
META_FIELD_KEYS = (
    "folder_name",
    "display_name",
    "launcher_title",
    "launcher_description",
    "site_title",
    "site_description",
    "update_summary",
    "exe_name",
    "release_url",
    "screenshot_path",
    "status",
    "show_in_launcher",
    "show_on_site",
)
FILTER_KEYS = ("all", "implementation", "distribution_ready", "released", "booth", "needs_review")
DAKE_META_PATTERN = re.compile(
    r"##\s*DAKE_META\s*```(?:json)?\s*(\{.*?\})\s*```",
    re.IGNORECASE | re.DOTALL,
)
WORKER_POLL_MS = 80
LAUNCH_CHECK_TIMEOUT_MS = 8000
AUTO_RELOAD_MS = 30000
WATCH_DEBOUNCE_MS = 700
NOTIFICATION_HIDE_MS = 3600
WATCHED_FILENAMES = {README_NAME, RELEASE_BODY_NAME, BOOTH_PRODUCT_NAME}
WATCHED_DIR_NAMES = {"assets", DIST_DIR_NAME}


@dataclass(frozen=True)
class FileChecks:
    has_readme: bool
    has_release_body: bool
    has_screenshot: bool
    has_booth_thumbnail: bool
    has_booth_product: bool
    has_booth_ready: bool
    has_dist_exe: bool
    has_release_url: bool
    dist_exes: tuple[Path, ...]

    @property
    def booth_materials_ready(self) -> bool:
        return self.has_booth_product and self.has_booth_thumbnail and self.has_booth_ready

    @property
    def booth_materials_partial(self) -> bool:
        return self.has_booth_product or self.has_booth_thumbnail or self.has_booth_ready


@dataclass(frozen=True)
class AppRecord:
    folder_name: str
    folder_path: Path
    meta_fields: dict[str, str]
    checks: FileChecks
    status_key: str
    missing_keys: tuple[str, ...]
    issue_messages: tuple[str, ...]
    last_modified: float

    @property
    def display_name(self) -> str:
        for key in ("display_name", "launcher_title", "site_title"):
            value = self.meta_fields.get(key, "").strip()
            if value:
                return value
        return self.folder_name

    @property
    def release_url(self) -> str:
        return self.meta_fields.get("release_url", "").strip()

    @property
    def status_text(self) -> str:
        return UI_TEXT[f"status_{self.status_key}"]


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        if exe_dir.name.lower() == DIST_DIR_NAME:
            return exe_dir.parent
        return exe_dir
    return Path(__file__).resolve().parent


def apps_root() -> Path:
    return app_dir().parent


def series_root() -> Path:
    return apps_root().parent


def icon_path() -> Path:
    if getattr(sys, "frozen", False):
        meipass = Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
        bundled_icon = meipass / "dake_icon.ico"
        if bundled_icon.exists():
            return bundled_icon
    return series_root() / "02_assets" / "dake_icon.ico"


def set_windows_app_id() -> None:
    if not sys.platform.startswith("win") or ctypes is None:
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("dake.app.dashboard")
    except Exception:
        pass


def apply_window_icon(window: tk.Misc) -> None:
    try:
        icon = icon_path()
        if icon.exists():
            window.iconbitmap(str(icon))
            window.iconbitmap(default=str(icon))
    except Exception:
        pass


def choose_font_family(root: tk.Tk) -> str:
    try:
        available = set(tkfont.families(root))
    except Exception:
        return "TkDefaultFont"
    for family in FONT_CANDIDATES:
        if family in available:
            return family
    return "TkDefaultFont"


def safe_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return UI_TEXT["value_yes"] if value else UI_TEXT["value_no"]
    return str(value).strip()


def safe_bool(value: object) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.strip().lower() in {"true", "1", "yes", "on"}
    if isinstance(value, (int, float)):
        return bool(value)
    return False


def read_text_utf8(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError:
        return path.read_text(encoding="utf-8", errors="replace")


def extract_dake_meta(readme_path: Path) -> tuple[dict[str, object], tuple[str, ...]]:
    if not readme_path.exists():
        return {}, (UI_TEXT["issue_readme_missing"],)

    try:
        content = read_text_utf8(readme_path)
    except OSError as exc:
        return {}, (UI_TEXT["issue_readme_read_error"].format(error=exc),)

    match = DAKE_META_PATTERN.search(content)
    if not match:
        return {}, (UI_TEXT["issue_meta_missing"],)

    try:
        loaded = json.loads(match.group(1))
    except json.JSONDecodeError as exc:
        return {}, (UI_TEXT["issue_meta_json_error"].format(error=exc),)

    if not isinstance(loaded, dict):
        return {}, (UI_TEXT["issue_meta_type_error"],)
    return loaded, ()


def existing_dist_exes(folder: Path) -> tuple[Path, ...]:
    dist_dir = folder / DIST_DIR_NAME
    if not dist_dir.exists() or not dist_dir.is_dir():
        return ()
    try:
        return tuple(sorted(path for path in dist_dir.glob("*.exe") if path.is_file()))
    except OSError:
        return ()


def latest_mtime(paths: list[Path], fallback: Path) -> float:
    values: list[float] = []
    for path in paths:
        try:
            if path.exists():
                values.append(path.stat().st_mtime)
        except OSError:
            continue
    try:
        values.append(fallback.stat().st_mtime)
    except OSError:
        pass
    return max(values) if values else 0.0


def format_datetime(timestamp: float) -> str:
    if timestamp <= 0:
        return UI_TEXT["value_unknown"]
    return datetime.fromtimestamp(timestamp).strftime("%Y-%m-%d %H:%M")


def bool_label(value: bool) -> str:
    return UI_TEXT["value_yes"] if value else UI_TEXT["value_no"]


def booth_label(checks: FileChecks) -> str:
    if checks.booth_materials_ready:
        return UI_TEXT["value_yes"]
    if checks.booth_materials_partial:
        return UI_TEXT["value_partial"]
    return UI_TEXT["value_no"]


def build_meta_fields(folder: Path, meta: dict[str, object]) -> dict[str, str]:
    fields = {key: safe_text(meta.get(key, "")) for key in META_FIELD_KEYS}
    fields["folder_name"] = fields["folder_name"] or folder.name
    if not fields["display_name"]:
        fields["display_name"] = fields["launcher_title"] or fields["site_title"] or folder.name
    return fields


def classify_status(folder: Path, meta_fields: dict[str, str], checks: FileChecks, issues: tuple[str, ...]) -> str:
    if not checks.has_readme or issues:
        return "needs_review"

    if checks.booth_materials_ready:
        return "booth_ready"

    if checks.has_release_url and checks.has_dist_exe and checks.has_screenshot:
        return "released"

    if checks.has_dist_exe and checks.has_screenshot and not checks.has_release_url:
        return "distribution_ready"

    status = meta_fields.get("status", "").strip().lower()
    if status in {"draft", "working", "development", "develop", "dev", "wip", "internal"}:
        return "implementation"

    if folder.name == APP_FOLDER_NAME:
        return "implementation"

    if checks.has_readme and (not checks.has_release_url or not checks.has_dist_exe):
        return "implementation"

    return "needs_review"


def build_missing_keys(checks: FileChecks, issues: tuple[str, ...]) -> tuple[str, ...]:
    missing: list[str] = []
    if not checks.has_readme:
        missing.append("readme")
    if issues and checks.has_readme:
        missing.append("dake_meta")
    if not checks.has_release_body:
        missing.append("release_body")
    if not checks.has_screenshot:
        missing.append("screenshot")
    if not checks.has_booth_thumbnail:
        missing.append("booth_thumbnail")
    if not checks.has_booth_product:
        missing.append("booth_product")
    if not checks.has_booth_ready:
        missing.append("booth_ready")
    if not checks.has_dist_exe:
        missing.append("dist_exe")
    if not checks.has_release_url:
        missing.append("release_url")
    return tuple(missing)


def scan_folder(folder: Path) -> AppRecord:
    readme_path = folder / README_NAME
    meta, issues = extract_dake_meta(readme_path)
    meta_fields = build_meta_fields(folder, meta)
    release_url = meta_fields.get("release_url", "").strip()
    dist_exes = existing_dist_exes(folder)
    checks = FileChecks(
        has_readme=readme_path.exists(),
        has_release_body=(folder / RELEASE_BODY_NAME).exists(),
        has_screenshot=(folder / SCREENSHOT_RELATIVE).exists(),
        has_booth_thumbnail=(folder / BOOTH_THUMBNAIL_RELATIVE).exists(),
        has_booth_product=(folder / BOOTH_PRODUCT_NAME).exists(),
        has_booth_ready=(folder / BOOTH_READY_NAME).is_dir(),
        has_dist_exe=bool(dist_exes),
        has_release_url=bool(release_url),
        dist_exes=dist_exes,
    )
    status_key = classify_status(folder, meta_fields, checks, issues)
    missing_keys = build_missing_keys(checks, issues)
    last_modified = latest_mtime(
        [
            readme_path,
            folder / RELEASE_BODY_NAME,
            folder / "main.py",
            folder / "build.bat",
            folder / SCREENSHOT_RELATIVE,
            folder / BOOTH_THUMBNAIL_RELATIVE,
            folder / BOOTH_PRODUCT_NAME,
            *dist_exes,
        ],
        folder,
    )
    return AppRecord(
        folder_name=folder.name,
        folder_path=folder,
        meta_fields=meta_fields,
        checks=checks,
        status_key=status_key,
        missing_keys=missing_keys,
        issue_messages=issues,
        last_modified=last_modified,
    )


def error_record(folder: Path, exc: Exception) -> AppRecord:
    meta_fields = build_meta_fields(folder, {})
    checks = FileChecks(
        has_readme=(folder / README_NAME).exists(),
        has_release_body=False,
        has_screenshot=False,
        has_booth_thumbnail=False,
        has_booth_product=False,
        has_booth_ready=False,
        has_dist_exe=False,
        has_release_url=False,
        dist_exes=(),
    )
    issue = UI_TEXT["issue_scan_error"].format(error=exc)
    return AppRecord(
        folder_name=folder.name,
        folder_path=folder,
        meta_fields=meta_fields,
        checks=checks,
        status_key="needs_review",
        missing_keys=build_missing_keys(checks, (issue,)),
        issue_messages=(issue,),
        last_modified=latest_mtime([folder / README_NAME], folder),
    )


def scan_apps(root: Path) -> list[AppRecord]:
    if not root.exists() or not root.is_dir():
        return []

    records: list[AppRecord] = []
    try:
        folders = sorted((path for path in root.iterdir() if path.is_dir()), key=lambda path: path.name.lower())
    except OSError:
        return records

    for folder in folders:
        try:
            records.append(scan_folder(folder))
        except Exception as exc:
            records.append(error_record(folder, exc))
    return records


def formal_ship_line_reached(record: AppRecord) -> bool:
    return (
        record.checks.has_release_url
        and record.checks.has_dist_exe
        and record.checks.has_screenshot
        and record.checks.booth_materials_ready
    )


def watched_folder_for_path(path: Path, root: Path) -> str | None:
    try:
        relative = path.resolve().relative_to(root.resolve())
    except (OSError, ValueError):
        return None

    parts = relative.parts
    if not parts:
        return None

    folder_name = parts[0]
    if folder_name.startswith("."):
        return None

    if len(parts) == 1:
        return folder_name

    if parts[-1] in WATCHED_FILENAMES:
        return folder_name

    if len(parts) >= 2 and parts[1] in WATCHED_DIR_NAMES:
        return folder_name

    return None


def record_map(records: list[AppRecord]) -> dict[str, AppRecord]:
    return {record.folder_name: record for record in records}


class DashboardApp:
    def __init__(self, root: tk.Tk, launch_check: bool = False) -> None:
        self.root = root
        self.launch_check = launch_check
        self.font_family = choose_font_family(root)
        self.worker_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.worker_thread: threading.Thread | None = None
        self.records: list[AppRecord] = []
        self.visible_records: list[AppRecord] = []
        self.record_by_iid: dict[str, AppRecord] = {}
        self.previous_record_map: dict[str, AppRecord] = {}
        self.selected_record: AppRecord | None = None
        self.filter_key = "all"
        self.summary_vars = {
            "total": tk.StringVar(value="0"),
            "implementation": tk.StringVar(value="0"),
            "distribution_ready": tk.StringVar(value="0"),
            "released": tk.StringVar(value="0"),
            "booth_ready": tk.StringVar(value="0"),
            "needs_review": tk.StringVar(value="0"),
        }
        self.qpsc_vars = {
            "new_apps": tk.StringVar(value="0"),
            "distribution_ready": tk.StringVar(value="0"),
            "needs_review": tk.StringVar(value="0"),
            "ship_line": tk.StringVar(value="0"),
        }
        self.search_var = tk.StringVar()
        self.last_loaded_var = tk.StringVar(value=UI_TEXT["last_loaded_waiting"])
        self.status_var = tk.StringVar(value=UI_TEXT["last_loaded_waiting"])
        self.watch_status_var = tk.StringVar(value=UI_TEXT["watch_status_polling"])
        self.count_var = tk.StringVar(value=UI_TEXT["count_line"].format(visible=0, total=0))
        self.filter_buttons: dict[str, tk.Button] = {}
        self.watch_observer = None
        self.watch_pending_folder: str | None = None
        self.watch_debounce_job: str | None = None
        self.last_watch_event_at = 0.0
        self.pending_reload_folder: str | None = None

        self.root.title(WINDOW_TITLE)
        self.root.geometry("1280x820")
        self.root.minsize(1080, 700)
        self.root.configure(bg=THEME["bg"])
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        apply_window_icon(self.root)
        self.configure_styles()
        self.build_ui()
        self.build_notification()
        self.search_var.trace_add("write", lambda *_args: self.apply_filters())
        self.root.after(WORKER_POLL_MS, self.poll_worker)
        self.reload_data(source="startup")
        if not self.launch_check:
            self.start_watchdog()
            self.schedule_auto_reload()
        if self.launch_check:
            self.root.after(LAUNCH_CHECK_TIMEOUT_MS, self.finish_launch_check)

    def configure_styles(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(
            "Dashboard.Treeview",
            background=THEME["panel"],
            fieldbackground=THEME["panel"],
            foreground=THEME["text"],
            bordercolor=THEME["border"],
            lightcolor=THEME["border"],
            darkcolor=THEME["border"],
            rowheight=34,
            font=(self.font_family, 10),
        )
        style.map(
            "Dashboard.Treeview",
            background=[("selected", THEME["selection"])],
            foreground=[("selected", THEME["text"])],
        )
        style.configure(
            "Dashboard.Treeview.Heading",
            background=THEME["panel_alt"],
            foreground=THEME["muted"],
            bordercolor=THEME["border"],
            lightcolor=THEME["border"],
            darkcolor=THEME["border"],
            relief="flat",
            font=(self.font_family, 9, "bold"),
        )
        style.map(
            "Dashboard.Treeview.Heading",
            background=[("active", THEME["accent_soft"])],
            foreground=[("active", THEME["text"])],
        )
        style.configure(
            "Dashboard.Vertical.TScrollbar",
            background=THEME["panel_alt"],
            troughcolor=THEME["bg"],
            bordercolor=THEME["bg"],
            arrowcolor=THEME["muted"],
            lightcolor=THEME["panel_alt"],
            darkcolor=THEME["panel_alt"],
            relief="flat",
            width=12,
        )

    def build_ui(self) -> None:
        shell = tk.Frame(self.root, bg=THEME["bg"])
        shell.pack(fill="both", expand=True, padx=26, pady=22)

        self.build_header(shell)
        self.build_summary(shell)
        self.build_qpsc_card(shell)
        self.build_controls(shell)
        self.build_body(shell)
        self.build_footer(shell)

    def build_header(self, parent: tk.Frame) -> None:
        header = tk.Frame(parent, bg=THEME["bg"])
        header.pack(fill="x", pady=(0, 18))
        header.grid_columnconfigure(0, weight=1)

        title_area = tk.Frame(header, bg=THEME["bg"])
        title_area.grid(row=0, column=0, sticky="w")

        tk.Label(
            title_area,
            text=UI_TEXT["header_title"],
            bg=THEME["bg"],
            fg=THEME["text"],
            font=(self.font_family, 26, "bold"),
        ).pack(anchor="w")

        tk.Label(
            title_area,
            text=UI_TEXT["header_subtitle"],
            bg=THEME["bg"],
            fg=THEME["muted"],
            font=(self.font_family, 11),
        ).pack(anchor="w", pady=(4, 0))

        action_area = tk.Frame(header, bg=THEME["bg"])
        action_area.grid(row=0, column=1, sticky="e")

        tk.Label(
            action_area,
            textvariable=self.last_loaded_var,
            bg=THEME["bg"],
            fg=THEME["muted"],
            font=(self.font_family, 10),
        ).pack(anchor="e", pady=(0, 8))

        self.reload_button = self.make_button(
            action_area,
            UI_TEXT["button_reload"],
            self.reload_data,
            primary=True,
        )
        self.reload_button.pack(anchor="e")

    def build_summary(self, parent: tk.Frame) -> None:
        summary = tk.Frame(parent, bg=THEME["bg"])
        summary.pack(fill="x", pady=(0, 16))
        for index in range(6):
            summary.grid_columnconfigure(index, weight=1, uniform="summary")

        cards = [
            ("total", "summary_total"),
            ("implementation", "summary_implementation"),
            ("distribution_ready", "summary_distribution_ready"),
            ("released", "summary_released"),
            ("booth_ready", "summary_booth_ready"),
            ("needs_review", "summary_needs_review"),
        ]
        for index, (key, label_key) in enumerate(cards):
            card = tk.Frame(
                summary,
                bg=THEME["panel_alt"],
                highlightbackground=THEME["border"],
                highlightthickness=1,
                bd=0,
            )
            card.grid(row=0, column=index, sticky="ew", padx=(0 if index == 0 else 8, 0), ipady=3)
            tk.Label(
                card,
                text=UI_TEXT[label_key],
                bg=THEME["panel_alt"],
                fg=THEME["muted"],
                font=(self.font_family, 9),
            ).pack(anchor="w", padx=14, pady=(10, 2))
            tk.Label(
                card,
                textvariable=self.summary_vars[key],
                bg=THEME["panel_alt"],
                fg=THEME["text"],
                font=(self.font_family, 22, "bold"),
            ).pack(anchor="w", padx=14, pady=(0, 10))

    def build_qpsc_card(self, parent: tk.Frame) -> None:
        card = tk.Frame(
            parent,
            bg=THEME["panel"],
            highlightbackground=THEME["border"],
            highlightthickness=1,
            bd=0,
        )
        card.pack(fill="x", pady=(0, 16))
        card.grid_columnconfigure(1, weight=1)

        title_area = tk.Frame(card, bg=THEME["panel"])
        title_area.grid(row=0, column=0, sticky="w", padx=16, pady=13)

        tk.Label(
            title_area,
            text=UI_TEXT["qpsc_card_title"],
            bg=THEME["panel"],
            fg=THEME["text"],
            font=(self.font_family, 12, "bold"),
        ).pack(anchor="w")

        tk.Label(
            title_area,
            text=UI_TEXT["qpsc_card_subtitle"],
            bg=THEME["panel"],
            fg=THEME["muted"],
            font=(self.font_family, 9),
        ).pack(anchor="w", pady=(3, 0))

        metrics = tk.Frame(card, bg=THEME["panel"])
        metrics.grid(row=0, column=1, sticky="e", padx=(0, 16), pady=10)

        items = [
            ("new_apps", "qpsc_new_apps"),
            ("distribution_ready", "qpsc_distribution_ready"),
            ("needs_review", "qpsc_needs_review"),
            ("ship_line", "qpsc_ship_line"),
        ]
        for index, (key, label_key) in enumerate(items):
            item = tk.Frame(metrics, bg=THEME["panel_soft"], padx=12, pady=7)
            item.grid(row=0, column=index, sticky="e", padx=(0 if index == 0 else 8, 0))
            tk.Label(
                item,
                text=UI_TEXT[label_key],
                bg=THEME["panel_soft"],
                fg=THEME["muted"],
                font=(self.font_family, 8, "bold"),
            ).pack(anchor="w")
            tk.Label(
                item,
                textvariable=self.qpsc_vars[key],
                bg=THEME["panel_soft"],
                fg=THEME["accent_hover"],
                font=(self.font_family, 16, "bold"),
            ).pack(anchor="w")

    def build_controls(self, parent: tk.Frame) -> None:
        controls = tk.Frame(parent, bg=THEME["bg"])
        controls.pack(fill="x", pady=(0, 14))
        controls.grid_columnconfigure(1, weight=1)

        filters = tk.Frame(controls, bg=THEME["bg"])
        filters.grid(row=0, column=0, sticky="w", padx=(0, 18))
        for key in FILTER_KEYS:
            label_key = "filter_booth" if key == "booth" else f"filter_{key}"
            button = self.make_button(filters, UI_TEXT[label_key], lambda value=key: self.set_filter(value))
            button.pack(side="left", padx=(0, 7))
            self.filter_buttons[key] = button
        self.update_filter_buttons()

        search_area = tk.Frame(controls, bg=THEME["bg"])
        search_area.grid(row=0, column=1, sticky="ew")
        search_area.grid_columnconfigure(1, weight=1)
        tk.Label(
            search_area,
            text=UI_TEXT["search_label"],
            bg=THEME["bg"],
            fg=THEME["muted"],
            font=(self.font_family, 10, "bold"),
        ).grid(row=0, column=0, sticky="w", padx=(0, 8))

        self.search_entry = tk.Entry(
            search_area,
            textvariable=self.search_var,
            bg=THEME["input"],
            fg=THEME["text"],
            insertbackground=THEME["accent"],
            relief="flat",
            highlightbackground=THEME["border"],
            highlightcolor=THEME["border_active"],
            highlightthickness=1,
            font=(self.font_family, 11),
        )
        self.search_entry.grid(row=0, column=1, sticky="ew", ipady=7)

    def build_body(self, parent: tk.Frame) -> None:
        body = tk.PanedWindow(
            parent,
            orient="horizontal",
            bg=THEME["bg"],
            sashwidth=8,
            sashrelief="flat",
            bd=0,
        )
        body.pack(fill="both", expand=True)

        left = tk.Frame(body, bg=THEME["panel"], highlightbackground=THEME["border"], highlightthickness=1, bd=0)
        right = tk.Frame(body, bg=THEME["panel"], highlightbackground=THEME["border"], highlightthickness=1, bd=0)
        body.add(left, width=780, minsize=620)
        body.add(right, width=420, minsize=360)

        self.build_list(left)
        self.build_detail(right)

    def build_list(self, parent: tk.Frame) -> None:
        header = tk.Frame(parent, bg=THEME["panel"])
        header.pack(fill="x", padx=16, pady=(14, 10))
        header.grid_columnconfigure(0, weight=1)

        tk.Label(
            header,
            text=UI_TEXT["list_title"],
            bg=THEME["panel"],
            fg=THEME["text"],
            font=(self.font_family, 13, "bold"),
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            header,
            textvariable=self.count_var,
            bg=THEME["panel"],
            fg=THEME["muted"],
            font=(self.font_family, 10),
        ).grid(row=0, column=1, sticky="e")

        table_shell = tk.Frame(parent, bg=THEME["panel"])
        table_shell.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        table_shell.grid_rowconfigure(0, weight=1)
        table_shell.grid_columnconfigure(0, weight=1)

        columns = ("status", "folder", "display", "release", "booth", "screenshot", "exe", "updated")
        self.tree = ttk.Treeview(
            table_shell,
            columns=columns,
            show="headings",
            selectmode="browse",
            style="Dashboard.Treeview",
        )
        self.tree.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(
            table_shell,
            orient="vertical",
            command=self.tree.yview,
            style="Dashboard.Vertical.TScrollbar",
        )
        scrollbar.grid(row=0, column=1, sticky="ns", padx=(8, 0))
        self.tree.configure(yscrollcommand=scrollbar.set)

        headings = {
            "status": UI_TEXT["column_status"],
            "folder": UI_TEXT["column_folder"],
            "display": UI_TEXT["column_display"],
            "release": UI_TEXT["column_release"],
            "booth": UI_TEXT["column_booth"],
            "screenshot": UI_TEXT["column_screenshot"],
            "exe": UI_TEXT["column_exe"],
            "updated": UI_TEXT["column_updated"],
        }
        widths = {
            "status": 110,
            "folder": 170,
            "display": 170,
            "release": 76,
            "booth": 86,
            "screenshot": 88,
            "exe": 64,
            "updated": 126,
        }
        for column in columns:
            self.tree.heading(column, text=headings[column])
            anchor = "center" if column in {"status", "release", "booth", "screenshot", "exe", "updated"} else "w"
            self.tree.column(column, width=widths[column], minwidth=54, stretch=column in {"folder", "display"}, anchor=anchor)

        for key, (_bg, fg) in STATUS_THEME.items():
            self.tree.tag_configure(key, foreground=fg)

        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)

    def build_detail(self, parent: tk.Frame) -> None:
        container = tk.Frame(parent, bg=THEME["panel"])
        container.pack(fill="both", expand=True, padx=16, pady=14)
        container.grid_columnconfigure(0, weight=1)
        container.grid_rowconfigure(4, weight=1)

        tk.Label(
            container,
            text=UI_TEXT["detail_title"],
            bg=THEME["panel"],
            fg=THEME["text"],
            font=(self.font_family, 13, "bold"),
        ).grid(row=0, column=0, sticky="w")

        self.detail_name_var = tk.StringVar(value=UI_TEXT["detail_empty"])
        self.detail_status_var = tk.StringVar(value=UI_TEXT["value_unset"])
        self.detail_badge = tk.Label(
            container,
            textvariable=self.detail_status_var,
            bg=THEME["panel_soft"],
            fg=THEME["muted"],
            font=(self.font_family, 10, "bold"),
            padx=10,
            pady=5,
        )
        self.detail_badge.grid(row=1, column=0, sticky="w", pady=(12, 7))

        tk.Label(
            container,
            textvariable=self.detail_name_var,
            bg=THEME["panel"],
            fg=THEME["text"],
            font=(self.font_family, 12, "bold"),
            wraplength=360,
            justify="left",
        ).grid(row=2, column=0, sticky="ew", pady=(0, 10))

        button_row = tk.Frame(container, bg=THEME["panel"])
        button_row.grid(row=3, column=0, sticky="ew", pady=(0, 12))
        self.open_folder_button = self.make_button(button_row, UI_TEXT["button_open_folder"], self.open_selected_folder)
        self.open_folder_button.pack(side="left", padx=(0, 8))
        self.open_readme_button = self.make_button(button_row, UI_TEXT["button_open_readme"], self.open_selected_readme)
        self.open_readme_button.pack(side="left", padx=(0, 8))
        self.open_release_button = self.make_button(button_row, UI_TEXT["button_open_release"], self.open_selected_release)
        self.open_release_button.pack(side="left")

        detail_area = tk.Frame(container, bg=THEME["panel"])
        detail_area.grid(row=4, column=0, sticky="nsew")
        detail_area.grid_columnconfigure(0, weight=1)
        detail_area.grid_rowconfigure(1, weight=1)
        detail_area.grid_rowconfigure(3, weight=1)
        detail_area.grid_rowconfigure(5, weight=1)
        detail_area.grid_rowconfigure(7, weight=1)

        self.meta_text = self.create_detail_text(detail_area, 0, UI_TEXT["detail_meta_title"])
        self.files_text = self.create_detail_text(detail_area, 2, UI_TEXT["detail_files_title"])
        self.missing_text = self.create_detail_text(detail_area, 4, UI_TEXT["detail_missing_title"])
        self.next_text = self.create_detail_text(detail_area, 6, UI_TEXT["detail_next_title"])
        self.update_detail(None)

    def create_detail_text(self, parent: tk.Frame, row: int, title: str) -> tk.Text:
        tk.Label(
            parent,
            text=title,
            bg=THEME["panel"],
            fg=THEME["muted"],
            font=(self.font_family, 9, "bold"),
        ).grid(row=row, column=0, sticky="w", pady=(0 if row == 0 else 10, 5))
        text = tk.Text(
            parent,
            bg=THEME["input"],
            fg=THEME["text"],
            insertbackground=THEME["accent"],
            relief="flat",
            highlightbackground=THEME["border"],
            highlightcolor=THEME["border"],
            highlightthickness=1,
            height=5,
            wrap="word",
            font=(self.font_family, 10),
            padx=10,
            pady=8,
        )
        text.grid(row=row + 1, column=0, sticky="nsew")
        text.configure(state="disabled")
        return text

    def build_footer(self, parent: tk.Frame) -> None:
        footer = tk.Frame(parent, bg=THEME["bg"])
        footer.pack(fill="x", pady=(12, 0))
        footer.grid_columnconfigure(1, weight=1)
        tk.Label(
            footer,
            textvariable=self.status_var,
            bg=THEME["bg"],
            fg=THEME["muted"],
            font=(self.font_family, 9),
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            footer,
            textvariable=self.watch_status_var,
            bg=THEME["bg"],
            fg=THEME["quiet"],
            font=(self.font_family, 9),
        ).grid(row=0, column=1, sticky="w", padx=(18, 0))
        tk.Label(
            footer,
            text=UI_TEXT["footer_note"],
            bg=THEME["bg"],
            fg=THEME["quiet"],
            font=(self.font_family, 9),
        ).grid(row=0, column=2, sticky="e")

    def build_notification(self) -> None:
        self.notification_var = tk.StringVar(value="")
        self.notification_frame = tk.Frame(
            self.root,
            bg=THEME["accent_soft"],
            highlightbackground=THEME["border_active"],
            highlightthickness=1,
            bd=0,
        )
        tk.Label(
            self.notification_frame,
            textvariable=self.notification_var,
            bg=THEME["accent_soft"],
            fg=THEME["text"],
            font=(self.font_family, 10, "bold"),
            padx=16,
            pady=11,
        ).pack()
        self.notification_frame.place_forget()

    def make_button(self, parent: tk.Misc, label: str, command, primary: bool = False) -> tk.Button:
        bg = THEME["accent_soft"] if primary else THEME["panel_alt"]
        fg = THEME["text"] if primary else THEME["muted"]
        active_bg = THEME["accent"] if primary else THEME["selection"]
        active_fg = "#FFFFFF" if primary else THEME["text"]
        return tk.Button(
            parent,
            text=label,
            command=command,
            bg=bg,
            fg=fg,
            activebackground=active_bg,
            activeforeground=active_fg,
            disabledforeground=THEME["quiet"],
            relief="flat",
            bd=0,
            padx=12,
            pady=7,
            cursor="hand2",
            font=(self.font_family, 9, "bold"),
        )

    def set_filter(self, key: str) -> None:
        self.filter_key = key
        self.update_filter_buttons()
        self.apply_filters()

    def update_filter_buttons(self) -> None:
        for key, button in self.filter_buttons.items():
            selected = key == self.filter_key
            button.configure(
                bg=THEME["accent_soft"] if selected else THEME["panel_alt"],
                fg=THEME["text"] if selected else THEME["muted"],
                activebackground=THEME["accent"] if selected else THEME["selection"],
                activeforeground="#FFFFFF" if selected else THEME["text"],
            )

    def reload_data(self, source: str = "manual", changed_folder: str | None = None) -> None:
        if self.worker_thread is not None and self.worker_thread.is_alive():
            self.pending_reload_folder = changed_folder or self.pending_reload_folder
            return
        self.reload_button.configure(state="disabled")
        self.status_var.set(UI_TEXT["status_loading"])
        self.worker_thread = threading.Thread(target=self.scan_worker, args=(source, changed_folder), daemon=True)
        self.worker_thread.start()

    def scan_worker(self, source: str, changed_folder: str | None) -> None:
        records = scan_apps(apps_root())
        self.worker_queue.put(("scan_done", {"records": records, "source": source, "changed_folder": changed_folder}))

    def poll_worker(self) -> None:
        try:
            while True:
                event, payload = self.worker_queue.get_nowait()
                if event == "scan_done":
                    self.handle_scan_done(payload)
                elif event == "watch_event":
                    self.handle_watch_event(payload)
        except queue.Empty:
            pass
        self.root.after(WORKER_POLL_MS, self.poll_worker)

    def handle_scan_done(self, payload: object) -> None:
        if isinstance(payload, dict):
            records = payload.get("records", [])
            source = str(payload.get("source", "manual"))
            changed_folder = payload.get("changed_folder")
        else:
            records = payload
            source = "manual"
            changed_folder = None

        previous_map = dict(self.previous_record_map)
        self.records = list(records) if isinstance(records, list) else []
        current_map = record_map(self.records)
        self.reload_button.configure(state="normal")
        loaded_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.last_loaded_var.set(UI_TEXT["last_loaded_value"].format(time=loaded_at))
        self.status_var.set(UI_TEXT["status_ready"])
        self.update_summary(previous_map, current_map)
        self.apply_filters()
        self.previous_record_map = current_map
        if source == "watch" and isinstance(changed_folder, str) and changed_folder:
            self.show_reload_notification(changed_folder)
        if self.pending_reload_folder:
            folder = self.pending_reload_folder
            self.pending_reload_folder = None
            self.reload_data(source="watch", changed_folder=folder)
        if self.launch_check:
            self.root.after(150, self.finish_launch_check)

    def finish_launch_check(self) -> None:
        self.stop_watchdog()
        try:
            self.root.destroy()
        except tk.TclError:
            pass

    def update_summary(
        self,
        previous_map: dict[str, AppRecord] | None = None,
        current_map: dict[str, AppRecord] | None = None,
    ) -> None:
        counts = {key: 0 for key in self.summary_vars}
        counts["total"] = len(self.records)
        for record in self.records:
            if record.status_key in counts:
                counts[record.status_key] += 1
        for key, value in counts.items():
            self.summary_vars[key].set(str(value))

        previous_map = previous_map or {}
        current_map = current_map or record_map(self.records)
        new_app_count = len(set(current_map) - set(previous_map)) if previous_map else 0
        self.qpsc_vars["new_apps"].set(str(new_app_count))
        self.qpsc_vars["distribution_ready"].set(str(counts["distribution_ready"]))
        self.qpsc_vars["needs_review"].set(str(counts["needs_review"]))
        self.qpsc_vars["ship_line"].set(str(sum(1 for record in self.records if formal_ship_line_reached(record))))

    def schedule_auto_reload(self) -> None:
        self.root.after(AUTO_RELOAD_MS, self.auto_reload)

    def auto_reload(self) -> None:
        self.reload_data(source="auto")
        self.schedule_auto_reload()

    def start_watchdog(self) -> None:
        if Observer is None or FileSystemEventHandler is None:
            self.watch_status_var.set(UI_TEXT["watch_status_polling"])
            return
        try:
            handler = self.create_watchdog_handler()
            observer = Observer()
            observer.schedule(handler, str(apps_root()), recursive=True)
            observer.daemon = True
            observer.start()
            self.watch_observer = observer
            self.watch_status_var.set(UI_TEXT["watch_status_watchdog"])
        except Exception:
            self.watch_status_var.set(UI_TEXT["watch_status_error"])

    def create_watchdog_handler(self):
        app = self
        root_path = apps_root()

        class DashboardWatchHandler(FileSystemEventHandler):
            def on_any_event(self, event) -> None:
                if getattr(event, "is_directory", False) and getattr(event, "event_type", "") not in {"created", "deleted", "moved"}:
                    return
                paths = [Path(getattr(event, "src_path", ""))]
                dest_path = getattr(event, "dest_path", "")
                if dest_path:
                    paths.append(Path(dest_path))
                for path in paths:
                    folder_name = watched_folder_for_path(path, root_path)
                    if folder_name:
                        app.worker_queue.put(("watch_event", folder_name))
                        return

        return DashboardWatchHandler()

    def stop_watchdog(self) -> None:
        observer = self.watch_observer
        self.watch_observer = None
        if observer is None:
            return
        try:
            observer.stop()
            observer.join(timeout=1.0)
        except Exception:
            pass

    def handle_watch_event(self, payload: object) -> None:
        folder_name = str(payload).strip()
        if not folder_name:
            return
        self.watch_pending_folder = folder_name
        self.last_watch_event_at = time.monotonic()
        if self.watch_debounce_job is None:
            self.watch_debounce_job = self.root.after(WATCH_DEBOUNCE_MS, self.flush_watch_reload)

    def flush_watch_reload(self) -> None:
        self.watch_debounce_job = None
        if time.monotonic() - self.last_watch_event_at < WATCH_DEBOUNCE_MS / 1000:
            self.watch_debounce_job = self.root.after(WATCH_DEBOUNCE_MS, self.flush_watch_reload)
            return
        folder_name = self.watch_pending_folder
        self.watch_pending_folder = None
        if folder_name:
            self.reload_data(source="watch", changed_folder=folder_name)

    def show_reload_notification(self, folder_name: str) -> None:
        self.notification_var.set(UI_TEXT["notification_reloaded"].format(folder=folder_name))
        self.notification_frame.place(relx=1.0, rely=1.0, anchor="se", x=-22, y=-22)
        self.root.after(NOTIFICATION_HIDE_MS, self.notification_frame.place_forget)

    def on_close(self) -> None:
        self.stop_watchdog()
        self.root.destroy()

    def apply_filters(self) -> None:
        query = self.search_var.get().strip().lower()
        filtered: list[AppRecord] = []
        for record in self.records:
            if not self.matches_filter(record):
                continue
            if query and not self.matches_query(record, query):
                continue
            filtered.append(record)

        self.visible_records = filtered
        self.render_tree()
        self.count_var.set(UI_TEXT["count_line"].format(visible=len(filtered), total=len(self.records)))
        if self.selected_record not in filtered:
            self.update_detail(filtered[0] if filtered else None)
            if filtered:
                first_iid = str(filtered[0].folder_path)
                self.tree.selection_set(first_iid)
                self.tree.focus(first_iid)

    def matches_filter(self, record: AppRecord) -> bool:
        if self.filter_key == "all":
            return True
        if self.filter_key == "booth":
            return record.status_key == "booth_ready"
        return record.status_key == self.filter_key

    def matches_query(self, record: AppRecord, query: str) -> bool:
        haystack = " ".join(
            [
                record.folder_name,
                record.display_name,
                record.meta_fields.get("launcher_title", ""),
                record.meta_fields.get("launcher_description", ""),
                record.meta_fields.get("site_title", ""),
                record.meta_fields.get("site_description", ""),
                record.meta_fields.get("update_summary", ""),
                record.meta_fields.get("exe_name", ""),
                record.meta_fields.get("status", ""),
            ]
        ).lower()
        return query in haystack

    def render_tree(self) -> None:
        selected_path = str(self.selected_record.folder_path) if self.selected_record else ""
        self.tree.delete(*self.tree.get_children())
        self.record_by_iid.clear()
        for record in self.visible_records:
            iid = str(record.folder_path)
            self.record_by_iid[iid] = record
            self.tree.insert(
                "",
                "end",
                iid=iid,
                values=(
                    record.status_text,
                    record.folder_name,
                    record.display_name,
                    bool_label(record.checks.has_release_url),
                    booth_label(record.checks),
                    bool_label(record.checks.has_screenshot),
                    bool_label(record.checks.has_dist_exe),
                    format_datetime(record.last_modified),
                ),
                tags=(record.status_key,),
            )
        if selected_path and selected_path in self.record_by_iid:
            self.tree.selection_set(selected_path)
            self.tree.focus(selected_path)

    def on_tree_select(self, _event=None) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        record = self.record_by_iid.get(selection[0])
        self.update_detail(record)

    def update_detail(self, record: AppRecord | None) -> None:
        self.selected_record = record
        if record is None:
            self.detail_name_var.set(UI_TEXT["detail_empty"])
            self.detail_status_var.set(UI_TEXT["value_unset"])
            self.detail_badge.configure(bg=THEME["panel_soft"], fg=THEME["muted"])
            self.set_text(self.meta_text, UI_TEXT["detail_empty"])
            self.set_text(self.files_text, UI_TEXT["detail_empty"])
            self.set_text(self.missing_text, UI_TEXT["detail_empty"])
            self.set_text(self.next_text, UI_TEXT["detail_empty"])
            self.set_detail_buttons_state(False)
            return

        badge_bg, badge_fg = STATUS_THEME.get(record.status_key, (THEME["panel_soft"], THEME["muted"]))
        self.detail_status_var.set(record.status_text)
        self.detail_badge.configure(bg=badge_bg, fg=badge_fg)
        self.detail_name_var.set(f"{record.folder_name}\n{record.display_name}")
        self.set_text(self.meta_text, self.build_meta_detail(record))
        self.set_text(self.files_text, self.build_file_detail(record))
        self.set_text(self.missing_text, self.build_missing_detail(record))
        self.set_text(self.next_text, self.build_next_detail(record))
        self.set_detail_buttons_state(True)

    def set_detail_buttons_state(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        self.open_folder_button.configure(state=state)
        self.open_readme_button.configure(state=state)
        release_state = "normal" if enabled and self.selected_record and self.selected_record.release_url else "disabled"
        self.open_release_button.configure(state=release_state)
        if release_state == "disabled":
            self.open_release_button.configure(text=UI_TEXT["button_release_missing"])
        else:
            self.open_release_button.configure(text=UI_TEXT["button_open_release"])

    def set_text(self, widget: tk.Text, value: str) -> None:
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("end", value)
        widget.configure(state="disabled")

    def build_meta_detail(self, record: AppRecord) -> str:
        lines: list[str] = []
        for key in META_FIELD_KEYS:
            label = UI_TEXT[f"meta_{key}"]
            value = record.meta_fields.get(key, "").strip() or UI_TEXT["value_unset"]
            lines.append(f"{label}: {value}")
        if record.issue_messages:
            lines.append("")
            lines.extend(record.issue_messages)
        return "\n".join(lines)

    def build_file_detail(self, record: AppRecord) -> str:
        checks = record.checks
        rows = [
            ("file_readme", checks.has_readme),
            ("file_release_body", checks.has_release_body),
            ("file_screenshot", checks.has_screenshot),
            ("file_booth_thumbnail", checks.has_booth_thumbnail),
            ("file_booth_product", checks.has_booth_product),
            ("file_booth_ready", checks.has_booth_ready),
            ("file_dist_exe", checks.has_dist_exe),
            ("file_release_url", checks.has_release_url),
        ]
        lines = [f"{UI_TEXT[label_key]}: {bool_label(value)}" for label_key, value in rows]
        if checks.dist_exes:
            lines.append("")
            lines.extend(path.name for path in checks.dist_exes)
        return "\n".join(lines)

    def build_missing_detail(self, record: AppRecord) -> str:
        if not record.missing_keys:
            return UI_TEXT["missing_none"]
        return "\n".join(f"- {UI_TEXT[f'missing_{key}']}" for key in record.missing_keys)

    def build_next_detail(self, record: AppRecord) -> str:
        tasks: list[str] = []
        if record.status_key == "needs_review":
            tasks.append(UI_TEXT["next_review_readme"])
            if "dake_meta" in record.missing_keys:
                tasks.append(UI_TEXT["next_fix_meta"])
        elif record.status_key == "implementation":
            if not record.checks.has_dist_exe:
                tasks.append(UI_TEXT["next_build_exe"])
            if not record.checks.has_screenshot:
                tasks.append(UI_TEXT["next_capture_screenshot"])
            if record.folder_name == APP_FOLDER_NAME:
                tasks.append(UI_TEXT["next_internal_no_release"])
            elif not record.checks.has_release_url:
                tasks.append(UI_TEXT["next_prepare_release"])
        elif record.status_key == "distribution_ready":
            tasks.append(UI_TEXT["next_prepare_release"])
            if not record.checks.booth_materials_ready:
                tasks.append(UI_TEXT["next_prepare_booth"])
        elif record.status_key == "released":
            tasks.append(UI_TEXT["next_release_check"])
            if not record.checks.booth_materials_ready:
                tasks.append(UI_TEXT["next_prepare_booth"])
        elif record.status_key == "booth_ready":
            tasks.append(UI_TEXT["next_booth_check"])
            if not record.checks.has_release_url:
                tasks.append(UI_TEXT["next_prepare_release"])

        if not tasks:
            tasks.append(UI_TEXT["next_no_action"])
        return "\n".join(f"- {task}" for task in tasks)

    def open_selected_folder(self) -> None:
        if self.selected_record is not None:
            self.open_path(self.selected_record.folder_path)

    def open_selected_readme(self) -> None:
        if self.selected_record is not None:
            self.open_path(self.selected_record.folder_path / README_NAME)

    def open_selected_release(self) -> None:
        if self.selected_record is None:
            return
        url = self.selected_record.release_url
        if not url:
            messagebox.showinfo(UI_TEXT["dialog_notice_title"], UI_TEXT["dialog_release_missing"], parent=self.root)
            return
        webbrowser.open(url)

    def open_path(self, path: Path) -> None:
        if not path.exists():
            messagebox.showinfo(
                UI_TEXT["dialog_notice_title"],
                UI_TEXT["dialog_missing_path"].format(path=path),
                parent=self.root,
            )
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(path))
            else:
                webbrowser.open(path.as_uri())
        except OSError as exc:
            messagebox.showerror(
                UI_TEXT["dialog_error_title"],
                UI_TEXT["dialog_open_failed"].format(path=path, error=exc),
                parent=self.root,
            )


def run_gui(launch_check: bool = False) -> int:
    app: DashboardApp | None = None
    try:
        set_windows_app_id()
        root = tk.Tk()
        app = DashboardApp(root, launch_check=launch_check)
        root.mainloop()
        if launch_check:
            print(UI_TEXT["status_launch_check_ok"])
        return 0
    except Exception as exc:
        if app is not None:
            try:
                app.root.destroy()
            except Exception:
                pass
        if launch_check:
            print(f"LAUNCH CHECK FAILED: {exc}", file=sys.stderr)
            return 1
        raise


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--launch-check", action="store_true")
    args = parser.parse_args()
    return run_gui(launch_check=args.launch_check)


if __name__ == "__main__":
    raise SystemExit(main())
