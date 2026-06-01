# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
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
from datetime import datetime
from pathlib import Path
from tkinter import font as tkfont
import tkinter as tk
from tkinter import messagebox, ttk
from urllib.parse import urlparse

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
SCREENSHOT_JPG_RELATIVE = Path("assets") / "screenshot.jpg"
BOOTH_THUMBNAIL_RELATIVE = Path("assets") / "booth_thumbnail.jpg"
DIST_DIR_NAME = "dist"
BOOTH_ASSIST_FOLDER_NAME = "DAKE_BOOTH_Assist"
BOOTH_ASSIST_EXE_NAME = "DakeBOOTH_Assist.exe"
URL_BLOCKLIST = ("javascript:", "file:", "data:", "sk-", "token", "secret", "api_key")

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
    "button_next_work": "次の作業へ",
    "button_open_booth_assist": "BOOTHアシストで開く",
    "button_open_booth_ready": "booth_readyを開く",
    "button_open_booth_product": "booth_product.txtを開く",
    "button_open_booth_thumbnail": "booth_thumbnailを開く",
    "button_open_screenshot_jpg": "screenshot.jpgを開く",
    "button_open_screenshot_webp": "screenshot.webpを開く",
    "button_missing": "未作成",
    "button_open": "開く",
    "button_show_location": "場所を表示",
    "workspace_links_title": "作業リンク",
    "link_folder": "フォルダ",
    "link_readme": "README",
    "link_release_body": "release_body",
    "link_assets": "assets",
    "link_screenshot_webp": "screenshot.webp",
    "link_screenshot_jpg": "screenshot.jpg",
    "link_booth_thumbnail": "booth_thumbnail",
    "link_booth_product": "booth_product",
    "link_booth_ready": "booth_ready",
    "link_dist": "dist",
    "link_exe": "exe",
    "link_release_url": "Release URL",
    "link_github_url": "GitHub",
    "link_missing": "未作成",
    "link_unavailable": "未取得",
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
    "column_next_step": "次工程",
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
    "booth_status_title": "BOOTH状況 {ready}/3",
    "booth_status_ready": "✓ {label}",
    "booth_status_missing": "✗ {label}",
    "booth_missing_title": "不足:",
    "booth_ready_label_thumbnail": "booth_thumbnail.jpg",
    "booth_ready_label_product": "booth_product.txt",
    "booth_ready_label_ready": "booth_ready/",
    "next_step_review": "要確認",
    "next_step_readme": "README整備",
    "next_step_screenshot": "スクショ作成",
    "next_step_release": "Release作成",
    "next_step_booth_materials": "BOOTH素材作成",
    "next_step_booth": "BOOTH登録",
    "next_step_released": "配布中",
    "next_step_internal": "内部",
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
    "qpsc_booth_working": "BOOTH作業中: {folder}",
    "qpsc_new_apps": "新規アプリ検出",
    "qpsc_distribution_ready": "配布準備",
    "qpsc_needs_review": "要確認",
    "qpsc_ship_line": "正式出荷ライン到達",
    "qpsc_unshipped": "未出荷",
    "qpsc_booth_materials_missing": "BOOTH素材不足",
    "qpsc_booth_thumbnail_missing": "thumbnail未作成",
    "qpsc_booth_product_missing": "booth_product未作成",
    "qpsc_booth_ready_missing": "booth_ready未作成",
    "qpsc_booth_registration_ready": "BOOTH登録可能",
    "qpsc_booth_missing": "BOOTH未準備",
    "qpsc_release_missing": "Release未作成",
    "git_card_title": "Git状態",
    "git_card_subtitle": "DAKE_series 全体",
    "git_branch": "branch: {value}",
    "git_latest": "latest: {value}",
    "git_uncommitted": "未コミット: {value}",
    "git_untracked": "未追跡: {value}",
    "git_remote_clean": "remote: 同期",
    "git_push_waiting": "push待ち: {value}",
    "git_pull_waiting": "pull待ち: {value}",
    "git_dashboard": "Dashboard: {value}",
    "git_dashboard_clean": "clean",
    "git_dashboard_dirty": "dirty",
    "git_error": "Git状態を取得できません",
    "shipment_title": "正式出荷ライン",
    "shipment_rate": "正式出荷率 {percent}%",
    "shipment_internal": "内部アプリ",
    "shipment_missing_title": "不足:",
    "next_candidates_title": "次にやる候補",
    "next_candidates_empty": "候補はありません。",
    "next_candidate_line": "{folder}\n{reason}",
    "candidate_needs_review": "要確認",
    "candidate_meta_broken": "README / DAKE_META破損",
    "candidate_release_missing": "dist/*.exeあり / release_urlなし",
    "candidate_booth_missing": "Release URLあり / BOOTH未準備",
    "candidate_booth_registration": "次工程: BOOTH登録",
    "candidate_site_unknown": "BOOTH素材あり / サイト反映確認が不明",
    "candidate_screenshot_missing": "screenshot.webpなし",
    "candidate_release_body_missing": "release_body.mdなし",
    "candidate_internal": "内部アプリ / 出荷対象外",
    "next_process_booth": "次工程: BOOTH登録",
    "next_process_release": "次工程: Release作成",
    "next_process_screenshot": "次工程: スクショ作成",
    "next_process_readme": "次工程: README整備",
    "next_process_booth_materials": "次工程: BOOTH素材作成",
    "next_process_internal": "次工程: 内部アプリ",
    "next_process_unknown": "次工程: 確認",
    "booth_assist_missing": "BOOTHアシストが見つかりません",
    "booth_assist_launched": "BOOTHアシストを起動しました: {folder}",
    "booth_links_title": "BOOTH作業リンク",
    "notice_opened": "{target}を開きました",
    "notice_missing_target": "対象が見つかりません",
    "notice_url_opened": "Release URLを開きました",
    "notice_url_failed": "URLを開けませんでした",
    "notice_exe_location": "exeの場所を表示しました",
    "watch_status_watchdog": "watchdog監視: ON",
    "watch_status_polling": "watchdog未導入 / 30秒自動読込: ON",
    "watch_status_error": "watchdog監視: 起動できませんでした",
    "meta_app_type": "app_type",
    "meta_completion_goal": "completion_goal",
    "column_app_type": "種別",
    "column_completion_goal": "完成条件",
    "filter_market": "市場向け",
    "filter_system": "QPSC系",
    "filter_personal": "専用",
    "filter_frozen": "凍結",
    "app_type_market": "市場向け",
    "app_type_system": "QPSC / 補助脳系",
    "app_type_personal": "ユキズ専用",
    "app_type_frozen": "凍結",
    "app_type_archived": "保管",
    "app_type_unknown": "未分類",
    "completion_goal_formal_release": "正式出荷",
    "completion_goal_local_ready": "ローカル運用",
    "completion_goal_system_ready": "システム稼働",
    "completion_goal_reference_ready": "正本提示",
    "completion_goal_frozen_closed": "凍結完了",
    "completion_goal_unknown": "未設定",
    "qpsc_market": "市場向け",
    "qpsc_system": "QPSC / 補助脳系",
    "qpsc_personal": "ユキズ専用",
    "qpsc_frozen": "凍結",
    "status_system_ready": "システム稼働",
    "status_local_ready": "ローカル運用",
    "status_reference_ready": "正本提示",
    "status_frozen_closed": "凍結完了",
    "shipment_non_formal": "{app_type} / {goal}",
    "shipment_non_formal_note": "正式出荷ラインではなく、この完成条件で判定します。",
    "missing_non_formal": "この分類ではBOOTH / dakeapp.com素材不足を主警告にしません。",
    "next_non_formal_goal": "完成条件: {goal}",
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
    "system_ready": ("#142337", THEME["accent_hover"]),
    "local_ready": ("#1A2530", THEME["muted"]),
    "reference_ready": ("#202538", THEME["purple"]),
    "frozen_closed": ("#222531", THEME["quiet"]),
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
    "app_type",
    "completion_goal",
    "show_in_launcher",
    "show_on_site",
)
APP_TYPE_DEFAULT = "market"
COMPLETION_GOAL_DEFAULT = "formal_release"
APP_TYPE_KEYS = ("market", "system", "personal", "frozen", "archived", "unknown")
COMPLETION_GOAL_KEYS = (
    "formal_release",
    "local_ready",
    "system_ready",
    "reference_ready",
    "frozen_closed",
    "unknown",
)
NON_PUBLIC_STATUSES = {"internal", "frozen", "draft", "experimental", "private"}
FILTER_KEYS = (
    "all",
    "implementation",
    "distribution_ready",
    "released",
    "booth",
    "needs_review",
    "market",
    "system",
    "personal",
    "frozen",
)
DAKE_META_PATTERN = re.compile(
    r"##\s*DAKE_META\s*```(?:json)?\s*(\{.*?\})\s*```",
    re.IGNORECASE | re.DOTALL,
)
WORKER_POLL_MS = 80
LAUNCH_CHECK_TIMEOUT_MS = 8000
AUTO_RELOAD_MS = 30000
WATCH_DEBOUNCE_MS = 700
NOTIFICATION_HIDE_MS = 3600
BOOTH_HIGHLIGHT_MS = 7000
WATCHED_FILENAMES = {README_NAME, RELEASE_BODY_NAME, BOOTH_PRODUCT_NAME}
WATCHED_DIR_NAMES = {"assets", DIST_DIR_NAME}
GIT_TIMEOUT_SECONDS = 2.5
SHIPMENT_MISSING_KEYS = (
    "readme",
    "release_body",
    "screenshot",
    "booth_thumbnail",
    "booth_product",
    "booth_ready",
    "dist_exe",
    "release_url",
)
BOOTH_MATERIAL_KEYS = ("booth_thumbnail", "booth_product", "booth_ready")
WORKSPACE_LINK_KEYS = (
    "folder",
    "readme",
    "release_body",
    "assets",
    "screenshot_webp",
    "screenshot_jpg",
    "booth_thumbnail",
    "booth_product",
    "booth_ready",
    "dist",
    "exe",
    "release_url",
    "github_url",
)


@dataclass(frozen=True)
class GitStatus:
    branch: str
    latest: str
    uncommitted_count: int
    untracked_count: int
    ahead_count: int
    behind_count: int
    dashboard_dirty: bool
    error: str = ""


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

    @property
    def booth_materials_count(self) -> int:
        return sum((self.has_booth_thumbnail, self.has_booth_product, self.has_booth_ready))


@dataclass(frozen=True)
class AppRecord:
    folder_name: str
    folder_path: Path
    meta_fields: dict[str, str]
    checks: FileChecks
    app_type: str
    completion_goal: str
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
        return value.strip().lower() in {"true", "1", "yes", "on", UI_TEXT["value_yes"].lower()}
    if isinstance(value, (int, float)):
        return bool(value)
    return False


def meta_bool(record: AppRecord, key: str) -> bool:
    return safe_bool(record.meta_fields.get(key, ""))


def meta_false(record: AppRecord, key: str) -> bool:
    value = record.meta_fields.get(key, "").strip().lower()
    return value in {"false", "0", "no", "off", UI_TEXT["value_no"].lower()}


def normalized_choice(value: str, allowed: tuple[str, ...], default: str) -> str:
    normalized = value.strip().lower()
    return normalized if normalized in allowed else default


def normalize_app_type(value: str) -> str:
    return normalized_choice(value, APP_TYPE_KEYS, APP_TYPE_DEFAULT)


def normalize_completion_goal(value: str) -> str:
    return normalized_choice(value, COMPLETION_GOAL_KEYS, COMPLETION_GOAL_DEFAULT)


def app_type_label(value: str) -> str:
    key = normalize_app_type(value)
    return UI_TEXT.get(f"app_type_{key}", UI_TEXT["app_type_unknown"])


def completion_goal_label(value: str) -> str:
    key = normalize_completion_goal(value)
    return UI_TEXT.get(f"completion_goal_{key}", UI_TEXT["completion_goal_unknown"])


def is_formal_release_meta(meta_fields: dict[str, str]) -> bool:
    status = meta_fields.get("status", "").strip().lower()
    if status in NON_PUBLIC_STATUSES:
        return False
    app_type = normalize_app_type(meta_fields.get("app_type", ""))
    completion_goal = normalize_completion_goal(meta_fields.get("completion_goal", ""))
    return app_type == "market" and completion_goal == "formal_release"


def is_formal_release_app(record: AppRecord) -> bool:
    return is_formal_release_meta(record.meta_fields)


def is_internal_app(record: AppRecord) -> bool:
    status = record.meta_fields.get("status", "").strip().lower()
    if record.folder_name == APP_FOLDER_NAME or status == "internal":
        return True
    if record.app_type in {"system", "personal", "frozen", "archived"}:
        return True
    return meta_false(record, "show_on_site")


def shipment_missing_keys(record: AppRecord) -> tuple[str, ...]:
    if not is_formal_release_app(record):
        return ()
    return tuple(key for key in SHIPMENT_MISSING_KEYS if key in record.missing_keys)


def shipment_rate(record: AppRecord) -> int | None:
    if not is_formal_release_app(record):
        return None
    missing_count = len(shipment_missing_keys(record))
    done_count = len(SHIPMENT_MISSING_KEYS) - missing_count
    return round(done_count / len(SHIPMENT_MISSING_KEYS) * 100)


def hidden_subprocess_kwargs() -> dict[str, object]:
    if not sys.platform.startswith("win"):
        return {}

    kwargs: dict[str, object] = {}
    try:
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= getattr(subprocess, "STARTF_USESHOWWINDOW", 1)
        startupinfo.wShowWindow = 0
        kwargs["startupinfo"] = startupinfo
    except Exception:
        pass

    create_no_window = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    if create_no_window:
        kwargs["creationflags"] = create_no_window
    return kwargs


def git_run(repo: Path, args: list[str]) -> tuple[str, str]:
    try:
        result = subprocess.run(
            ["git", *args],
            cwd=str(repo),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=GIT_TIMEOUT_SECONDS,
            check=False,
            **hidden_subprocess_kwargs(),
        )
    except Exception as exc:
        return "", str(exc)
    if result.returncode != 0:
        return "", (result.stderr or result.stdout).strip()
    return result.stdout.strip(), ""


def parse_remote_counts(status_sb: str) -> tuple[int, int]:
    ahead = 0
    behind = 0
    first_line = status_sb.splitlines()[0] if status_sb else ""
    ahead_match = re.search(r"ahead\s+(\d+)", first_line)
    behind_match = re.search(r"behind\s+(\d+)", first_line)
    if ahead_match:
        ahead = int(ahead_match.group(1))
    if behind_match:
        behind = int(behind_match.group(1))
    return ahead, behind


def read_git_status(repo: Path) -> GitStatus:
    branch, branch_error = git_run(repo, ["rev-parse", "--abbrev-ref", "HEAD"])
    latest, latest_error = git_run(repo, ["log", "-1", "--pretty=%h"])
    short_status, short_error = git_run(repo, ["status", "--short"])
    status_sb, status_error = git_run(repo, ["status", "-sb"])

    error = branch_error or latest_error or short_error or status_error
    if error:
        return GitStatus(
            branch=UI_TEXT["value_unknown"],
            latest=UI_TEXT["value_unknown"],
            uncommitted_count=0,
            untracked_count=0,
            ahead_count=0,
            behind_count=0,
            dashboard_dirty=False,
            error=error,
        )

    lines = [line for line in short_status.splitlines() if line.strip()]
    untracked_count = sum(1 for line in lines if line.startswith("??"))
    uncommitted_count = len(lines) - untracked_count
    dashboard_dirty = any("01_apps/DAKE_App_Dashboard/" in line.replace("\\", "/") for line in lines)
    ahead_count, behind_count = parse_remote_counts(status_sb)
    return GitStatus(
        branch=branch or UI_TEXT["value_unknown"],
        latest=latest or UI_TEXT["value_unknown"],
        uncommitted_count=uncommitted_count,
        untracked_count=untracked_count,
        ahead_count=ahead_count,
        behind_count=behind_count,
        dashboard_dirty=dashboard_dirty,
    )


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
    return f"{checks.booth_materials_count}/3"


def booth_product_candidates(folder: Path) -> tuple[Path, ...]:
    return (
        folder / BOOTH_PRODUCT_NAME,
        folder / BOOTH_READY_NAME / BOOTH_PRODUCT_NAME,
    )


def find_booth_product_path(folder: Path) -> Path | None:
    for candidate in booth_product_candidates(folder):
        if candidate.is_file():
            return candidate
    return None


def booth_missing_keys(checks: FileChecks) -> tuple[str, ...]:
    missing: list[str] = []
    if not checks.has_booth_thumbnail:
        missing.append("booth_thumbnail")
    if not checks.has_booth_product:
        missing.append("booth_product")
    if not checks.has_booth_ready:
        missing.append("booth_ready")
    return tuple(missing)


def booth_item_label(key: str) -> str:
    labels = {
        "booth_thumbnail": UI_TEXT["booth_ready_label_thumbnail"],
        "booth_product": UI_TEXT["booth_ready_label_product"],
        "booth_ready": UI_TEXT["booth_ready_label_ready"],
    }
    return labels.get(key, key)


def build_meta_fields(folder: Path, meta: dict[str, object]) -> dict[str, str]:
    fields = {key: safe_text(meta.get(key, "")) for key in META_FIELD_KEYS}
    fields["folder_name"] = fields["folder_name"] or folder.name
    if not fields["display_name"]:
        fields["display_name"] = fields["launcher_title"] or fields["site_title"] or folder.name
    fields["app_type"] = normalize_app_type(fields.get("app_type", ""))
    fields["completion_goal"] = normalize_completion_goal(fields.get("completion_goal", ""))
    return fields


def classify_status(folder: Path, meta_fields: dict[str, str], checks: FileChecks, issues: tuple[str, ...]) -> str:
    if not checks.has_readme or issues:
        return "needs_review"

    if not is_formal_release_meta(meta_fields):
        goal = normalize_completion_goal(meta_fields.get("completion_goal", ""))
        app_type = normalize_app_type(meta_fields.get("app_type", ""))
        status = meta_fields.get("status", "").strip().lower()
        if goal in {"system_ready", "local_ready", "reference_ready", "frozen_closed"}:
            return goal
        if app_type == "frozen" or status == "frozen":
            return "frozen_closed"
        if app_type == "system":
            return "system_ready"
        if app_type == "personal":
            return "local_ready"
        return "implementation"

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


def build_missing_keys(
    checks: FileChecks,
    issues: tuple[str, ...],
    formal_release: bool,
) -> tuple[str, ...]:
    missing: list[str] = []
    if not checks.has_readme:
        missing.append("readme")
    if issues and checks.has_readme:
        missing.append("dake_meta")
    if not formal_release:
        return tuple(missing)
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
    booth_product_path = find_booth_product_path(folder)
    checks = FileChecks(
        has_readme=readme_path.exists(),
        has_release_body=(folder / RELEASE_BODY_NAME).exists(),
        has_screenshot=(folder / SCREENSHOT_RELATIVE).exists(),
        has_booth_thumbnail=(folder / BOOTH_THUMBNAIL_RELATIVE).exists(),
        has_booth_product=booth_product_path is not None,
        has_booth_ready=(folder / BOOTH_READY_NAME).is_dir(),
        has_dist_exe=bool(dist_exes),
        has_release_url=bool(release_url),
        dist_exes=dist_exes,
    )
    app_type = normalize_app_type(meta_fields.get("app_type", ""))
    completion_goal = normalize_completion_goal(meta_fields.get("completion_goal", ""))
    status_key = classify_status(folder, meta_fields, checks, issues)
    missing_keys = build_missing_keys(checks, issues, is_formal_release_meta(meta_fields))
    last_modified = latest_mtime(
        [
            readme_path,
            folder / RELEASE_BODY_NAME,
            folder / "main.py",
            folder / "build.bat",
            folder / SCREENSHOT_RELATIVE,
            folder / BOOTH_THUMBNAIL_RELATIVE,
            *booth_product_candidates(folder),
            *dist_exes,
        ],
        folder,
    )
    return AppRecord(
        folder_name=folder.name,
        folder_path=folder,
        meta_fields=meta_fields,
        checks=checks,
        app_type=app_type,
        completion_goal=completion_goal,
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
        app_type=normalize_app_type(meta_fields.get("app_type", "")),
        completion_goal=normalize_completion_goal(meta_fields.get("completion_goal", "")),
        status_key="needs_review",
        missing_keys=build_missing_keys(checks, (issue,), is_formal_release_meta(meta_fields)),
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
    if not is_formal_release_app(record):
        return False
    return (
        record.checks.has_release_url
        and record.checks.has_dist_exe
        and record.checks.has_screenshot
        and record.checks.booth_materials_ready
    )


def is_booth_registration_target(record: AppRecord) -> bool:
    if not is_formal_release_app(record) or not meta_bool(record, "show_on_site"):
        return False
    return (
        record.checks.has_release_url
        and record.checks.has_dist_exe
        and record.checks.has_booth_thumbnail
        and record.checks.has_booth_product
        and record.checks.has_booth_ready
    )


def next_process_key(record: AppRecord) -> str:
    if not is_formal_release_app(record):
        return "internal"
    if record.status_key == "needs_review" or record.issue_messages or "dake_meta" in record.missing_keys:
        return "readme"
    if record.checks.has_dist_exe and not record.checks.has_release_url:
        return "release"
    if record.checks.has_release_url and meta_bool(record, "show_on_site") and not record.checks.booth_materials_ready:
        return "booth_materials"
    if not record.checks.has_screenshot:
        return "screenshot"
    if is_booth_registration_target(record):
        return "booth"
    if not record.checks.has_readme or not record.checks.has_release_body:
        return "readme"
    return "unknown"


def next_process_text(record: AppRecord) -> str:
    return UI_TEXT[f"next_process_{next_process_key(record)}"]


def next_step_label(record: AppRecord) -> str:
    if not is_formal_release_app(record):
        return completion_goal_label(record.completion_goal)
    if record.status_key == "needs_review" or record.issue_messages or "dake_meta" in record.missing_keys:
        return UI_TEXT["next_step_review"]

    process_key = next_process_key(record)
    if process_key == "readme":
        return UI_TEXT["next_step_readme"]
    if process_key == "screenshot":
        return UI_TEXT["next_step_screenshot"]
    if process_key == "release":
        return UI_TEXT["next_step_release"]
    if process_key == "booth_materials":
        return UI_TEXT["next_step_booth_materials"]
    if process_key == "booth":
        return UI_TEXT["next_step_booth"]
    return UI_TEXT["next_step_released"]


def is_safe_url(url: str) -> bool:
    stripped = url.strip()
    if not stripped:
        return False
    lowered = stripped.lower()
    if not (lowered.startswith("http://") or lowered.startswith("https://")):
        return False
    return not any(blocked in lowered for blocked in URL_BLOCKLIST)


def github_repo_url(record: AppRecord) -> str:
    url = record.release_url
    if not is_safe_url(url):
        return ""
    parsed = urlparse(url)
    if parsed.netloc.lower() != "github.com":
        return ""
    parts = [part for part in parsed.path.split("/") if part]
    if len(parts) < 2:
        return ""
    return f"https://github.com/{parts[0]}/{parts[1]}"


def first_dist_exe(record: AppRecord) -> Path | None:
    return record.checks.dist_exes[0] if record.checks.dist_exes else None


def short_display(value: str, max_length: int = 42) -> str:
    if len(value) <= max_length:
        return value
    keep = max(8, (max_length - 3) // 2)
    return f"{value[:keep]}...{value[-keep:]}"


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


def next_candidate_for_record(record: AppRecord) -> tuple[int, str] | None:
    if record.status_key == "needs_review":
        return 1, UI_TEXT["candidate_needs_review"]
    if record.issue_messages or "dake_meta" in record.missing_keys:
        return 2, UI_TEXT["candidate_meta_broken"]
    if record.checks.has_dist_exe and not record.checks.has_release_url and is_formal_release_app(record):
        return 3, UI_TEXT["candidate_release_missing"]
    if (
        is_formal_release_app(record)
        and record.checks.has_release_url
        and meta_bool(record, "show_on_site")
        and not record.checks.booth_materials_ready
    ):
        return 4, UI_TEXT["candidate_booth_missing"]
    if is_booth_registration_target(record):
        return 5, UI_TEXT["candidate_booth_registration"]
    if is_formal_release_app(record) and record.checks.booth_materials_ready and meta_bool(record, "show_on_site"):
        return 6, UI_TEXT["candidate_site_unknown"]
    if not record.checks.has_screenshot and is_formal_release_app(record):
        return 7, UI_TEXT["candidate_screenshot_missing"]
    if not record.checks.has_release_body and is_formal_release_app(record):
        return 8, UI_TEXT["candidate_release_body_missing"]
    if not is_formal_release_app(record):
        return 90, UI_TEXT["candidate_internal"]
    return None


def next_candidates(records: list[AppRecord], limit: int = 5) -> list[tuple[AppRecord, str]]:
    candidates: list[tuple[int, str, AppRecord]] = []
    for record in records:
        candidate = next_candidate_for_record(record)
        if candidate is None:
            continue
        priority, reason = candidate
        candidates.append((priority, reason, record))
    candidates.sort(key=lambda item: (item[0], item[2].folder_name.lower()))
    return [(record, reason) for _priority, reason, record in candidates[:limit]]


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
            "unshipped": tk.StringVar(value="0"),
            "booth_materials_missing": tk.StringVar(value="0"),
            "booth_thumbnail_missing": tk.StringVar(value="0"),
            "booth_product_missing": tk.StringVar(value="0"),
            "booth_ready_missing": tk.StringVar(value="0"),
            "booth_registration_ready": tk.StringVar(value="0"),
            "market": tk.StringVar(value="0"),
            "system": tk.StringVar(value="0"),
            "personal": tk.StringVar(value="0"),
            "frozen": tk.StringVar(value="0"),
        }
        self.qpsc_status_var = tk.StringVar(value=UI_TEXT["qpsc_card_subtitle"])
        self.git_vars = {
            "branch": tk.StringVar(value=UI_TEXT["git_branch"].format(value=UI_TEXT["value_unknown"])),
            "latest": tk.StringVar(value=UI_TEXT["git_latest"].format(value=UI_TEXT["value_unknown"])),
            "uncommitted": tk.StringVar(value=UI_TEXT["git_uncommitted"].format(value=0)),
            "untracked": tk.StringVar(value=UI_TEXT["git_untracked"].format(value=0)),
            "remote": tk.StringVar(value=UI_TEXT["git_remote_clean"]),
            "dashboard": tk.StringVar(value=UI_TEXT["git_dashboard"].format(value=UI_TEXT["value_unknown"])),
        }
        self.search_var = tk.StringVar()
        self.last_loaded_var = tk.StringVar(value=UI_TEXT["last_loaded_waiting"])
        self.status_var = tk.StringVar(value=UI_TEXT["last_loaded_waiting"])
        self.watch_status_var = tk.StringVar(value=UI_TEXT["watch_status_polling"])
        self.count_var = tk.StringVar(value=UI_TEXT["count_line"].format(visible=0, total=0))
        self.filter_buttons: dict[str, tk.Button] = {}
        self.link_rows: dict[str, dict[str, object]] = {}
        self.watch_observer = None
        self.watch_pending_folder: str | None = None
        self.watch_debounce_job: str | None = None
        self.last_watch_event_at = 0.0
        self.pending_reload_folder: str | None = None
        self.booth_working_folder: str | None = None
        self.booth_highlight_job: str | None = None

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
        self.build_git_card(shell)
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
            textvariable=self.qpsc_status_var,
            bg=THEME["panel"],
            fg=THEME["muted"],
            font=(self.font_family, 9),
        ).pack(anchor="w", pady=(3, 0))

        metrics = tk.Frame(card, bg=THEME["panel"])
        metrics.grid(row=0, column=1, sticky="e", padx=(0, 16), pady=10)

        items = [
            ("new_apps", "qpsc_new_apps"),
            ("needs_review", "qpsc_needs_review"),
            ("ship_line", "qpsc_ship_line"),
            ("booth_materials_missing", "qpsc_booth_materials_missing"),
            ("booth_thumbnail_missing", "qpsc_booth_thumbnail_missing"),
            ("booth_product_missing", "qpsc_booth_product_missing"),
            ("booth_ready_missing", "qpsc_booth_ready_missing"),
            ("booth_registration_ready", "qpsc_booth_registration_ready"),
            ("market", "qpsc_market"),
            ("system", "qpsc_system"),
            ("personal", "qpsc_personal"),
            ("frozen", "qpsc_frozen"),
        ]
        for index, (key, label_key) in enumerate(items):
            item = tk.Frame(metrics, bg=THEME["panel_soft"], padx=12, pady=7)
            item.grid(row=index // 6, column=index % 6, sticky="e", padx=(0 if index % 6 == 0 else 8, 0), pady=(0 if index < 6 else 8, 0))
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

    def build_git_card(self, parent: tk.Frame) -> None:
        card = tk.Frame(
            parent,
            bg=THEME["panel_alt"],
            highlightbackground=THEME["border"],
            highlightthickness=1,
            bd=0,
        )
        card.pack(fill="x", pady=(0, 16))
        card.grid_columnconfigure(1, weight=1)

        title_area = tk.Frame(card, bg=THEME["panel_alt"])
        title_area.grid(row=0, column=0, sticky="w", padx=16, pady=12)
        tk.Label(
            title_area,
            text=UI_TEXT["git_card_title"],
            bg=THEME["panel_alt"],
            fg=THEME["text"],
            font=(self.font_family, 12, "bold"),
        ).pack(anchor="w")
        tk.Label(
            title_area,
            text=UI_TEXT["git_card_subtitle"],
            bg=THEME["panel_alt"],
            fg=THEME["muted"],
            font=(self.font_family, 9),
        ).pack(anchor="w", pady=(3, 0))

        values = tk.Frame(card, bg=THEME["panel_alt"])
        values.grid(row=0, column=1, sticky="e", padx=(0, 16), pady=12)
        for index, key in enumerate(("branch", "latest", "uncommitted", "untracked", "remote", "dashboard")):
            tk.Label(
                values,
                textvariable=self.git_vars[key],
                bg=THEME["panel_alt"],
                fg=THEME["muted"] if key not in {"remote", "dashboard"} else THEME["accent_hover"],
                font=(self.font_family, 9, "bold"),
            ).grid(row=0, column=index, sticky="e", padx=(0 if index == 0 else 14, 0))

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

        columns = (
            "status",
            "folder",
            "display",
            "app_type",
            "completion_goal",
            "release",
            "booth",
            "screenshot",
            "exe",
            "next_step",
            "updated",
        )
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
            "app_type": UI_TEXT["column_app_type"],
            "completion_goal": UI_TEXT["column_completion_goal"],
            "release": UI_TEXT["column_release"],
            "booth": UI_TEXT["column_booth"],
            "screenshot": UI_TEXT["column_screenshot"],
            "exe": UI_TEXT["column_exe"],
            "next_step": UI_TEXT["column_next_step"],
            "updated": UI_TEXT["column_updated"],
        }
        widths = {
            "status": 96,
            "folder": 150,
            "display": 150,
            "app_type": 96,
            "completion_goal": 92,
            "release": 70,
            "booth": 74,
            "screenshot": 82,
            "exe": 58,
            "next_step": 116,
            "updated": 122,
        }
        for column in columns:
            self.tree.heading(column, text=headings[column])
            anchor = "center" if column in {"status", "app_type", "completion_goal", "release", "booth", "screenshot", "exe", "next_step", "updated"} else "w"
            self.tree.column(column, width=widths[column], minwidth=54, stretch=column in {"folder", "display"}, anchor=anchor)

        for key, (_bg, fg) in STATUS_THEME.items():
            self.tree.tag_configure(key, foreground=fg)
        self.tree.tag_configure("booth_working", background=THEME["accent_soft"], foreground=THEME["text"])
        self.tree.tag_configure("booth_count_0", foreground=THEME["muted"])
        self.tree.tag_configure("booth_count_1", foreground=THEME["warning"])
        self.tree.tag_configure("booth_count_2", foreground=THEME["warning"])
        self.tree.tag_configure("booth_count_3", foreground=THEME["success"])
        self.tree.tag_configure("booth_internal", foreground=THEME["purple"])

        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)

    def build_detail(self, parent: tk.Frame) -> None:
        container = tk.Frame(parent, bg=THEME["panel"])
        container.pack(fill="both", expand=True, padx=16, pady=14)
        container.grid_columnconfigure(0, weight=1)
        container.grid_rowconfigure(6, weight=1)

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

        work_row = tk.Frame(container, bg=THEME["panel"])
        work_row.grid(row=4, column=0, sticky="ew", pady=(0, 10))
        self.next_work_button = self.make_button(work_row, UI_TEXT["button_next_work"], self.open_next_work, primary=True)
        self.next_work_button.pack(side="left", padx=(0, 8))
        self.open_booth_assist_button = self.make_button(
            work_row,
            UI_TEXT["button_open_booth_assist"],
            self.open_selected_booth_assist,
        )
        self.open_booth_assist_button.pack(side="left")

        self.build_workspace_links(container, 5)

        detail_area = tk.Frame(container, bg=THEME["panel"])
        detail_area.grid(row=6, column=0, sticky="nsew")
        detail_area.grid_columnconfigure(0, weight=1)
        detail_area.grid_rowconfigure(1, weight=1)
        detail_area.grid_rowconfigure(3, weight=1)
        detail_area.grid_rowconfigure(5, weight=1)
        detail_area.grid_rowconfigure(7, weight=1)
        detail_area.grid_rowconfigure(9, weight=1)
        detail_area.grid_rowconfigure(11, weight=1)

        self.meta_text = self.create_detail_text(detail_area, 0, UI_TEXT["detail_meta_title"], height=3)
        self.shipment_text = self.create_detail_text(detail_area, 2, UI_TEXT["shipment_title"], height=3)
        self.files_text = self.create_detail_text(detail_area, 4, UI_TEXT["detail_files_title"], height=5)
        self.missing_text = self.create_detail_text(detail_area, 6, UI_TEXT["detail_missing_title"], height=3)
        self.next_text = self.create_detail_text(detail_area, 8, UI_TEXT["detail_next_title"], height=3)
        self.next_candidates_text = self.create_detail_text(detail_area, 10, UI_TEXT["next_candidates_title"], height=4)
        self.update_detail(None)

    def build_workspace_links(self, parent: tk.Frame, row: int) -> None:
        links_area = tk.Frame(parent, bg=THEME["panel"])
        links_area.grid(row=row, column=0, sticky="ew", pady=(0, 12))
        links_area.grid_columnconfigure(0, weight=1)
        links_area.grid_columnconfigure(1, weight=1)

        tk.Label(
            links_area,
            text=UI_TEXT["workspace_links_title"],
            bg=THEME["panel"],
            fg=THEME["muted"],
            font=(self.font_family, 9, "bold"),
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))

        for index, key in enumerate(WORKSPACE_LINK_KEYS):
            cell = tk.Frame(
                links_area,
                bg=THEME["panel_soft"],
                highlightbackground=THEME["border"],
                highlightthickness=1,
                bd=0,
            )
            cell.grid(row=1 + index // 2, column=index % 2, sticky="ew", padx=(0 if index % 2 == 0 else 8, 0), pady=(0, 6))
            cell.grid_columnconfigure(0, weight=1)

            label_var = tk.StringVar(value=UI_TEXT[f"link_{key}"])
            value_var = tk.StringVar(value=UI_TEXT["link_missing"])
            tk.Label(
                cell,
                textvariable=label_var,
                bg=THEME["panel_soft"],
                fg=THEME["text"],
                font=(self.font_family, 8, "bold"),
            ).grid(row=0, column=0, sticky="w", padx=(8, 6), pady=(5, 0))
            tk.Label(
                cell,
                textvariable=value_var,
                bg=THEME["panel_soft"],
                fg=THEME["muted"],
                font=(self.font_family, 8),
                anchor="w",
            ).grid(row=1, column=0, sticky="ew", padx=(8, 6), pady=(0, 5))
            button = self.make_button(
                cell,
                UI_TEXT["button_show_location"] if key == "exe" else UI_TEXT["button_open"],
                lambda value=key: self.open_workspace_link(value),
                compact=True,
            )
            button.grid(row=0, column=1, rowspan=2, sticky="e", padx=(0, 6), pady=5)
            self.link_rows[key] = {"value": value_var, "button": button}

        self.open_booth_ready_button = self.link_rows["booth_ready"]["button"]
        self.open_booth_product_button = self.link_rows["booth_product"]["button"]
        self.open_booth_thumbnail_button = self.link_rows["booth_thumbnail"]["button"]
        self.open_screenshot_jpg_button = self.link_rows["screenshot_jpg"]["button"]
        self.open_screenshot_webp_button = self.link_rows["screenshot_webp"]["button"]

    def create_detail_text(self, parent: tk.Frame, row: int, title: str, height: int = 5) -> tk.Text:
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
            height=height,
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

    def make_button(self, parent: tk.Misc, label: str, command, primary: bool = False, compact: bool = False) -> tk.Button:
        bg = THEME["accent_soft"] if primary else THEME["panel_alt"]
        fg = THEME["text"] if primary else THEME["muted"]
        active_bg = THEME["accent"] if primary else THEME["selection"]
        active_fg = "#FFFFFF" if primary else THEME["text"]
        button = tk.Button(
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
            padx=8 if compact else 12,
            pady=4 if compact else 7,
            cursor="hand2",
            font=(self.font_family, 8 if compact else 9, "bold"),
        )
        if compact:
            button.bind("<Enter>", lambda _event: button.configure(bg=active_bg if button["state"] != "disabled" else bg))
            button.bind("<Leave>", lambda _event: button.configure(bg=bg))
        return button

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
        git_status = read_git_status(series_root())
        self.worker_queue.put(
            (
                "scan_done",
                {
                    "records": records,
                    "git_status": git_status,
                    "source": source,
                    "changed_folder": changed_folder,
                },
            )
        )

    def poll_worker(self) -> None:
        try:
            while True:
                event, payload = self.worker_queue.get_nowait()
                if event == "scan_done":
                    self.handle_scan_done(payload)
                elif event == "watch_event":
                    self.handle_watch_event(payload)
                elif event == "booth_assist_status":
                    self.show_action_notification(str(payload))
                elif event == "link_notice":
                    self.show_action_notification(str(payload))
        except queue.Empty:
            pass
        self.root.after(WORKER_POLL_MS, self.poll_worker)

    def handle_scan_done(self, payload: object) -> None:
        if isinstance(payload, dict):
            records = payload.get("records", [])
            git_status = payload.get("git_status")
            source = str(payload.get("source", "manual"))
            changed_folder = payload.get("changed_folder")
        else:
            records = payload
            git_status = None
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
        if isinstance(git_status, GitStatus):
            self.update_git_card(git_status)
        self.apply_filters()
        if self.selected_record is not None:
            refreshed_record = current_map.get(self.selected_record.folder_name)
            if refreshed_record is not None:
                self.update_detail(refreshed_record)
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
        formal_records = [record for record in self.records if is_formal_release_app(record)]
        self.qpsc_vars["new_apps"].set(str(new_app_count))
        self.qpsc_vars["needs_review"].set(str(counts["needs_review"]))
        self.qpsc_vars["ship_line"].set(str(sum(1 for record in self.records if formal_ship_line_reached(record))))
        self.qpsc_vars["booth_materials_missing"].set(str(sum(1 for record in formal_records if not record.checks.booth_materials_ready)))
        self.qpsc_vars["booth_thumbnail_missing"].set(str(sum(1 for record in formal_records if not record.checks.has_booth_thumbnail)))
        self.qpsc_vars["booth_product_missing"].set(str(sum(1 for record in formal_records if not record.checks.has_booth_product)))
        self.qpsc_vars["booth_ready_missing"].set(str(sum(1 for record in formal_records if not record.checks.has_booth_ready)))
        self.qpsc_vars["booth_registration_ready"].set(str(sum(1 for record in formal_records if is_booth_registration_target(record))))
        self.qpsc_vars["market"].set(str(sum(1 for record in self.records if record.app_type == "market")))
        self.qpsc_vars["system"].set(str(sum(1 for record in self.records if record.app_type == "system")))
        self.qpsc_vars["personal"].set(str(sum(1 for record in self.records if record.app_type == "personal")))
        self.qpsc_vars["frozen"].set(str(sum(1 for record in self.records if record.app_type == "frozen")))

    def update_git_card(self, status: GitStatus) -> None:
        if status.error:
            self.git_vars["branch"].set(UI_TEXT["git_branch"].format(value=UI_TEXT["value_unknown"]))
            self.git_vars["latest"].set(UI_TEXT["git_latest"].format(value=UI_TEXT["value_unknown"]))
            self.git_vars["uncommitted"].set(UI_TEXT["git_uncommitted"].format(value=0))
            self.git_vars["untracked"].set(UI_TEXT["git_untracked"].format(value=0))
            self.git_vars["remote"].set(UI_TEXT["git_error"])
            self.git_vars["dashboard"].set(UI_TEXT["git_dashboard"].format(value=UI_TEXT["value_unknown"]))
            return

        self.git_vars["branch"].set(UI_TEXT["git_branch"].format(value=status.branch))
        self.git_vars["latest"].set(UI_TEXT["git_latest"].format(value=status.latest))
        self.git_vars["uncommitted"].set(UI_TEXT["git_uncommitted"].format(value=status.uncommitted_count))
        self.git_vars["untracked"].set(UI_TEXT["git_untracked"].format(value=status.untracked_count))
        if status.ahead_count:
            remote_text = UI_TEXT["git_push_waiting"].format(value=status.ahead_count)
        elif status.behind_count:
            remote_text = UI_TEXT["git_pull_waiting"].format(value=status.behind_count)
        else:
            remote_text = UI_TEXT["git_remote_clean"]
        self.git_vars["remote"].set(remote_text)
        dashboard_value = UI_TEXT["git_dashboard_dirty"] if status.dashboard_dirty else UI_TEXT["git_dashboard_clean"]
        self.git_vars["dashboard"].set(UI_TEXT["git_dashboard"].format(value=dashboard_value))

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
        self.show_action_notification(UI_TEXT["notification_reloaded"].format(folder=folder_name))

    def show_action_notification(self, message: str) -> None:
        self.status_var.set(message)
        self.notification_var.set(message)
        self.notification_frame.place(relx=1.0, rely=1.0, anchor="se", x=-22, y=-22)
        self.root.after(NOTIFICATION_HIDE_MS, self.notification_frame.place_forget)

    def on_close(self) -> None:
        if self.booth_highlight_job is not None:
            try:
                self.root.after_cancel(self.booth_highlight_job)
            except tk.TclError:
                pass
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
        if self.filter_key in {"market", "system", "personal"}:
            return record.app_type == self.filter_key
        if self.filter_key == "frozen":
            return record.app_type == "frozen" or record.status_key == "frozen_closed"
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
                record.app_type,
                record.completion_goal,
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
            tags = [record.status_key]
            if not is_formal_release_app(record):
                tags.append("booth_internal")
            else:
                tags.append(f"booth_count_{record.checks.booth_materials_count}")
            if record.folder_name == self.booth_working_folder:
                tags.append("booth_working")
            self.tree.insert(
                "",
                "end",
                iid=iid,
                values=(
                    record.status_text,
                    record.folder_name,
                    record.display_name,
                    app_type_label(record.app_type),
                    completion_goal_label(record.completion_goal),
                    bool_label(record.checks.has_release_url),
                    booth_label(record.checks),
                    bool_label(record.checks.has_screenshot),
                    bool_label(record.checks.has_dist_exe),
                    next_step_label(record),
                    format_datetime(record.last_modified),
                ),
                tags=tuple(tags),
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
            self.set_text(self.shipment_text, UI_TEXT["detail_empty"])
            self.set_text(self.files_text, UI_TEXT["detail_empty"])
            self.set_text(self.missing_text, UI_TEXT["detail_empty"])
            self.set_text(self.next_text, UI_TEXT["detail_empty"])
            self.set_text(self.next_candidates_text, self.build_next_candidates_detail())
            self.set_detail_buttons_state(False)
            return

        badge_bg, badge_fg = STATUS_THEME.get(record.status_key, (THEME["panel_soft"], THEME["muted"]))
        self.detail_status_var.set(record.status_text)
        self.detail_badge.configure(bg=badge_bg, fg=badge_fg)
        self.detail_name_var.set(f"{record.folder_name}\n{record.display_name}")
        self.set_text(self.meta_text, self.build_meta_detail(record))
        self.set_text(self.shipment_text, self.build_shipment_detail(record))
        self.set_text(self.files_text, self.build_file_detail(record))
        self.set_text(self.missing_text, self.build_missing_detail(record))
        self.set_text(self.next_text, self.build_next_detail(record))
        self.set_text(self.next_candidates_text, self.build_next_candidates_detail())
        self.set_detail_buttons_state(True)

    def set_detail_buttons_state(self, enabled: bool) -> None:
        record = self.selected_record if enabled else None
        state = "normal" if record is not None else "disabled"
        self.open_folder_button.configure(state=state)
        self.open_readme_button.configure(state=state)
        self.next_work_button.configure(state=state)
        booth_assist_state = "normal" if record is not None and is_booth_registration_target(record) else "disabled"
        self.open_booth_assist_button.configure(state=booth_assist_state)

        release_state = "normal" if record is not None and record.release_url else "disabled"
        self.open_release_button.configure(state=release_state)
        if release_state == "disabled":
            self.open_release_button.configure(text=UI_TEXT["button_release_missing"])
        else:
            self.open_release_button.configure(text=UI_TEXT["button_open_release"])
        self.update_workspace_links(record)

    def set_path_button_state(self, button: tk.Button, path: Path | None) -> None:
        button.configure(state="normal" if path is not None and path.exists() else "disabled")

    def update_workspace_links(self, record: AppRecord | None) -> None:
        for key in WORKSPACE_LINK_KEYS:
            row = self.link_rows.get(key)
            if not row:
                continue
            value_var = row["value"]
            button = row["button"]
            if not isinstance(value_var, tk.StringVar) or not isinstance(button, tk.Button):
                continue
            if record is None:
                value_var.set(UI_TEXT["link_missing"])
                button.configure(state="disabled")
                continue
            display_value, enabled = self.workspace_link_state(record, key)
            value_var.set(display_value)
            button.configure(state="normal" if enabled else "disabled")

    def workspace_link_state(self, record: AppRecord, key: str) -> tuple[str, bool]:
        if key == "release_url":
            url = record.release_url
            return (short_display(url) if url else UI_TEXT["link_missing"], bool(url))
        if key == "github_url":
            url = github_repo_url(record)
            return (short_display(url) if url else UI_TEXT["link_unavailable"], bool(url))
        path = self.workspace_link_path(record, key)
        if path is None:
            return UI_TEXT["link_missing"], False
        display_value = self.display_path(record, path)
        return display_value if path.exists() else UI_TEXT["link_missing"], path.exists()

    def workspace_link_path(self, record: AppRecord, key: str) -> Path | None:
        paths = {
            "folder": record.folder_path,
            "readme": record.folder_path / README_NAME,
            "release_body": record.folder_path / RELEASE_BODY_NAME,
            "assets": record.folder_path / "assets",
            "screenshot_webp": record.folder_path / SCREENSHOT_RELATIVE,
            "screenshot_jpg": record.folder_path / SCREENSHOT_JPG_RELATIVE,
            "booth_thumbnail": record.folder_path / BOOTH_THUMBNAIL_RELATIVE,
            "booth_product": find_booth_product_path(record.folder_path) or record.folder_path / BOOTH_PRODUCT_NAME,
            "booth_ready": record.folder_path / BOOTH_READY_NAME,
            "dist": record.folder_path / DIST_DIR_NAME,
            "exe": first_dist_exe(record),
        }
        return paths.get(key)

    def display_path(self, record: AppRecord, path: Path) -> str:
        try:
            value = str(path.relative_to(record.folder_path))
        except ValueError:
            value = str(path)
        if value == ".":
            value = str(path)
        return short_display(value)

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
            if key == "app_type":
                value = f"{value} ({app_type_label(record.app_type)})"
            elif key == "completion_goal":
                value = f"{value} ({completion_goal_label(record.completion_goal)})"
            lines.append(f"{label}: {value}")
        if record.issue_messages:
            lines.append("")
            lines.extend(record.issue_messages)
        return "\n".join(lines)

    def build_file_detail(self, record: AppRecord) -> str:
        checks = record.checks
        booth_missing = booth_missing_keys(checks)
        booth_lines = [UI_TEXT["booth_status_title"].format(ready=checks.booth_materials_count)]
        for key in BOOTH_MATERIAL_KEYS:
            template = UI_TEXT["booth_status_missing"] if key in booth_missing else UI_TEXT["booth_status_ready"]
            booth_lines.append(template.format(label=booth_item_label(key)))
        if booth_missing:
            booth_lines.append("")
            booth_lines.append(UI_TEXT["booth_missing_title"])
            booth_lines.extend(f"- {booth_item_label(key)}" for key in booth_missing)

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
        lines = [*booth_lines, "", *[f"{UI_TEXT[label_key]}: {bool_label(value)}" for label_key, value in rows]]
        if checks.dist_exes:
            lines.append("")
            lines.extend(path.name for path in checks.dist_exes)
        return "\n".join(lines)

    def build_shipment_detail(self, record: AppRecord) -> str:
        rate = shipment_rate(record)
        if rate is None:
            return "\n".join(
                [
                    UI_TEXT["shipment_non_formal"].format(
                        app_type=app_type_label(record.app_type),
                        goal=completion_goal_label(record.completion_goal),
                    ),
                    UI_TEXT["shipment_non_formal_note"],
                ]
            )

        lines = [UI_TEXT["shipment_rate"].format(percent=rate)]
        missing = shipment_missing_keys(record)
        if missing:
            lines.append("")
            lines.append(UI_TEXT["shipment_missing_title"])
            lines.extend(f"- {UI_TEXT[f'missing_{key}']}" for key in missing)
        else:
            lines.append(UI_TEXT["missing_none"])
        return "\n".join(lines)

    def build_missing_detail(self, record: AppRecord) -> str:
        if not record.missing_keys:
            if not is_formal_release_app(record):
                return UI_TEXT["missing_non_formal"]
            return UI_TEXT["missing_none"]
        return "\n".join(f"- {UI_TEXT[f'missing_{key}']}" for key in record.missing_keys)

    def build_next_candidates_detail(self) -> str:
        candidates = next_candidates(self.records)
        if not candidates:
            return UI_TEXT["next_candidates_empty"]
        blocks = [
            UI_TEXT["next_candidate_line"].format(folder=record.folder_name, reason=reason)
            for record, reason in candidates
        ]
        return "\n\n".join(blocks)

    def build_next_detail(self, record: AppRecord) -> str:
        tasks: list[str] = [next_process_text(record)]
        if not is_formal_release_app(record) and record.status_key != "needs_review":
            tasks.append(UI_TEXT["next_non_formal_goal"].format(goal=completion_goal_label(record.completion_goal)))
            return "\n".join(f"- {task}" for task in tasks)
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
            self.open_path(self.selected_record.folder_path, UI_TEXT["link_folder"])

    def open_selected_readme(self) -> None:
        if self.selected_record is not None:
            self.open_path(self.selected_record.folder_path / README_NAME, UI_TEXT["link_readme"])

    def open_selected_release(self) -> None:
        if self.selected_record is None:
            return
        self.open_url(self.selected_record.release_url, UI_TEXT["link_release_url"])

    def open_selected_booth_ready(self) -> None:
        self.open_selected_relative_path(BOOTH_READY_NAME)

    def open_selected_booth_product(self) -> None:
        if self.selected_record is not None:
            path = find_booth_product_path(self.selected_record.folder_path) or self.selected_record.folder_path / BOOTH_PRODUCT_NAME
            self.open_path(path, UI_TEXT["link_booth_product"])

    def open_selected_booth_thumbnail(self) -> None:
        self.open_selected_relative_path(BOOTH_THUMBNAIL_RELATIVE)

    def open_selected_screenshot_jpg(self) -> None:
        self.open_selected_relative_path(SCREENSHOT_JPG_RELATIVE)

    def open_selected_screenshot_webp(self) -> None:
        self.open_selected_relative_path(SCREENSHOT_RELATIVE)

    def open_selected_relative_path(self, relative_path: str | Path) -> None:
        if self.selected_record is not None:
            path = self.selected_record.folder_path / relative_path
            self.open_path(path, self.target_label_for_path(path))

    def open_workspace_link(self, key: str) -> None:
        record = self.selected_record
        if record is None:
            return
        if key == "release_url":
            self.open_url(record.release_url, UI_TEXT["link_release_url"])
            return
        if key == "github_url":
            self.open_url(github_repo_url(record), UI_TEXT["link_github_url"])
            return
        if key == "exe":
            exe_path = first_dist_exe(record)
            if exe_path is None:
                self.show_action_notification(UI_TEXT["notice_missing_target"])
                return
            self.show_file_location(exe_path)
            return
        path = self.workspace_link_path(record, key)
        if path is None:
            self.show_action_notification(UI_TEXT["notice_missing_target"])
            return
        self.open_path(path, UI_TEXT[f"link_{key}"])

    def target_label_for_path(self, path: Path) -> str:
        name = path.name.lower()
        if name == BOOTH_READY_NAME:
            return UI_TEXT["link_booth_ready"]
        if name == BOOTH_PRODUCT_NAME:
            return UI_TEXT["link_booth_product"]
        if name == "booth_thumbnail.jpg":
            return UI_TEXT["link_booth_thumbnail"]
        if name == "screenshot.jpg":
            return UI_TEXT["link_screenshot_jpg"]
        if name == "screenshot.webp":
            return UI_TEXT["link_screenshot_webp"]
        if name == README_NAME.lower():
            return UI_TEXT["link_readme"]
        if name == RELEASE_BODY_NAME.lower():
            return UI_TEXT["link_release_body"]
        return path.name

    def open_selected_booth_assist(self) -> None:
        if self.selected_record is not None and is_booth_registration_target(self.selected_record):
            self.launch_booth_assist(self.selected_record)

    def open_next_work(self) -> None:
        record = self.selected_record
        if record is None:
            return
        process_key = next_process_key(record)
        if process_key == "booth":
            self.launch_booth_assist(record)
            self.open_path(record.folder_path / BOOTH_READY_NAME, UI_TEXT["link_booth_ready"])
            self.open_path(find_booth_product_path(record.folder_path) or record.folder_path / BOOTH_PRODUCT_NAME, UI_TEXT["link_booth_product"])
        elif process_key == "release":
            self.open_path(record.folder_path / README_NAME, UI_TEXT["link_readme"])
            release_body = record.folder_path / RELEASE_BODY_NAME
            if release_body.exists():
                self.open_path(release_body, UI_TEXT["link_release_body"])
            dist_dir = record.folder_path / DIST_DIR_NAME
            if dist_dir.exists():
                self.open_path(dist_dir, UI_TEXT["link_dist"])
            exe_path = first_dist_exe(record)
            if exe_path is not None:
                self.show_file_location(exe_path)
        elif process_key == "screenshot":
            assets_dir = record.folder_path / "assets"
            self.open_path(assets_dir if assets_dir.exists() else record.folder_path, UI_TEXT["link_assets"])
            self.open_path(record.folder_path, UI_TEXT["link_folder"])
            screenshot_path = record.folder_path / SCREENSHOT_RELATIVE
            if screenshot_path.exists():
                self.open_path(screenshot_path, UI_TEXT["link_screenshot_webp"])
        elif process_key == "readme":
            self.open_path(record.folder_path / README_NAME, UI_TEXT["link_readme"])
            self.open_path(record.folder_path, UI_TEXT["link_folder"])
        elif process_key == "booth_materials":
            assets_dir = record.folder_path / "assets"
            self.open_path(assets_dir if assets_dir.exists() else record.folder_path, UI_TEXT["link_assets"])
            booth_product = find_booth_product_path(record.folder_path)
            if booth_product is not None:
                self.open_path(booth_product, UI_TEXT["link_booth_product"])
            booth_ready = record.folder_path / BOOTH_READY_NAME
            if booth_ready.exists():
                self.open_path(booth_ready, UI_TEXT["link_booth_ready"])
        else:
            self.open_path(record.folder_path, UI_TEXT["link_folder"])

    def launch_booth_assist(self, record: AppRecord) -> None:
        self.highlight_booth_work(record)
        thread = threading.Thread(target=self.booth_assist_worker, args=(record.folder_name,), daemon=True)
        thread.start()

    def booth_assist_worker(self, folder_name: str) -> None:
        assist_dir = apps_root() / BOOTH_ASSIST_FOLDER_NAME
        exe_path = assist_dir / DIST_DIR_NAME / BOOTH_ASSIST_EXE_NAME
        main_path = assist_dir / "main.py"
        started = False

        if exe_path.exists():
            started = self.start_booth_assist_process([str(exe_path), "--app", folder_name], assist_dir)
            if not started:
                started = self.start_booth_assist_process([str(exe_path)], assist_dir)

        if not started and main_path.exists():
            python_cmd = sys.executable if not getattr(sys, "frozen", False) else "python"
            started = self.start_booth_assist_process([python_cmd, str(main_path), "--app", folder_name], assist_dir)
            if not started:
                started = self.start_booth_assist_process([python_cmd, str(main_path)], assist_dir)

        if started:
            message = UI_TEXT["booth_assist_launched"].format(folder=folder_name)
        else:
            message = UI_TEXT["booth_assist_missing"]
        self.worker_queue.put(("booth_assist_status", message))

    @staticmethod
    def start_booth_assist_process(command: list[str], cwd: Path) -> bool:
        try:
            process = subprocess.Popen(command, cwd=str(cwd), **hidden_subprocess_kwargs())
        except Exception:
            return False
        try:
            return process.wait(timeout=1.0) == 0
        except subprocess.TimeoutExpired:
            return True

    def highlight_booth_work(self, record: AppRecord) -> None:
        self.booth_working_folder = record.folder_name
        message = UI_TEXT["qpsc_booth_working"].format(folder=record.folder_name)
        self.qpsc_status_var.set(message)
        self.status_var.set(message)
        self.render_tree()
        iid = str(record.folder_path)
        if iid in self.record_by_iid:
            self.tree.selection_set(iid)
            self.tree.focus(iid)
            self.tree.see(iid)
        if self.booth_highlight_job is not None:
            try:
                self.root.after_cancel(self.booth_highlight_job)
            except tk.TclError:
                pass
        self.booth_highlight_job = self.root.after(BOOTH_HIGHLIGHT_MS, lambda: self.clear_booth_work_highlight(record.folder_name))

    def clear_booth_work_highlight(self, folder_name: str) -> None:
        self.booth_highlight_job = None
        if self.booth_working_folder != folder_name:
            return
        self.booth_working_folder = None
        self.qpsc_status_var.set(UI_TEXT["qpsc_card_subtitle"])
        self.render_tree()

    def open_path(self, path: Path, target_label: str) -> None:
        if not path.exists():
            self.show_action_notification(UI_TEXT["notice_missing_target"])
            return
        thread = threading.Thread(target=self.open_path_worker, args=(path, target_label), daemon=True)
        thread.start()

    def open_path_worker(self, path: Path, target_label: str) -> None:
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(path))
            else:
                webbrowser.open(path.as_uri())
            message = UI_TEXT["notice_opened"].format(target=target_label)
        except OSError as exc:
            message = UI_TEXT["notice_missing_target"] if not str(exc) else UI_TEXT["notice_missing_target"]
        self.worker_queue.put(("link_notice", message))

    def open_url(self, url: str, _target_label: str) -> None:
        if not is_safe_url(url):
            self.show_action_notification(UI_TEXT["notice_url_failed"])
            return
        thread = threading.Thread(target=self.open_url_worker, args=(url,), daemon=True)
        thread.start()

    def open_url_worker(self, url: str) -> None:
        try:
            opened = webbrowser.open(url)
        except Exception:
            opened = False
        message = UI_TEXT["notice_url_opened"] if opened else UI_TEXT["notice_url_failed"]
        self.worker_queue.put(("link_notice", message))

    def show_file_location(self, path: Path) -> None:
        if not path.exists():
            self.show_action_notification(UI_TEXT["notice_missing_target"])
            return
        thread = threading.Thread(target=self.show_file_location_worker, args=(path,), daemon=True)
        thread.start()

    def show_file_location_worker(self, path: Path) -> None:
        try:
            if sys.platform.startswith("win"):
                subprocess.Popen(["explorer.exe", f"/select,{str(path)}"], **hidden_subprocess_kwargs())
            else:
                webbrowser.open(path.parent.as_uri())
            message = UI_TEXT["notice_exe_location"]
        except Exception:
            message = UI_TEXT["notice_missing_target"]
        self.worker_queue.put(("link_notice", message))


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
