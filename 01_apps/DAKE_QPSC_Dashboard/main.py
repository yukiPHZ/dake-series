# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import importlib.util
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

try:
    import ctypes
except Exception:
    ctypes = None


APP_ID = "dake.qpsc.dashboard"
APP_FOLDER_NAME = "DAKE_QPSC_Dashboard"
APP_DASHBOARD_FOLDER = "DAKE_App_Dashboard"
WEB_DASHBOARD_FOLDER = "DAKE_Web_Dashboard"
APP_DASHBOARD_EXE = "DakeApp_Dashboard.exe"
WEB_DASHBOARD_EXE = "DakeWeb_Dashboard.exe"
BOOTH_ASSIST_FOLDER = "DAKE_BOOTH_Assist"
BOOTH_ASSIST_EXE = "DakeBOOTH_Assist.exe"
DIST_DIR_NAME = "dist"
README_NAME = "README.md"
RELEASE_BODY_NAME = "release_body.md"
DEFAULT_SERIES_ROOT = Path(os.environ.get("DAKE_SERIES_ROOT", r"C:\Users\yukiz\devlop\DAKE_series"))
DEFAULT_APPS_ROOT = Path(os.environ.get("QPSC_SERIES_APPS_ROOT", str(DEFAULT_SERIES_ROOT / "01_apps")))
WORKER_POLL_MS = 80
MAX_NEXT_ACTIONS = 5

UI_TEXT = {
    "window_title": "Quiet Personal Cognitive System",
    "app_title": "Quiet Personal Cognitive System",
    "header_title": "Quiet Personal Cognitive System",
    "header_kicker": "QPSC",
    "header_subtitle": "正本を読み、次にやることだけを静かに表示します。",
    "section_current_title": "現在地",
    "button_reload": "再確認",
    "button_open_app_dashboard": "App詳細",
    "button_open_web_dashboard": "Web詳細",
    "button_open_target": "選択対象を開く",
    "button_close": "閉じる",
    "button_open_failed_title": "起動できません",
    "button_open_failed": "対象が見つかりません。\n\n{path}",
    "button_open_error": "起動に失敗しました。\n\n{path}\n\n{error}",
    "card_app_title": "アプリ",
    "card_site_title": "サイト",
    "card_git_title": "Git",
    "card_next_title": "いま見るもの",
    "label_app_total": "アプリ総数",
    "label_booth_missing": "BOOTH未登録",
    "label_release_missing": "Release未作成",
    "label_screenshot_missing": "スクショ未作成",
    "label_readme_missing": "README不足",
    "label_role_attention": "分類別 要確認",
    "label_market_count": "市場向け",
    "label_system_count": "QPSC / 補助脳系",
    "label_personal_count": "ユキズ専用",
    "label_frozen_count": "凍結",
    "label_site_total": "サイト総数",
    "label_cloudflare_unchecked": "Cloudflare未確認",
    "label_health_attention": "health未確認または異常",
    "label_site_git_uncommitted": "Git未反映または未commit",
    "label_series_uncommitted": "DAKE_series 未commit件数",
    "label_series_untracked": "DAKE_series 未追跡件数",
    "label_git_attention": "要確認",
    "value_waiting": "確認待ち",
    "value_none": "なし",
    "value_yes": "あり",
    "value_git_ok": "なし",
    "value_git_attention": "要確認あり",
    "priority_urgent": "優先",
    "priority_active": "通常",
    "priority_later": "保留",
    "priority_summary": "優先 {urgent} / 全体 {total}\n通常 {active} / 保留 {later}",
    "priority_header": "【{label}】",
    "status_checking": "確認中…",
    "status_ready": "確認完了",
    "status_attention": "要確認あり",
    "status_launch_check_ok": "LAUNCH CHECK OK",
    "last_loaded_waiting": "未確認",
    "last_loaded_value": "確認: {time}",
    "summary_template": "優先 {urgent}件 / 未処理全体 {total}件 / アプリ {apps}件 / サイト {sites}件",
    "dialog_empty": "対象はありません。",
    "dialog_booth_title": "BOOTH未登録アプリ",
    "dialog_release_title": "Release未作成アプリ",
    "dialog_screenshot_title": "スクショ未作成アプリ",
    "dialog_readme_title": "README不足アプリ",
    "dialog_role_title": "分類別 要確認アプリ",
    "dialog_cloudflare_title": "Cloudflare未確認サイト",
    "dialog_health_title": "health未確認または異常サイト",
    "dialog_site_git_title": "Git未反映または未commitサイト",
    "dialog_series_git_title": "DAKE_series Git要確認",
    "dialog_select_notice": "一覧から対象を選択してください。",
    "next_none": "現時点で明確な候補はありません。",
    "next_booth": "BOOTH登録が必要な正式出荷候補",
    "next_screenshot": "スクショ作成が必要な公開アプリ",
    "next_release": "Release作成が必要な出荷候補",
    "next_readme": "README不足のアプリ",
    "next_role": "分類別の確認が必要なアプリ",
    "next_cloudflare": "Cloudflare確認が必要な公開サイト",
    "next_health": "health確認が必要な公開サイト",
    "next_site_git": "サイト系Git未反映を確認",
    "next_series_git": "DAKE_seriesの未commitを確認",
    "next_line": "{index}. {label}（{count}件）",
    "reason_booth_missing": "BOOTH素材またはbooth_urlが未整備",
    "reason_release_missing": "release_urlが未設定",
    "reason_screenshot_missing": "screenshot_pathまたは実ファイルが未整備",
    "reason_readme_missing": "README.mdがありません",
    "reason_meta_missing": "DAKE_METAが未整備です",
    "reason_meta_json": "DAKE_META JSONを解析できません",
    "reason_release_body_missing": "RELEASE_BODYまたはrelease_body.mdがありません",
    "reason_meta_fields_missing": "DAKE_META必須項目が不足",
    "reason_update_summary_missing": "update_summaryが不足",
    "reason_role_unknown": "app_typeまたはcompletion_goalが未分類です",
    "reason_role_goal_mismatch": "app_typeとcompletion_goalの組み合わせを確認してください",
    "reason_system_ready_attention": "システム稼働の完成条件を確認してください",
    "reason_reference_ready_attention": "正本提示の完成条件を確認してください",
    "reason_personal_ready_attention": "ローカル運用の完成条件を確認してください",
    "reason_frozen_ready_attention": "凍結理由と再開条件を確認してください",
    "reason_role_missing_items": "不足: {items}",
    "app_type_market": "市場向け",
    "app_type_system": "QPSC / 補助脳系",
    "app_type_personal": "ユキズ専用",
    "app_type_frozen": "凍結",
    "app_type_archived": "保管",
    "app_type_unknown": "未分類",
    "completion_goal_formal_release": "正式出荷",
    "completion_goal_system_ready": "システム稼働",
    "completion_goal_reference_ready": "正本提示",
    "completion_goal_local_ready": "ローカル運用",
    "completion_goal_frozen_closed": "凍結完了",
    "completion_goal_unknown": "未設定",
    "reason_cloudflare_missing": "Cloudflare URLまたはProjectが未確認",
    "reason_health_attention": "health_urlまたはFunctions healthが未確認",
    "reason_site_git": "Git未反映または未commitがあります",
    "reason_series_git": "DAKE_seriesに未commitがあります",
    "source_app": "App Dashboard",
    "source_web": "Web Dashboard",
    "source_git": "DAKE_series",
    "source_error": "{source}: 状態取得に失敗しました",
    "error_missing_main": "main.py が見つかりません",
    "error_import_failed": "既存Dashboardを読み込めません: {error}",
    "error_missing_function": "既存Dashboardの取得関数が見つかりません",
    "error_scan_failed": "既存Dashboardの状態取得に失敗しました: {error}",
    "booth_assist_notice": "DAKE_BOOTH_Assistを起動しました: {folder}",
    "booth_assist_fallback": "DAKE_BOOTH_Assistが見つからないためフォルダを開きます。",
    "launch_check_template": "LAUNCH CHECK OK: apps={apps} booth={booth_urgent}/{booth_total} release={release_urgent}/{release_total} screenshot={screenshot_urgent}/{screenshot_total} readme={readme_urgent}/{readme_total} role={role_urgent}/{role_total} sites={sites} cloudflare={cloudflare_urgent}/{cloudflare_total} health={health_urgent}/{health_total} site_git={site_git_urgent}/{site_git_total} series_git={series_git}",
}

THEME = {
    "bg": "#F5F7FB",
    "panel": "#FFFFFF",
    "panel_alt": "#F9FAFD",
    "border": "#D7DCE7",
    "text": "#182230",
    "muted": "#667085",
    "quiet": "#98A2B3",
    "accent": "#2457C5",
    "accent_hover": "#1D4CB4",
    "danger": "#B42318",
    "danger_bg": "#FFF0EE",
    "warning": "#9A5B00",
    "warning_bg": "#FFF4D8",
    "success": "#1F7A4D",
    "success_bg": "#EAF7F0",
    "shadow": "#EEF2F7",
    "input": "#FFFFFF",
}

FONT_CANDIDATES = ["BIZ UDPGothic", "Yu Gothic UI", "Meiryo", "MS Gothic"]
EXCLUDED_APP_STATUS = {"internal", "frozen", "deprecated", "archived", "draft", "experimental", "private"}
APP_TYPE_DEFAULT = "market"
COMPLETION_GOAL_DEFAULT = "formal_release"
APP_TYPE_KEYS = ("market", "system", "personal", "frozen", "archived", "unknown")
COMPLETION_GOAL_KEYS = ("formal_release", "system_ready", "reference_ready", "local_ready", "frozen_closed", "unknown")
LATER_APP_NAMES = {"qpsc", "brainz", "oikawa", "orbit"}
LATER_SITE_STATUS = {"internal", "draft", "archived", "frozen", "deprecated"}
PRIORITY_URGENT = "urgent"
PRIORITY_ACTIVE = "active"
PRIORITY_LATER = "later"
PRIORITY_ORDER = (PRIORITY_URGENT, PRIORITY_ACTIVE, PRIORITY_LATER)
DAKE_META_PATTERN = re.compile(r"##\s*DAKE_META\b.*?```(?:json)?\s*(\{.*?\})\s*```", re.IGNORECASE | re.DOTALL)
DAKE_WEB_META_PATTERN = re.compile(r"##\s*DAKE_WEB_META\b.*?```(?:json)?\s*(\{.*?\})\s*```", re.IGNORECASE | re.DOTALL)


@dataclass(frozen=True)
class TargetItem:
    title: str
    detail: str
    path: Path
    folder_name: str
    url: str = ""
    priority: str = PRIORITY_ACTIVE

    def line(self) -> str:
        return f"{self.title} - {self.detail}"


@dataclass(frozen=True)
class AppRadar:
    total: int = 0
    market_count: int = 0
    system_count: int = 0
    personal_count: int = 0
    frozen_count: int = 0
    booth_missing: tuple[TargetItem, ...] = ()
    release_missing: tuple[TargetItem, ...] = ()
    screenshot_missing: tuple[TargetItem, ...] = ()
    readme_missing: tuple[TargetItem, ...] = ()
    role_attention: tuple[TargetItem, ...] = ()
    error: str = ""


@dataclass(frozen=True)
class SiteRadar:
    total: int = 0
    cloudflare_unchecked: tuple[TargetItem, ...] = ()
    health_attention: tuple[TargetItem, ...] = ()
    git_uncommitted: tuple[TargetItem, ...] = ()
    error: str = ""


@dataclass(frozen=True)
class GitRadar:
    series_uncommitted: int = 0
    series_untracked: int = 0
    error: str = ""


@dataclass(frozen=True)
class RadarSummary:
    app: AppRadar
    site: SiteRadar
    git: GitRadar
    loaded_at: datetime
    warnings: tuple[str, ...] = ()

    @property
    def urgent_total(self) -> int:
        return (
            priority_count(self.app.booth_missing, PRIORITY_URGENT)
            + priority_count(self.app.release_missing, PRIORITY_URGENT)
            + priority_count(self.app.screenshot_missing, PRIORITY_URGENT)
            + priority_count(self.app.readme_missing, PRIORITY_URGENT)
            + priority_count(self.app.role_attention, PRIORITY_URGENT)
            + priority_count(self.site.cloudflare_unchecked, PRIORITY_URGENT)
            + priority_count(self.site.health_attention, PRIORITY_URGENT)
            + priority_count(self.site.git_uncommitted, PRIORITY_URGENT)
            + self.git.series_uncommitted
            + len(self.warnings)
        )

    @property
    def action_total(self) -> int:
        return (
            len(self.app.booth_missing)
            + len(self.app.release_missing)
            + len(self.app.screenshot_missing)
            + len(self.app.readme_missing)
            + len(self.app.role_attention)
            + len(self.site.cloudflare_unchecked)
            + len(self.site.health_attention)
            + len(self.site.git_uncommitted)
            + self.git.series_uncommitted
            + len(self.warnings)
        )


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        if exe_dir.name.lower() == DIST_DIR_NAME:
            return exe_dir.parent
        return exe_dir
    return Path(__file__).resolve().parent


def icon_path() -> Path:
    if getattr(sys, "frozen", False):
        bundled = Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent)) / "dake_icon.ico"
        if bundled.exists():
            return bundled
    return DEFAULT_SERIES_ROOT / "02_assets" / "dake_icon.ico"


def set_windows_app_id() -> None:
    if not sys.platform.startswith("win") or ctypes is None:
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
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


def dashboard_dir(env_name: str, folder_name: str) -> Path:
    env_value = os.environ.get(env_name, "").strip()
    candidates: list[Path] = []
    if env_value:
        candidates.append(Path(env_value))
    if folder_name == APP_DASHBOARD_FOLDER:
        staging = app_dir().parent / "codex_staging_dake_dashboard_phase4"
    else:
        staging = app_dir().parent / "codex_staging_web_dashboard" / WEB_DASHBOARD_FOLDER
    candidates.extend([app_dir().parent / folder_name, DEFAULT_APPS_ROOT / folder_name, staging])
    for candidate in candidates:
        if (candidate / "main.py").exists():
            return candidate
    return candidates[0] if candidates else DEFAULT_APPS_ROOT / folder_name


def app_dashboard_dir() -> Path:
    return dashboard_dir("QPSC_APP_DASHBOARD_DIR", APP_DASHBOARD_FOLDER)


def web_dashboard_dir() -> Path:
    return dashboard_dir("QPSC_WEB_DASHBOARD_DIR", WEB_DASHBOARD_FOLDER)


def load_dashboard_module(name: str, folder: Path):
    module_path = folder / "main.py"
    if not module_path.exists():
        raise FileNotFoundError(UI_TEXT["error_missing_main"])
    module_name = f"qpsc_external_{name}_{abs(hash(str(module_path.resolve())))}"
    spec = importlib.util.spec_from_file_location(module_name, module_path)
    if spec is None or spec.loader is None:
        raise ImportError(UI_TEXT["error_import_failed"].format(error=module_path))
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    try:
        spec.loader.exec_module(module)
    except Exception as exc:
        raise ImportError(UI_TEXT["error_import_failed"].format(error=exc)) from exc
    return module


def read_text_utf8(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError:
        return path.read_text(encoding="utf-8", errors="replace")


def extract_meta(readme_path: Path, pattern: re.Pattern[str]) -> tuple[dict[str, object], str]:
    if not readme_path.exists():
        return {}, "missing_readme"
    try:
        text = read_text_utf8(readme_path)
    except OSError:
        return {}, "missing_readme"
    match = pattern.search(text)
    if not match:
        return {}, "missing_meta"
    try:
        data = json.loads(match.group(1))
    except json.JSONDecodeError:
        return {}, "json_error"
    if not isinstance(data, dict):
        return {}, "missing_meta"
    return data, ""


def safe_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def safe_bool(value: object) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.strip().lower() in {"true", "1", "yes", "on", UI_TEXT["value_yes"].lower()}
    if isinstance(value, (int, float)):
        return bool(value)
    return False


def priority_count(items: tuple[TargetItem, ...], priority: str) -> int:
    return sum(1 for item in items if item.priority == priority)


def priority_label(priority: str) -> str:
    return UI_TEXT.get(f"priority_{priority}", priority)


def priority_summary(items: tuple[TargetItem, ...]) -> str:
    return UI_TEXT["priority_summary"].format(
        urgent=priority_count(items, PRIORITY_URGENT),
        total=len(items),
        active=priority_count(items, PRIORITY_ACTIVE),
        later=priority_count(items, PRIORITY_LATER),
    )


def priority_items(items: tuple[TargetItem, ...], priority: str) -> tuple[TargetItem, ...]:
    return tuple(item for item in items if item.priority == priority)


def app_display_name(record: object, meta: dict[str, object]) -> str:
    for value in (meta.get("display_name"), meta.get("launcher_title"), getattr(record, "display_name", ""), getattr(record, "folder_name", "")):
        text = safe_text(value)
        if text:
            return text
    return safe_text(getattr(record, "folder_name", ""))


def target_for_app(record: object, meta: dict[str, object], reason_key: str, detail: str = "", priority: str = PRIORITY_ACTIVE) -> TargetItem:
    folder = Path(getattr(record, "folder_path", ""))
    return TargetItem(
        title=app_display_name(record, meta),
        detail=detail or UI_TEXT[reason_key],
        path=folder,
        folder_name=safe_text(getattr(record, "folder_name", folder.name)),
        priority=priority,
    )


def target_for_site(record: object, reason_key: str, url: str = "", priority: str = PRIORITY_ACTIVE) -> TargetItem:
    folder = Path(getattr(record, "folder_path", ""))
    title = safe_text(getattr(record, "display_name", "")) or safe_text(getattr(record, "folder_name", folder.name))
    return TargetItem(title=title, detail=UI_TEXT[reason_key], path=folder, folder_name=folder.name, url=url, priority=priority)


def meta_status(record: object, meta: dict[str, object]) -> str:
    value = safe_text(meta.get("status", ""))
    if value:
        return value.lower()
    fields = getattr(record, "meta_fields", {})
    if isinstance(fields, dict):
        return safe_text(fields.get("status", "")).lower()
    return ""


def normalized_choice(value: str, allowed: tuple[str, ...], default: str) -> str:
    normalized = value.strip().lower()
    return normalized if normalized in allowed else default


def meta_or_record_field(record: object, meta: dict[str, object], key: str) -> str:
    value = safe_text(meta.get(key, ""))
    if value:
        return value
    fields = getattr(record, "meta_fields", {})
    if isinstance(fields, dict):
        return safe_text(fields.get(key, ""))
    return ""


def app_type_value(record: object, meta: dict[str, object]) -> str:
    return normalized_choice(meta_or_record_field(record, meta, "app_type"), APP_TYPE_KEYS, APP_TYPE_DEFAULT)


def completion_goal_value(record: object, meta: dict[str, object]) -> str:
    return normalized_choice(meta_or_record_field(record, meta, "completion_goal"), COMPLETION_GOAL_KEYS, COMPLETION_GOAL_DEFAULT)


def app_type_label(value: str) -> str:
    key = normalized_choice(value, APP_TYPE_KEYS, "unknown")
    return UI_TEXT.get(f"app_type_{key}", UI_TEXT["app_type_unknown"])


def completion_goal_label(value: str) -> str:
    key = normalized_choice(value, COMPLETION_GOAL_KEYS, "unknown")
    return UI_TEXT.get(f"completion_goal_{key}", UI_TEXT["completion_goal_unknown"])


def is_market_formal_app(record: object, meta: dict[str, object]) -> bool:
    status = meta_status(record, meta)
    if status in EXCLUDED_APP_STATUS:
        return False
    return app_type_value(record, meta) == "market" and completion_goal_value(record, meta) == "formal_release"


def meta_show_on_site(meta: dict[str, object]) -> bool:
    return safe_bool(meta.get("show_on_site", False))


def meta_show_in_launcher(record: object, meta: dict[str, object]) -> bool:
    value = meta.get("show_in_launcher")
    if value is None:
        fields = getattr(record, "meta_fields", {})
        if isinstance(fields, dict):
            value = fields.get("show_in_launcher")
    return safe_bool(value)


def app_has_dist(record: object) -> bool:
    checks = getattr(record, "checks", None)
    if bool(getattr(checks, "has_dist_exe", False)):
        return True
    folder = Path(getattr(record, "folder_path", ""))
    try:
        return any((folder / DIST_DIR_NAME).glob("*.exe"))
    except OSError:
        return False


def app_booth_product_candidates(folder: Path) -> tuple[Path, ...]:
    return (
        folder / "booth_product.txt",
        folder / "booth_ready" / "booth_product.txt",
    )


def find_app_booth_product_path(folder: Path) -> Path | None:
    for candidate in app_booth_product_candidates(folder):
        if candidate.is_file():
            return candidate
    return None


def app_booth_material_flags(record: object) -> tuple[bool, bool, bool, bool]:
    folder = Path(getattr(record, "folder_path", ""))
    checks = getattr(record, "checks", None)
    ready_booth_product = folder / "booth_ready" / "booth_product.txt"
    ready_dir = folder / "booth_ready"
    thumbnail = folder / "assets" / "booth_thumbnail.jpg"
    has_product = bool(getattr(checks, "has_booth_product", False)) or find_app_booth_product_path(folder) is not None
    has_ready_product = ready_booth_product.exists()
    has_ready_dir = bool(getattr(checks, "has_booth_ready", ready_dir.is_dir()))
    has_thumbnail = bool(getattr(checks, "has_booth_thumbnail", thumbnail.exists()))
    return has_product, has_ready_product, has_ready_dir, has_thumbnail


def app_has_release_body(folder: Path) -> bool:
    return (folder / RELEASE_BODY_NAME).exists()


def app_is_later_market_target(record: object, meta: dict[str, object]) -> bool:
    if not is_market_formal_app(record, meta):
        return True
    folder_name = safe_text(getattr(record, "folder_name", ""))
    title = app_display_name(record, meta)
    combined = f"{folder_name} {title}".lower()
    if not meta_show_on_site(meta) and not meta_show_in_launcher(record, meta):
        return True
    return any(token in combined for token in LATER_APP_NAMES)


def is_public_app(record: object, meta: dict[str, object]) -> bool:
    return is_market_formal_app(record, meta) and (meta_status(record, meta) == "available" or meta_show_on_site(meta))


def is_launcher_internal_screenshot_target(record: object, meta: dict[str, object]) -> bool:
    if is_public_app(record, meta):
        return True
    return meta_show_in_launcher(record, meta)


def app_release_url(record: object, meta: dict[str, object]) -> str:
    return safe_text(meta.get("release_url", "")) or safe_text(getattr(record, "release_url", ""))


def app_screenshot_missing(record: object, meta: dict[str, object]) -> bool:
    folder = Path(getattr(record, "folder_path", ""))
    screenshot_value = safe_text(meta.get("screenshot_path", ""))
    if not screenshot_value:
        return True
    screenshot_path = Path(screenshot_value)
    if not screenshot_path.is_absolute():
        screenshot_path = folder / screenshot_path
    return not screenshot_path.exists()


def app_readme_issue(folder: Path, record: object, meta: dict[str, object], issue_key: str) -> tuple[str, str]:
    status = meta_status(record, meta)
    downgrade_later = status in {"frozen", "deprecated", "archived"}
    if issue_key == "missing_readme":
        return UI_TEXT["reason_readme_missing"], PRIORITY_URGENT
    if issue_key == "json_error":
        return UI_TEXT["reason_meta_json"], PRIORITY_URGENT
    if issue_key == "missing_meta":
        return UI_TEXT["reason_meta_missing"], PRIORITY_URGENT
    missing = [key for key in ("display_name", "folder_name", "exe_name", "status") if not safe_text(meta.get(key, ""))]
    if missing:
        return f"{UI_TEXT['reason_meta_fields_missing']}: {', '.join(missing)}", PRIORITY_URGENT
    readme_text = ""
    try:
        readme_text = read_text_utf8(folder / README_NAME)
    except OSError:
        pass
    if "RELEASE_BODY" not in readme_text or not app_has_release_body(folder):
        priority = PRIORITY_LATER if downgrade_later else PRIORITY_ACTIVE
        return UI_TEXT["reason_release_body_missing"], priority
    if not safe_text(meta.get("update_summary", "")):
        priority = PRIORITY_LATER if downgrade_later else PRIORITY_ACTIVE
        return UI_TEXT["reason_update_summary_missing"], priority
    return "", PRIORITY_ACTIVE


def classify_booth_priority(record: object, meta: dict[str, object]) -> str:
    if not is_market_formal_app(record, meta):
        return PRIORITY_LATER
    status = meta_status(record, meta)
    has_release = bool(app_release_url(record, meta))
    has_distribution = app_has_dist(record) or has_release
    visible = meta_show_on_site(meta) or meta_show_in_launcher(record, meta)
    has_booth_product, _has_ready_product, has_ready_dir, has_thumbnail = app_booth_material_flags(record)
    materials_complete = has_booth_product and has_ready_dir and has_thumbnail
    needs_only_booth_url = materials_complete and not safe_text(meta.get("booth_url", ""))
    if app_is_later_market_target(record, meta):
        return PRIORITY_LATER
    if status == "available" and has_release and has_distribution and visible and needs_only_booth_url:
        return PRIORITY_URGENT
    if status == "available":
        return PRIORITY_ACTIVE
    return PRIORITY_LATER


def classify_release_priority(record: object, meta: dict[str, object], has_readme_issue: bool) -> str:
    if not is_market_formal_app(record, meta):
        return PRIORITY_LATER
    status = meta_status(record, meta)
    if status in EXCLUDED_APP_STATUS:
        return PRIORITY_LATER
    if status == "available" and meta_show_on_site(meta):
        if not has_readme_issue and not app_screenshot_missing(record, meta) and app_has_dist(record):
            return PRIORITY_URGENT
        return PRIORITY_ACTIVE
    if status == "available":
        return PRIORITY_ACTIVE
    return PRIORITY_LATER


def classify_screenshot_priority(record: object, meta: dict[str, object]) -> str:
    if not is_market_formal_app(record, meta):
        return PRIORITY_LATER
    status = meta_status(record, meta)
    if status in {"frozen", "deprecated", "archived"}:
        return PRIORITY_LATER
    if status == "available" and meta_show_on_site(meta):
        return PRIORITY_URGENT
    if meta_show_in_launcher(record, meta):
        return PRIORITY_ACTIVE
    return PRIORITY_LATER


def app_booth_missing(record: object, meta: dict[str, object]) -> bool:
    if not is_market_formal_app(record, meta):
        return False
    booth_url = safe_text(meta.get("booth_url", ""))
    has_booth_product, _has_ready_product, has_booth_ready, has_booth_thumbnail = app_booth_material_flags(record)
    status = meta_status(record, meta)
    materials_missing = not has_booth_product or not has_booth_ready or not has_booth_thumbnail
    return (
        not has_booth_product
        or not booth_url
        or (status == "available" and materials_missing)
    )


def app_role_family(record: object, meta: dict[str, object]) -> str:
    app_type = app_type_value(record, meta)
    if app_type in {"market", "system", "personal", "frozen"}:
        return app_type
    return "frozen" if app_type == "archived" else "unknown"


def read_optional_text(path: Path) -> str:
    try:
        return read_text_utf8(path)
    except OSError:
        return ""


def app_role_attention(record: object, meta: dict[str, object], readme_text: str) -> tuple[str, str]:
    if is_market_formal_app(record, meta):
        return "", PRIORITY_ACTIVE

    app_type = app_type_value(record, meta)
    goal = completion_goal_value(record, meta)
    folder = Path(getattr(record, "folder_path", ""))
    detail_bits = [f"{app_type_label(app_type)} / {completion_goal_label(goal)}"]

    if app_type == "unknown" or goal == "unknown":
        return UI_TEXT["reason_role_unknown"], PRIORITY_ACTIVE

    expected_goals = {
        "system": {"system_ready", "reference_ready"},
        "personal": {"local_ready"},
        "frozen": {"frozen_closed"},
        "archived": {"frozen_closed", "reference_ready"},
    }
    if goal not in expected_goals.get(app_type, {"formal_release"}):
        return f"{UI_TEXT['reason_role_goal_mismatch']} ({' / '.join(detail_bits)})", PRIORITY_ACTIVE

    missing: list[str] = []
    if app_type == "system" and goal == "system_ready":
        if not (folder / "build.bat").exists():
            missing.append("build.bat")
        main_text = read_optional_text(folder / "main.py")
        if "--launch-check" not in readme_text and "--launch-check" not in main_text:
            missing.append("--launch-check")
        if "system_ready" not in readme_text and "Positioning" not in readme_text:
            missing.append("role docs")
        if missing:
            return f"{UI_TEXT['reason_system_ready_attention']} {UI_TEXT['reason_role_missing_items'].format(items=', '.join(missing))}", PRIORITY_ACTIVE
    elif app_type == "system" and goal == "reference_ready":
        if "reference_ready" not in readme_text and "正本" not in readme_text:
            return UI_TEXT["reason_reference_ready_attention"], PRIORITY_ACTIVE
    elif app_type == "personal" and goal == "local_ready":
        if "local_ready" not in readme_text and "ローカル" not in readme_text:
            return UI_TEXT["reason_personal_ready_attention"], PRIORITY_LATER
    elif app_type == "frozen" and goal == "frozen_closed":
        if "frozen" not in readme_text.lower() and "凍結" not in readme_text:
            return UI_TEXT["reason_frozen_ready_attention"], PRIORITY_LATER

    return "", PRIORITY_ACTIVE


def collect_app_radar() -> tuple[AppRadar, GitRadar, tuple[str, ...]]:
    folder = app_dashboard_dir()
    try:
        module = load_dashboard_module("app", folder)
        if not hasattr(module, "scan_apps"):
            raise AttributeError(UI_TEXT["error_missing_function"])
        records = list(module.scan_apps(folder.parent))
        booth_missing: list[TargetItem] = []
        release_missing: list[TargetItem] = []
        screenshot_missing: list[TargetItem] = []
        readme_missing: list[TargetItem] = []
        role_attention: list[TargetItem] = []
        role_counts = {"market": 0, "system": 0, "personal": 0, "frozen": 0}
        for record in records:
            record_folder = Path(getattr(record, "folder_path", ""))
            meta, issue_key = extract_meta(record_folder / README_NAME, DAKE_META_PATTERN)
            role_family = app_role_family(record, meta)
            if role_family in role_counts:
                role_counts[role_family] += 1
            reason, readme_priority = app_readme_issue(record_folder, record, meta, issue_key)
            readme_text = read_optional_text(record_folder / README_NAME)
            if reason:
                target = target_for_app(record, meta, "reason_meta_missing", reason, readme_priority)
                if is_market_formal_app(record, meta):
                    readme_missing.append(target)
                else:
                    role_attention.append(target)
            if is_market_formal_app(record, meta):
                if app_booth_missing(record, meta):
                    booth_missing.append(target_for_app(record, meta, "reason_booth_missing", priority=classify_booth_priority(record, meta)))
                if not app_release_url(record, meta):
                    release_priority = classify_release_priority(record, meta, bool(reason))
                    release_missing.append(target_for_app(record, meta, "reason_release_missing", priority=release_priority))
                if app_screenshot_missing(record, meta):
                    screenshot_missing.append(target_for_app(record, meta, "reason_screenshot_missing", priority=classify_screenshot_priority(record, meta)))
            else:
                role_reason, role_priority = app_role_attention(record, meta, readme_text)
                if role_reason:
                    role_attention.append(target_for_app(record, meta, "reason_role_unknown", role_reason, role_priority))
        return (
            AppRadar(
                total=len(records),
                market_count=role_counts["market"],
                system_count=role_counts["system"],
                personal_count=role_counts["personal"],
                frozen_count=role_counts["frozen"],
                booth_missing=tuple(booth_missing),
                release_missing=tuple(release_missing),
                screenshot_missing=tuple(screenshot_missing),
                readme_missing=tuple(readme_missing),
                role_attention=tuple(role_attention),
            ),
            collect_git_radar(module, folder.parent.parent),
            (),
        )
    except Exception as exc:
        warning = UI_TEXT["source_error"].format(source=UI_TEXT["source_app"])
        return AppRadar(error=UI_TEXT["error_scan_failed"].format(error=exc)), GitRadar(error=str(exc)), (warning,)


def collect_git_radar(app_module, series_root_path: Path) -> GitRadar:
    try:
        read_git_status = getattr(app_module, "read_git_status")
        status = read_git_status(series_root_path)
        return GitRadar(
            series_uncommitted=int(getattr(status, "uncommitted_count", 0) or 0),
            series_untracked=int(getattr(status, "untracked_count", 0) or 0),
            error=safe_text(getattr(status, "error", "")),
        )
    except Exception as exc:
        return GitRadar(error=str(exc))


def site_cloudflare_unchecked(record: object, meta: dict[str, object]) -> bool:
    files = getattr(record, "files", None)
    cloudflare = getattr(record, "cloudflare", None)
    class_key = safe_text(getattr(record, "class_key", ""))
    production_url = safe_text(getattr(record, "production_url", ""))
    domain = safe_text(getattr(record, "domain", ""))
    health_url = safe_text(getattr(record, "health_url", ""))
    cloudflare_url = safe_text(getattr(record, "cloudflare_url", ""))
    cloudflare_project = safe_text(getattr(record, "cloudflare_project", ""))
    meta_cloudflare_url = first_safe_text(
        meta.get("cloudflare_url"),
        meta.get("cloudflare_project_url"),
        meta.get("cloudflare_dashboard_url"),
    )
    likely_pages = bool(getattr(cloudflare, "likely_pages", False)) or bool(getattr(files, "has_wrangler", False))
    has_production = bool(production_url or domain)
    health_unknown = bool(health_url) and (
        not bool(getattr(cloudflare, "has_health_file", False)) or not bool(getattr(cloudflare, "has_routes_api", False))
    )
    return (
        class_key == "deploy_review"
        or (likely_pages and not meta_cloudflare_url and not cloudflare_url)
        or (has_production and not cloudflare_url)
        or (has_production and not cloudflare_project)
        or health_unknown
    )


def site_status(record: object, meta: dict[str, object]) -> str:
    return (safe_text(meta.get("status", "")) or safe_text(getattr(record, "status_value", ""))).lower()


def site_has_production(record: object) -> bool:
    return bool(safe_text(getattr(record, "production_url", "")) or safe_text(getattr(record, "domain", "")))


def site_is_public_target(record: object, meta: dict[str, object]) -> bool:
    status = site_status(record, meta)
    if status in LATER_SITE_STATUS:
        return False
    if not bool(getattr(record, "show_on_dashboard", True)):
        return False
    return site_has_production(record)


def classify_cloudflare_priority(record: object, meta: dict[str, object]) -> str:
    status = site_status(record, meta)
    if not site_has_production(record) or status in LATER_SITE_STATUS:
        return PRIORITY_LATER
    if site_is_public_target(record, meta):
        return PRIORITY_URGENT
    return PRIORITY_ACTIVE


def classify_health_priority(record: object, meta: dict[str, object]) -> str:
    files = getattr(record, "files", None)
    cloudflare = getattr(record, "cloudflare", None)
    has_api = bool(getattr(files, "has_functions", False)) or bool(getattr(cloudflare, "has_functions_api", False))
    if has_api and site_is_public_target(record, meta):
        return PRIORITY_URGENT
    if has_api:
        return PRIORITY_ACTIVE
    return PRIORITY_LATER


def classify_site_git_priority(record: object, meta: dict[str, object]) -> str:
    if site_is_public_target(record, meta):
        return PRIORITY_URGENT
    if site_has_production(record):
        return PRIORITY_ACTIVE
    return PRIORITY_LATER


def site_health_attention(record: object) -> bool:
    files = getattr(record, "files", None)
    cloudflare = getattr(record, "cloudflare", None)
    has_functions = bool(getattr(files, "has_functions", False))
    health_url = safe_text(getattr(record, "health_url", ""))
    return has_functions and (
        not health_url
        or not bool(getattr(cloudflare, "has_health_file", False))
        or not bool(getattr(cloudflare, "has_routes_api", False))
    )


def first_safe_text(*values: object) -> str:
    for value in values:
        text = safe_text(value)
        if text:
            return text
    return ""


def collect_site_radar() -> tuple[SiteRadar, tuple[str, ...]]:
    folder = web_dashboard_dir()
    try:
        module = load_dashboard_module("web", folder)
        if not hasattr(module, "scan_sites"):
            raise AttributeError(UI_TEXT["error_missing_function"])
        dev_root = getattr(module, "DEV_ROOT", DEFAULT_SERIES_ROOT.parent)
        records = list(module.scan_sites(dev_root))
        cloudflare_unchecked: list[TargetItem] = []
        health_attention: list[TargetItem] = []
        git_uncommitted: list[TargetItem] = []
        for record in records:
            record_folder = Path(getattr(record, "folder_path", ""))
            meta, _issue_key = extract_meta(record_folder / README_NAME, DAKE_WEB_META_PATTERN)
            cloudflare_url = safe_text(getattr(record, "cloudflare_url", ""))
            if site_cloudflare_unchecked(record, meta):
                cloudflare_unchecked.append(
                    target_for_site(record, "reason_cloudflare_missing", cloudflare_url, classify_cloudflare_priority(record, meta))
                )
            if site_health_attention(record):
                health_attention.append(
                    target_for_site(record, "reason_health_attention", safe_text(getattr(record, "health_url", "")), classify_health_priority(record, meta))
                )
            git = getattr(record, "git", None)
            if bool(getattr(git, "has_dirty", False)) or int(getattr(git, "ahead", 0) or 0) > 0:
                git_uncommitted.append(target_for_site(record, "reason_site_git", priority=classify_site_git_priority(record, meta)))
        return (
            SiteRadar(
                total=len(records),
                cloudflare_unchecked=tuple(cloudflare_unchecked),
                health_attention=tuple(health_attention),
                git_uncommitted=tuple(git_uncommitted),
            ),
            (),
        )
    except Exception as exc:
        warning = UI_TEXT["source_error"].format(source=UI_TEXT["source_web"])
        return SiteRadar(error=UI_TEXT["error_scan_failed"].format(error=exc)), (warning,)


def collect_radar_summary() -> RadarSummary:
    app_radar, git_radar, app_warnings = collect_app_radar()
    site_radar, site_warnings = collect_site_radar()
    return RadarSummary(app=app_radar, site=site_radar, git=git_radar, loaded_at=datetime.now(), warnings=tuple(app_warnings + site_warnings))


def launch_target(folder: Path, exe_name: str) -> Path:
    exe_path = folder / DIST_DIR_NAME / exe_name
    if exe_path.exists():
        return exe_path
    main_path = folder / "main.py"
    if main_path.exists():
        return main_path
    return exe_path


def launch_dashboard(folder: Path, exe_name: str) -> None:
    target = launch_target(folder, exe_name)
    if not target.exists():
        raise FileNotFoundError(str(target))
    if target.suffix.lower() == ".exe":
        subprocess.Popen([str(target)], cwd=str(folder))
        return
    subprocess.Popen([sys.executable, str(target)], cwd=str(folder))


def open_path(path: Path) -> None:
    if sys.platform.startswith("win"):
        os.startfile(str(path))  # type: ignore[attr-defined]
        return
    webbrowser.open(path.as_uri())


def open_url_or_path(url: str, path: Path) -> None:
    if url:
        webbrowser.open(url)
        return
    open_path(path)


class QpscDashboardApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.font_family = choose_font_family(root)
        self.worker_thread: threading.Thread | None = None
        self.worker_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.summary: RadarSummary | None = None
        self.app_vars = {
            "total": tk.StringVar(value=UI_TEXT["value_waiting"]),
            "booth": tk.StringVar(value=UI_TEXT["value_waiting"]),
            "release": tk.StringVar(value=UI_TEXT["value_waiting"]),
            "screenshot": tk.StringVar(value=UI_TEXT["value_waiting"]),
            "readme": tk.StringVar(value=UI_TEXT["value_waiting"]),
            "role": tk.StringVar(value=UI_TEXT["value_waiting"]),
            "market": tk.StringVar(value=UI_TEXT["value_waiting"]),
            "system": tk.StringVar(value=UI_TEXT["value_waiting"]),
            "personal": tk.StringVar(value=UI_TEXT["value_waiting"]),
            "frozen": tk.StringVar(value=UI_TEXT["value_waiting"]),
        }
        self.site_vars = {
            "total": tk.StringVar(value=UI_TEXT["value_waiting"]),
            "cloudflare": tk.StringVar(value=UI_TEXT["value_waiting"]),
            "health": tk.StringVar(value=UI_TEXT["value_waiting"]),
            "git": tk.StringVar(value=UI_TEXT["value_waiting"]),
        }
        self.git_vars = {
            "series_uncommitted": tk.StringVar(value=UI_TEXT["value_waiting"]),
            "series_untracked": tk.StringVar(value=UI_TEXT["value_waiting"]),
            "attention": tk.StringVar(value=UI_TEXT["value_waiting"]),
        }
        self.summary_var = tk.StringVar(value=UI_TEXT["last_loaded_waiting"])
        self.status_var = tk.StringVar(value=UI_TEXT["status_checking"])
        self.next_frame: tk.Frame | None = None
        self.configure_root()
        self.build_ui()
        self.root.after(120, self.refresh)

    def configure_root(self) -> None:
        self.root.title(UI_TEXT["window_title"])
        self.root.geometry("1200x800")
        self.root.minsize(1080, 720)
        self.root.configure(bg=THEME["bg"])
        apply_window_icon(self.root)

    def build_ui(self) -> None:
        outer = tk.Frame(self.root, bg=THEME["bg"])
        outer.pack(fill="both", expand=True, padx=24, pady=22)
        header = tk.Frame(outer, bg=THEME["bg"])
        header.pack(fill="x")
        title_area = tk.Frame(header, bg=THEME["bg"])
        title_area.pack(side="left", anchor="w")
        tk.Label(title_area, text=UI_TEXT["header_title"], bg=THEME["bg"], fg=THEME["text"], font=(self.font_family, 26, "bold")).pack(anchor="w")
        tk.Label(title_area, text=UI_TEXT["header_kicker"], bg=THEME["bg"], fg=THEME["quiet"], font=(self.font_family, 10, "bold")).pack(anchor="w", pady=(3, 0))
        tk.Label(title_area, text=UI_TEXT["header_subtitle"], bg=THEME["bg"], fg=THEME["muted"], font=(self.font_family, 12)).pack(anchor="w", pady=(5, 0))
        actions = tk.Frame(header, bg=THEME["bg"])
        actions.pack(side="right", anchor="e")
        self.reload_button = self.make_button(actions, UI_TEXT["button_reload"], self.refresh, primary=True)
        self.reload_button.pack(side="left", padx=(0, 8))
        self.make_button(actions, UI_TEXT["button_open_app_dashboard"], self.open_app_dashboard).pack(side="left", padx=(0, 8))
        self.make_button(actions, UI_TEXT["button_open_web_dashboard"], self.open_web_dashboard).pack(side="left")

        tk.Label(outer, text=UI_TEXT["section_current_title"], bg=THEME["bg"], fg=THEME["text"], font=(self.font_family, 13, "bold")).pack(anchor="w", pady=(24, 8))
        cards = tk.Frame(outer, bg=THEME["bg"])
        cards.pack(fill="x", pady=(0, 18))
        cards.grid_columnconfigure(0, weight=1, uniform="cards")
        cards.grid_columnconfigure(1, weight=1, uniform="cards")
        self.build_app_card(cards).grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        self.build_site_card(cards).grid(row=0, column=1, sticky="nsew", padx=(10, 0))

        lower = tk.Frame(outer, bg=THEME["bg"])
        lower.pack(fill="both", expand=True)
        lower.grid_columnconfigure(0, weight=3, uniform="lower")
        lower.grid_columnconfigure(1, weight=2, uniform="lower")
        self.build_next_card(lower).grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        self.build_git_card(lower).grid(row=0, column=1, sticky="nsew", padx=(10, 0))

        footer = tk.Frame(self.root, bg=THEME["bg"])
        footer.pack(fill="x", padx=24, pady=(0, 12))
        tk.Label(footer, textvariable=self.status_var, bg=THEME["bg"], fg=THEME["muted"], font=(self.font_family, 10)).pack(side="left")
        tk.Label(footer, textvariable=self.summary_var, bg=THEME["bg"], fg=THEME["quiet"], font=(self.font_family, 9)).pack(side="right")

    def make_card(self, parent: tk.Misc) -> tk.Frame:
        return tk.Frame(parent, bg=THEME["panel"], highlightthickness=1, highlightbackground=THEME["border"])

    def make_button(self, parent: tk.Misc, text: str, command, primary: bool = False) -> tk.Button:
        bg = THEME["accent"] if primary else THEME["panel"]
        fg = "#FFFFFF" if primary else THEME["text"]
        active_bg = THEME["accent_hover"] if primary else THEME["panel_alt"]
        return tk.Button(parent, text=text, command=command, bg=bg, fg=fg, activebackground=active_bg, activeforeground=fg, relief="flat", bd=0, padx=14, pady=8, cursor="hand2", font=(self.font_family, 9, "bold" if primary else "normal"))

    def build_app_card(self, parent: tk.Misc) -> tk.Frame:
        frame = self.make_card(parent)
        self.card_title(frame, UI_TEXT["card_app_title"])
        body = tk.Frame(frame, bg=THEME["panel"])
        body.pack(fill="x", padx=16, pady=(2, 16))
        self.metric_row(body, UI_TEXT["label_app_total"], self.app_vars["total"], None).pack(fill="x", pady=(0, 8))
        self.metric_row(body, UI_TEXT["label_booth_missing"], self.app_vars["booth"], lambda: self.show_targets("booth")).pack(fill="x", pady=(0, 8))
        self.metric_row(body, UI_TEXT["label_release_missing"], self.app_vars["release"], lambda: self.show_targets("release")).pack(fill="x", pady=(0, 8))
        self.metric_row(body, UI_TEXT["label_screenshot_missing"], self.app_vars["screenshot"], lambda: self.show_targets("screenshot")).pack(fill="x", pady=(0, 8))
        self.metric_row(body, UI_TEXT["label_readme_missing"], self.app_vars["readme"], lambda: self.show_targets("readme")).pack(fill="x", pady=(0, 8))
        self.metric_row(body, UI_TEXT["label_role_attention"], self.app_vars["role"], lambda: self.show_targets("role")).pack(fill="x", pady=(0, 8))
        self.metric_row(body, UI_TEXT["label_market_count"], self.app_vars["market"], None).pack(fill="x", pady=(0, 8))
        self.metric_row(body, UI_TEXT["label_system_count"], self.app_vars["system"], None).pack(fill="x", pady=(0, 8))
        self.metric_row(body, UI_TEXT["label_personal_count"], self.app_vars["personal"], None).pack(fill="x", pady=(0, 8))
        self.metric_row(body, UI_TEXT["label_frozen_count"], self.app_vars["frozen"], None).pack(fill="x")
        return frame

    def build_site_card(self, parent: tk.Misc) -> tk.Frame:
        frame = self.make_card(parent)
        self.card_title(frame, UI_TEXT["card_site_title"])
        body = tk.Frame(frame, bg=THEME["panel"])
        body.pack(fill="x", padx=16, pady=(2, 16))
        self.metric_row(body, UI_TEXT["label_site_total"], self.site_vars["total"], None).pack(fill="x", pady=(0, 8))
        self.metric_row(body, UI_TEXT["label_cloudflare_unchecked"], self.site_vars["cloudflare"], lambda: self.show_targets("cloudflare")).pack(fill="x", pady=(0, 8))
        self.metric_row(body, UI_TEXT["label_health_attention"], self.site_vars["health"], lambda: self.show_targets("health")).pack(fill="x", pady=(0, 8))
        self.metric_row(body, UI_TEXT["label_site_git_uncommitted"], self.site_vars["git"], lambda: self.show_targets("site_git")).pack(fill="x")
        return frame

    def build_git_card(self, parent: tk.Misc) -> tk.Frame:
        frame = self.make_card(parent)
        self.card_title(frame, UI_TEXT["card_git_title"])
        body = tk.Frame(frame, bg=THEME["panel"])
        body.pack(fill="x", padx=16, pady=(2, 16))
        self.metric_row(body, UI_TEXT["label_series_uncommitted"], self.git_vars["series_uncommitted"], lambda: self.show_targets("series_git")).pack(fill="x", pady=(0, 8))
        self.metric_row(body, UI_TEXT["label_series_untracked"], self.git_vars["series_untracked"], lambda: self.show_targets("series_git")).pack(fill="x", pady=(0, 8))
        self.metric_row(body, UI_TEXT["label_git_attention"], self.git_vars["attention"], lambda: self.show_targets("series_git")).pack(fill="x")
        return frame

    def build_next_card(self, parent: tk.Misc) -> tk.Frame:
        frame = self.make_card(parent)
        self.card_title(frame, UI_TEXT["card_next_title"])
        self.next_frame = tk.Frame(frame, bg=THEME["panel"])
        self.next_frame.pack(fill="both", expand=True, padx=16, pady=(2, 16))
        self.render_next_actions(None)
        return frame

    def card_title(self, parent: tk.Misc, text: str) -> None:
        tk.Label(parent, text=text, bg=THEME["panel"], fg=THEME["text"], font=(self.font_family, 14, "bold")).pack(anchor="w", padx=16, pady=(14, 10))

    def metric_row(self, parent: tk.Misc, label: str, value_var: tk.StringVar, command) -> tk.Frame:
        frame = tk.Frame(parent, bg=THEME["panel_alt"], highlightthickness=1, highlightbackground=THEME["shadow"])
        cursor = "hand2" if command else ""
        label_widget = tk.Label(frame, text=label, bg=THEME["panel_alt"], fg=THEME["muted"], font=(self.font_family, 9), cursor=cursor)
        label_widget.pack(side="left", padx=12, pady=11)
        value_widget = tk.Label(frame, textvariable=value_var, bg=THEME["panel_alt"], fg=THEME["text"], justify="right", font=(self.font_family, 12, "bold"), cursor=cursor)
        value_widget.pack(side="right", padx=12, pady=8)
        if command:
            frame.configure(cursor="hand2")
            for widget in (frame, label_widget, value_widget):
                widget.bind("<Button-1>", lambda _event, action=command: action())
        return frame

    def refresh(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            return
        self.status_var.set(UI_TEXT["status_checking"])
        self.reload_button.configure(state="disabled")
        self.worker_thread = threading.Thread(target=self.scan_worker, daemon=True)
        self.worker_thread.start()
        self.root.after(WORKER_POLL_MS, self.poll_worker)

    def scan_worker(self) -> None:
        try:
            self.worker_queue.put(("summary", collect_radar_summary()))
        except Exception as exc:
            self.worker_queue.put(("error", exc))

    def poll_worker(self) -> None:
        try:
            while True:
                event, payload = self.worker_queue.get_nowait()
                if event == "summary" and isinstance(payload, RadarSummary):
                    self.apply_summary(payload)
                elif event == "error":
                    self.status_var.set(UI_TEXT["status_attention"])
        except queue.Empty:
            pass
        if self.worker_thread and self.worker_thread.is_alive():
            self.root.after(WORKER_POLL_MS, self.poll_worker)
        else:
            self.reload_button.configure(state="normal")

    def apply_summary(self, summary: RadarSummary) -> None:
        self.summary = summary
        self.app_vars["total"].set(str(summary.app.total))
        self.app_vars["booth"].set(priority_summary(summary.app.booth_missing))
        self.app_vars["release"].set(priority_summary(summary.app.release_missing))
        self.app_vars["screenshot"].set(priority_summary(summary.app.screenshot_missing))
        self.app_vars["readme"].set(priority_summary(summary.app.readme_missing))
        self.app_vars["role"].set(priority_summary(summary.app.role_attention))
        self.app_vars["market"].set(str(summary.app.market_count))
        self.app_vars["system"].set(str(summary.app.system_count))
        self.app_vars["personal"].set(str(summary.app.personal_count))
        self.app_vars["frozen"].set(str(summary.app.frozen_count))
        self.site_vars["total"].set(str(summary.site.total))
        self.site_vars["cloudflare"].set(priority_summary(summary.site.cloudflare_unchecked))
        self.site_vars["health"].set(priority_summary(summary.site.health_attention))
        self.site_vars["git"].set(priority_summary(summary.site.git_uncommitted))
        self.git_vars["series_uncommitted"].set(str(summary.git.series_uncommitted))
        self.git_vars["series_untracked"].set(str(summary.git.series_untracked))
        self.git_vars["attention"].set(UI_TEXT["value_git_attention"] if summary.git.error or summary.git.series_uncommitted else UI_TEXT["value_git_ok"])
        self.summary_var.set(UI_TEXT["summary_template"].format(urgent=summary.urgent_total, total=summary.action_total, apps=summary.app.total, sites=summary.site.total))
        self.status_var.set(UI_TEXT["status_attention"] if summary.urgent_total else UI_TEXT["status_ready"])
        self.render_next_actions(summary)

    def render_next_actions(self, summary: RadarSummary | None) -> None:
        if self.next_frame is None:
            return
        for child in self.next_frame.winfo_children():
            child.destroy()
        actions = self.next_actions(summary)
        if not actions:
            tk.Label(self.next_frame, text=UI_TEXT["next_none"], bg=THEME["panel"], fg=THEME["muted"], font=(self.font_family, 11)).pack(anchor="w", pady=(4, 0))
            return
        for index, (label_key, count, target_key) in enumerate(actions, start=1):
            text = UI_TEXT["next_line"].format(index=index, label=UI_TEXT[label_key], count=count)
            button = self.make_button(self.next_frame, text, lambda key=target_key: self.show_targets(key))
            button.pack(fill="x", anchor="w", pady=(0, 8))

    def next_actions(self, summary: RadarSummary | None) -> list[tuple[str, int, str]]:
        if summary is None:
            return []
        candidates = [
            ("next_booth", priority_count(summary.app.booth_missing, PRIORITY_URGENT), "booth"),
            ("next_screenshot", priority_count(summary.app.screenshot_missing, PRIORITY_URGENT), "screenshot"),
            ("next_release", priority_count(summary.app.release_missing, PRIORITY_URGENT), "release"),
            ("next_readme", priority_count(summary.app.readme_missing, PRIORITY_URGENT), "readme"),
            ("next_role", priority_count(summary.app.role_attention, PRIORITY_URGENT), "role"),
            ("next_cloudflare", priority_count(summary.site.cloudflare_unchecked, PRIORITY_URGENT), "cloudflare"),
            ("next_health", priority_count(summary.site.health_attention, PRIORITY_URGENT), "health"),
            ("next_site_git", priority_count(summary.site.git_uncommitted, PRIORITY_URGENT), "site_git"),
            ("next_series_git", summary.git.series_uncommitted, "series_git"),
        ]
        return [item for item in candidates if item[1]][:MAX_NEXT_ACTIONS]

    def targets_for_key(self, key: str) -> tuple[str, tuple[TargetItem, ...]]:
        summary = self.summary
        if summary is None:
            return UI_TEXT["dialog_select_notice"], ()
        if key == "booth":
            return UI_TEXT["dialog_booth_title"], summary.app.booth_missing
        if key == "release":
            return UI_TEXT["dialog_release_title"], summary.app.release_missing
        if key == "screenshot":
            return UI_TEXT["dialog_screenshot_title"], summary.app.screenshot_missing
        if key == "readme":
            return UI_TEXT["dialog_readme_title"], summary.app.readme_missing
        if key == "role":
            return UI_TEXT["dialog_role_title"], summary.app.role_attention
        if key == "cloudflare":
            return UI_TEXT["dialog_cloudflare_title"], summary.site.cloudflare_unchecked
        if key == "health":
            return UI_TEXT["dialog_health_title"], summary.site.health_attention
        if key == "site_git":
            return UI_TEXT["dialog_site_git_title"], summary.site.git_uncommitted
        if key == "series_git":
            target = TargetItem(UI_TEXT["source_git"], UI_TEXT["reason_series_git"], DEFAULT_SERIES_ROOT, DEFAULT_SERIES_ROOT.name)
            return UI_TEXT["dialog_series_git_title"], (target,) if summary.git.series_uncommitted or summary.git.series_untracked or summary.git.error else ()
        return UI_TEXT["dialog_select_notice"], ()

    def show_targets(self, key: str) -> None:
        title, targets = self.targets_for_key(key)
        if not targets:
            messagebox.showinfo(title, UI_TEXT["dialog_empty"], parent=self.root)
            return
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("720x420")
        dialog.configure(bg=THEME["bg"])
        apply_window_icon(dialog)
        tk.Label(dialog, text=title, bg=THEME["bg"], fg=THEME["text"], font=(self.font_family, 14, "bold")).pack(anchor="w", padx=16, pady=(14, 4))
        tk.Label(dialog, text=UI_TEXT["dialog_select_notice"], bg=THEME["bg"], fg=THEME["muted"], font=(self.font_family, 9)).pack(anchor="w", padx=16)
        listbox = tk.Listbox(dialog, bg=THEME["input"], fg=THEME["text"], selectbackground=THEME["accent"], selectforeground="#FFFFFF", relief="flat", highlightthickness=1, highlightbackground=THEME["border"], font=(self.font_family, 10))
        listbox.pack(fill="both", expand=True, padx=16, pady=12)
        display_targets: list[TargetItem | None] = []
        for priority in PRIORITY_ORDER:
            group = priority_items(targets, priority)
            if not group:
                continue
            listbox.insert("end", UI_TEXT["priority_header"].format(label=priority_label(priority)))
            display_targets.append(None)
            for target in group:
                listbox.insert("end", target.line())
                display_targets.append(target)
        buttons = tk.Frame(dialog, bg=THEME["bg"])
        buttons.pack(fill="x", padx=16, pady=(0, 14))
        self.make_button(buttons, UI_TEXT["button_open_target"], lambda: self.open_selected_target(key, display_targets, listbox, dialog), primary=True).pack(side="left")
        self.make_button(buttons, UI_TEXT["button_close"], dialog.destroy).pack(side="right")
        listbox.bind("<Double-Button-1>", lambda _event: self.open_selected_target(key, display_targets, listbox, dialog))

    def open_selected_target(self, key: str, targets: list[TargetItem | None], listbox: tk.Listbox, dialog: tk.Toplevel) -> None:
        selection = listbox.curselection()
        if not selection:
            return
        target = targets[selection[0]]
        if target is None:
            return
        dialog.destroy()
        if key == "booth":
            self.open_booth_target(target)
        elif key == "readme":
            readme = target.path / README_NAME
            open_path(readme if readme.exists() else target.path)
        elif key in {"cloudflare", "health"}:
            open_url_or_path(target.url, target.path)
        else:
            open_path(target.path)

    def open_booth_target(self, target: TargetItem) -> None:
        assist_dir = DEFAULT_APPS_ROOT / BOOTH_ASSIST_FOLDER
        exe_path = assist_dir / DIST_DIR_NAME / BOOTH_ASSIST_EXE
        main_path = assist_dir / "main.py"
        try:
            if exe_path.exists():
                subprocess.Popen([str(exe_path), "--app", target.folder_name], cwd=str(assist_dir))
                self.status_var.set(UI_TEXT["booth_assist_notice"].format(folder=target.folder_name))
                return
            if main_path.exists():
                subprocess.Popen([sys.executable, str(main_path), "--app", target.folder_name], cwd=str(assist_dir))
                self.status_var.set(UI_TEXT["booth_assist_notice"].format(folder=target.folder_name))
                return
        except Exception:
            pass
        self.status_var.set(UI_TEXT["booth_assist_fallback"])
        open_path(target.path)

    def open_app_dashboard(self) -> None:
        self.open_dashboard(app_dashboard_dir(), APP_DASHBOARD_EXE)

    def open_web_dashboard(self) -> None:
        self.open_dashboard(web_dashboard_dir(), WEB_DASHBOARD_EXE)

    def open_dashboard(self, folder: Path, exe_name: str) -> None:
        try:
            launch_dashboard(folder, exe_name)
        except FileNotFoundError:
            messagebox.showinfo(UI_TEXT["button_open_failed_title"], UI_TEXT["button_open_failed"].format(path=launch_target(folder, exe_name)), parent=self.root)
        except Exception as exc:
            messagebox.showerror(UI_TEXT["button_open_failed_title"], UI_TEXT["button_open_error"].format(path=folder, error=exc), parent=self.root)


def run_gui() -> int:
    set_windows_app_id()
    root = tk.Tk()
    QpscDashboardApp(root)
    root.mainloop()
    return 0


def run_launch_check() -> int:
    summary = collect_radar_summary()
    print(
        UI_TEXT["launch_check_template"].format(
            apps=summary.app.total,
            booth_urgent=priority_count(summary.app.booth_missing, PRIORITY_URGENT),
            booth_total=len(summary.app.booth_missing),
            release_urgent=priority_count(summary.app.release_missing, PRIORITY_URGENT),
            release_total=len(summary.app.release_missing),
            screenshot_urgent=priority_count(summary.app.screenshot_missing, PRIORITY_URGENT),
            screenshot_total=len(summary.app.screenshot_missing),
            readme_urgent=priority_count(summary.app.readme_missing, PRIORITY_URGENT),
            readme_total=len(summary.app.readme_missing),
            role_urgent=priority_count(summary.app.role_attention, PRIORITY_URGENT),
            role_total=len(summary.app.role_attention),
            sites=summary.site.total,
            cloudflare_urgent=priority_count(summary.site.cloudflare_unchecked, PRIORITY_URGENT),
            cloudflare_total=len(summary.site.cloudflare_unchecked),
            health_urgent=priority_count(summary.site.health_attention, PRIORITY_URGENT),
            health_total=len(summary.site.health_attention),
            site_git_urgent=priority_count(summary.site.git_uncommitted, PRIORITY_URGENT),
            site_git_total=len(summary.site.git_uncommitted),
            series_git=summary.git.series_uncommitted,
        )
    )
    for warning in summary.warnings:
        print(warning)
    if summary.app.error:
        print(summary.app.error)
    if summary.site.error:
        print(summary.site.error)
    if summary.git.error:
        print(summary.git.error)
    return 0 if not summary.warnings and not summary.app.error and not summary.site.error else 1


def main() -> int:
    parser = argparse.ArgumentParser(description=UI_TEXT["window_title"])
    parser.add_argument("--launch-check", action="store_true")
    args = parser.parse_args()
    if args.launch_check:
        return run_launch_check()
    return run_gui()


if __name__ == "__main__":
    raise SystemExit(main())
