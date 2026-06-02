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
WEB_INDEX_FOLDER = "DAKE_Web_Index"
APP_DASHBOARD_EXE = "DakeApp_Dashboard.exe"
WEB_INDEX_EXE = "DakeWeb_Index.exe"
BOOTH_ASSIST_FOLDER = "DAKE_BOOTH_Assist"
BOOTH_ASSIST_EXE = "DakeBOOTH_Assist.exe"
DIST_DIR_NAME = "dist"
README_NAME = "README.md"
RELEASE_BODY_NAME = "release_body.md"
DEFAULT_SERIES_ROOT = Path(os.environ.get("DAKE_SERIES_ROOT", r"C:\Users\yukiz\devlop\DAKE_series"))
DEFAULT_APPS_ROOT = Path(os.environ.get("QPSC_SERIES_APPS_ROOT", str(DEFAULT_SERIES_ROOT / "01_apps")))
WORKER_POLL_MS = 80
MAX_NEXT_ACTIONS = 3
MAX_WEB_SITE_ROWS = 5

UI_TEXT = {
    "window_title": "Quiet Personal Cognitive System",
    "app_title": "Quiet Personal Cognitive System",
    "header_title": "Quiet Personal Cognitive System",
    "header_kicker": "QPCS",
    "header_subtitle": "正本を読み、次にやることだけを静かに表示します。",
    "section_current_title": "現在地",
    "button_reload": "再確認",
    "button_open_app_dashboard": "App詳細",
    "button_open_web_index": "Open Web Index",
    "button_open_target": "選択対象を開く",
    "button_close": "閉じる",
    "button_open_failed_title": "起動できません",
    "button_open_failed": "対象が見つかりません。\n\n{path}",
    "button_open_error": "起動に失敗しました。\n\n{path}\n\n{error}",
    "card_app_title": "アプリ",
    "card_web_sites_title": "Web Sites",
    "card_web_sites_subtitle": "README正本から自動生成",
    "column_folder_name": "フォルダ名",
    "column_site_name": "サイト名",
    "column_updated": "最終更新日時",
    "sites_count": "{count} sites",
    "web_index_opened": "DAKE_Web_Indexを起動しました。",
    "web_index_fallback": "DAKE_Web_Indexフォルダを開きました。",
    "value_unset": "-",
    "card_git_title": "Git",
    "card_next_title": "いま見るもの",
    "section_unresolved_title": "未処理",
    "section_classification_title": "分類",
    "label_app_total": "アプリ総数",
    "label_booth_missing": "BOOTH URL未設定",
    "label_booth_materials_missing": "BOOTH素材不足",
    "label_release_missing": "Release未作成",
    "label_screenshot_missing": "スクショ未作成",
    "label_readme_missing": "README不足",
    "label_role_attention": "分類別 要確認",
    "label_market_count": "市場向け",
    "label_system_count": "QPCS系",
    "label_personal_count": "ユキズ専用",
    "label_frozen_count": "凍結",
    "label_series_uncommitted": "DAKE_series 未commit件数",
    "label_series_untracked": "DAKE_series 未追跡件数",
    "value_waiting": "確認待ち",
    "value_none": "なし",
    "value_yes": "あり",
    "priority_urgent": "先に見る",
    "priority_active": "通常",
    "priority_later": "保留",
    "priority_later_ok": "後でよい",
    "priority_summary": "先に見る {urgent} / 全体 {total}\n通常 {active} / 保留 {later}",
    "priority_brief": "{urgent}/{total}",
    "priority_header": "【{label}】",
    "status_checking": "確認中…",
    "status_ready": "確認完了",
    "status_attention": "未処理あり",
    "status_launch_check_ok": "LAUNCH CHECK OK",
    "last_loaded_waiting": "未確認",
    "last_loaded_value": "確認: {time}",
    "summary_template": "未処理 {total}件 / アプリ {apps}件 / Web {sites} sites",
    "dialog_empty": "対象はありません。",
    "dialog_booth_title": "BOOTH URL未設定アプリ",
    "dialog_booth_materials_title": "BOOTH素材不足アプリ",
    "dialog_release_title": "Release未作成アプリ",
    "dialog_screenshot_title": "スクショ未作成アプリ",
    "dialog_readme_title": "README不足アプリ",
    "dialog_role_title": "分類別 要確認アプリ",
    "dialog_series_git_title": "DAKE_series Git状態",
    "dialog_select_notice": "一覧から対象を選択してください。",
    "dialog_booth_notice": "BOOTH素材が存在し、\nbooth_url が未設定のアプリです。",
    "booth_waiting_items": "登録前チェック: {items}",
    "booth_wait_release": "release_urlなし",
    "booth_wait_screenshot": "スクショ不足",
    "booth_wait_readme": "補足説明不足",
    "next_none": "現時点で明確な候補はありません。",
    "next_booth": "市場向けアプリのBOOTH URL未設定",
    "next_screenshot": "スクショ作成が必要な公開アプリ",
    "next_release": "Release作成が必要な出荷候補",
    "next_readme": "README不足のアプリ",
    "next_role": "分類別の確認が必要なアプリ",
    "next_series_git": "DAKE_seriesのGit確認",
    "next_line": "{index}. {label}（{count}件）",
    "reason_booth_missing": "BOOTH素材あり / booth_url 未設定",
    "reason_booth_materials_missing": "BOOTH登録前の素材が不足",
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
    "target_app_detail": "{detail} / {app_type} / {completion_goal}",
    "app_type_market": "市場向け",
    "app_type_qpcs": "QPCS系",
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
    "reason_series_git": "DAKE_seriesのGit状態",
    "source_app": "App Dashboard",
    "source_web_index": "DAKE_Web_Index",
    "source_git": "DAKE_series",
    "source_error": "{source}: 状態取得に失敗しました",
    "error_missing_main": "main.py が見つかりません",
    "error_import_failed": "外部ノードを読み込めません: {error}",
    "error_missing_function": "外部ノードの取得関数が見つかりません",
    "error_scan_failed": "外部ノードの取得に失敗しました: {error}",
    "booth_assist_notice": "DAKE_BOOTH_Assistを起動しました: {folder}",
    "booth_assist_fallback": "DAKE_BOOTH_Assistが見つからないためフォルダを開きます。",
    "launch_check_template": "LAUNCH CHECK OK: apps={apps} booth_url={booth_urgent}/{booth_total} release={release_urgent}/{release_total} screenshot={screenshot_urgent}/{screenshot_total} readme={readme_urgent}/{readme_total} role={role_urgent}/{role_total} web_sites={sites} series_git={series_git}",
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
APP_TYPE_KEYS = ("market", "qpcs", "personal", "frozen", "archived", "unknown")
COMPLETION_GOAL_KEYS = ("formal_release", "system_ready", "reference_ready", "local_ready", "frozen_closed", "unknown")
LATER_APP_NAMES = {"qpcs", "qpsc", "brainz", "oikawa", "orbit"}
LATER_SITE_STATUS = {"internal", "draft", "archived", "frozen", "deprecated"}
PRIORITY_URGENT = "urgent"
PRIORITY_ACTIVE = "active"
PRIORITY_LATER = "later"
PRIORITY_ORDER = (PRIORITY_URGENT, PRIORITY_ACTIVE, PRIORITY_LATER)
DAKE_META_PATTERN = re.compile(r"##\s*DAKE_META\b.*?```(?:json)?\s*(\{.*?\})\s*```", re.IGNORECASE | re.DOTALL)


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
    booth_materials_missing: tuple[TargetItem, ...] = ()
    release_missing: tuple[TargetItem, ...] = ()
    screenshot_missing: tuple[TargetItem, ...] = ()
    readme_missing: tuple[TargetItem, ...] = ()
    role_attention: tuple[TargetItem, ...] = ()
    error: str = ""


@dataclass(frozen=True)
class WebSiteItem:
    folder_name: str
    site_name: str
    updated_text: str
    folder_path: Path


@dataclass(frozen=True)
class SiteRadar:
    total: int = 0
    items: tuple[WebSiteItem, ...] = ()
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
            + priority_count(self.app.booth_materials_missing, PRIORITY_URGENT)
            + priority_count(self.app.release_missing, PRIORITY_URGENT)
            + priority_count(self.app.screenshot_missing, PRIORITY_URGENT)
            + priority_count(self.app.readme_missing, PRIORITY_URGENT)
            + priority_count(self.app.role_attention, PRIORITY_URGENT)
            + self.git.series_uncommitted
            + self.git.series_untracked
            + len(self.warnings)
        )

    @property
    def action_total(self) -> int:
        return (
            len(self.app.booth_missing)
            + len(self.app.booth_materials_missing)
            + len(self.app.release_missing)
            + len(self.app.screenshot_missing)
            + len(self.app.readme_missing)
            + len(self.app.role_attention)
            + self.git.series_uncommitted
            + self.git.series_untracked
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
    candidates.extend([app_dir().parent / folder_name, DEFAULT_APPS_ROOT / folder_name])
    if folder_name == APP_DASHBOARD_FOLDER:
        candidates.append(app_dir().parent / "codex_staging_dake_dashboard_phase4")
    for candidate in candidates:
        if (candidate / "main.py").exists() or (candidate / DIST_DIR_NAME).exists():
            return candidate
    return candidates[0] if candidates else DEFAULT_APPS_ROOT / folder_name


def app_dashboard_dir() -> Path:
    return dashboard_dir("QPSC_APP_DASHBOARD_DIR", APP_DASHBOARD_FOLDER)


def web_index_dir() -> Path:
    return dashboard_dir("QPSC_WEB_INDEX_DIR", WEB_INDEX_FOLDER)


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


def priority_brief(items: tuple[TargetItem, ...]) -> str:
    total = len(items)
    if total == 0:
        return "0"
    urgent = priority_count(items, PRIORITY_URGENT)
    if urgent == total:
        return str(total)
    return UI_TEXT["priority_brief"].format(urgent=urgent, total=total)


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
    detail_text = UI_TEXT["target_app_detail"].format(
        detail=detail or UI_TEXT[reason_key],
        app_type=app_type_label(app_type_value(record, meta)),
        completion_goal=completion_goal_label(completion_goal_value(record, meta)),
    )
    return TargetItem(
        title=app_display_name(record, meta),
        detail=detail_text,
        path=folder,
        folder_name=safe_text(getattr(record, "folder_name", folder.name)),
        priority=priority,
    )



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


def normalize_app_type_value(value: str) -> str:
    normalized = value.strip().lower()
    if normalized == "system":
        normalized = "qpcs"
    return normalized_choice(normalized, APP_TYPE_KEYS, APP_TYPE_DEFAULT)


def app_type_value(record: object, meta: dict[str, object]) -> str:
    return normalize_app_type_value(meta_or_record_field(record, meta, "app_type"))


def completion_goal_value(record: object, meta: dict[str, object]) -> str:
    return normalized_choice(meta_or_record_field(record, meta, "completion_goal"), COMPLETION_GOAL_KEYS, COMPLETION_GOAL_DEFAULT)


def app_type_label(value: str) -> str:
    key = normalize_app_type_value(value)
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


def booth_url_waiting_items(record: object, meta: dict[str, object], has_readme_issue: bool) -> tuple[str, ...]:
    items: list[str] = []
    if not app_release_url(record, meta):
        items.append(UI_TEXT["booth_wait_release"])
    if app_screenshot_missing(record, meta):
        items.append(UI_TEXT["booth_wait_screenshot"])
    if has_readme_issue:
        items.append(UI_TEXT["booth_wait_readme"])
    return tuple(items)


def booth_url_detail(record: object, meta: dict[str, object], has_readme_issue: bool) -> str:
    waiting_items = booth_url_waiting_items(record, meta, has_readme_issue)
    if not waiting_items:
        return UI_TEXT["reason_booth_missing"]
    return f"{UI_TEXT['reason_booth_missing']} / {UI_TEXT['booth_waiting_items'].format(items=', '.join(waiting_items))}"


def classify_booth_url_priority(record: object, meta: dict[str, object], has_readme_issue: bool = False) -> str:
    if not app_booth_url_missing(record, meta):
        return PRIORITY_LATER
    status = meta_status(record, meta)
    visible = meta_show_on_site(meta) or meta_show_in_launcher(record, meta)
    ready_to_register = (
        status == "available"
        and visible
        and bool(app_release_url(record, meta))
        and not app_screenshot_missing(record, meta)
        and not has_readme_issue
    )
    return PRIORITY_URGENT if ready_to_register else PRIORITY_ACTIVE


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


def app_booth_materials_complete(record: object) -> bool:
    has_booth_product, _has_ready_product, has_booth_ready, has_booth_thumbnail = app_booth_material_flags(record)
    return has_booth_product and has_booth_ready and has_booth_thumbnail


def app_booth_url_missing(record: object, meta: dict[str, object]) -> bool:
    if not is_market_formal_app(record, meta):
        return False
    if app_is_later_market_target(record, meta):
        return False
    if safe_text(meta.get("booth_url", "")):
        return False
    return app_booth_materials_complete(record)


def app_booth_materials_missing(record: object, meta: dict[str, object]) -> bool:
    if not is_market_formal_app(record, meta):
        return False
    if app_is_later_market_target(record, meta):
        return False
    return not app_booth_materials_complete(record)


def app_booth_missing(record: object, meta: dict[str, object]) -> bool:
    return app_booth_url_missing(record, meta)


def app_role_family(record: object, meta: dict[str, object]) -> str:
    app_type = app_type_value(record, meta)
    if app_type in {"market", "qpcs", "personal", "frozen"}:
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
        "qpcs": {"system_ready", "reference_ready"},
        "personal": {"local_ready"},
        "frozen": {"frozen_closed"},
        "archived": {"frozen_closed", "reference_ready"},
    }
    if goal not in expected_goals.get(app_type, {"formal_release"}):
        return f"{UI_TEXT['reason_role_goal_mismatch']} ({' / '.join(detail_bits)})", PRIORITY_ACTIVE

    missing: list[str] = []
    if app_type == "qpcs" and goal == "system_ready":
        if not (folder / "build.bat").exists():
            missing.append("build.bat")
        main_text = read_optional_text(folder / "main.py")
        if "--launch-check" not in readme_text and "--launch-check" not in main_text:
            missing.append("--launch-check")
        if "system_ready" not in readme_text and "Positioning" not in readme_text:
            missing.append("role docs")
        if missing:
            return f"{UI_TEXT['reason_system_ready_attention']} {UI_TEXT['reason_role_missing_items'].format(items=', '.join(missing))}", PRIORITY_ACTIVE
    elif app_type == "qpcs" and goal == "reference_ready":
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
        booth_materials_missing: list[TargetItem] = []
        release_missing: list[TargetItem] = []
        screenshot_missing: list[TargetItem] = []
        readme_missing: list[TargetItem] = []
        role_attention: list[TargetItem] = []
        role_counts = {"market": 0, "qpcs": 0, "personal": 0, "frozen": 0}
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
                if app_booth_url_missing(record, meta):
                    booth_missing.append(
                        target_for_app(
                            record,
                            meta,
                            "reason_booth_missing",
                            booth_url_detail(record, meta, bool(reason)),
                            classify_booth_url_priority(record, meta, bool(reason)),
                        )
                    )
                if app_booth_materials_missing(record, meta):
                    booth_materials_missing.append(target_for_app(record, meta, "reason_booth_materials_missing", priority=PRIORITY_ACTIVE))
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
                system_count=role_counts["qpcs"],
                personal_count=role_counts["personal"],
                frozen_count=role_counts["frozen"],
                booth_missing=tuple(booth_missing),
                booth_materials_missing=tuple(booth_materials_missing),
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



def first_safe_text(*values: object) -> str:
    for value in values:
        text = safe_text(value)
        if text:
            return text
    return ""


def web_site_item_from_record(record: object) -> WebSiteItem:
    folder_path = Path(getattr(record, "folder_path", ""))
    folder_name = safe_text(getattr(record, "folder_name", "")) or folder_path.name
    site_name = safe_text(getattr(record, "site_name", "")) or UI_TEXT["value_unset"]
    updated_getter = getattr(record, "updated_text", None)
    updated_text = safe_text(updated_getter()) if callable(updated_getter) else safe_text(getattr(record, "last_updated", ""))
    return WebSiteItem(
        folder_name=folder_name or UI_TEXT["value_unset"],
        site_name=site_name,
        updated_text=updated_text or UI_TEXT["value_unset"],
        folder_path=folder_path,
    )


def collect_site_radar() -> tuple[SiteRadar, tuple[str, ...]]:
    folder = web_index_dir()
    try:
        module = load_dashboard_module("web_index", folder)
        if not hasattr(module, "scan_sites"):
            raise AttributeError(UI_TEXT["error_missing_function"])
        dev_root = getattr(module, "DEFAULT_DEV_ROOT", DEFAULT_SERIES_ROOT.parent)
        records = list(module.scan_sites(dev_root))
        records.sort(key=lambda record: (safe_text(getattr(record, "folder_name", "")).lower(), safe_text(getattr(record, "site_name", "")).lower()))
        items = tuple(web_site_item_from_record(record) for record in records[:MAX_WEB_SITE_ROWS])
        return SiteRadar(total=len(records), items=items), ()
    except Exception as exc:
        warning = UI_TEXT["source_error"].format(source=UI_TEXT["source_web_index"])
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
            "booth_materials": tk.StringVar(value=UI_TEXT["value_waiting"]),
            "release": tk.StringVar(value=UI_TEXT["value_waiting"]),
            "screenshot": tk.StringVar(value=UI_TEXT["value_waiting"]),
            "readme": tk.StringVar(value=UI_TEXT["value_waiting"]),
            "role": tk.StringVar(value=UI_TEXT["value_waiting"]),
            "market": tk.StringVar(value=UI_TEXT["value_waiting"]),
            "system": tk.StringVar(value=UI_TEXT["value_waiting"]),
            "personal": tk.StringVar(value=UI_TEXT["value_waiting"]),
            "frozen": tk.StringVar(value=UI_TEXT["value_waiting"]),
        }
        self.web_site_rows: list[tuple[tk.StringVar, tk.StringVar, tk.StringVar]] = []
        self.web_sites_total_var = tk.StringVar(value=UI_TEXT["value_waiting"])
        self.git_vars = {
            "series_uncommitted": tk.StringVar(value=UI_TEXT["value_waiting"]),
            "series_untracked": tk.StringVar(value=UI_TEXT["value_waiting"]),
        }
        self.summary_var = tk.StringVar(value=UI_TEXT["last_loaded_waiting"])
        self.status_var = tk.StringVar(value=UI_TEXT["status_checking"])
        self.next_frame: tk.Frame | None = None
        self.configure_root()
        self.build_ui()
        self.root.after(120, self.refresh)

    def configure_root(self) -> None:
        self.root.title(UI_TEXT["window_title"])
        self.root.geometry("1240x840")
        self.root.minsize(1120, 760)
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
        self.make_button(actions, UI_TEXT["button_open_web_index"], self.open_web_index).pack(side="left")

        tk.Label(outer, text=UI_TEXT["section_current_title"], bg=THEME["bg"], fg=THEME["text"], font=(self.font_family, 13, "bold")).pack(anchor="w", pady=(24, 8))
        cards = tk.Frame(outer, bg=THEME["bg"])
        cards.pack(fill="x", pady=(0, 18))
        cards.grid_columnconfigure(0, weight=1, uniform="cards")
        cards.grid_columnconfigure(1, weight=1, uniform="cards")
        self.build_app_card(cards).grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        self.build_web_sites_card(cards).grid(row=0, column=1, sticky="nsew", padx=(10, 0))

        lower = tk.Frame(outer, bg=THEME["bg"])
        lower.pack(fill="both", expand=True)
        lower.grid_columnconfigure(0, weight=3, uniform="lower")
        lower.grid_columnconfigure(1, weight=2, uniform="lower")
        lower.grid_rowconfigure(0, weight=1, minsize=220)
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
        body.pack(fill="both", expand=True, padx=16, pady=(2, 16))
        self.metric_row(body, UI_TEXT["label_app_total"], self.app_vars["total"], None).pack(fill="x", pady=(0, 10))

        sections = tk.Frame(body, bg=THEME["panel"])
        sections.pack(fill="both", expand=True)
        sections.grid_columnconfigure(0, weight=3, uniform="app_sections")
        sections.grid_columnconfigure(1, weight=2, uniform="app_sections")

        unresolved = tk.Frame(sections, bg=THEME["panel"])
        unresolved.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        self.section_title(unresolved, UI_TEXT["section_unresolved_title"])
        self.metric_grid(
            unresolved,
            (
                (UI_TEXT["label_booth_missing"], self.app_vars["booth"], lambda: self.show_targets("booth")),
                (UI_TEXT["label_booth_materials_missing"], self.app_vars["booth_materials"], lambda: self.show_targets("booth_materials")),
                (UI_TEXT["label_release_missing"], self.app_vars["release"], lambda: self.show_targets("release")),
                (UI_TEXT["label_screenshot_missing"], self.app_vars["screenshot"], lambda: self.show_targets("screenshot")),
                (UI_TEXT["label_readme_missing"], self.app_vars["readme"], lambda: self.show_targets("readme")),
                (UI_TEXT["label_role_attention"], self.app_vars["role"], lambda: self.show_targets("role")),
            ),
        )

        classification = tk.Frame(sections, bg=THEME["panel"])
        classification.grid(row=0, column=1, sticky="nsew")
        self.section_title(classification, UI_TEXT["section_classification_title"])
        self.metric_grid(
            classification,
            (
                (UI_TEXT["label_market_count"], self.app_vars["market"], None),
                (UI_TEXT["label_system_count"], self.app_vars["system"], None),
                (UI_TEXT["label_personal_count"], self.app_vars["personal"], None),
                (UI_TEXT["label_frozen_count"], self.app_vars["frozen"], None),
            ),
        )
        return frame

    def build_web_sites_card(self, parent: tk.Misc) -> tk.Frame:
        frame = self.make_card(parent)
        self.card_title(frame, UI_TEXT["card_web_sites_title"])
        body = tk.Frame(frame, bg=THEME["panel"])
        body.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        tk.Label(body, text=UI_TEXT["card_web_sites_subtitle"], bg=THEME["panel"], fg=THEME["muted"], font=(self.font_family, 9)).pack(anchor="w", pady=(0, 8))

        table = tk.Frame(body, bg=THEME["panel"])
        table.pack(fill="both", expand=True)
        for column, weight in enumerate((3, 3, 2)):
            table.grid_columnconfigure(column, weight=weight, uniform="web_sites")
        headers = (UI_TEXT["column_folder_name"], UI_TEXT["column_site_name"], UI_TEXT["column_updated"])
        for column, label in enumerate(headers):
            tk.Label(table, text=label, bg=THEME["panel"], fg=THEME["quiet"], anchor="w", font=(self.font_family, 8, "bold")).grid(row=0, column=column, sticky="ew", padx=(0, 8 if column < 2 else 0), pady=(0, 4))

        for row_index in range(MAX_WEB_SITE_ROWS):
            row_vars = (
                tk.StringVar(value=""),
                tk.StringVar(value=""),
                tk.StringVar(value=""),
            )
            self.web_site_rows.append(row_vars)
            for column, value_var in enumerate(row_vars):
                cell = tk.Label(table, textvariable=value_var, bg=THEME["panel_alt"], fg=THEME["text"], anchor="w", font=(self.font_family, 9), padx=8, pady=5)
                cell.grid(row=row_index + 1, column=column, sticky="ew", padx=(0, 8 if column < 2 else 0), pady=(0, 4))
                cell.bind("<Double-Button-1>", lambda _event: self.open_web_index())

        footer = tk.Frame(body, bg=THEME["panel"])
        footer.pack(fill="x", pady=(8, 0))
        tk.Label(footer, textvariable=self.web_sites_total_var, bg=THEME["panel"], fg=THEME["muted"], font=(self.font_family, 10, "bold")).pack(side="left")
        self.make_button(footer, UI_TEXT["button_open_web_index"], self.open_web_index, primary=True).pack(side="right")
        return frame

    def build_git_card(self, parent: tk.Misc) -> tk.Frame:
        frame = self.make_card(parent)
        self.card_title(frame, UI_TEXT["card_git_title"])
        body = tk.Frame(frame, bg=THEME["panel"])
        body.pack(fill="x", padx=16, pady=(2, 16))
        self.metric_row(body, UI_TEXT["label_series_uncommitted"], self.git_vars["series_uncommitted"], lambda: self.show_targets("series_git")).pack(fill="x", pady=(0, 8))
        self.metric_row(body, UI_TEXT["label_series_untracked"], self.git_vars["series_untracked"], lambda: self.show_targets("series_git")).pack(fill="x")
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

    def section_title(self, parent: tk.Misc, text: str) -> None:
        tk.Label(parent, text=text, bg=THEME["panel"], fg=THEME["text"], font=(self.font_family, 9, "bold")).pack(anchor="w", pady=(0, 6))

    def metric_grid(self, parent: tk.Misc, items: tuple[tuple[str, tk.StringVar, object], ...]) -> None:
        grid = tk.Frame(parent, bg=THEME["panel"])
        grid.pack(fill="both", expand=True)
        for column in range(2):
            grid.grid_columnconfigure(column, weight=1, uniform="metric_grid")
        for index, (label, value_var, command) in enumerate(items):
            self.compact_metric_row(grid, label, value_var, command).grid(
                row=index // 2,
                column=index % 2,
                sticky="ew",
                padx=(0 if index % 2 == 0 else 8, 0),
                pady=(0, 7),
            )

    def compact_metric_row(self, parent: tk.Misc, label: str, value_var: tk.StringVar, command) -> tk.Frame:
        frame = tk.Frame(parent, bg=THEME["panel_alt"], highlightthickness=1, highlightbackground=THEME["shadow"])
        cursor = "hand2" if command else ""
        label_widget = tk.Label(frame, text=label, bg=THEME["panel_alt"], fg=THEME["muted"], font=(self.font_family, 8), cursor=cursor)
        label_widget.pack(side="left", padx=10, pady=7)
        value_widget = tk.Label(frame, textvariable=value_var, bg=THEME["panel_alt"], fg=THEME["text"], justify="right", font=(self.font_family, 11, "bold"), cursor=cursor)
        value_widget.pack(side="right", padx=10, pady=6)
        if command:
            frame.configure(cursor="hand2")
            for widget in (frame, label_widget, value_widget):
                widget.bind("<Button-1>", lambda _event, action=command: action())
        return frame

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
        self.app_vars["booth"].set(priority_brief(summary.app.booth_missing))
        self.app_vars["booth_materials"].set(priority_brief(summary.app.booth_materials_missing))
        self.app_vars["release"].set(priority_brief(summary.app.release_missing))
        self.app_vars["screenshot"].set(priority_brief(summary.app.screenshot_missing))
        self.app_vars["readme"].set(priority_brief(summary.app.readme_missing))
        self.app_vars["role"].set(str(len(summary.app.role_attention)))
        self.app_vars["market"].set(str(summary.app.market_count))
        self.app_vars["system"].set(str(summary.app.system_count))
        self.app_vars["personal"].set(str(summary.app.personal_count))
        self.app_vars["frozen"].set(str(summary.app.frozen_count))
        self.web_sites_total_var.set(UI_TEXT["sites_count"].format(count=summary.site.total))
        for index, row_vars in enumerate(self.web_site_rows):
            item = summary.site.items[index] if index < len(summary.site.items) else None
            row_vars[0].set(item.folder_name if item else "")
            row_vars[1].set(item.site_name if item else "")
            row_vars[2].set(item.updated_text if item else "")
        self.git_vars["series_uncommitted"].set(str(summary.git.series_uncommitted))
        self.git_vars["series_untracked"].set(str(summary.git.series_untracked))
        self.summary_var.set(UI_TEXT["summary_template"].format(total=summary.action_total, apps=summary.app.total, sites=summary.site.total))
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
            ("next_role", len(summary.app.role_attention), "role"),
            ("next_booth", priority_count(summary.app.booth_missing, PRIORITY_URGENT), "booth"),
            ("next_series_git", summary.git.series_uncommitted + summary.git.series_untracked, "series_git"),
        ]
        return [item for item in candidates if item[1]][:MAX_NEXT_ACTIONS]

    def targets_for_key(self, key: str) -> tuple[str, tuple[TargetItem, ...]]:
        summary = self.summary
        if summary is None:
            return UI_TEXT["dialog_select_notice"], ()
        if key == "booth":
            return UI_TEXT["dialog_booth_title"], summary.app.booth_missing
        if key == "booth_materials":
            return UI_TEXT["dialog_booth_materials_title"], summary.app.booth_materials_missing
        if key == "release":
            return UI_TEXT["dialog_release_title"], summary.app.release_missing
        if key == "screenshot":
            return UI_TEXT["dialog_screenshot_title"], summary.app.screenshot_missing
        if key == "readme":
            return UI_TEXT["dialog_readme_title"], summary.app.readme_missing
        if key == "role":
            return UI_TEXT["dialog_role_title"], summary.app.role_attention
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
        notice = UI_TEXT["dialog_booth_notice"] if key == "booth" else UI_TEXT["dialog_select_notice"]
        tk.Label(dialog, text=notice, bg=THEME["bg"], fg=THEME["muted"], justify="left", font=(self.font_family, 9)).pack(anchor="w", padx=16)
        listbox = tk.Listbox(dialog, bg=THEME["input"], fg=THEME["text"], selectbackground=THEME["accent"], selectforeground="#FFFFFF", relief="flat", highlightthickness=1, highlightbackground=THEME["border"], font=(self.font_family, 10))
        listbox.pack(fill="both", expand=True, padx=16, pady=12)
        display_targets: list[TargetItem | None] = []
        if key == "booth":
            groups = (
                (UI_TEXT["priority_urgent"], priority_items(targets, PRIORITY_URGENT), True),
                (UI_TEXT["priority_later_ok"], tuple(target for target in targets if target.priority != PRIORITY_URGENT), True),
            )
        else:
            groups = tuple((priority_label(priority), priority_items(targets, priority), False) for priority in PRIORITY_ORDER)
        for label, group, show_empty in groups:
            if not group and not show_empty:
                continue
            listbox.insert("end", UI_TEXT["priority_header"].format(label=label))
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

    def open_web_index(self) -> None:
        folder = web_index_dir()
        exe_path = folder / DIST_DIR_NAME / WEB_INDEX_EXE
        try:
            if exe_path.exists():
                subprocess.Popen([str(exe_path)], cwd=str(folder))
                self.status_var.set(UI_TEXT["web_index_opened"])
                return
            if folder.exists():
                open_path(folder)
                self.status_var.set(UI_TEXT["web_index_fallback"])
                return
            raise FileNotFoundError(str(folder))
        except FileNotFoundError:
            messagebox.showinfo(UI_TEXT["button_open_failed_title"], UI_TEXT["button_open_failed"].format(path=folder), parent=self.root)
        except Exception as exc:
            messagebox.showerror(UI_TEXT["button_open_failed_title"], UI_TEXT["button_open_error"].format(path=folder, error=exc), parent=self.root)

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
            series_git=summary.git.series_uncommitted + summary.git.series_untracked,
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
