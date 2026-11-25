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


APP_NAME = "DAKE Web Dashboard"
WINDOW_TITLE = "DAKE Web Dashboard"
APP_ID = "dake.web.dashboard"
DEV_ROOT = Path(os.environ.get("DAKE_WEB_DASHBOARD_ROOT", r"C:\Users\yukiz\devlop"))
README_NAME = "README.md"
WRANGLER_NAME = "wrangler.toml"
PACKAGE_NAME = "package.json"
ROUTES_NAME = "_routes.json"
PUBLIC_ROUTES = Path("public") / ROUTES_NAME
FUNCTIONS_DIR = "functions"
PUBLIC_DIR = "public"
GIT_DIR = ".git"
GITIGNORE_NAME = ".gitignore"
AUTO_RELOAD_MS = 30000
WORKER_POLL_MS = 80
WATCH_DEBOUNCE_MS = 700
NOTIFICATION_HIDE_MS = 3600
GIT_TIMEOUT_SECONDS = 3.0
MAX_SCAN_DEPTH = 2
MAX_CODE_SCAN_FILES = 220
MAX_CODE_FILE_BYTES = 700_000

UI_TEXT = {
    "header_title": "DAKE Web Dashboard",
    "header_subtitle": "サイト群の現在地を見る",
    "last_loaded_waiting": "最終読込: 未読込",
    "last_loaded_value": "最終読込: {time}",
    "watch_status_on": "監視状態: 監視中",
    "watch_status_off": "監視状態: 停止中",
    "watch_status_polling": "監視状態: watchdogなし / 30秒更新",
    "watch_status_error": "監視状態: 起動不可",
    "auto_status_on": "自動更新: ON",
    "auto_status_off": "自動更新: OFF",
    "button_reload": "再読み込み",
    "button_auto_start": "自動更新開始",
    "button_auto_stop": "自動更新停止",
    "button_watch_start": "監視開始",
    "button_watch_stop": "監視停止",
    "button_open_folder": "サイトフォルダを開く",
    "button_open_readme": "README.md を開く",
    "button_open_production": "production_url を開く",
    "button_open_health": "health_url を開く",
    "button_open_github": "GitHub を開く",
    "button_open_cloudflare": "Cloudflare を開く",
    "button_url_missing": "URL なし",
    "summary_total": "総サイト数",
    "summary_normal": "正常",
    "summary_needs_review": "要確認",
    "summary_api_review": "API確認",
    "summary_deploy_review": "デプロイ確認",
    "summary_dirty": "未コミットあり",
    "qpsc_title": "QPSC通知カード",
    "qpsc_subtitle": "正本と構成差分を監視中",
    "qpsc_new_sites": "新規サイト検出",
    "qpsc_readme_missing": "README未整備",
    "qpsc_meta_missing": "DAKE_WEB_META未整備",
    "qpsc_api_review": "API確認が必要",
    "qpsc_cloudflare_review": "Cloudflare確認が必要",
    "qpsc_git_dirty": "Git未コミットあり",
    "qpsc_git_error": "Git取得不可",
    "qpsc_candidate_components": "誤検出候補",
    "qpsc_hint_meta": "README正本化が必要",
    "qpsc_hint_api": "Functions / OpenAI / health を確認",
    "qpsc_hint_git": "repo外またはGit状態取得失敗",
    "qpsc_hint_candidate": "public / functions / sitemap など",
    "git_card_title": "Git状態カード",
    "git_dirty_sites": "未コミット変更ありサイト",
    "git_untracked_sites": "未追跡ありサイト",
    "git_ahead_sites": "push待ち疑い",
    "git_error_sites": "Git取得不可",
    "filter_all": "全部",
    "filter_normal": "正常",
    "filter_needs_review": "要確認",
    "filter_api_review": "API確認",
    "filter_deploy_review": "デプロイ確認",
    "filter_dirty": "未コミット",
    "filter_candidate": "候補",
    "filter_internal": "内部 / 凍結",
    "search_label": "検索",
    "search_placeholder": "フォルダ名・表示名・domain・Cloudflare・status を検索",
    "list_title": "サイト一覧",
    "detail_title": "詳細ペイン",
    "detail_empty": "サイトを選択すると、正本情報と不足項目を確認できます。",
    "next_title": "次にやる候補",
    "count_line": "{visible} / {total} 件を表示",
    "column_status": "状態",
    "column_folder": "フォルダ名",
    "column_display": "表示名",
    "column_domain": "ドメイン",
    "column_cloudflare": "Cloudflare Project",
    "column_git": "Git状態",
    "column_api": "API",
    "column_functions": "Functions",
    "column_reason": "判定理由",
    "column_updated": "最終更新",
    "class_normal": "正常",
    "class_needs_review": "要確認",
    "class_api_review": "API確認",
    "class_deploy_review": "デプロイ確認",
    "class_candidate": "候補",
    "class_internal": "内部 / 凍結",
    "status_loading": "読み込み中",
    "status_ready": "正本を読み込みました",
    "status_error": "確認が必要です",
    "status_launch_check_ok": "LAUNCH CHECK OK",
    "value_yes": "あり",
    "value_no": "なし",
    "value_unknown": "不明",
    "value_unset": "未設定",
    "value_ok": "OK",
    "value_ng": "NG",
    "value_suspect": "疑いあり",
    "value_clean": "clean",
    "value_get_failed": "取得不可",
    "value_not_repo": "Gitなし",
    "value_static": "静的",
    "value_functions": "Functions",
    "value_api_review": "API要確認",
    "value_direct_key": "直書き疑い",
    "value_front_direct": "直叩き疑い",
    "value_env_design": "env経由",
    "value_watchdog_missing": "watchdog未導入",
    "label_folder_path": "フォルダパス",
    "label_readme_path": "README.md",
    "label_domain": "domain",
    "label_cloudflare_project": "cloudflare_project",
    "label_cloudflare_url": "Cloudflare URL",
    "label_readme": "README状態",
    "label_meta": "DAKE_WEB_META状態",
    "label_git": "Git状態",
    "label_cloudflare": "Cloudflare構成",
    "label_api": "API構成",
    "label_openai_safety": "OpenAI API安全性チェック",
    "label_detection_basis": "判定根拠",
    "label_status_reasons": "状態理由",
    "label_api_reasons": "API理由",
    "label_git_reasons": "Git理由",
    "label_missing": "不足項目",
    "label_next": "次に必要そうな作業",
    "label_next_reasons": "次にやる理由",
    "label_detection_score": "検出スコア",
    "label_candidate_type": "candidate_type",
    "label_repo_url": "GitHub URL",
    "label_production_url": "production_url",
    "label_health_url": "health_url",
    "label_site_type": "site_type",
    "label_status": "status",
    "label_show_dashboard": "show_on_dashboard",
    "label_has_readme": "README.md",
    "label_has_meta": "DAKE_WEB_META",
    "label_has_git": "Gitリポジトリ",
    "label_has_package": "package.json",
    "label_has_wrangler": "wrangler.toml",
    "label_has_public": "public/",
    "label_has_functions": "functions/",
    "label_has_routes": "_routes.json",
    "label_has_gitignore": ".gitignore",
    "label_node_modules_ignored": "node_modules Git対象外推定",
    "label_pages_output": "pages_build_output_dir",
    "label_pages_like": "Cloudflare Pagesらしさ",
    "label_functions_api": "functions/api/",
    "label_health_file": "/api/health 相当",
    "label_routes_api": "_routes.json /api/*",
    "label_openai_env": "OPENAI_API_KEY参照",
    "label_openai_key": "sk-直書き疑い",
    "label_openai_front": "フロント直叩き疑い",
    "label_openai_functions": "functions側OpenAI接続",
    "label_openai_env_design": "env経由設計",
    "label_branch": "ブランチ",
    "label_commit": "最新コミット",
    "label_dirty_count": "未コミット変更数",
    "label_untracked_count": "未追跡ファイル数",
    "label_ahead": "push待ち疑い",
    "label_behind": "pull待ち疑い",
    "missing_none": "不足は見つかっていません。",
    "missing_readme": "README.md がありません。",
    "missing_meta": "DAKE_WEB_META が未整備です。",
    "missing_git": "Gitリポジトリではありません。",
    "missing_git_error": "Git状態を取得できません。",
    "missing_package": "package.json がありません。",
    "missing_wrangler": "wrangler.toml がありません。",
    "missing_domain": "domain / production_url が不明です。",
    "missing_cloudflare_project": "cloudflare_project が不明です。",
    "missing_routes": "_routes.json がありません。",
    "missing_health": "/api/health 相当がありません。",
    "missing_api_key": "APIキー直書き疑いがあります。",
    "missing_front_direct": "フロントからOpenAI API直叩き疑いがあります。",
    "missing_git_dirty": "Git未コミット変更があります。",
    "next_fix_key": "APIキー直書き疑いを確認し、env参照へ移す",
    "next_fix_front_direct": "フロント直叩きをFunctions経由へ移す",
    "next_add_readme": "README.md を整備する",
    "next_add_meta": "README.md に DAKE_WEB_META を追加する",
    "next_git_commit": "Git未コミット変更の内容を確認する",
    "next_add_wrangler": "Cloudflare Pages構成として wrangler.toml を確認する",
    "next_add_health": "Functionsありサイトに /api/health を用意する",
    "next_add_domain": "production_url / domain を正本へ記録する",
    "next_no_action": "現時点で明確な追加作業はありません。",
    "reason_none": "明確な追加理由はありません。",
    "reason_git_exists": ".git あり",
    "reason_readme_exists": "README.md あり",
    "reason_wrangler_exists": "wrangler.toml あり",
    "reason_package_exists": "package.json あり",
    "reason_public_exists": "public/ あり",
    "reason_functions_exists": "functions/ あり",
    "reason_meta_exists": "DAKE_WEB_META あり",
    "reason_package_public": "package.json + public/ あり",
    "reason_wrangler_functions": "wrangler.toml + functions/ あり",
    "reason_name_site": "フォルダ名がサイト系",
    "reason_component_public": "public単体のためサイト本体ではない可能性",
    "reason_component_functions": "functions単体のためサイト本体ではない可能性",
    "reason_component_sitemap": "sitemap単体のためサイト本体ではない可能性",
    "reason_parent_site_component": "親フォルダがサイトの構成要素",
    "reason_low_signal": ".git / README.md / wrangler.toml がありません",
    "reason_status_meta_missing": "DAKE_WEB_META が未整備",
    "reason_status_readme_missing": "README.md がありません",
    "reason_status_git_error": "Git状態を取得できません",
    "reason_status_production_missing": "production_url が未設定",
    "reason_status_cloudflare_missing": "cloudflare_project が未設定",
    "reason_status_functions_health": "functions あり / health 未確認",
    "reason_status_api_review": "OpenAI API確認が必要",
    "reason_status_internal": "内部 / 凍結扱い",
    "reason_status_clean": "主要な不足は見つかっていません",
    "reason_api_key": "APIキー直書き疑い",
    "reason_api_front_direct": "フロントOpenAI直叩き疑い",
    "reason_api_env": "OPENAI_API_KEY参照あり",
    "reason_api_functions": "functions側OpenAI接続あり",
    "reason_api_health_missing": "health_url または /api/health 未確認",
    "reason_api_routes_missing": "_routes.json /api/* 未確認",
    "reason_api_none": "API確認理由なし",
    "reason_git_no_repo": "Gitリポジトリではありません",
    "reason_git_error": "Git取得不可",
    "reason_git_dirty": "未コミットあり",
    "reason_git_untracked": "未追跡あり",
    "reason_git_ahead": "push待ち疑い",
    "reason_git_behind": "pull待ち疑い",
    "reason_git_clean": "Git clean",
    "candidate_type_site": "site",
    "candidate_type_candidate": "candidate",
    "candidate_type_ignored": "ignored_component",
    "dialog_error_title": "エラー",
    "dialog_notice_title": "確認",
    "dialog_open_failed": "開けませんでした。\n\n{path}\n\n{error}",
    "dialog_missing_path": "対象が見つかりません。\n\n{path}",
    "dialog_url_missing": "URL が設定されていません。",
    "dialog_url_invalid": "安全に開ける http/https URL ではありません。",
    "notification_reloaded": "{folder} を再読込しました",
    "footer_note": "内部用。一般公開・BOOTH登録・GitHub Release作成・dakeapp.com反映は行いません。",
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
    "normal": ("#152A25", THEME["success"]),
    "needs_review": ("#301B34", THEME["review"]),
    "api_review": ("#32201B", THEME["warning"]),
    "deploy_review": ("#1D213B", THEME["purple"]),
    "candidate": ("#202736", THEME["muted"]),
    "internal": ("#1B2333", THEME["muted"]),
}

FONT_CANDIDATES = ["BIZ UDPGothic", "Yu Gothic UI", "Meiryo", "MS Gothic"]
FILTER_KEYS = ("all", "normal", "needs_review", "api_review", "deploy_review", "dirty", "candidate", "internal")
COMPONENT_DIR_NAMES = {"public", "functions", "sitemap"}
EXCLUDED_DIR_NAMES = {
    ".git",
    ".wrangler",
    ".next",
    ".venv",
    "__pycache__",
    "build",
    "dake_series",
    "dist",
    "node_modules",
    "venv",
}
CODE_SUFFIXES = {".js", ".mjs", ".cjs", ".jsx", ".ts", ".tsx", ".html", ".htm", ".py", ".json", ".toml"}
FRONTEND_SUFFIXES = {".js", ".mjs", ".cjs", ".jsx", ".ts", ".tsx", ".html", ".htm"}
WATCHED_FILE_NAMES = {README_NAME, WRANGLER_NAME, PACKAGE_NAME, ROUTES_NAME}
WATCHED_DIR_NAMES = {FUNCTIONS_DIR, PUBLIC_DIR}

DAKE_WEB_META_PATTERN = re.compile(
    r"##\s*DAKE_WEB_META\b.*?```(?:json)?\s*(\{.*?\})\s*```",
    re.IGNORECASE | re.DOTALL,
)
URL_PATTERN = re.compile(r"https?://[^\s)>\]\"']+")
GITHUB_PATTERN = re.compile(r"https?://github\.com/[^\s)>\]\"']+", re.IGNORECASE)
SK_KEY_PATTERN = re.compile(r"\bsk-[A-Za-z0-9_\-]{12,}")
SECRET_URL_PATTERN = re.compile(r"(sk-[A-Za-z0-9_\-]{8,}|[?&](?:api[_-]?key|apikey|token|secret|key)=)", re.IGNORECASE)
WRANGLER_NAME_PATTERN = re.compile(r"(?m)^\s*name\s*=\s*[\"']([^\"']+)[\"']")
PAGES_OUTPUT_PATTERN = re.compile(r"(?m)^\s*pages_build_output_dir\s*=")


@dataclass(frozen=True)
class GitInfo:
    is_repo: bool
    branch: str = ""
    commit: str = ""
    dirty_count: int = 0
    untracked_count: int = 0
    ahead: int = 0
    behind: int = 0
    error: str = ""

    @property
    def has_dirty(self) -> bool:
        return self.dirty_count > 0

    @property
    def has_untracked(self) -> bool:
        return self.untracked_count > 0


@dataclass(frozen=True)
class FileChecks:
    has_readme: bool
    has_meta: bool
    meta_error: str
    has_git: bool
    has_package: bool
    has_wrangler: bool
    has_public: bool
    has_functions: bool
    has_routes: bool
    has_gitignore: bool
    node_modules_ignored: bool


@dataclass(frozen=True)
class CloudflareInfo:
    has_pages_output: bool
    likely_pages: bool
    has_functions_api: bool
    has_health_file: bool
    has_routes_api: bool
    route_file: str


@dataclass(frozen=True)
class ApiInfo:
    has_openai_env_ref: bool
    has_hardcoded_key_suspect: bool
    has_frontend_openai_direct: bool
    has_functions_openai: bool
    has_env_design: bool


@dataclass(frozen=True)
class CandidateInfo:
    score: int
    candidate_type: str
    reasons: tuple[str, ...]


@dataclass(frozen=True)
class SiteRecord:
    folder_name: str
    folder_path: Path
    display_name: str
    domain: str
    cloudflare_project: str
    site_type: str
    status_value: str
    show_on_dashboard: bool
    production_url: str
    health_url: str
    github_url: str
    cloudflare_url: str
    files: FileChecks
    cloudflare: CloudflareInfo
    api: ApiInfo
    git: GitInfo
    class_key: str
    detection_score: int
    detection_reasons: tuple[str, ...]
    status_reasons: tuple[str, ...]
    api_reasons: tuple[str, ...]
    git_reasons: tuple[str, ...]
    next_action_reasons: tuple[str, ...]
    candidate_type: str
    missing_items: tuple[str, ...]
    next_items: tuple[str, ...]
    last_modified: float

    @property
    def class_text(self) -> str:
        return UI_TEXT[f"class_{self.class_key}"]


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        if exe_dir.name.lower() == "dist":
            return exe_dir.parent
        return exe_dir
    return Path(__file__).resolve().parent


def series_root() -> Path:
    return app_dir().parent.parent


def icon_path() -> Path:
    if getattr(sys, "frozen", False):
        meipass = Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
        bundled = meipass / "dake_icon.ico"
        if bundled.exists():
            return bundled
    return series_root() / "02_assets" / "dake_icon.ico"


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


def safe_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return UI_TEXT["value_yes"] if value else UI_TEXT["value_no"]
    return str(value).strip()


def safe_bool(value: object, default: bool = True) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        lowered = value.strip().lower()
        if lowered in {"true", "1", "yes", "on"}:
            return True
        if lowered in {"false", "0", "no", "off"}:
            return False
    if isinstance(value, (int, float)):
        return bool(value)
    return default


def bool_label(value: bool) -> str:
    return UI_TEXT["value_yes"] if value else UI_TEXT["value_no"]


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError:
        return path.read_text(encoding="utf-8", errors="replace")


def load_json_file(path: Path) -> dict[str, object]:
    try:
        loaded = json.loads(read_text(path))
    except Exception:
        return {}
    return loaded if isinstance(loaded, dict) else {}


def extract_dake_web_meta(readme_path: Path) -> tuple[dict[str, object], str]:
    if not readme_path.exists():
        return {}, UI_TEXT["missing_readme"]
    try:
        content = read_text(readme_path)
    except OSError as exc:
        return {}, str(exc)
    match = DAKE_WEB_META_PATTERN.search(content)
    if not match:
        return {}, UI_TEXT["missing_meta"]
    try:
        data = json.loads(match.group(1))
    except json.JSONDecodeError as exc:
        return {}, str(exc)
    if not isinstance(data, dict):
        return {}, UI_TEXT["missing_meta"]
    return data, ""


def readme_title(readme_path: Path) -> str:
    if not readme_path.exists():
        return ""
    try:
        for line in read_text(readme_path).splitlines():
            stripped = line.strip()
            if stripped.startswith("# "):
                return stripped[2:].strip()
    except OSError:
        return ""
    return ""


def first_url_from_readme(readme_path: Path, prefer_github: bool = False) -> str:
    if not readme_path.exists():
        return ""
    try:
        text = read_text(readme_path)
    except OSError:
        return ""
    if prefer_github:
        match = GITHUB_PATTERN.search(text)
        return match.group(0).rstrip(".,") if match else ""
    urls = URL_PATTERN.findall(text)
    for url in urls:
        lowered = url.lower()
        if "github.com" not in lowered and "localhost" not in lowered and "127.0.0.1" not in lowered:
            return url.rstrip(".,")
    return ""


def first_cloudflare_url_from_readme(readme_path: Path) -> str:
    if not readme_path.exists():
        return ""
    try:
        text = read_text(readme_path)
    except OSError:
        return ""
    for url in URL_PATTERN.findall(text):
        lowered = url.lower()
        if "dash.cloudflare.com" in lowered or ".pages.dev" in lowered or "workers.dev" in lowered:
            return url.rstrip(".,")
    return ""


def domain_from_url(url: str) -> str:
    if not url:
        return ""
    try:
        without_scheme = re.sub(r"^https?://", "", url.strip(), flags=re.IGNORECASE)
        return without_scheme.split("/", 1)[0].strip()
    except Exception:
        return ""


def normalize_http_url(url: str) -> str:
    value = str(url or "").strip()
    if not re.match(r"^https?://", value, flags=re.IGNORECASE):
        return ""
    if SECRET_URL_PATTERN.search(value):
        return ""
    return value


def url_from_domain(domain: str) -> str:
    value = str(domain or "").strip()
    if not value or SECRET_URL_PATTERN.search(value):
        return ""
    if re.match(r"^https?://", value, flags=re.IGNORECASE):
        return normalize_http_url(value)
    if re.match(r"^[A-Za-z0-9.-]+(?::\d+)?$", value):
        return f"https://{value}"
    return ""


def first_safe_url(*values: str) -> str:
    for value in values:
        normalized = normalize_http_url(value)
        if normalized:
            return normalized
    return ""


def package_display_name(package_path: Path) -> str:
    data = load_json_file(package_path)
    name = safe_text(data.get("name", ""))
    return name


def wrangler_project_name(wrangler_path: Path) -> str:
    if not wrangler_path.exists():
        return ""
    try:
        text = read_text(wrangler_path)
    except OSError:
        return ""
    match = WRANGLER_NAME_PATTERN.search(text)
    return match.group(1).strip() if match else ""


def wrangler_has_pages_output(wrangler_path: Path) -> bool:
    if not wrangler_path.exists():
        return False
    try:
        return bool(PAGES_OUTPUT_PATTERN.search(read_text(wrangler_path)))
    except OSError:
        return False


def route_file_for(folder: Path) -> Path | None:
    public_route = folder / PUBLIC_ROUTES
    root_route = folder / ROUTES_NAME
    if public_route.exists():
        return public_route
    if root_route.exists():
        return root_route
    return None


def routes_include_api(route_path: Path | None) -> bool:
    if route_path is None:
        return False
    try:
        text = read_text(route_path)
    except OSError:
        return False
    try:
        data = json.loads(text)
        includes = data.get("include", []) if isinstance(data, dict) else []
        if isinstance(includes, list):
            return any("/api/" in str(item) or str(item).strip() == "/api/*" for item in includes)
    except Exception:
        pass
    return "/api/*" in text or "/api/" in text


def has_health_file(folder: Path) -> bool:
    candidates = [
        folder / "functions" / "api" / "health.js",
        folder / "functions" / "api" / "health.ts",
        folder / "functions" / "api" / "health.mjs",
        folder / "functions" / "api" / "health" / "index.js",
        folder / "functions" / "api" / "health" / "index.ts",
    ]
    if any(path.exists() for path in candidates):
        return True
    api_dir = folder / "functions" / "api"
    if not api_dir.exists() or not api_dir.is_dir():
        return False
    try:
        return any("health" in path.stem.lower() for path in api_dir.rglob("*") if path.is_file())
    except OSError:
        return False


def gitignore_mentions_node_modules(gitignore_path: Path) -> bool:
    if not gitignore_path.exists():
        return False
    try:
        text = read_text(gitignore_path).lower()
    except OSError:
        return False
    return "node_modules" in text


def node_modules_ignored_guess(folder: Path) -> bool:
    node_modules = folder / "node_modules"
    gitignore_path = folder / GITIGNORE_NAME
    if gitignore_mentions_node_modules(gitignore_path):
        return True
    return not node_modules.exists()


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


def is_excluded_dir(path: Path) -> bool:
    return path.name.lower() in EXCLUDED_DIR_NAMES


def build_candidate_info(folder: Path, parent_is_site: bool = False) -> CandidateInfo:
    name = folder.name.lower()
    reasons: list[str] = []
    score = 0

    has_git = (folder / GIT_DIR).exists()
    has_readme = (folder / README_NAME).exists()
    has_wrangler = (folder / WRANGLER_NAME).exists()
    has_package = (folder / PACKAGE_NAME).exists()
    has_public = (folder / PUBLIC_DIR).is_dir()
    has_functions = (folder / FUNCTIONS_DIR).is_dir()
    has_meta = False

    if has_git:
        score += 4
        reasons.append(UI_TEXT["reason_git_exists"])
    if has_readme:
        score += 2
        reasons.append(UI_TEXT["reason_readme_exists"])
        try:
            has_meta = bool(DAKE_WEB_META_PATTERN.search(read_text(folder / README_NAME)))
        except OSError:
            has_meta = False
    if has_wrangler:
        score += 3
        reasons.append(UI_TEXT["reason_wrangler_exists"])
    if has_package:
        score += 2
        reasons.append(UI_TEXT["reason_package_exists"])
    if has_public:
        score += 1
        reasons.append(UI_TEXT["reason_public_exists"])
    if has_functions:
        score += 2
        reasons.append(UI_TEXT["reason_functions_exists"])
    if has_meta:
        score += 4
        reasons.append(UI_TEXT["reason_meta_exists"])
    if has_package and has_public:
        score += 2
        reasons.append(UI_TEXT["reason_package_public"])
    if has_wrangler and has_functions:
        score += 2
        reasons.append(UI_TEXT["reason_wrangler_functions"])
    if name.endswith("-site") or name == "site" or "site" in name:
        score += 1
        reasons.append(UI_TEXT["reason_name_site"])

    if name in COMPONENT_DIR_NAMES:
        if name == "public":
            reasons.append(UI_TEXT["reason_component_public"])
        elif name == "functions":
            reasons.append(UI_TEXT["reason_component_functions"])
        else:
            reasons.append(UI_TEXT["reason_component_sitemap"])
        if parent_is_site and not has_git:
            reasons.append(UI_TEXT["reason_parent_site_component"])
            return CandidateInfo(score=score, candidate_type="ignored_component", reasons=tuple(reasons))
        if not has_git:
            return CandidateInfo(score=score, candidate_type="ignored_component", reasons=tuple(reasons))

    if not has_git and not has_readme and not has_wrangler:
        reasons.append(UI_TEXT["reason_low_signal"])
        return CandidateInfo(score=score, candidate_type="ignored_component", reasons=tuple(reasons))

    strong_site = has_git or has_wrangler or (has_package and has_public) or (has_readme and has_meta)
    if strong_site or score >= 5:
        candidate_type = "site"
    elif score >= 2:
        candidate_type = "candidate"
    else:
        candidate_type = "ignored_component"
    return CandidateInfo(score=score, candidate_type=candidate_type, reasons=tuple(reasons))


def discover_site_candidates(root: Path) -> list[tuple[Path, CandidateInfo]]:
    found: dict[str, tuple[Path, CandidateInfo]] = {}
    if not root.exists() or not root.is_dir():
        return []

    def walk(folder: Path, depth: int, parent_is_site: bool = False) -> None:
        if depth > MAX_SCAN_DEPTH:
            return
        try:
            children = sorted((path for path in folder.iterdir() if path.is_dir()), key=lambda path: path.name.lower())
        except OSError:
            return
        for child in children:
            if is_excluded_dir(child):
                continue
            info = build_candidate_info(child, parent_is_site=parent_is_site)
            if info.candidate_type != "ignored_component":
                try:
                    found[str(child.resolve()).lower()] = (child, info)
                except OSError:
                    found[str(child).lower()] = (child, info)
            if depth < MAX_SCAN_DEPTH:
                walk(child, depth + 1, parent_is_site=info.candidate_type == "site")

    walk(root, 0)
    return [item for _key, item in sorted(found.items(), key=lambda pair: str(pair[1][0]).lower())]


def has_site_marker(folder: Path, depth: int) -> bool:
    info = build_candidate_info(folder)
    return info.candidate_type != "ignored_component"


def discover_site_folders(root: Path) -> list[Path]:
    return [folder for folder, _info in discover_site_candidates(root)]


def run_git_command(folder: Path, args: list[str]) -> tuple[int, str, str]:
    flags = 0
    if sys.platform.startswith("win") and hasattr(subprocess, "CREATE_NO_WINDOW"):
        flags = subprocess.CREATE_NO_WINDOW
    try:
        completed = subprocess.run(
            ["git", *args],
            cwd=str(folder),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=GIT_TIMEOUT_SECONDS,
            creationflags=flags,
            check=False,
        )
        return completed.returncode, completed.stdout.strip(), completed.stderr.strip()
    except Exception as exc:
        return 1, "", str(exc)


def parse_ahead_behind(status_line: str) -> tuple[int, int]:
    ahead = 0
    behind = 0
    ahead_match = re.search(r"ahead\s+(\d+)", status_line)
    behind_match = re.search(r"behind\s+(\d+)", status_line)
    if ahead_match:
        ahead = int(ahead_match.group(1))
    if behind_match:
        behind = int(behind_match.group(1))
    return ahead, behind


def read_git_info(folder: Path) -> GitInfo:
    if not (folder / GIT_DIR).exists():
        return GitInfo(is_repo=False)

    code, short_status, error = run_git_command(folder, ["status", "--short"])
    if code != 0:
        return GitInfo(is_repo=True, error=error or UI_TEXT["value_get_failed"])

    branch_code, branch, branch_error = run_git_command(folder, ["rev-parse", "--abbrev-ref", "HEAD"])
    log_code, commit, log_error = run_git_command(folder, ["log", "-1", "--pretty=%h"])
    sb_code, status_sb, sb_error = run_git_command(folder, ["status", "-sb"])

    if branch_code != 0:
        branch = ""
    if log_code != 0:
        commit = ""
    if sb_code != 0:
        status_sb = ""

    lines = [line for line in short_status.splitlines() if line.strip()]
    dirty_count = len(lines)
    untracked_count = sum(1 for line in lines if line.startswith("??"))
    first_status_line = status_sb.splitlines()[0] if status_sb else ""
    ahead, behind = parse_ahead_behind(first_status_line)
    combined_error = branch_error or log_error or sb_error
    return GitInfo(
        is_repo=True,
        branch=branch,
        commit=commit,
        dirty_count=dirty_count,
        untracked_count=untracked_count,
        ahead=ahead,
        behind=behind,
        error="" if branch or commit or status_sb else combined_error,
    )


def iter_code_files(folder: Path) -> list[Path]:
    files: list[Path] = []

    def walk(current: Path, depth: int) -> None:
        if len(files) >= MAX_CODE_SCAN_FILES or depth > 4:
            return
        try:
            children = list(current.iterdir())
        except OSError:
            return
        for child in children:
            if len(files) >= MAX_CODE_SCAN_FILES:
                return
            if child.is_dir():
                if is_excluded_dir(child):
                    continue
                walk(child, depth + 1)
                continue
            if child.suffix.lower() not in CODE_SUFFIXES:
                continue
            try:
                if child.stat().st_size > MAX_CODE_FILE_BYTES:
                    continue
            except OSError:
                continue
            files.append(child)

    walk(folder, 0)
    return files


def is_under_functions(path: Path, folder: Path) -> bool:
    try:
        rel = path.relative_to(folder)
    except ValueError:
        return False
    return bool(rel.parts) and rel.parts[0].lower() == FUNCTIONS_DIR


def looks_like_frontend_file(path: Path, folder: Path) -> bool:
    if path.suffix.lower() not in FRONTEND_SUFFIXES:
        return False
    if is_under_functions(path, folder):
        return False
    try:
        rel = path.relative_to(folder)
    except ValueError:
        return False
    lowered_parts = {part.lower() for part in rel.parts}
    return bool(lowered_parts & {"public", "src", "app", "pages", "components"}) or len(rel.parts) <= 2


def scan_openai_usage(folder: Path) -> ApiInfo:
    has_env_ref = False
    has_key = False
    has_front_direct = False
    has_functions_openai = False

    for path in iter_code_files(folder):
        try:
            text = read_text(path)
        except OSError:
            continue
        if "OPENAI_API_KEY" in text:
            has_env_ref = True
        if SK_KEY_PATTERN.search(text):
            has_key = True
        if "api.openai.com" in text:
            if is_under_functions(path, folder):
                has_functions_openai = True
            elif looks_like_frontend_file(path, folder):
                has_front_direct = True
        if is_under_functions(path, folder) and ("OPENAI_API_KEY" in text or "openai" in text.lower()):
            has_functions_openai = True

    has_env_design = has_env_ref and has_functions_openai and not has_key and not has_front_direct
    return ApiInfo(
        has_openai_env_ref=has_env_ref,
        has_hardcoded_key_suspect=has_key,
        has_frontend_openai_direct=has_front_direct,
        has_functions_openai=has_functions_openai,
        has_env_design=has_env_design,
    )


def build_file_checks(folder: Path, meta_error: str, meta: dict[str, object]) -> FileChecks:
    return FileChecks(
        has_readme=(folder / README_NAME).exists(),
        has_meta=bool(meta),
        meta_error=meta_error,
        has_git=(folder / GIT_DIR).exists(),
        has_package=(folder / PACKAGE_NAME).exists(),
        has_wrangler=(folder / WRANGLER_NAME).exists(),
        has_public=(folder / PUBLIC_DIR).is_dir(),
        has_functions=(folder / FUNCTIONS_DIR).is_dir(),
        has_routes=route_file_for(folder) is not None,
        has_gitignore=(folder / GITIGNORE_NAME).exists(),
        node_modules_ignored=node_modules_ignored_guess(folder),
    )


def build_cloudflare_info(folder: Path, files: FileChecks) -> CloudflareInfo:
    wrangler_path = folder / WRANGLER_NAME
    route_path = route_file_for(folder)
    has_pages_output = wrangler_has_pages_output(wrangler_path)
    has_functions_api = (folder / "functions" / "api").is_dir()
    has_health = has_health_file(folder)
    has_routes_api = routes_include_api(route_path)
    likely_pages = files.has_wrangler or has_pages_output or files.has_routes or files.has_functions
    return CloudflareInfo(
        has_pages_output=has_pages_output,
        likely_pages=likely_pages,
        has_functions_api=has_functions_api,
        has_health_file=has_health,
        has_routes_api=has_routes_api,
        route_file=str(route_path.relative_to(folder)) if route_path else "",
    )


def classify_record(
    files: FileChecks,
    cloudflare: CloudflareInfo,
    api: ApiInfo,
    git: GitInfo,
    status_value: str,
    show_on_dashboard: bool,
    domain: str,
    production_url: str,
    cloudflare_project: str,
    health_url: str,
) -> str:
    lowered_status = status_value.strip().lower()
    if lowered_status in {"internal", "frozen", "archived"} or not show_on_dashboard:
        return "internal"
    if (
        api.has_hardcoded_key_suspect
        or api.has_frontend_openai_direct
        or (
            files.has_functions
            and (api.has_openai_env_ref or api.has_functions_openai)
            and (not health_url or not cloudflare.has_health_file or not cloudflare.has_routes_api)
        )
    ):
        return "api_review"
    if (
        not files.has_readme
        or not files.has_meta
        or not git.is_repo
        or bool(git.error)
        or (files.has_functions and not files.has_wrangler)
        or (cloudflare.likely_pages and not files.has_package)
        or git.dirty_count >= 10
    ):
        return "needs_review"
    if (
        cloudflare.likely_pages
        and not git.has_dirty
        and (not (production_url or domain) or not cloudflare_project or git.ahead > 0 or git.behind > 0)
    ):
        return "deploy_review"
    if files.has_readme and git.is_repo and files.has_wrangler and files.has_package and (production_url or domain) and not git.has_dirty:
        return "normal"
    return "needs_review"


def build_missing_items(
    files: FileChecks,
    cloudflare: CloudflareInfo,
    api: ApiInfo,
    git: GitInfo,
    domain: str,
    production_url: str,
    cloudflare_project: str,
) -> tuple[str, ...]:
    items: list[str] = []
    if not files.has_readme:
        items.append(UI_TEXT["missing_readme"])
    if not files.has_meta:
        items.append(UI_TEXT["missing_meta"])
    if not git.is_repo:
        items.append(UI_TEXT["missing_git"])
    elif git.error:
        items.append(UI_TEXT["missing_git_error"])
    if cloudflare.likely_pages and not files.has_package:
        items.append(UI_TEXT["missing_package"])
    if files.has_functions and not files.has_wrangler:
        items.append(UI_TEXT["missing_wrangler"])
    if cloudflare.likely_pages and not (domain or production_url):
        items.append(UI_TEXT["missing_domain"])
    if cloudflare.likely_pages and not cloudflare_project:
        items.append(UI_TEXT["missing_cloudflare_project"])
    if files.has_functions and not files.has_routes:
        items.append(UI_TEXT["missing_routes"])
    if files.has_functions and not cloudflare.has_health_file:
        items.append(UI_TEXT["missing_health"])
    if api.has_hardcoded_key_suspect:
        items.append(UI_TEXT["missing_api_key"])
    if api.has_frontend_openai_direct:
        items.append(UI_TEXT["missing_front_direct"])
    if git.has_dirty:
        items.append(UI_TEXT["missing_git_dirty"])
    return tuple(items) if items else (UI_TEXT["missing_none"],)


def build_next_items(
    files: FileChecks,
    cloudflare: CloudflareInfo,
    api: ApiInfo,
    git: GitInfo,
    domain: str,
    production_url: str,
) -> tuple[str, ...]:
    tasks: list[str] = []
    if api.has_hardcoded_key_suspect:
        tasks.append(UI_TEXT["next_fix_key"])
    if api.has_frontend_openai_direct:
        tasks.append(UI_TEXT["next_fix_front_direct"])
    if not files.has_readme:
        tasks.append(UI_TEXT["next_add_readme"])
    if not files.has_meta:
        tasks.append(UI_TEXT["next_add_meta"])
    if git.has_dirty:
        tasks.append(UI_TEXT["next_git_commit"])
    if cloudflare.likely_pages and not files.has_wrangler:
        tasks.append(UI_TEXT["next_add_wrangler"])
    if files.has_functions and not cloudflare.has_health_file:
        tasks.append(UI_TEXT["next_add_health"])
    if cloudflare.likely_pages and not (domain or production_url):
        tasks.append(UI_TEXT["next_add_domain"])
    if not tasks:
        tasks.append(UI_TEXT["next_no_action"])
    return tuple(tasks[:5])


def unique_reasons(items: list[str]) -> tuple[str, ...]:
    result: list[str] = []
    for item in items:
        if item and item not in result:
            result.append(item)
    return tuple(result) if result else (UI_TEXT["reason_none"],)


def build_status_reasons(
    files: FileChecks,
    cloudflare: CloudflareInfo,
    api: ApiInfo,
    git: GitInfo,
    class_key: str,
    status_value: str,
    show_on_dashboard: bool,
    domain: str,
    production_url: str,
    cloudflare_project: str,
) -> tuple[str, ...]:
    reasons: list[str] = []
    if not files.has_readme:
        reasons.append(UI_TEXT["reason_status_readme_missing"])
    if not files.has_meta:
        reasons.append(UI_TEXT["reason_status_meta_missing"])
    if git.error:
        reasons.append(UI_TEXT["reason_status_git_error"])
    if cloudflare.likely_pages and not (domain or production_url):
        reasons.append(UI_TEXT["reason_status_production_missing"])
    if cloudflare.likely_pages and not cloudflare_project:
        reasons.append(UI_TEXT["reason_status_cloudflare_missing"])
    if files.has_functions and not cloudflare.has_health_file:
        reasons.append(UI_TEXT["reason_status_functions_health"])
    if class_key == "api_review" or api.has_hardcoded_key_suspect or api.has_frontend_openai_direct:
        reasons.append(UI_TEXT["reason_status_api_review"])
    if class_key == "internal" or status_value.lower() in {"internal", "frozen", "archived"} or not show_on_dashboard:
        reasons.append(UI_TEXT["reason_status_internal"])
    if not reasons:
        reasons.append(UI_TEXT["reason_status_clean"])
    return unique_reasons(reasons)


def build_api_reasons(files: FileChecks, cloudflare: CloudflareInfo, api: ApiInfo, health_url: str) -> tuple[str, ...]:
    reasons: list[str] = []
    if api.has_hardcoded_key_suspect:
        reasons.append(UI_TEXT["reason_api_key"])
    if api.has_frontend_openai_direct:
        reasons.append(UI_TEXT["reason_api_front_direct"])
    if api.has_openai_env_ref:
        reasons.append(UI_TEXT["reason_api_env"])
    if api.has_functions_openai:
        reasons.append(UI_TEXT["reason_api_functions"])
    if files.has_functions and (not health_url or not cloudflare.has_health_file):
        reasons.append(UI_TEXT["reason_api_health_missing"])
    if files.has_functions and not cloudflare.has_routes_api:
        reasons.append(UI_TEXT["reason_api_routes_missing"])
    if not reasons:
        reasons.append(UI_TEXT["reason_api_none"])
    return unique_reasons(reasons)


def build_git_reasons(git: GitInfo) -> tuple[str, ...]:
    reasons: list[str] = []
    if not git.is_repo:
        reasons.append(UI_TEXT["reason_git_no_repo"])
    elif git.error:
        reasons.append(UI_TEXT["reason_git_error"])
    else:
        if git.has_dirty:
            reasons.append(UI_TEXT["reason_git_dirty"])
        if git.has_untracked:
            reasons.append(UI_TEXT["reason_git_untracked"])
        if git.ahead > 0:
            reasons.append(UI_TEXT["reason_git_ahead"])
        if git.behind > 0:
            reasons.append(UI_TEXT["reason_git_behind"])
        if not reasons:
            reasons.append(UI_TEXT["reason_git_clean"])
    return unique_reasons(reasons)


def build_next_action_reasons(
    record_name: str,
    class_key: str,
    status_reasons: tuple[str, ...],
    api_reasons: tuple[str, ...],
    git_reasons: tuple[str, ...],
) -> tuple[str, ...]:
    reasons: list[str] = []
    if class_key == "api_review":
        reasons.append(f"{record_name}: {UI_TEXT['summary_api_review']}: {api_reasons[0]}")
    if class_key == "needs_review":
        reasons.append(f"{record_name}: {UI_TEXT['summary_needs_review']}: {status_reasons[0]}")
    if UI_TEXT["reason_git_dirty"] in git_reasons:
        reasons.append(f"{record_name}: {UI_TEXT['label_git']}: {UI_TEXT['reason_git_dirty']}")
    if class_key == "deploy_review":
        reasons.append(f"{record_name}: {UI_TEXT['summary_deploy_review']}: {status_reasons[0]}")
    if not reasons:
        reasons.append(f"{record_name}: {UI_TEXT['next_no_action']}")
    return tuple(reasons[:5])


def short_reason(record: SiteRecord) -> str:
    source: tuple[str, ...]
    if record.candidate_type == "candidate":
        source = record.detection_reasons
    elif record.class_key == "api_review":
        source = record.api_reasons
    elif record.git.error or record.git.has_dirty:
        source = record.git_reasons
    else:
        source = record.status_reasons
    return " / ".join(reason for reason in source[:2] if reason)


def scan_site(folder: Path, candidate_info: CandidateInfo | None = None) -> SiteRecord:
    candidate_info = candidate_info or build_candidate_info(folder)
    readme_path = folder / README_NAME
    package_path = folder / PACKAGE_NAME
    wrangler_path = folder / WRANGLER_NAME
    meta, meta_error = extract_dake_web_meta(readme_path)
    files = build_file_checks(folder, meta_error, meta)
    cloudflare = build_cloudflare_info(folder, files)
    api = scan_openai_usage(folder)
    git = read_git_info(folder)

    inferred_production = first_url_from_readme(readme_path)
    production_url = safe_text(meta.get("production_url", "")) or inferred_production
    health_url = safe_text(meta.get("health_url", ""))
    domain = safe_text(meta.get("domain", "")) or domain_from_url(production_url)
    cloudflare_project = safe_text(meta.get("cloudflare_project", "")) or wrangler_project_name(wrangler_path)
    package_name = package_display_name(package_path) if package_path.exists() else ""
    display_name = safe_text(meta.get("display_name", "")) or readme_title(readme_path) or package_name or folder.name
    site_type = safe_text(meta.get("site_type", "")) or (UI_TEXT["value_functions"] if files.has_functions else UI_TEXT["value_static"])
    status_value = safe_text(meta.get("status", "")) or UI_TEXT["value_unknown"]
    show_on_dashboard = safe_bool(meta.get("show_on_dashboard", True), default=True)
    github_url = first_url_from_readme(readme_path, prefer_github=True)
    cloudflare_url = first_safe_url(
        safe_text(meta.get("cloudflare_url", "")),
        safe_text(meta.get("cloudflare_project_url", "")),
        safe_text(meta.get("cloudflare_dashboard_url", "")),
        first_cloudflare_url_from_readme(readme_path),
    )
    class_key = classify_record(
        files,
        cloudflare,
        api,
        git,
        status_value,
        show_on_dashboard,
        domain,
        production_url,
        cloudflare_project,
        health_url,
    )
    if candidate_info.candidate_type == "candidate":
        class_key = "candidate"
    status_reasons = build_status_reasons(
        files,
        cloudflare,
        api,
        git,
        class_key,
        status_value,
        show_on_dashboard,
        domain,
        production_url,
        cloudflare_project,
    )
    api_reasons = build_api_reasons(files, cloudflare, api, health_url)
    git_reasons = build_git_reasons(git)
    missing_items = build_missing_items(files, cloudflare, api, git, domain, production_url, cloudflare_project)
    next_action_reasons = build_next_action_reasons(folder.name, class_key, status_reasons, api_reasons, git_reasons)
    next_items = next_action_reasons or build_next_items(files, cloudflare, api, git, domain, production_url)
    last_modified = latest_mtime(
        [
            readme_path,
            package_path,
            wrangler_path,
            folder / PUBLIC_ROUTES,
            folder / ROUTES_NAME,
            folder / FUNCTIONS_DIR,
            folder / PUBLIC_DIR,
        ],
        folder,
    )
    return SiteRecord(
        folder_name=folder.name,
        folder_path=folder,
        display_name=display_name,
        domain=domain,
        cloudflare_project=cloudflare_project,
        site_type=site_type,
        status_value=status_value,
        show_on_dashboard=show_on_dashboard,
        production_url=production_url,
        health_url=health_url,
        github_url=github_url,
        cloudflare_url=cloudflare_url,
        files=files,
        cloudflare=cloudflare,
        api=api,
        git=git,
        class_key=class_key,
        detection_score=candidate_info.score,
        detection_reasons=candidate_info.reasons,
        status_reasons=status_reasons,
        api_reasons=api_reasons,
        git_reasons=git_reasons,
        next_action_reasons=next_action_reasons,
        candidate_type=candidate_info.candidate_type,
        missing_items=missing_items,
        next_items=next_items,
        last_modified=last_modified,
    )


def error_site(folder: Path, exc: Exception) -> SiteRecord:
    candidate_info = build_candidate_info(folder)
    files = FileChecks(
        has_readme=(folder / README_NAME).exists(),
        has_meta=False,
        meta_error=str(exc),
        has_git=(folder / GIT_DIR).exists(),
        has_package=False,
        has_wrangler=False,
        has_public=False,
        has_functions=False,
        has_routes=False,
        has_gitignore=False,
        node_modules_ignored=True,
    )
    cloudflare = CloudflareInfo(False, False, False, False, False, "")
    api = ApiInfo(False, False, False, False, False)
    git = GitInfo(is_repo=files.has_git, error=str(exc) if files.has_git else "")
    return SiteRecord(
        folder_name=folder.name,
        folder_path=folder,
        display_name=folder.name,
        domain="",
        cloudflare_project="",
        site_type=UI_TEXT["value_unknown"],
        status_value=UI_TEXT["value_unknown"],
        show_on_dashboard=True,
        production_url="",
        health_url="",
        github_url="",
        cloudflare_url="",
        files=files,
        cloudflare=cloudflare,
        api=api,
        git=git,
        class_key="needs_review",
        detection_score=candidate_info.score,
        detection_reasons=candidate_info.reasons,
        status_reasons=(str(exc),),
        api_reasons=(UI_TEXT["reason_api_none"],),
        git_reasons=build_git_reasons(git),
        next_action_reasons=(UI_TEXT["next_no_action"],),
        candidate_type=candidate_info.candidate_type,
        missing_items=(str(exc),),
        next_items=(UI_TEXT["next_no_action"],),
        last_modified=latest_mtime([folder / README_NAME], folder),
    )


def scan_sites(root: Path = DEV_ROOT) -> list[SiteRecord]:
    records: list[SiteRecord] = []
    for folder, candidate_info in discover_site_candidates(root):
        try:
            records.append(scan_site(folder, candidate_info))
        except Exception as exc:
            records.append(error_site(folder, exc))
    return records


def git_status_label(git: GitInfo) -> str:
    if not git.is_repo:
        return UI_TEXT["value_not_repo"]
    if git.error:
        return UI_TEXT["value_get_failed"]
    if git.dirty_count:
        return f"{UI_TEXT['summary_dirty']} {git.dirty_count}"
    if git.ahead or git.behind:
        return f"{UI_TEXT['value_clean']} +{git.ahead}/-{git.behind}"
    return UI_TEXT["value_clean"]


def api_label(record: SiteRecord) -> str:
    if record.api.has_hardcoded_key_suspect:
        return UI_TEXT["value_direct_key"]
    if record.api.has_frontend_openai_direct:
        return UI_TEXT["value_front_direct"]
    if record.class_key == "api_review":
        return UI_TEXT["value_api_review"]
    if record.api.has_env_design:
        return UI_TEXT["value_env_design"]
    if record.files.has_functions:
        return UI_TEXT["value_functions"]
    return UI_TEXT["value_static"]


def functions_label(record: SiteRecord) -> str:
    if not record.files.has_functions:
        return UI_TEXT["value_no"]
    if record.cloudflare.has_health_file:
        return f"{UI_TEXT['value_yes']} / health"
    return UI_TEXT["value_yes"]


def record_key(record: SiteRecord) -> str:
    return str(record.folder_path).lower()


def watched_folder_for_path(path: Path, root: Path) -> str | None:
    try:
        relative = path.resolve().relative_to(root.resolve())
    except (OSError, ValueError):
        return None
    parts = relative.parts
    if not parts:
        return None
    lowered_parts = [part.lower() for part in parts]
    if any(part in EXCLUDED_DIR_NAMES for part in lowered_parts):
        return None
    if path.name in WATCHED_FILE_NAMES or any(part in WATCHED_DIR_NAMES for part in lowered_parts):
        return parts[0]
    return None


class WebDashboardApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.geometry("1280x820")
        self.root.minsize(1080, 680)
        self.root.configure(bg=THEME["bg"])
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        apply_window_icon(self.root)

        self.font_family = choose_font_family(root)
        self.records: list[SiteRecord] = []
        self.visible_records: list[SiteRecord] = []
        self.record_by_iid: dict[str, SiteRecord] = {}
        self.selected_record: SiteRecord | None = None
        self.previous_map: dict[str, SiteRecord] = {}
        self.filter_key = "all"
        self.worker_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.worker_thread: threading.Thread | None = None
        self.pending_reload_folder: str | None = None
        self.auto_update_enabled = True
        self.watch_enabled = True
        self.watch_observer = None
        self.watch_pending_folder: str | None = None
        self.watch_debounce_job: str | None = None
        self.last_watch_event_at = 0.0

        self.last_loaded_var = tk.StringVar(value=UI_TEXT["last_loaded_waiting"])
        self.watch_status_var = tk.StringVar(value=UI_TEXT["watch_status_off"])
        self.auto_status_var = tk.StringVar(value=UI_TEXT["auto_status_on"])
        self.status_var = tk.StringVar(value="")
        self.search_var = tk.StringVar()
        self.count_var = tk.StringVar(value=UI_TEXT["count_line"].format(visible=0, total=0))
        self.detail_title_var = tk.StringVar(value=UI_TEXT["detail_empty"])
        self.detail_badge_var = tk.StringVar(value=UI_TEXT["value_unset"])
        self.notification_var = tk.StringVar(value="")
        self.summary_vars = {key: tk.StringVar(value="0") for key in ("total", "normal", "needs_review", "api_review", "deploy_review", "dirty")}
        self.qpsc_vars = {
            key: tk.StringVar(value="0")
            for key in (
                "new_sites",
                "readme_missing",
                "meta_missing",
                "api_review",
                "cloudflare_review",
                "git_dirty",
                "git_error",
                "candidate_components",
            )
        }
        self.git_vars = {key: tk.StringVar(value="0") for key in ("dirty_sites", "untracked_sites", "ahead_sites", "error_sites")}
        self.filter_buttons: dict[str, tk.Button] = {}
        self.detail_link_counter = 0

        self.configure_styles()
        self.build_ui()
        self.search_var.trace_add("write", lambda *_args: self.apply_filters())
        self.root.after(WORKER_POLL_MS, self.poll_worker)
        self.reload_data(source="startup")
        self.start_watchdog()
        self.schedule_auto_reload()

    def configure_styles(self) -> None:
        style = ttk.Style()
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
            rowheight=31,
            font=(self.font_family, 9),
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
            relief="flat",
            font=(self.font_family, 9, "bold"),
        )
        style.configure("Vertical.TScrollbar", background=THEME["panel_alt"], troughcolor=THEME["bg"])

    def build_ui(self) -> None:
        outer = tk.Frame(self.root, bg=THEME["bg"])
        outer.pack(fill="both", expand=True, padx=22, pady=(18, 0))

        self.build_header(outer)
        self.build_summary(outer)
        self.build_controls(outer)
        self.build_main_area(outer)
        self.build_notification()

        footer = tk.Frame(self.root, bg=THEME["bg"])
        footer.pack(fill="x", padx=22, pady=(8, 12))
        tk.Label(
            footer,
            text=UI_TEXT["footer_note"],
            bg=THEME["bg"],
            fg=THEME["quiet"],
            font=(self.font_family, 8),
        ).pack(side="left")
        tk.Label(
            footer,
            textvariable=self.status_var,
            bg=THEME["bg"],
            fg=THEME["muted"],
            font=(self.font_family, 9),
        ).pack(side="right")

    def build_header(self, parent: tk.Misc) -> None:
        header = tk.Frame(parent, bg=THEME["bg"])
        header.pack(fill="x", pady=(0, 14))
        left = tk.Frame(header, bg=THEME["bg"])
        left.pack(side="left", fill="x", expand=True)
        tk.Label(
            left,
            text=UI_TEXT["header_title"],
            bg=THEME["bg"],
            fg=THEME["text"],
            font=(self.font_family, 22, "bold"),
            anchor="w",
        ).pack(anchor="w")
        tk.Label(
            left,
            text=UI_TEXT["header_subtitle"],
            bg=THEME["bg"],
            fg=THEME["muted"],
            font=(self.font_family, 10),
            anchor="w",
        ).pack(anchor="w", pady=(4, 0))

        right = tk.Frame(header, bg=THEME["bg"])
        right.pack(side="right", anchor="ne")
        tk.Label(right, textvariable=self.last_loaded_var, bg=THEME["bg"], fg=THEME["muted"], font=(self.font_family, 9)).pack(anchor="e")
        tk.Label(right, textvariable=self.watch_status_var, bg=THEME["bg"], fg=THEME["muted"], font=(self.font_family, 9)).pack(anchor="e", pady=(3, 0))
        tk.Label(right, textvariable=self.auto_status_var, bg=THEME["bg"], fg=THEME["muted"], font=(self.font_family, 9)).pack(anchor="e", pady=(3, 0))

        button_row = tk.Frame(right, bg=THEME["bg"])
        button_row.pack(anchor="e", pady=(8, 0))
        self.reload_button = self.make_button(button_row, UI_TEXT["button_reload"], lambda: self.reload_data(source="manual"), primary=True)
        self.reload_button.pack(side="left", padx=(0, 8))
        self.auto_button = self.make_button(button_row, UI_TEXT["button_auto_stop"], self.toggle_auto_update)
        self.auto_button.pack(side="left", padx=(0, 8))
        self.watch_button = self.make_button(button_row, UI_TEXT["button_watch_stop"], self.toggle_watchdog)
        self.watch_button.pack(side="left")

    def build_summary(self, parent: tk.Misc) -> None:
        row = tk.Frame(parent, bg=THEME["bg"])
        row.pack(fill="x", pady=(0, 12))
        summary_items = [
            ("total", UI_TEXT["summary_total"]),
            ("normal", UI_TEXT["summary_normal"]),
            ("needs_review", UI_TEXT["summary_needs_review"]),
            ("api_review", UI_TEXT["summary_api_review"]),
            ("deploy_review", UI_TEXT["summary_deploy_review"]),
            ("dirty", UI_TEXT["summary_dirty"]),
        ]
        for index, (key, label) in enumerate(summary_items):
            row.columnconfigure(index, weight=1)
            self.metric_card(row, label, self.summary_vars[key]).grid(row=0, column=index, sticky="ew", padx=(0 if index == 0 else 8, 0))

        lower = tk.Frame(parent, bg=THEME["bg"])
        lower.pack(fill="x", pady=(0, 12))
        lower.columnconfigure(0, weight=1)
        lower.columnconfigure(1, weight=1)
        self.build_qpsc_card(lower).grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self.build_git_card(lower).grid(row=0, column=1, sticky="nsew")

    def metric_card(self, parent: tk.Misc, label: str, var: tk.StringVar) -> tk.Frame:
        frame = tk.Frame(parent, bg=THEME["panel"], highlightthickness=1, highlightbackground=THEME["border"])
        tk.Label(frame, text=label, bg=THEME["panel"], fg=THEME["muted"], font=(self.font_family, 9)).pack(anchor="w", padx=14, pady=(10, 2))
        tk.Label(frame, textvariable=var, bg=THEME["panel"], fg=THEME["text"], font=(self.font_family, 18, "bold")).pack(anchor="w", padx=14, pady=(0, 10))
        return frame

    def build_qpsc_card(self, parent: tk.Misc) -> tk.Frame:
        frame = self.panel(parent)
        tk.Label(frame, text=UI_TEXT["qpsc_title"], bg=THEME["panel"], fg=THEME["text"], font=(self.font_family, 11, "bold")).pack(anchor="w", padx=14, pady=(12, 2))
        tk.Label(frame, text=UI_TEXT["qpsc_subtitle"], bg=THEME["panel"], fg=THEME["muted"], font=(self.font_family, 8)).pack(anchor="w", padx=14)
        body = tk.Frame(frame, bg=THEME["panel"])
        body.pack(fill="x", padx=14, pady=(10, 12))
        pairs = [
            ("new_sites", UI_TEXT["qpsc_new_sites"]),
            ("readme_missing", UI_TEXT["qpsc_readme_missing"]),
            ("meta_missing", UI_TEXT["qpsc_meta_missing"], UI_TEXT["qpsc_hint_meta"]),
            ("api_review", UI_TEXT["qpsc_api_review"], UI_TEXT["qpsc_hint_api"]),
            ("cloudflare_review", UI_TEXT["qpsc_cloudflare_review"]),
            ("git_dirty", UI_TEXT["qpsc_git_dirty"]),
            ("git_error", UI_TEXT["qpsc_git_error"], UI_TEXT["qpsc_hint_git"]),
            ("candidate_components", UI_TEXT["qpsc_candidate_components"], UI_TEXT["qpsc_hint_candidate"]),
        ]
        self.compact_metrics(body, pairs, self.qpsc_vars)
        return frame

    def build_git_card(self, parent: tk.Misc) -> tk.Frame:
        frame = self.panel(parent)
        tk.Label(frame, text=UI_TEXT["git_card_title"], bg=THEME["panel"], fg=THEME["text"], font=(self.font_family, 11, "bold")).pack(anchor="w", padx=14, pady=(12, 2))
        body = tk.Frame(frame, bg=THEME["panel"])
        body.pack(fill="x", padx=14, pady=(29, 12))
        pairs = [
            ("dirty_sites", UI_TEXT["git_dirty_sites"]),
            ("untracked_sites", UI_TEXT["git_untracked_sites"]),
            ("ahead_sites", UI_TEXT["git_ahead_sites"]),
            ("error_sites", UI_TEXT["git_error_sites"]),
        ]
        self.compact_metrics(body, pairs, self.git_vars)
        return frame

    def compact_metrics(self, parent: tk.Misc, pairs: list[tuple[str, ...]], vars_map: dict[str, tk.StringVar]) -> None:
        for index, item_data in enumerate(pairs):
            key, label = item_data[0], item_data[1]
            hint = item_data[2] if len(item_data) > 2 else ""
            row = index // 4
            column = index % 4
            parent.columnconfigure(column, weight=1)
            item = tk.Frame(parent, bg=THEME["panel_alt"], highlightthickness=1, highlightbackground=THEME["border"])
            item.grid(row=row, column=column, sticky="ew", padx=(0 if column == 0 else 7, 0), pady=(0 if row == 0 else 7, 0))
            tk.Label(item, text=label, bg=THEME["panel_alt"], fg=THEME["muted"], font=(self.font_family, 8)).pack(anchor="w", padx=10, pady=(8, 1))
            tk.Label(item, textvariable=vars_map[key], bg=THEME["panel_alt"], fg=THEME["text"], font=(self.font_family, 15, "bold")).pack(anchor="w", padx=10, pady=(0, 8))
            if hint:
                tk.Label(item, text=hint, bg=THEME["panel_alt"], fg=THEME["quiet"], font=(self.font_family, 7)).pack(anchor="w", padx=10, pady=(0, 8))

    def build_controls(self, parent: tk.Misc) -> None:
        controls = tk.Frame(parent, bg=THEME["bg"])
        controls.pack(fill="x", pady=(0, 12))
        for key in FILTER_KEYS:
            button = self.make_button(controls, UI_TEXT[f"filter_{key}"], lambda value=key: self.set_filter(value))
            button.pack(side="left", padx=(0, 8))
            self.filter_buttons[key] = button
        self.update_filter_buttons()

        search_box = tk.Frame(controls, bg=THEME["bg"])
        search_box.pack(side="right")
        tk.Label(search_box, text=UI_TEXT["search_label"], bg=THEME["bg"], fg=THEME["muted"], font=(self.font_family, 9)).pack(side="left", padx=(0, 8))
        entry = tk.Entry(
            search_box,
            textvariable=self.search_var,
            bg=THEME["input"],
            fg=THEME["text"],
            insertbackground=THEME["text"],
            relief="flat",
            width=38,
            font=(self.font_family, 10),
        )
        entry.pack(side="left", ipady=7)

    def build_main_area(self, parent: tk.Misc) -> None:
        panes = tk.PanedWindow(parent, orient="horizontal", bg=THEME["bg"], sashwidth=8, bd=0, relief="flat")
        panes.pack(fill="both", expand=True)

        list_panel = self.panel(panes)
        detail_panel = self.panel(panes)
        panes.add(list_panel, minsize=680, stretch="always")
        panes.add(detail_panel, minsize=330)

        list_header = tk.Frame(list_panel, bg=THEME["panel"])
        list_header.pack(fill="x", padx=14, pady=(12, 8))
        tk.Label(list_header, text=UI_TEXT["list_title"], bg=THEME["panel"], fg=THEME["text"], font=(self.font_family, 12, "bold")).pack(side="left")
        tk.Label(list_header, textvariable=self.count_var, bg=THEME["panel"], fg=THEME["muted"], font=(self.font_family, 9)).pack(side="right")

        tree_frame = tk.Frame(list_panel, bg=THEME["panel"])
        tree_frame.pack(fill="both", expand=True, padx=14, pady=(0, 14))
        columns = ("status", "folder", "display", "domain", "cloudflare", "git", "api", "functions", "reason", "updated")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", style="Dashboard.Treeview", selectmode="browse")
        headings = {
            "status": UI_TEXT["column_status"],
            "folder": UI_TEXT["column_folder"],
            "display": UI_TEXT["column_display"],
            "domain": UI_TEXT["column_domain"],
            "cloudflare": UI_TEXT["column_cloudflare"],
            "git": UI_TEXT["column_git"],
            "api": UI_TEXT["column_api"],
            "functions": UI_TEXT["column_functions"],
            "reason": UI_TEXT["column_reason"],
            "updated": UI_TEXT["column_updated"],
        }
        widths = {
            "status": 92,
            "folder": 150,
            "display": 170,
            "domain": 150,
            "cloudflare": 135,
            "git": 100,
            "api": 92,
            "functions": 88,
            "reason": 180,
            "updated": 118,
        }
        for column in columns:
            self.tree.heading(column, text=headings[column])
            self.tree.column(column, width=widths[column], minwidth=70, stretch=column in {"display", "domain"})
        for key, (bg, fg) in STATUS_THEME.items():
            self.tree.tag_configure(key, background=bg, foreground=fg)
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview, style="Vertical.TScrollbar")
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        self.tree.bind("<ButtonRelease-1>", self.on_tree_click_open)

        self.build_detail_panel(detail_panel)

    def build_detail_panel(self, parent: tk.Misc) -> None:
        header = tk.Frame(parent, bg=THEME["panel"])
        header.pack(fill="x", padx=14, pady=(12, 8))
        tk.Label(header, text=UI_TEXT["detail_title"], bg=THEME["panel"], fg=THEME["text"], font=(self.font_family, 12, "bold")).pack(anchor="w")
        tk.Label(header, textvariable=self.detail_title_var, bg=THEME["panel"], fg=THEME["muted"], font=(self.font_family, 9), justify="left").pack(anchor="w", pady=(3, 0))
        self.detail_badge = tk.Label(
            header,
            textvariable=self.detail_badge_var,
            bg=THEME["panel_soft"],
            fg=THEME["muted"],
            font=(self.font_family, 9, "bold"),
            padx=10,
            pady=5,
        )
        self.detail_badge.pack(anchor="w", pady=(8, 0))

        buttons = tk.Frame(parent, bg=THEME["panel"])
        buttons.pack(fill="x", padx=14, pady=(0, 10))
        self.open_folder_button = self.make_button(buttons, UI_TEXT["button_open_folder"], self.open_selected_folder)
        self.open_readme_button = self.make_button(buttons, UI_TEXT["button_open_readme"], self.open_selected_readme)
        self.open_production_button = self.make_button(buttons, UI_TEXT["button_open_production"], self.open_selected_production)
        self.open_health_button = self.make_button(buttons, UI_TEXT["button_open_health"], self.open_selected_health)
        self.open_github_button = self.make_button(buttons, UI_TEXT["button_open_github"], self.open_selected_github)
        self.open_cloudflare_button = self.make_button(buttons, UI_TEXT["button_open_cloudflare"], self.open_selected_cloudflare)
        for index, button in enumerate(
            [
                self.open_folder_button,
                self.open_readme_button,
                self.open_production_button,
                self.open_health_button,
                self.open_github_button,
                self.open_cloudflare_button,
            ]
        ):
            button.grid(row=index // 2, column=index % 2, sticky="ew", padx=(0 if index % 2 == 0 else 8, 0), pady=(0, 7))
        buttons.columnconfigure(0, weight=1)
        buttons.columnconfigure(1, weight=1)

        self.detail_text = tk.Text(
            parent,
            bg=THEME["panel_alt"],
            fg=THEME["text"],
            insertbackground=THEME["text"],
            relief="flat",
            bd=0,
            padx=13,
            pady=12,
            wrap="word",
            height=12,
            font=(self.font_family, 9),
        )
        self.detail_text.pack(fill="both", expand=True, padx=14, pady=(0, 10))
        self.detail_text.configure(state="disabled")

        tk.Label(parent, text=UI_TEXT["next_title"], bg=THEME["panel"], fg=THEME["text"], font=(self.font_family, 10, "bold")).pack(anchor="w", padx=14, pady=(0, 6))
        self.next_text = tk.Text(
            parent,
            bg=THEME["panel_alt"],
            fg=THEME["text"],
            insertbackground=THEME["text"],
            relief="flat",
            bd=0,
            padx=13,
            pady=10,
            wrap="word",
            height=5,
            font=(self.font_family, 9),
        )
        self.next_text.pack(fill="x", padx=14, pady=(0, 14))
        self.next_text.configure(state="disabled")
        self.set_detail_buttons_state(False)

    def build_notification(self) -> None:
        self.notification_frame = tk.Frame(
            self.root,
            bg=THEME["accent_soft"],
            highlightthickness=1,
            highlightbackground=THEME["border_active"],
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

    def panel(self, parent: tk.Misc) -> tk.Frame:
        return tk.Frame(parent, bg=THEME["panel"], highlightthickness=1, highlightbackground=THEME["border"])

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
        records = scan_sites(DEV_ROOT)
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
        previous = dict(self.previous_map)
        self.records = list(records) if isinstance(records, list) else []
        current = {record_key(record): record for record in self.records}
        self.previous_map = current
        self.reload_button.configure(state="normal")
        self.last_loaded_var.set(UI_TEXT["last_loaded_value"].format(time=datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        self.status_var.set(UI_TEXT["status_ready"])
        self.update_summary(previous, current)
        self.apply_filters()
        if source == "watch" and isinstance(changed_folder, str) and changed_folder:
            self.show_reload_notification(changed_folder)
        if self.pending_reload_folder:
            folder = self.pending_reload_folder
            self.pending_reload_folder = None
            self.reload_data(source="watch", changed_folder=folder)

    def update_summary(self, previous: dict[str, SiteRecord], current: dict[str, SiteRecord]) -> None:
        counts = {key: 0 for key in self.summary_vars}
        site_records = [record for record in self.records if record.candidate_type == "site"]
        candidate_records = [record for record in self.records if record.candidate_type == "candidate"]
        counts["total"] = len(site_records)
        for record in site_records:
            if record.class_key in counts:
                counts[record.class_key] += 1
            if record.git.has_dirty:
                counts["dirty"] += 1
        for key, value in counts.items():
            self.summary_vars[key].set(str(value))

        new_sites = len(set(current) - set(previous)) if previous else 0
        self.qpsc_vars["new_sites"].set(str(new_sites))
        self.qpsc_vars["readme_missing"].set(str(sum(1 for record in site_records if not record.files.has_readme)))
        self.qpsc_vars["meta_missing"].set(str(sum(1 for record in site_records if not record.files.has_meta)))
        self.qpsc_vars["api_review"].set(str(sum(1 for record in site_records if record.class_key == "api_review")))
        self.qpsc_vars["cloudflare_review"].set(str(sum(1 for record in site_records if record.class_key == "deploy_review")))
        self.qpsc_vars["git_dirty"].set(str(sum(1 for record in site_records if record.git.has_dirty)))
        self.qpsc_vars["git_error"].set(str(sum(1 for record in site_records if record.git.error)))
        self.qpsc_vars["candidate_components"].set(str(len(candidate_records)))

        self.git_vars["dirty_sites"].set(str(sum(1 for record in site_records if record.git.has_dirty)))
        self.git_vars["untracked_sites"].set(str(sum(1 for record in site_records if record.git.has_untracked)))
        self.git_vars["ahead_sites"].set(str(sum(1 for record in site_records if record.git.ahead > 0)))
        self.git_vars["error_sites"].set(str(sum(1 for record in site_records if record.git.error)))

    def schedule_auto_reload(self) -> None:
        self.root.after(AUTO_RELOAD_MS, self.auto_reload)

    def auto_reload(self) -> None:
        if self.auto_update_enabled:
            self.reload_data(source="auto")
        self.schedule_auto_reload()

    def toggle_auto_update(self) -> None:
        self.auto_update_enabled = not self.auto_update_enabled
        self.auto_status_var.set(UI_TEXT["auto_status_on"] if self.auto_update_enabled else UI_TEXT["auto_status_off"])
        self.auto_button.configure(text=UI_TEXT["button_auto_stop"] if self.auto_update_enabled else UI_TEXT["button_auto_start"])

    def toggle_watchdog(self) -> None:
        if self.watch_enabled:
            self.watch_enabled = False
            self.stop_watchdog()
            self.watch_status_var.set(UI_TEXT["watch_status_off"])
            self.watch_button.configure(text=UI_TEXT["button_watch_start"])
            return
        self.watch_enabled = True
        self.watch_button.configure(text=UI_TEXT["button_watch_stop"])
        self.start_watchdog()

    def start_watchdog(self) -> None:
        if not self.watch_enabled:
            return
        if Observer is None or FileSystemEventHandler is None:
            self.watch_status_var.set(UI_TEXT["watch_status_polling"])
            return
        if self.watch_observer is not None:
            return
        try:
            handler = self.create_watchdog_handler()
            observer = Observer()
            observer.schedule(handler, str(DEV_ROOT), recursive=True)
            observer.daemon = True
            observer.start()
            self.watch_observer = observer
            self.watch_status_var.set(UI_TEXT["watch_status_on"])
        except Exception:
            self.watch_status_var.set(UI_TEXT["watch_status_error"])

    def create_watchdog_handler(self):
        app = self
        root_path = DEV_ROOT

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

    def apply_filters(self) -> None:
        query = self.search_var.get().strip().lower()
        filtered: list[SiteRecord] = []
        for record in self.records:
            if not self.matches_filter(record):
                continue
            if query and not self.matches_query(record, query):
                continue
            filtered.append(record)
        self.visible_records = filtered
        self.render_tree()
        total_base = sum(1 for record in self.records if record.candidate_type == ("candidate" if self.filter_key == "candidate" else "site"))
        self.count_var.set(UI_TEXT["count_line"].format(visible=len(filtered), total=total_base))
        if self.selected_record not in filtered:
            self.update_detail(filtered[0] if filtered else None)
            if filtered:
                iid = str(filtered[0].folder_path)
                self.tree.selection_set(iid)
                self.tree.focus(iid)

    def matches_filter(self, record: SiteRecord) -> bool:
        if self.filter_key == "all":
            return record.candidate_type == "site"
        if self.filter_key == "candidate":
            return record.candidate_type == "candidate"
        if record.candidate_type != "site":
            return False
        if self.filter_key == "dirty":
            return record.git.has_dirty
        return record.class_key == self.filter_key

    def matches_query(self, record: SiteRecord, query: str) -> bool:
        haystack = " ".join(
            [
                record.folder_name,
                record.display_name,
                record.domain,
                record.cloudflare_project,
                record.site_type,
                record.status_value,
                record.production_url,
                record.github_url,
                record.cloudflare_url,
                record.candidate_type,
                " ".join(record.detection_reasons),
                " ".join(record.status_reasons),
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
                    record.class_text,
                    record.folder_name,
                    record.display_name,
                    record.domain or UI_TEXT["value_unset"],
                    record.cloudflare_project or UI_TEXT["value_unset"],
                    git_status_label(record.git),
                    api_label(record),
                    functions_label(record),
                    short_reason(record),
                    format_datetime(record.last_modified),
                ),
                tags=(record.class_key,),
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

    def on_tree_click_open(self, event) -> None:
        if self.tree.identify("region", event.x, event.y) != "cell":
            return
        row_id = self.tree.identify_row(event.y)
        column_id = self.tree.identify_column(event.x)
        record = self.record_by_iid.get(row_id)
        if record is None:
            return
        if column_id == "#4":
            self.open_url(first_safe_url(record.production_url, url_from_domain(record.domain)), quiet=True)
        elif column_id == "#5":
            self.open_url(record.cloudflare_url, quiet=True)

    def update_detail(self, record: SiteRecord | None) -> None:
        self.selected_record = record
        if record is None:
            self.detail_title_var.set(UI_TEXT["detail_empty"])
            self.detail_badge_var.set(UI_TEXT["value_unset"])
            self.detail_badge.configure(bg=THEME["panel_soft"], fg=THEME["muted"])
            self.set_text(self.detail_text, UI_TEXT["detail_empty"])
            self.set_text(self.next_text, UI_TEXT["detail_empty"])
            self.set_detail_buttons_state(False)
            return
        badge_bg, badge_fg = STATUS_THEME.get(record.class_key, (THEME["panel_soft"], THEME["muted"]))
        self.detail_badge_var.set(record.class_text)
        self.detail_badge.configure(bg=badge_bg, fg=badge_fg)
        self.detail_title_var.set(f"{record.folder_name}\n{record.display_name}")
        self.render_detail_text(record)
        self.set_text(self.next_text, "\n".join(f"- {item}" for item in record.next_items))
        self.set_detail_buttons_state(True)

    def set_detail_buttons_state(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        self.open_folder_button.configure(state=state)
        record = self.selected_record if enabled else None
        self.open_readme_button.configure(state=state if record and (record.folder_path / README_NAME).exists() else "disabled")
        self.open_production_button.configure(state=state if record and normalize_http_url(record.production_url) else "disabled")
        self.open_health_button.configure(state=state if record and normalize_http_url(record.health_url) else "disabled")
        self.open_github_button.configure(state=state if record and normalize_http_url(record.github_url) else "disabled")
        self.open_cloudflare_button.configure(state=state if record and normalize_http_url(record.cloudflare_url) else "disabled")

    def set_text(self, widget: tk.Text, value: str) -> None:
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("end", value)
        widget.configure(state="disabled")

    def render_detail_text(self, record: SiteRecord) -> None:
        self.detail_text.configure(state="normal")
        self.detail_text.delete("1.0", "end")
        self.detail_link_counter = 0
        readme_path = record.folder_path / README_NAME
        domain_url = url_from_domain(record.domain)
        production_url = normalize_http_url(record.production_url)
        health_url = normalize_http_url(record.health_url)
        github_url = normalize_http_url(record.github_url)
        cloudflare_url = normalize_http_url(record.cloudflare_url)
        self.insert_detail_field(UI_TEXT["label_folder_path"], str(record.folder_path), lambda: self.open_path(record.folder_path))
        self.insert_detail_field(
            UI_TEXT["label_readme_path"],
            str(readme_path) if readme_path.exists() else UI_TEXT["value_unset"],
            lambda: self.open_path(readme_path) if readme_path.exists() else None,
        )
        self.insert_detail_field(UI_TEXT["label_domain"], record.domain or UI_TEXT["value_unset"], (lambda url=domain_url: self.open_url(url)) if domain_url else None)
        self.insert_detail_field(UI_TEXT["label_production_url"], record.production_url or UI_TEXT["value_unset"], (lambda url=production_url: self.open_url(url)) if production_url else None)
        self.insert_detail_field(UI_TEXT["label_health_url"], record.health_url or UI_TEXT["value_unset"], (lambda url=health_url: self.open_url(url)) if health_url else None)
        self.insert_detail_field(UI_TEXT["label_repo_url"], record.github_url or UI_TEXT["value_unset"], (lambda url=github_url: self.open_url(url)) if github_url else None)
        self.insert_detail_field(UI_TEXT["label_cloudflare_project"], record.cloudflare_project or UI_TEXT["value_unset"], (lambda url=cloudflare_url: self.open_url(url)) if cloudflare_url else None)
        self.insert_detail_field(UI_TEXT["label_cloudflare_url"], record.cloudflare_url or UI_TEXT["value_unset"], (lambda url=cloudflare_url: self.open_url(url)) if cloudflare_url else None)
        self.detail_text.insert("end", "\n")
        self.detail_text.insert("end", self.build_detail_text(record))
        self.detail_text.configure(state="disabled")

    def insert_detail_field(self, label: str, value: str, action=None) -> None:
        self.detail_text.insert("end", f"{label}: ")
        start = self.detail_text.index("end")
        self.detail_text.insert("end", value)
        end = self.detail_text.index("end")
        if action is not None and value != UI_TEXT["value_unset"]:
            tag = f"detail_link_{self.detail_link_counter}"
            self.detail_link_counter += 1
            self.detail_text.tag_add(tag, start, end)
            self.detail_text.tag_configure(tag, foreground=THEME["accent_hover"])
            self.detail_text.tag_bind(tag, "<Button-1>", lambda _event, callback=action: callback())
            self.detail_text.tag_bind(tag, "<Enter>", lambda _event: self.detail_text.configure(cursor="hand2"))
            self.detail_text.tag_bind(tag, "<Leave>", lambda _event: self.detail_text.configure(cursor=""))
        self.detail_text.insert("end", "\n")

    def build_detail_text(self, record: SiteRecord) -> str:
        lines = [
            f"{UI_TEXT['label_detection_score']}: {record.detection_score}",
            f"{UI_TEXT['label_candidate_type']}: {UI_TEXT.get(f'candidate_type_{record.candidate_type}', record.candidate_type)}",
            f"{UI_TEXT['label_site_type']}: {record.site_type or UI_TEXT['value_unset']}",
            f"{UI_TEXT['label_status']}: {record.status_value or UI_TEXT['value_unset']}",
            f"{UI_TEXT['label_show_dashboard']}: {bool_label(record.show_on_dashboard)}",
            "",
            UI_TEXT["label_detection_basis"],
            *[f"- {reason}" for reason in record.detection_reasons],
            "",
            UI_TEXT["label_status_reasons"],
            *[f"- {reason}" for reason in record.status_reasons],
            "",
            UI_TEXT["label_api_reasons"],
            *[f"- {reason}" for reason in record.api_reasons],
            "",
            UI_TEXT["label_git_reasons"],
            *[f"- {reason}" for reason in record.git_reasons],
            "",
            UI_TEXT["label_next_reasons"],
            *[f"- {reason}" for reason in record.next_action_reasons],
            "",
            UI_TEXT["label_readme"],
            f"- {UI_TEXT['label_has_readme']}: {bool_label(record.files.has_readme)}",
            f"- {UI_TEXT['label_has_meta']}: {bool_label(record.files.has_meta)}",
        ]
        if record.files.meta_error and not record.files.has_meta:
            lines.append(f"- {record.files.meta_error}")
        lines.extend(
            [
                "",
                UI_TEXT["label_git"],
                f"- {UI_TEXT['label_has_git']}: {bool_label(record.git.is_repo)}",
                f"- {UI_TEXT['label_branch']}: {record.git.branch or UI_TEXT['value_unset']}",
                f"- {UI_TEXT['label_commit']}: {record.git.commit or UI_TEXT['value_unset']}",
                f"- {UI_TEXT['label_dirty_count']}: {record.git.dirty_count}",
                f"- {UI_TEXT['label_untracked_count']}: {record.git.untracked_count}",
                f"- {UI_TEXT['label_ahead']}: {record.git.ahead}",
                f"- {UI_TEXT['label_behind']}: {record.git.behind}",
                "",
                UI_TEXT["label_cloudflare"],
                f"- {UI_TEXT['label_has_package']}: {bool_label(record.files.has_package)}",
                f"- {UI_TEXT['label_has_wrangler']}: {bool_label(record.files.has_wrangler)}",
                f"- {UI_TEXT['label_has_public']}: {bool_label(record.files.has_public)}",
                f"- {UI_TEXT['label_has_functions']}: {bool_label(record.files.has_functions)}",
                f"- {UI_TEXT['label_has_routes']}: {bool_label(record.files.has_routes)}",
                f"- {UI_TEXT['label_has_gitignore']}: {bool_label(record.files.has_gitignore)}",
                f"- {UI_TEXT['label_node_modules_ignored']}: {bool_label(record.files.node_modules_ignored)}",
                f"- {UI_TEXT['label_pages_output']}: {bool_label(record.cloudflare.has_pages_output)}",
                f"- {UI_TEXT['label_pages_like']}: {bool_label(record.cloudflare.likely_pages)}",
                f"- {UI_TEXT['label_functions_api']}: {bool_label(record.cloudflare.has_functions_api)}",
                f"- {UI_TEXT['label_health_file']}: {bool_label(record.cloudflare.has_health_file)}",
                f"- {UI_TEXT['label_routes_api']}: {bool_label(record.cloudflare.has_routes_api)}",
                "",
                UI_TEXT["label_api"],
                f"- {UI_TEXT['label_openai_env']}: {bool_label(record.api.has_openai_env_ref)}",
                f"- {UI_TEXT['label_openai_key']}: {bool_label(record.api.has_hardcoded_key_suspect)}",
                f"- {UI_TEXT['label_openai_front']}: {bool_label(record.api.has_frontend_openai_direct)}",
                f"- {UI_TEXT['label_openai_functions']}: {bool_label(record.api.has_functions_openai)}",
                f"- {UI_TEXT['label_openai_env_design']}: {bool_label(record.api.has_env_design)}",
                "",
                UI_TEXT["label_openai_safety"],
                f"- {UI_TEXT['value_direct_key']}: {bool_label(record.api.has_hardcoded_key_suspect)}",
                f"- {UI_TEXT['value_front_direct']}: {bool_label(record.api.has_frontend_openai_direct)}",
                "",
                UI_TEXT["label_missing"],
            ]
        )
        lines.extend(f"- {item}" for item in record.missing_items)
        return "\n".join(lines)

    def open_selected_folder(self) -> None:
        if self.selected_record:
            self.open_path(self.selected_record.folder_path)

    def open_selected_readme(self) -> None:
        if self.selected_record:
            self.open_path(self.selected_record.folder_path / README_NAME)

    def open_selected_production(self) -> None:
        if self.selected_record:
            self.open_url(self.selected_record.production_url)

    def open_selected_health(self) -> None:
        if self.selected_record:
            self.open_url(self.selected_record.health_url)

    def open_selected_github(self) -> None:
        if self.selected_record:
            self.open_url(self.selected_record.github_url)

    def open_selected_cloudflare(self) -> None:
        if self.selected_record:
            self.open_url(self.selected_record.cloudflare_url)

    def open_url(self, url: str, quiet: bool = False) -> None:
        normalized = normalize_http_url(url)
        if not str(url or "").strip():
            if quiet:
                return
            messagebox.showinfo(UI_TEXT["dialog_notice_title"], UI_TEXT["dialog_url_missing"], parent=self.root)
            return
        if not normalized:
            if quiet:
                return
            messagebox.showinfo(UI_TEXT["dialog_notice_title"], UI_TEXT["dialog_url_invalid"], parent=self.root)
            return
        webbrowser.open(normalized)

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

    def on_close(self) -> None:
        self.stop_watchdog()
        self.root.destroy()


def run_gui() -> int:
    set_windows_app_id()
    root = tk.Tk()
    WebDashboardApp(root)
    root.mainloop()
    return 0


def run_launch_check() -> int:
    records = scan_sites(DEV_ROOT)
    site_records = [record for record in records if record.candidate_type == "site"]
    candidate_count = sum(1 for record in records if record.candidate_type == "candidate")
    component_site_count = sum(1 for record in site_records if record.folder_name.lower() in COMPONENT_DIR_NAMES)
    api_review_count = sum(1 for record in site_records if record.class_key == "api_review")
    dirty_count = sum(1 for record in site_records if record.git.has_dirty)
    git_error_count = sum(1 for record in site_records if record.git.error)
    print(
        f"{UI_TEXT['status_launch_check_ok']}: sites={len(site_records)} candidates={candidate_count} component_sites={component_site_count} "
        f"api_review={api_review_count} dirty_sites={dirty_count} git_errors={git_error_count}"
    )
    return 0


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--launch-check", action="store_true")
    args = parser.parse_args()
    if args.launch_check:
        return run_launch_check()
    return run_gui()


if __name__ == "__main__":
    raise SystemExit(main())
