# -*- coding: utf-8 -*-
from __future__ import annotations

import sys

import builtins
import json
import os
import queue
import re
import shutil
import subprocess
import threading
import tempfile
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import urlparse
from tkinter import font as tkfont
from tkinter import messagebox, ttk


APP_NAME = "DakeBOOTHアシスト"
WINDOW_TITLE = "DakeBOOTHアシスト"
COPYRIGHT = "© 2026 しまリス不動産 / Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "main_title": "BOOTH登録を補助する",
    "main_description": "booth_product.txt と booth_ready から、登録作業を静かに進めます。",
    "section_apps": "商品選択",
    "section_product": "読み取り結果",
    "section_actions": "操作",
    "section_status": "ステータス",
    "selected_app_label": "選択中商品",
    "field_product_type": "商品種別",
    "product_type_single": "単品",
    "product_type_pack": "パック",
    "button_reload": "再読み込み",
    "button_launch_chrome": "ログイン済みChromeを起動",
    "button_start_assist": "Chrome接続で入力補助",
    "button_open_ready": "booth_readyフォルダを開く",
    "button_copy_title": "商品名をコピー",
    "button_copy_description": "説明文をコピー",
    "button_copy_tags": "タグをコピー",
    "button_copy_thumbnail_path": "商品画像パスをコピー",
    "button_copy_zip_path": "zipパスをコピー",
    "button_get_chrome_url": "現在のChrome URLを取得",
    "button_save_booth_url": "BOOTH URLを保存",
    "field_title": "商品名",
    "field_price": "価格",
    "field_description": "説明文",
    "field_product_info": "商品情報",
    "field_tags": "タグ",
    "field_product_image": "商品画像",
    "field_product_image_path": "商品画像パス",
    "field_ready": "booth_ready",
    "field_zip": "zipファイル",
    "field_zip_path": "zipパス",
    "field_thumbnail": "booth_thumbnail.jpg",
    "field_category": "カテゴリ",
    "field_proxy_purchase": "代理購入サービス",
    "field_screenshot": "screenshot.webp",
    "field_github_release": "GitHub Release URL",
    "field_booth_url": "BOOTH商品URL",
    "status_label": "状態",
    "status_no_selection": "未選択",
    "status_loading": "読み込み中",
    "status_ready": "準備完了",
    "status_opening_booth": "BOOTHを開いています",
    "status_assisting": "入力補助中",
    "status_chrome_launching": "Chromeを起動しています",
    "status_chrome_ready": "Chromeを開きました。BOOTHにログインしてください",
    "status_chrome_connecting": "Chromeへ接続しています",
    "status_chrome_connected": "Chromeへ接続しました",
    "status_login_required": "BOOTHへログインしてください",
    "status_edit_page_required": "BOOTHの商品登録または編集画面を開いてください",
    "status_edit_page_found": "BOOTH編集画面を確認しました",
    "status_assist_complete": "入力補助が完了しました。内容を確認してください",
    "status_confirm": "確認してください",
    "status_error": "エラー",
    "value_unset": "未設定",
    "value_no_selection": "未選択",
    "value_yes": "あり",
    "value_no": "なし",
    "value_file_found": "あり: {name}",
    "value_file_missing": "なし",
    "value_zip_multiple": "あり: {name} ほか {count}件",
    "list_has_product": "商品情報あり: {name}",
    "list_no_product": "商品情報なし: {name}",
    "detail_no_selection": "アプリを選択してください。",
    "detail_loaded": "商品情報を読み取りました。内容を確認してからBOOTHへ進んでください。",
    "detail_missing_product": "booth_product.txt が見つかりません。必要な項目は手動で入力してください。",
    "detail_no_apps": "01_apps 配下にアプリフォルダが見つかりません。",
    "detail_copied": "{label}をコピーしました。",
    "detail_copy_empty": "{label}が未設定です。booth_product.txt を確認してください。",
    "detail_ready_missing": "booth_ready フォルダが見つかりません。",
    "detail_opened_ready": "booth_ready フォルダを開きました。",
    "detail_booth_url_fetched": "ChromeからBOOTH URLを取得しました。",
    "detail_booth_url_saved": "BOOTH URLを保存しました。保存先: {targets}",
    "detail_chrome_launched": "Chromeを開きました。BOOTHへログインし、商品登録または編集画面を開いてください。",
    "detail_chrome_connected": "ログイン済みChromeへ接続しました。開いているBOOTH編集画面を確認しています。",
    "detail_login_required": "Chrome上でBOOTHへログインしてから、もう一度実行してください。",
    "detail_edit_page_required": "ChromeでBOOTHの商品登録または編集画面を開いてから、もう一度実行してください。",
    "detail_bad_edit_page": "404ページを検出しました。Chromeで商品編集画面を開き直してください。",
    "detail_used_page": "使用ページ:\n{url}",
    "detail_assist_complete": "公開ボタンは押していません。内容を確認してください。",
    "assist_filled": "{label}: 入力しました",
    "assist_file_set": "{label}: 入力しました ({name})",

    "assist_tags_filled": "{label}: {count}件入力",
    "assist_manual": "{label}: 手動で入力してください",
    "assist_manual_set": "{label}: 手動で設定してください",
    "assist_path": "{label}: {path}",
    "assist_path_missing": "{label}: 未設定",
    "assist_ready_folder": "booth_readyフォルダ: 開いて確認してください",
    "assist_ready_folder_missing": "booth_readyフォルダ: 見つかりません",
    "assist_missing_value": "{label}: 未設定のためスキップしました",
    "assist_publish_guard": "公開ボタン・保存ボタンは押していません。内容を確認し、公開/保存判断は人間が行ってください。",
    "assist_publish_guard_result": "公開ボタン・保存ボタンは押していません",
    "assist_summary_template": "使用ページ:\n{url}\n\nページ情報:\ntitle: {title}\ninput数: {input_count}\ntextarea数: {textarea_count}\nfile input数: {file_count}\n\nfile input診断:\n{file_input_details}\n\n入力結果:\n{results}",
    "log_page_diagnostics": "使用URL={url} title={title} input数={input_count} textarea数={textarea_count} file input数={file_count}",
    "log_file_input_detail": "file input {index}: accept={accept} name={name} id={id} aria-label={aria_label} nearby={nearby}",
    "log_file_input_try": "set_input_files試行: {label} index={index} path={path}",
    "log_file_input_success": "set_input_files成功: {label} index={index}",
    "log_file_input_failure": "set_input_files失敗: {label} index={index} error={error}",
    "log_file_input_manual": "初期DOMに安全なfile inputがないため手動案内: {label}",
    "log_category_diagnostics": "カテゴリ候補: label={label_count} select={select_count} text={text_count}",
    "log_proxy_purchase_diagnostics": "代理購入サービス候補: text={text_count}",
    "file_input_detail_line": "- {index}: accept={accept} name={name} id={id} aria-label={aria_label} nearby={nearby}",
    "file_input_details_empty": "なし",
    "result_bullet": "- {line}",
    "dialog_error_title": "エラー",
    "dialog_notice_title": "確認してください",
    "dialog_open_folder_error": "フォルダを開けませんでした。\n\n{error}",
    "dialog_no_ready_folder": "booth_ready フォルダが見つかりません。\n\n{path}",
    "dialog_playwright_busy": "Chrome接続の入力補助中です。完了してから、もう一度お試しください。",
    "dialog_playwright_setup": "Playwright Pythonを利用する準備がまだ完了していません。\n\n初回のみ次を実行してください。\npython -m pip install -r requirements.txt\n\n{error}",
    "dialog_playwright_error": "Chrome接続でBOOTH画面の操作中に止まりました。\n画面仕様が変わっている可能性があります。コピー補助として使い、手動で入力してください。\n\n{error}",
    "dialog_chrome_not_found": "Chromeが見つかりませんでした。Chromeをインストールするか、以下を手動で実行してください。\n\n{command}",
    "dialog_chrome_launch_error": "Chromeを起動できませんでした。以下を手動で実行してください。\n\n{command}\n\n{error}",
    "dialog_chrome_connect_error": "ログイン済みChromeへ接続できませんでした。\n先に「ログイン済みChromeを起動」ボタンを押し、BOOTHへログインしてからもう一度実行してください。\n\n{error}",
    "dialog_booth_url_empty": "BOOTH URLが空です。URLを入力するか、現在のChrome URLを取得してください。",
    "dialog_booth_url_invalid": "BOOTHの商品URLまたは編集URLではありません。\nbooth.pm または manage.booth.pm の商品URLを指定してください。\n\n{url}",
    "dialog_booth_url_confirm": "以下のBOOTH URLを保存します。\n\n{url}\n\n保存先:\n{targets}\n\nよろしいですか？",
    "dialog_booth_url_saved": "BOOTH URLを保存しました。\n\n保存先:\n{targets}",
    "dialog_booth_url_save_error": "BOOTH URLを保存できませんでした。\n\n{error}",
    "dialog_assist_summary_title": "入力補助の結果",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_separator": " / ",
    "footer_note": "公開前の最後の判断は人間が行います。",
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
    "warning": "#B54708",
    "warning_bg": "#FFF7E6",
    "error": "#C92A2A",
    "error_bg": "#FDECEC",
}

STATUS_THEME = {
    "status_no_selection": (THEME["subtle"], THEME["muted"]),
    "status_loading": (THEME["warning_bg"], THEME["warning"]),
    "status_ready": (THEME["success_bg"], THEME["success"]),
    "status_opening_booth": ("#EAF2FF", THEME["accent"]),
    "status_assisting": ("#EAF2FF", THEME["accent"]),
    "status_chrome_launching": ("#EAF2FF", THEME["accent"]),
    "status_chrome_ready": (THEME["warning_bg"], THEME["warning"]),
    "status_chrome_connecting": ("#EAF2FF", THEME["accent"]),
    "status_chrome_connected": ("#EAF2FF", THEME["accent"]),
    "status_login_required": (THEME["warning_bg"], THEME["warning"]),
    "status_edit_page_required": (THEME["warning_bg"], THEME["warning"]),
    "status_edit_page_found": ("#EAF2FF", THEME["accent"]),
    "status_assist_complete": (THEME["success_bg"], THEME["success"]),
    "status_confirm": (THEME["warning_bg"], THEME["warning"]),
    "status_error": (THEME["error_bg"], THEME["error"]),
}

WINDOW_SIZE = "1040x760"
WINDOW_MIN_SIZE = (960, 700)
CONFIG_FILE_NAME = "dake_booth_assist_config.json"
CHROME_REMOTE_DEBUGGING_PORT = 9222
CHROME_CDP_URL = f"http://127.0.0.1:{CHROME_REMOTE_DEBUGGING_PORT}"
CHROME_PROFILE_PARENT_NAME = "DakeBOOTH_Assist"
CHROME_PROFILE_DIR_NAME = "chrome_profile"
PRODUCT_FILE_NAME = "booth_product.txt"
READY_DIR_NAME = "booth_ready"
PACKS_DIR_NAME = "04_packs"
PACK_READY_DIR_NAME = "pack_ready"
SHIMARISU_PACK_ROOT_NAME = "SHIMARISU"
SHIMARISU_PACK_ENTRY_NAME = "SHIMARISU Pack v1.0"
SHIMARISU_PACK_IMAGE_ORDER = (
    "booth_thumbnail.jpg",
    "01_start.jpg",
    "02_checking.jpg",
    "03_decision.jpg",
    "screenshot.jpg",
)
PRODUCT_TYPE_SINGLE = "single"
PRODUCT_TYPE_PACK = "pack"
THUMBNAIL_NAME = "booth_thumbnail.jpg"
SCREENSHOT_NAME = "screenshot.webp"
BOOTH_ADMIN_URL = "https://manage.booth.pm/items"
BOOTH_EDIT_URL_HINT = "https://manage.booth.pm/items/数字/edit"
QUEUE_POLL_MS = 100


@dataclass(frozen=True)
class AppEntry:
    name: str
    path: Path
    has_product: bool
    product_type: str = PRODUCT_TYPE_SINGLE



@dataclass(frozen=True)
class ProductInfo:
    app_name: str
    product_type: str
    app_dir: Path
    product_path: Path
    product_source: str
    title: str
    price: str
    description: str
    tags: str
    url: str
    github_release: str
    booth_ready_dir: Path
    booth_ready_exists: bool
    zip_files: tuple[Path, ...]
    selected_zip: Path | None
    thumbnail_path: Path | None
    image_paths: tuple[Path, ...]
    screenshot_path: Path | None


@dataclass(frozen=True)
class FileInputInfo:
    index: int
    accept: str
    name: str
    input_id: str
    aria_label: str
    nearby: str


WorkerEvent = tuple[str, str]


FIELD_ALIASES = {
    "商品名": "title",
    "タイトル": "title",
    "name": "title",
    "title": "title",
    "価格": "price",
    "価格案": "price",
    "販売価格": "price",
    "price": "price",
    "概要": "description",
    "できること": "description",
    "説明": "description",
    "説明文": "description",
    "商品説明": "description",
    "商品紹介文": "description",
    "紹介文": "description",
    "description": "description",
    "body": "description",
    "タグ": "tags",
    "tag": "tags",
    "tags": "tags",
    "url": "url",
    "githubrelease": "github_release",
    "githubreleaseurl": "github_release",
    "github_release": "github_release",
    "releaseurl": "github_release",
    "release_url": "github_release",
    "ファイル": "zip_path",
    "作品ファイル": "zip_path",
    "配布zip": "zip_path",
    "zip": "zip_path",
    "zipfile": "zip_path",
    "zippath": "zip_path",
    "zip_path": "zip_path",
    "商品画像": "thumbnail_path",
    "booth商品画像": "thumbnail_path",
    "画像": "thumbnail_path",
    "thumbnail": "thumbnail_path",
    "thumbnailpath": "thumbnail_path",
    "thumbnail_path": "thumbnail_path",
    "boothurl": "url",
}

KEY_VALUE_RE = re.compile(r"^\s*([^:=]+?)\s*[:=]\s*(.*)$")
BOOTH_ITEM_ID_RE = re.compile(r"/items/(\d+)")
BOOTH_URL_SAVE_KEYS = {"boothurl", "url"}
GITHUB_RELEASE_RE = re.compile(r"https?://github\.com/[^\s)]+/releases/[^\s)]+", re.IGNORECASE)
DAKE_META_RE = re.compile(r"##\s*DAKE_META\b.*?```(?:json)?\s*(\{.*?\})\s*```", re.IGNORECASE | re.DOTALL)
PACK_META_RE = re.compile(r"##\s*PACK_META\b.*?```(?:json)?\s*(\{.*?\})\s*```", re.IGNORECASE | re.DOTALL)
STOP_SECTION_LABELS = {
    "使い方",
    "注意事項",
    "DAKEシリーズ表記",
    "対象",
    "配布物",
    "免責",
    "コピーライト",
}


TITLE_LOCATORS = (
    ("label", "商品名"),
    ("label", "タイトル"),
    ("placeholder", "商品名"),
    ("placeholder", "タイトル"),
    ("css", "input[name*='title' i]"),
    ("css", "input[name*='name' i]"),
)
DESCRIPTION_LOCATORS = (
    ("label", "説明文"),
    ("label", "商品説明"),
    ("label", "説明"),
    ("placeholder", "説明"),
    ("css", "textarea[name*='description' i]"),
    ("css", "textarea[name*='body' i]"),
    ("css", "textarea"),
)
PRICE_LOCATORS = (
    ("label", "価格"),
    ("label", "販売価格"),
    ("placeholder", "価格"),
    ("css", "input[name*='price' i]"),
    ("css", "input[type='number']"),
)
TAG_LOCATORS = (
    ("label", "タグ"),
    ("placeholder", "タグ"),
    ("css", "input[name*='tag' i]"),
    ("css", "textarea[name*='tag' i]"),
)

def get_source_dir() -> Path:
    return Path(__file__).resolve().parent


def get_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return get_source_dir()


def get_apps_root() -> Path:
    source_dir = get_source_dir()
    base_dir = get_base_dir()
    candidates = [
        source_dir.parent,
        base_dir.parent,
        base_dir.parent.parent,
    ]
    for candidate in candidates:
        if candidate.name == "01_apps" and candidate.exists():
            return candidate
    return source_dir.parent


def get_series_root() -> Path:
    return get_apps_root().parent


def get_packs_root() -> Path:
    return get_series_root() / PACKS_DIR_NAME


def get_shimarisu_pack_root() -> Path:
    return get_series_root().parent / SHIMARISU_PACK_ROOT_NAME


def get_external_pack_dirs() -> tuple[Path, ...]:
    candidates = (get_shimarisu_pack_root(),)
    return tuple(
        candidate
        for candidate in candidates
        if (candidate / READY_DIR_NAME / PRODUCT_FILE_NAME).is_file()
    )


def is_external_pack_dir(product_dir: Path) -> bool:
    try:
        resolved = product_dir.resolve()
    except OSError:
        resolved = product_dir
    for candidate in get_external_pack_dirs():
        try:
            if resolved == candidate.resolve():
                return True
        except OSError:
            if product_dir == candidate:
                return True
    return False


def is_pack_dir(product_dir: Path) -> bool:
    if is_external_pack_dir(product_dir):
        return True
    if product_dir.parent.name == PACKS_DIR_NAME:
        return True
    readme_path = product_dir / "README.md"
    if readme_path.exists() and PACK_META_RE.search(read_text_safely(readme_path)):
        return True
    return False


def product_type_for(product_dir: Path) -> str:
    return PRODUCT_TYPE_PACK if is_pack_dir(product_dir) else PRODUCT_TYPE_SINGLE


def ready_dir_name_for(product_dir: Path) -> str:
    if is_external_pack_dir(product_dir):
        return READY_DIR_NAME
    return PACK_READY_DIR_NAME if is_pack_dir(product_dir) else READY_DIR_NAME


def product_type_label(product_type: str) -> str:
    return UI_TEXT["product_type_pack"] if product_type == PRODUCT_TYPE_PACK else UI_TEXT["product_type_single"]


def get_config_path() -> Path:
    return get_base_dir() / CONFIG_FILE_NAME


def get_chrome_profile_dir() -> Path:
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        return Path(local_app_data) / CHROME_PROFILE_PARENT_NAME / CHROME_PROFILE_DIR_NAME
    return get_base_dir() / CHROME_PROFILE_DIR_NAME


def build_manual_chrome_command() -> str:
    profile_arg = f"%LOCALAPPDATA%\\{CHROME_PROFILE_PARENT_NAME}\\{CHROME_PROFILE_DIR_NAME}"
    return (
        f"chrome.exe --remote-debugging-port={CHROME_REMOTE_DEBUGGING_PORT} "
        f"--user-data-dir=\"{profile_arg}\" {BOOTH_ADMIN_URL}"
    )


def find_chrome_executable() -> str | None:
    candidates = [
        Path(os.environ.get("PROGRAMFILES", "")) / "Google" / "Chrome" / "Application" / "chrome.exe",
        Path(os.environ.get("PROGRAMFILES(X86)", "")) / "Google" / "Chrome" / "Application" / "chrome.exe",
        Path("C:/Program Files/Google/Chrome/Application/chrome.exe"),
        Path("C:/Program Files (x86)/Google/Chrome/Application/chrome.exe"),
    ]
    for candidate in candidates:
        try:
            if candidate.is_file():
                return str(candidate)
        except OSError:
            continue

    for command_name in ("chrome.exe", "chrome"):
        found = shutil.which(command_name)
        if found:
            return found
    return None


def build_chrome_launch_args(chrome_path: str) -> list[str]:
    return [
        chrome_path,
        f"--remote-debugging-port={CHROME_REMOTE_DEBUGGING_PORT}",
        f"--user-data-dir={get_chrome_profile_dir()}",
        BOOTH_ADMIN_URL,
    ]


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


def normalize_key(value: str) -> str:
    normalized = value.strip().lower()
    normalized = normalized.replace("＃", "#")
    normalized = re.sub(r"^[#\s]+", "", normalized)
    normalized = re.sub(r"[\s_\-:：/／]+", "", normalized)
    return normalized


def normalized_field_aliases() -> dict[str, str]:
    return {normalize_key(alias): field_name for alias, field_name in FIELD_ALIASES.items()}


def resolve_field_name(label: str) -> str | None:
    key = normalize_key(label)
    aliases = normalized_field_aliases()
    if key in aliases:
        return aliases[key]
    for alias, field_name in aliases.items():
        if alias and alias in key:
            return field_name
    return None


def is_stop_section_label(value: str) -> bool:
    key = normalize_key(value)
    return any(key == normalize_key(label) for label in STOP_SECTION_LABELS)


def resolve_section_field(value: str) -> tuple[str | None, bool]:
    field_name = normalized_field_aliases().get(normalize_key(value))
    if field_name:
        return field_name, True
    if is_stop_section_label(value):
        return None, True
    return None, False


def read_text_safely(path: Path) -> str:
    for encoding_name in ("utf-8-sig", "utf-8"):
        try:
            return path.read_text(encoding=encoding_name)
        except UnicodeDecodeError:
            continue
        except OSError:
            return ""
    try:
        return path.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return ""


def safe_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def extract_readme_meta(app_dir: Path, pattern: re.Pattern[str] | None = None) -> dict[str, object]:
    readme_path = app_dir / "README.md"
    text = read_text_safely(readme_path)
    if not text:
        return {}
    pattern = pattern or (PACK_META_RE if is_pack_dir(app_dir) else DAKE_META_RE)
    match = pattern.search(text)
    if not match:
        return {}
    try:
        loaded = json.loads(match.group(1))
    except json.JSONDecodeError:
        return {}
    return loaded if isinstance(loaded, dict) else {}


def booth_product_candidates(app_dir: Path) -> tuple[Path, ...]:
    if is_pack_dir(app_dir):
        if is_external_pack_dir(app_dir):
            return (app_dir / READY_DIR_NAME / PRODUCT_FILE_NAME,)
        return (app_dir / PACK_READY_DIR_NAME / PRODUCT_FILE_NAME,)
    return (
        app_dir / READY_DIR_NAME / PRODUCT_FILE_NAME,
        app_dir / PRODUCT_FILE_NAME,
    )


def find_booth_product_path(app_dir: Path) -> Path | None:
    for candidate in booth_product_candidates(app_dir):
        try:
            if candidate.is_file():
                return candidate
        except OSError:
            continue
    return None


def format_product_source(app_dir: Path, product_path: Path | None) -> str:
    if product_path is None:
        return ""
    try:
        return product_path.relative_to(app_dir).as_posix()
    except ValueError:
        return str(product_path)


def clean_single_line(value: str) -> str:
    for line in value.splitlines():
        cleaned = line.strip().strip("`")
        cleaned = re.sub(r"^\s*[-*・]\s*", "", cleaned)
        if cleaned:
            return cleaned
    return ""


def clean_multiline(value: str) -> str:
    return value.strip()


def split_tags(value: str) -> list[str]:
    tags: list[str] = []
    for raw_line in value.replace("、", ",").splitlines():
        line = re.sub(r"^\s*[-*・]\s*", "", raw_line.strip())
        for part in line.split(","):
            tag = part.strip()
            if tag:
                tags.append(tag)
    return list(dict.fromkeys(tags))


def clean_tags(value: str) -> str:
    return ", ".join(split_tags(value))


def clean_url(value: str) -> str:
    return clean_single_line(value).rstrip("、。，,)")

def ensure_url_scheme(value: str) -> str:
    cleaned = value.strip()
    if cleaned and not re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://", cleaned):
        return "https://" + cleaned
    return cleaned


def booth_url_parts(value: str):
    try:
        return urlparse(ensure_url_scheme(value))
    except Exception:
        return urlparse("")


def is_booth_host(host: str | None) -> bool:
    if not host:
        return False
    normalized = host.lower()
    return normalized == "booth.pm" or normalized.endswith(".booth.pm")


def is_public_booth_shop_host(host: str | None) -> bool:
    return is_booth_host(host) and host.lower() != "manage.booth.pm" if host else False


def extract_booth_item_id(value: str) -> str:
    match = BOOTH_ITEM_ID_RE.search(booth_url_parts(value).path)
    return match.group(1) if match else ""


def clean_booth_url(value: str) -> str:
    parts = booth_url_parts(value)
    if not parts.scheme or not parts.netloc:
        return ""
    return f"{parts.scheme}://{parts.netloc}{parts.path}".rstrip("/")


def is_valid_booth_save_url(value: str) -> bool:
    parts = booth_url_parts(value)
    return is_booth_host(parts.hostname) and bool(extract_booth_item_id(value))


def shop_host_from_url(value: str) -> str:
    parts = booth_url_parts(value)
    host = parts.hostname or ""
    return host if is_public_booth_shop_host(host) else ""


def collect_shop_hosts_from_page(page) -> list[str]:
    try:
        hrefs = page.evaluate("""
() => Array.from(document.links || [])
  .map((link) => link.href || "")
  .filter(Boolean)
  .slice(0, 300)
""")
    except Exception:
        return []
    hosts: list[str] = []
    for href in hrefs if isinstance(hrefs, list) else []:
        host = shop_host_from_url(str(href))
        if host and host not in hosts:
            hosts.append(host)
    return hosts


def infer_booth_shop_host(product: ProductInfo | None = None, browser=None, page=None) -> str:
    if product is not None:
        host = shop_host_from_url(product.url)
        if host:
            return host

    if browser is not None:
        for browser_page in collect_browser_pages(browser):
            host = shop_host_from_url(safe_page_url(browser_page))
            if host:
                return host

    if page is not None:
        hosts = collect_shop_hosts_from_page(page)
        if hosts:
            return hosts[0]
    return ""


def normalize_booth_url_for_save(value: str, product: ProductInfo | None = None, browser=None, page=None) -> str:
    cleaned = clean_booth_url(value)
    if not cleaned:
        return ""
    parts = booth_url_parts(cleaned)
    if not is_booth_host(parts.hostname):
        return ""

    item_id = extract_booth_item_id(cleaned)
    if not item_id:
        return cleaned

    if (parts.hostname or "").lower() == "manage.booth.pm" and "/edit" in parts.path:
        shop_host = infer_booth_shop_host(product, browser, page)
        if shop_host:
            return f"https://{shop_host}/items/{item_id}"
    return cleaned


def is_booth_url_save_label(label: str) -> bool:
    return normalize_key(label) in BOOTH_URL_SAVE_KEYS


def update_key_value_url_line(line: str, url: str) -> str:
    match = re.match(r"^(\s*)([^:=]+?)(\s*)([:=]).*$", line)
    if not match:
        return f"BOOTH URL: {url}"
    leading, key, spacing, separator = match.groups()
    value_spacing = "" if separator == "=" else " "
    return f"{leading}{key.rstrip()}{spacing}{separator}{value_spacing}{url}"


def choose_booth_url_append_line(text: str, url: str) -> str:
    for line in text.splitlines():
        if "=" in line and KEY_VALUE_RE.match(line):
            return f"BOOTH_URL={url}"
    return f"BOOTH URL: {url}"


def is_structured_key_value_line(line: str) -> bool:
    stripped = line.strip()
    if re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://", stripped):
        return False
    return bool(KEY_VALUE_RE.match(stripped))


def update_booth_url_text(text: str, url: str) -> str:
    lines = text.splitlines()
    output: list[str] = []
    updated = False
    index = 0

    while index < len(lines):
        line = lines[index]
        stripped = line.strip()

        if stripped.startswith("#") and is_booth_url_save_label(stripped.lstrip("#").strip()):
            output.append(line)
            output.append(url)
            updated = True
            index += 1
            while index < len(lines):
                next_stripped = lines[index].strip()
                if next_stripped.startswith("#") or is_structured_key_value_line(next_stripped):
                    break
                index += 1
            continue

        match = KEY_VALUE_RE.match(line)
        if match and is_booth_url_save_label(match.group(1)):
            output.append(update_key_value_url_line(line, url))
            updated = True
            index += 1
            continue

        output.append(line)
        index += 1

    if not updated:
        if output and output[-1].strip():
            output.append("")
        output.append(choose_booth_url_append_line(text, url))

    return "\n".join(output).rstrip() + "\n"


def booth_url_save_targets(app_dir: Path) -> tuple[Path, ...]:
    if is_external_pack_dir(app_dir):
        return (app_dir / READY_DIR_NAME / PRODUCT_FILE_NAME,)
    if is_pack_dir(app_dir):
        targets = [app_dir / PACK_READY_DIR_NAME / PRODUCT_FILE_NAME]
        readme_path = app_dir / "README.md"
        if readme_path.exists():
            targets.append(readme_path)
        return tuple(targets)
    primary = app_dir / READY_DIR_NAME / PRODUCT_FILE_NAME
    root_product = app_dir / PRODUCT_FILE_NAME
    targets = [primary]
    if root_product.exists() and root_product != primary:
        targets.append(root_product)
    return tuple(targets)


def update_pack_meta_booth_url(text: str, url: str) -> str:
    match = PACK_META_RE.search(text)
    if not match:
        return text
    try:
        meta = json.loads(match.group(1))
    except json.JSONDecodeError:
        return text
    if not isinstance(meta, dict):
        return text
    meta["booth_url"] = url
    replacement = json.dumps(meta, ensure_ascii=False, indent=2)
    return text[: match.start(1)] + replacement + text[match.end(1) :]


def format_save_target(app_dir: Path, target: Path) -> str:
    try:
        return target.relative_to(app_dir).as_posix()
    except ValueError:
        return str(target)


def format_save_targets(app_dir: Path, targets: tuple[Path, ...]) -> str:
    return "\n".join(format_save_target(app_dir, target) for target in targets)


def save_booth_url_to_products(app_dir: Path, url: str) -> tuple[Path, ...]:
    targets = booth_url_save_targets(app_dir)
    for target in targets:
        target.parent.mkdir(parents=True, exist_ok=True)
        current_text = read_text_safely(target) if target.exists() else ""
        if is_pack_dir(app_dir) and target.name == "README.md":
            target.write_text(update_pack_meta_booth_url(current_text, url), encoding="utf-8")
        else:
            target.write_text(update_booth_url_text(current_text, url), encoding="utf-8")
    return targets


def parse_booth_product(path: Path) -> dict[str, str]:
    if not path.exists():
        return {}

    fields: dict[str, str] = {}
    buffers: dict[str, list[str]] = {}
    current_field: str | None = None
    text = read_text_safely(path)

    for raw_line in text.splitlines():
        line = raw_line.rstrip()
        stripped = line.strip()

        if stripped.startswith("#"):
            field_name, _is_section = resolve_section_field(stripped.lstrip("#").strip())
            current_field = field_name
            if field_name:
                buffers.setdefault(field_name, [])
            continue

        key_value = KEY_VALUE_RE.match(line)
        if key_value:
            key, value = key_value.groups()
            field_name = resolve_field_name(key)
            if field_name:
                fields[field_name] = value.strip()
                current_field = field_name if not value.strip() else None
                buffers.setdefault(field_name, [])
                continue

        if current_field != "tags":
            field_name, is_section = resolve_section_field(stripped)
            if is_section:
                current_field = field_name
                if field_name:
                    buffers.setdefault(field_name, [])
                continue

        if current_field:
            buffers.setdefault(current_field, []).append(line)

    for field_name, lines in buffers.items():
        if fields.get(field_name, "").strip():
            continue
        if any(line.strip() for line in lines):
            fields[field_name] = "\n".join(lines).strip()

    if "github_release" not in fields:
        release_match = GITHUB_RELEASE_RE.search(text)
        if release_match:
            fields["github_release"] = release_match.group(0).rstrip("、。，,)")

    if fields.get("url") and "github.com" in fields["url"].lower() and "github_release" not in fields:
        fields["github_release"] = fields["url"]

    return {
        "title": clean_single_line(fields.get("title", "")),
        "price": clean_single_line(fields.get("price", "")),
        "description": clean_multiline(fields.get("description", "")),
        "tags": clean_tags(fields.get("tags", "")),
        "url": clean_url(fields.get("url", "")),
        "github_release": clean_url(fields.get("github_release", "")),
        "zip_path": clean_single_line(fields.get("zip_path", "")),
        "thumbnail_path": clean_single_line(fields.get("thumbnail_path", "")),
    }


def resolve_product_relative_path(app_dir: Path, value: str) -> Path | None:
    if not value:
        return None
    candidate = Path(value.strip().strip('"'))
    candidates = [candidate] if candidate.is_absolute() else [app_dir / candidate]
    for path in candidates:
        if path.exists():
            return path.resolve()
    return None


def find_file_case_insensitive(directory: Path, name: str) -> Path | None:
    if not directory.exists():
        return None
    target = name.lower()
    try:
        for child in directory.iterdir():
            if child.is_file() and child.name.lower() == target:
                return child
    except OSError:
        return None
    return None


def find_zip_files(ready_dir: Path) -> tuple[Path, ...]:
    if not ready_dir.exists():
        return ()
    try:
        return tuple(sorted((path for path in ready_dir.iterdir() if path.is_file() and path.suffix.lower() == ".zip"), key=lambda item: item.name.lower()))
    except OSError:
        return ()


def shimarisu_pack_defaults() -> dict[str, str]:
    return {
        "title": "しまりすくん 実務判断Pack",
        "price": "3,000円",
        "description": (
            "PDFや画像を投げると、内容を見て、次にできることを静かに出す実務判断Pack。\n\n"
            "PDF結合、PDF圧縮、PDF画像化、画像PDF化、画像リサイズなどを、しまりすくん経由で進められます。"
        ),
        "tags": "Windows, PDF, 画像変換, 業務効率化, 実務ツール, DAKE, しまりすくん, 不動産実務, 書類整理, デスクトップアプリ",
        "zip_path": "booth_ready/SHIMARISU_Pack.zip",
        "thumbnail_path": "booth_ready/images/booth_thumbnail.jpg",
    }


def apply_external_pack_defaults(app_dir: Path, parsed: dict[str, str]) -> dict[str, str]:
    if not is_external_pack_dir(app_dir):
        return parsed
    merged = dict(parsed)
    defaults = shimarisu_pack_defaults()
    for key in ("title", "price", "description", "tags", "zip_path", "thumbnail_path"):
        merged[key] = defaults[key]
    return merged


def ordered_pack_image_paths(ready_dir: Path) -> tuple[Path, ...]:
    image_dir = ready_dir / "images"
    if not image_dir.exists():
        return ()
    paths: list[Path] = []
    for name in SHIMARISU_PACK_IMAGE_ORDER:
        path_item = find_file_case_insensitive(image_dir, name)
        if path_item is not None:
            paths.append(path_item)
    return tuple(paths)


def build_product_info(app_dir: Path) -> ProductInfo:
    product_type = product_type_for(app_dir)
    found_product_path = find_booth_product_path(app_dir)
    product_path = found_product_path or app_dir / PRODUCT_FILE_NAME
    parsed = {
        "title": "",
        "price": "",
        "description": "",
        "tags": "",
        "url": "",
        "github_release": "",
        "zip_path": "",
        "thumbnail_path": "",
    }
    parsed.update(parse_booth_product(product_path))
    parsed = apply_external_pack_defaults(app_dir, parsed)
    readme_meta = extract_readme_meta(app_dir)
    readme_title = safe_text(readme_meta.get("display_name")) or safe_text(readme_meta.get("site_title")) or safe_text(readme_meta.get("launcher_title"))
    readme_description = safe_text(readme_meta.get("summary")) or safe_text(readme_meta.get("site_description")) or safe_text(readme_meta.get("update_summary"))
    parsed["title"] = parsed["title"] or readme_title
    parsed["description"] = parsed["description"] or readme_description
    parsed["price"] = parsed["price"] or safe_text(readme_meta.get("price"))
    parsed["url"] = parsed["url"] or safe_text(readme_meta.get("booth_url"))
    if product_type != PRODUCT_TYPE_PACK:
        parsed["github_release"] = parsed["github_release"] or safe_text(readme_meta.get("release_url"))
    product_source = format_product_source(app_dir, found_product_path)
    ready_dir = app_dir / ready_dir_name_for(app_dir)
    ready_exists = ready_dir.exists() and ready_dir.is_dir()
    zip_files = find_zip_files(ready_dir)
    zip_from_text = resolve_product_relative_path(app_dir, parsed.get("zip_path", ""))
    selected_zip = zip_files[0] if zip_files else zip_from_text

    if product_type == PRODUCT_TYPE_PACK and not is_external_pack_dir(app_dir):
        thumbnail_path = find_file_case_insensitive(app_dir / "assets", THUMBNAIL_NAME)
        if thumbnail_path is None:
            thumbnail_path = find_file_case_insensitive(ready_dir, THUMBNAIL_NAME)
        if thumbnail_path is None:
            thumbnail_path = find_file_case_insensitive(ready_dir / "images", THUMBNAIL_NAME)
    else:
        thumbnail_path = find_file_case_insensitive(ready_dir, THUMBNAIL_NAME)
        if thumbnail_path is None:
            thumbnail_path = find_file_case_insensitive(ready_dir / "images", THUMBNAIL_NAME)
        if thumbnail_path is None:
            thumbnail_path = find_file_case_insensitive(app_dir / "assets", THUMBNAIL_NAME)
    if thumbnail_path is None:
        thumbnail_path = resolve_product_relative_path(app_dir, parsed.get("thumbnail_path", ""))
    image_paths = ordered_pack_image_paths(ready_dir)
    if not image_paths and thumbnail_path is not None:
        image_paths = (thumbnail_path,)

    screenshot_path = find_file_case_insensitive(app_dir / "assets", SCREENSHOT_NAME)
    if screenshot_path is None:
        screenshot_path = find_file_case_insensitive(ready_dir, SCREENSHOT_NAME)
    if screenshot_path is None:
        screenshot_path = find_file_case_insensitive(ready_dir / "images", "screenshot.jpg")

    return ProductInfo(
        app_name=app_dir.name,
        product_type=product_type,
        app_dir=app_dir,
        product_path=product_path,
        product_source=product_source,
        title=parsed["title"],
        price=parsed["price"],
        description=parsed["description"],
        tags=parsed["tags"],
        url=parsed["url"],
        github_release=parsed["github_release"],
        booth_ready_dir=ready_dir,
        booth_ready_exists=ready_exists,
        zip_files=zip_files,
        selected_zip=selected_zip,
        thumbnail_path=thumbnail_path,
        image_paths=image_paths,
        screenshot_path=screenshot_path,
    )


def discover_apps() -> list[AppEntry]:
    entries: list[AppEntry] = []
    for root, product_type in ((get_apps_root(), PRODUCT_TYPE_SINGLE), (get_packs_root(), PRODUCT_TYPE_PACK)):
        if not root.exists():
            continue
        try:
            for child in root.iterdir():
                if child.is_dir() and not child.name.startswith("."):
                    product_path = find_booth_product_path(child)
                    entries.append(AppEntry(child.name, child, product_path is not None, product_type))
        except OSError:
            continue
    for pack_dir in get_external_pack_dirs():
        product_path = find_booth_product_path(pack_dir)
        entries.append(AppEntry(SHIMARISU_PACK_ENTRY_NAME, pack_dir, product_path is not None, PRODUCT_TYPE_PACK))
    return sorted(entries, key=lambda entry: (entry.product_type != PRODUCT_TYPE_PACK, not entry.has_product, entry.name.lower()))


def load_config() -> dict[str, str]:
    try:
        data = json.loads(get_config_path().read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return data if isinstance(data, dict) else {}


def save_config(selected_app: Path) -> None:
    data = {"last_app_path": str(selected_app)}
    try:
        get_config_path().write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        pass


def price_for_input(price: str) -> str:
    digits = re.sub(r"[^\d]", "", price)
    return digits or price.strip()


def make_locator(page, kind: str, query: str):
    if kind == "label":
        return page.get_by_label(query, exact=False)
    if kind == "placeholder":
        return page.get_by_placeholder(query, exact=False)
    return page.locator(query)


def fill_first_available(page, label: str, value: str, locators: tuple[tuple[str, str], ...]) -> str:
    if not value.strip():
        return UI_TEXT["assist_missing_value"].format(label=label)

    for kind, query in locators:
        try:
            locator = make_locator(page, kind, query)
            count = locator.count()
        except Exception:
            continue
        for index in range(min(count, 5)):
            try:
                target = locator.nth(index)
                target.fill(value, timeout=2500)
                return UI_TEXT["assist_filled"].format(label=label)
            except Exception:
                continue
    return UI_TEXT["assist_manual"].format(label=label)


def input_tags_one_by_one(page, label: str, value: str) -> str:
    tags = split_tags(value)
    if not tags:
        return UI_TEXT["assist_missing_value"].format(label=label)

    for kind, query in TAG_LOCATORS:
        try:
            locator = make_locator(page, kind, query)
            count = locator.count()
        except Exception:
            continue
        for index in range(min(count, 5)):
            target = locator.nth(index)
            entered_count = 0
            for tag in tags:
                try:
                    try:
                        target.fill(tag, timeout=2500)
                    except Exception:
                        target.click(timeout=2500)
                        target.type(tag, delay=20, timeout=2500)
                    target.press("Enter", timeout=2500)
                    try:
                        page.wait_for_timeout(180)
                    except Exception:
                        pass
                    entered_count += 1
                except Exception:
                    continue
            if entered_count > 0:
                return UI_TEXT["assist_tags_filled"].format(label=label, count=entered_count)
    return UI_TEXT["assist_manual"].format(label=label)


def display_detail(value: str) -> str:
    stripped = (value or "").strip()
    return stripped if stripped else UI_TEXT["value_unset"]


def safe_get_attribute(locator, name: str) -> str:
    try:
        return locator.get_attribute(name, timeout=1000) or ""
    except Exception:
        return ""


def safe_nearby_text(locator) -> str:
    try:
        value = locator.evaluate(
            r"""
element => {
  const pieces = [];
  if (element.labels) {
    for (const label of element.labels) {
      pieces.push(label.innerText || label.textContent || "");
    }
  }
  let node = element.parentElement;
  for (let depth = 0; node && depth < 4; depth += 1) {
    pieces.push(node.innerText || node.textContent || "");
    node = node.parentElement;
  }
  return pieces.join(" ").replace(/\s+/g, " ").trim().slice(0, 180);
}
"""
        )
        return str(value or "")
    except Exception:
        return ""


def collect_file_input_diagnostics(page) -> tuple[FileInputInfo, ...]:
    try:
        inputs = page.locator("input[type='file']")
        count = inputs.count()
    except Exception:
        return ()

    details: list[FileInputInfo] = []
    for index in range(min(count, 12)):
        try:
            target = inputs.nth(index)
            details.append(
                FileInputInfo(
                    index=index,
                    accept=safe_get_attribute(target, "accept"),
                    name=safe_get_attribute(target, "name"),
                    input_id=safe_get_attribute(target, "id"),
                    aria_label=safe_get_attribute(target, "aria-label"),
                    nearby=safe_nearby_text(target),
                )
            )
        except Exception:
            continue
    return tuple(details)


def file_input_attr_text(info: FileInputInfo) -> str:
    return " ".join((info.accept, info.name, info.input_id, info.aria_label)).lower()


def contains_any(text: str, keywords: tuple[str, ...]) -> bool:
    return any(keyword in text for keyword in keywords)


def looks_like_image_file_input(info: FileInputInfo) -> bool:
    image_keywords = ("image", "img", "thumbnail", "thumb", "photo", "picture", "cover", "商品画像", "画像", "サムネ")
    zip_keywords = ("zip", ".zip", "application/zip", "application/x-zip", "octet", "作品ファイル", "ダウンロード", "download", "attachment", "digital", "item_file")
    attr_text = file_input_attr_text(info)
    nearby_text = info.nearby.lower()
    if contains_any(attr_text, image_keywords):
        return True
    if contains_any(attr_text, zip_keywords):
        return False
    return contains_any(nearby_text, image_keywords) and not contains_any(nearby_text, zip_keywords)


def looks_like_zip_file_input(info: FileInputInfo) -> bool:
    image_keywords = ("image", "img", "thumbnail", "thumb", "photo", "picture", "cover", "商品画像", "画像", "サムネ")
    zip_keywords = ("zip", ".zip", "application/zip", "application/x-zip", "octet", "作品ファイル", "ダウンロード", "download", "attachment", "digital", "item_file")
    attr_text = file_input_attr_text(info)
    nearby_text = info.nearby.lower()
    if contains_any(attr_text, image_keywords):
        return False
    if contains_any(attr_text, zip_keywords) or ("application" in info.accept.lower() and "image" not in info.accept.lower()):
        return True
    return contains_any(nearby_text, zip_keywords) and not contains_any(nearby_text, image_keywords)



def set_file_input_by_indexes(page, label: str, path: Path | None, indexes: list[int]) -> str:
    if path is None or not path.exists():
        return UI_TEXT["assist_missing_value"].format(label=label)
    if not indexes:
        print(UI_TEXT["log_file_input_manual"].format(label=label), flush=True)
        return UI_TEXT["assist_manual_set"].format(label=label)

    for index in indexes:
        print(UI_TEXT["log_file_input_try"].format(label=label, index=index, path=path), flush=True)
        try:
            target = page.locator("input[type='file']").nth(index)
            target.set_input_files(str(path), timeout=5000)
            print(UI_TEXT["log_file_input_success"].format(label=label, index=index), flush=True)
            return UI_TEXT["assist_file_set"].format(label=label, name=path.name)
        except Exception as exc:
            print(UI_TEXT["log_file_input_failure"].format(label=label, index=index, error=exc), flush=True)
            continue
    return UI_TEXT["assist_manual_set"].format(label=label)


def set_product_image_file(page, label: str, path: Path | None, file_inputs: tuple[FileInputInfo, ...]) -> str:
    if path is None or not path.exists():
        return UI_TEXT["assist_missing_value"].format(label=label)
    if not file_inputs:
        print(UI_TEXT["log_file_input_manual"].format(label=label), flush=True)
        return UI_TEXT["assist_manual_set"].format(label=label)

    candidates = [info.index for info in file_inputs if looks_like_image_file_input(info)]
    if not candidates:
        non_zip = [info.index for info in file_inputs if not looks_like_zip_file_input(info)]
        if len(file_inputs) > 1 and len(non_zip) == 1:
            candidates = non_zip
    return set_file_input_by_indexes(page, label, path, candidates)


def set_product_zip_file(page, label: str, path: Path | None, file_inputs: tuple[FileInputInfo, ...]) -> str:
    if path is None or not path.exists():
        return UI_TEXT["assist_missing_value"].format(label=label)
    if not file_inputs:
        print(UI_TEXT["log_file_input_manual"].format(label=label), flush=True)
        return UI_TEXT["assist_manual_set"].format(label=label)

    candidates = [info.index for info in file_inputs if looks_like_zip_file_input(info)]
    if not candidates:
        non_image = [info.index for info in file_inputs if not looks_like_image_file_input(info)]
        if len(file_inputs) > 1 and len(non_image) == 1:
            candidates = non_image
    return set_file_input_by_indexes(page, label, path, candidates)


def format_path_result(label_key: str, path: Path | None) -> str:
    label = UI_TEXT[label_key]
    if path is None:
        return UI_TEXT["assist_path_missing"].format(label=label)
    return UI_TEXT["assist_path"].format(label=label, path=path)


def format_ready_folder_result(product: ProductInfo) -> str:
    if product.booth_ready_exists:
        return UI_TEXT["assist_ready_folder"]
    return UI_TEXT["assist_ready_folder_missing"]


def category_manual_result(page) -> str:
    label_count = 0
    text_count = 0
    try:
        label_count = page.get_by_label(UI_TEXT["field_category"], exact=False).count()
    except Exception:
        pass
    try:
        text_count = page.get_by_text(UI_TEXT["field_category"], exact=False).count()
    except Exception:
        pass
    select_count = count_selector(page, "select")
    print(
        UI_TEXT["log_category_diagnostics"].format(
            label_count=label_count,
            select_count=select_count,
            text_count=text_count,
        ),
        flush=True,
    )
    return UI_TEXT["assist_manual"].format(label=UI_TEXT["field_category"])


def proxy_purchase_manual_result(page) -> str:
    text_count = 0
    try:
        text_count = page.get_by_text(UI_TEXT["field_proxy_purchase"], exact=False).count()
    except Exception:
        pass
    print(UI_TEXT["log_proxy_purchase_diagnostics"].format(text_count=text_count), flush=True)
    return UI_TEXT["assist_manual"].format(label=UI_TEXT["field_proxy_purchase"])


def format_result_lines(results: list[str]) -> str:
    return "\n".join(UI_TEXT["result_bullet"].format(line=result) for result in results)


def assist_booth_form(page, product: ProductInfo, diagnostics: dict[str, object]) -> str:
    raw_file_inputs = diagnostics.get("file_inputs", ())
    file_inputs = raw_file_inputs if isinstance(raw_file_inputs, tuple) else ()
    results = [
        fill_first_available(page, UI_TEXT["field_title"], product.title, TITLE_LOCATORS),
        fill_first_available(page, UI_TEXT["field_description"], product.description, DESCRIPTION_LOCATORS),
        fill_first_available(page, UI_TEXT["field_price"], price_for_input(product.price), PRICE_LOCATORS),
        input_tags_one_by_one(page, UI_TEXT["field_tags"], product.tags),
        set_product_image_file(page, UI_TEXT["field_product_image"], product.thumbnail_path, file_inputs),
        set_product_zip_file(page, UI_TEXT["field_zip"], product.selected_zip, file_inputs),
        format_path_result("field_product_image_path", product.thumbnail_path),
        format_path_result("field_zip_path", product.selected_zip),
        format_ready_folder_result(product),
        category_manual_result(page),
        proxy_purchase_manual_result(page),
        UI_TEXT["assist_publish_guard_result"],
    ]
    return format_result_lines(results)


def safe_page_url(page) -> str:
    try:
        return page.url or ""
    except Exception:
        return ""


def safe_page_title(page) -> str:
    try:
        return page.title(timeout=2000) or ""
    except Exception:
        return ""


def count_selector(page, selector: str) -> int:
    try:
        return page.locator(selector).count()
    except Exception:
        return 0


def collect_page_diagnostics(page) -> dict[str, object]:
    return {
        "url": safe_page_url(page),
        "title": safe_page_title(page),
        "input_count": count_selector(page, "input"),
        "textarea_count": count_selector(page, "textarea"),
        "file_count": count_selector(page, "input[type='file']"),
        "file_inputs": collect_file_input_diagnostics(page),
    }


def format_file_input_detail(info: FileInputInfo) -> str:
    return UI_TEXT["file_input_detail_line"].format(
        index=info.index,
        accept=display_detail(info.accept),
        name=display_detail(info.name),
        id=display_detail(info.input_id),
        aria_label=display_detail(info.aria_label),
        nearby=display_detail(info.nearby),
    )


def format_file_input_details(file_inputs: object) -> str:
    if not isinstance(file_inputs, tuple) or not file_inputs:
        return UI_TEXT["file_input_details_empty"]
    lines = [format_file_input_detail(info) for info in file_inputs if isinstance(info, FileInputInfo)]
    return "\n".join(lines) if lines else UI_TEXT["file_input_details_empty"]


def format_assist_summary(diagnostics: dict[str, object], results: str) -> str:
    return UI_TEXT["assist_summary_template"].format(
        url=diagnostics["url"],
        title=diagnostics["title"],
        input_count=diagnostics["input_count"],
        textarea_count=diagnostics["textarea_count"],
        file_count=diagnostics["file_count"],
        file_input_details=format_file_input_details(diagnostics.get("file_inputs", ())),
        results=results,
    )


def print_page_diagnostics(diagnostics: dict[str, object]) -> None:
    print(
        UI_TEXT["log_page_diagnostics"].format(
            url=diagnostics["url"],
            title=diagnostics["title"],
            input_count=diagnostics["input_count"],
            textarea_count=diagnostics["textarea_count"],
            file_count=diagnostics["file_count"],
        ),
        flush=True,
    )
    raw_file_inputs = diagnostics.get("file_inputs", ())
    if isinstance(raw_file_inputs, tuple):
        for info in raw_file_inputs:
            if not isinstance(info, FileInputInfo):
                continue
            print(
                UI_TEXT["log_file_input_detail"].format(
                    index=info.index,
                    accept=display_detail(info.accept),
                    name=display_detail(info.name),
                    id=display_detail(info.input_id),
                    aria_label=display_detail(info.aria_label),
                    nearby=display_detail(info.nearby),
                ),
                flush=True,
            )

def is_booth_login_page(page) -> bool:
    url = safe_page_url(page).lower()
    return "accounts.pixiv.net" in url or ("login" in url and ("booth.pm" in url or "pixiv" in url))


def is_manage_booth_page(page) -> bool:
    return "manage.booth.pm" in safe_page_url(page).lower()


def is_booth_edit_url(url: str) -> bool:
    normalized_url = url.lower()
    return "manage.booth.pm/items/" in normalized_url and "/edit" in normalized_url


def is_booth_new_url(url: str) -> bool:
    normalized_url = url.lower().rstrip("/")
    return "manage.booth.pm/items" in normalized_url and normalized_url.endswith("/new")


def is_booth_edit_page(page) -> bool:
    return is_booth_edit_url(safe_page_url(page))


def is_bad_booth_page(page) -> bool:
    if not is_manage_booth_page(page):
        return False
    if is_booth_new_url(safe_page_url(page)):
        return True
    title = safe_page_title(page).lower()
    url = safe_page_url(page).lower()
    return "404" in title or "404" in url or "not found" in title


def is_active_page(page) -> bool:
    try:
        return page.evaluate("document.visibilityState") == "visible"
    except Exception:
        return False


def collect_browser_pages(browser) -> list:
    pages = []
    for context in browser.contexts:
        try:
            pages.extend(context.pages)
        except Exception:
            continue
    return pages


def has_login_page(browser) -> bool:
    return any(is_booth_login_page(page) for page in collect_browser_pages(browser))


def find_bad_booth_page(browser):
    pages = collect_browser_pages(browser)
    for page in pages:
        if is_bad_booth_page(page) and is_active_page(page):
            return page
    for page in pages:
        if is_bad_booth_page(page):
            return page
    return None


def choose_booth_edit_page(browser):
    edit_pages = []
    for page in collect_browser_pages(browser):
        if not is_booth_edit_page(page):
            continue
        if is_active_page(page):
            return page
        edit_pages.append(page)
    return edit_pages[0] if edit_pages else None


def is_booth_page_url(url: str) -> bool:
    return is_booth_host(booth_url_parts(url).hostname)


def choose_booth_url_page(browser):
    pages = collect_browser_pages(browser)
    for page in pages:
        if is_active_page(page) and is_valid_booth_save_url(safe_page_url(page)):
            return page
    for page in pages:
        if is_active_page(page) and is_booth_page_url(safe_page_url(page)):
            return page
    for page in pages:
        if is_valid_booth_save_url(safe_page_url(page)):
            return page
    for page in pages:
        if is_booth_page_url(safe_page_url(page)):
            return page
    return None


def run_current_booth_url_worker(product: ProductInfo | None, events: queue.Queue[WorkerEvent]) -> None:
    try:
        from playwright.sync_api import sync_playwright
    except Exception as exc:
        events.put(("error_setup", UI_TEXT["dialog_playwright_setup"].format(error=exc)))
        return

    try:
        events.put(("status", "status_chrome_connecting"))
        with sync_playwright() as playwright:
            try:
                browser = playwright.chromium.connect_over_cdp(CHROME_CDP_URL, timeout=7000)
            except Exception as exc:
                events.put(("error_connect", UI_TEXT["dialog_chrome_connect_error"].format(error=exc)))
                return

            events.put(("status", "status_chrome_connected"))
            page = choose_booth_url_page(browser)
            if page is None:
                events.put(("booth_url_invalid", UI_TEXT["dialog_booth_url_invalid"].format(url=UI_TEXT["value_unset"])))
                return

            raw_url = safe_page_url(page)
            booth_url = normalize_booth_url_for_save(raw_url, product, browser, page)
            if not is_valid_booth_save_url(booth_url):
                events.put(("booth_url_invalid", UI_TEXT["dialog_booth_url_invalid"].format(url=raw_url or UI_TEXT["value_unset"])))
                return

            events.put(("booth_url_value", booth_url))
            events.put(("url_done", UI_TEXT["detail_booth_url_fetched"]))
    except Exception as exc:
        events.put(("error", UI_TEXT["dialog_playwright_error"].format(error=exc)))


def run_chrome_assist_worker(product: ProductInfo | None, events: queue.Queue[WorkerEvent]) -> None:
    try:
        from playwright.sync_api import sync_playwright
    except Exception as exc:
        events.put(("error_setup", UI_TEXT["dialog_playwright_setup"].format(error=exc)))
        return

    if product is None:
        events.put(("edit_page_required", UI_TEXT["detail_no_selection"]))
        return

    try:
        events.put(("status", "status_chrome_connecting"))
        with sync_playwright() as playwright:
            try:
                browser = playwright.chromium.connect_over_cdp(CHROME_CDP_URL, timeout=7000)
            except Exception as exc:
                events.put(("error_connect", UI_TEXT["dialog_chrome_connect_error"].format(error=exc)))
                return

            events.put(("status", "status_chrome_connected"))
            events.put(("detail", UI_TEXT["detail_chrome_connected"]))

            bad_page = find_bad_booth_page(browser)
            if bad_page is not None:
                diagnostics = collect_page_diagnostics(bad_page)
                print_page_diagnostics(diagnostics)
                events.put(("bad_page", UI_TEXT["detail_bad_edit_page"]))
                return

            page = choose_booth_edit_page(browser)
            if page is None:
                if has_login_page(browser):
                    events.put(("login_required", UI_TEXT["detail_login_required"]))
                    return
                events.put(("edit_page_required", UI_TEXT["detail_edit_page_required"]))
                return

            diagnostics = collect_page_diagnostics(page)
            print_page_diagnostics(diagnostics)
            events.put(("status", "status_edit_page_found"))
            events.put(("detail", UI_TEXT["detail_used_page"].format(url=diagnostics["url"])))
            events.put(("status", "status_assisting"))
            results = assist_booth_form(page, product, diagnostics)
            events.put(("summary", format_assist_summary(diagnostics, results)))
            events.put(("done", UI_TEXT["detail_assist_complete"]))
    except Exception as exc:
        events.put(("error", UI_TEXT["dialog_playwright_error"].format(error=exc)))

class DakeBoothAssistApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(*WINDOW_MIN_SIZE)
        self.root.configure(bg=THEME["background"])

        self.font_family = choose_font_family(self.root)
        self.root.option_add("*Font", (self.font_family, 10))

        self.apps: list[AppEntry] = []
        self.selected_product: ProductInfo | None = None
        self.booth_url_var = tk.StringVar(value="")
        self.playwright_active = False
        self.playwright_events: queue.Queue[WorkerEvent] = queue.Queue()

        self.selected_app_var = tk.StringVar(value=UI_TEXT["value_no_selection"])
        self.status_var = tk.StringVar(value=UI_TEXT["status_no_selection"])
        self.status_detail_var = tk.StringVar(value=UI_TEXT["detail_no_selection"])
        self.field_vars = {
            "product_type": tk.StringVar(value=UI_TEXT["value_unset"]),
            "title": tk.StringVar(value=UI_TEXT["value_unset"]),
            "price": tk.StringVar(value=UI_TEXT["value_unset"]),
            "tags": tk.StringVar(value=UI_TEXT["value_unset"]),
            "product_info": tk.StringVar(value=UI_TEXT["value_file_missing"]),
            "ready": tk.StringVar(value=UI_TEXT["value_no"]),
            "zip": tk.StringVar(value=UI_TEXT["value_file_missing"]),
            "thumbnail": tk.StringVar(value=UI_TEXT["value_file_missing"]),
            "screenshot": tk.StringVar(value=UI_TEXT["value_file_missing"]),
            "github_release": tk.StringVar(value=UI_TEXT["value_unset"]),
        }

        self._apply_window_icon()
        self._build_styles()
        self._build_ui()
        self.refresh_apps()

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
            padding=(14, 9),
            font=(self.font_family, 10, "bold"),
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
            padding=(12, 8),
            font=(self.font_family, 10, "bold"),
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

        self._build_header(container)

        body = tk.Frame(container, bg=THEME["background"])
        body.pack(fill="both", expand=True, pady=(16, 0))
        body.grid_columnconfigure(0, weight=1, uniform="body")
        body.grid_columnconfigure(1, weight=2, uniform="body")
        body.grid_rowconfigure(0, weight=1)

        self._build_app_panel(body)
        self._build_result_panel(body)
        self._build_actions(container)
        self._build_status(container)
        self._build_footer(container)

    def _build_header(self, parent: tk.Frame) -> None:
        header = tk.Frame(parent, bg=THEME["background"])
        header.pack(fill="x")

        tk.Label(
            header,
            text=UI_TEXT["main_title"],
            bg=THEME["background"],
            fg=THEME["text"],
            font=(self.font_family, 21, "bold"),
            anchor="w",
        ).pack(fill="x")

        tk.Label(
            header,
            text=UI_TEXT["main_description"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 10),
            anchor="w",
        ).pack(fill="x", pady=(6, 0))

    def _build_panel(self, parent: tk.Frame, title_key: str) -> tk.Frame:
        panel = tk.Frame(parent, bg=THEME["panel"], highlightbackground=THEME["border"], highlightthickness=1)
        tk.Label(
            panel,
            text=UI_TEXT[title_key],
            bg=THEME["panel"],
            fg=THEME["text"],
            font=(self.font_family, 12, "bold"),
            anchor="w",
        ).pack(fill="x", padx=16, pady=(13, 9))
        return panel

    def _build_app_panel(self, parent: tk.Frame) -> None:
        panel = self._build_panel(parent, "section_apps")
        panel.grid(row=0, column=0, sticky="nsew", padx=(0, 12))

        selected_row = tk.Frame(panel, bg=THEME["panel"])
        selected_row.pack(fill="x", padx=16)

        tk.Label(
            selected_row,
            text=UI_TEXT["selected_app_label"],
            bg=THEME["panel"],
            fg=THEME["muted"],
            font=(self.font_family, 9),
            anchor="w",
        ).pack(anchor="w")

        tk.Label(
            selected_row,
            textvariable=self.selected_app_var,
            bg=THEME["panel"],
            fg=THEME["text"],
            font=(self.font_family, 11, "bold"),
            anchor="w",
            wraplength=280,
            justify="left",
        ).pack(fill="x", pady=(4, 10))

        self.app_listbox = tk.Listbox(
            panel,
            height=10,
            activestyle="none",
            bg="#FFFFFF",
            fg=THEME["text"],
            selectbackground="#EAF2FF",
            selectforeground=THEME["text"],
            highlightthickness=1,
            highlightbackground=THEME["border"],
            bd=0,
            exportselection=False,
        )
        self.app_listbox.pack(fill="both", expand=True, padx=16, pady=(0, 12))
        self.app_listbox.bind("<<ListboxSelect>>", self._on_app_selected)

        self.reload_button = ttk.Button(
            panel,
            text=UI_TEXT["button_reload"],
            style="Secondary.TButton",
            command=self.refresh_apps,
        )
        self.reload_button.pack(anchor="e", padx=16, pady=(0, 16))

    def _build_result_panel(self, parent: tk.Frame) -> None:
        panel = self._build_panel(parent, "section_product")
        panel.grid(row=0, column=1, sticky="nsew")

        grid = tk.Frame(panel, bg=THEME["panel"])
        grid.pack(fill="x", padx=16)
        grid.grid_columnconfigure(1, weight=1)
        rows = (
            ("field_product_type", "product_type"),
            ("field_title", "title"),
            ("field_price", "price"),
            ("field_tags", "tags"),
            ("field_product_info", "product_info"),
            ("field_ready", "ready"),
            ("field_zip", "zip"),
            ("field_thumbnail", "thumbnail"),
            ("field_screenshot", "screenshot"),
            ("field_github_release", "github_release"),
        )
        for row_index, (label_key, value_key) in enumerate(rows):
            self._build_info_row(grid, row_index, label_key, self.field_vars[value_key])

        tk.Label(
            panel,
            text=UI_TEXT["field_description"],
            bg=THEME["panel"],
            fg=THEME["muted"],
            font=(self.font_family, 9, "bold"),
            anchor="w",
        ).pack(fill="x", padx=16, pady=(10, 4))

        self.description_text = tk.Text(
            panel,
            height=5,
            wrap="word",
            bg="#FFFFFF",
            fg=THEME["text"],
            bd=0,
            highlightthickness=1,
            highlightbackground=THEME["border"],
            padx=10,
            pady=8,
        )
        self.description_text.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        self.description_text.configure(state="disabled")

    def _build_info_row(self, parent: tk.Frame, row_index: int, label_key: str, variable: tk.StringVar) -> None:
        tk.Label(
            parent,
            text=UI_TEXT[label_key],
            bg=THEME["panel"],
            fg=THEME["muted"],
            font=(self.font_family, 9, "bold"),
            anchor="w",
            width=18,
        ).grid(row=row_index, column=0, sticky="nw", pady=4, padx=(0, 10))

        tk.Label(
            parent,
            textvariable=variable,
            bg=THEME["panel"],
            fg=THEME["text"],
            font=(self.font_family, 9),
            anchor="w",
            justify="left",
            wraplength=430,
        ).grid(row=row_index, column=1, sticky="ew", pady=4)

    def _build_actions(self, parent: tk.Frame) -> None:
        actions = self._build_panel(parent, "section_actions")
        actions.pack(fill="x", pady=(14, 0))

        booth_url_frame = tk.Frame(actions, bg=THEME["panel"])
        booth_url_frame.pack(fill="x", padx=16, pady=(0, 10))
        booth_url_frame.grid_columnconfigure(1, weight=1)
        tk.Label(
            booth_url_frame,
            text=UI_TEXT["field_booth_url"],
            bg=THEME["panel"],
            fg=THEME["muted"],
            font=(self.font_family, 9, "bold"),
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=(0, 10))
        self.booth_url_entry = ttk.Entry(booth_url_frame, textvariable=self.booth_url_var)
        self.booth_url_entry.grid(row=0, column=1, sticky="ew")

        button_grid = tk.Frame(actions, bg=THEME["panel"])
        button_grid.pack(fill="x", padx=16, pady=(0, 16))
        for column in range(3):
            button_grid.grid_columnconfigure(column, weight=1)

        self.launch_chrome_button = ttk.Button(
            button_grid,
            text=UI_TEXT["button_launch_chrome"],
            style="Primary.TButton",
            command=self.launch_logged_in_chrome,
        )
        self.assist_button = ttk.Button(
            button_grid,
            text=UI_TEXT["button_start_assist"],
            style="Primary.TButton",
            command=self.start_input_assist,
        )
        self.fetch_booth_url_button = ttk.Button(
            button_grid,
            text=UI_TEXT["button_get_chrome_url"],
            style="Secondary.TButton",
            command=self.fetch_current_chrome_url,
        )
        self.save_booth_url_button = ttk.Button(
            button_grid,
            text=UI_TEXT["button_save_booth_url"],
            style="Secondary.TButton",
            command=self.save_booth_url,
        )
        self.open_ready_button = ttk.Button(
            button_grid,
            text=UI_TEXT["button_open_ready"],
            style="Secondary.TButton",
            command=self.open_booth_ready,
        )
        self.copy_title_button = ttk.Button(
            button_grid,
            text=UI_TEXT["button_copy_title"],
            style="Secondary.TButton",
            command=lambda: self.copy_product_field("title", "field_title"),
        )
        self.copy_description_button = ttk.Button(
            button_grid,
            text=UI_TEXT["button_copy_description"],
            style="Secondary.TButton",
            command=lambda: self.copy_product_field("description", "field_description"),
        )
        self.copy_tags_button = ttk.Button(
            button_grid,
            text=UI_TEXT["button_copy_tags"],
            style="Secondary.TButton",
            command=lambda: self.copy_product_field("tags", "field_tags"),
        )
        self.copy_thumbnail_path_button = ttk.Button(
            button_grid,
            text=UI_TEXT["button_copy_thumbnail_path"],
            style="Secondary.TButton",
            command=lambda: self.copy_product_field("thumbnail_path", "field_product_image_path"),
        )
        self.copy_zip_path_button = ttk.Button(
            button_grid,
            text=UI_TEXT["button_copy_zip_path"],
            style="Secondary.TButton",
            command=lambda: self.copy_product_field("selected_zip", "field_zip_path"),
        )

        buttons = (
            self.launch_chrome_button,
            self.assist_button,
            self.fetch_booth_url_button,
            self.save_booth_url_button,
            self.open_ready_button,
            self.copy_title_button,
            self.copy_description_button,
            self.copy_tags_button,
            self.copy_thumbnail_path_button,
            self.copy_zip_path_button,
        )
        for index, button in enumerate(buttons):
            button.grid(row=index // 3, column=index % 3, sticky="ew", padx=5, pady=5)

    def _build_status(self, parent: tk.Frame) -> None:
        status_panel = tk.Frame(parent, bg=THEME["background"])
        status_panel.pack(fill="x", pady=(12, 0))

        tk.Label(
            status_panel,
            text=UI_TEXT["status_label"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 9, "bold"),
        ).pack(side="left", padx=(0, 8))

        self.status_badge = tk.Label(
            status_panel,
            textvariable=self.status_var,
            bg=THEME["subtle"],
            fg=THEME["muted"],
            font=(self.font_family, 9, "bold"),
            padx=12,
            pady=5,
        )
        self.status_badge.pack(side="left")

        tk.Label(
            status_panel,
            textvariable=self.status_detail_var,
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 9),
            anchor="w",
            justify="left",
        ).pack(side="left", fill="x", expand=True, padx=(12, 0))

    def _build_footer(self, parent: tk.Frame) -> None:
        footer = tk.Frame(parent, bg=THEME["background"])
        footer.pack(fill="x", pady=(10, 0))
        footer_text = UI_TEXT["footer_left"] + UI_TEXT["footer_separator"] + UI_TEXT["footer_note"] + UI_TEXT["footer_separator"] + UI_TEXT["footer_copyright"]
        tk.Label(
            footer,
            text=footer_text,
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 8),
            anchor="w",
        ).pack(fill="x")

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

    def refresh_apps(self) -> None:
        self._set_status("status_loading")
        self.apps = discover_apps()
        self.app_listbox.delete(0, tk.END)

        for entry in self.apps:
            key = "list_has_product" if entry.has_product else "list_no_product"
            label = f"{product_type_label(entry.product_type)}: {entry.name}"
            self.app_listbox.insert(tk.END, UI_TEXT[key].format(name=label))

        if not self.apps:
            self.selected_product = None
            self.selected_app_var.set(UI_TEXT["value_no_selection"])
            self._clear_product_fields()
            self._set_status("status_no_selection", UI_TEXT["detail_no_apps"])
            self._update_buttons()
            return

        selected_index = self._preferred_app_index()
        self.app_listbox.selection_clear(0, tk.END)
        self.app_listbox.selection_set(selected_index)
        self.app_listbox.see(selected_index)
        self.load_product_for_index(selected_index)

    def _preferred_app_index(self) -> int:
        config = load_config()
        last_app_path = config.get("last_app_path", "")
        if last_app_path:
            for index, entry in enumerate(self.apps):
                if str(entry.path) == last_app_path:
                    return index
        for index, entry in enumerate(self.apps):
            if entry.has_product:
                return index
        return 0

    def _on_app_selected(self, _event=None) -> None:
        selection = self.app_listbox.curselection()
        if not selection:
            self.selected_product = None
            self.selected_app_var.set(UI_TEXT["value_no_selection"])
            self._clear_product_fields()
            self._set_status("status_no_selection", UI_TEXT["detail_no_selection"])
            self._update_buttons()
            return
        self.load_product_for_index(selection[0])

    def load_product_for_index(self, index: int) -> None:
        if index < 0 or index >= len(self.apps):
            return

        entry = self.apps[index]
        save_config(entry.path)
        self.selected_app_var.set(entry.name)
        product = build_product_info(entry.path)
        self.selected_product = product
        self._render_product(product)
        detail_key = "detail_loaded" if product.product_path.exists() else "detail_missing_product"
        self._set_status("status_ready", UI_TEXT[detail_key])
        self._update_buttons()

    def _clear_product_fields(self) -> None:
        for key, variable in self.field_vars.items():
            if key in {"ready"}:
                variable.set(UI_TEXT["value_no"])
            elif key in {"product_info", "zip", "thumbnail", "screenshot"}:
                variable.set(UI_TEXT["value_file_missing"])
            else:
                variable.set(UI_TEXT["value_unset"])
        self._set_description(UI_TEXT["value_unset"])
        self.booth_url_var.set("")

    def _render_product(self, product: ProductInfo) -> None:
        self.field_vars["product_type"].set(product_type_label(product.product_type))
        self.field_vars["title"].set(product.title or UI_TEXT["value_unset"])
        self.field_vars["price"].set(product.price or UI_TEXT["value_unset"])
        self.field_vars["tags"].set(product.tags or UI_TEXT["value_unset"])
        self.field_vars["product_info"].set(product.product_source or UI_TEXT["value_file_missing"])
        self.field_vars["ready"].set(UI_TEXT["value_yes"] if product.booth_ready_exists else UI_TEXT["value_no"])
        self.field_vars["zip"].set(self._format_zip_value(product))
        self.field_vars["thumbnail"].set(self._format_path_value(product.thumbnail_path))
        self.field_vars["screenshot"].set(self._format_path_value(product.screenshot_path))
        self.field_vars["github_release"].set(product.github_release or UI_TEXT["value_unset"])
        self.booth_url_var.set(product.url if is_valid_booth_save_url(product.url) else "")
        self._set_description(product.description or UI_TEXT["value_unset"])

    def _format_zip_value(self, product: ProductInfo) -> str:
        if not product.selected_zip:
            return UI_TEXT["value_file_missing"]
        if len(product.zip_files) > 1:
            return UI_TEXT["value_zip_multiple"].format(name=product.selected_zip.name, count=len(product.zip_files) - 1)
        return UI_TEXT["value_file_found"].format(name=product.selected_zip.name)

    def _format_path_value(self, path: Path | None) -> str:
        if path is None:
            return UI_TEXT["value_file_missing"]
        return UI_TEXT["value_file_found"].format(name=path.name)

    def _set_description(self, value: str) -> None:
        self.description_text.configure(state="normal")
        self.description_text.delete("1.0", tk.END)
        self.description_text.insert("1.0", value)
        self.description_text.configure(state="disabled")

    def _set_status(self, status_key: str, detail: str | None = None) -> None:
        self.status_var.set(UI_TEXT[status_key])
        background, foreground = STATUS_THEME.get(status_key, STATUS_THEME["status_no_selection"])
        self.status_badge.configure(bg=background, fg=foreground)
        if detail is not None:
            self.status_detail_var.set(detail)

    def _update_buttons(self) -> None:
        has_selection = self.selected_product is not None
        normal_if_selection = tk.NORMAL if has_selection else tk.DISABLED
        playwright_state = tk.DISABLED if self.playwright_active else normal_if_selection
        self.launch_chrome_button.configure(state=playwright_state)
        self.assist_button.configure(state=playwright_state)
        self.fetch_booth_url_button.configure(state=playwright_state)
        self.save_booth_url_button.configure(state=normal_if_selection)
        self.open_ready_button.configure(state=normal_if_selection)
        self.copy_title_button.configure(state=normal_if_selection)
        self.copy_description_button.configure(state=normal_if_selection)
        self.copy_tags_button.configure(state=normal_if_selection)
        self.copy_thumbnail_path_button.configure(state=normal_if_selection)
        self.copy_zip_path_button.configure(state=normal_if_selection)

    def copy_product_field(self, attribute: str, label_key: str) -> None:
        if self.selected_product is None:
            self._set_status("status_no_selection", UI_TEXT["detail_no_selection"])
            return

        if attribute == "thumbnail_path" and self.selected_product.image_paths:
            value = "\n".join(str(path_item) for path_item in self.selected_product.image_paths)
        else:
            value = getattr(self.selected_product, attribute, "")
        label = UI_TEXT[label_key]
        if not value:
            self._set_status("status_ready", UI_TEXT["detail_copy_empty"].format(label=label))
            return

        self.root.clipboard_clear()
        self.root.clipboard_append(str(value))
        self._set_status("status_ready", UI_TEXT["detail_copied"].format(label=label))

    def open_booth_ready(self) -> None:
        if self.selected_product is None:
            self._set_status("status_no_selection", UI_TEXT["detail_no_selection"])
            return
        path = self.selected_product.booth_ready_dir
        if not path.exists():
            self._set_status("status_ready", UI_TEXT["detail_ready_missing"])
            messagebox.showinfo(UI_TEXT["dialog_notice_title"], UI_TEXT["dialog_no_ready_folder"].format(path=path))
            return
        try:
            os.startfile(path)
        except OSError as exc:
            self._set_status("status_error")
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["dialog_open_folder_error"].format(error=exc))
            return
        self._set_status("status_ready", UI_TEXT["detail_opened_ready"])

    def launch_logged_in_chrome(self) -> None:
        if self.playwright_active:
            messagebox.showinfo(UI_TEXT["dialog_notice_title"], UI_TEXT["dialog_playwright_busy"])
            return

        self._set_status("status_chrome_launching")
        chrome_path = find_chrome_executable()
        manual_command = build_manual_chrome_command()
        if chrome_path is None:
            self._set_status("status_error")
            messagebox.showerror(
                UI_TEXT["dialog_error_title"],
                UI_TEXT["dialog_chrome_not_found"].format(command=manual_command),
            )
            return

        try:
            get_chrome_profile_dir().mkdir(parents=True, exist_ok=True)
            creationflags = 0
            if os.name == "nt":
                creationflags = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0) | getattr(subprocess, "DETACHED_PROCESS", 0)
            subprocess.Popen(
                build_chrome_launch_args(chrome_path),
                stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=creationflags,
            )
        except Exception as exc:
            self._set_status("status_error")
            messagebox.showerror(
                UI_TEXT["dialog_error_title"],
                UI_TEXT["dialog_chrome_launch_error"].format(command=manual_command, error=exc),
            )
            return

        self._set_status("status_chrome_ready", UI_TEXT["detail_chrome_launched"])

    def fetch_current_chrome_url(self) -> None:
        self._start_booth_url_fetch_task()

    def _start_booth_url_fetch_task(self) -> None:
        if self.playwright_active:
            messagebox.showinfo(UI_TEXT["dialog_notice_title"], UI_TEXT["dialog_playwright_busy"])
            return
        if self.selected_product is None:
            self._set_status("status_no_selection", UI_TEXT["detail_no_selection"])
            return

        self.playwright_active = True
        self._update_buttons()
        product = self.selected_product
        worker = threading.Thread(target=run_current_booth_url_worker, args=(product, self.playwright_events), daemon=True)
        worker.start()
        self.root.after(QUEUE_POLL_MS, self._poll_playwright_events)

    def save_booth_url(self) -> None:
        if self.selected_product is None:
            self._set_status("status_no_selection", UI_TEXT["detail_no_selection"])
            return

        raw_url = self.booth_url_var.get().strip()
        if not raw_url:
            messagebox.showinfo(UI_TEXT["dialog_notice_title"], UI_TEXT["dialog_booth_url_empty"])
            return

        booth_url = normalize_booth_url_for_save(raw_url, self.selected_product)
        if not is_valid_booth_save_url(booth_url):
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["dialog_booth_url_invalid"].format(url=raw_url))
            return

        targets = booth_url_save_targets(self.selected_product.app_dir)
        targets_text = format_save_targets(self.selected_product.app_dir, targets)
        confirmed = messagebox.askyesno(
            UI_TEXT["dialog_notice_title"],
            UI_TEXT["dialog_booth_url_confirm"].format(url=booth_url, targets=targets_text),
        )
        if not confirmed:
            return

        try:
            saved_targets = save_booth_url_to_products(self.selected_product.app_dir, booth_url)
        except Exception as exc:
            self._set_status("status_error")
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["dialog_booth_url_save_error"].format(error=exc))
            return

        saved_targets_text = format_save_targets(self.selected_product.app_dir, saved_targets)
        refreshed_product = build_product_info(self.selected_product.app_dir)
        self.selected_product = refreshed_product
        self._render_product(refreshed_product)
        self.booth_url_var.set(booth_url)
        self._set_status("status_ready", UI_TEXT["detail_booth_url_saved"].format(targets=saved_targets_text.replace("\n", ", ")))
        messagebox.showinfo(
            UI_TEXT["dialog_notice_title"],
            UI_TEXT["dialog_booth_url_saved"].format(targets=saved_targets_text),
        )

    def start_input_assist(self) -> None:
        self._start_chrome_assist_task()

    def _start_chrome_assist_task(self) -> None:
        if self.playwright_active:
            messagebox.showinfo(UI_TEXT["dialog_notice_title"], UI_TEXT["dialog_playwright_busy"])
            return
        if self.selected_product is None:
            self._set_status("status_no_selection", UI_TEXT["detail_no_selection"])
            return

        self.playwright_active = True
        self._update_buttons()
        product = self.selected_product
        worker = threading.Thread(target=run_chrome_assist_worker, args=(product, self.playwright_events), daemon=True)
        worker.start()
        self.root.after(QUEUE_POLL_MS, self._poll_playwright_events)
    def _poll_playwright_events(self) -> None:
        while True:
            try:
                kind, payload = self.playwright_events.get_nowait()
            except queue.Empty:
                break

            if kind == "status":
                self._set_status(payload)
            elif kind == "detail":
                self.status_detail_var.set(payload)
            elif kind == "summary":
                self._set_status("status_assist_complete", UI_TEXT["assist_publish_guard"])
                messagebox.showinfo(UI_TEXT["dialog_assist_summary_title"], payload)
            elif kind == "booth_url_value":
                self.booth_url_var.set(payload)
            elif kind == "url_done":
                self.playwright_active = False
                self._set_status("status_ready", payload)
                self._update_buttons()
            elif kind == "booth_url_invalid":
                self.playwright_active = False
                self._set_status("status_error")
                self._update_buttons()
                messagebox.showerror(UI_TEXT["dialog_error_title"], payload)
            elif kind == "login_required":
                self.playwright_active = False
                self._set_status("status_login_required", payload)
                self._update_buttons()
                messagebox.showinfo(UI_TEXT["dialog_notice_title"], payload)
            elif kind == "edit_page_required":
                self.playwright_active = False
                self._set_status("status_edit_page_required", payload)
                self._update_buttons()
                messagebox.showinfo(UI_TEXT["dialog_notice_title"], payload)
            elif kind == "bad_page":
                self.playwright_active = False
                self._set_status("status_edit_page_required", payload)
                self._update_buttons()
                messagebox.showinfo(UI_TEXT["dialog_notice_title"], payload)
            elif kind == "error_setup":
                self.playwright_active = False
                self._set_status("status_error")
                self._update_buttons()
                messagebox.showerror(UI_TEXT["dialog_error_title"], payload)
            elif kind == "error_connect":
                self.playwright_active = False
                self._set_status("status_error")
                self._update_buttons()
                messagebox.showerror(UI_TEXT["dialog_error_title"], payload)
            elif kind == "error":
                self.playwright_active = False
                self._set_status("status_error")
                self._update_buttons()
                messagebox.showerror(UI_TEXT["dialog_error_title"], payload)
            elif kind == "done":
                self.playwright_active = False
                self._set_status("status_assist_complete", payload)
                self._update_buttons()

        if self.playwright_active:
            self.root.after(QUEUE_POLL_MS, self._poll_playwright_events)


def run_launch_check() -> int:
    apps = discover_apps()
    if apps:
        build_product_info(apps[0].path)

    with tempfile.TemporaryDirectory(dir=get_base_dir()) as temp_dir:
        temp_root = Path(temp_dir)

        missing_product_app = temp_root / "MissingProduct"
        missing_product_app.mkdir()
        missing_product = build_product_info(missing_product_app)
        if missing_product.title or missing_product.booth_ready_exists or missing_product.selected_zip is not None:
            raise RuntimeError("missing product fixture failed")

        zip_app = temp_root / "ZipApp"
        ready_dir = zip_app / READY_DIR_NAME
        ready_dir.mkdir(parents=True)
        (zip_app / PRODUCT_FILE_NAME).write_text(
            "TITLE=Sample\nPRICE=1000\nDESCRIPTION=Sample description\nTAGS=alpha,beta\n",
            encoding="utf-8",
        )
        (ready_dir / "b.zip").write_bytes(b"")
        (ready_dir / "a.zip").write_bytes(b"")
        zip_product = build_product_info(zip_app)
        if not zip_product.booth_ready_exists or len(zip_product.zip_files) != 2:
            raise RuntimeError("zip fixture failed")
        if zip_product.selected_zip is None or zip_product.selected_zip.name != "a.zip":
            raise RuntimeError("zip selection fixture failed")

        ready_product_app = temp_root / "ReadyProduct"
        ready_product_dir = ready_product_app / READY_DIR_NAME
        ready_product_dir.mkdir(parents=True)
        (ready_product_dir / PRODUCT_FILE_NAME).write_text(
            "TITLE=Ready Sample\nPRICE=1200\nDESCRIPTION=Ready description\nTAGS=ready,booth\n",
            encoding="utf-8",
        )
        ready_product = build_product_info(ready_product_app)
        if ready_product.title != "Ready Sample":
            raise RuntimeError("ready product fixture failed")
        if ready_product.product_source != f"{READY_DIR_NAME}/{PRODUCT_FILE_NAME}":
            raise RuntimeError("ready product source fixture failed")

        both_product_app = temp_root / "BothProduct"
        both_ready_dir = both_product_app / READY_DIR_NAME
        both_ready_dir.mkdir(parents=True)
        both_product_app.mkdir(exist_ok=True)
        (both_product_app / PRODUCT_FILE_NAME).write_text("TITLE=Root Sample\n", encoding="utf-8")
        (both_ready_dir / PRODUCT_FILE_NAME).write_text("TITLE=Ready Sample\n", encoding="utf-8")
        both_product = build_product_info(both_product_app)
        if both_product.title != "Ready Sample" or both_product.product_source != f"{READY_DIR_NAME}/{PRODUCT_FILE_NAME}":
            raise RuntimeError("ready product priority fixture failed")

        factory_app = temp_root / "FactoryV2"
        factory_ready_dir = factory_app / READY_DIR_NAME
        factory_ready_dir.mkdir(parents=True)
        (factory_app / PRODUCT_FILE_NAME).write_text(
            "商品名\nRoot Title\n\n概要\nRoot description\n\nできること\n- Root feature\n\n使い方\n1. Root usage\n",
            encoding="utf-8",
        )
        (factory_ready_dir / PRODUCT_FILE_NAME).write_text(
            "# 商品名\nReady Title\n\n# 価格案\n500円\n\n# 商品紹介文\nReady description\n\n# タグ\nPDF\nWindows\n\n# GitHub Release\nhttps://github.com/yukiPHZ/dake-series/releases/tag/Ready_v1.0.0\n\n# URL\n",
            encoding="utf-8",
        )
        factory_product = build_product_info(factory_app)
        if factory_product.title != "Ready Title" or factory_product.price != "500円":
            raise RuntimeError("factory v2 title/price fixture failed")
        if split_tags(factory_product.tags) != ["PDF", "Windows"]:
            raise RuntimeError("factory v2 tags fixture failed")
        if factory_product.github_release != "https://github.com/yukiPHZ/dake-series/releases/tag/Ready_v1.0.0":
            raise RuntimeError("factory v2 github release fixture failed")

        readme_fallback_app = temp_root / "ReadmeFallback"
        readme_fallback_app.mkdir()
        (readme_fallback_app / "README.md").write_text(
            "## DAKE_META\n\n```json\n{\"display_name\":\"Fallback Name\",\"update_summary\":\"Fallback summary\",\"release_url\":\"https://github.com/yukiPHZ/dake-series/releases/tag/Fallback_v1.0.0\"}\n```\n",
            encoding="utf-8",
        )
        readme_fallback_product = build_product_info(readme_fallback_app)
        if readme_fallback_product.title != "Fallback Name":
            raise RuntimeError("readme display_name fallback fixture failed")
        if readme_fallback_product.description != "Fallback summary":
            raise RuntimeError("readme update_summary fallback fixture failed")
        if readme_fallback_product.github_release != "https://github.com/yukiPHZ/dake-series/releases/tag/Fallback_v1.0.0":
            raise RuntimeError("readme release_url fallback fixture failed")

        pack_app = temp_root / "PackProduct"
        pack_ready_dir = pack_app / PACK_READY_DIR_NAME
        pack_ready_dir.mkdir(parents=True)
        (pack_app / "README.md").write_text(
            "## PACK_META\n\n```json\n{\"display_name\":\"Pack Sample\",\"summary\":\"Pack summary\",\"price\":980,\"booth_url\":\"\",\"included_apps\":[\"Sample\"]}\n```\n",
            encoding="utf-8",
        )
        (pack_ready_dir / PRODUCT_FILE_NAME).write_text("# 商品名\nPack Ready\n\n# 価格案\n980円\n\n# URL\n", encoding="utf-8")
        (pack_ready_dir / "pack.zip").write_bytes(b"")
        pack_image_dir = pack_ready_dir / "images"
        pack_image_dir.mkdir()
        (pack_image_dir / THUMBNAIL_NAME).write_bytes(b"")
        pack_product = build_product_info(pack_app)
        if pack_product.product_type != PRODUCT_TYPE_PACK or pack_product.product_source != f"{PACK_READY_DIR_NAME}/{PRODUCT_FILE_NAME}":
            raise RuntimeError("pack product fixture failed")
        if not pack_product.image_paths:
            raise RuntimeError("pack image path fixture failed")
        pack_saved_url = "https://peakheadz.booth.pm/items/8417562"
        pack_saved_targets = save_booth_url_to_products(pack_app, pack_saved_url)
        if tuple(format_save_target(pack_app, target) for target in pack_saved_targets) != (f"{PACK_READY_DIR_NAME}/{PRODUCT_FILE_NAME}", "README.md"):
            raise RuntimeError("pack booth url target fixture failed")
        if pack_saved_url not in read_text_safely(pack_ready_dir / PRODUCT_FILE_NAME):
            raise RuntimeError("pack product booth url save fixture failed")
        if '"booth_url": "https://peakheadz.booth.pm/items/8417562"' not in read_text_safely(pack_app / "README.md"):
            raise RuntimeError("pack readme booth url save fixture failed")

        url_app = temp_root / "UrlApp"
        url_ready_dir = url_app / READY_DIR_NAME
        url_ready_dir.mkdir(parents=True)
        url_app_product = url_app / PRODUCT_FILE_NAME
        url_ready_product = url_ready_dir / PRODUCT_FILE_NAME
        url_app_product.write_text("TITLE=Root URL\nURL=https://old.booth.pm/items/1\n", encoding="utf-8")
        url_ready_product.write_text("TITLE=Ready URL\nBOOTH URL: https://old.booth.pm/items/2\n", encoding="utf-8")
        saved_url = "https://peakheadz.booth.pm/items/8417561"
        saved_targets = save_booth_url_to_products(url_app, saved_url)
        if tuple(format_save_target(url_app, target) for target in saved_targets) != (f"{READY_DIR_NAME}/{PRODUCT_FILE_NAME}", PRODUCT_FILE_NAME):
            raise RuntimeError("booth url target fixture failed")
        if saved_url not in read_text_safely(url_ready_product):
            raise RuntimeError("ready booth url save fixture failed")
        if "URL=https://peakheadz.booth.pm/items/8417561" not in read_text_safely(url_app_product):
            raise RuntimeError("root booth url update fixture failed")
        if parse_booth_product(url_ready_product).get("url") != saved_url:
            raise RuntimeError("booth url parse fixture failed")

    if not is_booth_edit_url("https://manage.booth.pm/items/8417561/edit"):
        raise RuntimeError("booth edit URL fixture failed")
    items_new_url = "https://manage.booth.pm/items" + "/new"
    if not is_booth_new_url(items_new_url):
        raise RuntimeError("items new URL fixture failed")
    if is_booth_edit_url(items_new_url):
        raise RuntimeError("items new URL should not match edit fixture")
    if is_booth_edit_url("https://manage.booth.pm/items/8417561"):
        raise RuntimeError("item URL without edit should not match edit fixture")
    edit_url = "https://manage.booth.pm/items/8417561/edit"
    if normalize_booth_url_for_save(edit_url) != edit_url:
        raise RuntimeError("edit booth url fallback fixture failed")
    if not is_valid_booth_save_url("https://peakheadz.booth.pm/items/8417561"):
        raise RuntimeError("public booth url validation fixture failed")
    if is_valid_booth_save_url("https://example.com/items/8417561"):
        raise RuntimeError("external booth url rejection fixture failed")
    if "BOOTH_URL=https://peakheadz.booth.pm/items/8417561" not in update_booth_url_text("TITLE=Sample\n", "https://peakheadz.booth.pm/items/8417561"):
        raise RuntimeError("booth url append equals fixture failed")
    if "# URL\nhttps://peakheadz.booth.pm/items/8417561\n# Tags" not in update_booth_url_text("# URL\nhttps://old.booth.pm/items/1\n# Tags\nalpha\n", "https://peakheadz.booth.pm/items/8417561"):
        raise RuntimeError("booth url heading update fixture failed")

    tag_sample = "Windows, 実務、ツール\n仕事効率化, 軽量, シンプル"
    expected_tags = ["Windows", "実務", "ツール", "仕事効率化", "軽量", "シンプル"]
    if split_tags(tag_sample) != expected_tags:
        raise RuntimeError("tag split fixture failed")

    image_file_info = FileInputInfo(0, "image/jpeg,image/png", "item[images][]", "thumb", "", "商品画像")
    zip_file_info = FileInputInfo(1, ".zip,application/zip", "item[files][]", "zip", "", "作品ファイル")
    if not looks_like_image_file_input(image_file_info):
        raise RuntimeError("image file input fixture failed")
    if not looks_like_zip_file_input(zip_file_info):
        raise RuntimeError("zip file input fixture failed")
    if looks_like_zip_file_input(image_file_info):
        raise RuntimeError("file input separation fixture failed")
    manual_upload_path = Path(__file__).resolve()
    image_manual_result = set_product_image_file(None, UI_TEXT["field_product_image"], manual_upload_path, ())
    if image_manual_result != UI_TEXT["assist_manual_set"].format(label=UI_TEXT["field_product_image"]):
        raise RuntimeError("manual image upload fixture failed")
    zip_manual_result = set_product_zip_file(None, UI_TEXT["field_zip"], manual_upload_path, ())
    if zip_manual_result != UI_TEXT["assist_manual_set"].format(label=UI_TEXT["field_zip"]):
        raise RuntimeError("manual zip upload fixture failed")
    if not format_path_result("field_product_image_path", manual_upload_path).startswith(UI_TEXT["field_product_image_path"]):
        raise RuntimeError("manual path result fixture failed")

    real_import = builtins.__import__

    def blocked_import(name, globals=None, locals=None, fromlist=(), level=0):
        if name.startswith("playwright"):
            raise ModuleNotFoundError("playwright")
        return real_import(name, globals, locals, fromlist, level)

    setup_events: queue.Queue[WorkerEvent] = queue.Queue()
    try:
        builtins.__import__ = blocked_import
        run_chrome_assist_worker(None, setup_events)
    finally:
        builtins.__import__ = real_import

    setup_kind, _setup_message = setup_events.get_nowait()
    if setup_kind != "error_setup":
        raise RuntimeError("playwright setup guidance fixture failed")

    product_apps = sum(1 for entry in apps if entry.has_product)
    pack_products = sum(1 for entry in apps if entry.product_type == PRODUCT_TYPE_PACK)
    shimarisu_pack = next((build_product_info(entry.path) for entry in apps if entry.name == SHIMARISU_PACK_ENTRY_NAME), None)
    if shimarisu_pack is not None:
        if shimarisu_pack.title != "しまりすくん 実務判断Pack":
            raise RuntimeError("shimarisu pack title fixture failed")
        if shimarisu_pack.price != "3,000円":
            raise RuntimeError("shimarisu pack price fixture failed")
        if shimarisu_pack.selected_zip is None or shimarisu_pack.selected_zip.name != "SHIMARISU_Pack.zip":
            raise RuntimeError("shimarisu pack zip fixture failed")
        if not shimarisu_pack.image_paths or shimarisu_pack.image_paths[0].name != THUMBNAIL_NAME:
            raise RuntimeError("shimarisu pack image fixture failed")
    dake_backup_source = "missing"
    dake_backup_fields = "missing"
    dake_backup_tag_count = 0
    for entry in apps:
        if entry.name == "DAKE_Backup":
            backup_info = build_product_info(entry.path)
            dake_backup_source = backup_info.product_source or "missing"
            dake_backup_fields = (
                f"title={bool(backup_info.title)},"
                f"price={bool(backup_info.price)},"
                f"description={bool(backup_info.description)},"
                f"tags={bool(backup_info.tags)}"
            )
            dake_backup_tag_count = len(split_tags(backup_info.tags))
            break

    print(f"{APP_NAME} launch-check OK")
    print(f"apps_root={get_apps_root()}")
    print(f"apps={len(apps)}")
    print(f"product_apps={product_apps}")
    print(f"pack_products={pack_products}")
    print(f"shimarisu_pack={'ready' if shimarisu_pack is not None else 'missing'}")
    print(f"dake_backup_product={dake_backup_source}")
    print(f"dake_backup_fields={dake_backup_fields}")
    print(f"dake_backup_tag_count={dake_backup_tag_count}")
    print(f"cdp_url={CHROME_CDP_URL}")
    print(f"edit_url_hint={BOOTH_EDIT_URL_HINT}")
    print(f"chrome_path={find_chrome_executable() or 'missing'}")
    print(f"chrome_profile={get_chrome_profile_dir()}")
    print("fixtures=missing_product, missing_ready, multiple_zip, missing_playwright_python, ready_product_lookup, ready_product_priority, factory_v2_product, readme_fallback, open_edit_url_only, bad_page_detection, tag_split, file_input_detection, manual_file_upload, booth_url_save, pack_product")
    return 0


def main() -> None:
    if "--launch-check" in sys.argv:
        raise SystemExit(run_launch_check())
    app = DakeBoothAssistApp()
    app.run()


if __name__ == "__main__":
    main()
