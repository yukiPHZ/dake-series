# -*- coding: utf-8 -*-
from __future__ import annotations

import builtins
import json
import os
import queue
import re
import shutil
import subprocess
import sys
import threading
import tempfile
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import font as tkfont
from tkinter import messagebox, ttk


APP_NAME = "DakeBOOTHアシスト"
WINDOW_TITLE = "DakeBOOTHアシスト"
COPYRIGHT = "© 2026 しまリス不動産 / Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "main_title": "BOOTH登録を補助する",
    "main_description": "booth_product.txt と booth_ready から、登録作業を静かに進めます。",
    "section_apps": "アプリ選択",
    "section_product": "読み取り結果",
    "section_actions": "操作",
    "section_status": "ステータス",
    "selected_app_label": "選択中アプリ",
    "button_reload": "再読み込み",
    "button_launch_chrome": "ログイン済みChromeを起動",
    "button_start_assist": "Chrome接続で入力補助",
    "button_open_ready": "booth_readyフォルダを開く",
    "button_copy_title": "商品名をコピー",
    "button_copy_description": "説明文をコピー",
    "button_copy_tags": "タグをコピー",
    "field_title": "商品名",
    "field_price": "価格",
    "field_description": "説明文",
    "field_product_info": "商品情報",
    "field_tags": "タグ",
    "field_ready": "booth_ready",
    "field_zip": "zipファイル",
    "field_thumbnail": "booth_thumbnail.jpg",
    "field_screenshot": "screenshot.webp",
    "field_github_release": "GitHub Release URL",
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
    "detail_chrome_launched": "Chromeを開きました。BOOTHへログインし、商品登録または編集画面を開いてください。",
    "detail_chrome_connected": "ログイン済みChromeへ接続しました。開いているBOOTH編集画面を確認しています。",
    "detail_login_required": "Chrome上でBOOTHへログインしてから、もう一度実行してください。",
    "detail_edit_page_required": "ChromeでBOOTHの商品登録または編集画面を開いてから、もう一度実行してください。",
    "detail_assist_complete": "公開ボタンは押していません。内容を確認してください。",
    "assist_filled": "{label}: 入力しました",
    "assist_file_set": "{label}: ファイルを選択しました",
    "assist_manual": "{label}: 手動で入力してください",
    "assist_missing_value": "{label}: 未設定のためスキップしました",
    "assist_publish_guard": "公開ボタンは押していません。内容を確認し、公開判断は人間が行ってください。",
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


@dataclass(frozen=True)
class ProductInfo:
    app_name: str
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
    screenshot_path: Path | None


WorkerEvent = tuple[str, str]


FIELD_ALIASES = {
    "商品名": "title",
    "タイトル": "title",
    "name": "title",
    "title": "title",
    "価格": "price",
    "販売価格": "price",
    "price": "price",
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
    "zip": "zip_path",
    "zipfile": "zip_path",
    "zippath": "zip_path",
    "zip_path": "zip_path",
    "商品画像": "thumbnail_path",
    "画像": "thumbnail_path",
    "thumbnail": "thumbnail_path",
    "thumbnailpath": "thumbnail_path",
    "thumbnail_path": "thumbnail_path",
}

KEY_VALUE_RE = re.compile(r"^\s*([^:=]+?)\s*[:=]\s*(.*)$")
GITHUB_RELEASE_RE = re.compile(r"https?://github\.com/[^\s)]+/releases/[^\s)]+", re.IGNORECASE)


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
THUMBNAIL_FILE_LOCATORS = (
    ("css", "input[type='file'][accept*='image']"),
    ("css", "input[type='file'][name*='image' i]"),
    ("css", "input[type='file'][name*='thumbnail' i]"),
)
ZIP_FILE_LOCATORS = (
    ("css", "input[type='file'][accept*='zip']"),
    ("css", "input[type='file'][accept*='application']"),
    ("css", "input[type='file'][name*='zip' i]"),
    ("css", "input[type='file'][name*='file' i]"),
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


def resolve_field_name(label: str) -> str | None:
    key = normalize_key(label)
    if key in FIELD_ALIASES:
        return FIELD_ALIASES[key]
    for alias, field_name in FIELD_ALIASES.items():
        if alias and alias in key:
            return field_name
    return None


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


def find_booth_product_path(app_dir: Path) -> Path | None:
    candidates = (
        app_dir / PRODUCT_FILE_NAME,
        app_dir / READY_DIR_NAME / PRODUCT_FILE_NAME,
    )
    for candidate in candidates:
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


def clean_tags(value: str) -> str:
    tags: list[str] = []
    for raw_line in value.replace("、", ",").splitlines():
        line = re.sub(r"^\s*[-*・]\s*", "", raw_line.strip())
        for part in line.split(","):
            tag = part.strip()
            if tag:
                tags.append(tag)
    return ", ".join(dict.fromkeys(tags))


def clean_url(value: str) -> str:
    return clean_single_line(value).rstrip("、。，,)")


def parse_booth_product(path: Path) -> dict[str, str]:
    if not path.exists():
        return {}

    fields: dict[str, str] = {}
    buffers: dict[str, list[str]] = {}
    current_field: str | None = None

    for raw_line in read_text_safely(path).splitlines():
        line = raw_line.rstrip()
        stripped = line.strip()

        if stripped.startswith("#"):
            field_name = resolve_field_name(stripped.lstrip("#").strip())
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

        if current_field:
            buffers.setdefault(current_field, []).append(line)

    for field_name, lines in buffers.items():
        if field_name not in fields and any(line.strip() for line in lines):
            fields[field_name] = "\n".join(lines).strip()

    if "github_release" not in fields:
        release_match = GITHUB_RELEASE_RE.search(read_text_safely(path))
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


def build_product_info(app_dir: Path) -> ProductInfo:
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
    product_source = format_product_source(app_dir, found_product_path)
    ready_dir = app_dir / READY_DIR_NAME
    ready_exists = ready_dir.exists() and ready_dir.is_dir()
    zip_files = find_zip_files(ready_dir)
    zip_from_text = resolve_product_relative_path(app_dir, parsed.get("zip_path", ""))
    selected_zip = zip_files[0] if zip_files else zip_from_text

    thumbnail_path = find_file_case_insensitive(ready_dir, THUMBNAIL_NAME)
    if thumbnail_path is None:
        thumbnail_path = find_file_case_insensitive(app_dir / "assets", THUMBNAIL_NAME)
    if thumbnail_path is None:
        thumbnail_path = resolve_product_relative_path(app_dir, parsed.get("thumbnail_path", ""))

    screenshot_path = find_file_case_insensitive(app_dir / "assets", SCREENSHOT_NAME)
    if screenshot_path is None:
        screenshot_path = find_file_case_insensitive(ready_dir, SCREENSHOT_NAME)

    return ProductInfo(
        app_name=app_dir.name,
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
        screenshot_path=screenshot_path,
    )


def discover_apps() -> list[AppEntry]:
    apps_root = get_apps_root()
    if not apps_root.exists():
        return []
    entries: list[AppEntry] = []
    try:
        for child in apps_root.iterdir():
            if child.is_dir() and not child.name.startswith("."):
                product_path = find_booth_product_path(child)
                entries.append(AppEntry(child.name, child, product_path is not None))
    except OSError:
        return []

    return sorted(entries, key=lambda entry: (not entry.has_product, entry.name.lower()))


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


def set_first_file_available(page, label: str, path: Path | None, locators: tuple[tuple[str, str], ...]) -> str:
    if path is None or not path.exists():
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
                target.set_input_files(str(path), timeout=3000)
                return UI_TEXT["assist_file_set"].format(label=label)
            except Exception:
                continue
    return UI_TEXT["assist_manual"].format(label=label)


def set_zip_file(page, label: str, path: Path | None) -> str:
    result = set_first_file_available(page, label, path, ZIP_FILE_LOCATORS)
    if result != UI_TEXT["assist_manual"].format(label=label):
        return result
    if path is None or not path.exists():
        return result

    try:
        inputs = page.locator("input[type='file']")
        count = inputs.count()
    except Exception:
        return result

    for index in range(min(count, 8)):
        try:
            target = inputs.nth(index)
            accept = (target.get_attribute("accept", timeout=1000) or "").lower()
            name = (target.get_attribute("name", timeout=1000) or "").lower()
            if "image" in accept or "image" in name or "thumbnail" in name:
                continue
            if accept and not any(marker in accept for marker in ("zip", "application", "octet", "*")):
                continue
            target.set_input_files(str(path), timeout=3000)
            return UI_TEXT["assist_file_set"].format(label=label)
        except Exception:
            continue
    return result


def assist_booth_form(page, product: ProductInfo) -> str:
    results = [
        fill_first_available(page, UI_TEXT["field_title"], product.title, TITLE_LOCATORS),
        fill_first_available(page, UI_TEXT["field_description"], product.description, DESCRIPTION_LOCATORS),
        fill_first_available(page, UI_TEXT["field_price"], price_for_input(product.price), PRICE_LOCATORS),
        fill_first_available(page, UI_TEXT["field_tags"], product.tags, TAG_LOCATORS),
        set_first_file_available(page, UI_TEXT["field_thumbnail"], product.thumbnail_path, THUMBNAIL_FILE_LOCATORS),
        set_zip_file(page, UI_TEXT["field_zip"], product.selected_zip),
        UI_TEXT["assist_publish_guard"],
    ]
    return "\n".join(results)


def safe_page_url(page) -> str:
    try:
        return page.url or ""
    except Exception:
        return ""


def is_booth_login_page(page) -> bool:
    url = safe_page_url(page).lower()
    return "accounts.pixiv.net" in url or ("login" in url and ("booth.pm" in url or "pixiv" in url))


def is_manage_booth_page(page) -> bool:
    return "manage.booth.pm" in safe_page_url(page).lower()


def is_booth_edit_url(url: str) -> bool:
    normalized_url = url.lower()
    return "manage.booth.pm/items/" in normalized_url and "/edit" in normalized_url


def is_booth_edit_page(page) -> bool:
    return is_booth_edit_url(safe_page_url(page))


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


def choose_booth_edit_page(browser):
    edit_pages = []
    for page in collect_browser_pages(browser):
        if not is_booth_edit_page(page):
            continue
        if is_active_page(page):
            return page
        edit_pages.append(page)
    return edit_pages[0] if edit_pages else None

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
            page = choose_booth_edit_page(browser)
            if page is None:
                if has_login_page(browser):
                    events.put(("login_required", UI_TEXT["detail_login_required"]))
                    return
                events.put(("edit_page_required", UI_TEXT["detail_edit_page_required"]))
                return

            events.put(("status", "status_edit_page_found"))
            events.put(("detail", UI_TEXT["detail_chrome_connected"]))
            events.put(("status", "status_assisting"))
            summary = assist_booth_form(page, product)
            events.put(("summary", summary))
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
        self.playwright_active = False
        self.playwright_events: queue.Queue[WorkerEvent] = queue.Queue()

        self.selected_app_var = tk.StringVar(value=UI_TEXT["value_no_selection"])
        self.status_var = tk.StringVar(value=UI_TEXT["status_no_selection"])
        self.status_detail_var = tk.StringVar(value=UI_TEXT["detail_no_selection"])
        self.field_vars = {
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

        buttons = (
            self.launch_chrome_button,
            self.assist_button,
            self.open_ready_button,
            self.copy_title_button,
            self.copy_description_button,
            self.copy_tags_button,
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
            self.app_listbox.insert(tk.END, UI_TEXT[key].format(name=entry.name))

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

    def _render_product(self, product: ProductInfo) -> None:
        self.field_vars["title"].set(product.title or UI_TEXT["value_unset"])
        self.field_vars["price"].set(product.price or UI_TEXT["value_unset"])
        self.field_vars["tags"].set(product.tags or UI_TEXT["value_unset"])
        self.field_vars["product_info"].set(product.product_source or UI_TEXT["value_file_missing"])
        self.field_vars["ready"].set(UI_TEXT["value_yes"] if product.booth_ready_exists else UI_TEXT["value_no"])
        self.field_vars["zip"].set(self._format_zip_value(product))
        self.field_vars["thumbnail"].set(self._format_path_value(product.thumbnail_path))
        self.field_vars["screenshot"].set(self._format_path_value(product.screenshot_path))
        self.field_vars["github_release"].set(product.github_release or UI_TEXT["value_unset"])
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
        self.open_ready_button.configure(state=normal_if_selection)
        self.copy_title_button.configure(state=normal_if_selection)
        self.copy_description_button.configure(state=normal_if_selection)
        self.copy_tags_button.configure(state=normal_if_selection)

    def copy_product_field(self, attribute: str, label_key: str) -> None:
        if self.selected_product is None:
            self._set_status("status_no_selection", UI_TEXT["detail_no_selection"])
            return

        value = getattr(self.selected_product, attribute, "")
        label = UI_TEXT[label_key]
        if not value:
            self._set_status("status_ready", UI_TEXT["detail_copy_empty"].format(label=label))
            return

        self.root.clipboard_clear()
        self.root.clipboard_append(value)
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
        if both_product.title != "Root Sample" or both_product.product_source != PRODUCT_FILE_NAME:
            raise RuntimeError("root product priority fixture failed")

    if not is_booth_edit_url("https://manage.booth.pm/items/8417561/edit"):
        raise RuntimeError("booth edit URL fixture failed")
    items_new_url = "https://manage.booth.pm/items" + "/new"
    if is_booth_edit_url(items_new_url):
        raise RuntimeError("items new URL should not match edit fixture")
    if is_booth_edit_url("https://manage.booth.pm/items/8417561"):
        raise RuntimeError("item URL without edit should not match edit fixture")

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
    dake_backup_source = "missing"
    dake_backup_fields = "missing"
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
            break

    print(f"{APP_NAME} launch-check OK")
    print(f"apps_root={get_apps_root()}")
    print(f"apps={len(apps)}")
    print(f"product_apps={product_apps}")
    print(f"dake_backup_product={dake_backup_source}")
    print(f"dake_backup_fields={dake_backup_fields}")
    print(f"cdp_url={CHROME_CDP_URL}")
    print(f"edit_url_hint={BOOTH_EDIT_URL_HINT}")
    print(f"chrome_path={find_chrome_executable() or 'missing'}")
    print(f"chrome_profile={get_chrome_profile_dir()}")
    print("fixtures=missing_product, missing_ready, multiple_zip, missing_playwright_python, ready_product_lookup, root_product_priority, open_edit_url_only")
    return 0


def main() -> None:
    if "--launch-check" in sys.argv:
        raise SystemExit(run_launch_check())
    app = DakeBoothAssistApp()
    app.run()


if __name__ == "__main__":
    main()
