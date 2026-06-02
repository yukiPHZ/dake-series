from __future__ import annotations

import ctypes
import ctypes.wintypes
import datetime as dt
import hashlib
import json
import math
import os
import random
import re
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
from dataclasses import asdict, dataclass
from pathlib import Path
from tkinter import BOTH, DISABLED, NORMAL, Canvas, Entry, Frame, Label, StringVar, Tk, filedialog, messagebox, ttk


UI_TEXT = {
    "app_name": "DAKE_Note_Inbox",
    "window_title": "DAKE_Note_Inbox",
    "display_name": "note素材受信箱",
    "subtitle": "Slack素材をPEAKHEADZ_ROOTへ置く受信箱",
    "section_status": "同期状態",
    "section_article_candidates": "記事候補",
    "section_settings": "設定",
    "slack_status": "Slack接続状態",
    "last_synced_at": "最終同期日時",
    "sync_count": "同期件数",
    "save_to": "保存先",
    "today_count": "今日の同期件数",
    "tag_count": "札付け件数",
    "ollama_status": "Ollama接続状態",
    "last_tagged_at": "最終札付け日時",
    "not_connected": "未接続",
    "ready": "待機中",
    "syncing": "同期中",
    "connected": "接続済み",
    "failed": "失敗",
    "never": "未実行",
    "button_sync": "今すぐ同期",
    "button_open_obsidian": "Obsidianを開く",
    "button_open_inbox": "INBOXを開く",
    "button_open_notes": "NOTESを開く",
    "button_open_articles": "ARTICLESを開く",
    "button_tag_materials": "札付けする",
    "button_update_candidates": "記事候補更新",
    "button_save_settings": "設定保存",
    "button_browse": "参照",
    "label_token": "Slack Bot Token",
    "label_channel": "Slack Channel ID",
    "label_root": "PEAKHEADZ_ROOT",
    "label_obsidian": "Obsidian実行ファイル",
    "label_interval": "同期間隔（秒）",
    "label_ollama_enabled": "Ollama使用",
    "label_ollama_model": "Ollamaモデル名",
    "ollama_on": "ON",
    "ollama_off": "OFF",
    "settings_saved": "設定を保存しました。",
    "missing_slack": "Slack Bot Token と Slack Channel ID を設定してください。",
    "missing_root": "PEAKHEADZ_ROOT を設定してください。",
    "sync_done": "{count}件を保存しました。",
    "sync_none": "新しいSlack素材はありません。",
    "tagging": "札付け中",
    "tag_done": "{count}件をNOTESへ保存しました。",
    "tag_none": "未処理のINBOX素材はありません。",
    "missing_inbox": "INBOXが見つかりません。",
    "missing_notes": "NOTESが見つかりません。",
    "candidate_updating": "記事候補生成中",
    "candidate_empty": "記事候補はまだありません。",
    "candidate_saved": "記事候補を更新しました。",
    "candidate_title": "タイトル",
    "candidate_reason": "理由",
    "candidate_material_count": "使用素材",
    "candidate_count_unit": "{count}件",
    "candidate_heading": "候補{index}",
    "article_candidates_heading": "記事候補",
    "candidate_default_title": "note素材を読み返す入口",
    "candidate_default_reason": "最近の札付き素材から、Obsidianで巡れるまとまりが見えるため。",
    "open_failed": "開けませんでした: {path}",
    "obsidian_failed": "Obsidianを開けませんでした。設定を確認してください。",
    "obsidian_browse_title": "Obsidian.exeを選択",
    "soft_error": "処理に失敗しました。",
    "filetype_executable": "実行ファイル",
    "markdown_heading": "Slack原文",
    "material_original_heading": "原文",
    "material_tag_heading": "札付け",
    "material_tags_heading": "タグ",
    "material_links_heading": "Obsidianリンク",
    "material_hint_heading": "記事化メモ",
    "tag_note_material": "note素材",
    "fallback_article_hint": "この断片は、note記事の素材として後から読み返せる。",
    "ollama_ok": "ok",
    "ollama_fallback": "fallback",
    "ollama_disabled": "disabled",
    "tray_open": "開く",
    "tray_sync": "今すぐ同期",
    "tray_obsidian": "Obsidianを開く",
    "tray_exit": "終了",
    "self_test_ok": "SELF TEST OK",
    "launch_check_ok": "LAUNCH CHECK OK",
}


APP_NAME = "DAKE_Note_Inbox"
APP_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
DEFAULT_PEAKHEADZ_ROOT = Path.home() / "Documents" / "PEAKHEADZ_ROOT"
SLACK_HISTORY_URL = "https://slack.com/api/conversations.history"
OLLAMA_GENERATE_URL = "http://127.0.0.1:11434/api/generate"
DEFAULT_OLLAMA_MODEL = "qwen2.5:7b"
OBSIDIAN_LINK_CANDIDATES = ["在る", "握らない強さ", "側に", "BORINEF", "熾火", "線", "ワー", "DAKE", "Codex", "Slack", "Obsidian", "note"]
MAX_ARTICLE_CANDIDATES = 3


def appdata_dir() -> Path:
    base = os.getenv("APPDATA")
    if base:
        return Path(base)
    return Path.home() / "AppData" / "Roaming"


CONFIG_PATH = appdata_dir() / APP_NAME / "note_inbox_config.json"


COLORS = {
    "bg": "#09101a",
    "panel": "#111b28",
    "panel_2": "#142033",
    "line": "#27364a",
    "text": "#eaf1f8",
    "muted": "#99a7b7",
    "accent": "#8fb8ff",
    "accent_2": "#9dd7c5",
    "danger": "#ff9a9a",
    "entry": "#0d1624",
}


@dataclass
class AppConfig:
    slack_bot_token: str = ""
    slack_channel_id: str = ""
    peakheadz_root: str = str(DEFAULT_PEAKHEADZ_ROOT)
    obsidian_path: str = ""
    sync_interval_seconds: int = 300
    slack_last_ts: str = ""
    last_synced_at: str = ""
    last_sync_count: int = 0
    today_sync_date: str = ""
    today_sync_count: int = 0
    ollama_enabled: bool = True
    ollama_model: str = DEFAULT_OLLAMA_MODEL
    last_tagged_at: str = ""
    last_tag_count: int = 0
    ollama_status: str = "unknown"


@dataclass
class SlackMessage:
    ts: str
    text: str
    user: str = ""


@dataclass
class MaterialSummary:
    path: Path
    tags: list[str]
    links: list[str]
    article_hint: str


@dataclass
class ArticleCandidate:
    title: str
    reason: str
    material_count: int


class ConfigStore:
    def __init__(self, path: Path = CONFIG_PATH) -> None:
        self.path = path

    def load(self) -> AppConfig:
        source_path = self.path
        if not source_path.exists():
            legacy = legacy_config_path()
            if legacy.exists():
                source_path = legacy
            else:
                return AppConfig()
        try:
            data = json.loads(source_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return AppConfig()
        base = AppConfig()
        for key in asdict(base):
            if key in data:
                setattr(base, key, data[key])
        base.sync_interval_seconds = normalize_interval(base.sync_interval_seconds)
        base.last_sync_count = safe_int(base.last_sync_count, 0)
        base.today_sync_count = safe_int(base.today_sync_count, 0)
        base.ollama_enabled = safe_bool(base.ollama_enabled, True)
        base.ollama_model = str(base.ollama_model or DEFAULT_OLLAMA_MODEL)
        base.last_tag_count = safe_int(base.last_tag_count, 0)
        base.ollama_status = str(base.ollama_status or "unknown")
        if source_path != self.path:
            self.save(base)
        return base

    def save(self, config: AppConfig) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        payload = json.dumps(asdict(config), ensure_ascii=False, indent=2)
        tmp_path = self.path.with_suffix(".tmp")
        tmp_path.write_text(payload + "\n", encoding="utf-8")
        tmp_path.replace(self.path)


def safe_int(value: object, default: int) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def safe_bool(value: object, default: bool) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        normalized = value.strip().lower()
        if normalized in {"1", "true", "yes", "on"}:
            return True
        if normalized in {"0", "false", "no", "off"}:
            return False
    if isinstance(value, int):
        return value != 0
    return default


def normalize_interval(value: object) -> int:
    seconds = safe_int(value, 300)
    if seconds < 0:
        return 0
    return min(seconds, 86_400)


def now_iso() -> str:
    return dt.datetime.now().replace(microsecond=0).isoformat()


def today_key() -> str:
    return dt.date.today().isoformat()


def legacy_config_path() -> Path:
    return APP_DIR / "data" / "note_inbox_config.json"


def common_icon_candidates() -> list[Path]:
    candidates = [
        APP_DIR.parent.parent / "02_assets" / "dake_icon.ico",
        APP_DIR.parent.parent.parent / "02_assets" / "dake_icon.ico",
    ]
    bundle_root = Path(getattr(sys, "_MEIPASS", APP_DIR))
    candidates.append(bundle_root / "assets" / "dake_icon.ico")
    return candidates


def app_icon_path() -> Path | None:
    for candidate in common_icon_candidates():
        if candidate.exists():
            return candidate
    return None


def safe_filename(value: str) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|\s]+', "_", value).strip("._")
    return cleaned[:80] or "slack"


def slack_ts_to_local(ts: str) -> dt.datetime:
    try:
        return dt.datetime.fromtimestamp(float(ts))
    except (TypeError, ValueError, OSError):
        return dt.datetime.now()


def yaml_quote(value: str) -> str:
    return json.dumps(value or "", ensure_ascii=False)


def slack_error_message(error: str) -> str:
    known = {
        "not_authed": "Slack Bot Token が設定されていません。",
        "invalid_auth": "Slack Bot Token が無効です。",
        "channel_not_found": "Slack Channel ID が見つかりません。",
        "not_in_channel": "Bot が対象チャンネルに参加していません。",
        "missing_scope": "Slack Bot Token の権限が不足しています。",
    }
    return known.get(error, f"Slack API error: {error}")


def slack_api_get(token: str, params: dict[str, str], timeout: int = 20) -> dict:
    query = urllib.parse.urlencode(params)
    request = urllib.request.Request(
        f"{SLACK_HISTORY_URL}?{query}",
        headers={
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
        },
        method="GET",
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            body = response.read().decode("utf-8")
    except urllib.error.HTTPError as exc:
        raise RuntimeError(f"HTTP {exc.code}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(str(exc.reason)) from exc
    data = json.loads(body)
    if not data.get("ok"):
        raise RuntimeError(slack_error_message(str(data.get("error", "unknown"))))
    return data


def fetch_slack_messages(config: AppConfig) -> list[SlackMessage]:
    params: dict[str, str] = {
        "channel": config.slack_channel_id,
        "limit": "100",
    }
    if config.slack_last_ts:
        params["oldest"] = config.slack_last_ts
        params["inclusive"] = "false"
    data = slack_api_get(config.slack_bot_token, params)
    messages: list[SlackMessage] = []
    for item in data.get("messages", []):
        text = str(item.get("text", ""))
        ts = str(item.get("ts", ""))
        if not ts or not text:
            continue
        messages.append(SlackMessage(ts=ts, text=text, user=str(item.get("user", ""))))
    messages.sort(key=lambda message: float(message.ts))
    return messages


def target_inbox(root_path: str) -> Path:
    return Path(root_path).expanduser() / "INBOX"


def target_notes(root_path: str) -> Path:
    return Path(root_path).expanduser() / "NOTES"


def target_articles(root_path: str) -> Path:
    return Path(root_path).expanduser() / "ARTICLES"


def article_candidates_path(root_path: str) -> Path:
    return target_articles(root_path) / "article_candidates.md"


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    for index in range(2, 10_000):
        candidate = parent / f"{stem}_{index}{suffix}"
        if not candidate.exists():
            return candidate
    raise RuntimeError("filename collision limit reached")


def strip_yaml_scalar(value: str) -> str:
    value = value.strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in {'"', "'"}:
        return value[1:-1]
    return value


def split_markdown_frontmatter(content: str) -> tuple[dict[str, str], str]:
    if not content.startswith("---"):
        return {}, content
    lines = content.splitlines()
    if not lines or lines[0].strip() != "---":
        return {}, content
    end_index = None
    for index in range(1, len(lines)):
        if lines[index].strip() == "---":
            end_index = index
            break
    if end_index is None:
        return {}, content
    meta: dict[str, str] = {}
    for line in lines[1:end_index]:
        if ":" not in line:
            continue
        key, value = line.split(":", 1)
        meta[key.strip()] = strip_yaml_scalar(value)
    body = "\n".join(lines[end_index + 1 :]).lstrip("\n")
    return meta, body


def split_markdown_frontmatter_with_lists(content: str) -> tuple[dict[str, object], str]:
    if not content.startswith("---"):
        return {}, content
    lines = content.splitlines()
    if not lines or lines[0].strip() != "---":
        return {}, content
    end_index = None
    for index in range(1, len(lines)):
        if lines[index].strip() == "---":
            end_index = index
            break
    if end_index is None:
        return {}, content
    meta: dict[str, object] = {}
    current_list_key = ""
    for line in lines[1:end_index]:
        stripped = line.strip()
        if not stripped:
            continue
        if stripped.startswith("- ") and current_list_key:
            current = meta.setdefault(current_list_key, [])
            if isinstance(current, list):
                current.append(strip_yaml_scalar(stripped[2:]))
            continue
        current_list_key = ""
        if ":" not in line:
            continue
        key, value = line.split(":", 1)
        key = key.strip()
        value = value.strip()
        if value:
            meta[key] = strip_yaml_scalar(value)
        else:
            meta[key] = []
            current_list_key = key
    body = "\n".join(lines[end_index + 1 :]).lstrip("\n")
    return meta, body


def extract_original_text(body: str) -> str:
    lines = body.splitlines()
    headings = {f"# {UI_TEXT['markdown_heading']}", f"# {UI_TEXT['material_original_heading']}"}
    for index, line in enumerate(lines):
        if line.strip() in headings:
            return "\n".join(lines[index + 1 :]).lstrip("\n")
    return body.strip()


def yaml_list(items: list[str]) -> list[str]:
    return [f"  - {yaml_quote(item)}" for item in items]


def normalize_tag(value: str) -> str:
    return value.strip().lstrip("#").replace(" ", "")


def normalize_link(value: str) -> str:
    value = value.strip()
    if value.startswith("[[") and value.endswith("]]"):
        return value
    name = value.strip("[]")
    return f"[[{name}]]" if name else ""


def heuristic_labels(text: str, status: str = "fallback") -> dict[str, object]:
    lowered = text.lower()
    tags = [UI_TEXT["tag_note_material"]]
    selected: list[str] = []
    for candidate in OBSIDIAN_LINK_CANDIDATES:
        if candidate.lower() in lowered or candidate in text:
            selected.append(candidate)
    if "凍結" in text or "作らない" in text or "削る" in text:
        selected.extend(["在る", "握らない強さ"])
    if not selected:
        selected.append("note")
    links: list[str] = []
    for candidate in selected:
        link = f"[[{candidate}]]"
        if link not in links:
            links.append(link)
        if len(links) >= 4:
            break
    for link in links:
        tag = link.strip("[]")
        if tag and tag not in tags:
            tags.append(tag)
        if len(tags) >= 5:
            break
    return {
        "tags": tags,
        "links": links,
        "article_hint": UI_TEXT["fallback_article_hint"],
        "ollama_status": status,
    }


def material_prompt(text: str) -> str:
    candidates = ", ".join(OBSIDIAN_LINK_CANDIDATES)
    trimmed = text[:2400]
    return (
        "あなたはSlack原文に札を貼る係です。記事本文は生成しません。"
        "本文を要約しすぎず、Obsidianで巡るための最小JSONだけを返してください。"
        f"候補リンク: {candidates}\n"
        "必ず次のJSON形式のみで返してください: "
        '{"tags":["note素材"],"links":["[[在る]]"],"article_hint":"短い記事化メモ"}\n'
        f"原文:\n{trimmed}"
    )


def extract_json_object(text: str) -> dict:
    start = text.find("{")
    end = text.rfind("}")
    if start < 0 or end <= start:
        raise ValueError("json object not found")
    return json.loads(text[start : end + 1])


def call_ollama_for_labels(text: str, model: str) -> dict:
    payload = {
        "model": model or DEFAULT_OLLAMA_MODEL,
        "prompt": material_prompt(text),
        "stream": False,
        "options": {"temperature": 0.2},
    }
    request = urllib.request.Request(
        OLLAMA_GENERATE_URL,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=30) as response:
        data = json.loads(response.read().decode("utf-8"))
    return extract_json_object(str(data.get("response", "")))


def normalize_label_result(result: dict, status: str) -> dict[str, object]:
    tags = [UI_TEXT["tag_note_material"]]
    for item in result.get("tags", []):
        if not isinstance(item, str):
            continue
        tag = normalize_tag(item)
        if tag and tag not in tags:
            tags.append(tag)
        if len(tags) >= 6:
            break
    links: list[str] = []
    for item in result.get("links", []):
        if not isinstance(item, str):
            continue
        link = normalize_link(item)
        name = link.strip("[]")
        if link and name in OBSIDIAN_LINK_CANDIDATES and link not in links:
            links.append(link)
        if len(links) >= 5:
            break
    if not links:
        links = [str(item) for item in heuristic_labels("", status)["links"]]
    hint = str(result.get("article_hint", "")).strip()
    if not hint:
        hint = UI_TEXT["fallback_article_hint"]
    return {
        "tags": tags,
        "links": links,
        "article_hint": hint[:180],
        "ollama_status": status,
    }


def labels_for_text(text: str, config: AppConfig) -> dict[str, object]:
    if not config.ollama_enabled:
        return heuristic_labels(text, UI_TEXT["ollama_disabled"])
    try:
        return normalize_label_result(call_ollama_for_labels(text, config.ollama_model), UI_TEXT["ollama_ok"])
    except Exception:
        return heuristic_labels(text, UI_TEXT["ollama_fallback"])


def material_output_path(notes: Path, raw_path: Path) -> Path:
    source_id = hashlib.sha1(str(raw_path).encode("utf-8")).hexdigest()[:8]
    return notes / f"{safe_filename(raw_path.stem)}_{source_id}_material.md"


def markdown_for_material(raw_path: Path, meta: dict[str, str], original_text: str, labels: dict[str, object]) -> str:
    tags = [str(item) for item in labels.get("tags", [UI_TEXT["tag_note_material"]])]
    links = [str(item) for item in labels.get("links", [])]
    hint = str(labels.get("article_hint", UI_TEXT["fallback_article_hint"]))
    ollama_status = str(labels.get("ollama_status", UI_TEXT["ollama_fallback"]))
    source = meta.get("source", "slack") or "slack"
    slack_ts = meta.get("slack_ts", meta.get("timestamp", ""))
    tag_line = " ".join(f"#{normalize_tag(tag)}" for tag in tags if normalize_tag(tag))
    return "\n".join(
        [
            "---",
            f"source: {yaml_quote(source)}",
            "status: material",
            f"original_path: {yaml_quote(str(raw_path))}",
            f"slack_ts: {yaml_quote(slack_ts)}",
            "tags:",
            *yaml_list(tags),
            "links:",
            *yaml_list(links),
            f"article_hint: {yaml_quote(hint)}",
            f"ollama_status: {yaml_quote(ollama_status)}",
            "---",
            "",
            f"# {UI_TEXT['material_original_heading']}",
            "",
            original_text.rstrip(),
            "",
            "---",
            "",
            f"# {UI_TEXT['material_tag_heading']}",
            "",
            f"## {UI_TEXT['material_tags_heading']}",
            "",
            tag_line,
            "",
            f"## {UI_TEXT['material_links_heading']}",
            "",
            "\n".join(links),
            "",
            f"## {UI_TEXT['material_hint_heading']}",
            "",
            hint,
            "",
        ]
    )


def is_raw_markdown(meta: dict[str, str]) -> bool:
    return meta.get("status", "").strip().lower() == "raw"


def save_material_from_raw(raw_path: Path, notes: Path, config: AppConfig) -> tuple[Path | None, str]:
    target = material_output_path(notes, raw_path)
    if target.exists():
        return None, "skipped"
    content = raw_path.read_text(encoding="utf-8")
    meta, body = split_markdown_frontmatter(content)
    if not is_raw_markdown(meta):
        return None, "skipped"
    original_text = extract_original_text(body)
    labels = labels_for_text(original_text, config)
    notes.mkdir(parents=True, exist_ok=True)
    target.write_text(markdown_for_material(raw_path, meta, original_text, labels), encoding="utf-8", newline="\n")
    return target, str(labels.get("ollama_status", UI_TEXT["ollama_fallback"]))


def tag_inbox_materials(config: AppConfig, limit: int = 100) -> tuple[int, str]:
    inbox = target_inbox(config.peakheadz_root)
    if not inbox.exists():
        raise FileNotFoundError(UI_TEXT["missing_inbox"])
    notes = target_notes(config.peakheadz_root)
    count = 0
    statuses: list[str] = []
    for raw_path in sorted(inbox.rglob("*.md")):
        if count >= limit:
            break
        saved, status = save_material_from_raw(raw_path, notes, config)
        if saved:
            count += 1
            statuses.append(status)
    if not statuses:
        return 0, config.ollama_status or "unknown"
    if UI_TEXT["ollama_ok"] in statuses:
        return count, UI_TEXT["ollama_ok"]
    if UI_TEXT["ollama_fallback"] in statuses:
        return count, UI_TEXT["ollama_fallback"]
    return count, statuses[-1]


def string_list(value: object) -> list[str]:
    if not isinstance(value, list):
        return []
    result: list[str] = []
    for item in value:
        if isinstance(item, str) and item.strip():
            result.append(item.strip())
    return result


def material_summary_from_path(path: Path) -> MaterialSummary | None:
    content = path.read_text(encoding="utf-8")
    meta, _body = split_markdown_frontmatter_with_lists(content)
    if str(meta.get("status", "")).strip().lower() != "material":
        return None
    tags = string_list(meta.get("tags"))
    links = string_list(meta.get("links"))
    hint = str(meta.get("article_hint", "")).strip()
    if not tags and not links and not hint:
        return None
    return MaterialSummary(path=path, tags=tags, links=links, article_hint=hint)


def collect_material_summaries(root_path: str) -> list[MaterialSummary]:
    notes = target_notes(root_path)
    if not notes.exists():
        raise FileNotFoundError(UI_TEXT["missing_notes"])
    summaries: list[MaterialSummary] = []
    for path in sorted(notes.rglob("*.md")):
        summary = material_summary_from_path(path)
        if summary:
            summaries.append(summary)
    return summaries


def article_candidate_prompt(summaries: list[MaterialSummary]) -> str:
    lines = []
    for index, summary in enumerate(summaries[:80], 1):
        lines.append(
            "\n".join(
                [
                    f"素材{index}",
                    "tags: " + ", ".join(summary.tags[:8]),
                    "links: " + ", ".join(summary.links[:8]),
                    "article_hint: " + summary.article_hint[:180],
                ]
            )
        )
    return (
        "あなたはnote記事本文を書かない編集補助です。"
        "以下の札付き素材から、記事候補を最大3件だけJSONで返してください。"
        "ランキング、score、本文生成は禁止です。"
        "原文本文は渡していません。tags, links, article_hintだけで判断してください。"
        '形式: {"candidates":[{"title":"短いタイトル","reason":"短い理由","material_count":3}]}\n'
        + "\n---\n".join(lines)
    )


def call_ollama_for_article_candidates(summaries: list[MaterialSummary], model: str) -> dict:
    payload = {
        "model": model or DEFAULT_OLLAMA_MODEL,
        "prompt": article_candidate_prompt(summaries),
        "stream": False,
        "options": {"temperature": 0.25},
    }
    request = urllib.request.Request(
        OLLAMA_GENERATE_URL,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=35) as response:
        data = json.loads(response.read().decode("utf-8"))
    return extract_json_object(str(data.get("response", "")))


def clamp_material_count(value: object, total: int) -> int:
    count = safe_int(value, 1)
    if total <= 0:
        return 0
    return max(1, min(count, total))


def normalize_article_candidates(data: dict, summaries: list[MaterialSummary]) -> list[ArticleCandidate]:
    raw_candidates = data.get("candidates", [])
    if not isinstance(raw_candidates, list):
        return []
    candidates: list[ArticleCandidate] = []
    for item in raw_candidates[:MAX_ARTICLE_CANDIDATES]:
        if not isinstance(item, dict):
            continue
        title = str(item.get("title", "")).strip()
        reason = str(item.get("reason", "")).strip()
        if not title or not reason:
            continue
        candidates.append(
            ArticleCandidate(
                title=title[:80],
                reason=reason[:220],
                material_count=clamp_material_count(item.get("material_count", 1), len(summaries)),
            )
        )
    return candidates[:MAX_ARTICLE_CANDIDATES]


def fallback_title_for_key(key: str) -> str:
    titles = {
        "在る": "在るをめぐる断片",
        "握らない強さ": "握らない強さを残す",
        "DAKE": "DAKEを入口にする素材整理",
        "Slack": "Slackに投げた断片を育てる",
        "Obsidian": "Obsidianで巡るnote素材",
        "note": "note素材を読み返す入口",
    }
    return titles.get(key, f"{key}から始めるnote素材")


def fallback_article_candidates(summaries: list[MaterialSummary]) -> list[ArticleCandidate]:
    if not summaries:
        return []
    buckets: dict[str, int] = {}
    for summary in summaries:
        keys = summary.links + summary.tags
        seen: set[str] = set()
        for key in keys:
            normalized = normalize_link(key).strip("[]") if key.startswith("[[") else normalize_tag(key)
            if not normalized or normalized in seen or normalized == UI_TEXT["tag_note_material"]:
                continue
            seen.add(normalized)
            buckets[normalized] = buckets.get(normalized, 0) + 1
    if not buckets:
        buckets["note"] = len(summaries)
    ordered = sorted(buckets.items(), key=lambda item: (-item[1], item[0]))[:MAX_ARTICLE_CANDIDATES]
    candidates: list[ArticleCandidate] = []
    for key, count in ordered:
        candidates.append(
            ArticleCandidate(
                title=fallback_title_for_key(key),
                reason=f"NOTESの素材で「{key}」につながる札やメモがまとまっているため。",
                material_count=count,
            )
        )
    return candidates[:MAX_ARTICLE_CANDIDATES]


def generate_article_candidates(config: AppConfig) -> list[ArticleCandidate]:
    summaries = collect_material_summaries(config.peakheadz_root)
    if not summaries:
        return []
    if config.ollama_enabled:
        try:
            candidates = normalize_article_candidates(call_ollama_for_article_candidates(summaries, config.ollama_model), summaries)
            if candidates:
                return candidates[:MAX_ARTICLE_CANDIDATES]
        except Exception:
            pass
    return fallback_article_candidates(summaries)


def markdown_for_article_candidates(candidates: list[ArticleCandidate]) -> str:
    lines = [f"# {UI_TEXT['article_candidates_heading']}", ""]
    for index, candidate in enumerate(candidates[:MAX_ARTICLE_CANDIDATES], 1):
        lines.extend(
            [
                f"## {UI_TEXT['candidate_heading'].format(index=index)}",
                "",
                UI_TEXT["candidate_title"],
                "",
                candidate.title,
                "",
                UI_TEXT["candidate_reason"],
                "",
                candidate.reason,
                "",
                UI_TEXT["candidate_material_count"],
                "",
                UI_TEXT["candidate_count_unit"].format(count=candidate.material_count),
                "",
                "---",
                "",
            ]
        )
    return "\n".join(lines).rstrip() + "\n"


def save_article_candidates(config: AppConfig, candidates: list[ArticleCandidate]) -> Path:
    articles = target_articles(config.peakheadz_root)
    articles.mkdir(parents=True, exist_ok=True)
    path = article_candidates_path(config.peakheadz_root)
    path.write_text(markdown_for_article_candidates(candidates), encoding="utf-8", newline="\n")
    return path


def markdown_for_slack_message(message: SlackMessage, channel_id: str) -> str:
    local_time = slack_ts_to_local(message.ts).strftime("%Y-%m-%d %H:%M:%S")
    return "\n".join(
        [
            "---",
            "source: slack",
            f"channel_id: {yaml_quote(channel_id)}",
            f"timestamp: {yaml_quote(local_time)}",
            f"slack_ts: {yaml_quote(message.ts)}",
            "status: raw",
            "---",
            "",
            f"# {UI_TEXT['markdown_heading']}",
            "",
            message.text,
            "",
        ]
    )


def save_slack_message(root_path: str, channel_id: str, message: SlackMessage) -> Path:
    inbox = target_inbox(root_path)
    inbox.mkdir(parents=True, exist_ok=True)
    stamp = slack_ts_to_local(message.ts).strftime("%Y%m%d_%H%M%S")
    filename = f"{stamp}_{safe_filename(message.ts)}.md"
    path = unique_path(inbox / filename)
    path.write_text(markdown_for_slack_message(message, channel_id), encoding="utf-8", newline="\n")
    return path


def open_path(path: Path) -> None:
    resolved = path.expanduser()
    if not resolved.exists():
        raise FileNotFoundError(str(resolved))
    if os.name == "nt":
        os.startfile(str(resolved))  # type: ignore[attr-defined]
    else:
        subprocess.Popen(["xdg-open", str(resolved)])


def localappdata_dir() -> Path:
    base = os.getenv("LOCALAPPDATA")
    if base:
        return Path(base)
    return Path.home() / "AppData" / "Local"


def obsidian_launcher_candidates() -> list[Path]:
    return [
        localappdata_dir() / "Programs" / "Obsidian" / "Obsidian.exe",
        appdata_dir() / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Obsidian.lnk",
    ]


def find_obsidian_launcher() -> Path | None:
    for candidate in obsidian_launcher_candidates():
        if candidate.exists():
            return candidate
    return None


def obsidian_vault_uri(root_path: Path) -> str:
    vault_name = root_path.expanduser().name or "PEAKHEADZ_ROOT"
    return "obsidian://open?vault=" + urllib.parse.quote(vault_name)


def open_obsidian_with_launcher(launcher: Path, root_path: Path) -> None:
    launcher = launcher.expanduser()
    root_path = root_path.expanduser()
    suffix = launcher.suffix.lower()
    if suffix == ".exe":
        subprocess.Popen([str(launcher), str(root_path)])
        return
    if suffix == ".lnk" and os.name == "nt":
        os.startfile(str(launcher))  # type: ignore[attr-defined]
        webbrowser.open(obsidian_vault_uri(root_path))
        return
    raise FileNotFoundError(str(launcher))


class WindowsTrayIcon:
    def __init__(self, app: "NoteInboxApp") -> None:
        self.app = app
        self.active = False
        self.thread: threading.Thread | None = None
        self.hwnd = None
        self._class_atom = None
        self._callback = None

    def start(self) -> None:
        if os.name != "nt" or self.active:
            return
        self.thread = threading.Thread(target=self._run, name="note-inbox-tray", daemon=True)
        self.thread.start()

    def stop(self) -> None:
        if os.name != "nt" or not self.hwnd:
            return
        try:
            self._delete_icon()
            ctypes.windll.user32.PostMessageW(self.hwnd, 0x0010, 0, 0)
        except Exception:
            pass

    def _run(self) -> None:
        try:
            self._message_loop()
        except Exception:
            self.active = False

    def _message_loop(self) -> None:
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        kernel32.GetModuleHandleW.restype = ctypes.c_void_p
        user32.CreateWindowExW.restype = ctypes.c_void_p
        user32.CreateWindowExW.argtypes = [
            ctypes.wintypes.DWORD,
            ctypes.wintypes.LPCWSTR,
            ctypes.wintypes.LPCWSTR,
            ctypes.wintypes.DWORD,
            ctypes.c_int,
            ctypes.c_int,
            ctypes.c_int,
            ctypes.c_int,
            ctypes.c_void_p,
            ctypes.c_void_p,
            ctypes.c_void_p,
            ctypes.c_void_p,
        ]
        user32.DefWindowProcW.restype = ctypes.c_ssize_t
        user32.DefWindowProcW.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_size_t, ctypes.c_size_t]
        user32.RegisterClassW.restype = ctypes.c_ushort
        user32.TrackPopupMenu.restype = ctypes.c_uint

        WNDPROC = ctypes.WINFUNCTYPE(ctypes.c_ssize_t, ctypes.c_void_p, ctypes.c_uint, ctypes.c_size_t, ctypes.c_size_t)

        class WNDCLASS(ctypes.Structure):
            _fields_ = [
                ("style", ctypes.c_uint),
                ("lpfnWndProc", WNDPROC),
                ("cbClsExtra", ctypes.c_int),
                ("cbWndExtra", ctypes.c_int),
                ("hInstance", ctypes.c_void_p),
                ("hIcon", ctypes.c_void_p),
                ("hCursor", ctypes.c_void_p),
                ("hbrBackground", ctypes.c_void_p),
                ("lpszMenuName", ctypes.c_wchar_p),
                ("lpszClassName", ctypes.c_wchar_p),
            ]

        def wnd_proc(hwnd, msg, wparam, lparam):
            if msg == 0x0400 + 20:
                if lparam in (0x0205, 0x0206):
                    self._show_menu(hwnd)
                elif lparam == 0x0203:
                    self.app.root.after(0, self.app.show_window)
                return 0
            if msg == 0x0010:
                self._delete_icon()
                user32.DestroyWindow(hwnd)
                return 0
            if msg == 0x0002:
                user32.PostQuitMessage(0)
                return 0
            return user32.DefWindowProcW(hwnd, msg, wparam, lparam)

        self._callback = WNDPROC(wnd_proc)
        instance = kernel32.GetModuleHandleW(None)
        class_name = f"{APP_NAME}_TrayWindow"
        wndclass = WNDCLASS()
        wndclass.lpfnWndProc = self._callback
        wndclass.hInstance = instance
        wndclass.lpszClassName = class_name
        self._class_atom = user32.RegisterClassW(ctypes.byref(wndclass))
        hwnd = user32.CreateWindowExW(0, class_name, class_name, 0, 0, 0, 0, 0, None, None, instance, None)
        if not hwnd:
            return
        self.hwnd = hwnd
        self._add_icon(hwnd)
        self.active = True
        msg = ctypes.wintypes.MSG()
        while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))
        self.active = False

    def _notify_data(self, hwnd):
        class GUID(ctypes.Structure):
            _fields_ = [
                ("Data1", ctypes.c_ulong),
                ("Data2", ctypes.c_ushort),
                ("Data3", ctypes.c_ushort),
                ("Data4", ctypes.c_ubyte * 8),
            ]

        class NOTIFYICONDATA(ctypes.Structure):
            _fields_ = [
                ("cbSize", ctypes.c_ulong),
                ("hWnd", ctypes.c_void_p),
                ("uID", ctypes.c_uint),
                ("uFlags", ctypes.c_uint),
                ("uCallbackMessage", ctypes.c_uint),
                ("hIcon", ctypes.c_void_p),
                ("szTip", ctypes.c_wchar * 128),
                ("dwState", ctypes.c_ulong),
                ("dwStateMask", ctypes.c_ulong),
                ("szInfo", ctypes.c_wchar * 256),
                ("uTimeoutOrVersion", ctypes.c_uint),
                ("szInfoTitle", ctypes.c_wchar * 64),
                ("dwInfoFlags", ctypes.c_ulong),
                ("guidItem", GUID),
                ("hBalloonIcon", ctypes.c_void_p),
            ]

        shell32 = ctypes.windll.shell32
        user32 = ctypes.windll.user32
        user32.LoadImageW.restype = ctypes.c_void_p
        user32.LoadIconW.restype = ctypes.c_void_p
        icon_handle = None
        icon_path = app_icon_path()
        if icon_path:
            icon_handle = user32.LoadImageW(None, str(icon_path), 1, 0, 0, 0x00000010 | 0x00000040)
        if not icon_handle:
            icon_handle = user32.LoadIconW(None, 32512)
        data = NOTIFYICONDATA()
        data.cbSize = ctypes.sizeof(NOTIFYICONDATA)
        data.hWnd = hwnd
        data.uID = 1
        data.uFlags = 0x00000001 | 0x00000002 | 0x00000004
        data.uCallbackMessage = 0x0400 + 20
        data.hIcon = icon_handle
        data.szTip = UI_TEXT["app_name"]
        return shell32, data

    def _add_icon(self, hwnd) -> None:
        shell32, data = self._notify_data(hwnd)
        shell32.Shell_NotifyIconW(0x00000000, ctypes.byref(data))

    def _delete_icon(self) -> None:
        if not self.hwnd:
            return
        shell32, data = self._notify_data(self.hwnd)
        shell32.Shell_NotifyIconW(0x00000002, ctypes.byref(data))

    def _show_menu(self, hwnd) -> None:
        user32 = ctypes.windll.user32
        menu = user32.CreatePopupMenu()
        commands = [
            (1001, UI_TEXT["tray_open"], self.app.show_window),
            (1002, UI_TEXT["tray_sync"], self.app.sync_now),
            (1003, UI_TEXT["tray_obsidian"], self.app.open_obsidian),
            (1004, UI_TEXT["tray_exit"], self.app.exit_app),
        ]
        for command_id, label, _callback in commands:
            user32.AppendMenuW(menu, 0x00000000, command_id, label)
        point = ctypes.wintypes.POINT()
        user32.GetCursorPos(ctypes.byref(point))
        user32.SetForegroundWindow(hwnd)
        selected = user32.TrackPopupMenu(menu, 0x0100, point.x, point.y, 0, hwnd, None)
        user32.DestroyMenu(menu)
        for command_id, _label, callback in commands:
            if selected == command_id:
                self.app.root.after(0, callback)
                break


class StarField:
    def __init__(self, canvas: Canvas, count: int = 38) -> None:
        self.canvas = canvas
        self.count = count
        self.stars: list[dict[str, float | int]] = []
        self.lines: list[int] = []
        self.running = True
        self.resize_after_id: str | None = None
        self.canvas.bind("<Configure>", self._on_resize, add="+")
        self._create_stars()
        self._tick()

    def _create_stars(self) -> None:
        self.resize_after_id = None
        self.canvas.update_idletasks()
        width = max(self.canvas.winfo_width(), 920)
        height = max(self.canvas.winfo_height(), 620)
        for star in self.stars:
            self.canvas.delete(int(star["item"]))
        for line in self.lines:
            self.canvas.delete(line)
        self.stars.clear()
        self.lines.clear()
        for _index in range(self.count):
            x = float(random.randint(12, width - 12))
            y = float(random.randint(12, height - 12))
            size = random.choice([1, 1, 2])
            phase = random.random() * 6.28
            speed = random.uniform(0.25, 0.55)
            item = self.canvas.create_oval(x, y, x + size, y + size, fill="#53667f", outline="")
            self.stars.append(
                {
                    "item": item,
                    "x": x,
                    "y": y,
                    "size": float(size),
                    "phase": phase,
                    "speed": speed,
                    "vx": random.uniform(-0.18, 0.18),
                    "vy": random.uniform(-0.12, 0.12),
                }
            )

    def _on_resize(self, _event) -> None:
        if self.resize_after_id:
            self.canvas.after_cancel(self.resize_after_id)
        self.resize_after_id = self.canvas.after(350, self._create_stars)

    def _tick(self) -> None:
        if not self.running:
            return
        current = time.time()
        palette = ["#405066", "#53667f", "#7288a8", "#9fb5d6"]
        width = max(self.canvas.winfo_width(), 920)
        height = max(self.canvas.winfo_height(), 620)
        for star in self.stars:
            x = float(star["x"]) + float(star["vx"])
            y = float(star["y"]) + float(star["vy"])
            if x < 8 or x > width - 8:
                star["vx"] = -float(star["vx"])
                x = max(8, min(width - 8, x))
            if y < 8 or y > height - 8:
                star["vy"] = -float(star["vy"])
                y = max(8, min(height - 8, y))
            star["x"] = x
            star["y"] = y
            size = float(star["size"])
            value = int((1 + math.sin(current * float(star["speed"]) + float(star["phase"]))) * 1.5)
            color = palette[max(0, min(value, len(palette) - 1))]
            item = int(star["item"])
            self.canvas.coords(item, x, y, x + size, y + size)
            self.canvas.itemconfigure(item, fill=color)
        self._draw_connections()
        self.canvas.after(180, self._tick)

    def _draw_connections(self) -> None:
        for line in self.lines:
            self.canvas.delete(line)
        self.lines.clear()
        max_lines = 54
        max_distance = 118.0
        for index, star in enumerate(self.stars):
            if len(self.lines) >= max_lines:
                break
            x1 = float(star["x"])
            y1 = float(star["y"])
            for other in self.stars[index + 1 :]:
                dx = float(other["x"]) - x1
                dy = float(other["y"]) - y1
                distance = math.sqrt(dx * dx + dy * dy)
                if distance <= max_distance:
                    tone = "#27364a" if distance > 72 else "#334a68"
                    line = self.canvas.create_line(x1, y1, float(other["x"]), float(other["y"]), fill=tone, width=1)
                    self.canvas.tag_lower(line)
                    self.lines.append(line)
                    if len(self.lines) >= max_lines:
                        break


class NoteInboxApp:
    def __init__(self, root: Tk, store: ConfigStore | None = None) -> None:
        self.root = root
        self.store = store or ConfigStore()
        self.config = self.store.load()
        if not self.config.obsidian_path:
            launcher = find_obsidian_launcher()
            if launcher and launcher.suffix.lower() == ".exe":
                self.config.obsidian_path = str(launcher)
        self.sync_lock = threading.Lock()
        self.tag_lock = threading.Lock()
        self.candidate_lock = threading.Lock()
        self.status_vars: dict[str, StringVar] = {}
        self.entry_vars: dict[str, StringVar] = {}
        self.check_vars: dict[str, StringVar] = {}
        self.candidate_vars: list[dict[str, StringVar]] = []
        self.candidate_frames: list[Frame] = []
        self.candidate_empty_var: StringVar | None = None
        self.toggle_buttons: dict[str, ttk.Button] = {}
        self.form_row = 0
        self.buttons: list[ttk.Button] = []
        self.tray = WindowsTrayIcon(self)
        self.auto_sync_after_id: str | None = None
        self._setup_window()
        self._build_ui()
        self._refresh_status()
        self.tray.start()
        self._schedule_auto_sync()

    def _setup_window(self) -> None:
        self.root.title(UI_TEXT["window_title"])
        self.root.geometry("980x720")
        self.root.minsize(680, 500)
        icon_path = app_icon_path()
        if icon_path:
            try:
                self.root.iconbitmap(str(icon_path))
            except Exception:
                pass
        self.root.configure(bg=COLORS["bg"])
        self.root.protocol("WM_DELETE_WINDOW", self.hide_to_tray)
        self.root.bind("<Unmap>", self._on_unmap)
        self.root.after(80, self._maximize_window)

    def _maximize_window(self) -> None:
        try:
            self.root.state("zoomed")
        except Exception:
            self.root.attributes("-zoomed", True)

    def _build_ui(self) -> None:
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "TButton",
            padding=(14, 8),
            font=("Yu Gothic UI", 10),
            borderwidth=0,
            relief="flat",
            background=COLORS["panel_2"],
            foreground=COLORS["text"],
            focuscolor=COLORS["panel_2"],
        )
        style.map(
            "TButton",
            background=[("active", "#1d3049"), ("disabled", "#182333")],
            foreground=[("disabled", COLORS["muted"])],
        )
        style.configure(
            "Accent.TButton",
            padding=(18, 9),
            font=("Yu Gothic UI", 10, "bold"),
            borderwidth=0,
            relief="flat",
            background="#2563eb",
            foreground="#f8fbff",
            focuscolor="#2563eb",
        )
        style.map("Accent.TButton", background=[("active", "#3b82f6"), ("disabled", "#284267")])
        style.configure("Browse.TButton", padding=(10, 5), font=("Yu Gothic UI", 9), borderwidth=0, relief="flat")
        style.configure(
            "ToggleOn.TButton",
            padding=(18, 6),
            font=("Yu Gothic UI", 10, "bold"),
            borderwidth=0,
            relief="flat",
            background="#2f8f75",
            foreground="#f8fffc",
            focuscolor="#2f8f75",
        )
        style.map("ToggleOn.TButton", background=[("active", "#36a789"), ("disabled", "#24483f")])
        style.configure(
            "ToggleOff.TButton",
            padding=(18, 6),
            font=("Yu Gothic UI", 10, "bold"),
            borderwidth=0,
            relief="flat",
            background="#263143",
            foreground=COLORS["muted"],
            focuscolor="#263143",
        )
        style.map("ToggleOff.TButton", background=[("active", "#303c52"), ("disabled", "#1b2534")])

        canvas = Canvas(self.root, bg=COLORS["bg"], highlightthickness=0)
        canvas.pack(fill=BOTH, expand=True)
        StarField(canvas)

        frame = Frame(canvas, bg=COLORS["bg"])
        window = canvas.create_window(0, 18, anchor="n", window=frame)

        def resize_content(event) -> None:
            content_width = min(max(event.width - 72, 680), 980)
            canvas.itemconfigure(window, width=content_width)
            canvas.coords(window, event.width / 2, 18)

        canvas.bind("<Configure>", resize_content, add="+")

        title = Label(frame, text=UI_TEXT["app_name"], bg=COLORS["bg"], fg=COLORS["text"], font=("Yu Gothic UI", 20, "bold"))
        title.pack(anchor="center", pady=(24, 8))
        subtitle = Label(frame, text=UI_TEXT["subtitle"], bg=COLORS["bg"], fg=COLORS["muted"], font=("Yu Gothic UI", 10))
        subtitle.pack(anchor="center", pady=(0, 12))

        status_panel = Frame(frame, bg=COLORS["panel"], highlightbackground=COLORS["line"], highlightthickness=1)
        status_panel.pack(fill="x", pady=(8, 14))
        Label(status_panel, text=UI_TEXT["section_status"], bg=COLORS["panel"], fg=COLORS["accent"], font=("Yu Gothic UI", 12, "bold")).pack(anchor="w", padx=18, pady=(14, 8))

        grid = Frame(status_panel, bg=COLORS["panel"])
        grid.pack(fill="x", padx=18, pady=(0, 16))
        status_items = [
            ("slack_status", UI_TEXT["slack_status"]),
            ("last_synced_at", UI_TEXT["last_synced_at"]),
            ("sync_count", UI_TEXT["sync_count"]),
            ("today_count", UI_TEXT["today_count"]),
            ("tag_count", UI_TEXT["tag_count"]),
            ("ollama_status", UI_TEXT["ollama_status"]),
            ("last_tagged_at", UI_TEXT["last_tagged_at"]),
            ("save_to", UI_TEXT["save_to"]),
        ]
        for row, (key, label_text) in enumerate(status_items):
            Label(grid, text=label_text, bg=COLORS["panel"], fg=COLORS["muted"], font=("Yu Gothic UI", 9)).grid(row=row, column=0, sticky="w", pady=3)
            var = StringVar(value="")
            self.status_vars[key] = var
            Label(grid, textvariable=var, bg=COLORS["panel"], fg=COLORS["text"], font=("Yu Gothic UI", 10), wraplength=540, justify="left").grid(row=row, column=1, sticky="w", padx=(18, 0), pady=3)
        grid.columnconfigure(1, weight=1)

        candidate_panel = Frame(frame, bg=COLORS["panel"], highlightbackground=COLORS["line"], highlightthickness=1)
        candidate_panel.pack(fill="x", pady=(0, 14))
        candidate_header = Frame(candidate_panel, bg=COLORS["panel"])
        candidate_header.pack(fill="x", padx=18, pady=(14, 8))
        Label(candidate_header, text=UI_TEXT["section_article_candidates"], bg=COLORS["panel"], fg=COLORS["accent"], font=("Yu Gothic UI", 12, "bold")).pack(side="left")
        self._add_button(candidate_header, UI_TEXT["button_update_candidates"], self.update_article_candidates_now)

        candidate_body = Frame(candidate_panel, bg=COLORS["panel"])
        candidate_body.pack(fill="x", padx=18, pady=(0, 16))
        self.candidate_empty_var = StringVar(value=UI_TEXT["candidate_empty"])
        Label(candidate_body, textvariable=self.candidate_empty_var, bg=COLORS["panel"], fg=COLORS["muted"], font=("Yu Gothic UI", 10)).pack(anchor="w", pady=(0, 6))
        for index in range(MAX_ARTICLE_CANDIDATES):
            card = Frame(candidate_body, bg=COLORS["panel_2"], highlightbackground=COLORS["line"], highlightthickness=1)
            card.pack(fill="x", pady=(6, 0))
            vars_for_card = {
                "heading": StringVar(value=UI_TEXT["candidate_heading"].format(index=index + 1)),
                "title": StringVar(value=""),
                "reason": StringVar(value=""),
                "count": StringVar(value=""),
            }
            self.candidate_vars.append(vars_for_card)
            self.candidate_frames.append(card)
            Label(card, textvariable=vars_for_card["heading"], bg=COLORS["panel_2"], fg=COLORS["accent_2"], font=("Yu Gothic UI", 10, "bold")).pack(anchor="w", padx=14, pady=(10, 3))
            Label(card, textvariable=vars_for_card["title"], bg=COLORS["panel_2"], fg=COLORS["text"], font=("Yu Gothic UI", 11, "bold"), wraplength=820, justify="left").pack(anchor="w", padx=14, pady=2)
            Label(card, textvariable=vars_for_card["reason"], bg=COLORS["panel_2"], fg=COLORS["muted"], font=("Yu Gothic UI", 9), wraplength=820, justify="left").pack(anchor="w", padx=14, pady=2)
            Label(card, textvariable=vars_for_card["count"], bg=COLORS["panel_2"], fg=COLORS["accent"], font=("Yu Gothic UI", 9, "bold")).pack(anchor="w", padx=14, pady=(2, 10))
            card.pack_forget()

        button_panel = Frame(frame, bg=COLORS["bg"])
        button_panel.pack(anchor="center", pady=(0, 14))
        self._add_button(button_panel, UI_TEXT["button_sync"], self.sync_now, "Accent.TButton")
        self._add_button(button_panel, UI_TEXT["button_tag_materials"], self.tag_materials_now)
        self._add_button(button_panel, UI_TEXT["button_open_obsidian"], self.open_obsidian)
        self._add_button(button_panel, UI_TEXT["button_open_inbox"], self.open_inbox)
        self._add_button(button_panel, UI_TEXT["button_open_notes"], self.open_notes)
        self._add_button(button_panel, UI_TEXT["button_open_articles"], self.open_articles)

        settings_panel = Frame(frame, bg=COLORS["panel"], highlightbackground=COLORS["line"], highlightthickness=1)
        settings_panel.pack(fill="x", pady=(0, 24))
        Label(settings_panel, text=UI_TEXT["section_settings"], bg=COLORS["panel"], fg=COLORS["accent_2"], font=("Yu Gothic UI", 12, "bold")).pack(anchor="w", padx=18, pady=(14, 8))

        form = Frame(settings_panel, bg=COLORS["panel"])
        form.pack(fill="x", padx=18, pady=(0, 12))
        self._add_entry(form, "slack_bot_token", UI_TEXT["label_token"], self.config.slack_bot_token, show="*")
        self._add_entry(form, "slack_channel_id", UI_TEXT["label_channel"], self.config.slack_channel_id)
        self._add_entry(form, "peakheadz_root", UI_TEXT["label_root"], self.config.peakheadz_root)
        self._add_entry(form, "obsidian_path", UI_TEXT["label_obsidian"], self.config.obsidian_path, browse_command=self.browse_obsidian_exe)
        self._add_entry(form, "sync_interval_seconds", UI_TEXT["label_interval"], str(self.config.sync_interval_seconds))
        self._add_check(form, "ollama_enabled", UI_TEXT["label_ollama_enabled"], self.config.ollama_enabled)
        self._add_entry(form, "ollama_model", UI_TEXT["label_ollama_model"], self.config.ollama_model)
        form.columnconfigure(1, weight=1)

        save_row = Frame(settings_panel, bg=COLORS["panel"])
        save_row.pack(anchor="center", pady=(0, 16))
        self._add_button(save_row, UI_TEXT["button_save_settings"], self.save_settings, "Accent.TButton")

    def _add_button(self, parent: Frame, text: str, command, style_name: str = "TButton") -> None:
        button = ttk.Button(parent, text=text, command=command, style=style_name)
        button.pack(side="left", padx=(0, 8), pady=4)
        self.buttons.append(button)

    def _add_entry(self, parent: Frame, key: str, label_text: str, value: str, show: str | None = None, browse_command=None) -> None:
        row = self.form_row
        self.form_row += 1
        Label(parent, text=label_text, bg=COLORS["panel"], fg=COLORS["muted"], font=("Yu Gothic UI", 9)).grid(row=row, column=0, sticky="w", pady=5)
        var = StringVar(value=value)
        self.entry_vars[key] = var
        entry = Entry(
            parent,
            textvariable=var,
            show=show,
            bg=COLORS["entry"],
            fg=COLORS["text"],
            insertbackground=COLORS["text"],
            relief="flat",
            highlightbackground=COLORS["line"],
            highlightcolor=COLORS["accent"],
            highlightthickness=1,
            font=("Yu Gothic UI", 10),
        )
        entry.grid(row=row, column=1, sticky="ew", padx=(18, 0), pady=5, ipady=5)
        if browse_command:
            button = ttk.Button(parent, text=UI_TEXT["button_browse"], command=browse_command, style="Browse.TButton")
            button.grid(row=row, column=2, sticky="e", padx=(8, 0), pady=5)
            self.buttons.append(button)

    def _add_check(self, parent: Frame, key: str, label_text: str, value: bool) -> None:
        row = self.form_row
        self.form_row += 1
        Label(parent, text=label_text, bg=COLORS["panel"], fg=COLORS["muted"], font=("Yu Gothic UI", 9)).grid(row=row, column=0, sticky="w", pady=5)
        var = StringVar(value="1" if value else "0")
        self.check_vars[key] = var
        button = ttk.Button(parent, command=lambda: self._toggle_setting(key))
        button.grid(row=row, column=1, sticky="w", padx=(18, 0), pady=5)
        self.toggle_buttons[key] = button
        self.buttons.append(button)
        self._refresh_toggle(key)

    def _toggle_setting(self, key: str) -> None:
        var = self.check_vars[key]
        var.set("0" if var.get() == "1" else "1")
        self._refresh_toggle(key)

    def _refresh_toggle(self, key: str) -> None:
        button = self.toggle_buttons.get(key)
        var = self.check_vars.get(key)
        if not button or not var:
            return
        enabled = var.get() == "1"
        button.configure(
            text=UI_TEXT["ollama_on"] if enabled else UI_TEXT["ollama_off"],
            style="ToggleOn.TButton" if enabled else "ToggleOff.TButton",
        )

    def _refresh_status(self) -> None:
        connected = UI_TEXT["connected"] if self.config.slack_bot_token and self.config.slack_channel_id else UI_TEXT["not_connected"]
        self.status_vars["slack_status"].set(connected)
        self.status_vars["last_synced_at"].set(self.config.last_synced_at or UI_TEXT["never"])
        self.status_vars["sync_count"].set(str(self.config.last_sync_count))
        self.status_vars["today_count"].set(str(self.config.today_sync_count))
        self.status_vars["tag_count"].set(str(self.config.last_tag_count))
        self.status_vars["ollama_status"].set(self.config.ollama_status or "unknown")
        self.status_vars["last_tagged_at"].set(self.config.last_tagged_at or UI_TEXT["never"])
        self.status_vars["save_to"].set(str(target_inbox(self.config.peakheadz_root)))

    def _set_buttons(self, state: str) -> None:
        for button in self.buttons:
            button.configure(state=state)

    def _on_unmap(self, _event) -> None:
        if self.root.state() == "iconic":
            self.root.after(120, self.hide_to_tray)

    def hide_to_tray(self) -> None:
        self.root.withdraw()

    def show_window(self) -> None:
        self.root.deiconify()
        self._maximize_window()
        self.root.lift()
        self.root.focus_force()

    def exit_app(self) -> None:
        if self.auto_sync_after_id:
            self.root.after_cancel(self.auto_sync_after_id)
            self.auto_sync_after_id = None
        self.tray.stop()
        self.root.after(150, self.root.destroy)

    def save_settings(self) -> None:
        self.config.slack_bot_token = self.entry_vars["slack_bot_token"].get().strip()
        self.config.slack_channel_id = self.entry_vars["slack_channel_id"].get().strip()
        self.config.peakheadz_root = self.entry_vars["peakheadz_root"].get().strip()
        self.config.obsidian_path = self.entry_vars["obsidian_path"].get().strip()
        self.config.sync_interval_seconds = normalize_interval(self.entry_vars["sync_interval_seconds"].get().strip())
        self.config.ollama_enabled = self.check_vars["ollama_enabled"].get() == "1"
        self.config.ollama_model = self.entry_vars["ollama_model"].get().strip() or DEFAULT_OLLAMA_MODEL
        self.entry_vars["sync_interval_seconds"].set(str(self.config.sync_interval_seconds))
        self.entry_vars["ollama_model"].set(self.config.ollama_model)
        self.store.save(self.config)
        self._refresh_status()
        self._schedule_auto_sync()
        messagebox.showinfo(UI_TEXT["app_name"], UI_TEXT["settings_saved"])

    def sync_now(self, show_dialog: bool = True) -> None:
        self.save_settings_without_dialog()
        if not self.config.slack_bot_token or not self.config.slack_channel_id:
            if show_dialog:
                messagebox.showwarning(UI_TEXT["app_name"], UI_TEXT["missing_slack"])
            return
        if not self.config.peakheadz_root:
            if show_dialog:
                messagebox.showwarning(UI_TEXT["app_name"], UI_TEXT["missing_root"])
            return
        if not self.sync_lock.acquire(blocking=False):
            return
        self.status_vars["slack_status"].set(UI_TEXT["syncing"])
        self._set_buttons(DISABLED)
        threading.Thread(target=self._sync_worker, args=(show_dialog,), name="note-inbox-sync", daemon=True).start()

    def save_settings_without_dialog(self) -> None:
        self.config.slack_bot_token = self.entry_vars["slack_bot_token"].get().strip()
        self.config.slack_channel_id = self.entry_vars["slack_channel_id"].get().strip()
        self.config.peakheadz_root = self.entry_vars["peakheadz_root"].get().strip()
        self.config.obsidian_path = self.entry_vars["obsidian_path"].get().strip()
        self.config.sync_interval_seconds = normalize_interval(self.entry_vars["sync_interval_seconds"].get().strip())
        self.config.ollama_enabled = self.check_vars["ollama_enabled"].get() == "1"
        self.config.ollama_model = self.entry_vars["ollama_model"].get().strip() or DEFAULT_OLLAMA_MODEL
        self.store.save(self.config)

    def _sync_worker(self, show_dialog: bool) -> None:
        try:
            messages = fetch_slack_messages(self.config)
            saved_paths: list[Path] = []
            for message in messages:
                saved_paths.append(save_slack_message(self.config.peakheadz_root, self.config.slack_channel_id, message))
            if messages:
                self.config.slack_last_ts = messages[-1].ts
            now = now_iso()
            if self.config.today_sync_date != today_key():
                self.config.today_sync_date = today_key()
                self.config.today_sync_count = 0
            self.config.last_sync_count = len(saved_paths)
            self.config.today_sync_count += len(saved_paths)
            self.config.last_synced_at = now
            self.store.save(self.config)
            self.root.after(0, lambda: self._sync_complete(len(saved_paths), None, show_dialog))
        except Exception as exc:
            self.root.after(0, lambda: self._sync_complete(0, str(exc), show_dialog))
        finally:
            self.sync_lock.release()

    def _sync_complete(self, count: int, error: str | None, show_dialog: bool) -> None:
        self._set_buttons(NORMAL)
        if error:
            self.status_vars["slack_status"].set(UI_TEXT["failed"])
            if show_dialog:
                messagebox.showerror(UI_TEXT["app_name"], f"{UI_TEXT['soft_error']}\n{error}")
        else:
            self.status_vars["slack_status"].set(UI_TEXT["connected"])
            message = UI_TEXT["sync_done"].format(count=count) if count else UI_TEXT["sync_none"]
            if show_dialog:
                messagebox.showinfo(UI_TEXT["app_name"], message)
        self._refresh_status()

    def tag_materials_now(self) -> None:
        self.save_settings_without_dialog()
        if not self.config.peakheadz_root:
            messagebox.showwarning(UI_TEXT["app_name"], UI_TEXT["missing_root"])
            return
        if not self.tag_lock.acquire(blocking=False):
            return
        self.status_vars["ollama_status"].set(UI_TEXT["tagging"])
        self._set_buttons(DISABLED)
        threading.Thread(target=self._tag_worker, name="note-inbox-tagging", daemon=True).start()

    def _tag_worker(self) -> None:
        try:
            count, status = tag_inbox_materials(self.config)
            self.config.last_tag_count = count
            self.config.ollama_status = status
            self.config.last_tagged_at = now_iso()
            self.store.save(self.config)
            self.root.after(0, lambda: self._tag_complete(count, status, None))
        except Exception as exc:
            self.config.ollama_status = UI_TEXT["ollama_fallback"]
            self.store.save(self.config)
            self.root.after(0, lambda: self._tag_complete(0, UI_TEXT["ollama_fallback"], str(exc)))
        finally:
            self.tag_lock.release()

    def _tag_complete(self, count: int, status: str, error: str | None) -> None:
        self._set_buttons(NORMAL)
        if error:
            self.status_vars["ollama_status"].set(status)
            messagebox.showerror(UI_TEXT["app_name"], f"{UI_TEXT['soft_error']}\n{error}")
        else:
            self.status_vars["ollama_status"].set(status)
            message = UI_TEXT["tag_done"].format(count=count) if count else UI_TEXT["tag_none"]
            messagebox.showinfo(UI_TEXT["app_name"], message)
        self._refresh_status()

    def update_article_candidates_now(self) -> None:
        self.save_settings_without_dialog()
        if not self.config.peakheadz_root:
            messagebox.showwarning(UI_TEXT["app_name"], UI_TEXT["missing_root"])
            return
        if not self.candidate_lock.acquire(blocking=False):
            return
        if self.candidate_empty_var:
            self.candidate_empty_var.set(UI_TEXT["candidate_updating"])
        self._set_buttons(DISABLED)
        threading.Thread(target=self._candidate_worker, name="note-inbox-candidates", daemon=True).start()

    def _candidate_worker(self) -> None:
        try:
            candidates = generate_article_candidates(self.config)
            path = save_article_candidates(self.config, candidates)
            self.root.after(0, lambda: self._candidate_complete(candidates, path, None))
        except Exception as exc:
            error = str(exc)
            self.root.after(0, lambda: self._candidate_complete([], None, error))
        finally:
            self.candidate_lock.release()

    def _candidate_complete(self, candidates: list[ArticleCandidate], path: Path | None, error: str | None) -> None:
        self._set_buttons(NORMAL)
        if error:
            if self.candidate_empty_var:
                self.candidate_empty_var.set(UI_TEXT["candidate_empty"])
            messagebox.showerror(UI_TEXT["app_name"], f"{UI_TEXT['soft_error']}\n{error}")
            return
        self._render_article_candidates(candidates)
        if self.candidate_empty_var:
            self.candidate_empty_var.set(UI_TEXT["candidate_saved"] if path else UI_TEXT["candidate_empty"])

    def _render_article_candidates(self, candidates: list[ArticleCandidate]) -> None:
        for index, frame in enumerate(self.candidate_frames):
            if index >= len(candidates):
                frame.pack_forget()
                continue
            candidate = candidates[index]
            vars_for_card = self.candidate_vars[index]
            vars_for_card["heading"].set(UI_TEXT["candidate_heading"].format(index=index + 1))
            vars_for_card["title"].set(candidate.title)
            vars_for_card["reason"].set(candidate.reason)
            count_text = UI_TEXT["candidate_count_unit"].format(count=candidate.material_count)
            vars_for_card["count"].set(f"{UI_TEXT['candidate_material_count']}: {count_text}")
            frame.pack(fill="x", pady=(6, 0))

    def _schedule_auto_sync(self) -> None:
        if self.auto_sync_after_id:
            self.root.after_cancel(self.auto_sync_after_id)
            self.auto_sync_after_id = None
        seconds = normalize_interval(self.config.sync_interval_seconds)
        if seconds <= 0:
            return
        self.auto_sync_after_id = self.root.after(seconds * 1000, self._auto_sync)

    def _auto_sync(self) -> None:
        self.auto_sync_after_id = None
        if self.config.slack_bot_token and self.config.slack_channel_id:
            self.sync_now(show_dialog=False)
        self._schedule_auto_sync()

    def browse_obsidian_exe(self) -> None:
        raw_path = self.entry_vars["obsidian_path"].get().strip()
        initial = Path(raw_path).expanduser() if raw_path else None
        initial_dir = str(initial.parent) if initial and initial.exists() else str(localappdata_dir() / "Programs" / "Obsidian")
        selected = filedialog.askopenfilename(
            title=UI_TEXT["obsidian_browse_title"],
            initialdir=initial_dir,
            filetypes=[(UI_TEXT["filetype_executable"], "*.exe"), ("All files", "*.*")],
        )
        if selected:
            self.entry_vars["obsidian_path"].set(selected)

    def open_obsidian(self) -> None:
        self.save_settings_without_dialog()
        root_path = Path(self.config.peakheadz_root).expanduser()
        configured = Path(self.config.obsidian_path).expanduser() if self.config.obsidian_path else None
        launcher = configured if configured and configured.exists() else find_obsidian_launcher()
        try:
            if launcher:
                open_obsidian_with_launcher(launcher, root_path)
            else:
                webbrowser.open(obsidian_vault_uri(root_path))
        except Exception:
            try:
                if webbrowser.open(obsidian_vault_uri(root_path)):
                    return
            except Exception:
                pass
            messagebox.showerror(UI_TEXT["app_name"], UI_TEXT["obsidian_failed"])

    def open_inbox(self) -> None:
        self._open_root_child("INBOX")

    def open_notes(self) -> None:
        self._open_root_child("NOTES")

    def open_articles(self) -> None:
        self._open_root_child("ARTICLES")

    def _open_root_child(self, child: str) -> None:
        self.save_settings_without_dialog()
        path = Path(self.config.peakheadz_root).expanduser() / child
        try:
            open_path(path)
        except Exception:
            messagebox.showerror(UI_TEXT["app_name"], UI_TEXT["open_failed"].format(path=path))


def run_self_test() -> int:
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        store = ConfigStore(tmp_path / "data" / "note_inbox_config.json")
        config = AppConfig(
            slack_bot_token="test-token-placeholder",
            slack_channel_id="C123456",
            peakheadz_root=str(tmp_path / "PEAKHEADZ_ROOT"),
            obsidian_path="",
            sync_interval_seconds=60,
        )
        store.save(config)
        loaded = store.load()
        if loaded.slack_channel_id != "C123456":
            raise AssertionError("config restore failed")
        message = SlackMessage(ts="1700000000.000100", text="hello from slack")
        saved = save_slack_message(loaded.peakheadz_root, loaded.slack_channel_id, message)
        body = saved.read_text(encoding="utf-8")
        if "source: slack" not in body or "hello from slack" not in body:
            raise AssertionError("markdown save failed")
        raw_before = saved.read_text(encoding="utf-8")
        original_call = globals()["call_ollama_for_labels"]
        try:
            globals()["call_ollama_for_labels"] = lambda _text, _model: (_ for _ in ()).throw(RuntimeError("offline"))
            loaded.ollama_enabled = True
            loaded.ollama_model = DEFAULT_OLLAMA_MODEL
            count, status = tag_inbox_materials(loaded)
        finally:
            globals()["call_ollama_for_labels"] = original_call
        if count != 1 or status != UI_TEXT["ollama_fallback"]:
            raise AssertionError("fallback tagging failed")
        if saved.read_text(encoding="utf-8") != raw_before:
            raise AssertionError("raw markdown changed")
        material_files = list(target_notes(loaded.peakheadz_root).glob("*_material.md"))
        if len(material_files) != 1:
            raise AssertionError("material markdown not created")
        material = material_files[0].read_text(encoding="utf-8")
        if "status: material" not in material or "ollama_status" not in material or "[[" not in material:
            raise AssertionError("material markdown format failed")
        loaded.ollama_enabled = False
        candidates = generate_article_candidates(loaded)
        if not candidates or len(candidates) > MAX_ARTICLE_CANDIDATES:
            raise AssertionError("article candidate generation failed")
        candidate_path = save_article_candidates(loaded, candidates)
        candidate_markdown = candidate_path.read_text(encoding="utf-8")
        if "# " not in candidate_markdown or "## " not in candidate_markdown:
            raise AssertionError("article candidate markdown failed")
        too_many = [
            ArticleCandidate(title=f"title {index}", reason=f"reason {index}", material_count=index)
            for index in range(1, 5)
        ]
        capped_markdown = markdown_for_article_candidates(too_many)
        if UI_TEXT["candidate_heading"].format(index=4) in capped_markdown:
            raise AssertionError("article candidate cap failed")
        if "score" in capped_markdown.lower() or "ranking" in capped_markdown.lower():
            raise AssertionError("article candidate forbidden text failed")
    print(UI_TEXT["self_test_ok"])
    return 0


def run_launch_check() -> int:
    ConfigStore().load()
    if not UI_TEXT["app_name"]:
        raise AssertionError("UI_TEXT is empty")
    print(UI_TEXT["launch_check_ok"])
    return 0


def main() -> int:
    if "--self-test" in sys.argv:
        return run_self_test()
    if "--launch-check" in sys.argv:
        return run_launch_check()
    root = Tk()
    NoteInboxApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
