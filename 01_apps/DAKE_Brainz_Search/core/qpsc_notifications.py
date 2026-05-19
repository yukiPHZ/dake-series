from __future__ import annotations

import json
import uuid
from pathlib import Path
from typing import Any

from core.app_config import config_dir, now_iso


NOTIFICATIONS_FILE_NAME = "qpsc_notifications.json"
MAX_EVENTS = 100

UI_TEXT = {
    "title_chatgpt_export": "ChatGPT exportを取り込みました",
    "title_codex_report": "Codex報告を保存しました",
    "title_codex_result": "Codex結果を保存しました",
    "title_paste_import": "ペースト投稿を取り込みました",
    "title_slack_import": "Slackから取り込みました",
    "message_saved_count": "{count}件の記憶を保存しました。",
}


def notifications_path(path: Path | None = None) -> Path:
    return path or (config_dir() / NOTIFICATIONS_FILE_NAME)


def read_notification_events(path: Path | None = None) -> list[dict[str, Any]]:
    target = notifications_path(path)
    try:
        data = json.loads(target.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    if not isinstance(data, list):
        return []
    return [item for item in data if isinstance(item, dict)]


def write_notification_events(events: list[dict[str, Any]], path: Path | None = None) -> Path:
    target = notifications_path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    temp_path = target.with_suffix(target.suffix + ".tmp")
    temp_path.write_text(json.dumps(events[-MAX_EVENTS:], ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    temp_path.replace(target)
    return target


def append_import_notification(
    source: str,
    title: str,
    message: str,
    related_path: str = "",
    path: Path | None = None,
) -> dict[str, Any]:
    event = {
        "id": f"{now_iso().replace(':', '').replace('-', '').replace('T', '_')}_{uuid.uuid4().hex[:8]}",
        "created_at": now_iso(),
        "source": source,
        "title": title,
        "message": message,
        "status": "unread",
        "kind": "import",
        "related_path": related_path,
    }
    events = read_notification_events(path)
    events.append(event)
    write_notification_events(events, path)
    return event


def append_saved_count_notification(
    source: str,
    title: str,
    count: int,
    related_path: str = "",
    path: Path | None = None,
) -> dict[str, Any] | None:
    if count <= 0:
        return None
    try:
        return append_import_notification(
            source=source,
            title=title,
            message=UI_TEXT["message_saved_count"].format(count=count),
            related_path=related_path,
            path=path,
        )
    except OSError:
        return None
