from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from core.config import app_dir, series_root


NOTIFICATIONS_FILE_NAME = "qpsc_notifications.json"


@dataclass(frozen=True)
class QpscNotification:
    id: str
    created_at: str
    source: str
    title: str
    message: str
    status: str
    kind: str
    related_path: str


def brainz_notification_candidates() -> list[Path]:
    apps_dir = app_dir().parent
    return [
        apps_dir / "DAKE_Brainz_Search" / "data" / "config" / NOTIFICATIONS_FILE_NAME,
        series_root() / "01_apps" / "DAKE_Brainz_Search" / "data" / "config" / NOTIFICATIONS_FILE_NAME,
    ]


def read_qpsc_notification_events(limit: int | None = None) -> list[QpscNotification]:
    events = _read_events(_notification_path())
    events.sort(key=lambda event: str(event.get("created_at", "") or ""), reverse=True)
    if limit is not None:
        events = events[:limit]
    return [_to_notification(event) for event in events]


def read_qpsc_notifications(limit: int = 3) -> list[QpscNotification]:
    unread = [event for event in read_qpsc_notification_events() if event.status == "unread"]
    return unread[:limit]


def mark_qpsc_notification_read(notification_id: str) -> bool:
    target = _notification_path()
    events = _read_events(target)
    changed = False
    for event in events:
        if str(event.get("id", "") or "") == notification_id:
            event["status"] = "read"
            changed = True
            break
    if changed:
        try:
            _write_events(target, events)
        except OSError:
            return False
    return changed


def _notification_path() -> Path:
    candidates = brainz_notification_candidates()
    for path in candidates:
        if path.exists():
            return path
    return candidates[0]


def _read_events(path: Path) -> list[dict[str, Any]]:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    if not isinstance(data, list):
        return []
    return [item for item in data if isinstance(item, dict)]


def _write_events(path: Path, events: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_suffix(path.suffix + ".tmp")
    temp_path.write_text(json.dumps(events, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    temp_path.replace(path)


def _to_notification(event: dict[str, Any]) -> QpscNotification:
    return QpscNotification(
        id=str(event.get("id", "") or ""),
        created_at=str(event.get("created_at", "") or ""),
        source=str(event.get("source", "") or ""),
        title=str(event.get("title", "") or ""),
        message=str(event.get("message", "") or ""),
        status=str(event.get("status", "") or ""),
        kind=str(event.get("kind", "") or ""),
        related_path=str(event.get("related_path", "") or ""),
    )
