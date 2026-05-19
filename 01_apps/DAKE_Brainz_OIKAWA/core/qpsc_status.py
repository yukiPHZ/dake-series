from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any

from core.config import app_dir, read_json_file, series_root


STATUS_FILE_NAME = "qpsc_brainz_status.json"
STALE_AFTER_SECONDS = 90


def brainz_status_candidates() -> list[Path]:
    apps_dir = app_dir().parent
    return [
        apps_dir / "DAKE_Brainz_Search" / "data" / "config" / STATUS_FILE_NAME,
        series_root() / "01_apps" / "DAKE_Brainz_Search" / "data" / "config" / STATUS_FILE_NAME,
    ]


def read_brainz_status() -> dict[str, Any]:
    for path in brainz_status_candidates():
        data = read_json_file(path)
        if data:
            return data
    return {}


def is_brainz_awake(data: dict[str, Any]) -> bool:
    if data.get("brainz_awake") is not True:
        return False
    heartbeat = _parse_datetime(str(data.get("last_heartbeat_at", "") or ""))
    if heartbeat is None:
        return False
    age = (datetime.now() - heartbeat).total_seconds()
    return age <= STALE_AFTER_SECONDS


def _parse_datetime(value: str) -> datetime | None:
    try:
        return datetime.fromisoformat(value)
    except ValueError:
        return None
