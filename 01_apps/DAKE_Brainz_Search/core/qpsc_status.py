from __future__ import annotations

from pathlib import Path
from typing import Any

from core.app_config import config_dir, now_iso, write_json_file


STATUS_FILE_NAME = "qpsc_brainz_status.json"
AWAKE_STATUS_MESSAGE = "BRAINZ is awake."


def status_path() -> Path:
    return config_dir() / STATUS_FILE_NAME


def build_brainz_awake_status(started_at: str | None = None) -> dict[str, Any]:
    timestamp = now_iso()
    return {
        "brainz_awake": True,
        "started_at": started_at or timestamp,
        "last_heartbeat_at": timestamp,
        "status_message": AWAKE_STATUS_MESSAGE,
    }


def write_brainz_awake_status(
    path: Path | None = None,
    started_at: str | None = None,
) -> dict[str, Any]:
    payload = build_brainz_awake_status(started_at=started_at)
    write_json_file(path or status_path(), payload)
    return payload
