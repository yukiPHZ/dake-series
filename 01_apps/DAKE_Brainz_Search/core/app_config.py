from __future__ import annotations

import json
import os
import sys
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Any


APP_FOLDER_NAME = "DAKE_Brainz_Search"
CONFIG_FILE_NAME = "brainz_config.json"
DATA_DIR_NAME = "data"
LOG_DIR_NAME = "logs"
EXPORT_DIR_NAME = "exports"
CONFIG_DIR_NAME = "config"
DB_FILE_NAME = "brainz.db"


@dataclass
class AppConfig:
    memory_folder: str = ""
    watch_folder: str = ""
    auto_index_enabled: bool = False
    remote_queue_folder: str = ""
    enable_remote_queue: bool = False
    auto_run_remote_search: bool = False
    codex_reports_folder: str = ""
    enable_slack_inbox: bool = False
    slack_bot_token: str = ""
    slack_channel_id: str = ""
    slack_poll_interval_seconds: int = 10
    slack_last_ts: str = ""
    enable_aru_inbox: bool = False
    aru_slack_token: str = ""
    aru_channel_id: str = ""
    aru_poll_interval_seconds: int = 10
    aru_last_ts: str = ""
    slack_notify_enabled: bool = False
    slack_webhook_url: str = ""
    slack_notify_max_per_day: int = 3
    slack_notify_quiet_hours: str = "22:00-04:00"
    enable_notifications: bool = True
    last_query: str = ""
    last_indexed_at: str = ""


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        if exe_dir.name.lower() == "dist":
            return exe_dir.parent
        return exe_dir
    return Path(__file__).resolve().parent.parent


def series_root() -> Path:
    return app_dir().parent.parent


def data_dir() -> Path:
    return app_dir() / DATA_DIR_NAME


def logs_dir() -> Path:
    return data_dir() / LOG_DIR_NAME


def exports_dir() -> Path:
    return data_dir() / EXPORT_DIR_NAME


def config_dir() -> Path:
    return data_dir() / CONFIG_DIR_NAME


def db_path() -> Path:
    return data_dir() / DB_FILE_NAME


def assets_dir() -> Path:
    return app_dir() / "assets"


def peakheadz_logo_path() -> Path:
    return assets_dir() / "peakheadz_logo.png"


def peakheadz_icon_path() -> Path:
    return assets_dir() / "peakheadz_logo.ico"


def common_icon_path() -> Path:
    return series_root() / "02_assets" / "dake_icon.ico"


def ensure_app_dirs() -> None:
    for path in (data_dir(), logs_dir(), exports_dir(), config_dir(), assets_dir()):
        path.mkdir(parents=True, exist_ok=True)


def read_json_file(path: Path) -> dict[str, Any]:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    if isinstance(data, dict):
        return data
    return {}


def write_json_file(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_suffix(path.suffix + ".tmp")
    temp_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    temp_path.replace(path)


class ConfigStore:
    def __init__(self, path: Path | None = None) -> None:
        ensure_app_dirs()
        self.path = path or (config_dir() / CONFIG_FILE_NAME)

    def load(self) -> AppConfig:
        data = read_json_file(self.path)
        return AppConfig(
            memory_folder=str(data.get("memory_folder", "") or ""),
            watch_folder=str(data.get("watch_folder", "") or ""),
            auto_index_enabled=bool(data.get("auto_index_enabled", False)),
            remote_queue_folder=str(data.get("remote_queue_folder", "") or ""),
            enable_remote_queue=bool(data.get("enable_remote_queue", False)),
            auto_run_remote_search=bool(data.get("auto_run_remote_search", False)),
            codex_reports_folder=str(data.get("codex_reports_folder", "") or ""),
            enable_slack_inbox=bool(data.get("enable_slack_inbox", False)),
            slack_bot_token=str(data.get("slack_bot_token", "") or ""),
            slack_channel_id=str(data.get("slack_channel_id", "") or ""),
            slack_poll_interval_seconds=parse_int(data.get("slack_poll_interval_seconds", 10), default=10, minimum=5, maximum=15),
            slack_last_ts=str(data.get("slack_last_ts", "") or ""),
            enable_aru_inbox=bool(data.get("enable_aru_inbox", False)),
            aru_slack_token=str(data.get("aru_slack_token", "") or ""),
            aru_channel_id=str(data.get("aru_channel_id", "") or ""),
            aru_poll_interval_seconds=parse_int(data.get("aru_poll_interval_seconds", 10), default=10, minimum=5, maximum=15),
            aru_last_ts=str(data.get("aru_last_ts", "") or ""),
            slack_notify_enabled=bool(data.get("slack_notify_enabled", False)),
            slack_webhook_url=str(data.get("slack_webhook_url", "") or ""),
            slack_notify_max_per_day=parse_int(data.get("slack_notify_max_per_day", 3), default=3, minimum=1, maximum=5),
            slack_notify_quiet_hours=str(data.get("slack_notify_quiet_hours", "22:00-04:00") or "22:00-04:00"),
            enable_notifications=bool(data.get("enable_notifications", True)),
            last_query=str(data.get("last_query", "") or ""),
            last_indexed_at=str(data.get("last_indexed_at", "") or ""),
        )

    def save(self, config: AppConfig) -> None:
        payload = asdict(config)
        payload["updated_at"] = now_iso()
        write_json_file(self.path, payload)


def parse_int(value: Any, default: int, minimum: int, maximum: int) -> int:
    try:
        parsed = int(str(value).strip())
    except (TypeError, ValueError):
        return default
    return max(minimum, min(maximum, parsed))


def read_text_safe(path: Path) -> str:
    for encoding in ("utf-8-sig", "utf-8", "cp932", "utf-16"):
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
        except OSError:
            return ""
    try:
        return path.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return ""


def open_path(path: Path) -> None:
    if sys.platform.startswith("win") and hasattr(os, "startfile"):
        os.startfile(str(path))  # type: ignore[attr-defined]
        return

    import webbrowser

    webbrowser.open(path.resolve().as_uri())
