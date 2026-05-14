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
            last_query=str(data.get("last_query", "") or ""),
            last_indexed_at=str(data.get("last_indexed_at", "") or ""),
        )

    def save(self, config: AppConfig) -> None:
        payload = asdict(config)
        payload["updated_at"] = now_iso()
        write_json_file(self.path, payload)


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
