# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import sys
from dataclasses import asdict, dataclass
from pathlib import Path


APP_FOLDER_NAME = "DAKE_Backup"


@dataclass
class AppSettings:
    source_folder: str = ""
    destination_folder: str = ""
    last_saved_at: str = ""


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def data_dir() -> Path:
    return app_dir() / "data"


def logs_dir() -> Path:
    return data_dir() / "logs"


def settings_path() -> Path:
    return data_dir() / "settings.json"


def ensure_data_files() -> None:
    data_dir().mkdir(parents=True, exist_ok=True)
    logs_dir().mkdir(parents=True, exist_ok=True)
    if not settings_path().exists():
        save_settings(AppSettings())


def load_settings() -> AppSettings:
    ensure_data_files()
    try:
        raw = json.loads(settings_path().read_text(encoding="utf-8"))
    except Exception:
        return AppSettings()
    return AppSettings(
        source_folder=str(raw.get("source_folder", "") or ""),
        destination_folder=str(raw.get("destination_folder", "") or ""),
        last_saved_at=str(raw.get("last_saved_at", "") or ""),
    )


def save_settings(settings: AppSettings) -> None:
    data_dir().mkdir(parents=True, exist_ok=True)
    logs_dir().mkdir(parents=True, exist_ok=True)
    settings_path().write_text(
        json.dumps(asdict(settings), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

