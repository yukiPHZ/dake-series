# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import os
import sys
from dataclasses import asdict, dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

from app.constants import APP_DATA_DIR_NAME, USER_DATA_DIR_NAME


@dataclass
class AppConfig:
    save_root: str
    mic_device: str = ""
    mic_label: str = ""
    sensitivity: str = "standard"
    silence_timeout_sec: int = 30
    transcription_quality: str = "standard"
    summary_enabled: bool = False
    auto_start_windows: bool = False
    auto_delete_days: int = 0
    first_run_complete: bool = False
    tray_notice_shown: bool = False


@dataclass(frozen=True)
class RuntimePaths:
    app_data: Path
    config_path: Path
    state_db: Path
    logs: Path
    spool: Path
    models: Path
    save_root: Path


def documents_dir() -> Path:
    return Path.home() / "Documents"


def default_save_root() -> Path:
    return documents_dir() / USER_DATA_DIR_NAME


def app_data_dir() -> Path:
    base = os.environ.get("LOCALAPPDATA")
    if base:
        return Path(base) / APP_DATA_DIR_NAME
    return Path.home() / f".{APP_DATA_DIR_NAME}"


def default_config() -> AppConfig:
    return AppConfig(save_root=str(default_save_root()))


def load_config(config_path: Path) -> AppConfig:
    config = default_config()
    if not config_path.is_file():
        return config
    try:
        payload = json.loads(config_path.read_text(encoding="utf-8"))
    except Exception:
        return config
    if not isinstance(payload, dict):
        return config
    values = asdict(config)
    values.update({key: value for key, value in payload.items() if key in values})
    try:
        values["silence_timeout_sec"] = int(values.get("silence_timeout_sec") or 30)
        values["auto_delete_days"] = int(values.get("auto_delete_days") or 0)
    except Exception:
        values["silence_timeout_sec"] = 30
        values["auto_delete_days"] = 0
    if values["silence_timeout_sec"] not in (10, 30, 60):
        values["silence_timeout_sec"] = 30
    if values["sensitivity"] not in ("low", "standard", "high"):
        values["sensitivity"] = "standard"
    if values["transcription_quality"] not in ("light", "standard"):
        values["transcription_quality"] = "standard"
    if values["auto_delete_days"] not in (0, 30, 90):
        values["auto_delete_days"] = 0
    return AppConfig(**values)


def save_config(config_path: Path, config: AppConfig) -> None:
    config_path.parent.mkdir(parents=True, exist_ok=True)
    config_path.write_text(
        json.dumps(asdict(config), ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def runtime_paths(config: AppConfig | None = None) -> RuntimePaths:
    app_data = app_data_dir()
    config_path = app_data / "config.json"
    active_config = config or load_config(config_path)
    save_root = Path(active_config.save_root).expanduser()
    return RuntimePaths(
        app_data=app_data,
        config_path=config_path,
        state_db=app_data / "state.db",
        logs=app_data / "logs",
        spool=app_data / "spool",
        models=app_data / "models",
        save_root=save_root,
    )


def ensure_runtime_paths(paths: RuntimePaths) -> None:
    for path in (paths.app_data, paths.logs, paths.spool, paths.models, paths.save_root):
        path.mkdir(parents=True, exist_ok=True)


def day_dir(save_root: Path, record_date: date | str) -> Path:
    day = record_date if isinstance(record_date, str) else record_date.isoformat()
    return save_root / day


def audio_dir(save_root: Path, record_date: date | str) -> Path:
    return day_dir(save_root, record_date) / "audio"


def transcript_dir(save_root: Path, record_date: date | str) -> Path:
    return day_dir(save_root, record_date) / "transcript"


def ensure_day_dirs(save_root: Path, record_date: date | str) -> None:
    audio_dir(save_root, record_date).mkdir(parents=True, exist_ok=True)
    transcript_dir(save_root, record_date).mkdir(parents=True, exist_ok=True)


def clip_stem(started_at: datetime, existing_names: set[str] | None = None) -> str:
    base = started_at.strftime("%Y%m%d_%H%M%S")
    existing = existing_names or set()
    if base not in existing:
        return base
    milliseconds = f"{started_at.microsecond // 1000:03d}"
    candidate = f"{base}_{milliseconds}"
    if candidate not in existing:
        return candidate
    for index in range(2, 1000):
        candidate = f"{base}_{milliseconds}_{index:02d}"
        if candidate not in existing:
            return candidate
    raise RuntimeError("Could not create a unique clip name.")


def unique_audio_path(save_root: Path, started_at: datetime) -> Path:
    record_date = started_at.date()
    ensure_day_dirs(save_root, record_date)
    names = {path.stem for path in audio_dir(save_root, record_date).glob("*.wav")}
    return audio_dir(save_root, record_date) / f"{clip_stem(started_at, names)}.wav"


def transcript_path_for_audio(save_root: Path, audio_path: Path, record_date: date | str) -> Path:
    ensure_day_dirs(save_root, record_date)
    return transcript_dir(save_root, record_date) / f"{audio_path.stem}.txt"


def daily_transcript_path(save_root: Path, record_date: date | str) -> Path:
    return day_dir(save_root, record_date) / "daily_transcript.md"


def daily_summary_path(save_root: Path, record_date: date | str) -> Path:
    return day_dir(save_root, record_date) / "daily_summary.md"


def app_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[2]


def icon_candidates() -> list[Path]:
    base = app_base_dir()
    source = Path(__file__).resolve().parents[2]
    return [
        base / "dake_icon.ico",
        source / "dake_icon.ico",
        source / ".." / ".." / "02_assets" / "dake_icon.ico",
        base / ".." / ".." / "02_assets" / "dake_icon.ico",
        Path("C:/Users/yukiz/devlop/DAKE_series/02_assets/dake_icon.ico"),
    ]

