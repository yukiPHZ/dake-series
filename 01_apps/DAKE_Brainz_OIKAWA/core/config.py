# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import os
import sys
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any


APP_FOLDER_NAME = "DAKE_Brainz_OIKAWA"
CONFIG_FILE_NAME = "oikawa_config.json"
DEFAULT_MEMORY_FOLDER = Path.home() / "Documents" / "PEAKHEADZ_ROOT"
LEGACY_MEMORY_FOLDER = Path.home() / "Documents" / "brainz_memory"


@dataclass
class AppConfig:
    memory_folder: str = ""


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        if exe_dir.name.lower() == "dist":
            return exe_dir.parent
        return exe_dir
    return Path(__file__).resolve().parent.parent


def series_root() -> Path:
    return app_dir().parent.parent


def config_path() -> Path:
    return app_dir() / CONFIG_FILE_NAME


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


def normalize_folder(value: str | Path) -> Path:
    text = os.path.expandvars(str(value)).strip()
    return Path(text).expanduser()


def existing_folder(value: str | Path) -> Path | None:
    if not value:
        return None
    path = normalize_folder(value)
    try:
        resolved = path.resolve()
    except OSError:
        resolved = path
    if resolved.exists() and resolved.is_dir():
        return resolved
    return None


class ConfigStore:
    def __init__(self, path: Path | None = None) -> None:
        self.path = path or config_path()

    def load(self) -> AppConfig:
        data = read_json_file(self.path)
        return AppConfig(memory_folder=str(data.get("memory_folder", "") or ""))

    def save(self, config: AppConfig) -> None:
        write_json_file(self.path, asdict(config))


def brainz_config_candidates() -> list[Path]:
    apps_dir = app_dir().parent
    return [
        apps_dir / "DAKE_Brainz_Search" / "data" / "config" / "brainz_config.json",
        series_root() / "01_apps" / "DAKE_Brainz_Search" / "data" / "config" / "brainz_config.json",
    ]


def find_brainz_config_memory() -> Path | None:
    for path in brainz_config_candidates():
        data = read_json_file(path)
        folder = existing_folder(str(data.get("memory_folder", "") or ""))
        if folder:
            return folder
    return None


def resolve_memory_folder(config: AppConfig) -> Path | None:
    configured = existing_folder(config.memory_folder)
    if configured:
        return configured

    brainz_configured = find_brainz_config_memory()
    if brainz_configured:
        return brainz_configured

    default_folder = existing_folder(DEFAULT_MEMORY_FOLDER)
    if default_folder:
        return default_folder

    legacy_folder = existing_folder(LEGACY_MEMORY_FOLDER)
    if legacy_folder:
        return legacy_folder

    return None


def open_path(path: Path) -> None:
    if sys.platform.startswith("win") and hasattr(os, "startfile"):
        os.startfile(str(path))  # type: ignore[attr-defined]
        return

    import webbrowser

    webbrowser.open(path.resolve().as_uri())
