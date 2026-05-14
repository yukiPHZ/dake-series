# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import os
import shutil
from pathlib import Path


APP_KEY = "DAKE_Music_Otooku"
DISPLAY_NAME = "音を置く"
EXE_NAME = "DakeMusic_Otooku.exe"

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
PROMPTS_DIR = DATA_DIR / "prompts"
REFERENCES_DIR = DATA_DIR / "references"
OUTPUTS_DIR = DATA_DIR / "outputs"
LOGS_DIR = DATA_DIR / "logs"
PRESETS_DIR = DATA_DIR / "presets"

OLLAMA_BASE_URL = "http://localhost:11434"
OLLAMA_DEFAULT_MODEL = "qwen2.5:7b"
OLLAMA_SETTINGS_PATH = PRESETS_DIR / "ollama_settings.json"
FFMPEG_SETTINGS_PATH = PRESETS_DIR / "ffmpeg_settings.json"

DEFAULT_DURATION_SECONDS = 12
MIN_DURATION_SECONDS = 10
MAX_DURATION_SECONDS = 20

AUDIOCRAFT_MODEL_NAME = "facebook/musicgen-small"


def load_ollama_model_name() -> str:
    env_model = os.environ.get("DAKE_OTOOKU_OLLAMA_MODEL", "").strip()
    if env_model:
        return env_model

    try:
        data = json.loads(OLLAMA_SETTINGS_PATH.read_text(encoding="utf-8"))
        model = str(data.get("ollama_model", "")).strip()
        if model:
            return model
    except Exception:
        pass

    return OLLAMA_DEFAULT_MODEL


def _configured_tool_path(tool_name: str) -> str:
    env_name = f"DAKE_OTOOKU_{tool_name.upper()}"
    env_path = os.environ.get(env_name, "").strip()
    if env_path:
        return env_path

    try:
        data = json.loads(FFMPEG_SETTINGS_PATH.read_text(encoding="utf-8"))
        value = str(data.get(tool_name, "")).strip()
        if value:
            return value
    except Exception:
        pass

    return ""


def _winget_tool_candidates(tool_name: str) -> tuple[Path, ...]:
    local_app_data = Path(os.environ.get("LOCALAPPDATA", ""))
    if not local_app_data:
        return ()
    packages = local_app_data / "Microsoft" / "WinGet" / "Packages"
    return (
        packages / "Gyan.FFmpeg_Microsoft.Winget.Source_8wekyb3d8bbwe" / "ffmpeg-8.1.1-full_build" / "bin" / f"{tool_name}.exe",
        packages / "yt-dlp.FFmpeg_Microsoft.Winget.Source_8wekyb3d8bbwe" / "ffmpeg-N-124279-g0f6ba39122-win64-gpl" / "bin" / f"{tool_name}.exe",
    )


def resolve_tool_command(tool_name: str) -> str | None:
    configured = _configured_tool_path(tool_name)
    if configured and Path(configured).exists():
        return configured

    found = shutil.which(tool_name)
    if found:
        return found

    for candidate in _winget_tool_candidates(tool_name):
        if candidate.exists():
            return str(candidate)

    return None


def ensure_data_dirs() -> None:
    for path in (
        PROMPTS_DIR,
        REFERENCES_DIR,
        OUTPUTS_DIR,
        LOGS_DIR,
        PRESETS_DIR,
    ):
        path.mkdir(parents=True, exist_ok=True)
