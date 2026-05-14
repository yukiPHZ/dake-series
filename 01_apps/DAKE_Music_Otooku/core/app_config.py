# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import os
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
OLLAMA_DEFAULT_MODEL = "llama3.1"
OLLAMA_SETTINGS_PATH = PRESETS_DIR / "ollama_settings.json"

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


def ensure_data_dirs() -> None:
    for path in (
        PROMPTS_DIR,
        REFERENCES_DIR,
        OUTPUTS_DIR,
        LOGS_DIR,
        PRESETS_DIR,
    ):
        path.mkdir(parents=True, exist_ok=True)

