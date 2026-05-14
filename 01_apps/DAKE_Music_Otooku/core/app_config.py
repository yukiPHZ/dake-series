# -*- coding: utf-8 -*-
from __future__ import annotations

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
OLLAMA_MODEL_CANDIDATES = (
    "llama3.1:8b",
    "llama3.1",
    "llama3",
    "gemma2",
    "mistral",
)

DEFAULT_DURATION_SECONDS = 12
MIN_DURATION_SECONDS = 10
MAX_DURATION_SECONDS = 20

AUDIOCRAFT_MODEL_NAME = "facebook/musicgen-small"


def ensure_data_dirs() -> None:
    for path in (
        PROMPTS_DIR,
        REFERENCES_DIR,
        OUTPUTS_DIR,
        LOGS_DIR,
        PRESETS_DIR,
    ):
        path.mkdir(parents=True, exist_ok=True)

