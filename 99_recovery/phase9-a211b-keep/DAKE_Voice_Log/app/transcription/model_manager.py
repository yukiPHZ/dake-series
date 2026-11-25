# -*- coding: utf-8 -*-
from __future__ import annotations

from pathlib import Path

from app.constants import TRANSCRIPTION_MODELS, UI_TEXT


class TranscriptionModelManager:
    def __init__(self, model_root: Path) -> None:
        self.model_root = model_root

    def model_name(self, quality: str) -> str:
        return TRANSCRIPTION_MODELS.get(quality, TRANSCRIPTION_MODELS["standard"])

    def quality_label(self, quality: str) -> str:
        return UI_TEXT["quality_light"] if quality == "light" else UI_TEXT["quality_standard"]

    def model_hint(self, quality: str) -> str:
        name = self.model_name(quality)
        return f"faster-whisper-{name}"

    def is_model_ready(self, quality: str) -> bool:
        name = self.model_name(quality)
        if not self.model_root.exists():
            return False
        expected_tokens = (name, f"faster-whisper-{name}")
        for path in self.model_root.rglob("*"):
            lowered = path.name.lower()
            if any(token in lowered for token in expected_tokens):
                return True
        return False

