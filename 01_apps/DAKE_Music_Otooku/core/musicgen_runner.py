# -*- coding: utf-8 -*-
from __future__ import annotations

import importlib.util
from dataclasses import dataclass
from pathlib import Path

from .app_config import AUDIOCRAFT_MODEL_NAME, DEFAULT_DURATION_SECONDS


@dataclass(frozen=True)
class MusicGenResult:
    success: bool
    output_path: Path | None = None
    message: str = ""


def is_musicgen_available() -> bool:
    return importlib.util.find_spec("audiocraft") is not None


def generate_music(
    prompt: str,
    output_path: Path,
    duration_seconds: int = DEFAULT_DURATION_SECONDS,
    model_name: str = AUDIOCRAFT_MODEL_NAME,
) -> MusicGenResult:
    if not is_musicgen_available():
        return MusicGenResult(False, message="MusicGen unavailable")

    try:
        from audiocraft.data.audio import audio_write
        from audiocraft.models import MusicGen

        output_path.parent.mkdir(parents=True, exist_ok=True)
        model = MusicGen.get_pretrained(model_name)
        model.set_generation_params(duration=int(duration_seconds))
        wav = model.generate([prompt], progress=False)[0].cpu()
        audio_write(
            str(output_path.with_suffix("")),
            wav,
            model.sample_rate,
            strategy="loudness",
            loudness_compressor=True,
        )
        if output_path.exists():
            return MusicGenResult(True, output_path=output_path, message="MusicGen complete")
        return MusicGenResult(False, message="MusicGen finished but output file was not found")
    except Exception as exc:
        return MusicGenResult(False, message=str(exc))

