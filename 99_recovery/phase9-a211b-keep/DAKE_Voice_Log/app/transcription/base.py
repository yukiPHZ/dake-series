# -*- coding: utf-8 -*-
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Protocol


@dataclass(frozen=True)
class TranscriptionSegment:
    start: float
    end: float
    text: str


@dataclass(frozen=True)
class TranscriptionResult:
    text: str
    segments: tuple[TranscriptionSegment, ...]
    no_speech: bool = False


class TranscriptionEngine(Protocol):
    def transcribe(self, audio_path: Path) -> TranscriptionResult:
        ...

