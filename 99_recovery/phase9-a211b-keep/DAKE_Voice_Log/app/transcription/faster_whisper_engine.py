# -*- coding: utf-8 -*-
from __future__ import annotations

import os
from pathlib import Path

from app.transcription.base import TranscriptionResult, TranscriptionSegment


class FasterWhisperEngine:
    def __init__(
        self,
        model_name: str,
        model_root: Path,
        language: str = "ja",
        cpu_threads: int | None = None,
    ) -> None:
        self.model_name = model_name
        self.model_root = model_root
        self.language = language
        self.cpu_threads = cpu_threads or max(1, min(4, (os.cpu_count() or 2) // 2))
        self._model = None

    def _load_model(self):  # noqa: ANN202
        if self._model is not None:
            return self._model
        try:
            from faster_whisper import WhisperModel  # type: ignore
        except Exception as exc:
            raise RuntimeError("faster-whisper is not installed.") from exc
        self.model_root.mkdir(parents=True, exist_ok=True)
        self._model = WhisperModel(
            self.model_name,
            device="cpu",
            compute_type="int8",
            download_root=str(self.model_root),
            cpu_threads=self.cpu_threads,
            num_workers=1,
        )
        return self._model

    def transcribe(self, audio_path: Path) -> TranscriptionResult:
        model = self._load_model()
        segments_iter, _info = model.transcribe(
            str(audio_path),
            language=self.language,
            beam_size=1,
            vad_filter=False,
            condition_on_previous_text=False,
        )
        segments: list[TranscriptionSegment] = []
        texts: list[str] = []
        for segment in segments_iter:
            text = str(segment.text).strip()
            if not text:
                continue
            segments.append(TranscriptionSegment(float(segment.start), float(segment.end), text))
            texts.append(text)
        text = "\n".join(texts).strip()
        return TranscriptionResult(text=text, segments=tuple(segments), no_speech=not bool(text))

