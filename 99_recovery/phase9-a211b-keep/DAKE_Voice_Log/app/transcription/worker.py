# -*- coding: utf-8 -*-
from __future__ import annotations

import queue
import threading
import time
from dataclasses import dataclass
from pathlib import Path

from app.constants import UI_TEXT
from app.core.app_state import TranscriptionState
from app.core.job_queue import RetryPolicy
from app.storage.daily_builder import build_daily_transcript, write_clip_transcript
from app.storage.database import (
    TRANSCRIPTION_DONE,
    TRANSCRIPTION_FAILED,
    TRANSCRIPTION_NO_SPEECH,
    TRANSCRIPTION_PROCESSING,
    TRANSCRIPTION_QUEUED,
    ClipRecord,
    VoiceLogDatabase,
)
from app.transcription.faster_whisper_engine import FasterWhisperEngine
from app.transcription.model_manager import TranscriptionModelManager


@dataclass(frozen=True)
class TranscriptionEvent:
    state: TranscriptionState
    message: str = ""
    clip_id: int | None = None


class TranscriptionService:
    def __init__(
        self,
        db: VoiceLogDatabase,
        save_root: Path,
        model_root: Path,
        quality: str,
        events: "queue.Queue[TranscriptionEvent]",
    ) -> None:
        self.db = db
        self.save_root = save_root
        self.model_manager = TranscriptionModelManager(model_root)
        self.quality = quality
        self.events = events
        self.retry_policy = RetryPolicy()
        self._thread: threading.Thread | None = None
        self._stop_event = threading.Event()
        self._wake = threading.Event()
        self._engine: FasterWhisperEngine | None = None
        self.state = TranscriptionState.IDLE

    def start(self) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._stop_event.clear()
        self.db.reset_processing_jobs()
        self.db.queue_recorded_clips()
        self._thread = threading.Thread(target=self._run, name="dake-voice-transcription", daemon=True)
        self._thread.start()

    def stop(self) -> None:
        self._stop_event.set()
        self._wake.set()

    def wake(self) -> None:
        self.db.queue_recorded_clips()
        self._wake.set()

    def set_quality(self, quality: str) -> None:
        if quality != self.quality:
            self.quality = quality
            self._engine = None
        self.wake()

    def _run(self) -> None:
        while not self._stop_event.is_set():
            clip = self.db.next_transcription_clip()
            if clip is None:
                self._set_state(TranscriptionState.IDLE)
                self._wake.wait(timeout=5.0)
                self._wake.clear()
                continue
            self._process_clip(clip)

    def _process_clip(self, clip: ClipRecord) -> None:
        self.db.update_clip_status(clip.id, TRANSCRIPTION_PROCESSING)
        self._set_state(TranscriptionState.PROCESSING, UI_TEXT["processing_transcribing"], clip.id)
        try:
            engine = self._get_engine()
            result = engine.transcribe(Path(clip.audio_path))
            transcript_path = Path(clip.transcript_path)
            if result.no_speech:
                self.db.update_clip_status(clip.id, TRANSCRIPTION_NO_SPEECH, transcript_path=transcript_path)
            else:
                write_clip_transcript(transcript_path, clip, result.text)
                self.db.update_clip_status(clip.id, TRANSCRIPTION_DONE, transcript_path=transcript_path)
                build_daily_transcript(self.db, self.save_root, clip.record_date)
        except Exception as exc:
            retries = clip.transcription_retries + 1
            next_retry = self.retry_policy.next_retry_at(retries).isoformat(timespec="seconds")
            status = TRANSCRIPTION_QUEUED if self.retry_policy.can_retry(retries) else TRANSCRIPTION_FAILED
            self.db.update_clip_status(
                clip.id,
                status,
                error=str(exc),
                retries=retries,
                next_retry_at=next_retry if status == TRANSCRIPTION_QUEUED else "",
            )
            self._set_state(TranscriptionState.ERROR, str(exc), clip.id)
            time.sleep(1.0)

    def _get_engine(self) -> FasterWhisperEngine:
        if self._engine is None:
            self._engine = FasterWhisperEngine(
                self.model_manager.model_name(self.quality),
                self.model_manager.model_root,
            )
        return self._engine

    def _set_state(self, state: TranscriptionState, message: str = "", clip_id: int | None = None) -> None:
        self.state = state
        self.events.put(TranscriptionEvent(state, message, clip_id))

