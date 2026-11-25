# -*- coding: utf-8 -*-
from __future__ import annotations

import queue
import threading
from dataclasses import dataclass
from datetime import date
from pathlib import Path

from app.constants import SUMMARY_MODEL, UI_TEXT
from app.core.app_state import CaptureState, SummaryState
from app.storage.database import DAILY_COMPLETE, DAILY_FAILED, DAILY_PROCESSING, VoiceLogDatabase
from app.storage.paths import daily_summary_path, daily_transcript_path
from app.summary.ollama_client import OllamaClient
from app.summary.prompts import chunk_prompt, final_prompt


@dataclass(frozen=True)
class SummaryEvent:
    state: SummaryState
    message: str = ""
    record_date: str = ""


def split_daily_text(text: str, max_chars: int = 7000) -> list[str]:
    sections = text.split("\n## ")
    if not sections:
        return []
    header = sections[0]
    chunks: list[str] = []
    current = header.strip()
    for section in sections[1:]:
        part = "## " + section.strip()
        if len(current) + len(part) + 2 > max_chars and current.strip():
            chunks.append(current.strip())
            current = part
        else:
            current = f"{current}\n\n{part}" if current else part
    if current.strip():
        chunks.append(current.strip())
    return chunks


class SummaryService:
    def __init__(
        self,
        db: VoiceLogDatabase,
        save_root: Path,
        events: "queue.Queue[SummaryEvent]",
        client: OllamaClient | None = None,
        model_name: str = SUMMARY_MODEL,
    ) -> None:
        self.db = db
        self.save_root = save_root
        self.events = events
        self.client = client or OllamaClient()
        self.model_name = model_name
        self._thread: threading.Thread | None = None
        self._stop_event = threading.Event()
        self._wake = threading.Event()
        self.enabled = False
        self.capture_state = CaptureState.WAITING
        self.backlog_count = 0
        self.state = SummaryState.DISABLED

    def start(self, enabled: bool) -> None:
        self.enabled = enabled
        if self._thread and self._thread.is_alive():
            self.wake()
            return
        self._stop_event.clear()
        self._thread = threading.Thread(target=self._run, name="dake-voice-summary", daemon=True)
        self._thread.start()

    def stop(self) -> None:
        self._stop_event.set()
        self._wake.set()

    def set_enabled(self, enabled: bool) -> None:
        self.enabled = enabled
        self.wake()

    def update_activity(self, capture_state: CaptureState, backlog_count: int) -> None:
        self.capture_state = capture_state
        self.backlog_count = backlog_count
        if capture_state == CaptureState.WAITING and backlog_count == 0:
            self.wake()

    def wake(self) -> None:
        self._wake.set()

    def _run(self) -> None:
        while not self._stop_event.is_set():
            if not self.enabled:
                self._set_state(SummaryState.DISABLED)
                self._wake.wait(timeout=10.0)
                self._wake.clear()
                continue
            if self.capture_state != CaptureState.WAITING or self.backlog_count:
                self._set_state(SummaryState.WAITING)
                self._wake.wait(timeout=30.0)
                self._wake.clear()
                continue
            pending = self.db.summary_pending_dates(date.today(), self.model_name)
            if not pending:
                self._set_state(SummaryState.WAITING)
                self._wake.wait(timeout=60.0)
                self._wake.clear()
                continue
            status = self.client.status()
            if not status.reachable:
                self._set_state(SummaryState.ERROR, UI_TEXT["status_ollama_missing"])
                self._wake.wait(timeout=120.0)
                self._wake.clear()
                continue
            if self.model_name not in status.models:
                self._set_state(SummaryState.ERROR, UI_TEXT["status_ollama_model_missing"])
                self._wake.wait(timeout=120.0)
                self._wake.clear()
                continue
            try:
                self._summarize_date(pending[0])
            except Exception:
                self._wake.wait(timeout=120.0)
                self._wake.clear()

    def _summarize_date(self, record_date: str) -> Path:
        self._set_state(SummaryState.PROCESSING, UI_TEXT["processing_summarizing"], record_date)
        output_path = daily_summary_path(self.save_root, record_date)
        self.db.update_summary(record_date, output_path, DAILY_PROCESSING, self.model_name)
        try:
            source_path = daily_transcript_path(self.save_root, record_date)
            text = source_path.read_text(encoding="utf-8")
            chunks = split_daily_text(text)
            chunk_summaries = [self.client.generate(self.model_name, chunk_prompt(record_date, chunk)) for chunk in chunks]
            final = self.client.generate(self.model_name, final_prompt(record_date, chunk_summaries))
            output_path.parent.mkdir(parents=True, exist_ok=True)
            output_path.write_text(final.rstrip() + "\n", encoding="utf-8")
            self.db.update_summary(record_date, output_path, DAILY_COMPLETE, self.model_name)
            self._set_state(SummaryState.COMPLETE, record_date=record_date)
            return output_path
        except Exception as exc:
            self.db.update_summary(record_date, output_path, DAILY_FAILED, self.model_name, str(exc))
            self._set_state(SummaryState.ERROR, str(exc), record_date)
            raise

    def summarize_now(self, record_date: str) -> None:
        if not self.enabled:
            self.enabled = True
        threading.Thread(target=lambda: self._summarize_date(record_date), daemon=True).start()

    def _set_state(self, state: SummaryState, message: str = "", record_date: str = "") -> None:
        self.state = state
        self.events.put(SummaryEvent(state, message, record_date))
