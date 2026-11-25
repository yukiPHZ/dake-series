# -*- coding: utf-8 -*-
from __future__ import annotations

from dataclasses import dataclass
from enum import Enum


class CaptureState(str, Enum):
    INITIALIZING = "INITIALIZING"
    WAITING = "WAITING"
    RECORDING = "RECORDING"
    PAUSED = "PAUSED"
    MIC_ERROR = "MIC_ERROR"
    STOPPED = "STOPPED"


class TranscriptionState(str, Enum):
    IDLE = "IDLE"
    QUEUED = "QUEUED"
    PROCESSING = "PROCESSING"
    ERROR = "ERROR"


class SummaryState(str, Enum):
    DISABLED = "DISABLED"
    WAITING = "WAITING"
    PROCESSING = "PROCESSING"
    COMPLETE = "COMPLETE"
    ERROR = "ERROR"


@dataclass(frozen=True)
class TodayStats:
    recorded: int = 0
    transcribed: int = 0
    pending: int = 0
    summary_status: str = "WAITING"


@dataclass(frozen=True)
class AppSnapshot:
    capture_state: CaptureState
    transcription_state: TranscriptionState
    summary_state: SummaryState
    mic_name: str
    level: float
    today: TodayStats
    backlog_count: int
    active_recording_started_at: str | None = None
    last_message: str = ""

