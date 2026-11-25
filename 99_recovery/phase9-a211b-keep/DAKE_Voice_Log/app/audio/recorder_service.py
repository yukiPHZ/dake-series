# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import uuid
import wave
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import numpy as np

from app.storage.recovery import finalize_pcm_to_wav


@dataclass(frozen=True)
class RecorderConfig:
    sample_rate: int = 16000
    channels: int = 1
    sample_width: int = 2
    max_clip_sec: int = 30 * 60


class PcmSpoolWriter:
    def __init__(self, spool_dir: Path, config: RecorderConfig) -> None:
        self.spool_dir = spool_dir
        self.config = config
        self.pcm_path: Path | None = None
        self.meta_path: Path | None = None
        self._handle = None
        self.started_at: datetime | None = None
        self.session_id = ""

    def begin(self, started_at: datetime, session_id: str) -> None:
        self.spool_dir.mkdir(parents=True, exist_ok=True)
        token = f"{started_at:%Y%m%d_%H%M%S}_{uuid.uuid4().hex[:8]}"
        self.pcm_path = self.spool_dir / f"{token}.pcm"
        self.meta_path = self.spool_dir / f"{token}.json"
        self.started_at = started_at
        self.session_id = session_id
        payload = {
            "pcm_path": str(self.pcm_path),
            "sample_rate": self.config.sample_rate,
            "channels": self.config.channels,
            "sample_width": self.config.sample_width,
            "start_time": started_at.isoformat(timespec="seconds"),
            "session_id": session_id,
        }
        self.meta_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        self._handle = self.pcm_path.open("ab")

    def write(self, samples: np.ndarray) -> None:
        if self._handle is None:
            return
        data = np.asarray(samples)
        if data.ndim > 1 and data.shape[1] > self.config.channels:
            data = data[:, : self.config.channels]
        if data.dtype != np.int16:
            data = np.clip(data, -1.0, 1.0)
            data = (data * 32767).astype(np.int16)
        self._handle.write(data.tobytes())
        self._handle.flush()

    def finalize(self, wav_path: Path) -> Path:
        if self._handle is not None:
            self._handle.close()
            self._handle = None
        if self.pcm_path is None:
            raise RuntimeError("Recorder was not started.")
        finalize_pcm_to_wav(
            self.pcm_path,
            wav_path,
            self.config.sample_rate,
            self.config.channels,
            self.config.sample_width,
        )
        self.pcm_path.unlink(missing_ok=True)
        if self.meta_path is not None:
            self.meta_path.unlink(missing_ok=True)
        return wav_path

    def abort_close(self) -> None:
        if self._handle is not None:
            self._handle.close()
            self._handle = None


class RotationPolicy:
    def __init__(self, max_clip_sec: int = 30 * 60) -> None:
        self.max_clip_sec = max_clip_sec

    def should_rotate(self, started_at: datetime, now: datetime) -> bool:
        return (now - started_at).total_seconds() >= self.max_clip_sec

    @staticmethod
    def should_roll_date(started_at: datetime, now: datetime) -> bool:
        return started_at.date() != now.date()

