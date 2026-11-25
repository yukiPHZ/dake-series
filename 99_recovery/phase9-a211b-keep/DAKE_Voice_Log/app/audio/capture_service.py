# -*- coding: utf-8 -*-
from __future__ import annotations

import queue
import threading
import time
import logging
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable

import numpy as np

from app.audio.recorder_service import PcmSpoolWriter, RecorderConfig, RotationPolicy
from app.audio.vad_service import EnergyVad, VadConfig
from app.core.app_state import CaptureState
from app.storage.database import VoiceLogDatabase
from app.storage.paths import AppConfig, transcript_path_for_audio, unique_audio_path


@dataclass(frozen=True)
class CaptureEvent:
    kind: str
    state: CaptureState
    message: str = ""
    level: float = 0.0
    started_at: datetime | None = None


class CaptureService:
    def __init__(
        self,
        config: AppConfig,
        save_root: Path,
        spool_dir: Path,
        db: VoiceLogDatabase,
        events: "queue.Queue[CaptureEvent]",
        on_clip_recorded: Callable[[int], None] | None = None,
        sample_rate: int = 16000,
        logger: logging.Logger | None = None,
    ) -> None:
        self.config = config
        self.save_root = save_root
        self.spool_dir = spool_dir
        self.db = db
        self.events = events
        self.on_clip_recorded = on_clip_recorded
        self.sample_rate = sample_rate
        self.logger = logger or logging.getLogger("dake_voice_log")
        self.state = CaptureState.INITIALIZING
        self.level = 0.0
        self.mic_name = config.mic_label
        self._sd = None
        self._stream = None
        self._thread: threading.Thread | None = None
        self._stop_event = threading.Event()
        self._paused = threading.Event()
        self._audio_queue: queue.Queue[np.ndarray] = queue.Queue(maxsize=60)
        self._current_writer: PcmSpoolWriter | None = None
        self._current_started_at: datetime | None = None
        self._current_session_id = ""
        self._rotation = RotationPolicy()
        self._vad = EnergyVad(
            VadConfig(
                sample_rate=sample_rate,
                silence_timeout_sec=config.silence_timeout_sec,
                sensitivity=config.sensitivity,
            )
        )
        try:
            import sounddevice as sd  # type: ignore

            self._sd = sd
        except Exception:
            self._sd = None

    def start(self) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._stop_event.clear()
        self._paused.clear()
        self._thread = threading.Thread(target=self._run, name="dake-voice-capture", daemon=True)
        self._thread.start()

    def pause(self) -> None:
        self._paused.set()
        if self.state == CaptureState.RECORDING:
            self._finish_recording(datetime.now())
        self._set_state(CaptureState.PAUSED)

    def resume(self) -> None:
        self._paused.clear()
        if self._sd is None:
            self._set_state(CaptureState.MIC_ERROR, "sounddevice is not installed.")
        else:
            self._set_state(CaptureState.WAITING)

    def stop(self) -> None:
        self._stop_event.set()
        if self.state == CaptureState.RECORDING:
            self._finish_recording(datetime.now())
        self._set_state(CaptureState.STOPPED)

    def join(self, timeout: float = 2.0) -> None:
        if self._thread:
            self._thread.join(timeout=timeout)

    def _run(self) -> None:
        if self._sd is None:
            self.logger.error("audio.capture.sounddevice_missing")
            self._set_state(CaptureState.MIC_ERROR, "sounddevice is not installed.")
            return
        while not self._stop_event.is_set():
            if self._paused.is_set():
                time.sleep(0.2)
                continue
            try:
                self._open_stream()
                self._set_state(CaptureState.WAITING)
                self._process_audio_queue()
            except Exception as exc:
                self._close_stream()
                if self.state == CaptureState.RECORDING:
                    self._finish_recording(datetime.now())
                self.logger.error("audio.capture.stream_error error=%s", exc)
                self._set_state(CaptureState.MIC_ERROR, str(exc))
                time.sleep(3.0)
        self._close_stream()

    def _open_stream(self) -> None:
        if self._sd is None:
            raise RuntimeError("sounddevice is not installed.")
        blocksize = int(self.sample_rate * 0.1)
        device = int(self.config.mic_device) if self.config.mic_device.strip().isdigit() else None

        def callback(indata, frames, time_info, status) -> None:  # noqa: ANN001
            if status:
                pass
            try:
                self._audio_queue.put_nowait(np.array(indata, copy=True))
            except queue.Full:
                try:
                    self._audio_queue.get_nowait()
                    self._audio_queue.put_nowait(np.array(indata, copy=True))
                except Exception:
                    pass

        self._stream = self._sd.InputStream(
            samplerate=self.sample_rate,
            blocksize=blocksize,
            device=device,
            channels=1,
            dtype="int16",
            callback=callback,
        )
        self._stream.start()
        try:
            info = self._sd.query_devices(device, "input")
            self.mic_name = str(info.get("name") or self.mic_name)
        except Exception:
            pass
        self.logger.info("audio.capture.stream_opened mic=%s sample_rate=%s", self.mic_name, self.sample_rate)

    def _close_stream(self) -> None:
        stream = self._stream
        self._stream = None
        if stream is None:
            return
        try:
            stream.stop()
            stream.close()
        except Exception:
            pass

    def _process_audio_queue(self) -> None:
        while not self._stop_event.is_set() and not self._paused.is_set():
            try:
                samples = self._audio_queue.get(timeout=0.3)
            except queue.Empty:
                continue
            decision = self._vad.consume(samples)
            self.level = min(1.0, decision.rms * 20.0)
            self.events.put(CaptureEvent("level", self.state, level=self.level))
            now = datetime.now()
            if decision.start_recording and self.state != CaptureState.RECORDING:
                self._start_recording(now)
                for block in self._vad.pre_roll_blocks():
                    self._write_block(block)
            elif self.state == CaptureState.RECORDING:
                self._write_block(samples)
            if self.state == CaptureState.RECORDING and self._current_started_at is not None:
                if self._rotation.should_rotate(self._current_started_at, now) or self._rotation.should_roll_date(self._current_started_at, now):
                    self._finish_recording(now, reset_session=False)
                    self._start_recording(now)
                    for block in self._vad.pre_roll_blocks():
                        self._write_block(block)
            if decision.stop_recording and self.state == CaptureState.RECORDING:
                self._finish_recording(now)

    def _start_recording(self, started_at: datetime) -> None:
        if not self._current_session_id:
            self._current_session_id = started_at.strftime("%Y%m%d_%H%M%S")
        writer = PcmSpoolWriter(self.spool_dir, RecorderConfig(sample_rate=self.sample_rate))
        writer.begin(started_at, self._current_session_id)
        self._current_writer = writer
        self._current_started_at = started_at
        self.logger.info("audio.recording.started session_id=%s start=%s", self._current_session_id, started_at.isoformat(timespec="seconds"))
        self._set_state(CaptureState.RECORDING, started_at=started_at)

    def _write_block(self, samples: np.ndarray) -> None:
        if self._current_writer is None:
            return
        self._current_writer.write(samples)

    def _finish_recording(self, ended_at: datetime, reset_session: bool = True) -> None:
        writer = self._current_writer
        started_at = self._current_started_at
        self._current_writer = None
        self._current_started_at = None
        if writer is None or started_at is None:
            return
        try:
            wav_path = unique_audio_path(self.save_root, started_at)
            writer.finalize(wav_path)
            transcript_path = transcript_path_for_audio(self.save_root, wav_path, started_at.date())
            clip_id = self.db.create_clip(
                session_id=writer.session_id,
                started_at=started_at,
                ended_at=ended_at,
                audio_path=wav_path,
                transcript_path=transcript_path,
            )
            if self.on_clip_recorded:
                self.on_clip_recorded(clip_id)
            self.logger.info("audio.recording.finished clip_id=%s session_id=%s path=%s", clip_id, writer.session_id, wav_path)
        except Exception as exc:
            writer.abort_close()
            self.logger.error("audio.recording.save_failed error=%s", exc)
            self._set_state(CaptureState.MIC_ERROR, str(exc))
            return
        self._set_state(CaptureState.WAITING)
        if reset_session:
            self._current_session_id = ""

    def _set_state(self, state: CaptureState, message: str = "", started_at: datetime | None = None) -> None:
        self.state = state
        self.events.put(CaptureEvent("state", state, message=message, level=self.level, started_at=started_at))
