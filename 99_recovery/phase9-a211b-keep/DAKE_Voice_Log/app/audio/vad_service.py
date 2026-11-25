# -*- coding: utf-8 -*-
from __future__ import annotations

from collections import deque
from dataclasses import dataclass

import numpy as np


SENSITIVITY_MULTIPLIER = {
    "low": 5.5,
    "standard": 3.2,
    "high": 2.2,
}


@dataclass(frozen=True)
class VadConfig:
    sample_rate: int = 16000
    frame_ms: int = 100
    pre_roll_sec: float = 1.5
    min_speech_sec: float = 0.25
    silence_timeout_sec: float = 30.0
    sensitivity: str = "standard"
    min_noise_floor: float = 0.0015
    threshold_margin: float = 0.0025


@dataclass(frozen=True)
class VadDecision:
    voice: bool
    start_recording: bool
    stop_recording: bool
    rms: float
    noise_floor: float


class EnergyVad:
    def __init__(self, config: VadConfig) -> None:
        self.config = config
        self.noise_floor = config.min_noise_floor
        self._recording = False
        self._speech_candidate_sec = 0.0
        self._silence_sec = 0.0
        self._calibration_frames = max(1, int(2.0 / self.frame_seconds))
        self._calibration_values: list[float] = []
        self._pre_roll: deque[np.ndarray] = deque(maxlen=max(1, int(config.pre_roll_sec / self.frame_seconds) + 2))

    @property
    def frame_seconds(self) -> float:
        return self.config.frame_ms / 1000.0

    def reset(self) -> None:
        self._recording = False
        self._speech_candidate_sec = 0.0
        self._silence_sec = 0.0
        self._pre_roll.clear()

    def consume(self, samples: np.ndarray) -> VadDecision:
        mono = self._to_mono_float(samples)
        rms = self._rms(mono)
        self._pre_roll.append(np.array(samples, copy=True))
        self._update_noise_floor(rms)
        voice = rms >= self._start_threshold()
        stop_voice = rms >= self._stop_threshold()

        start = False
        stop = False
        if not self._recording:
            if voice:
                self._speech_candidate_sec += self.frame_seconds
            else:
                self._speech_candidate_sec = max(0.0, self._speech_candidate_sec - self.frame_seconds)
            if self._speech_candidate_sec >= self.config.min_speech_sec:
                self._recording = True
                start = True
                self._silence_sec = 0.0
        else:
            if stop_voice:
                self._silence_sec = 0.0
            else:
                self._silence_sec += self.frame_seconds
            if self._silence_sec >= self.config.silence_timeout_sec:
                self._recording = False
                stop = True
                self._speech_candidate_sec = 0.0
                self._silence_sec = 0.0

        return VadDecision(
            voice=voice,
            start_recording=start,
            stop_recording=stop,
            rms=rms,
            noise_floor=self.noise_floor,
        )

    def pre_roll_blocks(self) -> list[np.ndarray]:
        return [np.array(block, copy=True) for block in self._pre_roll]

    def _update_noise_floor(self, rms: float) -> None:
        if len(self._calibration_values) < self._calibration_frames:
            self._calibration_values.append(rms)
            self.noise_floor = max(self.config.min_noise_floor, float(np.percentile(self._calibration_values, 70)))
            return
        if self._recording:
            return
        if rms < self._start_threshold() * 0.9:
            self.noise_floor = max(
                self.config.min_noise_floor,
                self.noise_floor * 0.985 + rms * 0.015,
            )

    def _start_threshold(self) -> float:
        multiplier = SENSITIVITY_MULTIPLIER.get(self.config.sensitivity, SENSITIVITY_MULTIPLIER["standard"])
        return self.noise_floor * multiplier + self.config.threshold_margin

    def _stop_threshold(self) -> float:
        return self._start_threshold() * 0.55

    @staticmethod
    def _to_mono_float(samples: np.ndarray) -> np.ndarray:
        array = np.asarray(samples)
        if array.ndim > 1:
            array = array.mean(axis=1)
        if np.issubdtype(array.dtype, np.integer):
            return array.astype(np.float32) / np.iinfo(array.dtype).max
        return array.astype(np.float32)

    @staticmethod
    def _rms(samples: np.ndarray) -> float:
        if samples.size == 0:
            return 0.0
        return float(np.sqrt(np.mean(np.square(samples, dtype=np.float32))))

