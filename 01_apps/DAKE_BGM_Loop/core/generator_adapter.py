# -*- coding: utf-8 -*-
from __future__ import annotations

import math
import os
import random
import shlex
import shutil
import struct
import subprocess
import wave
from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class GenerateRequest:
    prompt: str
    mood: str
    mood_slug: str
    duration_sec: int
    seed: int
    output_path: Path


@dataclass(frozen=True)
class GenerateResult:
    output_path: Path
    adapter_name: str
    message: str = ""
    mock_profile: dict[str, Any] | None = None


class AdapterUnavailable(RuntimeError):
    """Raised when a generator adapter is selected but not configured."""


class BaseGeneratorAdapter:
    name = "base"

    def is_available(self) -> bool:
        return True

    def status_message(self) -> str:
        return ""

    def generate(self, request: GenerateRequest) -> GenerateResult:
        raise NotImplementedError


class MockGeneratorAdapter(BaseGeneratorAdapter):
    """Small deterministic WAV generator used for Phase 3 UI and file-flow checks."""

    name = "mock"
    sample_rate = 22050

    _profiles = {
        "のんき": {
            "label": "soft mallet and bell",
            "base": (329.63, 392.0, 440.0, 523.25),
            "pad_amp": 0.055,
            "low_amp": 0.025,
            "bell_amp": 0.30,
            "bell_events_per_15": 8,
            "click_amp": 0.006,
            "clicks_per_15": 4,
            "noise_amp": 0.002,
            "echo": 0.08,
            "space": 0.0,
        },
        "静か": {
            "label": "low soft pad",
            "base": (98.0, 110.0, 130.81, 146.83),
            "pad_amp": 0.26,
            "low_amp": 0.13,
            "bell_amp": 0.015,
            "bell_events_per_15": 1,
            "click_amp": 0.0,
            "clicks_per_15": 0,
            "noise_amp": 0.002,
            "echo": 0.04,
            "space": 0.0,
        },
        "作業用": {
            "label": "steady work pad",
            "base": (130.81, 164.81, 196.0, 246.94),
            "pad_amp": 0.18,
            "low_amp": 0.08,
            "bell_amp": 0.035,
            "bell_events_per_15": 3,
            "click_amp": 0.018,
            "clicks_per_15": 8,
            "noise_amp": 0.004,
            "echo": 0.05,
            "space": 0.0,
        },
        "神社": {
            "label": "distant bell with airy echo",
            "base": (196.0, 246.94, 329.63, 392.0),
            "pad_amp": 0.09,
            "low_amp": 0.035,
            "bell_amp": 0.20,
            "bell_events_per_15": 3,
            "click_amp": 0.004,
            "clicks_per_15": 2,
            "noise_amp": 0.003,
            "echo": 0.32,
            "space": 0.0,
        },
        "雨": {
            "label": "tiny grain noise and warm pad",
            "base": (110.0, 146.83, 174.61, 220.0),
            "pad_amp": 0.13,
            "low_amp": 0.06,
            "bell_amp": 0.0,
            "bell_events_per_15": 0,
            "click_amp": 0.026,
            "clicks_per_15": 32,
            "noise_amp": 0.060,
            "echo": 0.10,
            "space": 0.0,
        },
        "夜": {
            "label": "dark pad with light bass",
            "base": (123.47, 146.83, 164.81, 196.0),
            "pad_amp": 0.20,
            "low_amp": 0.035,
            "bell_amp": 0.015,
            "bell_events_per_15": 1,
            "click_amp": 0.0,
            "clicks_per_15": 0,
            "noise_amp": 0.006,
            "echo": 0.09,
            "space": 0.0,
        },
        "ミシン": {
            "label": "small sewing click rhythm",
            "base": (130.81, 164.81, 174.61, 196.0),
            "pad_amp": 0.06,
            "low_amp": 0.035,
            "bell_amp": 0.0,
            "bell_events_per_15": 0,
            "click_amp": 0.080,
            "clicks_per_15": 28,
            "noise_amp": 0.010,
            "echo": 0.06,
            "space": 0.0,
        },
        "コード": {
            "label": "keyboard click and soft pad",
            "base": (130.81, 146.83, 196.0, 246.94),
            "pad_amp": 0.13,
            "low_amp": 0.045,
            "bell_amp": 0.015,
            "bell_events_per_15": 2,
            "click_amp": 0.070,
            "clicks_per_15": 18,
            "noise_amp": 0.008,
            "echo": 0.05,
            "space": 0.0,
        },
        "余白": {
            "label": "sparse spacious texture",
            "base": (98.0, 146.83, 196.0),
            "pad_amp": 0.050,
            "low_amp": 0.018,
            "bell_amp": 0.040,
            "bell_events_per_15": 1,
            "click_amp": 0.0,
            "clicks_per_15": 0,
            "noise_amp": 0.001,
            "echo": 0.12,
            "space": 0.62,
        },
    }

    def generate(self, request: GenerateRequest) -> GenerateResult:
        request.output_path.parent.mkdir(parents=True, exist_ok=True)
        self._write_mock_wav(request)
        return GenerateResult(
            output_path=request.output_path,
            adapter_name=self.name,
            message="Mock loop generated.",
            mock_profile=self.describe_profile(request.mood),
        )

    def describe_profile(self, mood: str) -> dict[str, Any]:
        profile = self._profile_for_mood(mood)
        return {
            "label": profile["label"],
            "base": list(profile["base"]),
            "pad_amp": profile["pad_amp"],
            "low_amp": profile["low_amp"],
            "bell_amp": profile["bell_amp"],
            "bell_events_per_15": profile["bell_events_per_15"],
            "click_amp": profile["click_amp"],
            "clicks_per_15": profile["clicks_per_15"],
            "noise_amp": profile["noise_amp"],
            "echo": profile["echo"],
            "space": profile["space"],
        }

    def _profile_for_mood(self, mood: str) -> dict[str, Any]:
        return self._profiles.get(mood, self._profiles["のんき"])

    def _write_mock_wav(self, request: GenerateRequest) -> None:
        profile = self._profile_for_mood(request.mood)
        duration_sec = max(1, int(request.duration_sec))
        total_samples = self.sample_rate * duration_sec
        rng = random.Random(request.seed + _stable_mood_number(request.mood))

        base_freq = rng.choice(profile["base"])
        low_freq = max(41.0, base_freq / rng.choice((2.0, 2.5, 3.0, 4.0)))
        air_freq = base_freq * rng.choice((1.5, 2.0, 2.5, 3.0))
        phase_a = rng.random() * math.tau
        phase_b = rng.random() * math.tau
        phase_c = rng.random() * math.tau
        slow_cycles = rng.choice((1, 2, 3))
        slow_cycles_2 = rng.choice((2, 3, 4))

        samples: list[float] = [0.0] * total_samples
        for index in range(total_samples):
            loop_phase = index / total_samples
            motion = 0.70 + 0.20 * math.sin(math.tau * slow_cycles * loop_phase + phase_a)
            motion += 0.10 * math.sin(math.tau * slow_cycles_2 * loop_phase + phase_b)
            if profile["space"]:
                motion *= self._space_gate(loop_phase, float(profile["space"]), phase_c)

            pad = self._cyclic_sine(base_freq, duration_sec, loop_phase, phase_a)
            pad += 0.42 * self._cyclic_sine(air_freq, duration_sec, loop_phase, phase_b)
            low = self._cyclic_sine(low_freq, duration_sec, loop_phase, phase_b)
            texture = self._cyclic_sine(rng.choice((31.0, 43.0, 53.0)), duration_sec, loop_phase, phase_c)
            texture *= self._cyclic_sine(rng.choice((67.0, 79.0, 97.0)), duration_sec, loop_phase, phase_b)
            samples[index] += float(profile["pad_amp"]) * motion * pad
            samples[index] += float(profile["low_amp"]) * motion * low
            samples[index] += float(profile["noise_amp"]) * 0.18 * texture

        self._add_bell_events(samples, duration_sec, rng, profile)
        self._add_click_events(samples, duration_sec, rng, profile, request.mood)
        self._add_grain_noise(samples, rng, float(profile["noise_amp"]), request.mood)
        self._apply_circular_echo(samples, int(self.sample_rate * rng.choice((0.18, 0.24, 0.31))), float(profile["echo"]))
        self._apply_loop_crossfade(samples)
        self._soften_loop_edges(samples)

        peak = max(0.001, max(abs(sample) for sample in samples))
        scale = min(0.82 / peak, 1.0) * 0.72
        frames = bytearray()
        for sample in samples:
            sample = max(-0.92, min(0.92, sample * scale))
            frames.extend(struct.pack("<h", int(sample * 32767)))

        with wave.open(str(request.output_path), "wb") as wav_file:
            wav_file.setnchannels(1)
            wav_file.setsampwidth(2)
            wav_file.setframerate(self.sample_rate)
            wav_file.writeframes(frames)

    def _cyclic_sine(self, freq: float, duration_sec: int, loop_phase: float, phase: float) -> float:
        cycles = max(1, round(freq * duration_sec))
        return math.sin(math.tau * cycles * loop_phase + phase)

    def _space_gate(self, loop_phase: float, amount: float, phase: float) -> float:
        cycles = 3
        value = 0.5 + 0.5 * math.sin(math.tau * cycles * loop_phase + phase)
        threshold = min(0.88, 0.40 + amount * 0.45)
        if value < threshold:
            return 0.08
        return 0.08 + 0.92 * ((value - threshold) / max(0.001, 1.0 - threshold))

    def _add_bell_events(
        self,
        samples: list[float],
        duration_sec: int,
        rng: random.Random,
        profile: dict[str, Any],
    ) -> None:
        event_count = int(round(float(profile["bell_events_per_15"]) * duration_sec / 15))
        if event_count <= 0 or float(profile["bell_amp"]) <= 0:
            return
        total_samples = len(samples)
        segment = total_samples / event_count
        for event_index in range(event_count):
            center = int((event_index + 0.22 + rng.random() * 0.56) * segment) % total_samples
            freq = rng.choice(profile["base"]) * rng.choice((1.0, 1.5, 2.0))
            tail_sec = rng.choice((0.65, 0.85, 1.10))
            amp = float(profile["bell_amp"]) * rng.uniform(0.62, 1.0)
            self._add_tone_event(samples, center, freq, tail_sec, amp, decay=4.8, brightness=0.55)

    def _add_click_events(
        self,
        samples: list[float],
        duration_sec: int,
        rng: random.Random,
        profile: dict[str, Any],
        mood: str,
    ) -> None:
        event_count = int(round(float(profile["clicks_per_15"]) * duration_sec / 15))
        if event_count <= 0 or float(profile["click_amp"]) <= 0:
            return
        total_samples = len(samples)
        period = total_samples / event_count
        for event_index in range(event_count):
            jitter = rng.uniform(-0.12, 0.12)
            if mood in {"コード", "雨"}:
                jitter = rng.uniform(-0.36, 0.36)
            center = int((event_index + 0.5 + jitter) * period) % total_samples
            if mood == "ミシン" and event_index % 4 in (1, 2):
                center = (center + int(self.sample_rate * 0.055)) % total_samples
            freq = rng.choice((880.0, 1174.66, 1567.98, 2093.0))
            tail_sec = rng.choice((0.018, 0.026, 0.040))
            amp = float(profile["click_amp"]) * rng.uniform(0.65, 1.0)
            self._add_tone_event(samples, center, freq, tail_sec, amp, decay=22.0, brightness=0.18)

    def _add_tone_event(
        self,
        samples: list[float],
        center: int,
        freq: float,
        tail_sec: float,
        amp: float,
        decay: float,
        brightness: float,
    ) -> None:
        total_samples = len(samples)
        tail_samples = max(1, int(self.sample_rate * tail_sec))
        for offset in range(tail_samples):
            index = (center + offset) % total_samples
            t = offset / self.sample_rate
            env = math.exp(-decay * (offset / tail_samples))
            tone = math.sin(math.tau * freq * t)
            tone += brightness * math.sin(math.tau * freq * 2.01 * t)
            samples[index] += amp * env * tone

    def _add_grain_noise(self, samples: list[float], rng: random.Random, amp: float, mood: str) -> None:
        if amp <= 0:
            return
        total_samples = len(samples)
        grain = max(12, int(self.sample_rate * 0.018))
        previous = rng.uniform(-1.0, 1.0)
        values = [previous]
        for _ in range(max(1, total_samples // grain) + 2):
            target = rng.uniform(-1.0, 1.0)
            values.append(0.62 * previous + 0.38 * target)
            previous = values[-1]
        for index in range(total_samples):
            slot = index // grain
            blend = (index % grain) / grain
            value = values[slot] * (1.0 - blend) + values[slot + 1] * blend
            if mood == "雨":
                value += 0.42 * math.sin(math.tau * (index % 31) / 31)
            samples[index] += amp * 0.42 * value

    def _apply_circular_echo(self, samples: list[float], delay_samples: int, amount: float) -> None:
        if amount <= 0:
            return
        original = samples[:]
        total_samples = len(samples)
        for index in range(total_samples):
            samples[index] += original[(index - delay_samples) % total_samples] * amount
            samples[index] += original[(index - delay_samples * 2) % total_samples] * amount * 0.36

    def _apply_loop_crossfade(self, samples: list[float]) -> None:
        fade_samples = min(len(samples) // 8, int(self.sample_rate * 0.090))
        if fade_samples <= 1:
            return
        start = samples[:fade_samples]
        end = samples[-fade_samples:]
        for offset in range(fade_samples):
            t = offset / (fade_samples - 1)
            edge_weight = 0.18
            samples[offset] = start[offset] * (1.0 - edge_weight * (1.0 - t)) + end[offset] * edge_weight * (1.0 - t)
            samples[-fade_samples + offset] = end[offset] * (1.0 - edge_weight * t) + start[offset] * edge_weight * t

    def _soften_loop_edges(self, samples: list[float]) -> None:
        fade_samples = min(len(samples) // 8, int(self.sample_rate * 0.055))
        if fade_samples <= 1:
            return
        for offset in range(fade_samples):
            t = offset / (fade_samples - 1)
            fade_in = 0.5 - 0.5 * math.cos(math.pi * t)
            fade_out = 1.0 - fade_in
            samples[offset] *= fade_in
            samples[-fade_samples + offset] *= fade_out


class AceStepGeneratorAdapter(BaseGeneratorAdapter):
    """Prepared wrapper for a future ACE-Step integration.

    Phase 3 intentionally does not call ACE-Step yet.
    """

    name = "ace_step"

    def __init__(self, command: str | None = None) -> None:
        self.command = command or os.environ.get("ACE_STEP_COMMAND", "ace-step")

    def is_available(self) -> bool:
        return False

    def status_message(self) -> str:
        return "ACE-Step real generation is not connected yet. Mock mode only."

    def generate(self, request: GenerateRequest) -> GenerateResult:
        raise AdapterUnavailable(self.status_message())

    def _run_command(self, request: GenerateRequest) -> subprocess.CompletedProcess[str]:
        if "{" in self.command and "}" in self.command:
            command = self.command.format(
                prompt=request.prompt,
                duration=request.duration_sec,
                duration_sec=request.duration_sec,
                seed=request.seed,
                output=str(request.output_path),
                output_path=str(request.output_path),
            )
            return subprocess.run(
                command,
                shell=True,
                cwd=str(request.output_path.parent),
                capture_output=True,
                text=True,
                timeout=900,
            )

        args = [
            *self._split_command(),
            "--prompt",
            request.prompt,
            "--duration",
            str(request.duration_sec),
            "--seed",
            str(request.seed),
            "--output",
            str(request.output_path),
        ]
        return subprocess.run(
            args,
            cwd=str(request.output_path.parent),
            capture_output=True,
            text=True,
            timeout=900,
        )

    def _split_command(self) -> list[str]:
        return shlex.split(self.command, posix=os.name != "nt")


def _stable_mood_number(mood: str) -> int:
    return sum((index + 1) * ord(char) for index, char in enumerate(mood))
