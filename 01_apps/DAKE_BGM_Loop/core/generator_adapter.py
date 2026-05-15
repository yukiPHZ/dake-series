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
    """Small local WAV generator used for Phase 1 UI and file-flow checks."""

    name = "mock"
    sample_rate = 22050

    _profiles = {
        "のんき": {
            "base": (174.0, 196.0, 220.0),
            "pad": 0.26,
            "low": 0.10,
            "click": 0.035,
            "texture": 0.010,
            "clicks_per_15": 8,
        },
        "静か": {
            "base": (130.8, 146.8, 164.8),
            "pad": 0.22,
            "low": 0.08,
            "click": 0.004,
            "texture": 0.006,
            "clicks_per_15": 3,
        },
        "作業用": {
            "base": (146.8, 164.8, 196.0),
            "pad": 0.23,
            "low": 0.09,
            "click": 0.025,
            "texture": 0.008,
            "clicks_per_15": 10,
        },
        "神社": {
            "base": (164.8, 196.0, 246.9),
            "pad": 0.21,
            "low": 0.06,
            "click": 0.018,
            "texture": 0.009,
            "clicks_per_15": 4,
        },
        "雨": {
            "base": (110.0, 130.8, 146.8),
            "pad": 0.21,
            "low": 0.07,
            "click": 0.020,
            "texture": 0.024,
            "clicks_per_15": 18,
        },
        "夜": {
            "base": (98.0, 110.0, 130.8),
            "pad": 0.23,
            "low": 0.12,
            "click": 0.005,
            "texture": 0.012,
            "clicks_per_15": 3,
        },
        "ミシン": {
            "base": (146.8, 164.8, 174.0),
            "pad": 0.18,
            "low": 0.08,
            "click": 0.055,
            "texture": 0.012,
            "clicks_per_15": 22,
        },
        "コード": {
            "base": (130.8, 146.8, 164.8),
            "pad": 0.20,
            "low": 0.08,
            "click": 0.045,
            "texture": 0.010,
            "clicks_per_15": 16,
        },
        "余白": {
            "base": (98.0, 123.5, 146.8),
            "pad": 0.16,
            "low": 0.05,
            "click": 0.003,
            "texture": 0.004,
            "clicks_per_15": 2,
        },
    }

    def generate(self, request: GenerateRequest) -> GenerateResult:
        request.output_path.parent.mkdir(parents=True, exist_ok=True)
        self._write_mock_wav(request)
        return GenerateResult(
            output_path=request.output_path,
            adapter_name=self.name,
            message="Mock loop generated.",
        )

    def _write_mock_wav(self, request: GenerateRequest) -> None:
        profile = self._profiles.get(request.mood, self._profiles["のんき"])
        duration_sec = max(1, int(request.duration_sec))
        total_samples = self.sample_rate * duration_sec
        rng = random.Random(request.seed + _stable_mood_number(request.mood))

        base_freq = rng.choice(profile["base"])
        low_freq = base_freq / rng.choice((2.0, 2.5, 3.0))
        air_freq = base_freq * rng.choice((2.01, 2.49, 3.02))

        base_cycles = max(1, round(base_freq * duration_sec))
        low_cycles = max(1, round(low_freq * duration_sec))
        air_cycles = max(1, round(air_freq * duration_sec))
        slow_cycles = rng.choice((1, 2, 3))
        slow_cycles_2 = rng.choice((2, 3, 5))
        click_cycles = max(1, int(profile["clicks_per_15"] * duration_sec / 15))
        click_period = max(1, total_samples // click_cycles)
        click_width = max(6, int(self.sample_rate * rng.choice((0.004, 0.006, 0.009))))
        high_click_cycles = max(1, round(rng.choice((880.0, 1175.0, 1568.0)) * duration_sec))
        texture_cycles_a = max(1, round(rng.choice((37.0, 43.0, 53.0)) * duration_sec))
        texture_cycles_b = max(1, round(rng.choice((71.0, 89.0, 97.0)) * duration_sec))
        phase_a = rng.random() * math.tau
        phase_b = rng.random() * math.tau

        frames = bytearray()
        for index in range(total_samples):
            loop_phase = index / total_samples
            motion = 0.72 + 0.20 * math.sin(math.tau * slow_cycles * loop_phase + phase_a)
            motion += 0.08 * math.sin(math.tau * slow_cycles_2 * loop_phase + phase_b)

            pad = math.sin(math.tau * base_cycles * loop_phase + phase_a)
            pad += 0.38 * math.sin(math.tau * air_cycles * loop_phase + phase_b)
            low = math.sin(math.tau * low_cycles * loop_phase + phase_b)

            position = index % click_period
            click_distance = min(position, click_period - position)
            click_env = math.exp(-((click_distance / click_width) ** 2))
            click = click_env * math.sin(math.tau * high_click_cycles * loop_phase)

            texture = math.sin(math.tau * texture_cycles_a * loop_phase + phase_a)
            texture *= math.sin(math.tau * texture_cycles_b * loop_phase + phase_b)

            sample = 0.0
            sample += profile["pad"] * motion * pad
            sample += profile["low"] * low
            sample += profile["click"] * click
            sample += profile["texture"] * texture
            sample *= 0.55
            sample = max(-0.92, min(0.92, sample))
            frames.extend(struct.pack("<h", int(sample * 32767)))

        with wave.open(str(request.output_path), "wb") as wav_file:
            wav_file.setnchannels(1)
            wav_file.setsampwidth(2)
            wav_file.setframerate(self.sample_rate)
            wav_file.writeframes(frames)


class AceStepGeneratorAdapter(BaseGeneratorAdapter):
    """Thin wrapper for a future ACE-Step command-line integration."""

    name = "ace_step"

    def __init__(self, command: str | None = None) -> None:
        self.command = command or os.environ.get("ACE_STEP_COMMAND", "ace-step")

    def is_available(self) -> bool:
        parts = self._split_command()
        return bool(parts and shutil.which(parts[0]))

    def status_message(self) -> str:
        if self.is_available():
            return "ACE-Step adapter is available."
        return "ACE-Step is not configured. Mock mode only."

    def generate(self, request: GenerateRequest) -> GenerateResult:
        if not self.is_available():
            raise AdapterUnavailable(self.status_message())

        request.output_path.parent.mkdir(parents=True, exist_ok=True)
        completed = self._run_command(request)
        if completed.returncode != 0:
            stderr = (completed.stderr or "").strip()
            raise RuntimeError(f"ACE-Step generation failed: {stderr or completed.returncode}")
        if not request.output_path.is_file():
            raise RuntimeError("ACE-Step finished, but the expected WAV file was not created.")
        return GenerateResult(
            output_path=request.output_path,
            adapter_name=self.name,
            message="ACE-Step loop generated.",
        )

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

