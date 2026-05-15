# -*- coding: utf-8 -*-
from __future__ import annotations

import math
import struct
import wave
from dataclasses import dataclass, field
from pathlib import Path

from .app_config import resolve_tool_command
from .presets import MusicPreset
from .prompt_builder import MusicDirection
from .subprocess_utils import run_hidden


@dataclass(frozen=True)
class TinyAmbientPlan:
    duration: float
    base_frequency: float
    color: str
    pulse: bool
    filter_text: str


@dataclass
class TinyAmbientResult:
    success: bool
    wav_path: Path | None = None
    mp3_path: Path | None = None
    plan: TinyAmbientPlan | None = None
    messages: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)


def _contains_any(text: str, words: tuple[str, ...]) -> bool:
    return any(word in text for word in words)


def plan_tiny_ambient(user_text: str, direction: MusicDirection, preset: MusicPreset | None = None) -> TinyAmbientPlan:
    source = " ".join(
        [
            user_text,
            direction.mood,
            direction.texture,
            direction.usage_idea,
            preset.name if preset else "",
            preset.mood if preset else "",
            preset.texture if preset else "",
            " ".join(preset.tags) if preset else "",
        ]
    ).lower()

    duration = 9.0
    base_frequency = 220.0
    color = "pink"
    pulse = False

    if _contains_any(source, ("borinef", "ember", "low heat", "余熱", "低温")):
        duration = 12.0
        base_frequency = 165.0
        color = "pink"
    elif _contains_any(source, ("yukiz", "稼働", "work", "code", "sewing", "ミシン", "コード")):
        duration = 10.0
        base_frequency = 220.0
        color = "pink"
        pulse = True
    elif _contains_any(source, ("朝", "morning", "shrine", "jinja", "神社", "holiday")):
        duration = 8.0
        base_frequency = 360.0
        color = "white"
    elif _contains_any(source, ("深夜", "midnight", "night", "夜", "blue", "memory")):
        duration = 11.0
        base_frequency = 145.0
        color = "pink"

    hash_shift = sum(ord(char) for char in user_text) % 31
    base_frequency += hash_shift - 15
    base_frequency = max(120.0, min(base_frequency, 600.0))
    fade_out_start = max(duration - 1.8, 0.0)
    noise_lowpass = 2600 if color == "white" else 1800
    filter_text = (
        f"[0:a]volume=1.00[a0];"
        f"[1:a]volume=0.35[a1];"
        f"[2:a]highpass=f=160,lowpass=f={noise_lowpass},volume=0.14[a2];"
        f"[a0][a1][a2]amix=inputs=3:duration=longest:normalize=0,"
        f"afade=t=in:st=0:d=0.60,"
        f"afade=t=out:st={fade_out_start:.2f}:d=1.80,"
        f"alimiter=limit=0.92,"
        f"volume=1.40[out]"
    )
    return TinyAmbientPlan(
        duration=duration,
        base_frequency=base_frequency,
        color=color,
        pulse=pulse,
        filter_text=filter_text,
    )


def _run(command: list[str], timeout: int = 60) -> tuple[bool, str]:
    try:
        result = run_hidden(
            command,
            capture_output=True,
            text=True,
            timeout=timeout,
            check=False,
        )
    except Exception as exc:
        return False, str(exc)
    if result.returncode != 0:
        return False, (result.stderr or result.stdout or "").strip()
    return True, (result.stdout or result.stderr or "").strip()


def _write_python_fallback_wav(path: Path, plan: TinyAmbientPlan) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    sample_rate = 44100
    frames = int(sample_rate * plan.duration)
    frequency = plan.base_frequency
    with wave.open(str(path), "wb") as audio_file:
        audio_file.setnchannels(2)
        audio_file.setsampwidth(2)
        audio_file.setframerate(sample_rate)
        for index in range(frames):
            t = index / sample_rate
            fade_in = min(t / 0.6, 1.0)
            fade_out = min((plan.duration - t) / 1.8, 1.0)
            envelope = max(0.0, min(fade_in, fade_out))
            pulse = 0.72 + 0.28 * math.sin(2.0 * math.pi * 1.15 * t) if plan.pulse else 1.0
            value = (
                math.sin(2.0 * math.pi * frequency * t) * 0.34
                + math.sin(2.0 * math.pi * frequency * 1.5 * t) * 0.12
            )
            sample = int(32767 * value * envelope * pulse)
            sample = max(-30000, min(30000, sample))
            packed = struct.pack("<hh", sample, sample)
            audio_file.writeframes(packed)


def generate_tiny_ambient(
    user_text: str,
    direction: MusicDirection,
    audio_dir: Path,
    preset: MusicPreset | None = None,
    ffmpeg_command: str = "ffmpeg",
) -> TinyAmbientResult:
    audio_dir.mkdir(parents=True, exist_ok=True)
    wav_path = audio_dir / "generated_preview.wav"
    mp3_path = audio_dir / "generated_preview.mp3"
    plan = plan_tiny_ambient(user_text, direction, preset)
    result = TinyAmbientResult(success=False, wav_path=wav_path, plan=plan)

    resolved_ffmpeg = resolve_tool_command(ffmpeg_command)
    if resolved_ffmpeg:
        primary = plan.base_frequency
        overtone = min(plan.base_frequency * 1.5, 880.0)
        wav_command = [
            resolved_ffmpeg,
            "-y",
            "-f",
            "lavfi",
            "-i",
            f"sine=frequency={primary:.2f}:duration={plan.duration:.2f}:sample_rate=44100",
            "-f",
            "lavfi",
            "-i",
            f"sine=frequency={overtone:.2f}:duration={plan.duration:.2f}:sample_rate=44100",
            "-f",
            "lavfi",
            "-i",
            f"anoisesrc=color={plan.color}:amplitude=0.05:duration={plan.duration:.2f}:sample_rate=44100",
            "-filter_complex",
            plan.filter_text,
            "-map",
            "[out]",
            "-ac",
            "2",
            "-ar",
            "44100",
            str(wav_path),
        ]
        ok, message = _run(wav_command)
        if ok and wav_path.exists():
            result.success = True
            result.messages.append("generated_preview.wav")
            mp3_command = [
                resolved_ffmpeg,
                "-y",
                "-i",
                str(wav_path),
                "-codec:a",
                "libmp3lame",
                "-b:a",
                "192k",
                str(mp3_path),
            ]
            mp3_ok, mp3_message = _run(mp3_command)
            if mp3_ok and mp3_path.exists():
                result.mp3_path = mp3_path
                result.messages.append("generated_preview.mp3")
            elif mp3_message:
                result.errors.append(mp3_message)
            return result
        if message:
            result.errors.append(message)

    try:
        _write_python_fallback_wav(wav_path, plan)
        result.success = wav_path.exists()
        if result.success:
            result.messages.append("generated_preview.wav")
            result.errors.append("FFmpeg ambient generation unavailable; used Python fallback wav")
    except Exception as exc:
        result.errors.append(str(exc))
    return result
