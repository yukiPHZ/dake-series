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
    frequencies: tuple[float, float, float]
    delays: tuple[float, float, float]
    tone_durations: tuple[float, float, float]
    color: str
    pulse: bool
    noise_volume: float
    tone_volumes: tuple[float, float, float]
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

    duration = 14.0
    base_frequency = 220.0
    frequencies = (220.0, 277.0, 330.0)
    color = "pink"
    pulse = False
    noise_volume = 0.10
    tone_volumes = (0.78, 0.48, 0.38)
    preset_hint = ""
    if preset:
        preset_hint = " ".join([preset.name, preset.mood, preset.texture, " ".join(preset.tags)]).lower()

    if _contains_any(preset_hint, ("borinef", "ember")) or (not preset and _contains_any(source, ("borinef", "ember", "low heat", "余熱", "低温"))):
        duration = 16.0
        frequencies = (160.0, 220.0, 260.0)
        base_frequency = frequencies[0]
        color = "pink"
        noise_volume = 0.16
        tone_volumes = (0.84, 0.46, 0.34)
    elif _contains_any(preset_hint, ("yukiz", "稼働")) or (not preset and _contains_any(source, ("yukiz", "稼働", "work", "code", "sewing", "ミシン", "コード"))):
        duration = 14.0
        frequencies = (220.0, 330.0, 440.0)
        base_frequency = frequencies[0]
        color = "pink"
        pulse = True
        noise_volume = 0.11
        tone_volumes = (0.72, 0.46, 0.36)
    elif _contains_any(preset_hint, ("holiday", "jinja", "shrine")) or (not preset and _contains_any(source, ("朝", "morning", "shrine", "jinja", "神社", "holiday"))):
        duration = 13.0
        frequencies = (330.0, 440.0, 660.0)
        base_frequency = frequencies[0]
        color = "white"
        noise_volume = 0.06
        tone_volumes = (0.58, 0.40, 0.30)
    elif _contains_any(preset_hint, ("blue", "memory")) or (not preset and _contains_any(source, ("深夜", "midnight", "night", "夜", "blue", "memory"))):
        duration = 16.0
        frequencies = (196.0, 247.0, 330.0) if _contains_any(source, ("blue", "memory")) else (145.0, 196.0, 247.0)
        base_frequency = frequencies[0]
        color = "pink"
        noise_volume = 0.12
        tone_volumes = (0.76, 0.44, 0.34)
    elif _contains_any(source, ("雨", "rain")):
        duration = 15.0
        frequencies = (185.0, 247.0, 294.0)
        base_frequency = frequencies[0]
        noise_volume = 0.18

    hash_shift = sum(ord(char) for char in user_text) % 31
    frequency_shift = hash_shift - 15
    frequencies = tuple(max(120.0, min(frequency + frequency_shift, 720.0)) for frequency in frequencies)
    base_frequency = frequencies[0]
    delays = (0.0, min(4.0, duration * 0.28), min(8.5, duration * 0.58))
    tone_durations = (
        min(duration, duration * 0.64),
        min(duration - delays[1] + 0.4, duration * 0.54),
        min(duration - delays[2] + 0.4, duration * 0.42),
    )
    fade_out_start = max(duration - 2.2, 0.0)
    noise_lowpass = 2600 if color == "white" else 1800
    tone_parts = []
    for index, (delay, tone_duration, tone_volume) in enumerate(zip(delays, tone_durations, tone_volumes)):
        delay_ms = int(delay * 1000)
        tone_fade_out = max(tone_duration - 1.6, 0.0)
        tone_parts.append(
            f"[{index}:a]volume={tone_volume:.2f},"
            f"afade=t=in:st=0:d=0.60,"
            f"afade=t=out:st={tone_fade_out:.2f}:d=1.60,"
            f"adelay={delay_ms}|{delay_ms}[a{index}];"
        )
    movement_filter = ",tremolo=f=0.85:d=0.08" if pulse else ""
    filter_text = (
        f"{''.join(tone_parts)}"
        f"[3:a]highpass=f=120,lowpass=f={noise_lowpass},volume={noise_volume:.2f}[air];"
        f"[a0][a1][a2][air]amix=inputs=4:duration=longest:normalize=0,"
        f"atrim=0:{duration:.2f},"
        f"afade=t=in:st=0:d=0.90,"
        f"afade=t=out:st={fade_out_start:.2f}:d=2.20"
        f"{movement_filter},"
        f"alimiter=limit=0.92,"
        f"volume=1.12[out]"
    )
    return TinyAmbientPlan(
        duration=duration,
        base_frequency=base_frequency,
        frequencies=frequencies,
        delays=delays,
        tone_durations=tone_durations,
        color=color,
        pulse=pulse,
        noise_volume=noise_volume,
        tone_volumes=tone_volumes,
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
    with wave.open(str(path), "wb") as audio_file:
        audio_file.setnchannels(2)
        audio_file.setsampwidth(2)
        audio_file.setframerate(sample_rate)
        for index in range(frames):
            t = index / sample_rate
            fade_in = min(t / 0.6, 1.0)
            fade_out = min((plan.duration - t) / 1.8, 1.0)
            envelope = max(0.0, min(fade_in, fade_out))
            movement = 0.92 + 0.08 * math.sin(2.0 * math.pi * 0.85 * t) if plan.pulse else 1.0
            value = 0.0
            for frequency, delay, tone_duration, tone_volume in zip(
                plan.frequencies,
                plan.delays,
                plan.tone_durations,
                plan.tone_volumes,
            ):
                local_t = t - delay
                if local_t < 0.0 or local_t > tone_duration:
                    continue
                tone_fade_in = min(local_t / 0.6, 1.0)
                tone_fade_out = min((tone_duration - local_t) / 1.6, 1.0)
                tone_envelope = max(0.0, min(tone_fade_in, tone_fade_out))
                value += math.sin(2.0 * math.pi * frequency * local_t) * tone_volume * 0.36 * tone_envelope
            noise = (
                math.sin(2.0 * math.pi * 713.0 * t)
                + math.sin(2.0 * math.pi * 1189.0 * t)
            ) * plan.noise_volume * 0.06
            value = (value + noise) * movement
            sample = int(32767 * value * envelope)
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
        wav_command = [
            resolved_ffmpeg,
            "-y",
            "-f",
            "lavfi",
            "-i",
            f"sine=frequency={plan.frequencies[0]:.2f}:duration={plan.tone_durations[0]:.2f}:sample_rate=44100",
            "-f",
            "lavfi",
            "-i",
            f"sine=frequency={plan.frequencies[1]:.2f}:duration={plan.tone_durations[1]:.2f}:sample_rate=44100",
            "-f",
            "lavfi",
            "-i",
            f"sine=frequency={plan.frequencies[2]:.2f}:duration={plan.tone_durations[2]:.2f}:sample_rate=44100",
            "-f",
            "lavfi",
            "-i",
            f"anoisesrc=color={plan.color}:amplitude=0.06:duration={plan.duration:.2f}:sample_rate=44100",
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
