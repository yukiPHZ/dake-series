# -*- coding: utf-8 -*-
from __future__ import annotations

import math
import shutil
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
    variation: str
    variation_summary: str
    variation_use: str
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


@dataclass(frozen=True)
class TinyAmbientVariation:
    key: str
    wav_path: Path
    mp3_path: Path | None
    plan: TinyAmbientPlan


@dataclass
class TinyAmbientResult:
    success: bool
    wav_path: Path | None = None
    mp3_path: Path | None = None
    plan: TinyAmbientPlan | None = None
    variations: list[TinyAmbientVariation] = field(default_factory=list)
    messages: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)


VARIATION_KEYS = ("A", "B", "C")


def _contains_any(text: str, words: tuple[str, ...]) -> bool:
    return any(word in text for word in words)


def _variation_note(source: str, variation: str) -> tuple[str, str]:
    if _contains_any(source, ("borinef", "ember")):
        notes = {
            "A": ("low warm drone", "静かな余熱、夜、内省、長尺背景向け"),
            "B": ("embers pulse", "作業動画、タイピング、低温の制作ログ向け"),
            "C": ("distant air", "遠い余白、暗い画面、静かなShorts向け"),
        }
    elif _contains_any(source, ("yukiz", "稼働", "work", "code", "sewing")):
        notes = {
            "A": ("quiet machine hum", "深夜のコード作業、ミシン、待機画面向け"),
            "B": ("working pulse", "作業動画、タイピング、制作ショート向け"),
            "C": ("late night air", "作業終わり、余白、静かな制作ログ向け"),
        }
    elif _contains_any(source, ("holiday", "jinja", "shrine")):
        notes = {
            "A": ("calm bell", "神社、朝、空、短い映像向け"),
            "B": ("light rhythm", "散歩映像、朝の切り替わり、Shorts向け"),
            "C": ("morning wind", "余白、風、写真の背景向け"),
        }
    elif _contains_any(source, ("blue", "memory")):
        notes = {
            "A": ("soft memory drone", "写真、余韻、静かな回想向け"),
            "B": ("slow memory pulse", "制作ログ、短い映像のゆるい動き向け"),
            "C": ("blue floating air", "Japan Memory Lane、余白、夜明け前向け"),
        }
    else:
        notes = {
            "A": ("quiet stable drone", "作業・読書・深夜向け"),
            "B": ("slight pulse", "作業動画・タイピング向け"),
            "C": ("airy floating", "余白・遠景・静かなShorts向け"),
        }
    return notes.get(variation, notes["A"])


def plan_tiny_ambient(
    user_text: str,
    direction: MusicDirection,
    preset: MusicPreset | None = None,
    variation: str = "A",
) -> TinyAmbientPlan:
    variation = variation.upper()
    if variation not in VARIATION_KEYS:
        variation = "A"
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
    note_source = preset_hint or source

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

    if variation == "A":
        frequencies = tuple(max(120.0, min(frequency * 0.98, 720.0)) for frequency in frequencies)
        noise_volume *= 0.86
        tone_volumes = tuple(volume * 0.94 for volume in tone_volumes)
        delays = (0.0, min(4.2, duration * 0.30), min(8.8, duration * 0.62))
        duration_scales = (0.70, 0.55, 0.42)
        stereo_offsets = (0, 24, 48)
        tremolo_depth = 0.0
    elif variation == "B":
        duration = max(12.0, min(duration - 1.0, 20.0))
        frequencies = (
            max(120.0, min(frequencies[0] * 1.00, 720.0)),
            max(120.0, min(frequencies[1] * 1.03, 720.0)),
            max(120.0, min(frequencies[2] * 1.06, 720.0)),
        )
        pulse = True
        noise_volume = min(noise_volume + 0.03, 0.20)
        tone_volumes = (tone_volumes[0] * 0.86, tone_volumes[1] * 0.58, tone_volumes[2] * 0.50)
        delays = (0.0, min(2.6, duration * 0.22), min(6.4, duration * 0.48))
        duration_scales = (0.52, 0.48, 0.42)
        stereo_offsets = (0, 36, 72)
        tremolo_depth = 0.13
    else:
        duration = min(duration + 1.0, 20.0)
        frequencies = (
            max(120.0, min(frequencies[0] * 1.07, 720.0)),
            max(120.0, min(frequencies[1] * 1.10, 720.0)),
            max(120.0, min(frequencies[2] * 1.12, 720.0)),
        )
        color = "white"
        noise_volume = min(noise_volume + 0.04, 0.18)
        tone_volumes = (tone_volumes[0] * 0.62, tone_volumes[1] * 0.48, tone_volumes[2] * 0.42)
        delays = (0.0, min(5.0, duration * 0.34), min(9.5, duration * 0.66))
        duration_scales = (0.60, 0.52, 0.46)
        stereo_offsets = (48, 94, 140)
        tremolo_depth = 0.0

    base_frequency = frequencies[0]
    tone_durations = (
        max(1.0, min(duration - delays[0] + 0.4, duration * duration_scales[0])),
        max(1.0, min(duration - delays[1] + 0.4, duration * duration_scales[1])),
        max(1.0, min(duration - delays[2] + 0.4, duration * duration_scales[2])),
    )
    fade_out_start = max(duration - 2.2, 0.0)
    noise_lowpass = 3200 if color == "white" else 1800
    tone_parts = []
    for index, (delay, tone_duration, tone_volume, stereo_offset) in enumerate(
        zip(delays, tone_durations, tone_volumes, stereo_offsets)
    ):
        delay_ms = int(delay * 1000)
        right_delay_ms = delay_ms + stereo_offset
        tone_fade_out = max(tone_duration - 1.6, 0.0)
        tone_parts.append(
            f"[{index}:a]volume={tone_volume:.2f},"
            f"afade=t=in:st=0:d=0.60,"
            f"afade=t=out:st={tone_fade_out:.2f}:d=1.60,"
            f"pan=stereo|c0=c0|c1=c0,"
            f"adelay={delay_ms}|{right_delay_ms}[a{index}];"
        )
    movement_filter = f",tremolo=f=0.85:d={tremolo_depth:.2f}" if pulse and tremolo_depth > 0.0 else ""
    variation_summary, variation_use = _variation_note(note_source, variation)
    filter_text = (
        f"{''.join(tone_parts)}"
        f"[3:a]highpass=f=120,lowpass=f={noise_lowpass},volume={noise_volume:.2f},"
        f"pan=stereo|c0=c0|c1=c0[air];"
        f"[a0][a1][a2][air]amix=inputs=4:duration=longest:normalize=0,"
        f"atrim=0:{duration:.2f},"
        f"afade=t=in:st=0:d=0.90,"
        f"afade=t=out:st={fade_out_start:.2f}:d=2.20"
        f"{movement_filter},"
        f"alimiter=limit=0.92,"
        f"volume=1.12[out]"
    )
    return TinyAmbientPlan(
        variation=variation,
        variation_summary=variation_summary,
        variation_use=variation_use,
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


def _ffmpeg_wav_command(resolved_ffmpeg: str, plan: TinyAmbientPlan, wav_path: Path) -> list[str]:
    return [
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


def _write_variation_notes(audio_dir: Path, variations: list[TinyAmbientVariation], preset: MusicPreset | None) -> None:
    notes_dir = audio_dir.parent / "notes"
    notes_dir.mkdir(parents=True, exist_ok=True)
    lines = ["# Ambient Variation Notes", ""]
    lines.extend(["Preset:", preset.name if preset else "Custom", ""])
    for variation in variations:
        plan = variation.plan
        lines.extend(
            [
                f"{variation.key}:",
                plan.variation_summary,
                plan.variation_use,
                f"File: {variation.wav_path.name}",
                f"Duration: {plan.duration:.1f}s",
                "Frequencies: " + ", ".join(f"{frequency:.1f}Hz" for frequency in plan.frequencies),
                "",
            ]
        )
    (notes_dir / "variation_notes.txt").write_text("\n".join(lines).strip() + "\n", encoding="utf-8")


def _copy_legacy_preview(result: TinyAmbientResult, audio_dir: Path) -> None:
    if not result.variations:
        return
    first_variation = result.variations[0]
    legacy_wav = audio_dir / "generated_preview.wav"
    legacy_mp3 = audio_dir / "generated_preview.mp3"
    try:
        shutil.copyfile(first_variation.wav_path, legacy_wav)
        result.messages.append(legacy_wav.name)
        result.wav_path = first_variation.wav_path
        result.plan = first_variation.plan
    except Exception as exc:
        result.errors.append(f"legacy wav copy failed: {exc}")
    if first_variation.mp3_path:
        try:
            shutil.copyfile(first_variation.mp3_path, legacy_mp3)
            result.messages.append(legacy_mp3.name)
            result.mp3_path = first_variation.mp3_path
        except Exception as exc:
            result.errors.append(f"legacy mp3 copy failed: {exc}")


def generate_tiny_ambient(
    user_text: str,
    direction: MusicDirection,
    audio_dir: Path,
    preset: MusicPreset | None = None,
    ffmpeg_command: str = "ffmpeg",
) -> TinyAmbientResult:
    audio_dir.mkdir(parents=True, exist_ok=True)
    first_plan = plan_tiny_ambient(user_text, direction, preset, "A")
    result = TinyAmbientResult(success=False, plan=first_plan)

    resolved_ffmpeg = resolve_tool_command(ffmpeg_command)
    for variation_key in VARIATION_KEYS:
        plan = plan_tiny_ambient(user_text, direction, preset, variation_key)
        wav_path = audio_dir / f"generated_preview_{variation_key}.wav"
        mp3_path = audio_dir / f"generated_preview_{variation_key}.mp3"
        mp3_output: Path | None = None
        generated = False

        if resolved_ffmpeg:
            ok, message = _run(_ffmpeg_wav_command(resolved_ffmpeg, plan, wav_path))
            generated = ok and wav_path.exists()
            if not generated and message:
                result.errors.append(f"Variation {variation_key}: {message}")

        if not generated:
            try:
                _write_python_fallback_wav(wav_path, plan)
                generated = wav_path.exists()
                if generated:
                    result.errors.append(f"Variation {variation_key}: used Python fallback wav")
            except Exception as exc:
                result.errors.append(f"Variation {variation_key}: {exc}")

        if not generated:
            continue

        if resolved_ffmpeg:
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
                mp3_output = mp3_path
            elif mp3_message:
                result.errors.append(f"Variation {variation_key} mp3: {mp3_message}")

        result.variations.append(
            TinyAmbientVariation(
                key=variation_key,
                wav_path=wav_path,
                mp3_path=mp3_output,
                plan=plan,
            )
        )
        result.messages.append(wav_path.name)
        if mp3_output:
            result.messages.append(mp3_output.name)

    result.success = bool(result.variations)
    if result.success:
        _copy_legacy_preview(result, audio_dir)
        _write_variation_notes(audio_dir, result.variations, preset)
    return result
