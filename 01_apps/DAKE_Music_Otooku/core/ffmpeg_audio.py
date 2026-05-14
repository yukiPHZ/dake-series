# -*- coding: utf-8 -*-
from __future__ import annotations

import subprocess
from dataclasses import dataclass, field
from pathlib import Path

from .app_config import resolve_tool_command
from .audio_probe import probe_duration


@dataclass
class AudioExportResult:
    success: bool
    files: list[Path] = field(default_factory=list)
    messages: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)


def _run(command: list[str], timeout: int = 120) -> tuple[bool, str]:
    try:
        result = subprocess.run(
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


def _fade_filter(duration: float | None, fade_in: float = 1.5, fade_out: float = 2.0) -> str:
    filters = [f"afade=t=in:st=0:d={fade_in:.2f}"]
    if duration and duration > 0.4:
        filters.append(f"afade=t=out:st={max(duration - fade_out, 0):.2f}:d={fade_out:.2f}")
    return ",".join(filters)


def export_audio_material(
    source_path: Path,
    audio_dir: Path,
    ffmpeg_command: str = "ffmpeg",
    ffprobe_command: str = "ffprobe",
    output_stem: str = "generated",
) -> AudioExportResult:
    result = AudioExportResult(success=False)
    resolved_ffmpeg = resolve_tool_command(ffmpeg_command)
    if not resolved_ffmpeg:
        result.errors.append("FFmpeg is required for audio export")
        return result
    resolved_ffprobe = resolve_tool_command(ffprobe_command)

    audio_dir.mkdir(parents=True, exist_ok=True)
    source_path = source_path.resolve()
    wav_output = audio_dir / f"{output_stem}.wav"
    mp3_output = audio_dir / f"{output_stem}.mp3"
    loop_output = audio_dir / "loop_preview.mp3"
    work_wav = audio_dir / f"_{output_stem}_export.wav"

    duration = probe_duration(source_path, resolved_ffprobe) if resolved_ffprobe else None
    wav_command = [
        resolved_ffmpeg,
        "-y",
        "-i",
        str(source_path),
        "-vn",
        "-ac",
        "2",
        "-ar",
        "44100",
        "-af",
        _fade_filter(duration),
        str(work_wav),
    ]
    ok, message = _run(wav_command)
    if not ok:
        result.errors.append(message or "wav export failed")
        return result

    try:
        if wav_output.exists():
            wav_output.unlink()
        work_wav.replace(wav_output)
    except Exception as exc:
        result.errors.append(str(exc))
        return result

    result.files.append(wav_output)
    result.messages.append(f"wrote {wav_output.name}")

    mp3_command = [
        resolved_ffmpeg,
        "-y",
        "-i",
        str(wav_output),
        "-codec:a",
        "libmp3lame",
        "-b:a",
        "192k",
        str(mp3_output),
    ]
    ok, message = _run(mp3_command)
    if ok:
        result.files.append(mp3_output)
        result.messages.append(f"wrote {mp3_output.name}")
    else:
        result.errors.append(message or "mp3 export failed")

    loop_seconds = min(duration, 30.0) if duration else 30.0
    loop_fade_start = max(loop_seconds - 2.0, 0)
    loop_command = [
        resolved_ffmpeg,
        "-y",
        "-i",
        str(source_path),
        "-t",
        f"{loop_seconds:.2f}",
        "-vn",
        "-ac",
        "2",
        "-ar",
        "44100",
        "-af",
        f"afade=t=in:st=0:d=1.50,afade=t=out:st={loop_fade_start:.2f}:d=2.00",
        "-codec:a",
        "libmp3lame",
        "-b:a",
        "192k",
        str(loop_output),
    ]
    ok, message = _run(loop_command)
    if ok:
        result.files.append(loop_output)
        result.messages.append(f"wrote {loop_output.name}")
    else:
        result.errors.append(message or "loop preview export failed")

    result.success = len(result.files) == 3
    if result.success:
        result.messages.append("audio package created")
    elif not result.errors:
        result.errors.append("audio package was incomplete")
    return result
