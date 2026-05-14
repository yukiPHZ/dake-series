# -*- coding: utf-8 -*-
from __future__ import annotations

import shutil
import subprocess
from dataclasses import dataclass, field
from pathlib import Path

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


def _fade_filter(duration: float | None) -> str:
    filters = [
        "loudnorm=I=-16:TP=-1.5:LRA=11",
        "afade=t=in:st=0:d=0.05",
    ]
    if duration and duration > 1.2:
        filters.append(f"afade=t=out:st={max(duration - 0.8, 0):.2f}:d=0.8")
    return ",".join(filters)


def export_audio_material(
    source_path: Path,
    audio_dir: Path,
    ffmpeg_command: str = "ffmpeg",
    ffprobe_command: str = "ffprobe",
) -> AudioExportResult:
    result = AudioExportResult(success=False)
    if shutil.which(ffmpeg_command) is None:
        result.errors.append("FFmpeg is required for audio export")
        return result

    audio_dir.mkdir(parents=True, exist_ok=True)
    source_path = source_path.resolve()
    wav_output = audio_dir / "generated.wav"
    mp3_output = audio_dir / "generated.mp3"
    loop_output = audio_dir / "loop_preview.mp3"
    work_wav = audio_dir / "_otooku_export.wav"

    duration = probe_duration(source_path, ffprobe_command) if shutil.which(ffprobe_command) else None
    wav_command = [
        ffmpeg_command,
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

    mp3_command = [
        ffmpeg_command,
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
    else:
        result.errors.append(message or "mp3 export failed")

    loop_seconds = 30.0
    loop_fade_start = max(loop_seconds - 0.8, 0)
    loop_command = [
        ffmpeg_command,
        "-y",
        "-stream_loop",
        "-1",
        "-i",
        str(wav_output),
        "-t",
        f"{loop_seconds:.2f}",
        "-vn",
        "-af",
        f"afade=t=in:st=0:d=0.05,afade=t=out:st={loop_fade_start:.2f}:d=0.8",
        "-codec:a",
        "libmp3lame",
        "-b:a",
        "192k",
        str(loop_output),
    ]
    ok, message = _run(loop_command)
    if ok:
        result.files.append(loop_output)
    else:
        result.errors.append(message or "loop preview export failed")

    result.success = bool(result.files)
    result.messages.append("audio export complete")
    return result

