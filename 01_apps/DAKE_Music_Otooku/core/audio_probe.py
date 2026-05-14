# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import subprocess
from dataclasses import dataclass
from pathlib import Path

from .app_config import resolve_tool_command

@dataclass(frozen=True)
class AudioInfo:
    path: Path
    duration: float | None
    codec: str
    sample_rate: str
    channels: str

    def display_lines(self) -> list[str]:
        duration_text = "--" if self.duration is None else f"{self.duration:.1f}s"
        return [
            f"file name: {self.path.name}",
            f"duration: {duration_text}",
            f"codec: {self.codec}",
            f"sample rate: {self.sample_rate}",
            f"channels: {self.channels}",
        ]


def _run_ffprobe(audio_path: Path, ffprobe_command: str = "ffprobe") -> dict | None:
    resolved_ffprobe = resolve_tool_command(ffprobe_command) or ffprobe_command
    try:
        result = subprocess.run(
            [
                resolved_ffprobe,
                "-v",
                "error",
                "-select_streams",
                "a:0",
                "-show_entries",
                "stream=codec_name,sample_rate,channels,duration:format=duration",
                "-of",
                "json",
                str(audio_path),
            ],
            capture_output=True,
            text=True,
            timeout=10,
            check=False,
        )
    except Exception:
        return None
    if result.returncode != 0:
        return None
    try:
        return json.loads(result.stdout)
    except Exception:
        return None


def probe_duration(audio_path: Path, ffprobe_command: str = "ffprobe") -> float | None:
    info = probe_audio_info(audio_path, ffprobe_command)
    return info.duration if info else None


def probe_audio_info(audio_path: Path, ffprobe_command: str = "ffprobe") -> AudioInfo | None:
    data = _run_ffprobe(audio_path, ffprobe_command)
    if not data:
        return None

    streams = data.get("streams") or []
    stream = streams[0] if streams else {}
    format_info = data.get("format") or {}

    duration_raw = stream.get("duration") or format_info.get("duration")
    try:
        duration = float(duration_raw) if duration_raw is not None else None
    except Exception:
        duration = None

    sample_rate = str(stream.get("sample_rate") or "--")
    if sample_rate != "--":
        sample_rate = f"{sample_rate} Hz"

    return AudioInfo(
        path=audio_path,
        duration=duration,
        codec=str(stream.get("codec_name") or "--"),
        sample_rate=sample_rate,
        channels=str(stream.get("channels") or "--"),
    )
