from __future__ import annotations

import json
from dataclasses import dataclass
from fractions import Fraction
from pathlib import Path

from core.cli_checker import run_command


@dataclass(frozen=True)
class MediaInfo:
    file_name: str
    file_size_bytes: int
    duration: float
    width: int
    height: int
    fps: float
    video_codec: str
    audio_present: bool
    audio_codec: str


def _fraction_to_float(value: str | None) -> float:
    if not value or value == "0/0":
        return 0.0
    try:
        return float(Fraction(value))
    except Exception:
        try:
            return float(value)
        except Exception:
            return 0.0


def _float_value(value: object) -> float:
    try:
        return float(value)  # type: ignore[arg-type]
    except Exception:
        return 0.0


def probe_media(video_path: Path, ffprobe_path: str) -> MediaInfo:
    if not video_path.exists():
        raise FileNotFoundError(str(video_path))
    args = [
        ffprobe_path,
        "-v",
        "error",
        "-print_format",
        "json",
        "-show_format",
        "-show_streams",
        str(video_path),
    ]
    completed = run_command(args, timeout=25)
    if completed.returncode != 0:
        raise RuntimeError((completed.stderr or "ffprobe failed").strip())
    data = json.loads(completed.stdout)
    streams = data.get("streams", [])
    video_stream = next((stream for stream in streams if stream.get("codec_type") == "video"), {})
    audio_stream = next((stream for stream in streams if stream.get("codec_type") == "audio"), {})
    format_info = data.get("format", {})
    duration = _float_value(format_info.get("duration")) or _float_value(video_stream.get("duration"))
    fps = _fraction_to_float(video_stream.get("avg_frame_rate") or video_stream.get("r_frame_rate"))
    return MediaInfo(
        file_name=video_path.name,
        file_size_bytes=video_path.stat().st_size,
        duration=duration,
        width=int(video_stream.get("width") or 0),
        height=int(video_stream.get("height") or 0),
        fps=fps,
        video_codec=str(video_stream.get("codec_name") or ""),
        audio_present=bool(audio_stream),
        audio_codec=str(audio_stream.get("codec_name") or ""),
    )
