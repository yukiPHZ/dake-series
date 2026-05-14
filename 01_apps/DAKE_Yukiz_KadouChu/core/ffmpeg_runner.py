from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from core.cli_checker import run_command


@dataclass(frozen=True)
class PreviewResult:
    created: bool
    path: Path
    message: str
    used_encoder: str


def _duration_arg(start_seconds: float, end_seconds: float) -> str:
    return f"{max(1.0, end_seconds - start_seconds):.3f}"


def _build_preview_args(
    ffmpeg_path: str,
    video_path: Path,
    output_path: Path,
    start_seconds: float,
    end_seconds: float,
    encoder: str,
) -> list[str]:
    video_codec_args = ["-c:v", "h264_nvenc", "-preset", "p4", "-cq", "23"] if encoder == "nvenc" else [
        "-c:v",
        "libx264",
        "-preset",
        "veryfast",
        "-crf",
        "23",
    ]
    return [
        ffmpeg_path,
        "-y",
        "-ss",
        f"{max(0.0, start_seconds):.3f}",
        "-i",
        str(video_path),
        "-t",
        _duration_arg(start_seconds, end_seconds),
        "-map",
        "0:v:0",
        "-map",
        "0:a?",
        *video_codec_args,
        "-c:a",
        "aac",
        "-b:a",
        "128k",
        "-movflags",
        "+faststart",
        str(output_path),
    ]


def create_preview_clip(
    video_path: Path,
    output_path: Path,
    start_seconds: float,
    end_seconds: float,
    ffmpeg_path: str,
    use_nvenc: bool,
) -> PreviewResult:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    encoders = ["nvenc", "cpu"] if use_nvenc else ["cpu"]
    messages: list[str] = []
    for encoder in encoders:
        args = _build_preview_args(ffmpeg_path, video_path, output_path, start_seconds, end_seconds, encoder)
        completed = run_command(args, timeout=180)
        if completed.returncode == 0 and output_path.exists():
            return PreviewResult(True, output_path, "created", "h264_nvenc" if encoder == "nvenc" else "libx264")
        messages.append((completed.stderr or completed.stdout or f"{encoder} encode failed").strip()[-600:])
    return PreviewResult(False, output_path, "\n".join(messages), "none")
