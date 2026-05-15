from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

from core.app_config import LOG_TEXT
from core.cli_checker import run_command

LogCallback = Callable[[str], None]

VERTICAL_SIZE = "1080x1920"
VERTICAL_POLICY = "background blur + foreground centered"
HORIZONTAL_SIZE = "1920x1080"
HORIZONTAL_POLICY = "background blur + foreground centered"


def _read_json(path: Path) -> Any:
    if not path.exists() or not path.is_file():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8", errors="replace"))
    except Exception:
        return None


def _time_to_seconds(value: object) -> float:
    if isinstance(value, (int, float)):
        return max(0.0, float(value))
    text = str(value or "").strip()
    if not text or text == "--":
        return 0.0
    if re.fullmatch(r"\d+(\.\d+)?", text):
        return max(0.0, float(text))
    match = re.fullmatch(r"(\d{2}):(\d{2}):(\d{2})(?:[,.](\d{1,3}))?", text)
    if not match:
        return 0.0
    hours, minutes, seconds, millis = match.groups()
    total = int(hours) * 3600 + int(minutes) * 60 + int(seconds)
    if millis:
        total += int(millis.ljust(3, "0")[:3]) / 1000
    return max(0.0, float(total))


def _duration_seconds(short: dict[str, Any]) -> float:
    if short.get("duration_seconds") is not None:
        return max(0.0, float(short.get("duration_seconds") or 0))
    duration = _time_to_seconds(short.get("duration"))
    if duration > 0:
        return duration
    start = _time_to_seconds(short.get("start_seconds", short.get("start")))
    end = _time_to_seconds(short.get("end_seconds", short.get("end")))
    return max(0.0, end - start)


def _range_seconds(short: dict[str, Any]) -> tuple[float, float, float]:
    start = _time_to_seconds(short.get("start_seconds", short.get("start")))
    duration = _duration_seconds(short)
    end = _time_to_seconds(short.get("end_seconds", short.get("end")))
    if duration <= 0 and end > start:
        duration = end - start
    if end <= start and duration > 0:
        end = start + duration
    return start, end, duration


def _source_resolution(package_dir: Path) -> str:
    media = _read_json(package_dir / "media_info.json")
    if isinstance(media, dict):
        width = media.get("width") or media.get("video_width")
        height = media.get("height") or media.get("video_height")
        if width and height:
            return f"{width}x{height}"
    return "unknown"


def _source_duration(package_dir: Path) -> float:
    media = _read_json(package_dir / "media_info.json")
    if isinstance(media, dict):
        try:
            return max(0.0, float(media.get("duration") or 0))
        except Exception:
            return 0.0
    return 0.0


def _fallback_short() -> dict[str, Any]:
    return {
        "id": 0,
        "start": "--",
        "end": "--",
        "duration": "--",
        "reason": "Shorts候補は未生成です。",
        "status": "unavailable",
    }


def ensure_selected_short(package_dir: Path, log: LogCallback | None = None) -> tuple[dict[str, Any], bool]:
    selected_dir = package_dir / "selected"
    selected_dir.mkdir(parents=True, exist_ok=True)
    selected_short_path = selected_dir / "selected_short.json"
    selected_short = _read_json(selected_short_path)
    if isinstance(selected_short, dict):
        return selected_short, False

    shorts = _read_json(package_dir / "shorts_candidates.json")
    selected_short = shorts[0] if isinstance(shorts, list) and shorts and isinstance(shorts[0], dict) else _fallback_short()
    selected_short_path.write_text(json.dumps(selected_short, ensure_ascii=False, indent=2), encoding="utf-8")
    if log:
        log(LOG_TEXT["selected_default_short"])
    return selected_short, True


def find_source_video_path(package_dir: Path) -> Path | None:
    meta = _read_json(package_dir / "package_meta.json")
    if isinstance(meta, dict):
        value = meta.get("source_video_path")
        if value:
            path = Path(str(value))
            if path.exists():
                return path

    log_path = package_dir / "logs" / "package_log.txt"
    if log_path.exists():
        text = log_path.read_text(encoding="utf-8", errors="replace")
        for line in text.splitlines():
            if line.lower().startswith("source_video_path:"):
                value = line.split(":", 1)[1].strip()
                path = Path(value)
                if path.exists():
                    return path

    media = _read_json(package_dir / "media_info.json")
    if isinstance(media, dict):
        value = media.get("source_video_path")
        if value:
            path = Path(str(value))
            if path.exists():
                return path
    return None


def _preview_args(
    ffmpeg_path: str,
    source_video: Path,
    output_path: Path,
    start: float,
    duration: float,
    encoder: str,
) -> list[str]:
    video_codec_args = (
        ["-c:v", "h264_nvenc", "-preset", "p4", "-cq", "23"]
        if encoder == "h264_nvenc"
        else ["-c:v", "libx264", "-preset", "veryfast", "-crf", "23"]
    )
    return [
        ffmpeg_path,
        "-y",
        "-ss",
        f"{max(0.0, start):.3f}",
        "-i",
        str(source_video),
        "-t",
        f"{max(1.0, duration):.3f}",
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


def _write_preview_log(
    selected_dir: Path,
    source_video: Path | None,
    short: dict[str, Any],
    encoder: str,
    nvenc_used: bool,
    fallback: bool,
    output_path: Path,
    error: str,
) -> Path:
    start, end, duration = _range_seconds(short)
    lines = [
        f"executed_at: {datetime.now().isoformat(timespec='seconds')}",
        f"source_video: {source_video or ''}",
        f"start: {start:.3f}",
        f"end: {end:.3f}",
        f"duration: {duration:.3f}",
        f"encoder: {encoder}",
        f"nvenc_used: {str(nvenc_used).lower()}",
        f"fallback: {str(fallback).lower()}",
        f"output_path: {output_path}",
        f"error: {error}",
    ]
    path = selected_dir / "short_preview_log.txt"
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path


def _write_vertical_log(
    selected_dir: Path,
    source_video: Path | None,
    short: dict[str, Any],
    encoder: str,
    nvenc_used: bool,
    fallback: bool,
    output_path: Path,
    error: str,
) -> Path:
    start, end, duration = _range_seconds(short)
    lines = [
        f"executed_at: {datetime.now().isoformat(timespec='seconds')}",
        f"source_video: {source_video or ''}",
        f"start: {start:.3f}",
        f"end: {end:.3f}",
        f"duration: {duration:.3f}",
        f"output_size: {VERTICAL_SIZE}",
        f"crop_pad_policy: {VERTICAL_POLICY}",
        f"encoder: {encoder}",
        f"nvenc_used: {str(nvenc_used).lower()}",
        f"fallback: {str(fallback).lower()}",
        f"output_path: {output_path}",
        f"error: {error}",
    ]
    path = selected_dir / "short_vertical_log.txt"
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path


def _write_horizontal_video_log(
    selected_dir: Path,
    source_video: Path | None,
    source_resolution: str,
    encoder: str,
    nvenc_used: bool,
    fallback: bool,
    output_path: Path,
    error: str,
) -> Path:
    lines = [
        f"executed_at: {datetime.now().isoformat(timespec='seconds')}",
        f"source_video: {source_video or ''}",
        f"source_resolution: {source_resolution}",
        f"output_resolution: {HORIZONTAL_SIZE}",
        f"crop_pad_policy: {HORIZONTAL_POLICY}",
        f"encoder: {encoder}",
        f"nvenc_used: {str(nvenc_used).lower()}",
        f"fallback: {str(fallback).lower()}",
        f"output_path: {output_path}",
        f"error: {error}",
    ]
    path = selected_dir / "horizontal_video_log.txt"
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path


def _vertical_args(
    ffmpeg_path: str,
    source_video: Path,
    output_path: Path,
    start: float,
    duration: float,
    encoder: str,
) -> list[str]:
    video_codec_args = (
        ["-c:v", "h264_nvenc", "-preset", "p4", "-cq", "23"]
        if encoder == "h264_nvenc"
        else ["-c:v", "libx264", "-preset", "veryfast", "-crf", "23"]
    )
    filter_complex = (
        "[0:v]split=2[bg][fg];"
        "[bg]scale=1080:1920:force_original_aspect_ratio=increase,"
        "crop=1080:1920,boxblur=30:1[bgout];"
        "[fg]scale=1080:1920:force_original_aspect_ratio=decrease[fgout];"
        "[bgout][fgout]overlay=(W-w)/2:(H-h)/2,format=yuv420p[v]"
    )
    return [
        ffmpeg_path,
        "-y",
        "-ss",
        f"{max(0.0, start):.3f}",
        "-i",
        str(source_video),
        "-t",
        f"{max(1.0, duration):.3f}",
        "-filter_complex",
        filter_complex,
        "-map",
        "[v]",
        "-map",
        "0:a?",
        *video_codec_args,
        "-c:a",
        "aac",
        "-b:a",
        "192k",
        "-movflags",
        "+faststart",
        str(output_path),
    ]


def _horizontal_video_args(
    ffmpeg_path: str,
    source_video: Path,
    output_path: Path,
    encoder: str,
) -> list[str]:
    video_codec_args = (
        ["-c:v", "h264_nvenc", "-preset", "p4", "-cq", "23"]
        if encoder == "h264_nvenc"
        else ["-c:v", "libx264", "-preset", "veryfast", "-crf", "23"]
    )
    filter_complex = (
        "[0:v]split=2[bg][fg];"
        "[bg]scale=1920:1080:force_original_aspect_ratio=increase,"
        "crop=1920:1080,boxblur=24:1[bgout];"
        "[fg]scale=1920:1080:force_original_aspect_ratio=decrease[fgout];"
        "[bgout][fgout]overlay=(W-w)/2:(H-h)/2,format=yuv420p[v]"
    )
    return [
        ffmpeg_path,
        "-y",
        "-i",
        str(source_video),
        "-filter_complex",
        filter_complex,
        "-map",
        "[v]",
        "-map",
        "0:a?",
        *video_codec_args,
        "-c:a",
        "aac",
        "-b:a",
        "192k",
        "-movflags",
        "+faststart",
        str(output_path),
    ]


def generate_selected_short_preview(
    package_dir: Path,
    ffmpeg_path: str | None,
    nvenc_online: bool,
    source_video_path: Path | None = None,
    log: LogCallback | None = None,
) -> dict[str, Any]:
    selected_dir = package_dir / "selected"
    selected_dir.mkdir(parents=True, exist_ok=True)
    output_path = selected_dir / "short_preview.mp4"

    def emit(message: str) -> None:
        if log:
            log(message)

    emit(LOG_TEXT["selected_preview_start"])
    selected_short, used_default = ensure_selected_short(package_dir, log=log)
    start, end, duration = _range_seconds(selected_short)
    source_video = source_video_path or find_source_video_path(package_dir)

    if source_video is None or not source_video.exists():
        emit(LOG_TEXT["selected_preview_source_missing"])
        _write_preview_log(selected_dir, source_video, selected_short, "unavailable", False, False, output_path, "source video missing")
        return {
            "status": "FAILED",
            "package_dir": str(package_dir),
            "selected_dir": str(selected_dir),
            "output_path": str(output_path),
            "encoder": "unavailable",
            "nvenc_used": False,
            "fallback": False,
            "used_default": used_default,
            "message": "Source video is required.",
        }
    if not ffmpeg_path:
        emit(LOG_TEXT["selected_preview_ffmpeg_missing"])
        _write_preview_log(selected_dir, source_video, selected_short, "unavailable", False, False, output_path, "ffmpeg missing")
        return {
            "status": "FAILED",
            "package_dir": str(package_dir),
            "selected_dir": str(selected_dir),
            "output_path": str(output_path),
            "encoder": "unavailable",
            "nvenc_used": False,
            "fallback": False,
            "used_default": used_default,
            "message": "FFmpeg is required for selected short preview.",
        }
    if duration <= 0:
        emit(LOG_TEXT["selected_preview_failed"])
        _write_preview_log(selected_dir, source_video, selected_short, "unavailable", False, False, output_path, "invalid selected short range")
        return {
            "status": "FAILED",
            "package_dir": str(package_dir),
            "selected_dir": str(selected_dir),
            "output_path": str(output_path),
            "encoder": "unavailable",
            "nvenc_used": False,
            "fallback": False,
            "used_default": used_default,
            "message": "Selected short range is unavailable.",
        }

    if output_path.exists():
        output_path.unlink()

    errors: list[str] = []
    fallback = False
    if nvenc_online:
        emit(LOG_TEXT["selected_preview_nvenc"])
        completed = run_command(_preview_args(ffmpeg_path, source_video, output_path, start, duration, "h264_nvenc"), timeout=240)
        if completed.returncode == 0 and output_path.exists():
            emit(LOG_TEXT["selected_preview_created"])
            _write_preview_log(selected_dir, source_video, selected_short, "h264_nvenc", True, False, output_path, "")
            return {
                "status": "COMPLETED",
                "package_dir": str(package_dir),
                "selected_dir": str(selected_dir),
                "output_path": str(output_path),
                "encoder": "h264_nvenc",
                "nvenc_used": True,
                "fallback": False,
                "used_default": used_default,
                "message": LOG_TEXT["selected_preview_created"],
            }
        errors.append((completed.stderr or completed.stdout or "h264_nvenc failed").strip()[-800:])
        fallback = True
        emit(LOG_TEXT["selected_preview_fallback"])

    completed = run_command(_preview_args(ffmpeg_path, source_video, output_path, start, duration, "libx264"), timeout=240)
    if completed.returncode == 0 and output_path.exists():
        emit(LOG_TEXT["selected_preview_created"])
        _write_preview_log(selected_dir, source_video, selected_short, "libx264", False, fallback, output_path, "")
        return {
            "status": "COMPLETED",
            "package_dir": str(package_dir),
            "selected_dir": str(selected_dir),
            "output_path": str(output_path),
            "encoder": "libx264",
            "nvenc_used": False,
            "fallback": fallback,
            "used_default": used_default,
            "message": LOG_TEXT["selected_preview_created"],
        }

    errors.append((completed.stderr or completed.stdout or "libx264 failed").strip()[-800:])
    emit(LOG_TEXT["selected_preview_failed"])
    error_text = "\n".join(error for error in errors if error)
    _write_preview_log(selected_dir, source_video, selected_short, "unavailable", False, fallback, output_path, error_text)
    return {
        "status": "FAILED",
        "package_dir": str(package_dir),
        "selected_dir": str(selected_dir),
        "output_path": str(output_path),
        "encoder": "unavailable",
        "nvenc_used": False,
        "fallback": fallback,
        "used_default": used_default,
        "message": error_text or "Short preview generation failed.",
    }


def generate_vertical_short(
    package_dir: Path,
    ffmpeg_path: str | None,
    nvenc_online: bool,
    source_video_path: Path | None = None,
    log: LogCallback | None = None,
) -> dict[str, Any]:
    selected_dir = package_dir / "selected"
    selected_dir.mkdir(parents=True, exist_ok=True)
    output_path = selected_dir / "short_vertical_1080x1920.mp4"

    def emit(message: str) -> None:
        if log:
            log(message)

    emit(LOG_TEXT["vertical_short_start"])
    selected_short, used_default = ensure_selected_short(package_dir, log=log)
    start, end, duration = _range_seconds(selected_short)
    source_video = source_video_path or find_source_video_path(package_dir)

    if source_video is None or not source_video.exists():
        emit(LOG_TEXT["selected_preview_source_missing"])
        _write_vertical_log(selected_dir, source_video, selected_short, "unavailable", False, False, output_path, "source video missing")
        return {
            "status": "FAILED",
            "package_dir": str(package_dir),
            "selected_dir": str(selected_dir),
            "output_path": str(output_path),
            "size": VERTICAL_SIZE,
            "encoder": "unavailable",
            "nvenc_used": False,
            "fallback": False,
            "used_default": used_default,
            "message": "Source video is required.",
        }
    if not ffmpeg_path:
        emit(LOG_TEXT["selected_preview_ffmpeg_missing"])
        _write_vertical_log(selected_dir, source_video, selected_short, "unavailable", False, False, output_path, "ffmpeg missing")
        return {
            "status": "FAILED",
            "package_dir": str(package_dir),
            "selected_dir": str(selected_dir),
            "output_path": str(output_path),
            "size": VERTICAL_SIZE,
            "encoder": "unavailable",
            "nvenc_used": False,
            "fallback": False,
            "used_default": used_default,
            "message": "FFmpeg is required for vertical short export.",
        }
    if duration <= 0:
        emit(LOG_TEXT["vertical_short_failed"])
        _write_vertical_log(selected_dir, source_video, selected_short, "unavailable", False, False, output_path, "invalid selected short range")
        return {
            "status": "FAILED",
            "package_dir": str(package_dir),
            "selected_dir": str(selected_dir),
            "output_path": str(output_path),
            "size": VERTICAL_SIZE,
            "encoder": "unavailable",
            "nvenc_used": False,
            "fallback": False,
            "used_default": used_default,
            "message": "Selected short range is unavailable.",
        }

    emit(LOG_TEXT["vertical_short_layout"])
    if output_path.exists():
        output_path.unlink()

    errors: list[str] = []
    fallback = False
    if nvenc_online:
        completed = run_command(_vertical_args(ffmpeg_path, source_video, output_path, start, duration, "h264_nvenc"), timeout=360)
        if completed.returncode == 0 and output_path.exists():
            emit(LOG_TEXT["vertical_short_nvenc"])
            emit(LOG_TEXT["vertical_short_created"])
            _write_vertical_log(selected_dir, source_video, selected_short, "h264_nvenc", True, False, output_path, "")
            return {
                "status": "COMPLETED",
                "package_dir": str(package_dir),
                "selected_dir": str(selected_dir),
                "output_path": str(output_path),
                "size": VERTICAL_SIZE,
                "encoder": "h264_nvenc",
                "nvenc_used": True,
                "fallback": False,
                "used_default": used_default,
                "message": LOG_TEXT["vertical_short_created"],
            }
        errors.append((completed.stderr or completed.stdout or "h264_nvenc failed").strip()[-1000:])
        fallback = True
        emit(LOG_TEXT["vertical_short_fallback"])

    completed = run_command(_vertical_args(ffmpeg_path, source_video, output_path, start, duration, "libx264"), timeout=360)
    if completed.returncode == 0 and output_path.exists():
        emit(LOG_TEXT["vertical_short_created"])
        _write_vertical_log(selected_dir, source_video, selected_short, "libx264", False, fallback, output_path, "")
        return {
            "status": "COMPLETED",
            "package_dir": str(package_dir),
            "selected_dir": str(selected_dir),
            "output_path": str(output_path),
            "size": VERTICAL_SIZE,
            "encoder": "libx264",
            "nvenc_used": False,
            "fallback": fallback,
            "used_default": used_default,
            "message": LOG_TEXT["vertical_short_created"],
        }

    errors.append((completed.stderr or completed.stdout or "libx264 failed").strip()[-1000:])
    emit(LOG_TEXT["vertical_short_failed"])
    error_text = "\n".join(error for error in errors if error)
    _write_vertical_log(selected_dir, source_video, selected_short, "unavailable", False, fallback, output_path, error_text)
    return {
        "status": "FAILED",
        "package_dir": str(package_dir),
        "selected_dir": str(selected_dir),
        "output_path": str(output_path),
        "size": VERTICAL_SIZE,
        "encoder": "unavailable",
        "nvenc_used": False,
        "fallback": fallback,
        "used_default": used_default,
        "message": error_text or "Vertical short export failed.",
    }


def generate_horizontal_video(
    package_dir: Path,
    ffmpeg_path: str | None,
    nvenc_online: bool,
    source_video_path: Path | None = None,
    log: LogCallback | None = None,
) -> dict[str, Any]:
    selected_dir = package_dir / "selected"
    selected_dir.mkdir(parents=True, exist_ok=True)
    output_path = selected_dir / "horizontal_video.mp4"
    source_resolution = _source_resolution(package_dir)

    def emit(message: str) -> None:
        if log:
            log(message)

    emit(LOG_TEXT["horizontal_video_start"])
    source_video = source_video_path or find_source_video_path(package_dir)

    if source_video is None or not source_video.exists():
        emit(LOG_TEXT["selected_preview_source_missing"])
        _write_horizontal_video_log(
            selected_dir,
            source_video,
            source_resolution,
            "unavailable",
            False,
            False,
            output_path,
            "source video missing",
        )
        return {
            "status": "FAILED",
            "package_dir": str(package_dir),
            "selected_dir": str(selected_dir),
            "output_path": str(output_path),
            "size": HORIZONTAL_SIZE,
            "encoder": "unavailable",
            "nvenc_used": False,
            "fallback": False,
            "source_resolution": source_resolution,
            "message": "Source video is required.",
        }
    if not ffmpeg_path:
        emit(LOG_TEXT["selected_preview_ffmpeg_missing"])
        _write_horizontal_video_log(
            selected_dir,
            source_video,
            source_resolution,
            "unavailable",
            False,
            False,
            output_path,
            "ffmpeg missing",
        )
        return {
            "status": "FAILED",
            "package_dir": str(package_dir),
            "selected_dir": str(selected_dir),
            "output_path": str(output_path),
            "size": HORIZONTAL_SIZE,
            "encoder": "unavailable",
            "nvenc_used": False,
            "fallback": False,
            "source_resolution": source_resolution,
            "message": "FFmpeg is required for horizontal video export.",
        }

    emit(LOG_TEXT["horizontal_video_layout"])
    if output_path.exists():
        output_path.unlink()

    duration = _source_duration(package_dir)
    timeout = max(600, int(duration * 3 + 180)) if duration > 0 else 900
    errors: list[str] = []
    fallback = False
    if nvenc_online:
        completed = run_command(_horizontal_video_args(ffmpeg_path, source_video, output_path, "h264_nvenc"), timeout=timeout)
        if completed.returncode == 0 and output_path.exists():
            emit(LOG_TEXT["horizontal_video_nvenc"])
            emit(LOG_TEXT["horizontal_video_created"])
            _write_horizontal_video_log(selected_dir, source_video, source_resolution, "h264_nvenc", True, False, output_path, "")
            return {
                "status": "COMPLETED",
                "package_dir": str(package_dir),
                "selected_dir": str(selected_dir),
                "output_path": str(output_path),
                "size": HORIZONTAL_SIZE,
                "encoder": "h264_nvenc",
                "nvenc_used": True,
                "fallback": False,
                "source_resolution": source_resolution,
                "message": LOG_TEXT["horizontal_video_created"],
            }
        errors.append((completed.stderr or completed.stdout or "h264_nvenc failed").strip()[-1000:])
        fallback = True
        emit(LOG_TEXT["horizontal_video_fallback"])

    completed = run_command(_horizontal_video_args(ffmpeg_path, source_video, output_path, "libx264"), timeout=timeout)
    if completed.returncode == 0 and output_path.exists():
        emit(LOG_TEXT["horizontal_video_created"])
        _write_horizontal_video_log(selected_dir, source_video, source_resolution, "libx264", False, fallback, output_path, "")
        return {
            "status": "COMPLETED",
            "package_dir": str(package_dir),
            "selected_dir": str(selected_dir),
            "output_path": str(output_path),
            "size": HORIZONTAL_SIZE,
            "encoder": "libx264",
            "nvenc_used": False,
            "fallback": fallback,
            "source_resolution": source_resolution,
            "message": LOG_TEXT["horizontal_video_created"],
        }

    errors.append((completed.stderr or completed.stdout or "libx264 failed").strip()[-1000:])
    emit(LOG_TEXT["horizontal_video_failed"])
    error_text = "\n".join(error for error in errors if error)
    _write_horizontal_video_log(
        selected_dir,
        source_video,
        source_resolution,
        "unavailable",
        False,
        fallback,
        output_path,
        error_text,
    )
    return {
        "status": "FAILED",
        "package_dir": str(package_dir),
        "selected_dir": str(selected_dir),
        "output_path": str(output_path),
        "size": HORIZONTAL_SIZE,
        "encoder": "unavailable",
        "nvenc_used": False,
        "fallback": fallback,
        "source_resolution": source_resolution,
        "message": error_text or "Horizontal video export failed.",
    }
