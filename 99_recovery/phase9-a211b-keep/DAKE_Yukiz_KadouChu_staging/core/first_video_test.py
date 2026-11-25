from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Callable, Any

from core.app_config import LOG_TEXT, human_size, outputs_dir, seconds_to_timecode
from core.cli_checker import run_command
from core.media_probe import MediaInfo, probe_media

LogCallback = Callable[[str], None]


def first_video_test_dir() -> Path:
    return outputs_dir() / "first_video_test"


def _media_info_payload(info: MediaInfo) -> dict[str, Any]:
    return {
        "file_name": info.file_name,
        "file_size_bytes": info.file_size_bytes,
        "file_size": human_size(info.file_size_bytes),
        "duration": info.duration,
        "duration_timecode": seconds_to_timecode(info.duration),
        "width": info.width,
        "height": info.height,
        "fps": info.fps,
        "video_codec": info.video_codec,
        "audio_present": info.audio_present,
        "audio_codec": info.audio_codec,
    }


def _build_clip_args(ffmpeg_path: str, video_path: Path, output_path: Path, encoder: str) -> list[str]:
    video_args = (
        ["-c:v", "h264_nvenc", "-preset", "p4", "-cq", "23"]
        if encoder == "h264_nvenc"
        else ["-c:v", "libx264", "-preset", "veryfast", "-crf", "23"]
    )
    return [
        ffmpeg_path,
        "-y",
        "-ss",
        "0",
        "-i",
        str(video_path),
        "-t",
        "10",
        "-map",
        "0:v:0",
        "-map",
        "0:a?",
        *video_args,
        "-c:a",
        "aac",
        "-b:a",
        "128k",
        "-movflags",
        "+faststart",
        str(output_path),
    ]


def _try_encode(ffmpeg_path: str, video_path: Path, output_path: Path, encoder: str) -> tuple[bool, str]:
    completed = run_command(_build_clip_args(ffmpeg_path, video_path, output_path, encoder), timeout=180)
    message = (completed.stderr or completed.stdout or "").strip()
    return completed.returncode == 0 and output_path.exists(), message[-1200:]


def _write_log(log_path: Path, entries: list[str], payload: dict[str, Any]) -> None:
    lines = [*entries, "", "RESULT", json.dumps(payload, ensure_ascii=False, indent=2)]
    log_path.write_text("\n".join(lines).strip() + "\n", encoding="utf-8")


def run_first_video_test(
    video_path: Path,
    ffprobe_path: str | None,
    ffmpeg_path: str | None,
    nvenc_online: bool,
    log: LogCallback | None = None,
) -> dict[str, Any]:
    output_dir = first_video_test_dir()
    logs_dir = output_dir / "logs"
    output_dir.mkdir(parents=True, exist_ok=True)
    logs_dir.mkdir(parents=True, exist_ok=True)

    log_entries: list[str] = []

    def emit(message: str) -> None:
        line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} {message}"
        log_entries.append(line)
        if log is not None:
            log(message)

    result: dict[str, Any] = {
        "started_at": datetime.now().isoformat(timespec="seconds"),
        "selected_file": str(video_path),
        "output_dir": str(output_dir),
        "media_info_path": str(output_dir / "media_info.json"),
        "test_clip_path": str(output_dir / "test_clip.mp4"),
        "log_path": str(logs_dir / "test_log.txt"),
        "media_info": None,
        "test_clip": "SKIPPED",
        "nvenc": "SKIPPED",
        "encoder": "none",
        "fallback": False,
        "message": "",
    }

    emit(LOG_TEXT["first_test_start"])

    media_info: MediaInfo | None = None
    media_info_path = output_dir / "media_info.json"
    if ffprobe_path:
        try:
            media_info = probe_media(video_path, ffprobe_path)
            result["media_info"] = _media_info_payload(media_info)
            media_info_path.write_text(json.dumps(result["media_info"], ensure_ascii=False, indent=2), encoding="utf-8")
            emit(LOG_TEXT["first_test_probe_ready"])
        except Exception as exc:
            result["message"] = f"ffprobe failed: {exc}"
            media_info_path.write_text(
                json.dumps({"available": False, "reason": str(exc)}, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            emit(LOG_TEXT["first_test_probe_failed"])
    else:
        result["message"] = "FFprobe is required for media information."
        media_info_path.write_text(
            json.dumps({"available": False, "reason": result["message"]}, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        emit(LOG_TEXT["first_test_ffprobe_skipped"])

    output_clip = output_dir / "test_clip.mp4"
    if output_clip.exists():
        try:
            output_clip.unlink()
        except OSError:
            pass

    if not ffmpeg_path:
        result["test_clip"] = "SKIPPED"
        result["nvenc"] = "SKIPPED"
        result["message"] = "FFmpeg is required for first video test."
        emit(LOG_TEXT["first_test_ffmpeg_required"])
        _write_log(logs_dir / "test_log.txt", log_entries, result)
        return result

    nvenc_error = ""
    if nvenc_online:
        emit(LOG_TEXT["first_test_nvenc_try"])
        created, nvenc_error = _try_encode(ffmpeg_path, video_path, output_clip, "h264_nvenc")
        if created:
            result["test_clip"] = "READY"
            result["nvenc"] = "READY"
            result["encoder"] = "h264_nvenc"
            emit(LOG_TEXT["first_test_nvenc_ready"])
            _write_log(logs_dir / "test_log.txt", log_entries, result)
            return result
        result["fallback"] = True
        result["nvenc"] = "FALLBACK CPU"
        result["nvenc_error"] = nvenc_error
        emit(LOG_TEXT["first_test_nvenc_fallback"])
    else:
        result["nvenc"] = "SKIPPED"
        emit(LOG_TEXT["first_test_cpu_try"])

    created, cpu_message = _try_encode(ffmpeg_path, video_path, output_clip, "libx264")
    if created:
        result["test_clip"] = "READY"
        result["encoder"] = "libx264"
        result["cpu_message"] = cpu_message
        emit(LOG_TEXT["first_test_clip_ready"])
    else:
        result["test_clip"] = "FAILED"
        result["encoder"] = "none"
        result["cpu_error"] = cpu_message
        if nvenc_error:
            result["message"] = f"NVENC failed, CPU failed. NVENC: {nvenc_error[:300]} CPU: {cpu_message[:300]}"
        else:
            result["message"] = f"CPU encode failed. {cpu_message[:600]}"
        emit(LOG_TEXT["first_test_clip_failed"])

    _write_log(logs_dir / "test_log.txt", log_entries, result)
    return result
