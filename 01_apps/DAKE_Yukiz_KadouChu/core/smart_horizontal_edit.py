from __future__ import annotations

import json
import re
import shutil
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

from core.app_config import LOG_TEXT, app_root, seconds_to_timecode
from core.cli_checker import run_command
from core.ollama_client import generate_ollama_text
from core.selected_preview import find_source_video_path

LogCallback = Callable[[str], None]

OUTPUT_SIZE = "1920x1080"
OUTPUT_POLICY = "quiet segments + background blur + foreground centered"
OUTPUT_NAME = "smart_horizontal_edit.mp4"
SEQUENCE_NAME = "smart_horizontal_sequence.json"
LOG_NAME = "smart_horizontal_edit_log.txt"
BGM_EXTENSIONS = {".mp3", ".wav", ".m4a", ".aac", ".flac", ".ogg"}
BGM_VOLUME = 0.07
FADE_IN_SECONDS = 0.55
FADE_OUT_SECONDS = 0.65
MIN_SEGMENTS = 3
MAX_SEGMENTS = 10
MAX_SEGMENT_SECONDS = 90.0


def _read_json(path: Path) -> Any:
    if not path.exists() or not path.is_file():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8", errors="replace"))
    except Exception:
        return None


def _read_text(path: Path, limit: int = 1600) -> str:
    if not path.exists() or not path.is_file():
        return ""
    try:
        return path.read_text(encoding="utf-8", errors="replace").strip()[:limit]
    except Exception:
        return ""


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


def _duration_seconds(item: dict[str, Any]) -> float:
    if item.get("duration_seconds") is not None:
        try:
            return max(0.0, float(item.get("duration_seconds") or 0))
        except Exception:
            return 0.0
    duration = _time_to_seconds(item.get("duration"))
    if duration > 0:
        return duration
    start = _time_to_seconds(item.get("start_seconds", item.get("start")))
    end = _time_to_seconds(item.get("end_seconds", item.get("end")))
    return max(0.0, end - start)


def _range_seconds(item: dict[str, Any]) -> tuple[float, float, float]:
    start = _time_to_seconds(item.get("start_seconds", item.get("start")))
    duration = _duration_seconds(item)
    end = _time_to_seconds(item.get("end_seconds", item.get("end")))
    if duration <= 0 and end > start:
        duration = end - start
    if end <= start and duration > 0:
        end = start + duration
    if duration > MAX_SEGMENT_SECONDS:
        duration = MAX_SEGMENT_SECONDS
        end = start + duration
    return start, end, duration


def _media_duration(package_dir: Path) -> float:
    media = _read_json(package_dir / "media_info.json")
    if isinstance(media, dict):
        try:
            return max(0.0, float(media.get("duration") or 0))
        except Exception:
            return 0.0
    return 0.0


def _source_audio_present(package_dir: Path) -> bool:
    media = _read_json(package_dir / "media_info.json")
    if isinstance(media, dict) and "audio_present" in media:
        return bool(media.get("audio_present"))
    return True


def _selected_bgm(package_dir: Path) -> Path | None:
    bgm_dir = package_dir / "selected" / "bgm"
    if not bgm_dir.exists():
        return None
    try:
        files = sorted(path for path in bgm_dir.iterdir() if path.is_file() and path.suffix.lower() in BGM_EXTENSIONS)
    except Exception:
        return None
    return files[0] if files else None


def _normal_segment(item: dict[str, Any], fallback_type: str, fallback_reason: str) -> dict[str, Any] | None:
    start, end, duration = _range_seconds(item)
    if duration < 1.0:
        return None
    return {
        "start": seconds_to_timecode(start),
        "end": seconds_to_timecode(end),
        "duration": seconds_to_timecode(duration),
        "start_seconds": round(start, 3),
        "end_seconds": round(end, 3),
        "duration_seconds": round(duration, 3),
        "type": str(item.get("type") or fallback_type or "WORK").upper(),
        "reason": str(item.get("reason") or fallback_reason),
        "status": "candidate",
    }


def _dedupe_segments(segments: list[dict[str, Any]]) -> list[dict[str, Any]]:
    seen: set[tuple[int, int]] = set()
    clean: list[dict[str, Any]] = []
    for item in sorted(segments, key=lambda segment: float(segment.get("start_seconds") or 0)):
        key = (round(float(item.get("start_seconds") or 0)), round(float(item.get("end_seconds") or 0)))
        if key in seen:
            continue
        seen.add(key)
        clean.append(item)
    return clean[:MAX_SEGMENTS]


def _read_selected_short(package_dir: Path) -> list[dict[str, Any]]:
    selected = _read_json(package_dir / "selected" / "selected_short.json")
    if isinstance(selected, dict):
        segment = _normal_segment(selected, "SELECTED", "選択済みShorts候補")
        return [segment] if segment else []
    return []


def _read_pack_segments(package_dir: Path) -> list[dict[str, Any]]:
    payload = _read_json(package_dir / "selected" / "shorts_pack" / "shorts_pack.json")
    clips = payload.get("clips") if isinstance(payload, dict) else []
    if not isinstance(clips, list):
        return []
    segments: list[dict[str, Any]] = []
    for clip in clips:
        if not isinstance(clip, dict):
            continue
        segment = _normal_segment(clip, str(clip.get("type") or "WORK"), str(clip.get("reason") or "Shorts Pack候補"))
        if segment:
            segments.append(segment)
    return segments


def _read_candidate_segments(package_dir: Path) -> list[dict[str, Any]]:
    payload = _read_json(package_dir / "shorts_candidates.json")
    if not isinstance(payload, list):
        return []
    segments: list[dict[str, Any]] = []
    role_cycle = ["INTRO", "WORK", "AFTERGLOW", "WORK"]
    for index, item in enumerate(payload):
        if not isinstance(item, dict):
            continue
        segment = _normal_segment(item, role_cycle[index % len(role_cycle)], str(item.get("reason") or "Shorts候補"))
        if segment:
            segments.append(segment)
    return segments


def _fallback_segments(duration: float) -> list[dict[str, Any]]:
    if duration <= 0:
        duration = 180.0
    count = max(MIN_SEGMENTS, min(5, int(max(3, duration // 90))))
    clip_floor = 6.0 if duration >= 30 else 2.0
    clip_length = min(60.0, max(clip_floor, min(45.0, duration / (count + 1))))
    span = max(0.0, duration - clip_length)
    roles = ["INTRO", "WORK", "AFTERGLOW", "WORK", "AFTERGLOW"]
    reasons = ["静かな導入", "作業感", "余熱", "流れの継続", "静かな終わり"]
    segments: list[dict[str, Any]] = []
    for index in range(count):
        ratio = 0.0 if count == 1 else index / max(1, count - 1)
        start = span * ratio
        end = min(duration, start + clip_length)
        segments.append(
            {
                "start": seconds_to_timecode(start),
                "end": seconds_to_timecode(end),
                "duration": seconds_to_timecode(max(1.0, end - start)),
                "start_seconds": round(start, 3),
                "end_seconds": round(end, 3),
                "duration_seconds": round(max(1.0, end - start), 3),
                "type": roles[index % len(roles)],
                "reason": reasons[index % len(reasons)],
                "status": "candidate",
            }
        )
    return segments


def _build_template_sequence(package_dir: Path) -> list[dict[str, Any]]:
    segments = [
        *_read_selected_short(package_dir),
        *_read_pack_segments(package_dir),
        *_read_candidate_segments(package_dir),
    ]
    segments = _dedupe_segments(segments)
    if len(segments) < MIN_SEGMENTS:
        segments = _dedupe_segments([*segments, *_fallback_segments(_media_duration(package_dir))])
    return segments[:MAX_SEGMENTS]


def _ollama_sequence(package_dir: Path, segments: list[dict[str, Any]], ollama_ready: bool) -> tuple[list[dict[str, Any]], bool, str]:
    if not ollama_ready or not segments:
        return segments, False, ""
    memory_summary = app_root() / "data" / "memory" / "memory_summary.md"
    context = {
        "target": "5m-15m quiet horizontal edit",
        "segments": segments[:MAX_SEGMENTS],
        "assistant_review": _read_text(package_dir / "assistant_review.md", limit=1000),
        "assistant_recommendation": _read_text(package_dir / "assistant_recommendation.md", limit=1000),
        "memory_summary": _read_text(memory_summary, limit=800),
        "bridge_metadata": _read_text(package_dir / "selected" / "upload" / "metadata_draft.txt", limit=800),
        "bgm": [path.name for path in (package_dir / "selected" / "bgm").glob("*") if path.is_file()] if (package_dir / "selected" / "bgm").exists() else [],
    }
    prompt = (
        "You are the local assistant brain for Dakeユキズ稼働中.\n"
        "Choose 3 to 10 quiet segments for a horizontal edit. This is not timeline editing.\n"
        "Keep chronological flow. Avoid hype and flashy edits.\n"
        "Return only JSON array with keys: start_seconds, end_seconds, type, reason.\n\n"
        f"{json.dumps(context, ensure_ascii=False)}"
    )
    response = generate_ollama_text(prompt, timeout=45)
    if not response.get("ok"):
        return segments, False, str(response.get("reason") or "")
    text = str(response.get("text") or "")
    start = text.find("[")
    end = text.rfind("]")
    if start == -1 or end == -1 or end <= start:
        return segments, False, "Ollama response did not include a JSON array."
    try:
        parsed = json.loads(text[start : end + 1])
    except Exception as exc:
        return segments, False, str(exc)
    if not isinstance(parsed, list):
        return segments, False, "Ollama response was not a list."
    updated: list[dict[str, Any]] = []
    for index, item in enumerate(parsed[:MAX_SEGMENTS]):
        if not isinstance(item, dict):
            continue
        fallback = segments[min(index, len(segments) - 1)] if segments else {}
        merged = dict(fallback)
        merged.update(item)
        segment = _normal_segment(merged, str(fallback.get("type") or "WORK"), str(fallback.get("reason") or "静かな流れ"))
        if segment:
            updated.append(segment)
    updated = _dedupe_segments(updated)
    if len(updated) < MIN_SEGMENTS:
        return segments, False, "Ollama returned too few valid segments."
    return updated, True, str(response.get("model") or "")


def _video_filter(duration: float) -> str:
    fade_in = min(FADE_IN_SECONDS, max(0.2, duration * 0.16))
    fade_out = min(FADE_OUT_SECONDS, max(0.2, duration * 0.18))
    fade_out_start = max(0.0, duration - fade_out)
    return (
        "[0:v]split=2[bg][fg];"
        "[bg]scale=1920:1080:force_original_aspect_ratio=increase,"
        "crop=1920:1080,boxblur=24:1[bgout];"
        "[fg]scale=1920:1080:force_original_aspect_ratio=decrease[fgout];"
        "[bgout][fgout]overlay=(W-w)/2:(H-h)/2,format=yuv420p,"
        f"fade=t=in:st=0:d={fade_in:.3f},"
        f"fade=t=out:st={fade_out_start:.3f}:d={fade_out:.3f}[v]"
    )


def _part_args(ffmpeg_path: str, source_video: Path, output_path: Path, segment: dict[str, Any], source_audio_present: bool) -> list[str]:
    start = float(segment.get("start_seconds") or 0)
    duration = max(1.0, float(segment.get("duration_seconds") or 30))
    args = [ffmpeg_path, "-y", "-ss", f"{start:.3f}", "-i", str(source_video)]
    if not source_audio_present:
        args.extend(["-f", "lavfi", "-t", f"{duration:.3f}", "-i", "anullsrc=channel_layout=stereo:sample_rate=48000"])
    args.extend(["-t", f"{duration:.3f}", "-filter_complex", _video_filter(duration), "-map", "[v]"])
    if source_audio_present:
        args.extend(["-map", "0:a:0?"])
    else:
        args.extend(["-map", "1:a:0"])
    args.extend(
        [
            "-r",
            "30",
            "-c:v",
            "libx264",
            "-preset",
            "veryfast",
            "-crf",
            "20",
            "-c:a",
            "aac",
            "-b:a",
            "192k",
            "-ar",
            "48000",
            "-ac",
            "2",
            "-shortest",
            str(output_path),
        ]
    )
    return args


def _codec_args(encoder: str) -> list[str]:
    if encoder == "h264_nvenc":
        return ["-c:v", "h264_nvenc", "-preset", "p4", "-cq", "23"]
    return ["-c:v", "libx264", "-preset", "veryfast", "-crf", "22"]


def _concat_escape(path: Path) -> str:
    return str(path.resolve()).replace("\\", "/").replace("'", "'\\''")


def _write_concat_list(temp_root: Path, part_paths: list[Path]) -> Path:
    path = temp_root / "smart_horizontal_concat.txt"
    path.write_text("\n".join(f"file '{_concat_escape(part)}'" for part in part_paths) + "\n", encoding="utf-8")
    return path


def _concat_args(ffmpeg_path: str, concat_list: Path, output_path: Path, encoder: str) -> list[str]:
    return [
        ffmpeg_path,
        "-y",
        "-f",
        "concat",
        "-safe",
        "0",
        "-i",
        str(concat_list),
        *_codec_args(encoder),
        "-c:a",
        "aac",
        "-b:a",
        "192k",
        "-movflags",
        "+faststart",
        str(output_path),
    ]


def _mix_bgm_args(ffmpeg_path: str, base_path: Path, bgm_path: Path, output_path: Path, duration: float) -> list[str]:
    fade_out_start = max(0.0, duration - FADE_OUT_SECONDS)
    return [
        ffmpeg_path,
        "-y",
        "-i",
        str(base_path),
        "-stream_loop",
        "-1",
        "-i",
        str(bgm_path),
        "-t",
        f"{max(1.0, duration):.3f}",
        "-filter_complex",
        (
            "[0:a]volume=1.0[a0];"
            f"[1:a]volume={BGM_VOLUME:.3f},atrim=0:{max(1.0, duration):.3f},"
            f"afade=t=in:st=0:d=0.400,afade=t=out:st={fade_out_start:.3f}:d={FADE_OUT_SECONDS:.3f}[a1];"
            "[a0][a1]amix=inputs=2:duration=first:dropout_transition=0[a]"
        ),
        "-map",
        "0:v:0",
        "-map",
        "[a]",
        "-c:v",
        "copy",
        "-c:a",
        "aac",
        "-b:a",
        "192k",
        "-movflags",
        "+faststart",
        str(output_path),
    ]


def _write_sequence(selected_dir: Path, segments: list[dict[str, Any]]) -> Path:
    path = selected_dir / SEQUENCE_NAME
    path.write_text(json.dumps(segments, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def _write_log(selected_dir: Path, lines: list[str]) -> Path:
    path = selected_dir / LOG_NAME
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path


def generate_smart_horizontal_edit(
    package_dir: Path,
    ffmpeg_path: str | None,
    nvenc_online: bool,
    ollama_ready: bool,
    source_video_path: Path | None = None,
    log: LogCallback | None = None,
) -> dict[str, Any]:
    selected_dir = package_dir / "selected"
    selected_dir.mkdir(parents=True, exist_ok=True)
    output_path = selected_dir / OUTPUT_NAME

    def emit(message: str) -> None:
        if log:
            log(message)

    emit(LOG_TEXT["smart_horizontal_start"])
    source_video = source_video_path or find_source_video_path(package_dir)
    bgm_path = _selected_bgm(package_dir)
    segments = _build_template_sequence(package_dir)
    segments, used_ollama, ollama_note = _ollama_sequence(package_dir, segments, ollama_ready)
    sequence_path = _write_sequence(selected_dir, segments)
    total_duration = sum(float(segment.get("duration_seconds") or 0) for segment in segments)
    if bgm_path:
        emit(LOG_TEXT["smart_horizontal_bgm"])

    log_lines = [
        f"executed_at: {datetime.now().isoformat(timespec='seconds')}",
        f"package_dir: {package_dir}",
        f"source_video: {source_video or ''}",
        f"sequence_path: {sequence_path}",
        f"segment_count: {len(segments)}",
        f"total_duration: {round(total_duration, 3)}",
        f"output_size: {OUTPUT_SIZE}",
        f"policy: {OUTPUT_POLICY}",
        f"fade_used: true",
        f"fade_in_seconds: {FADE_IN_SECONDS}",
        f"fade_out_seconds: {FADE_OUT_SECONDS}",
        f"bgm: {bgm_path or ''}",
        f"bgm_volume: {BGM_VOLUME if bgm_path else 0}",
        f"used_ollama: {str(used_ollama).lower()}",
        f"ollama_note: {ollama_note}",
    ]

    if len(segments) < MIN_SEGMENTS:
        emit(LOG_TEXT["smart_horizontal_failed"])
        log_path = _write_log(selected_dir, [*log_lines, "status: FAILED", "error: no usable segments"])
        return _result(package_dir, selected_dir, output_path, sequence_path, log_path, segments, "FAILED", "unavailable", False, False, False, used_ollama, "No usable segments.")
    if source_video is None or not source_video.exists():
        emit(LOG_TEXT["selected_preview_source_missing"])
        log_path = _write_log(selected_dir, [*log_lines, "status: FAILED", "error: source video missing"])
        return _result(package_dir, selected_dir, output_path, sequence_path, log_path, segments, "FAILED", "unavailable", False, False, False, used_ollama, "Source video is required.")
    if not ffmpeg_path:
        emit(LOG_TEXT["selected_preview_ffmpeg_missing"])
        log_path = _write_log(selected_dir, [*log_lines, "status: FAILED", "error: ffmpeg missing"])
        return _result(package_dir, selected_dir, output_path, sequence_path, log_path, segments, "FAILED", "unavailable", False, False, False, used_ollama, "FFmpeg is required for Smart Horizontal Edit.")

    emit(LOG_TEXT["smart_horizontal_flow"])
    emit(LOG_TEXT["smart_horizontal_export"])
    errors: list[str] = []
    source_audio_present = _source_audio_present(package_dir)
    fallback = False
    bgm_used = False
    encoder = "h264_nvenc" if nvenc_online else "libx264"

    with tempfile.TemporaryDirectory(prefix="dake_smart_horizontal_") as temp_name:
        temp_root = Path(temp_name)
        part_paths: list[Path] = []
        for index, segment in enumerate(segments, start=1):
            part_path = temp_root / f"part_{index:02d}.mp4"
            completed = run_command(
                _part_args(ffmpeg_path, source_video, part_path, segment, source_audio_present),
                timeout=max(180, int(float(segment.get("duration_seconds") or 30) * 5 + 90)),
            )
            if completed.returncode != 0 or not part_path.exists():
                errors.append(f"part {index}: {(completed.stderr or completed.stdout or '').strip()[-1000:]}")
                continue
            part_paths.append(part_path)

        if len(part_paths) < MIN_SEGMENTS:
            emit(LOG_TEXT["smart_horizontal_failed"])
            log_path = _write_log(selected_dir, [*log_lines, "status: FAILED", "error: too few rendered parts", *errors])
            return _result(package_dir, selected_dir, output_path, sequence_path, log_path, segments, "FAILED", "unavailable", False, fallback, False, used_ollama, "Too few segments rendered.")

        concat_list = _write_concat_list(temp_root, part_paths)
        base_path = temp_root / "smart_horizontal_base.mp4"
        attempts = ["h264_nvenc", "libx264"] if nvenc_online else ["libx264"]
        concat_ok = False
        for active_encoder in attempts:
            completed = run_command(_concat_args(ffmpeg_path, concat_list, base_path, active_encoder), timeout=max(240, int(total_duration * 4 + 180)))
            if completed.returncode == 0 and base_path.exists():
                encoder = active_encoder
                concat_ok = True
                fallback = active_encoder == "libx264" and nvenc_online
                break
            errors.append(f"concat {active_encoder}: {(completed.stderr or completed.stdout or '').strip()[-1000:]}")

        if not concat_ok:
            emit(LOG_TEXT["smart_horizontal_failed"])
            log_path = _write_log(selected_dir, [*log_lines, "status: FAILED", "error: concat failed", *errors])
            return _result(package_dir, selected_dir, output_path, sequence_path, log_path, segments, "FAILED", "unavailable", False, fallback, False, used_ollama, "Smart horizontal concat failed.")

        if output_path.exists():
            output_path.unlink()
        if bgm_path:
            mixed = run_command(_mix_bgm_args(ffmpeg_path, base_path, bgm_path, output_path, total_duration), timeout=max(240, int(total_duration * 4 + 180)))
            bgm_used = mixed.returncode == 0 and output_path.exists()
            if not bgm_used:
                errors.append(f"bgm mix: {(mixed.stderr or mixed.stdout or '').strip()[-1000:]}")
                shutil.copy2(base_path, output_path)
        else:
            shutil.copy2(base_path, output_path)

    status = "COMPLETED" if output_path.exists() else "FAILED"
    if status == "COMPLETED":
        emit(LOG_TEXT["smart_horizontal_created"])
    else:
        emit(LOG_TEXT["smart_horizontal_failed"])
    log_path = _write_log(
        selected_dir,
        [
            *log_lines,
            f"status: {status}",
            f"encoder: {encoder if status == 'COMPLETED' else 'unavailable'}",
            f"nvenc_used: {str(encoder == 'h264_nvenc' and status == 'COMPLETED').lower()}",
            f"fallback: {str(fallback).lower()}",
            f"bgm_used: {str(bgm_used).lower()}",
            f"output_path: {output_path}",
            "segments:",
            *[f"- {item.get('start')} - {item.get('end')} / {item.get('type')} / {item.get('reason')}" for item in segments],
            "errors:",
            *[f"- {error}" for error in errors[:12]],
        ],
    )
    return _result(
        package_dir,
        selected_dir,
        output_path,
        sequence_path,
        log_path,
        segments,
        status,
        encoder if status == "COMPLETED" else "unavailable",
        encoder == "h264_nvenc" and status == "COMPLETED",
        fallback,
        bgm_used,
        used_ollama,
        "smart_horizontal_edit.mp4 created." if status == "COMPLETED" else "Smart Horizontal Edit failed.",
    )


def _result(
    package_dir: Path,
    selected_dir: Path,
    output_path: Path,
    sequence_path: Path,
    log_path: Path,
    segments: list[dict[str, Any]],
    status: str,
    encoder: str,
    nvenc_used: bool,
    fallback: bool,
    bgm_used: bool,
    used_ollama: bool,
    message: str,
) -> dict[str, Any]:
    return {
        "status": status,
        "package_dir": str(package_dir),
        "selected_dir": str(selected_dir),
        "output_path": str(output_path),
        "sequence_path": str(sequence_path),
        "log_path": str(log_path),
        "segments": segments,
        "segment_count": len(segments),
        "total_duration": round(sum(float(item.get("duration_seconds") or 0) for item in segments), 3),
        "size": OUTPUT_SIZE,
        "encoder": encoder,
        "nvenc_used": nvenc_used,
        "fallback": fallback,
        "bgm_used": bgm_used,
        "fade_used": True,
        "used_ollama": used_ollama,
        "message": message,
    }
