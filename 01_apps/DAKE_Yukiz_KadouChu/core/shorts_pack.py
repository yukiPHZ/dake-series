from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

from core.app_config import LOG_TEXT, seconds_to_timecode
from core.cli_checker import run_command
from core.ollama_client import generate_ollama_text
from core.selected_preview import find_source_video_path

LogCallback = Callable[[str], None]

PACK_TYPE = "quiet_flow"
PACK_DIR_NAME = "shorts_pack"
VERTICAL_SIZE = "1080x1920"
FADE_IN_SECONDS = 0.5
FADE_OUT_SECONDS = 0.6
BGM_VOLUME = 0.08
BGM_EXTENSIONS = {".mp3", ".wav", ".m4a", ".aac", ".flac", ".ogg"}

ROLE_SPECS = [
    {
        "type": "INTRO",
        "slug": "intro",
        "filename": "short_01_intro.mp4",
        "reason": "静かな導入",
        "text_direction": "minimal",
        "caption_idea": "稼働中。",
    },
    {
        "type": "WORK",
        "slug": "work",
        "filename": "short_02_work.mp4",
        "reason": "作業感",
        "text_direction": "minimal",
        "caption_idea": "まだ作ってる。",
    },
    {
        "type": "AFTERGLOW",
        "slug": "afterglow",
        "filename": "short_03_afterglow.mp4",
        "reason": "余熱",
        "text_direction": "minimal",
        "caption_idea": "整っています。",
    },
]


def _read_json(path: Path) -> Any:
    if not path.exists() or not path.is_file():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8", errors="replace"))
    except Exception:
        return None


def _read_text(path: Path, limit: int = 1800) -> str:
    if not path.exists() or not path.is_file():
        return ""
    return path.read_text(encoding="utf-8", errors="replace").strip()[:limit]


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


def _candidate_range(candidate: dict[str, Any]) -> tuple[float, float, float]:
    start = _time_to_seconds(candidate.get("start_seconds", candidate.get("start")))
    duration = _time_to_seconds(candidate.get("duration_seconds", candidate.get("duration")))
    end = _time_to_seconds(candidate.get("end_seconds", candidate.get("end")))
    if duration <= 0 and end > start:
        duration = end - start
    if end <= start and duration > 0:
        end = start + duration
    if duration <= 0:
        duration = 30.0
        end = start + duration
    return start, end, duration


def _read_candidates(package_dir: Path) -> list[dict[str, Any]]:
    payload = _read_json(package_dir / "shorts_candidates.json")
    if not isinstance(payload, list):
        return []
    candidates = [dict(item) for item in payload if isinstance(item, dict)]
    return sorted(candidates, key=lambda item: _candidate_range(item)[0])


def _fallback_candidates(duration: float) -> list[dict[str, Any]]:
    if duration <= 0:
        duration = 120.0
    clip_length = min(45.0, max(12.0, duration / 4))
    starts = [
        0.0,
        max(0.0, (duration - clip_length) * 0.46),
        max(0.0, duration - clip_length),
    ]
    candidates: list[dict[str, Any]] = []
    for index, start in enumerate(starts, start=1):
        end = min(duration, start + clip_length)
        candidates.append(
            {
                "id": index,
                "start": seconds_to_timecode(start),
                "end": seconds_to_timecode(end),
                "duration": seconds_to_timecode(max(1.0, end - start)),
                "start_seconds": round(start, 3),
                "end_seconds": round(end, 3),
                "duration_seconds": round(max(1.0, end - start), 3),
                "reason": "補助脳：素材尺から静かに分けました。",
                "status": "candidate",
            }
        )
    return candidates


def _selected_bgm(package_dir: Path) -> Path | None:
    bgm_dir = package_dir / "selected" / "bgm"
    if not bgm_dir.exists():
        return None
    try:
        files = sorted(path for path in bgm_dir.iterdir() if path.is_file() and path.suffix.lower() in BGM_EXTENSIONS)
    except Exception:
        return None
    return files[0] if files else None


def _role_from_candidate(spec: dict[str, str], candidate: dict[str, Any]) -> dict[str, Any]:
    start, end, duration = _candidate_range(candidate)
    return {
        "type": spec["type"],
        "start": seconds_to_timecode(start),
        "end": seconds_to_timecode(end),
        "duration": seconds_to_timecode(duration),
        "start_seconds": round(start, 3),
        "end_seconds": round(end, 3),
        "duration_seconds": round(duration, 3),
        "reason": candidate.get("reason") or spec["reason"],
        "status": "candidate",
        "text_direction": spec["text_direction"],
        "caption_idea": spec["caption_idea"],
    }


def _template_roles(candidates: list[dict[str, Any]], duration: float) -> list[dict[str, Any]]:
    source = candidates if len(candidates) >= 3 else _fallback_candidates(duration)
    indexes = [0, max(0, len(source) // 2), len(source) - 1]
    roles: list[dict[str, Any]] = []
    for spec, index in zip(ROLE_SPECS, indexes):
        candidate = source[min(max(0, index), len(source) - 1)]
        role = _role_from_candidate(spec, candidate)
        role["reason"] = spec["reason"]
        roles.append(role)
    return roles


def _ollama_roles(package_dir: Path, candidates: list[dict[str, Any]], roles: list[dict[str, Any]], ollama_ready: bool) -> tuple[list[dict[str, Any]], bool, str]:
    if not ollama_ready:
        return roles, False, ""
    context = {
        "pack_type": PACK_TYPE,
        "roles": [role["type"] for role in roles],
        "candidates": candidates[:5],
        "assistant_review": _read_text(package_dir / "assistant_review.md", limit=1000),
        "assistant_recommendation": _read_text(package_dir / "assistant_recommendation.md", limit=1000),
        "transcript": _read_text(package_dir / "transcript.txt", limit=1200),
        "metadata_title": _read_text(package_dir / "metadata" / "title_ideas.txt", limit=600),
        "bgm": [path.name for path in (package_dir / "selected" / "bgm").glob("*") if path.is_file()] if (package_dir / "selected" / "bgm").exists() else [],
    }
    prompt = (
        "You are the local assistant brain for Dakeユキズ稼働中.\n"
        "Assign three quiet YouTube Shorts roles from the candidates: INTRO, WORK, AFTERGLOW.\n"
        "Keep the timing practical. Do not suggest flashy edits.\n"
        "Return only JSON array with keys: type, start_seconds, end_seconds, reason, text_direction, caption_idea.\n\n"
        f"{json.dumps(context, ensure_ascii=False)}"
    )
    response = generate_ollama_text(prompt, timeout=45)
    if not response.get("ok"):
        return roles, False, str(response.get("reason") or "")
    text = str(response.get("text") or "")
    start = text.find("[")
    end = text.rfind("]")
    if start == -1 or end == -1 or end <= start:
        return roles, False, "Ollama response did not include a JSON array."
    try:
        parsed = json.loads(text[start : end + 1])
    except Exception as exc:
        return roles, False, str(exc)
    if not isinstance(parsed, list):
        return roles, False, "Ollama response was not a list."

    by_type = {str(item.get("type") or "").upper(): item for item in parsed if isinstance(item, dict)}
    updated: list[dict[str, Any]] = []
    for role in roles:
        item = by_type.get(str(role["type"]).upper())
        if not isinstance(item, dict):
            updated.append(role)
            continue
        start_seconds = _time_to_seconds(item.get("start_seconds", item.get("start")))
        end_seconds = _time_to_seconds(item.get("end_seconds", item.get("end")))
        if end_seconds <= start_seconds:
            updated.append(role)
            continue
        new_role = dict(role)
        new_role.update(
            {
                "start": seconds_to_timecode(start_seconds),
                "end": seconds_to_timecode(end_seconds),
                "duration": seconds_to_timecode(end_seconds - start_seconds),
                "start_seconds": round(start_seconds, 3),
                "end_seconds": round(end_seconds, 3),
                "duration_seconds": round(end_seconds - start_seconds, 3),
                "reason": str(item.get("reason") or role["reason"]),
                "text_direction": str(item.get("text_direction") or role["text_direction"]),
                "caption_idea": str(item.get("caption_idea") or role["caption_idea"]),
            }
        )
        updated.append(new_role)
    return updated, True, str(response.get("model") or "")


def _video_filter(duration: float) -> str:
    fade_in = min(FADE_IN_SECONDS, max(0.2, duration * 0.18))
    fade_out = min(FADE_OUT_SECONDS, max(0.2, duration * 0.22))
    fade_out_start = max(0.0, duration - fade_out)
    return (
        "[0:v]split=2[bg][fg];"
        "[bg]scale=1080:1920:force_original_aspect_ratio=increase,"
        "crop=1080:1920,boxblur=30:1[bgout];"
        "[fg]scale=1080:1920:force_original_aspect_ratio=decrease[fgout];"
        "[bgout][fgout]overlay=(W-w)/2:(H-h)/2,format=yuv420p,"
        f"fade=t=in:st=0:d={fade_in:.3f},"
        f"fade=t=out:st={fade_out_start:.3f}:d={fade_out:.3f}[v]"
    )


def _codec_args(encoder: str) -> list[str]:
    if encoder == "h264_nvenc":
        return ["-c:v", "h264_nvenc", "-preset", "p4", "-cq", "23"]
    return ["-c:v", "libx264", "-preset", "veryfast", "-crf", "23"]


def _render_args(
    ffmpeg_path: str,
    source_video: Path,
    output_path: Path,
    role: dict[str, Any],
    encoder: str,
    bgm_path: Path | None,
    source_audio_present: bool,
) -> list[str]:
    start = float(role.get("start_seconds") or 0)
    duration = max(1.0, float(role.get("duration_seconds") or 30))
    args = [
        ffmpeg_path,
        "-y",
        "-ss",
        f"{start:.3f}",
        "-i",
        str(source_video),
    ]
    if bgm_path is not None:
        args.extend(["-stream_loop", "-1", "-i", str(bgm_path)])
    args.extend(["-t", f"{duration:.3f}"])

    filters = [_video_filter(duration)]
    if bgm_path is not None and source_audio_present:
        fade_out_start = max(0.0, duration - FADE_OUT_SECONDS)
        filters.append(
            f"[0:a]volume=1.0,afade=t=in:st=0:d=0.300,afade=t=out:st={fade_out_start:.3f}:d={FADE_OUT_SECONDS:.3f}[a0];"
            f"[1:a]volume={BGM_VOLUME:.3f},atrim=0:{duration:.3f},"
            f"afade=t=in:st=0:d=0.300,afade=t=out:st={fade_out_start:.3f}:d={FADE_OUT_SECONDS:.3f}[a1];"
            "[a0][a1]amix=inputs=2:duration=first:dropout_transition=0[a]"
        )
    elif bgm_path is not None:
        fade_out_start = max(0.0, duration - FADE_OUT_SECONDS)
        filters.append(
            f"[1:a]volume={BGM_VOLUME:.3f},atrim=0:{duration:.3f},"
            f"afade=t=in:st=0:d=0.300,afade=t=out:st={fade_out_start:.3f}:d={FADE_OUT_SECONDS:.3f}[a]"
        )

    args.extend(["-filter_complex", ";".join(filters), "-map", "[v]"])
    if bgm_path is not None:
        args.extend(["-map", "[a]"])
    else:
        args.extend(["-map", "0:a?"])
    args.extend([*_codec_args(encoder), "-c:a", "aac", "-b:a", "192k", "-movflags", "+faststart", str(output_path)])
    return args


def _render_clip(
    ffmpeg_path: str,
    source_video: Path,
    pack_dir: Path,
    role: dict[str, Any],
    encoder: str,
    bgm_path: Path | None,
    source_audio_present: bool,
) -> dict[str, Any]:
    output_path = pack_dir / str(role["filename"])
    if output_path.exists():
        output_path.unlink()
    completed = run_command(
        _render_args(ffmpeg_path, source_video, output_path, role, encoder, bgm_path, source_audio_present),
        timeout=max(240, int(float(role.get("duration_seconds") or 30) * 6 + 120)),
    )
    return {
        "ok": completed.returncode == 0 and output_path.exists(),
        "output_path": output_path,
        "error": (completed.stderr or completed.stdout or "").strip()[-1200:],
    }


def _write_pack_log(pack_dir: Path, lines: list[str]) -> Path:
    path = pack_dir / "shorts_pack_log.txt"
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path


def generate_shorts_pack(
    package_dir: Path,
    ffmpeg_path: str | None,
    nvenc_online: bool,
    ollama_ready: bool,
    source_video_path: Path | None = None,
    log: LogCallback | None = None,
) -> dict[str, Any]:
    selected_dir = package_dir / "selected"
    pack_dir = selected_dir / PACK_DIR_NAME
    pack_dir.mkdir(parents=True, exist_ok=True)

    def emit(message: str) -> None:
        if log:
            log(message)

    emit(LOG_TEXT["shorts_pack_start"])
    source_video = source_video_path or find_source_video_path(package_dir)
    bgm_path = _selected_bgm(package_dir)
    candidates = _read_candidates(package_dir)
    duration = _media_duration(package_dir)
    roles = _template_roles(candidates, duration)
    roles, used_ollama, ollama_note = _ollama_roles(package_dir, candidates, roles, ollama_ready)
    emit(LOG_TEXT["shorts_pack_roles"])
    if bgm_path:
        emit(LOG_TEXT["shorts_pack_bgm"])

    log_lines = [
        f"executed_at: {datetime.now().isoformat(timespec='seconds')}",
        f"package_dir: {package_dir}",
        f"source_video: {source_video or ''}",
        f"pack_type: {PACK_TYPE}",
        f"target_count: {len(ROLE_SPECS)}",
        f"bgm: {bgm_path or ''}",
        f"bgm_volume: {BGM_VOLUME}",
        f"fade_in_seconds: {FADE_IN_SECONDS}",
        f"fade_out_seconds: {FADE_OUT_SECONDS}",
        f"used_ollama: {str(used_ollama).lower()}",
        f"ollama_note: {ollama_note}",
    ]

    if source_video is None or not source_video.exists():
        emit(LOG_TEXT["selected_preview_source_missing"])
        log_path = _write_pack_log(pack_dir, [*log_lines, "status: FAILED", "error: source video missing"])
        return {
            "status": "FAILED",
            "package_dir": str(package_dir),
            "selected_dir": str(selected_dir),
            "pack_dir": str(pack_dir),
            "pack_json": str(pack_dir / "shorts_pack.json"),
            "log_path": str(log_path),
            "clips": [],
            "generated_count": 0,
            "used_ollama": used_ollama,
            "bgm_applied": False,
            "nvenc_used": False,
            "fallback": False,
            "message": "Source video is required.",
        }
    if not ffmpeg_path:
        emit(LOG_TEXT["selected_preview_ffmpeg_missing"])
        log_path = _write_pack_log(pack_dir, [*log_lines, "status: FAILED", "error: ffmpeg missing"])
        return {
            "status": "FAILED",
            "package_dir": str(package_dir),
            "selected_dir": str(selected_dir),
            "pack_dir": str(pack_dir),
            "pack_json": str(pack_dir / "shorts_pack.json"),
            "log_path": str(log_path),
            "clips": [],
            "generated_count": 0,
            "used_ollama": used_ollama,
            "bgm_applied": False,
            "nvenc_used": False,
            "fallback": False,
            "message": "FFmpeg is required for Shorts Pack generation.",
        }

    emit(LOG_TEXT["shorts_pack_flow"])
    source_audio_present = _source_audio_present(package_dir)
    rendered: list[dict[str, Any]] = []
    any_fallback = False
    any_nvenc = False
    any_bgm = False
    errors: list[str] = []

    for spec, role in zip(ROLE_SPECS, roles):
        role = dict(role)
        role["filename"] = spec["filename"]
        attempts: list[tuple[str, Path | None]] = []
        if nvenc_online:
            attempts.append(("h264_nvenc", bgm_path))
        attempts.append(("libx264", bgm_path))
        if bgm_path is not None:
            attempts.append(("libx264", None))

        clip_result: dict[str, Any] | None = None
        for encoder, active_bgm in attempts:
            result = _render_clip(ffmpeg_path, source_video, pack_dir, role, encoder, active_bgm, source_audio_present)
            if result["ok"]:
                clip_result = result
                role["encoder"] = encoder
                role["nvenc_used"] = encoder == "h264_nvenc"
                role["fallback"] = encoder == "libx264" and nvenc_online
                role["bgm_applied"] = active_bgm is not None
                break
            errors.append(f"{role['type']} {encoder}: {result['error']}")
            if encoder == "h264_nvenc":
                any_fallback = True

        output_path = Path(str(clip_result["output_path"])) if clip_result else pack_dir / str(role["filename"])
        role["output_path"] = str(output_path)
        role["status"] = "completed" if clip_result else "failed"
        role["fade"] = {"in_seconds": FADE_IN_SECONDS, "out_seconds": FADE_OUT_SECONDS}
        rendered.append(role)
        any_nvenc = any_nvenc or bool(role.get("nvenc_used"))
        any_fallback = any_fallback or bool(role.get("fallback"))
        any_bgm = any_bgm or bool(role.get("bgm_applied"))

    generated_count = sum(1 for clip in rendered if clip.get("status") == "completed")
    pack_payload = {
        "pack_type": PACK_TYPE,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "source_video": str(source_video),
        "target_count": len(ROLE_SPECS),
        "generated_count": generated_count,
        "used_ollama": used_ollama,
        "ollama_note": ollama_note,
        "bgm": str(bgm_path or ""),
        "bgm_volume": BGM_VOLUME if any_bgm else 0,
        "output_size": VERTICAL_SIZE,
        "clips": rendered,
    }
    pack_json = pack_dir / "shorts_pack.json"
    pack_json.write_text(json.dumps(pack_payload, ensure_ascii=False, indent=2), encoding="utf-8")

    status = "COMPLETED" if generated_count == len(ROLE_SPECS) else "FAILED"
    if status == "COMPLETED":
        emit(LOG_TEXT["shorts_pack_created"])
    else:
        emit(LOG_TEXT["shorts_pack_failed"])
    log_path = _write_pack_log(
        pack_dir,
        [
            *log_lines,
            f"status: {status}",
            f"generated_count: {generated_count}",
            f"nvenc_used: {str(any_nvenc).lower()}",
            f"fallback: {str(any_fallback).lower()}",
            f"bgm_applied: {str(any_bgm).lower()}",
            f"pack_json: {pack_json}",
            "outputs:",
            *[f"- {clip.get('type')}: {clip.get('output_path')} ({clip.get('status')})" for clip in rendered],
            "errors:",
            *[f"- {error}" for error in errors[:12]],
        ],
    )
    return {
        "status": status,
        "package_dir": str(package_dir),
        "selected_dir": str(selected_dir),
        "pack_dir": str(pack_dir),
        "pack_json": str(pack_json),
        "log_path": str(log_path),
        "clips": rendered,
        "generated_count": generated_count,
        "used_ollama": used_ollama,
        "bgm_applied": any_bgm,
        "nvenc_used": any_nvenc,
        "fallback": any_fallback,
        "message": f"{generated_count} Shorts generated." if status == "COMPLETED" else "Shorts Pack generation failed.",
    }
