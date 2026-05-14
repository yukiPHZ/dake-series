from __future__ import annotations

import json
import re
import traceback
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

from core.app_config import LOG_TEXT, SHORTS_REASON_TEXT, human_size, outputs_dir, seconds_to_timecode
from core.cli_checker import is_ollama_api_ready
from core.media_probe import MediaInfo, probe_media
from core.ollama_client import build_metadata_draft
from core.project_writer import ProjectPaths
from core.shorts_analyzer import create_shorts_candidates, write_shorts_candidates
from core.transcription import TranscriptionResult, transcribe_media

LogCallback = Callable[[str], None]
ProgressCallback = Callable[[float], None]


def packages_dir() -> Path:
    return outputs_dir() / "packages"


def _safe_video_name(video_path: Path) -> str:
    stem = video_path.stem or "video"
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", stem)
    cleaned = re.sub(r"\s+", "_", cleaned).strip("._ ")
    return (cleaned or "video")[:48]


def _create_package_dir(video_path: Path) -> Path:
    base = packages_dir()
    base.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    name = f"{stamp}_{_safe_video_name(video_path)}"
    package_root = base / name
    counter = 2
    while package_root.exists():
        package_root = base / f"{name}_{counter:02d}"
        counter += 1
    for directory in [package_root, package_root / "metadata", package_root / "logs"]:
        directory.mkdir(parents=True, exist_ok=True)
    return package_root


def _media_payload(media_info: MediaInfo) -> dict[str, Any]:
    return {
        "available": True,
        "file_name": media_info.file_name,
        "file_size_bytes": media_info.file_size_bytes,
        "file_size_human": human_size(media_info.file_size_bytes),
        "duration": media_info.duration,
        "duration_timecode": seconds_to_timecode(media_info.duration),
        "width": media_info.width,
        "height": media_info.height,
        "fps": media_info.fps,
        "video_codec": media_info.video_codec,
        "audio_present": media_info.audio_present,
        "audio_codec": media_info.audio_codec,
    }


def _write_media_info(package_root: Path, media_info: MediaInfo) -> Path:
    path = package_root / "media_info.json"
    path.write_text(json.dumps(_media_payload(media_info), ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def _write_media_info_unavailable(package_root: Path, reason: str) -> Path:
    path = package_root / "media_info_unavailable.json"
    payload = {
        "available": False,
        "reason": reason,
        "created_at": datetime.now().isoformat(timespec="seconds"),
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def _fallback_shorts_candidates() -> list[dict[str, object]]:
    candidates: list[dict[str, object]] = []
    clip_length = 45.0
    for index, start in enumerate([0.0, 45.0, 90.0], start=1):
        end = start + clip_length
        candidates.append(
            {
                "id": index,
                "start": seconds_to_timecode(start),
                "end": seconds_to_timecode(end),
                "duration": seconds_to_timecode(clip_length),
                "start_seconds": round(start, 3),
                "end_seconds": round(end, 3),
                "duration_seconds": round(clip_length, 3),
                "reason": SHORTS_REASON_TEXT["duration_unknown"],
                "status": "candidate",
            }
        )
    return candidates


def _as_text_list(value: object, fallback: list[str], limit: int | None = None) -> list[str]:
    if isinstance(value, list):
        items = [str(item).strip() for item in value if str(item).strip()]
    else:
        items = []
    if not items:
        items = fallback
    if limit is not None:
        items = items[:limit]
    return items


def _metadata_payload(metadata: dict[str, Any], source_name: str) -> dict[str, Any]:
    title_ideas = _as_text_list(
        metadata.get("title_ideas"),
        ["稼働中。夜に作る。", "止まらず作る。", "静かな作業机。", "今日も、少し進める。"],
        limit=7,
    )
    if len(title_ideas) < 3:
        title_ideas.extend(["静かな作業机。", "今日も、少し進める。"])
    title_ideas = title_ideas[:7]

    description = str(metadata.get("description") or "").strip()
    if "Links" not in description:
        description = (
            f"{description}\n\n" if description else "制作記録です。\n\n"
        ) + "Links\nPEAKHEADZ:\nDAKE:\nGitHub:\n\nMemo\n- 自動公開はしていません。"
    if "PEAKHEADZ" not in description:
        description += "\nPEAKHEADZ:"
    if "DAKE" not in description:
        description += "\nDAKE:"
    if "GitHub" not in description:
        description += "\nGitHub:"

    tags = _as_text_list(
        metadata.get("tags"),
        ["稼働中", "DAKE", "PEAKHEADZ", "quiet workflow", "作業動画", "AI開発", "Python", "GitHub"],
    )
    notes = _as_text_list(
        metadata.get("notes"),
        [
            "- 公開前にタイトル確認",
            "- サムネ未生成",
            "- BGM未適用",
            "- Shorts候補は自動抽出のため要確認",
            "- 自動公開はしていません",
        ],
    )
    if not any("自動公開" in note for note in notes):
        notes.append("- 自動公開はしていません")
    if not any("Source" in note or "素材" in note for note in notes):
        notes.append(f"- Source: {source_name}")

    return {
        "title_ideas": title_ideas,
        "description": description,
        "tags": tags,
        "notes": notes,
        "used_ollama": bool(metadata.get("used_ollama")),
        "ollama_model": str(metadata.get("ollama_model") or ""),
    }


def _write_metadata_files(package_root: Path, metadata: dict[str, Any], source_name: str) -> list[Path]:
    payload = _metadata_payload(metadata, source_name)
    metadata_dir = package_root / "metadata"
    files = {
        "title_ideas.txt": "\n".join(f"- {item}" for item in payload["title_ideas"]) + "\n",
        "description_draft.txt": payload["description"].strip() + "\n",
        "tags.txt": "\n".join(str(tag) for tag in payload["tags"]) + "\n",
        "upload_notes.txt": "\n".join(str(note) for note in payload["notes"]) + "\n",
    }
    written: list[Path] = []
    for name, content in files.items():
        path = metadata_dir / name
        path.write_text(content, encoding="utf-8")
        written.append(path)
    return written


def _write_package_log(package_root: Path, entries: list[str]) -> Path:
    path = package_root / "logs" / "package_log.txt"
    if not entries:
        entries = [f"{datetime.now().isoformat(timespec='seconds')} No log entries."]
    path.write_text("\n".join(entries).strip() + "\n", encoding="utf-8")
    return path


def _write_package_meta(package_root: Path, video_path: Path) -> Path:
    path = package_root / "package_meta.json"
    payload = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "source_video_name": video_path.name,
        "source_video_path": str(video_path),
        "note": "Source path is stored locally for selected short preview generation. No upload is performed.",
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def _record_file(path: Path) -> str:
    try:
        return str(path.relative_to(outputs_dir()))
    except Exception:
        return str(path)


def generate_posting_package(
    video_path: Path,
    ffprobe_path: str | None,
    ollama_ready: bool,
    log: LogCallback | None = None,
    progress: ProgressCallback | None = None,
) -> dict[str, Any]:
    package_root = _create_package_dir(video_path)
    log_entries: list[str] = []
    generated: list[str] = []

    def emit(message: str) -> None:
        line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} {message}"
        log_entries.append(line)
        if log:
            log(message)

    def tick(value: float) -> None:
        if progress:
            progress(max(0.0, min(1.0, value)))

    media_info: MediaInfo | None = None
    transcript_result: TranscriptionResult | None = None
    used_ollama = False

    try:
        emit(LOG_TEXT["posting_package_start"])
        meta_path = _write_package_meta(package_root, video_path)
        generated.append(_record_file(meta_path))
        log_entries.append(f"Selected file: {video_path.name}")
        log_entries.append(f"source_video_path: {video_path}")
        log_entries.append(f"Source size: {human_size(video_path.stat().st_size) if video_path.exists() else 'unknown'}")
        tick(0.04)

        if ffprobe_path:
            try:
                media_info = probe_media(video_path, ffprobe_path)
                media_path = _write_media_info(package_root, media_info)
                generated.append(_record_file(media_path))
                emit(LOG_TEXT["posting_media_ready"])
            except Exception as exc:
                media_path = _write_media_info_unavailable(package_root, str(exc))
                generated.append(_record_file(media_path))
                emit(LOG_TEXT["posting_media_unavailable"])
        else:
            media_path = _write_media_info_unavailable(package_root, "FFprobe is missing.")
            generated.append(_record_file(media_path))
            emit(LOG_TEXT["posting_media_unavailable"])
        tick(0.18)

        transcript_result = transcribe_media(
            video_path=video_path,
            project_dir=package_root,
            log=emit,
            progress=lambda value: tick(0.20 + value * 0.35),
        )
        if transcript_result.transcript_path:
            generated.append(_record_file(transcript_result.transcript_path))
        if transcript_result.srt_path:
            generated.append(_record_file(transcript_result.srt_path))
        if transcript_result.unavailable_path:
            generated.append(_record_file(transcript_result.unavailable_path))
        tick(0.56)

        if media_info and media_info.duration > 0:
            candidates = create_shorts_candidates(media_info.duration, transcript_result.srt_path if transcript_result else None)
        else:
            candidates = _fallback_shorts_candidates()
        shorts_path = write_shorts_candidates(ProjectPaths.from_root(package_root), candidates)
        generated.append(_record_file(shorts_path))
        emit(LOG_TEXT["posting_shorts_ready"])
        tick(0.68)

        metadata = build_metadata_draft(
            project_name=package_root.name,
            source_name=video_path.name,
            media_info=media_info,
            transcript_path=transcript_result.transcript_path if transcript_result else None,
        )
        used_ollama = bool(metadata.get("used_ollama"))
        metadata_files = _write_metadata_files(package_root, metadata, video_path.name)
        generated.extend(_record_file(path) for path in metadata_files)
        emit(LOG_TEXT["posting_titles_ready"])
        if not used_ollama and (ollama_ready or is_ollama_api_ready()):
            emit(LOG_TEXT["posting_ollama_fallback"])
        elif not used_ollama:
            emit(LOG_TEXT["posting_ollama_fallback"])
        tick(0.88)

        log_entries.extend(
            [
                f"Media info: {'available' if media_info else 'unavailable'}",
                f"Transcription: {'available' if transcript_result and transcript_result.available else 'unavailable'}",
                f"Shorts candidates: {len(candidates)}",
                f"Metadata: {'Ollama' if used_ollama else 'template fallback'}",
                "YouTube auto upload: disabled",
                "Source video modified: no",
                f"Package folder: {package_root}",
            ]
        )
        emit(LOG_TEXT["posting_ready"])
        log_path = _write_package_log(package_root, log_entries)
        generated.append(_record_file(log_path))
        tick(1.0)
        return {
            "status": "COMPLETED",
            "package_dir": str(package_root),
            "generated": generated,
            "media_info": _media_payload(media_info) if media_info else None,
            "transcription": "READY" if transcript_result and transcript_result.available else "UNAVAILABLE",
            "shorts_count": len(candidates),
            "used_ollama": used_ollama,
            "message": LOG_TEXT["posting_ready"],
        }
    except Exception as exc:
        log_entries.append(traceback.format_exc())
        log_path = _write_package_log(package_root, log_entries)
        generated.append(_record_file(log_path))
        return {
            "status": "FAILED",
            "package_dir": str(package_root),
            "generated": generated,
            "media_info": _media_payload(media_info) if media_info else None,
            "transcription": "READY" if transcript_result and transcript_result.available else "UNAVAILABLE",
            "shorts_count": 0,
            "used_ollama": used_ollama,
            "message": str(exc),
        }
