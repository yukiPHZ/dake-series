from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

from core.app_config import human_size, outputs_dir, safe_project_name
from core.media_probe import MediaInfo


@dataclass(frozen=True)
class ProjectPaths:
    root: Path
    shorts_dir: Path
    metadata_dir: Path
    logs_dir: Path

    @classmethod
    def from_root(cls, root: Path) -> "ProjectPaths":
        return cls(root=root, shorts_dir=root / "shorts", metadata_dir=root / "metadata", logs_dir=root / "logs")


def create_project(source_path: Path) -> ProjectPaths:
    project_root = outputs_dir() / safe_project_name(source_path.name)
    project = ProjectPaths.from_root(project_root)
    for directory in [project.root, project.shorts_dir, project.metadata_dir, project.logs_dir]:
        directory.mkdir(parents=True, exist_ok=True)
    return project


def write_source_manifest(project: ProjectPaths, source_path: Path) -> Path:
    stat = source_path.stat()
    payload = {
        "source_file_name": source_path.name,
        "source_extension": source_path.suffix.lower(),
        "source_file_size": stat.st_size,
        "source_file_size_human": human_size(stat.st_size),
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "note": "Full local source paths are intentionally not stored.",
    }
    path = project.root / "source_manifest.json"
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def write_media_info(project: ProjectPaths, media_info: MediaInfo | None, reason: str = "") -> Path:
    path = project.root / "media_info.json"
    if media_info is None:
        payload: dict[str, Any] = {"available": False, "reason": reason}
    else:
        payload = {
            "available": True,
            "file_name": media_info.file_name,
            "file_size_bytes": media_info.file_size_bytes,
            "duration": media_info.duration,
            "width": media_info.width,
            "height": media_info.height,
            "fps": media_info.fps,
            "video_codec": media_info.video_codec,
            "audio_present": media_info.audio_present,
            "audio_codec": media_info.audio_codec,
        }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def write_preview_note(project: ProjectPaths, reason: str) -> Path:
    path = project.shorts_dir / "preview_unavailable.txt"
    path.write_text(f"Preview clip unavailable\nReason: {reason}\n", encoding="utf-8")
    return path


def write_metadata_files(project: ProjectPaths, metadata: dict[str, Any], preview_created: bool) -> list[Path]:
    title_ideas = metadata.get("title_ideas") or []
    description = str(metadata.get("description") or "")
    tags = metadata.get("tags") or []
    notes = metadata.get("notes") or []
    used_ollama = bool(metadata.get("used_ollama"))

    files = {
        "title_ideas.txt": "\n".join(f"- {item}" for item in title_ideas) + "\n",
        "description_draft.txt": description.strip() + "\n",
        "tags.txt": "\n".join(str(tag) for tag in tags) + "\n",
        "upload_notes.txt": "\n".join(
            [
                "Dakeユキズ稼働中 Phase 1",
                "自動公開はしません。",
                f"Ollama: {'used' if used_ollama else 'template fallback'}",
                f"Preview clip: {'created' if preview_created else 'not created'}",
                "",
                *[str(note) for note in notes],
            ]
        )
        + "\n",
    }
    written: list[Path] = []
    for name, content in files.items():
        path = project.metadata_dir / name
        path.write_text(content, encoding="utf-8")
        written.append(path)
    return written


def write_log_files(project: ProjectPaths, entries: list[str]) -> list[Path]:
    if not entries:
        entries = [f"{datetime.now().isoformat(timespec='seconds')} No log entries."]
    text = "\n".join(entries) + "\n"
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    project_log = project.logs_dir / "process.log"
    project_log.write_text(text, encoding="utf-8")

    app_log_dir = project.root.parents[1] / "logs"
    app_log_dir.mkdir(parents=True, exist_ok=True)
    app_log = app_log_dir / f"{stamp}_{project.root.name}.log"
    app_log.write_text(text, encoding="utf-8")
    return [project_log, app_log]
