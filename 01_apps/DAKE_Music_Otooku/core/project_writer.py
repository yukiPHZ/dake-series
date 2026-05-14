# -*- coding: utf-8 -*-
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from .app_config import OUTPUTS_DIR
from .prompt_builder import MusicDirection


@dataclass(frozen=True)
class ProjectPaths:
    root: Path
    audio: Path
    prompts: Path
    notes: Path
    logs: Path


def sanitize_project_name(raw: str) -> str:
    ascii_slug = re.sub(r"[^A-Za-z0-9_-]+", "_", raw).strip("_").lower()
    if not ascii_slug:
        ascii_slug = "otooku"
    return ascii_slug[:48]


def make_project_name(prompt: str) -> str:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    slug = sanitize_project_name(prompt)
    if slug == "otooku":
        return f"otooku_{timestamp}"
    return f"otooku_{timestamp}_{slug}"


def create_project(project_name: str | None = None, base_output_dir: Path = OUTPUTS_DIR) -> ProjectPaths:
    name = project_name or f"otooku_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    root = base_output_dir / name
    paths = ProjectPaths(
        root=root,
        audio=root / "audio",
        prompts=root / "prompts",
        notes=root / "notes",
        logs=root / "logs",
    )
    for path in (paths.audio, paths.prompts, paths.notes, paths.logs):
        path.mkdir(parents=True, exist_ok=True)
    return paths


def write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def write_project_files(
    paths: ProjectPaths,
    user_text: str,
    direction: MusicDirection,
    log_lines: list[str],
) -> None:
    write_text(paths.prompts / "music_direction.txt", direction.to_direction_text(user_text))
    write_text(paths.prompts / "musicgen_prompt.txt", direction.musicgen_prompt.strip() + "\n")
    write_text(paths.notes / "usage_note.txt", direction.to_usage_note())
    write_text(paths.logs / "process_log.txt", "\n".join(log_lines).strip() + "\n")


def write_loop_notes(
    paths: ProjectPaths,
    direction: MusicDirection,
    tags: tuple[str, ...],
    durations: tuple[int, ...],
    fade_in: float,
    fade_out: float,
    volume_mode: str,
    files: list[Path],
) -> None:
    lines = [
        "# Loop Pack Notes",
        "",
        "Mood:",
        direction.mood,
        "",
        "BPM:",
        direction.bpm,
        "",
        "Texture:",
        direction.texture,
        "",
        "Suggested Use:",
        direction.usage_idea,
        "",
        f"Tags: {', '.join(tags) if tags else 'quiet'}",
        f"Durations: {', '.join(str(duration) + 's' for duration in durations)}",
        f"Fade In: {fade_in:.1f}s",
        f"Fade Out: {fade_out:.1f}s",
        f"Volume: {volume_mode}",
        "",
        "Files:",
    ]
    lines.extend(f"- {path.name}" for path in files)
    write_text(paths.notes / "loop_notes.txt", "\n".join(lines).strip() + "\n")


def write_setup_needed(paths: ProjectPaths, reason: str) -> None:
    text = "\n".join(
        [
            "# setup_needed",
            "",
            reason,
            "",
            "AudioCraft / MusicGen が利用できる環境では、短い wav 生成まで進めます。",
            "未導入の場合も、musicgen_prompt.txt と設計メモを素材化の起点として使えます。",
        ]
    )
    write_text(paths.root / "setup_needed.txt", text.strip() + "\n")
