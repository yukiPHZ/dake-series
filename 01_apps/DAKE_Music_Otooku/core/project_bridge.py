# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import re
import shutil
import time
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path

from .app_config import OLLAMA_BASE_URL, OUTPUTS_DIR, load_ollama_model_name
from .favorites import FAVORITE_INDEX_PATH, FavoriteRecord


PROJECTS_DIR = OUTPUTS_DIR / "projects"
PROJECT_SUBDIRS = ("raw", "bgm", "shorts", "notes", "thumbnails", "export", "upload")


@dataclass(frozen=True)
class FavoriteChoice:
    label: str
    record: FavoriteRecord


@dataclass(frozen=True)
class ProjectSuggestion:
    suggested_title: str
    mood: str
    suggested_use: tuple[str, ...]
    shorts_direction: tuple[str, ...]
    source: str = "template"
    response_time: float | None = None
    error: str = ""


@dataclass(frozen=True)
class ProjectBridgeResult:
    success: bool
    project_root: Path | None = None
    copied_bgm: Path | None = None
    suggestion: ProjectSuggestion | None = None
    error: str = ""


def _record_from_dict(data: dict) -> FavoriteRecord:
    return FavoriteRecord(
        original_path=str(data.get("original_path", "")),
        favorite_path=str(data.get("favorite_path", "")),
        file_name=str(data.get("file_name", "")),
        created_at=str(data.get("created_at", "")),
        preset=str(data.get("preset", "")),
        tags=str(data.get("tags", "")),
        duration=str(data.get("duration", "")),
        note=str(data.get("note", "")),
    )


def load_favorite_choices(index_path: Path = FAVORITE_INDEX_PATH) -> tuple[FavoriteChoice, ...]:
    try:
        data = json.loads(index_path.read_text(encoding="utf-8"))
        records = data.get("favorites", []) if isinstance(data, dict) else data
    except Exception:
        records = []

    choices: list[FavoriteChoice] = []
    for index, item in enumerate(records):
        if not isinstance(item, dict):
            continue
        record = _record_from_dict(item)
        if not record.favorite_path or not Path(record.favorite_path).exists():
            continue
        preset = f" / {record.preset}" if record.preset else ""
        label = f"{record.file_name}{preset}"
        if any(choice.label == label for choice in choices):
            label = f"{label} #{index + 1}"
        choices.append(FavoriteChoice(label=label, record=record))
    return tuple(choices)


def sanitize_project_box_name(raw: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_-]+", "_", raw.strip()).strip("_").lower()
    return cleaned[:64] or f"project_{datetime.now().strftime('%Y%m%d_%H%M%S')}"


def _unique_project_root(base_dir: Path, project_name: str) -> Path:
    root = base_dir / project_name
    if not root.exists():
        return root
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    candidate = base_dir / f"{project_name}_{timestamp}"
    if not candidate.exists():
        return candidate
    for index in range(2, 1000):
        candidate = base_dir / f"{project_name}_{timestamp}_{index:02d}"
        if not candidate.exists():
            return candidate
    return base_dir / f"{project_name}_{timestamp}_copy"


def _unique_file_path(target_dir: Path, source_name: str) -> Path:
    candidate = target_dir / source_name
    if not candidate.exists():
        return candidate
    source = Path(source_name)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    candidate = target_dir / f"{source.stem}_{timestamp}{source.suffix}"
    if not candidate.exists():
        return candidate
    for index in range(2, 1000):
        candidate = target_dir / f"{source.stem}_{timestamp}_{index:02d}{source.suffix}"
        if not candidate.exists():
            return candidate
    return target_dir / f"{source.stem}_{timestamp}_copy{source.suffix}"


def _dedupe_lines(lines: list[str], fallback: tuple[str, ...]) -> tuple[str, ...]:
    values: list[str] = []
    seen: set[str] = set()
    for line in [*lines, *fallback]:
        cleaned = line.strip().lstrip("-•* ").strip()
        if not cleaned or cleaned.lower() in seen:
            continue
        values.append(cleaned)
        seen.add(cleaned.lower())
        if len(values) >= 6:
            break
    return tuple(values or fallback)


def fallback_project_suggestion(record: FavoriteRecord, error: str = "") -> ProjectSuggestion:
    return ProjectSuggestion(
        suggested_title=record.preset or "quiet work video",
        mood=record.tags or "quiet work",
        suggested_use=(
            record.note or "深夜のコード作業",
            "Shorts背景",
            "静かな作業動画",
        ),
        shorts_direction=(
            "静かなタイピング",
            "深夜の机",
            "ミシン作業",
            "画面録画",
            "まだ作ってる。",
        ),
        error=error,
    )


def _parse_project_suggestion(text: str, record: FavoriteRecord, source: str, response_time: float) -> ProjectSuggestion:
    sections: dict[str, list[str]] = {
        "suggested_title": [],
        "mood": [],
        "suggested_use": [],
        "shorts_direction": [],
    }
    label_map = {
        "suggested_title": "suggested_title",
        "suggested title": "suggested_title",
        "title": "suggested_title",
        "mood": "mood",
        "suggested_use": "suggested_use",
        "suggested use": "suggested_use",
        "use": "suggested_use",
        "shorts_direction": "shorts_direction",
        "shorts direction": "shorts_direction",
        "shorts": "shorts_direction",
    }
    current_key: str | None = None
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        match = re.match(r"^([A-Za-z _]+):\s*(.*)$", line)
        if match:
            mapped = label_map.get(match.group(1).strip().lower())
            if mapped:
                current_key = mapped
                value = match.group(2).strip()
                if value:
                    sections[mapped].append(value)
                continue
            current_key = None
            continue
        if current_key:
            sections[current_key].append(line)

    fallback = fallback_project_suggestion(record)
    return ProjectSuggestion(
        suggested_title=" ".join(sections["suggested_title"]).strip() or fallback.suggested_title,
        mood=" ".join(sections["mood"]).strip() or fallback.mood,
        suggested_use=_dedupe_lines(sections["suggested_use"], fallback.suggested_use),
        shorts_direction=_dedupe_lines(sections["shorts_direction"], fallback.shorts_direction),
        source=source,
        response_time=response_time,
    )


def generate_project_suggestion(
    record: FavoriteRecord,
    available_models: tuple[str, ...] = (),
    use_local_brain: bool = True,
    base_url: str = OLLAMA_BASE_URL,
    timeout: int = 60,
) -> ProjectSuggestion:
    if not use_local_brain:
        return fallback_project_suggestion(record)

    try:
        import requests
    except Exception as exc:
        return fallback_project_suggestion(record, f"requests import failed: {exc}")

    model = available_models[0] if available_models else load_ollama_model_name()
    prompt = (
        "You are a quiet local assistant for 音を置く. "
        "This is not video editing. Prepare a practical video production box idea for this BGM. "
        "Keep it concise and useful. Return exactly these labels:\n"
        "suggested_title:\n"
        "mood:\n"
        "suggested_use:\n"
        "shorts_direction:\n\n"
        f"BGM: {record.file_name}\n"
        f"Preset: {record.preset}\n"
        f"Tags: {record.tags}\n"
        f"Duration: {record.duration}\n"
        f"Note: {record.note}\n"
        "Question: このBGMはどんな動画に向いているか。"
    )
    payload = {
        "model": model,
        "prompt": prompt,
        "stream": False,
        "options": {
            "temperature": 0.25,
            "num_predict": 260,
        },
    }
    try:
        started_at = time.perf_counter()
        response = requests.post(f"{base_url.rstrip('/')}/api/generate", json=payload, timeout=timeout)
        response.raise_for_status()
        response_time = time.perf_counter() - started_at
        body = str(response.json().get("response", "")).strip()
        if not body:
            raise RuntimeError("Ollama returned an empty response")
        return _parse_project_suggestion(body, record, f"ollama:{model}", response_time)
    except Exception as exc:
        return fallback_project_suggestion(record, str(exc))


def _write_project_notes(
    project_root: Path,
    project_name: str,
    record: FavoriteRecord,
    copied_bgm: Path,
    suggestion: ProjectSuggestion,
) -> None:
    lines = [
        "Project:",
        project_name,
        "",
        "Selected Preset:",
        record.preset,
        "",
        "BGM:",
        copied_bgm.name,
        "",
        "Suggested Title:",
        suggestion.suggested_title,
        "",
        "Mood:",
        suggestion.mood,
        "",
        "Suggested Use:",
        *suggestion.suggested_use,
        "",
        "Shorts Direction:",
        *suggestion.shorts_direction,
        "",
        f"Source: {suggestion.source}",
    ]
    if suggestion.response_time is not None:
        lines.append(f"Response Time: {suggestion.response_time:.2f}s")
    if suggestion.error:
        lines.append(f"Fallback Reason: {suggestion.error}")
    (project_root / "notes" / "project_notes.txt").write_text("\n".join(lines).strip() + "\n", encoding="utf-8")


def _write_export_log(
    project_root: Path,
    project_name: str,
    record: FavoriteRecord,
    copied_bgm: Path,
    suggestion: ProjectSuggestion,
) -> None:
    lines = [
        "# Project Bridge Export Log",
        "",
        f"Created At: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"Project: {project_name}",
        f"Favorite: {record.favorite_path}",
        f"BGM: {copied_bgm}",
        f"Suggestion Source: {suggestion.source}",
    ]
    if suggestion.response_time is not None:
        lines.append(f"Response Time: {suggestion.response_time:.2f}s")
    (project_root / "export_log.txt").write_text("\n".join(lines).strip() + "\n", encoding="utf-8")


def create_project_box(
    project_name: str,
    favorite_record: FavoriteRecord,
    available_models: tuple[str, ...] = (),
    base_projects_dir: Path = PROJECTS_DIR,
    use_local_brain: bool = True,
) -> ProjectBridgeResult:
    source = Path(favorite_record.favorite_path)
    if not source.exists():
        return ProjectBridgeResult(False, error="favorite audio file was not found")

    clean_name = sanitize_project_box_name(project_name)
    project_root = _unique_project_root(base_projects_dir, clean_name)
    try:
        for subdir in PROJECT_SUBDIRS:
            (project_root / subdir).mkdir(parents=True, exist_ok=True)
        copied_bgm = _unique_file_path(project_root / "bgm", source.name)
        shutil.copy2(source, copied_bgm)
        suggestion = generate_project_suggestion(favorite_record, available_models, use_local_brain=use_local_brain)
        _write_project_notes(project_root, clean_name, favorite_record, copied_bgm, suggestion)
        _write_export_log(project_root, clean_name, favorite_record, copied_bgm, suggestion)
        return ProjectBridgeResult(True, project_root=project_root, copied_bgm=copied_bgm, suggestion=suggestion)
    except Exception as exc:
        return ProjectBridgeResult(False, project_root=project_root, error=str(exc))


def record_to_dict(record: FavoriteRecord) -> dict:
    return asdict(record)
