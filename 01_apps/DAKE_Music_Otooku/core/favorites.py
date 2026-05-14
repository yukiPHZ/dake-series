# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import re
import shutil
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path

from .app_config import DATA_DIR
from .audio_probe import probe_duration
from .presets import MusicPreset


FAVORITES_DIR = DATA_DIR / "favorites"
FAVORITES_AUDIO_DIR = FAVORITES_DIR / "audio"
FAVORITES_NOTES_DIR = FAVORITES_DIR / "notes"
FAVORITE_INDEX_PATH = FAVORITES_DIR / "favorite_index.json"
FAVORITE_NOTE_PATH = FAVORITES_NOTES_DIR / "favorite_note.txt"


@dataclass(frozen=True)
class FavoriteRecord:
    original_path: str
    favorite_path: str
    file_name: str
    created_at: str
    preset: str
    tags: str
    duration: str
    note: str


@dataclass(frozen=True)
class FavoriteSaveResult:
    success: bool
    record: FavoriteRecord | None = None
    error: str = ""


def ensure_favorites_dirs(base_dir: Path = FAVORITES_DIR) -> tuple[Path, Path, Path, Path]:
    audio_dir = base_dir / "audio"
    notes_dir = base_dir / "notes"
    index_path = base_dir / "favorite_index.json"
    note_path = notes_dir / "favorite_note.txt"
    audio_dir.mkdir(parents=True, exist_ok=True)
    notes_dir.mkdir(parents=True, exist_ok=True)
    return audio_dir, notes_dir, index_path, note_path


def _read_text_if_exists(path: Path) -> str:
    try:
        if path.exists():
            return path.read_text(encoding="utf-8")
    except Exception:
        pass
    return ""


def _extract_section(text: str, label: str) -> str:
    pattern = rf"^{re.escape(label)}:\s*\n(.+?)(?:\n\n|$)"
    match = re.search(pattern, text, flags=re.MULTILINE | re.DOTALL)
    if not match:
        return ""
    return " ".join(line.strip() for line in match.group(1).splitlines() if line.strip())


def infer_favorite_metadata(
    source_path: Path,
    project_root: Path | None = None,
    preset: MusicPreset | None = None,
) -> tuple[str, str, str]:
    preset_name = preset.name if preset else ""
    tags = ", ".join(preset.tags) if preset else ""
    note = ""

    if project_root:
        candidates = (
            project_root / "notes" / "loop_notes.txt",
            project_root / "video_bgm_pack" / "notes" / "usage_note.txt",
            project_root / "notes" / "usage_note.txt",
            project_root / "prompts" / "music_direction.txt",
        )
        combined = "\n\n".join(_read_text_if_exists(path) for path in candidates)
        preset_name = preset_name or _extract_section(combined, "Preset")
        tags = tags or _extract_section(combined, "Preset Tags")
        note = _extract_section(combined, "Preset Use") or _extract_section(combined, "Suggested Use")

    if not tags:
        name_parts = source_path.stem.split("_")
        tags = ", ".join(part for part in name_parts[:3] if part and not part.endswith("s"))
    if not note:
        note = "Shorts / quiet work / midnight"
    return preset_name, tags, note


def _unique_favorite_path(audio_dir: Path, source_path: Path, created_at: datetime) -> Path:
    candidate = audio_dir / source_path.name
    if not candidate.exists():
        return candidate

    timestamp = created_at.strftime("%Y%m%d_%H%M%S")
    base = audio_dir / f"{source_path.stem}_{timestamp}{source_path.suffix}"
    if not base.exists():
        return base

    for index in range(2, 1000):
        candidate = audio_dir / f"{source_path.stem}_{timestamp}_{index:02d}{source_path.suffix}"
        if not candidate.exists():
            return candidate
    return audio_dir / f"{source_path.stem}_{timestamp}_copy{source_path.suffix}"


def _load_index(index_path: Path) -> list[dict]:
    try:
        data = json.loads(index_path.read_text(encoding="utf-8"))
        if isinstance(data, dict):
            records = data.get("favorites", [])
        else:
            records = data
        if isinstance(records, list):
            return [record for record in records if isinstance(record, dict)]
    except Exception:
        pass
    return []


def _write_index(index_path: Path, records: list[dict]) -> None:
    payload = {
        "favorites": records,
    }
    index_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def _append_note(note_path: Path, record: FavoriteRecord) -> None:
    lines = [
        f"[{record.created_at[:16]}]",
        "File:",
        record.file_name,
        "",
        "Preset:",
        record.preset,
        "",
        "Use:",
        record.note,
        "",
        "---",
        "",
    ]
    with note_path.open("a", encoding="utf-8") as handle:
        handle.write("\n".join(lines))


def save_favorite_audio(
    source_path: Path,
    project_root: Path | None = None,
    preset: MusicPreset | None = None,
    favorites_dir: Path = FAVORITES_DIR,
) -> FavoriteSaveResult:
    source_path = Path(source_path)
    if not source_path.exists() or source_path.suffix.lower() not in {".mp3", ".wav"}:
        return FavoriteSaveResult(False, error="selected audio file was not found")

    try:
        audio_dir, _notes_dir, index_path, note_path = ensure_favorites_dirs(favorites_dir)
        created = datetime.now()
        favorite_path = _unique_favorite_path(audio_dir, source_path, created)
        shutil.copy2(source_path, favorite_path)

        duration = probe_duration(source_path)
        preset_name, tags, note = infer_favorite_metadata(source_path, project_root, preset)
        record = FavoriteRecord(
            original_path=str(source_path.resolve()),
            favorite_path=str(favorite_path.resolve()),
            file_name=favorite_path.name,
            created_at=created.strftime("%Y-%m-%d %H:%M:%S"),
            preset=preset_name,
            tags=tags,
            duration="" if duration is None else f"{duration:.1f}s",
            note=note,
        )
        records = _load_index(index_path)
        records.append(asdict(record))
        _write_index(index_path, records)
        _append_note(note_path, record)
        return FavoriteSaveResult(True, record=record)
    except Exception as exc:
        return FavoriteSaveResult(False, error=str(exc))
