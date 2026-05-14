from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Callable

from core.app_config import LOG_TEXT

LogCallback = Callable[[str], None]

SELECTED_DIR_NAME = "selected"


def _read_text(path: Path) -> str:
    if not path.exists() or not path.is_file():
        return ""
    return path.read_text(encoding="utf-8", errors="replace").strip()


def _read_json(path: Path) -> Any:
    if not path.exists() or not path.is_file():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8", errors="replace"))
    except Exception:
        return None


def _title_lines(text: str) -> list[str]:
    lines: list[str] = []
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        line = line.lstrip("-").strip()
        if line:
            lines.append(line)
    return lines[:7]


def _short_label(candidate: dict[str, Any], index: int) -> str:
    start = str(candidate.get("start") or "--")
    end = str(candidate.get("end") or "--")
    reason = str(candidate.get("reason") or "candidate")
    return f"#{index} {start} - {end} / {reason}"


def _title_label(title: str, index: int) -> str:
    return f"#{index} {title}"


def read_selected_candidates(package_dir: Path) -> dict[str, Any]:
    metadata_dir = package_dir / "metadata"
    shorts_payload = _read_json(package_dir / "shorts_candidates.json")
    shorts = shorts_payload if isinstance(shorts_payload, list) else []
    shorts = [item for item in shorts if isinstance(item, dict)][:5]

    title_text = _read_text(metadata_dir / "title_ideas.txt")
    titles = _title_lines(title_text)
    description = _read_text(metadata_dir / "description_draft.txt")
    tags = _read_text(metadata_dir / "tags.txt")
    notes = _read_text(metadata_dir / "upload_notes.txt")
    review_exists = (package_dir / "assistant_review.md").exists()

    return {
        "package_dir": str(package_dir),
        "shorts": shorts,
        "short_labels": [_short_label(candidate, index) for index, candidate in enumerate(shorts, start=1)],
        "titles": titles,
        "title_labels": [_title_label(title, index) for index, title in enumerate(titles, start=1)],
        "has_description": bool(description),
        "has_tags": bool(tags),
        "has_notes": bool(notes),
        "has_review": review_exists,
        "description": description,
        "tags": tags,
        "notes": notes,
    }


def _fallback_short() -> dict[str, Any]:
    return {
        "id": 0,
        "start": "--",
        "end": "--",
        "duration": "--",
        "reason": "Shorts候補は未生成です。",
        "status": "unavailable",
    }


def _selected_short_text(candidate: dict[str, Any]) -> str:
    start = str(candidate.get("start") or "--")
    end = str(candidate.get("end") or "--")
    reason = str(candidate.get("reason") or "")
    duration = str(candidate.get("duration") or candidate.get("duration_seconds") or "--")
    return "\n".join(
        [
            f"- start: {start}",
            f"- end: {end}",
            f"- duration: {duration}",
            f"- reason: {reason}",
        ]
    )


def _write_text(path: Path, text: str) -> Path:
    path.write_text(text.rstrip() + "\n", encoding="utf-8")
    return path


def export_selected_draft(
    package_dir: Path,
    short_index: int | None = None,
    title_index: int | None = None,
    log: LogCallback | None = None,
) -> dict[str, Any]:
    candidates = read_selected_candidates(package_dir)
    selected_dir = package_dir / SELECTED_DIR_NAME
    selected_dir.mkdir(parents=True, exist_ok=True)

    shorts: list[dict[str, Any]] = candidates["shorts"]
    titles: list[str] = candidates["titles"]
    used_default = short_index is None or title_index is None
    if used_default and log:
        log(LOG_TEXT["selected_default_short"])

    safe_short_index = short_index if short_index is not None else 0
    safe_title_index = title_index if title_index is not None else 0

    selected_short = shorts[safe_short_index] if 0 <= safe_short_index < len(shorts) else _fallback_short()
    selected_title = titles[safe_title_index] if 0 <= safe_title_index < len(titles) else "Title candidate unavailable."
    description = str(candidates["description"]) or "Description draft unavailable."
    tags = str(candidates["tags"]) or "Tags unavailable."
    notes = str(candidates["notes"]) or "Upload notes unavailable."

    selected_short_path = selected_dir / "selected_short.json"
    selected_short_path.write_text(json.dumps(selected_short, ensure_ascii=False, indent=2), encoding="utf-8")
    selected_title_path = _write_text(selected_dir / "selected_title.txt", selected_title)
    selected_description_path = _write_text(selected_dir / "selected_description.txt", description)
    selected_tags_path = _write_text(selected_dir / "selected_tags.txt", tags)
    selected_notes_path = _write_text(selected_dir / "selected_upload_notes.txt", notes)

    summary = (
        "# Selected Draft\n\n"
        "## Selected Short\n"
        f"{_selected_short_text(selected_short)}\n\n"
        "## Selected Title\n"
        f"{selected_title}\n\n"
        "## Description\n"
        f"{description}\n\n"
        "## Tags\n"
        f"{tags}\n\n"
        "## Upload Notes\n"
        f"{notes}\n\n"
        "## Human Decision\n"
        "最終判断はユーザーが行います。\n"
        "自動投稿はしていません。\n"
    )
    summary_path = _write_text(selected_dir / "selected_summary.md", summary)

    if log:
        if short_index is not None:
            log(LOG_TEXT["selected_short"])
        if title_index is not None:
            log(LOG_TEXT["selected_title"])
        log(LOG_TEXT["selected_draft_created"])
        log(LOG_TEXT["selected_human_decision"])

    return {
        "status": "COMPLETED",
        "package_dir": str(package_dir),
        "selected_dir": str(selected_dir),
        "selected_short_index": safe_short_index,
        "selected_title_index": safe_title_index,
        "used_default": used_default,
        "selected_short": selected_short,
        "selected_title": selected_title,
        "written": [
            str(selected_short_path),
            str(selected_title_path),
            str(selected_description_path),
            str(selected_tags_path),
            str(selected_notes_path),
            str(summary_path),
        ],
    }
