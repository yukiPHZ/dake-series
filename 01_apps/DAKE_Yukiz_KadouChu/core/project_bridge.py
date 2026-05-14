from __future__ import annotations

import json
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

from core.app_config import app_root
from core.ollama_client import generate_ollama_text

LogCallback = Callable[[str], None]

BGM_EXTENSIONS = {".mp3", ".wav"}

BRIDGE_TEXT = {
    "notes_unavailable": "Project notes unavailable.",
    "no_project_boxes": "No Project Boxes Found",
    "template_title": "深夜、まだ作ってる。",
    "template_mood": "quiet midnight work",
    "template_scene": "静かなタイピング / 深夜の机 / まだ作ってる。",
    "template_note": "Ollama unavailable. Template bridge metadata was used.",
}


def _app_dir() -> Path:
    root = app_root()
    return root.parent if root.name.lower() == "dist" else root


def otooku_projects_dir() -> Path:
    return _app_dir().parent / "DAKE_Music_Otooku" / "data" / "outputs" / "projects"


def _read_text(path: Path) -> str:
    if not path.exists() or not path.is_file():
        return ""
    return path.read_text(encoding="utf-8", errors="replace")


def _parse_sections(text: str) -> dict[str, list[str]]:
    sections: dict[str, list[str]] = {}
    current = ""
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if line.endswith(":") and len(line) <= 48:
            current = line[:-1].strip()
            sections.setdefault(current, [])
            continue
        if current and line:
            sections.setdefault(current, []).append(line)
    return sections


def _first(sections: dict[str, list[str]], key: str, fallback: str = "") -> str:
    values = sections.get(key, [])
    return values[0] if values else fallback


def _clean_lines(lines: list[str], fallback: list[str] | None = None) -> list[str]:
    cleaned = [line.strip("- ").strip() for line in lines if line.strip("- ").strip()]
    return cleaned or list(fallback or [])


def _list_bgm_files(project_root: Path) -> list[Path]:
    bgm_dir = project_root / "bgm"
    if not bgm_dir.exists():
        return []
    return sorted(
        [path for path in bgm_dir.iterdir() if path.is_file() and path.suffix.lower() in BGM_EXTENSIONS],
        key=lambda path: path.name.lower(),
    )


def list_project_boxes(base_dir: Path | None = None) -> list[dict[str, str]]:
    projects_root = base_dir or otooku_projects_dir()
    if not projects_root.exists():
        return []
    boxes = [path for path in projects_root.iterdir() if path.is_dir()]
    boxes.sort(key=lambda path: path.stat().st_mtime, reverse=True)
    return [{"name": path.name, "root": str(path)} for path in boxes]


def read_project_box(project_root: Path) -> dict[str, Any]:
    root = project_root.resolve()
    notes_path = root / "notes" / "project_notes.txt"
    notes_text = _read_text(notes_path)
    sections = _parse_sections(notes_text)
    bgm_files = _list_bgm_files(root)

    suggested_use = _clean_lines(
        sections.get("Suggested Use", []),
        ["quiet work", "Shorts background", "late night build"],
    )
    shorts_direction = _clean_lines(
        sections.get("Shorts Direction", []),
        ["静かなタイピング", "深夜の机", "まだ作ってる。"],
    )
    selected_bgm = _first(sections, "BGM", bgm_files[0].name if bgm_files else "")

    return {
        "name": root.name,
        "root": str(root),
        "notes_path": str(notes_path) if notes_path.exists() else "",
        "notes_available": bool(notes_text),
        "notes_text": notes_text,
        "preset": _first(sections, "Selected Preset", "--"),
        "suggested_title": _first(sections, "Suggested Title", BRIDGE_TEXT["template_title"]),
        "mood": _first(sections, "Mood", BRIDGE_TEXT["template_mood"]),
        "suggested_use": suggested_use,
        "shorts_direction": shorts_direction,
        "selected_bgm": selected_bgm,
        "bgm_files": [{"name": path.name, "path": str(path)} for path in bgm_files],
    }


def _unique_dest_path(directory: Path, filename: str) -> Path:
    safe_name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", filename).strip(" .")
    safe_name = safe_name or "bgm.mp3"
    candidate = directory / safe_name
    if not candidate.exists():
        return candidate
    stem = candidate.stem
    suffix = candidate.suffix
    for index in range(1, 1000):
        next_candidate = directory / f"{stem}_{index:02d}{suffix}"
        if not next_candidate.exists():
            return next_candidate
    raise RuntimeError("Could not create a unique BGM filename.")


def add_bgm_to_video_box(package_dir: Path, bgm_path: Path, log: LogCallback | None = None) -> dict[str, Any]:
    source = bgm_path.resolve()
    if not source.exists() or source.suffix.lower() not in BGM_EXTENSIONS:
        return {
            "status": "FAILED",
            "package_dir": str(package_dir),
            "selected_dir": str(package_dir / "selected"),
            "copied_bgm": "",
            "message": "Selected BGM was not found.",
        }

    selected_dir = package_dir / "selected"
    bgm_dir = selected_dir / "bgm"
    bgm_dir.mkdir(parents=True, exist_ok=True)
    dest = _unique_dest_path(bgm_dir, source.name)
    shutil.copy2(source, dest)
    if log:
        log("BGM copied.")
    return {
        "status": "COMPLETED",
        "package_dir": str(package_dir),
        "selected_dir": str(selected_dir),
        "copied_bgm": str(dest),
        "message": "BGM copied to selected/bgm.",
    }


def _metadata_template(project: dict[str, Any], bgm_name: str) -> str:
    suggested_use = [str(item) for item in project.get("suggested_use", [])][:5]
    shorts_direction = [str(item) for item in project.get("shorts_direction", [])][:5]
    lines = [
        "# Project Bridge Metadata Draft",
        "",
        "Project:",
        str(project.get("name") or "--"),
        "",
        "Selected Preset:",
        str(project.get("preset") or "--"),
        "",
        "Suggested Title:",
        str(project.get("suggested_title") or BRIDGE_TEXT["template_title"]),
        "",
        "Suggested Mood:",
        str(project.get("mood") or BRIDGE_TEXT["template_mood"]),
        "",
        "BGM:",
        bgm_name or str(project.get("selected_bgm") or "--"),
        "",
        "Suggested Use:",
        *[f"- {line}" for line in suggested_use],
        "",
        "Shorts Direction:",
        *[f"- {line}" for line in shorts_direction],
        "",
        "Safety:",
        "- This app does not upload automatically.",
        "- Original audio and video files are not modified.",
    ]
    return "\n".join(lines).strip() + "\n"


def _ollama_prompt(project: dict[str, Any], bgm_name: str) -> str:
    suggested_use = "\n".join(f"- {item}" for item in project.get("suggested_use", [])[:5])
    shorts_direction = "\n".join(f"- {item}" for item in project.get("shorts_direction", [])[:5])
    notes_excerpt = str(project.get("notes_text") or "")[:1200]
    return (
        "You are the local assistant brain for a quiet video production console.\n"
        "The user is bridging a BGM Project Box into a YouTube video package.\n"
        "Do not act like a DAW. Do not propose automatic upload.\n"
        "Return concise Japanese notes with these labels: editing_mood, suggested_scene, shorts_direction, title_direction.\n\n"
        f"Project: {project.get('name')}\n"
        f"Preset: {project.get('preset')}\n"
        f"BGM: {bgm_name}\n"
        f"Mood: {project.get('mood')}\n"
        f"Suggested Use:\n{suggested_use}\n\n"
        f"Shorts Direction:\n{shorts_direction}\n\n"
        f"Project notes excerpt:\n{notes_excerpt}\n"
    )


def generate_bridge_metadata_draft(
    package_dir: Path,
    project: dict[str, Any],
    bgm_path: Path | None,
    ollama_ready: bool,
) -> dict[str, Any]:
    selected_dir = package_dir / "selected"
    upload_dir = selected_dir / "upload"
    upload_dir.mkdir(parents=True, exist_ok=True)
    metadata_path = upload_dir / "metadata_draft.txt"
    bgm_name = bgm_path.name if bgm_path else str(project.get("selected_bgm") or "")

    text = _metadata_template(project, bgm_name)
    used_ollama = False
    ollama_model = ""
    ollama_reason = ""
    if ollama_ready:
        response = generate_ollama_text(_ollama_prompt(project, bgm_name), timeout=45)
        used_ollama = bool(response.get("ok"))
        ollama_model = str(response.get("model") or "")
        ollama_reason = str(response.get("reason") or "")
        if used_ollama:
            text += "\n## Ollama Assistant Proposal\n\n"
            text += f"Model: {ollama_model}\n\n"
            text += str(response.get("text") or "").strip() + "\n"
        else:
            text += "\n## Assistant Template\n\n"
            text += f"{BRIDGE_TEXT['template_note']} {ollama_reason}\n"
            text += f"editing_mood: {BRIDGE_TEXT['template_mood']}\n"
            text += f"suggested_scene: {BRIDGE_TEXT['template_scene']}\n"
            text += "title_direction: 深夜、まだ作ってる。 / 静かな余熱 / 止まらず作る。\n"
    else:
        text += "\n## Assistant Template\n\n"
        text += BRIDGE_TEXT["template_note"] + "\n"
        text += f"editing_mood: {BRIDGE_TEXT['template_mood']}\n"
        text += f"suggested_scene: {BRIDGE_TEXT['template_scene']}\n"
        text += "title_direction: 深夜、まだ作ってる。 / 静かな余熱 / 止まらず作る。\n"

    text += f"\nGenerated At: {datetime.now().isoformat(timespec='seconds')}\n"
    metadata_path.write_text(text, encoding="utf-8")
    return {
        "status": "COMPLETED",
        "package_dir": str(package_dir),
        "selected_dir": str(selected_dir),
        "metadata_path": str(metadata_path),
        "project": str(project.get("name") or ""),
        "preset": str(project.get("preset") or ""),
        "bgm": bgm_name,
        "used_ollama": used_ollama,
        "ollama_model": ollama_model,
        "ollama_reason": ollama_reason,
    }
