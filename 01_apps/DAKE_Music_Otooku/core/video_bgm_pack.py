# -*- coding: utf-8 -*-
from __future__ import annotations

import re
import shutil
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

from .app_config import OLLAMA_BASE_URL, load_ollama_model_name
from .prompt_builder import MusicDirection


AUDIO_EXTENSIONS = {".mp3", ".wav"}
VIDEO_BGM_CATEGORIES = ("shorts", "long", "ambient", "work")


@dataclass(frozen=True)
class VideoBgmSuggestion:
    shorts_ideas: tuple[str, ...]
    long_ideas: tuple[str, ...]
    atmosphere: str
    scenes: tuple[str, ...]
    source: str = "template"
    response_time: float | None = None
    error: str = ""


@dataclass
class VideoBgmPackResult:
    success: bool
    root: Path
    copied_files: list[Path] = field(default_factory=list)
    suggestion: VideoBgmSuggestion | None = None
    messages: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)


def _clean_line(line: str) -> str:
    return line.strip().lstrip("-•* ").strip()


def _dedupe_lines(lines: list[str], fallback: tuple[str, ...]) -> tuple[str, ...]:
    cleaned: list[str] = []
    seen: set[str] = set()
    for line in [*lines, *fallback]:
        value = _clean_line(line)
        if not value or value.lower() in seen:
            continue
        cleaned.append(value)
        seen.add(value.lower())
        if len(cleaned) >= 6:
            break
    return tuple(cleaned or fallback)


def _parse_video_suggestion(text: str, direction: MusicDirection, source: str, response_time: float) -> VideoBgmSuggestion:
    label_map = {
        "shorts": "shorts",
        "shorts ideas": "shorts",
        "shorts向け用途": "shorts",
        "long": "long",
        "long ideas": "long",
        "long動画向け用途": "long",
        "atmosphere": "atmosphere",
        "空気感": "atmosphere",
        "scenes": "scenes",
        "recommended scenes": "scenes",
        "推奨シーン": "scenes",
        "suggested use": "scenes",
        "use": "scenes",
    }
    sections: dict[str, list[str]] = {"shorts": [], "long": [], "atmosphere": [], "scenes": []}
    current_key: str | None = None
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        match = re.match(r"^([A-Za-z _]+|[一-龥ぁ-んァ-ヶー]+):\s*(.*)$", line)
        if match:
            label = match.group(1).strip().lower()
            mapped = label_map.get(label)
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

    fallback = fallback_video_suggestion(direction)
    atmosphere = " / ".join(_clean_line(line) for line in sections["atmosphere"] if _clean_line(line))
    return VideoBgmSuggestion(
        shorts_ideas=_dedupe_lines(sections["shorts"], fallback.shorts_ideas),
        long_ideas=_dedupe_lines(sections["long"], fallback.long_ideas),
        atmosphere=atmosphere or fallback.atmosphere,
        scenes=_dedupe_lines(sections["scenes"], fallback.scenes),
        source=source,
        response_time=response_time,
    )


def fallback_video_suggestion(direction: MusicDirection, error: str = "") -> VideoBgmSuggestion:
    return VideoBgmSuggestion(
        shorts_ideas=(
            "「深夜、まだ作ってる。」",
            "「止まらず作る。」",
            "「静かな作業机。」",
        ),
        long_ideas=(
            "1時間作業配信",
            "holiday-jinja ambient",
            "ミシンとコード",
            "夜の制作ログ",
        ),
        atmosphere=direction.mood or "quiet ambient",
        scenes=(
            direction.usage_idea,
            "静かなShorts",
            "作業配信待機",
            "BORINEF背景",
        ),
        source="template",
        error=error,
    )


def generate_video_bgm_suggestion(
    direction: MusicDirection,
    available_models: tuple[str, ...] = (),
    base_url: str = OLLAMA_BASE_URL,
    timeout: int = 60,
) -> VideoBgmSuggestion:
    try:
        import requests
    except Exception as exc:
        return fallback_video_suggestion(direction, f"requests import failed: {exc}")

    model = available_models[0] if available_models else load_ollama_model_name()
    prompt = (
        "You are a quiet local assistant for the DAKE app 音を置く.\n"
        "This app creates video BGM material packs, not finished songs and not video editing timelines.\n"
        "Suggest practical uses for this sound under Shorts, long videos, ambient scenes, and work scenes.\n"
        "Keep it quiet, minimal, useful, and concise. Avoid famous artists, existing songs, vocals, and hype.\n"
        "Return exactly these labels. Use short Japanese phrases or compact English when useful.\n\n"
        "Shorts:\n"
        "Long:\n"
        "Atmosphere:\n"
        "Scenes:\n\n"
        "Sound direction:\n"
        f"Mood: {direction.mood}\n"
        f"BPM: {direction.bpm}\n"
        f"Texture: {direction.texture}\n"
        f"Instruments: {direction.instrumentation}\n"
        f"Usage Idea: {direction.usage_idea}\n"
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
        return _parse_video_suggestion(body, direction, f"ollama:{model}", response_time)
    except Exception as exc:
        return fallback_video_suggestion(direction, str(exc))


def _loop_files(loop_dir: Path) -> list[Path]:
    if not loop_dir.exists():
        return []
    return sorted(
        path
        for path in loop_dir.iterdir()
        if path.is_file() and path.suffix.lower() in AUDIO_EXTENSIONS
    )


def _matches(name: str, needles: tuple[str, ...]) -> bool:
    return any(needle in name for needle in needles)


def _copy_category(
    files: list[Path],
    destination: Path,
    copied_files: list[Path],
    predicate: Callable[[str], bool],
    fallback_predicate: Callable[[str], bool] | None = None,
) -> None:
    matches = [path for path in files if predicate(path.name.lower())]
    if not matches and fallback_predicate:
        matches = [path for path in files if fallback_predicate(path.name.lower())]
    if not matches and files:
        matches = files[:1]

    destination.mkdir(parents=True, exist_ok=True)
    for source in matches:
        target = destination / source.name
        shutil.copy2(source, target)
        copied_files.append(target)


def _write_note(path: Path, lines: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines).strip() + "\n", encoding="utf-8")


def _relative(path: Path, root: Path) -> str:
    try:
        return path.relative_to(root).as_posix()
    except ValueError:
        return path.as_posix()


def _write_video_notes(
    pack_root: Path,
    direction: MusicDirection,
    suggestion: VideoBgmSuggestion,
    copied_files: list[Path],
    messages: list[str],
    errors: list[str],
) -> None:
    notes_dir = pack_root / "notes"
    _write_note(
        notes_dir / "usage_note.txt",
        [
            "Mood:",
            direction.mood,
            "",
            "Texture:",
            direction.texture,
            "",
            "Suggested Use:",
            *suggestion.scenes,
            "",
            "Atmosphere:",
            suggestion.atmosphere,
            "",
            f"Source: {suggestion.source}",
        ],
    )
    _write_note(notes_dir / "shorts_ideas.txt", [f"- {line}" for line in suggestion.shorts_ideas])
    _write_note(notes_dir / "long_video_ideas.txt", [f"- {line}" for line in suggestion.long_ideas])

    log_lines = [
        "# Video BGM Pack Export Log",
        "",
        f"Suggestion Source: {suggestion.source}",
    ]
    if suggestion.response_time is not None:
        log_lines.append(f"Response Time: {suggestion.response_time:.2f}s")
    if suggestion.error:
        log_lines.append(f"Fallback Reason: {suggestion.error}")
    log_lines.extend(["", "Copied Files:"])
    log_lines.extend(f"- {_relative(path, pack_root)}" for path in copied_files)
    if messages:
        log_lines.extend(["", "Messages:"])
        log_lines.extend(f"- {message}" for message in messages)
    if errors:
        log_lines.extend(["", "Errors:"])
        log_lines.extend(f"- {error}" for error in errors)
    _write_note(pack_root / "export_log.txt", log_lines)


def export_video_bgm_pack(
    project_root: Path,
    direction: MusicDirection,
    available_models: tuple[str, ...] = (),
) -> VideoBgmPackResult:
    project_root = Path(project_root)
    pack_root = project_root / "video_bgm_pack"
    result = VideoBgmPackResult(success=False, root=pack_root)

    bgm_root = pack_root / "bgm"
    for category in VIDEO_BGM_CATEGORIES:
        (bgm_root / category).mkdir(parents=True, exist_ok=True)
    (pack_root / "notes").mkdir(parents=True, exist_ok=True)

    suggestion = generate_video_bgm_suggestion(direction, available_models)
    result.suggestion = suggestion

    loop_dir = project_root / "audio" / "loop_pack"
    files = _loop_files(loop_dir)
    if not files:
        result.errors.append("loop_pack folder has no mp3/wav files")
        _write_video_notes(pack_root, direction, suggestion, result.copied_files, result.messages, result.errors)
        return result

    _copy_category(
        files,
        bgm_root / "shorts",
        result.copied_files,
        lambda name: "30s" in name,
        lambda name: name.endswith(".mp3"),
    )
    _copy_category(
        files,
        bgm_root / "long",
        result.copied_files,
        lambda name: "180s" in name,
        lambda name: "60s" in name,
    )
    _copy_category(
        files,
        bgm_root / "ambient",
        result.copied_files,
        lambda name: _matches(name, ("ambient", "quiet", "shrine", "midnight")),
        lambda name: "60s" in name,
    )
    _copy_category(
        files,
        bgm_root / "work",
        result.copied_files,
        lambda name: _matches(name, ("work", "borinef")),
        lambda name: "60s" in name,
    )

    result.messages.append("Loop Pack copied into video BGM categories")
    result.success = bool(result.copied_files)
    _write_video_notes(pack_root, direction, suggestion, result.copied_files, result.messages, result.errors)
    return result
