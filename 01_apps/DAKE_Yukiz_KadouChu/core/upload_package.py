from __future__ import annotations

import json
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

from core.memory_store import memory_summary_path
from core.ollama_client import generate_ollama_text

LogCallback = Callable[[str], None]

VIDEO_OUTPUTS = [
    ("selected/horizontal_video.mp4", "horizontal_video.mp4"),
    ("selected/smart_horizontal_edit.mp4", "smart_horizontal_edit.mp4"),
    ("selected/horizontal_edit.mp4", "horizontal_edit.mp4"),
]

SHORT_OUTPUTS = [
    ("selected/short_vertical_1080x1920.mp4", "short_vertical_1080x1920.mp4"),
]

UPLOAD_LOG_TEXT = {
    "start": "補助脳：投稿前セットを整えています。",
    "assets": "補助脳：Shortsと横動画をまとめています。",
    "ready": "補助脳：YouTubeへ持っていける形にしました。",
}


def _read_text(path: Path, limit: int | None = None) -> str:
    if not path.exists() or not path.is_file():
        return ""
    text = path.read_text(encoding="utf-8", errors="replace").strip()
    return text[:limit] if limit is not None else text


def _read_json(path: Path) -> Any:
    if not path.exists() or not path.is_file():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8", errors="replace"))
    except Exception:
        return None


def _write_text(path: Path, text: str) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text.rstrip() + "\n", encoding="utf-8")
    return path


def _copy_file(source: Path, destination: Path) -> Path | None:
    if not source.exists() or not source.is_file():
        return None
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, destination)
    return destination


def _unique_destination(destination: Path) -> Path:
    if not destination.exists():
        return destination
    stem = destination.stem
    suffix = destination.suffix
    for index in range(1, 100):
        candidate = destination.with_name(f"{stem}_{index:02d}{suffix}")
        if not candidate.exists():
            return candidate
    return destination.with_name(f"{stem}_{datetime.now().strftime('%H%M%S')}{suffix}")


def _nonempty_lines(text: str) -> list[str]:
    lines: list[str] = []
    for raw in text.splitlines():
        line = raw.strip().lstrip("-").strip()
        if line:
            lines.append(line)
    return lines


def _section_lines(text: str, headings: set[str]) -> list[str]:
    lines = text.splitlines()
    capture = False
    result: list[str] = []
    for raw in lines:
        line = raw.strip()
        if line.startswith("## "):
            heading = line.lstrip("#").strip().lower()
            capture = heading in headings
            continue
        if capture:
            if line.startswith("#"):
                break
            cleaned = line.lstrip("-").strip()
            if cleaned:
                result.append(cleaned)
    return result


def _first_title_from_recommendation(package_dir: Path) -> str:
    recommendation = _read_text(package_dir / "assistant_recommendation.md", limit=5000)
    for heading in [{"suggested title direction"}, {"title direction"}]:
        lines = _section_lines(recommendation, heading)
        if lines:
            return lines[0]
    return ""


def _resolve_title(package_dir: Path) -> tuple[str, str]:
    selected_title = _read_text(package_dir / "selected" / "selected_title.txt")
    if selected_title:
        return selected_title.splitlines()[0].strip(), "selected_title.txt"

    title_ideas = _nonempty_lines(_read_text(package_dir / "metadata" / "title_ideas.txt"))
    if title_ideas:
        return title_ideas[0], "metadata/title_ideas.txt"

    recommendation_title = _first_title_from_recommendation(package_dir)
    if recommendation_title:
        return recommendation_title, "assistant_recommendation.md"

    return "Title candidate unavailable.", "fallback"


def _resolve_text(package_dir: Path, selected_name: str, metadata_name: str, fallback: str) -> tuple[str, str]:
    selected = _read_text(package_dir / "selected" / selected_name)
    if selected:
        return selected, f"selected/{selected_name}"
    metadata = _read_text(package_dir / "metadata" / metadata_name)
    if metadata:
        return metadata, f"metadata/{metadata_name}"
    return fallback, "fallback"


def _copy_video_outputs(package_dir: Path, upload_dir: Path) -> tuple[list[Path], list[str]]:
    copied: list[Path] = []
    missing: list[str] = []
    video_dir = upload_dir / "video"
    for relative_source, destination_name in VIDEO_OUTPUTS:
        source = package_dir / relative_source
        destination = video_dir / destination_name
        result = _copy_file(source, destination)
        if result is None:
            missing.append(relative_source)
        else:
            copied.append(result)
    return copied, missing


def _copy_short_outputs(package_dir: Path, upload_dir: Path) -> tuple[list[Path], list[str]]:
    copied: list[Path] = []
    missing: list[str] = []
    shorts_dir = upload_dir / "shorts"
    for relative_source, destination_name in SHORT_OUTPUTS:
        source = package_dir / relative_source
        destination = shorts_dir / destination_name
        result = _copy_file(source, destination)
        if result is None:
            missing.append(relative_source)
        else:
            copied.append(result)

    pack_dir = package_dir / "selected" / "shorts_pack"
    pack_files = sorted(pack_dir.glob("*.mp4")) if pack_dir.exists() else []
    if not pack_files:
        missing.append("selected/shorts_pack/*.mp4")
    for source in pack_files:
        result = _copy_file(source, shorts_dir / source.name)
        if result is not None:
            copied.append(result)
    return copied, missing


def _copy_thumbnails(package_dir: Path, upload_dir: Path) -> list[Path]:
    copied: list[Path] = []
    thumbnail_dir = upload_dir / "thumbnails"
    sources = [package_dir / "selected" / "thumbnails", package_dir / "thumbnails"]
    for source_dir in sources:
        if not source_dir.exists() or not source_dir.is_dir():
            continue
        for source in sorted(source_dir.iterdir()):
            if not source.is_file():
                continue
            destination = _unique_destination(thumbnail_dir / source.name)
            result = _copy_file(source, destination)
            if result is not None:
                copied.append(result)
    title_match = _read_json(package_dir / "selected" / "title_match" / "title_match.json")
    best_pair = title_match.get("best_pair") if isinstance(title_match, dict) else {}
    best_thumbnail = str(best_pair.get("thumbnail") or "") if isinstance(best_pair, dict) else ""
    best_source = package_dir / "selected" / "thumbnails" / best_thumbnail
    if best_thumbnail and best_source.exists():
        result = _copy_file(best_source, thumbnail_dir / "best_thumbnail.png")
        if result is not None:
            copied.append(result)
    return copied


def _ollama_refine_metadata(
    package_dir: Path,
    title: str,
    description: str,
    review: str,
    recommendation: str,
    memory: str,
) -> tuple[str, str, bool]:
    prompt = (
        "You are the local assistant brain for a quiet production console.\n"
        "Prepare a compact YouTube Studio handoff description and one checklist note.\n"
        "Do not upload, do not mention automation as available, do not write a long article.\n"
        "Keep Japanese concise, calm, and practical.\n"
        "Return JSON only with keys: final_description, checklist_note.\n\n"
        f"Title:\n{title[:300]}\n\n"
        f"Existing description:\n{description[:1800]}\n\n"
        f"Assistant review:\n{review[:1000]}\n\n"
        f"Recommendation:\n{recommendation[:1000]}\n\n"
        f"Memory summary:\n{memory[:1000]}\n\n"
        f"Package: {package_dir.name}\n"
    )
    response = generate_ollama_text(prompt, timeout=45)
    if not response.get("ok"):
        return description, "", False
    text = str(response.get("text") or "").strip()
    start = text.find("{")
    end = text.rfind("}")
    if start == -1 or end == -1 or end <= start:
        return description, "", False
    try:
        payload = json.loads(text[start : end + 1])
    except Exception:
        return description, "", False
    final_description = str(payload.get("final_description") or "").strip()
    checklist_note = str(payload.get("checklist_note") or "").strip()
    return final_description or description, checklist_note, bool(final_description or checklist_note)


def _build_checklist(
    video_files: list[Path],
    short_files: list[Path],
    thumbnail_files: list[Path],
    checklist_note: str,
) -> str:
    video_names = [path.name for path in video_files]
    short_names = [path.name for path in short_files]
    thumbnail_names = [path.name for path in thumbnail_files]
    lines = [
        "# Upload Checklist",
        "",
        "## Video",
        "- [ ] horizontal video checked",
        "- [ ] shorts checked",
        "- [ ] audio level checked",
    ]
    if video_names:
        lines.extend(f"  - {name}" for name in video_names)
    else:
        lines.append("  - No horizontal videos available")
    if short_names:
        lines.extend(f"  - {name}" for name in short_names)
    else:
        lines.append("  - No shorts available")
    lines.extend(
        [
            "",
            "## Metadata",
            "- [ ] title checked",
            "- [ ] description checked",
            "- [ ] tags checked",
            "",
            "## Thumbnail",
            "- [ ] thumbnail prepared",
        ]
    )
    if thumbnail_names:
        lines.extend(f"  - {name}" for name in thumbnail_names)
    else:
        lines.append("  - No thumbnails available")
    lines.extend(
        [
            "",
            "## Publish",
            "- [ ] YouTube Studio upload",
            "- [ ] final confirmation",
            "",
            "## Assistant Note",
            checklist_note or "最終判断はユーザーが行います。",
            "自動公開はしていません。",
        ]
    )
    return "\n".join(lines)


def _write_log(
    log_path: Path,
    package_dir: Path,
    copied_video: list[Path],
    copied_shorts: list[Path],
    copied_thumbnails: list[Path],
    metadata_sources: dict[str, str],
    missing: list[str],
    used_ollama: bool,
) -> Path:
    lines = [
        "Upload Package Export Log",
        f"created_at: {datetime.now().isoformat(timespec='seconds')}",
        f"package_dir: {package_dir}",
        f"used_ollama: {used_ollama}",
        "",
        "video:",
        *[f"- {path}" for path in copied_video],
        "",
        "shorts:",
        *[f"- {path}" for path in copied_shorts],
        "",
        "thumbnails:",
        *[f"- {path}" for path in copied_thumbnails],
        "",
        "metadata_sources:",
        *[f"- {key}: {value}" for key, value in metadata_sources.items()],
        "",
        "missing:",
        *[f"- {item}" for item in missing],
        "",
        "policy:",
        "- Copy only.",
        "- YouTube auto upload is disabled.",
        "- Final judgment belongs to the user.",
    ]
    return _write_text(log_path, "\n".join(lines))


def generate_upload_package(
    package_dir: Path,
    ollama_ready: bool = False,
    log: LogCallback | None = None,
) -> dict[str, Any]:
    selected_dir = package_dir / "selected"
    upload_dir = selected_dir / "upload_ready"
    metadata_dir = upload_dir / "metadata"
    logs_dir = upload_dir / "logs"
    for directory in [upload_dir / "video", upload_dir / "shorts", metadata_dir, upload_dir / "thumbnails", logs_dir]:
        directory.mkdir(parents=True, exist_ok=True)

    if log:
        log(UPLOAD_LOG_TEXT["start"])

    copied_video, missing_video = _copy_video_outputs(package_dir, upload_dir)
    if log:
        log(UPLOAD_LOG_TEXT["assets"])
    copied_shorts, missing_shorts = _copy_short_outputs(package_dir, upload_dir)
    copied_thumbnails = _copy_thumbnails(package_dir, upload_dir)

    title, title_source = _resolve_title(package_dir)
    description, description_source = _resolve_text(
        package_dir,
        "selected_description.txt",
        "description_draft.txt",
        "Description draft unavailable.",
    )
    tags, tags_source = _resolve_text(package_dir, "selected_tags.txt", "tags.txt", "Tags unavailable.")

    review = _read_text(package_dir / "assistant_review.md", limit=3000)
    recommendation = _read_text(package_dir / "assistant_recommendation.md", limit=3000)
    memory = _read_text(memory_summary_path(), limit=3000)

    checklist_note = "最終判断はユーザーが行います。"
    used_ollama = False
    if ollama_ready:
        description, checklist_note, used_ollama = _ollama_refine_metadata(
            package_dir,
            title,
            description,
            review,
            recommendation,
            memory,
        )

    final_title_path = _write_text(metadata_dir / "final_title.txt", title)
    final_description_path = _write_text(metadata_dir / "final_description.txt", description)
    final_tags_path = _write_text(metadata_dir / "final_tags.txt", tags)
    checklist_path = _write_text(
        metadata_dir / "upload_checklist.md",
        _build_checklist(copied_video, copied_shorts, copied_thumbnails, checklist_note),
    )

    copied_metadata: list[Path] = [final_title_path, final_description_path, final_tags_path, checklist_path]
    review_copy = _copy_file(package_dir / "assistant_review.md", metadata_dir / "assistant_review.md")
    recommendation_copy = _copy_file(package_dir / "assistant_recommendation.md", metadata_dir / "assistant_recommendation.md")
    memory_copy = _copy_file(memory_summary_path(), metadata_dir / "memory_summary.md")
    title_match_md_copy = _copy_file(package_dir / "selected" / "title_match" / "title_match.md", metadata_dir / "title_match.md")
    title_match_json_copy = _copy_file(package_dir / "selected" / "title_match" / "title_match.json", metadata_dir / "title_match.json")
    for item in [review_copy, recommendation_copy, memory_copy, title_match_md_copy, title_match_json_copy]:
        if item is not None:
            copied_metadata.append(item)

    missing = missing_video + missing_shorts
    if review_copy is None:
        missing.append("assistant_review.md")
    if recommendation_copy is None:
        missing.append("assistant_recommendation.md")
    if memory_copy is None:
        missing.append("data/memory/memory_summary.md")
    if title_match_md_copy is None:
        missing.append("selected/title_match/title_match.md")
    if title_match_json_copy is None:
        missing.append("selected/title_match/title_match.json")
    if not copied_thumbnails:
        missing.append("thumbnails")

    metadata_sources = {
        "title": title_source,
        "description": description_source,
        "tags": tags_source,
    }
    log_path = _write_log(
        logs_dir / "upload_package_log.txt",
        package_dir,
        copied_video,
        copied_shorts,
        copied_thumbnails,
        metadata_sources,
        missing,
        used_ollama,
    )
    if log:
        log(UPLOAD_LOG_TEXT["ready"])

    return {
        "status": "COMPLETED",
        "package_dir": str(package_dir),
        "selected_dir": str(selected_dir),
        "upload_ready_dir": str(upload_dir),
        "video_count": len(copied_video),
        "shorts_count": len(copied_shorts),
        "thumbnail_count": len(copied_thumbnails),
        "metadata_count": len(copied_metadata),
        "copied_video": [str(path) for path in copied_video],
        "copied_shorts": [str(path) for path in copied_shorts],
        "copied_metadata": [str(path) for path in copied_metadata],
        "copied_thumbnails": [str(path) for path in copied_thumbnails],
        "missing": missing,
        "metadata_sources": metadata_sources,
        "final_title": title,
        "final_description_path": str(final_description_path),
        "final_tags_path": str(final_tags_path),
        "checklist_path": str(checklist_path),
        "log_path": str(log_path),
        "used_ollama": used_ollama,
        "message": "upload_ready package created.",
    }
