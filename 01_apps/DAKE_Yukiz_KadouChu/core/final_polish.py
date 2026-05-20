from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

from core.memory_store import memory_index_path, memory_summary_path
from core.ollama_client import generate_ollama_text

LogCallback = Callable[[str], None]

FINAL_POLISH_TEXT = {
    "start": "補助脳：投稿前の最終確認を整えています。",
    "upload": "補助脳：Upload Ready を確認しています。",
    "preview": "補助脳：プレビューウォールを整えています。",
    "ready": "補助脳：出したくなる状態に整えました。",
    "failed": "補助脳：Final Polish の生成に失敗しました。",
}

NOTE_FALLBACK = {
    "current_atmosphere": "quiet midnight work",
    "upload_feeling": "静かな作業感があります。",
    "shorts_balance": "INTRO / WORK / AFTERGLOW の流れが自然です。",
    "thumbnail_direction": "机と光の余熱感があります。",
    "assistant_note": "最後だけ、菊田さんが握ってください。",
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


def _write_json(path: Path, payload: Any) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def _nonempty_first_line(text: str) -> str:
    for raw in text.splitlines():
        line = raw.strip().lstrip("-").strip()
        if line and not line.startswith("#"):
            return line
    return ""


def _first_existing(paths: list[Path]) -> Path | None:
    for path in paths:
        if path.exists() and path.is_file():
            return path
    return None


def _title_match_payload(package_dir: Path) -> dict[str, Any]:
    payload = _read_json(package_dir / "selected" / "title_match" / "title_match.json")
    return payload if isinstance(payload, dict) else {}


def _best_pair(package_dir: Path) -> dict[str, str]:
    payload = _title_match_payload(package_dir)
    best = payload.get("best_pair")
    return {str(key): str(value) for key, value in best.items()} if isinstance(best, dict) else {}


def _resolve_title(package_dir: Path) -> tuple[str, str]:
    upload_title = _read_text(package_dir / "selected" / "upload_ready" / "metadata" / "final_title.txt")
    if upload_title:
        return _nonempty_first_line(upload_title), "selected/upload_ready/metadata/final_title.txt"

    best_title = _best_pair(package_dir).get("title", "").strip()
    if best_title:
        return best_title, "selected/title_match/title_match.json"

    selected_title = _read_text(package_dir / "selected" / "selected_title.txt")
    if selected_title:
        return _nonempty_first_line(selected_title), "selected/selected_title.txt"

    title_ideas = _read_text(package_dir / "metadata" / "title_ideas.txt")
    if title_ideas:
        return _nonempty_first_line(title_ideas), "metadata/title_ideas.txt"

    return "稼働中。", "fallback"


def _resolve_thumbnail(package_dir: Path) -> tuple[str, str, str]:
    upload_best = package_dir / "selected" / "upload_ready" / "thumbnails" / "best_thumbnail.png"
    if upload_best.exists():
        return upload_best.name, str(upload_best), "selected/upload_ready/thumbnails/best_thumbnail.png"

    best_thumbnail = _best_pair(package_dir).get("thumbnail", "").strip()
    title_match_thumb = package_dir / "selected" / "thumbnails" / best_thumbnail
    if best_thumbnail and title_match_thumb.exists():
        return best_thumbnail, str(title_match_thumb), "selected/title_match/title_match.json"

    upload_thumbs = sorted((package_dir / "selected" / "upload_ready" / "thumbnails").glob("*.png"))
    if upload_thumbs:
        path = upload_thumbs[0]
        return path.name, str(path), "selected/upload_ready/thumbnails"

    local_thumbs = sorted((package_dir / "selected" / "thumbnails").glob("thumb_*.png"))
    if local_thumbs:
        path = local_thumbs[0]
        return path.name, str(path), "selected/thumbnails"

    return "--", "", "missing"


def _resolve_horizontal_video(package_dir: Path) -> tuple[str, str, str]:
    path = _first_existing(
        [
            package_dir / "selected" / "upload_ready" / "video" / "smart_horizontal_edit.mp4",
            package_dir / "selected" / "smart_horizontal_edit.mp4",
            package_dir / "selected" / "upload_ready" / "video" / "horizontal_video.mp4",
            package_dir / "selected" / "horizontal_video.mp4",
            package_dir / "selected" / "upload_ready" / "video" / "horizontal_edit.mp4",
            package_dir / "selected" / "horizontal_edit.mp4",
        ]
    )
    if path is None:
        return "--", "", "missing"
    return path.name, str(path), str(path)


def _shorts_files(package_dir: Path) -> list[Path]:
    upload_shorts = sorted((package_dir / "selected" / "upload_ready" / "shorts").glob("*.mp4"))
    if upload_shorts:
        return upload_shorts
    pack_shorts = sorted((package_dir / "selected" / "shorts_pack").glob("*.mp4"))
    if pack_shorts:
        return pack_shorts
    vertical = package_dir / "selected" / "short_vertical_1080x1920.mp4"
    return [vertical] if vertical.exists() else []


def _shorts_roles(package_dir: Path) -> list[str]:
    payload = _title_match_payload(package_dir)
    direction = payload.get("shorts_direction")
    if isinstance(direction, dict):
        roles = []
        for key, label in [("intro", "INTRO"), ("work", "WORK"), ("afterglow", "AFTERGLOW")]:
            value = str(direction.get(key) or "").strip()
            roles.append(f"{label}: {value}" if value else label)
        return roles
    return ["INTRO", "WORK", "AFTERGLOW"]


def _bgm_files(package_dir: Path) -> list[str]:
    bgm_dir = package_dir / "selected" / "bgm"
    names = [path.name for path in sorted(bgm_dir.glob("*")) if path.suffix.lower() in {".mp3", ".wav"}]
    if names:
        return names
    metadata = _read_text(package_dir / "selected" / "upload" / "metadata_draft.txt", limit=2000)
    for raw in metadata.splitlines():
        if raw.strip().lower().startswith("bgm:"):
            value = raw.split(":", 1)[1].strip()
            return [value] if value and value != "--" else []
    return []


def _memory_saved(package_dir: Path) -> bool:
    upload_memory = package_dir / "selected" / "upload_ready" / "metadata" / "memory_summary.md"
    if upload_memory.exists():
        return True
    index_path = memory_index_path()
    payload = _read_json(index_path)
    if not isinstance(payload, list):
        return False
    try:
        target = str(package_dir.resolve(strict=False)).lower()
    except Exception:
        target = str(package_dir).lower()
    for item in payload:
        if not isinstance(item, dict):
            continue
        source = str(item.get("source_package") or "")
        try:
            source = str(Path(source).resolve(strict=False))
        except Exception:
            pass
        if source.lower() == target:
            return True
    return False


def _build_check(package_dir: Path) -> dict[str, Any]:
    title, title_source = _resolve_title(package_dir)
    thumbnail, thumbnail_path, thumbnail_source = _resolve_thumbnail(package_dir)
    horizontal, horizontal_path, horizontal_source = _resolve_horizontal_video(package_dir)
    shorts = _shorts_files(package_dir)
    bgm = _bgm_files(package_dir)
    upload_dir = package_dir / "selected" / "upload_ready"
    upload_ready = (upload_dir / "metadata" / "upload_checklist.md").exists()
    return {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "package_name": package_dir.name,
        "title": title,
        "title_source": title_source,
        "thumbnail": thumbnail,
        "thumbnail_path": thumbnail_path,
        "thumbnail_source": thumbnail_source,
        "horizontal_video": horizontal,
        "horizontal_video_path": horizontal_path,
        "horizontal_video_source": horizontal_source,
        "shorts_count": len(shorts),
        "shorts": [path.name for path in shorts],
        "shorts_roles": _shorts_roles(package_dir),
        "bgm": bgm[0] if bgm else "--",
        "bgm_files": bgm,
        "memory_saved": _memory_saved(package_dir),
        "upload_ready": upload_ready,
        "upload_ready_dir": str(upload_dir) if upload_dir.exists() else "",
        "status": "COMPLETED",
    }


def _ollama_note(package_dir: Path, check: dict[str, Any]) -> tuple[dict[str, str], bool, str]:
    review = _read_text(package_dir / "assistant_review.md", limit=1200)
    recommendation = _read_text(package_dir / "assistant_recommendation.md", limit=1200)
    title_match = _read_text(package_dir / "selected" / "title_match" / "title_match.md", limit=1000)
    memory = _read_text(memory_summary_path(), limit=1000)
    upload_metadata = "\n\n".join(
        [
            _read_text(package_dir / "selected" / "upload_ready" / "metadata" / "final_title.txt", limit=300),
            _read_text(package_dir / "selected" / "upload_ready" / "metadata" / "final_description.txt", limit=1200),
            _read_text(package_dir / "selected" / "upload_ready" / "metadata" / "final_tags.txt", limit=500),
        ]
    )
    prompt = (
        "You are the local assistant brain for a quiet production console.\n"
        "Create a compact final pre-upload note. Do not upload. Do not over-write a long article.\n"
        "Return JSON only with keys: current_atmosphere, upload_feeling, shorts_balance, thumbnail_direction, assistant_note.\n"
        "Keep Japanese calm and concise.\n\n"
        f"Final check:\n{json.dumps(check, ensure_ascii=False)[:2000]}\n\n"
        f"Assistant review:\n{review}\n\n"
        f"Recommendation:\n{recommendation}\n\n"
        f"Title match:\n{title_match}\n\n"
        f"Memory:\n{memory}\n\n"
        f"Upload metadata:\n{upload_metadata[:2000]}\n"
    )
    response = generate_ollama_text(prompt, timeout=35)
    if not response.get("ok"):
        return dict(NOTE_FALLBACK), False, str(response.get("reason") or "")
    text = str(response.get("text") or "")
    start = text.find("{")
    end = text.rfind("}")
    if start == -1 or end == -1 or end <= start:
        return dict(NOTE_FALLBACK), False, "Ollama response was not JSON."
    try:
        payload = json.loads(text[start : end + 1])
    except Exception:
        return dict(NOTE_FALLBACK), False, "Ollama JSON parse failed."
    note = dict(NOTE_FALLBACK)
    for key in note:
        value = str(payload.get(key) or "").strip()
        if value:
            note[key] = value
    return note, True, str(response.get("model") or "")


def _build_note_markdown(note: dict[str, str], used_ollama: bool) -> str:
    return "\n".join(
        [
            "# Final Polish Note",
            "",
            "## Current Atmosphere",
            note.get("current_atmosphere", NOTE_FALLBACK["current_atmosphere"]),
            "",
            "## Upload Feeling",
            note.get("upload_feeling", NOTE_FALLBACK["upload_feeling"]),
            "",
            "## Shorts Balance",
            note.get("shorts_balance", NOTE_FALLBACK["shorts_balance"]),
            "",
            "## Thumbnail Direction",
            note.get("thumbnail_direction", NOTE_FALLBACK["thumbnail_direction"]),
            "",
            "## Assistant Note",
            note.get("assistant_note", NOTE_FALLBACK["assistant_note"]),
            "",
            "## Source",
            "Ollama" if used_ollama else "template fallback",
        ]
    )


def _write_log(path: Path, package_dir: Path, check: dict[str, Any], used_ollama: bool, reason: str) -> Path:
    lines = [
        "Final Polish Log",
        f"created_at: {datetime.now().isoformat(timespec='seconds')}",
        f"package_dir: {package_dir}",
        f"used_ollama: {used_ollama}",
        f"ollama_reason: {reason}",
        "",
        "final_check:",
        json.dumps(check, ensure_ascii=False, indent=2),
        "",
        "policy:",
        "- Preview only.",
        "- No YouTube upload.",
        "- Final judgment belongs to the user.",
    ]
    return _write_text(path, "\n".join(lines))


def generate_final_polish(
    package_dir: Path,
    ollama_ready: bool = False,
    log: LogCallback | None = None,
) -> dict[str, Any]:
    output_dir = package_dir / "selected" / "final_polish"
    output_dir.mkdir(parents=True, exist_ok=True)

    if log:
        log(FINAL_POLISH_TEXT["start"])
        log(FINAL_POLISH_TEXT["upload"])

    check = _build_check(package_dir)
    if log:
        log(FINAL_POLISH_TEXT["preview"])

    note = dict(NOTE_FALLBACK)
    used_ollama = False
    ollama_reason = ""
    if ollama_ready:
        note, used_ollama, ollama_reason = _ollama_note(package_dir, check)

    final_check_path = _write_json(output_dir / "final_check.json", check)
    note_path = _write_text(output_dir / "final_polish_note.md", _build_note_markdown(note, used_ollama))
    log_path = _write_log(output_dir / "final_polish_log.txt", package_dir, check, used_ollama, ollama_reason)

    if log:
        log(FINAL_POLISH_TEXT["ready"])

    return {
        "status": "COMPLETED",
        "package_dir": str(package_dir),
        "final_polish_dir": str(output_dir),
        "final_check_path": str(final_check_path),
        "final_polish_note_path": str(note_path),
        "log_path": str(log_path),
        "final_check": check,
        "note": note,
        "used_ollama": used_ollama,
        "ollama_reason": ollama_reason,
        "message": "Final Polish ready.",
    }


def read_final_check(package_dir: Path) -> dict[str, Any]:
    payload = _read_json(package_dir / "selected" / "final_polish" / "final_check.json")
    return payload if isinstance(payload, dict) else {}
