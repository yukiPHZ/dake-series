from __future__ import annotations

import json
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

from core.cli_checker import run_command
from core.ollama_client import generate_ollama_text
from core.selected_preview import find_source_video_path

try:  # OpenCV is optional at runtime.
    import cv2  # type: ignore
    import numpy as np  # type: ignore
except Exception:  # pragma: no cover - exercised on machines without opencv
    cv2 = None
    np = None

LogCallback = Callable[[str], None]

THUMBNAIL_SIZE = (1280, 720)
THUMBNAIL_COUNT = 5
THUMBNAIL_TEXT = {
    "start": "補助脳：サムネ候補を整理しています。",
    "search": "補助脳：静かな導入向きの画を探しています。",
    "ready": "補助脳：サムネ候補を出力しました。",
    "added": "補助脳：サムネを投稿前セットへ追加しました。",
    "failed": "補助脳：サムネ候補の生成に失敗しました。",
}

FALLBACK_DIRECTIONS = [
    ("quiet work", "静かなタイピング"),
    ("desk light", "机と光"),
    ("midnight coding", "静かな導入向きです"),
    ("afterglow", "余熱感が強いです"),
    ("calm process", "タイトルと相性が良いです"),
]


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


def _seconds_to_timecode(seconds: float) -> str:
    total = max(0, int(round(seconds)))
    hours = total // 3600
    minutes = (total % 3600) // 60
    secs = total % 60
    return f"{hours:02d}:{minutes:02d}:{secs:02d}"


def _title_for_package(package_dir: Path) -> str:
    for path in [
        package_dir / "selected" / "selected_title.txt",
        package_dir / "selected" / "upload_ready" / "metadata" / "final_title.txt",
        package_dir / "metadata" / "title_ideas.txt",
    ]:
        text = _read_text(path)
        for raw in text.splitlines():
            line = raw.strip().lstrip("-").strip()
            if line:
                return line
    return "稼働中。"


def _resolve_thumbnail_source(package_dir: Path, source_video_path: Path | None = None) -> tuple[Path | None, str]:
    selected_dir = package_dir / "selected"
    candidates = [
        (selected_dir / "smart_horizontal_edit.mp4", "selected/smart_horizontal_edit.mp4"),
        (selected_dir / "horizontal_video.mp4", "selected/horizontal_video.mp4"),
        (selected_dir / "horizontal_edit.mp4", "selected/horizontal_edit.mp4"),
    ]
    for path, label in candidates:
        if path.exists() and path.is_file():
            return path, label
    if source_video_path is not None and source_video_path.exists():
        return source_video_path, "source video"
    source = find_source_video_path(package_dir)
    if source is not None and source.exists():
        return source, "source video"
    return None, "unavailable"


def _candidate_ratios(count: int = 24) -> list[float]:
    seed = [
        0.06,
        0.10,
        0.16,
        0.23,
        0.31,
        0.37,
        0.44,
        0.53,
        0.59,
        0.66,
        0.73,
        0.79,
        0.86,
        0.92,
    ]
    golden = 0.61803398875
    value = 0.09
    while len(seed) < count:
        value = (value + golden) % 1.0
        if 0.05 <= value <= 0.94:
            seed.append(value)
    return sorted(set(round(item, 4) for item in seed))[:count]


def _read_frame(cap: Any, seconds: float) -> Any:
    if cv2 is None:
        return None
    cap.set(cv2.CAP_PROP_POS_MSEC, max(0.0, seconds) * 1000.0)
    ok, frame = cap.read()
    return frame if ok else None


def _score_frame(frame: Any, previous: Any | None = None) -> dict[str, float]:
    if cv2 is None or np is None or frame is None:
        return {"score": 0.0, "brightness": 0.0, "sharpness": 0.0, "motion": 0.0}
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    brightness = float(gray.mean())
    contrast = float(gray.std())
    sharpness = float(cv2.Laplacian(gray, cv2.CV_64F).var())
    edges = float(cv2.Canny(gray, 80, 160).mean() / 255.0)
    motion = 0.0
    if previous is not None:
        previous_gray = cv2.cvtColor(previous, cv2.COLOR_BGR2GRAY)
        if previous_gray.shape == gray.shape:
            motion = float(cv2.absdiff(gray, previous_gray).mean())

    brightness_score = max(0.0, 1.0 - abs(brightness - 118.0) / 118.0) * 34.0
    contrast_score = min(contrast / 70.0, 1.0) * 20.0
    sharp_score = min(sharpness / 420.0, 1.0) * 22.0
    edge_score = min(edges / 0.10, 1.0) * 12.0
    motion_score = min(motion / 36.0, 1.0) * 12.0
    penalty = 0.0
    if brightness < 24.0 or brightness > 232.0:
        penalty += 42.0
    if sharpness < 18.0:
        penalty += 22.0
    if motion > 85.0:
        penalty += 10.0
    score = max(0.0, brightness_score + contrast_score + sharp_score + edge_score + motion_score - penalty)
    return {
        "score": round(score, 3),
        "brightness": round(brightness, 3),
        "contrast": round(contrast, 3),
        "sharpness": round(sharpness, 3),
        "motion": round(motion, 3),
    }


def _fit_thumbnail(frame: Any) -> Any:
    if cv2 is None or frame is None:
        return frame
    target_w, target_h = THUMBNAIL_SIZE
    height, width = frame.shape[:2]
    if width <= 0 or height <= 0:
        return frame
    source_ratio = width / height
    target_ratio = target_w / target_h
    if 1.55 <= source_ratio <= 1.98:
        scale = max(target_w / width, target_h / height)
        resized = cv2.resize(frame, (int(width * scale), int(height * scale)), interpolation=cv2.INTER_AREA)
        y = max(0, (resized.shape[0] - target_h) // 2)
        x = max(0, (resized.shape[1] - target_w) // 2)
        return resized[y : y + target_h, x : x + target_w]

    bg_scale = max(target_w / width, target_h / height)
    background = cv2.resize(frame, (int(width * bg_scale), int(height * bg_scale)), interpolation=cv2.INTER_AREA)
    y = max(0, (background.shape[0] - target_h) // 2)
    x = max(0, (background.shape[1] - target_w) // 2)
    background = background[y : y + target_h, x : x + target_w]
    background = cv2.GaussianBlur(background, (35, 35), 0)

    fg_scale = min(target_w / width, target_h / height)
    foreground = cv2.resize(frame, (int(width * fg_scale), int(height * fg_scale)), interpolation=cv2.INTER_AREA)
    output = background.copy()
    y = max(0, (target_h - foreground.shape[0]) // 2)
    x = max(0, (target_w - foreground.shape[1]) // 2)
    output[y : y + foreground.shape[0], x : x + foreground.shape[1]] = foreground
    return output


def _select_samples(source: Path, limit: int = THUMBNAIL_COUNT) -> tuple[list[dict[str, Any]], float]:
    if cv2 is None:
        return [], 0.0
    cap = cv2.VideoCapture(str(source))
    if not cap.isOpened():
        return [], 0.0
    try:
        fps = float(cap.get(cv2.CAP_PROP_FPS) or 0.0)
        frame_count = float(cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0.0)
        duration = frame_count / fps if fps > 0 and frame_count > 0 else 0.0
        if duration <= 0:
            duration = 180.0
        rows: list[dict[str, Any]] = []
        for ratio in _candidate_ratios():
            seconds = max(0.0, min(duration - 0.25, duration * ratio))
            previous = _read_frame(cap, max(0.0, seconds - 0.35))
            frame = _read_frame(cap, seconds)
            if frame is None:
                continue
            score = _score_frame(frame, previous)
            if score["brightness"] < 18.0 or score["brightness"] > 238.0:
                continue
            rows.append({"seconds": seconds, "frame": frame, **score})
        rows.sort(key=lambda item: float(item.get("score") or 0.0), reverse=True)

        selected: list[dict[str, Any]] = []
        min_gap = max(2.0, min(20.0, duration / 12.0))
        for row in rows:
            seconds = float(row["seconds"])
            if all(abs(seconds - float(item["seconds"])) >= min_gap for item in selected):
                selected.append(row)
            if len(selected) >= limit:
                break
        if len(selected) < limit:
            for row in rows:
                if row not in selected:
                    selected.append(row)
                if len(selected) >= limit:
                    break
        return selected[:limit], duration
    finally:
        cap.release()


def _ollama_thumbnail_notes(package_dir: Path, title: str, count: int) -> tuple[list[dict[str, str]], bool]:
    review = _read_text(package_dir / "assistant_review.md", limit=1200)
    recommendation = _read_text(package_dir / "assistant_recommendation.md", limit=1200)
    memory = _read_text(package_dir / "selected" / "upload_ready" / "metadata" / "memory_summary.md", limit=1000)
    bridge = _read_text(package_dir / "selected" / "upload" / "metadata_draft.txt", limit=1000)
    prompt = (
        "You are the local assistant brain for a quiet video production console.\n"
        "Create compact thumbnail directions. Do not generate images. Do not design text layers.\n"
        "Return JSON only: {\"items\":[{\"direction\":\"quiet work\",\"reason\":\"静かなタイピング\",\"title_match\":\"...\"}]}.\n"
        f"Need {count} items.\n"
        f"Title: {title[:200]}\n"
        f"Review:\n{review}\n\nRecommendation:\n{recommendation}\n\nMemory:\n{memory}\n\nProject Bridge:\n{bridge}\n"
    )
    response = generate_ollama_text(prompt, timeout=35)
    if not response.get("ok"):
        return [], False
    text = str(response.get("text") or "")
    start = text.find("{")
    end = text.rfind("}")
    if start == -1 or end == -1 or end <= start:
        return [], False
    try:
        payload = json.loads(text[start : end + 1])
    except Exception:
        return [], False
    items = payload.get("items")
    if not isinstance(items, list):
        return [], False
    cleaned: list[dict[str, str]] = []
    for item in items[:count]:
        if not isinstance(item, dict):
            continue
        cleaned.append(
            {
                "direction": str(item.get("direction") or "").strip(),
                "reason": str(item.get("reason") or "").strip(),
                "title_match": str(item.get("title_match") or "").strip(),
            }
        )
    return cleaned, bool(cleaned)


def _decorate_candidate(index: int, candidate: dict[str, Any], title: str, notes: list[dict[str, str]]) -> dict[str, Any]:
    direction, reason = FALLBACK_DIRECTIONS[(index - 1) % len(FALLBACK_DIRECTIONS)]
    title_match = title
    if index - 1 < len(notes):
        note = notes[index - 1]
        direction = note.get("direction") or direction
        reason = note.get("reason") or reason
        title_match = note.get("title_match") or title_match
    return {
        "file": f"thumb_{index:02d}.png",
        "seconds": round(float(candidate.get("seconds") or 0.0), 3),
        "timecode": _seconds_to_timecode(float(candidate.get("seconds") or 0.0)),
        "direction": direction,
        "reason": reason,
        "title_match": title_match,
        "score": candidate.get("score", 0.0),
        "brightness": candidate.get("brightness", 0.0),
        "sharpness": candidate.get("sharpness", 0.0),
        "status": "candidate",
    }


def _ffprobe_command_for(ffmpeg_path: str | None) -> str | None:
    if not ffmpeg_path:
        return "ffprobe"
    path = Path(ffmpeg_path)
    if path.name.lower().startswith("ffmpeg"):
        sibling = path.with_name("ffprobe.exe" if path.suffix.lower() == ".exe" else "ffprobe")
        if sibling.exists():
            return str(sibling)
        if str(path) == ffmpeg_path and path.parent == Path("."):
            return "ffprobe"
    return "ffprobe"


def _probe_duration_for_fallback(source: Path, ffmpeg_path: str | None) -> float | None:
    ffprobe = _ffprobe_command_for(ffmpeg_path)
    if ffprobe:
        try:
            completed = run_command(
                [
                    ffprobe,
                    "-v",
                    "error",
                    "-show_entries",
                    "format=duration",
                    "-of",
                    "default=noprint_wrappers=1:nokey=1",
                    str(source),
                ],
                timeout=20,
            )
            if completed.returncode == 0:
                value = float((completed.stdout or "").strip().splitlines()[0])
                if value > 0:
                    return value
        except Exception:
            pass
    if ffmpeg_path:
        try:
            completed = run_command([ffmpeg_path, "-hide_banner", "-i", str(source)], timeout=20)
            text = f"{completed.stdout}\n{completed.stderr}"
            match = re.search(r"Duration:\s*(\d+):(\d+):(\d+(?:\.\d+)?)", text)
            if match:
                hours = int(match.group(1))
                minutes = int(match.group(2))
                seconds = float(match.group(3))
                return hours * 3600 + minutes * 60 + seconds
        except Exception:
            pass
    return None


def _fallback_sample_seconds(duration: float | None) -> list[float]:
    ratios = [0.12, 0.27, 0.44, 0.63, 0.82]
    if duration and duration > 1:
        safe_duration = max(0.8, duration - 0.35)
        return [max(0.15, min(safe_duration, duration * ratio)) for ratio in ratios]
    return [3.0, 8.0, 15.0, 26.0, 42.0]


def _ffmpeg_fallback(
    source: Path,
    output_dir: Path,
    ffmpeg_path: str | None,
    title: str,
) -> list[dict[str, Any]]:
    if not ffmpeg_path:
        return []
    duration = _probe_duration_for_fallback(source, ffmpeg_path)
    sample_seconds = _fallback_sample_seconds(duration)
    candidates: list[dict[str, Any]] = []
    vf = "scale=1280:720:force_original_aspect_ratio=decrease,pad=1280:720:(ow-iw)/2:(oh-ih)/2"
    for index, seconds in enumerate(sample_seconds, start=1):
        output_path = output_dir / f"thumb_{index:02d}.png"
        args = [
            ffmpeg_path,
            "-hide_banner",
            "-loglevel",
            "error",
            "-y",
            "-ss",
            f"{seconds:.3f}",
            "-i",
            str(source),
            "-frames:v",
            "1",
            "-vf",
            vf,
            str(output_path),
        ]
        try:
            completed = run_command(args, timeout=30)
        except Exception:
            continue
        if completed.returncode == 0 and output_path.exists():
            candidates.append(
                _decorate_candidate(
                    index,
                    {"seconds": seconds, "score": 0.0, "brightness": 0.0, "sharpness": 0.0},
                    title,
                    [],
                )
            )
    return candidates


def _write_log(
    path: Path,
    package_dir: Path,
    source: Path | None,
    source_label: str,
    candidates: list[dict[str, Any]],
    used_ollama: bool,
    error: str = "",
) -> Path:
    lines = [
        "Thumbnail Flow Log",
        f"created_at: {datetime.now().isoformat(timespec='seconds')}",
        f"package_dir: {package_dir}",
        f"source: {source or '--'}",
        f"source_label: {source_label}",
        f"size: {THUMBNAIL_SIZE[0]}x{THUMBNAIL_SIZE[1]}",
        f"candidates: {len(candidates)}",
        f"used_ollama: {used_ollama}",
        f"error: {error}",
        "",
        "files:",
        *[f"- {item.get('file')} / {item.get('timecode')} / {item.get('direction')}" for item in candidates],
        "",
        "policy:",
        "- Source video is not modified.",
        "- This is thumbnail candidate organization, not image editing.",
        "- No automatic upload.",
    ]
    return _write_text(path, "\n".join(lines))


def generate_thumbnail_candidates(
    package_dir: Path,
    source_video_path: Path | None = None,
    ffmpeg_path: str | None = None,
    ollama_ready: bool = False,
    log: LogCallback | None = None,
) -> dict[str, Any]:
    selected_dir = package_dir / "selected"
    thumbnail_dir = selected_dir / "thumbnails"
    thumbnail_dir.mkdir(parents=True, exist_ok=True)
    if log:
        log(THUMBNAIL_TEXT["start"])

    source, source_label = _resolve_thumbnail_source(package_dir, source_video_path)
    title = _title_for_package(package_dir)
    if source is None:
        log_path = _write_log(thumbnail_dir / "thumbnail_flow_log.txt", package_dir, source, source_label, [], False, "Source video is unavailable.")
        if log:
            log(THUMBNAIL_TEXT["failed"])
        return {
            "status": "FAILED",
            "package_dir": str(package_dir),
            "selected_dir": str(selected_dir),
            "thumbnail_dir": str(thumbnail_dir),
            "candidates": [],
            "candidate_count": 0,
            "json_path": "",
            "log_path": str(log_path),
            "source": "",
            "source_label": source_label,
            "used_ollama": False,
            "message": "Source video is unavailable.",
        }

    if log:
        log(THUMBNAIL_TEXT["search"])
    samples, duration = _select_samples(source)
    notes: list[dict[str, str]] = []
    used_ollama = False
    if ollama_ready:
        notes, used_ollama = _ollama_thumbnail_notes(package_dir, title, max(THUMBNAIL_COUNT, len(samples)))

    candidates: list[dict[str, Any]] = []
    if samples and cv2 is not None:
        for index, sample in enumerate(samples[:THUMBNAIL_COUNT], start=1):
            output_path = thumbnail_dir / f"thumb_{index:02d}.png"
            frame = _fit_thumbnail(sample.get("frame"))
            if frame is None:
                continue
            cv2.imwrite(str(output_path), frame)
            if output_path.exists():
                candidates.append(_decorate_candidate(index, sample, title, notes))
    if len(candidates) < THUMBNAIL_COUNT:
        fallback = _ffmpeg_fallback(source, thumbnail_dir, ffmpeg_path, title)
        existing = {str(item.get("file")) for item in candidates}
        for item in fallback:
            if str(item.get("file")) not in existing:
                candidates.append(item)
            if len(candidates) >= THUMBNAIL_COUNT:
                break

    json_path = thumbnail_dir / "thumbnail_candidates.json"
    _write_json(json_path, candidates)
    log_path = _write_log(thumbnail_dir / "thumbnail_flow_log.txt", package_dir, source, source_label, candidates, used_ollama)
    if log and candidates:
        log(THUMBNAIL_TEXT["ready"])
    elif log:
        log(THUMBNAIL_TEXT["failed"])

    return {
        "status": "COMPLETED" if candidates else "FAILED",
        "package_dir": str(package_dir),
        "selected_dir": str(selected_dir),
        "thumbnail_dir": str(thumbnail_dir),
        "candidates": candidates,
        "candidate_count": len(candidates),
        "json_path": str(json_path),
        "log_path": str(log_path),
        "source": str(source),
        "source_label": source_label,
        "duration": duration,
        "size": f"{THUMBNAIL_SIZE[0]}x{THUMBNAIL_SIZE[1]}",
        "used_ollama": used_ollama,
        "message": f"{len(candidates)} thumbnail candidates created." if candidates else "Thumbnail candidates were not created.",
    }


def read_thumbnail_candidates(package_dir: Path) -> list[dict[str, Any]]:
    payload = _read_json(package_dir / "selected" / "thumbnails" / "thumbnail_candidates.json")
    if isinstance(payload, list):
        return [dict(item) for item in payload if isinstance(item, dict)]
    return []


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


def _update_upload_checklist(checklist_path: Path, copied_path: Path) -> None:
    if not checklist_path.exists():
        return
    text = _read_text(checklist_path)
    line = f"  - {copied_path.name}"
    if copied_path.name in text:
        return
    if "  - No thumbnails available" in text:
        text = text.replace("  - No thumbnails available", line)
    else:
        text = text.rstrip() + f"\n{line}"
    _write_text(checklist_path, text)


def add_thumbnail_to_upload_package(
    package_dir: Path,
    thumbnail_path: Path | None = None,
    log: LogCallback | None = None,
) -> dict[str, Any]:
    thumbnail_dir = package_dir / "selected" / "thumbnails"
    if thumbnail_path is None:
        candidates = read_thumbnail_candidates(package_dir)
        if candidates:
            thumbnail_path = thumbnail_dir / str(candidates[0].get("file") or "")
        else:
            files = sorted(thumbnail_dir.glob("thumb_*.png"))
            thumbnail_path = files[0] if files else None
    if thumbnail_path is None or not thumbnail_path.exists():
        return {
            "status": "FAILED",
            "package_dir": str(package_dir),
            "copied_path": "",
            "message": "Thumbnail candidate is not ready.",
        }
    upload_thumbnail_dir = package_dir / "selected" / "upload_ready" / "thumbnails"
    upload_thumbnail_dir.mkdir(parents=True, exist_ok=True)
    destination = _unique_destination(upload_thumbnail_dir / thumbnail_path.name)
    shutil.copy2(thumbnail_path, destination)
    _update_upload_checklist(package_dir / "selected" / "upload_ready" / "metadata" / "upload_checklist.md", destination)
    if log:
        log(THUMBNAIL_TEXT["added"])
    return {
        "status": "COMPLETED",
        "package_dir": str(package_dir),
        "copied_path": str(destination),
        "upload_thumbnail_dir": str(upload_thumbnail_dir),
        "message": "Thumbnail copied to upload_ready.",
    }
