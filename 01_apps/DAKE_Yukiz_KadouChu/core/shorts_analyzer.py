from __future__ import annotations

import json
import re
from pathlib import Path

from core.app_config import SHORTS_REASON_TEXT, seconds_to_timecode
from core.project_writer import ProjectPaths


def _read_srt_starts(srt_path: Path | None) -> list[float]:
    if not srt_path or not srt_path.exists():
        return []
    text = srt_path.read_text(encoding="utf-8", errors="replace")
    starts: list[float] = []
    for match in re.finditer(r"(\d{2}):(\d{2}):(\d{2}),(\d{3})\s+-->", text):
        hours, minutes, seconds, millis = [int(group) for group in match.groups()]
        starts.append(hours * 3600 + minutes * 60 + seconds + millis / 1000)
    return starts


def create_shorts_candidates(duration: float, srt_path: Path | None = None) -> list[dict[str, object]]:
    if duration <= 0:
        return []
    clip_length = 45.0
    if duration < 45:
        clip_length = max(8.0, duration)
    elif duration < 90:
        clip_length = min(30.0, duration)

    count = 5 if duration >= 300 else 4 if duration >= 180 else 3 if duration >= 90 else 1
    starts = _read_srt_starts(srt_path)
    selected: list[float] = []

    for start in starts:
        if len(selected) >= count:
            break
        safe_start = min(max(0.0, start - 3), max(0.0, duration - clip_length))
        if all(abs(safe_start - other) > clip_length * 0.8 for other in selected):
            selected.append(safe_start)

    if len(selected) < count:
        spacing = max(0.0, (duration - clip_length) / (count + 1))
        for index in range(count):
            if len(selected) >= count:
                break
            start = min(max(0.0, spacing * (index + 1)), max(0.0, duration - clip_length))
            if all(abs(start - other) > clip_length * 0.6 for other in selected):
                selected.append(start)

    selected = sorted(selected)[:count]
    candidates: list[dict[str, object]] = []
    for index, start in enumerate(selected, start=1):
        end = min(duration, start + clip_length)
        clip_duration = max(0.0, end - start)
        reason = SHORTS_REASON_TEXT["speech"] if starts else SHORTS_REASON_TEXT["even"]
        candidates.append(
            {
                "id": index,
                "start": seconds_to_timecode(start),
                "end": seconds_to_timecode(end),
                "duration": seconds_to_timecode(clip_duration),
                "start_seconds": round(start, 3),
                "end_seconds": round(end, 3),
                "duration_seconds": round(clip_duration, 3),
                "reason": reason,
                "status": "candidate",
            }
        )
    return candidates


def write_shorts_candidates(project: ProjectPaths, candidates: list[dict[str, object]]) -> Path:
    path = project.root / "shorts_candidates.json"
    path.write_text(json.dumps(candidates, ensure_ascii=False, indent=2), encoding="utf-8")
    return path
