from __future__ import annotations

import json
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

from core.memory_store import memory_summary_path
from core.ollama_client import generate_ollama_text

LogCallback = Callable[[str], None]

TITLE_MATCH_TEXT = {
    "start": "補助脳：サムネとタイトルの相性を整理しています。",
    "ready": "補助脳：入口セットを作成しました。",
    "human": "補助脳：最後だけ、菊田さんが握ってください。",
    "failed": "補助脳：Title Match の生成に失敗しました。",
    "added": "補助脳：入口セットを投稿前セットへ追加しました。",
}

FALLBACK_TITLES = [
    "深夜、まだ作ってる。",
    "稼働中。",
    "今日も少し進める。",
    "静かな作業机。",
    "止まらず作る。",
]

FALLBACK_REASONS = [
    "机と光の静かな入口感があり、タイトルの余白と合います。",
    "作業感が伝わりやすい。",
    "余熱感がある。",
]

DEFAULT_SHORTS_DIRECTION = {
    "intro": "静かな導入",
    "work": "作業感",
    "afterglow": "余熱",
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


def _nonempty_lines(text: str) -> list[str]:
    rows: list[str] = []
    for raw in text.splitlines():
        line = raw.strip().lstrip("-").strip()
        if line and not line.startswith("#"):
            rows.append(line)
    return rows


def _section_lines(text: str, names: set[str], limit: int = 5) -> list[str]:
    rows = text.splitlines()
    capture = False
    found: list[str] = []
    for raw in rows:
        line = raw.strip()
        if line.startswith("## "):
            heading = line.lstrip("#").strip().lower()
            capture = heading in names
            continue
        if capture:
            if line.startswith("#"):
                break
            cleaned = line.lstrip("-").strip()
            if cleaned:
                found.append(cleaned)
            if len(found) >= limit:
                break
    return found


def _unique(values: list[str], limit: int = 10) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for value in values:
        cleaned = value.strip()
        if not cleaned:
            continue
        key = cleaned.lower()
        if key in seen:
            continue
        seen.add(key)
        result.append(cleaned)
        if len(result) >= limit:
            break
    return result


def _title_candidates(package_dir: Path) -> list[str]:
    rows: list[str] = []
    selected_title = _read_text(package_dir / "selected" / "selected_title.txt")
    if selected_title:
        rows.append(selected_title.splitlines()[0].strip())
    rows.extend(_nonempty_lines(_read_text(package_dir / "metadata" / "title_ideas.txt")))
    rows.extend(_section_lines(_read_text(package_dir / "assistant_recommendation.md", limit=5000), {"suggested title direction", "title direction"}))
    thumbnails = _read_json(package_dir / "selected" / "thumbnails" / "thumbnail_candidates.json")
    if isinstance(thumbnails, list):
        for item in thumbnails:
            if isinstance(item, dict):
                rows.append(str(item.get("title_match") or "").strip())
    rows.extend(FALLBACK_TITLES)
    return _unique(rows, limit=10)


def _thumbnail_candidates(package_dir: Path) -> list[dict[str, str]]:
    thumbnail_dir = package_dir / "selected" / "thumbnails"
    payload = _read_json(thumbnail_dir / "thumbnail_candidates.json")
    rows: list[dict[str, str]] = []
    if isinstance(payload, list):
        for index, item in enumerate(payload, start=1):
            if not isinstance(item, dict):
                continue
            file_name = str(item.get("file") or f"thumb_{index:02d}.png").strip()
            if not file_name:
                continue
            rows.append(
                {
                    "file": file_name,
                    "direction": str(item.get("direction") or "quiet work").strip(),
                    "reason": str(item.get("reason") or "静かな入口").strip(),
                    "title_match": str(item.get("title_match") or "").strip(),
                }
            )
    if not rows and thumbnail_dir.exists():
        for index, path in enumerate(sorted(thumbnail_dir.glob("thumb_*.png")), start=1):
            rows.append(
                {
                    "file": path.name,
                    "direction": "quiet work",
                    "reason": FALLBACK_REASONS[(index - 1) % len(FALLBACK_REASONS)],
                    "title_match": "",
                }
            )
    if not rows:
        rows.append({"file": "", "direction": "quiet work", "reason": "サムネ候補は未生成です。", "title_match": ""})
    return rows[:7]


def _shorts_direction(package_dir: Path) -> dict[str, str]:
    direction = DEFAULT_SHORTS_DIRECTION.copy()
    payload = _read_json(package_dir / "selected" / "shorts_pack" / "shorts_pack.json")
    items: list[Any] = []
    if isinstance(payload, dict):
        raw = payload.get("items") or payload.get("shorts") or payload.get("clips")
        items = raw if isinstance(raw, list) else []
    elif isinstance(payload, list):
        items = payload
    for item in items:
        if not isinstance(item, dict):
            continue
        kind = str(item.get("type") or item.get("role") or "").lower()
        reason = str(item.get("reason") or item.get("caption_idea") or "").strip()
        if not reason:
            continue
        if "intro" in kind:
            direction["intro"] = reason
        elif "work" in kind:
            direction["work"] = reason
        elif "after" in kind:
            direction["afterglow"] = reason
    return direction


def _score_pair(thumbnail: dict[str, str], title: str, index: int) -> float:
    text = " ".join([thumbnail.get("direction", ""), thumbnail.get("reason", ""), thumbnail.get("title_match", ""), title]).lower()
    score = 100.0 - index
    if thumbnail.get("title_match") and thumbnail["title_match"] in title:
        score += 25
    if any(word in text for word in ["desk", "light", "机", "光"]):
        score += 8 if any(word in title for word in ["深夜", "机", "光", "まだ"]) else 0
    if any(word in text for word in ["quiet", "静か", "導入"]):
        score += 7 if any(word in title for word in ["稼働中", "静か", "深夜"]) else 0
    if any(word in text for word in ["afterglow", "余熱", "夜"]):
        score += 7 if any(word in title for word in ["まだ", "今日", "夜", "少し"]) else 0
    if any(word in text for word in ["work", "作業", "タイピング"]):
        score += 5 if any(word in title for word in ["作", "進める", "稼働"]) else 0
    return score


def _template_pairs(thumbnails: list[dict[str, str]], titles: list[str]) -> list[dict[str, str]]:
    scored: list[tuple[float, dict[str, str]]] = []
    for thumb_index, thumbnail in enumerate(thumbnails):
        for title in titles:
            reason = thumbnail.get("reason") or FALLBACK_REASONS[thumb_index % len(FALLBACK_REASONS)]
            scored.append(
                (
                    _score_pair(thumbnail, title, thumb_index),
                    {
                        "thumbnail": thumbnail.get("file", ""),
                        "title": title,
                        "reason": reason,
                    },
                )
            )
    scored.sort(key=lambda item: item[0], reverse=True)
    pairs: list[dict[str, str]] = []
    used_thumbnails: set[str] = set()
    used_titles: set[str] = set()
    for _score, pair in scored:
        thumbnail_name = pair["thumbnail"]
        title = pair["title"]
        if thumbnail_name in used_thumbnails or title in used_titles:
            continue
        pairs.append(pair)
        used_thumbnails.add(thumbnail_name)
        used_titles.add(title)
        if len(pairs) >= 3:
            break
    while len(pairs) < 3 and titles:
        index = len(pairs)
        thumbnail = thumbnails[index % len(thumbnails)]
        pairs.append(
            {
                "thumbnail": thumbnail.get("file", ""),
                "title": titles[index % len(titles)],
                "reason": FALLBACK_REASONS[index % len(FALLBACK_REASONS)],
            }
        )
    return pairs[:3]


def _normalize_pair(pair: Any, thumbnails: list[dict[str, str]], titles: list[str], fallback: dict[str, str]) -> dict[str, str]:
    available = {item.get("file", "") for item in thumbnails}
    if not isinstance(pair, dict):
        return fallback
    thumbnail = str(pair.get("thumbnail") or fallback.get("thumbnail") or "").strip()
    title = str(pair.get("title") or fallback.get("title") or "").strip()
    reason = str(pair.get("reason") or fallback.get("reason") or "").strip()
    if thumbnail not in available:
        thumbnail = fallback.get("thumbnail", "")
    if not title:
        title = titles[0] if titles else FALLBACK_TITLES[0]
    if not reason:
        reason = fallback.get("reason", FALLBACK_REASONS[0])
    return {"thumbnail": thumbnail, "title": title, "reason": reason}


def _ollama_title_match(
    package_dir: Path,
    thumbnails: list[dict[str, str]],
    titles: list[str],
    shorts_direction: dict[str, str],
    fallback_pairs: list[dict[str, str]],
) -> tuple[list[dict[str, str]], dict[str, str], bool]:
    review = _read_text(package_dir / "assistant_review.md", limit=1200)
    recommendation = _read_text(package_dir / "assistant_recommendation.md", limit=1200)
    metadata_draft = _read_text(package_dir / "selected" / "upload" / "metadata_draft.txt", limit=900)
    memory = _read_text(memory_summary_path(), limit=1000)
    prompt = (
        "You are the local assistant brain for a quiet production console.\n"
        "Match existing thumbnail candidates and title candidates. Do not upload. Do not create images.\n"
        "Return JSON only with keys: best_pair, alternatives, shorts_direction.\n"
        "best_pair and alternatives items require thumbnail, title, reason.\n"
        "Keep reasons short Japanese.\n\n"
        f"Thumbnails:\n{json.dumps(thumbnails, ensure_ascii=False)[:2500]}\n\n"
        f"Titles:\n{json.dumps(titles, ensure_ascii=False)[:1000]}\n\n"
        f"Shorts direction:\n{json.dumps(shorts_direction, ensure_ascii=False)}\n\n"
        f"Assistant review:\n{review}\n\nRecommendation:\n{recommendation}\n\nMetadata draft:\n{metadata_draft}\n\nMemory:\n{memory}\n"
    )
    response = generate_ollama_text(prompt, timeout=40)
    if not response.get("ok"):
        return fallback_pairs, shorts_direction, False
    text = str(response.get("text") or "")
    start = text.find("{")
    end = text.rfind("}")
    if start == -1 or end == -1 or end <= start:
        return fallback_pairs, shorts_direction, False
    try:
        payload = json.loads(text[start : end + 1])
    except Exception:
        return fallback_pairs, shorts_direction, False
    pairs = [_normalize_pair(payload.get("best_pair"), thumbnails, titles, fallback_pairs[0])]
    alternatives = payload.get("alternatives")
    if isinstance(alternatives, list):
        for index, item in enumerate(alternatives[:2], start=1):
            pairs.append(_normalize_pair(item, thumbnails, titles, fallback_pairs[min(index, len(fallback_pairs) - 1)]))
    while len(pairs) < 3:
        pairs.append(fallback_pairs[len(pairs)])
    raw_direction = payload.get("shorts_direction")
    if isinstance(raw_direction, dict):
        merged = shorts_direction.copy()
        for key in ["intro", "work", "afterglow"]:
            value = str(raw_direction.get(key) or "").strip()
            if value:
                merged[key] = value
        shorts_direction = merged
    return pairs[:3], shorts_direction, True


def _build_markdown(pairs: list[dict[str, str]], shorts_direction: dict[str, str]) -> str:
    best = pairs[0]
    alternatives = pairs[1:3]
    lines = [
        "# Title Match",
        "",
        "## Best Pair",
        "Thumbnail:",
        best["thumbnail"] or "--",
        "",
        "Title:",
        best["title"],
        "",
        "Reason:",
        best["reason"],
        "",
        "## Alternative Pairs",
    ]
    for index, pair in enumerate(alternatives, start=2):
        lines.extend(
            [
                "",
                f"### Pair {index}",
                "Thumbnail:",
                pair["thumbnail"] or "--",
                "",
                "Title:",
                pair["title"],
                "",
                "Reason:",
                pair["reason"],
            ]
        )
    lines.extend(
        [
            "",
            "## Shorts Direction",
            f"- INTRO: {shorts_direction.get('intro', DEFAULT_SHORTS_DIRECTION['intro'])}",
            f"- WORK: {shorts_direction.get('work', DEFAULT_SHORTS_DIRECTION['work'])}",
            f"- AFTERGLOW: {shorts_direction.get('afterglow', DEFAULT_SHORTS_DIRECTION['afterglow'])}",
            "",
            "## Assistant Note",
            "最終判断はユーザーが行います。",
        ]
    )
    return "\n".join(lines)


def _write_log(
    path: Path,
    package_dir: Path,
    pairs: list[dict[str, str]],
    thumbnails: list[dict[str, str]],
    titles: list[str],
    used_ollama: bool,
    error: str = "",
) -> Path:
    lines = [
        "Title Match Log",
        f"created_at: {datetime.now().isoformat(timespec='seconds')}",
        f"package_dir: {package_dir}",
        f"used_ollama: {used_ollama}",
        f"thumbnail_candidates: {len(thumbnails)}",
        f"title_candidates: {len(titles)}",
        f"error: {error}",
        "",
        "pairs:",
        *[f"- {item.get('thumbnail')}: {item.get('title')} / {item.get('reason')}" for item in pairs],
        "",
        "policy:",
        "- No YouTube upload.",
        "- No automatic publish.",
        "- Final judgment belongs to the user.",
    ]
    return _write_text(path, "\n".join(lines))


def generate_title_match(
    package_dir: Path,
    ollama_ready: bool = False,
    log: LogCallback | None = None,
) -> dict[str, Any]:
    selected_dir = package_dir / "selected"
    output_dir = selected_dir / "title_match"
    output_dir.mkdir(parents=True, exist_ok=True)
    if log:
        log(TITLE_MATCH_TEXT["start"])

    thumbnails = _thumbnail_candidates(package_dir)
    titles = _title_candidates(package_dir)
    shorts_direction = _shorts_direction(package_dir)
    fallback_pairs = _template_pairs(thumbnails, titles)
    pairs = fallback_pairs
    used_ollama = False
    if ollama_ready:
        pairs, shorts_direction, used_ollama = _ollama_title_match(package_dir, thumbnails, titles, shorts_direction, fallback_pairs)

    best_pair = pairs[0] if pairs else {"thumbnail": "", "title": titles[0] if titles else FALLBACK_TITLES[0], "reason": FALLBACK_REASONS[0]}
    alternatives = pairs[1:3]
    payload = {
        "best_pair": best_pair,
        "alternatives": alternatives,
        "shorts_direction": shorts_direction,
        "used_ollama": used_ollama,
        "created_at": datetime.now().isoformat(timespec="seconds"),
    }
    md_path = _write_text(output_dir / "title_match.md", _build_markdown([best_pair, *alternatives], shorts_direction))
    json_path = _write_json(output_dir / "title_match.json", payload)
    log_path = _write_log(output_dir / "title_match_log.txt", package_dir, [best_pair, *alternatives], thumbnails, titles, used_ollama)
    if log:
        log(TITLE_MATCH_TEXT["ready"])
        log(TITLE_MATCH_TEXT["human"])
    return {
        "status": "COMPLETED",
        "package_dir": str(package_dir),
        "selected_dir": str(selected_dir),
        "title_match_dir": str(output_dir),
        "md_path": str(md_path),
        "json_path": str(json_path),
        "log_path": str(log_path),
        "best_pair": best_pair,
        "alternatives": alternatives,
        "shorts_direction": shorts_direction,
        "used_ollama": used_ollama,
        "message": "Title Match created.",
    }


def read_title_match(package_dir: Path) -> dict[str, Any]:
    payload = _read_json(package_dir / "selected" / "title_match" / "title_match.json")
    return payload if isinstance(payload, dict) else {}


def add_title_match_to_upload_package(
    package_dir: Path,
    log: LogCallback | None = None,
) -> dict[str, Any]:
    title_match_dir = package_dir / "selected" / "title_match"
    md_path = title_match_dir / "title_match.md"
    json_path = title_match_dir / "title_match.json"
    if not md_path.exists() or not json_path.exists():
        return {
            "status": "FAILED",
            "package_dir": str(package_dir),
            "message": "Title Match is not ready.",
        }

    upload_dir = package_dir / "selected" / "upload_ready"
    metadata_dir = upload_dir / "metadata"
    thumbnail_dir = upload_dir / "thumbnails"
    metadata_dir.mkdir(parents=True, exist_ok=True)
    thumbnail_dir.mkdir(parents=True, exist_ok=True)
    md_destination = metadata_dir / "title_match.md"
    json_destination = metadata_dir / "title_match.json"
    shutil.copy2(md_path, md_destination)
    shutil.copy2(json_path, json_destination)

    payload = read_title_match(package_dir)
    best_pair = payload.get("best_pair") if isinstance(payload, dict) else {}
    best_thumbnail_name = str(best_pair.get("thumbnail") or "") if isinstance(best_pair, dict) else ""
    best_source = package_dir / "selected" / "thumbnails" / best_thumbnail_name
    best_destination = thumbnail_dir / "best_thumbnail.png"
    copied_thumbnail = ""
    if best_thumbnail_name and best_source.exists():
        shutil.copy2(best_source, best_destination)
        copied_thumbnail = str(best_destination)
    if log:
        log(TITLE_MATCH_TEXT["added"])
    return {
        "status": "COMPLETED",
        "package_dir": str(package_dir),
        "upload_metadata_dir": str(metadata_dir),
        "md_path": str(md_destination),
        "json_path": str(json_destination),
        "best_thumbnail_path": copied_thumbnail,
        "message": "Title Match copied to upload_ready.",
    }
