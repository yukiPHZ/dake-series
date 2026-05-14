from __future__ import annotations

import json
from collections import Counter
from pathlib import Path
from typing import Any

from core.memory_store import (
    collect_package_memory,
    ensure_memory_dirs,
    memory_index_path,
    memory_projects_dir,
    memory_summary_path,
)
from core.ollama_client import generate_ollama_text


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


def _as_list(value: Any) -> list[str]:
    if isinstance(value, list):
        return [str(item).strip() for item in value if str(item).strip()]
    text = str(value or "").strip()
    return [text] if text else []


def _top(counter: Counter[str], limit: int = 8) -> list[dict[str, Any]]:
    return [{"value": key, "count": count} for key, count in counter.most_common(limit)]


def load_memory_records(base_dir: Path | None = None, limit: int = 8) -> list[dict[str, Any]]:
    root = ensure_memory_dirs(base_dir)
    project_paths = sorted(memory_projects_dir(root).glob("*.json"), key=lambda path: path.stat().st_mtime, reverse=True)
    records: list[dict[str, Any]] = []
    for path in project_paths[:limit]:
        loaded = _read_json(path)
        if isinstance(loaded, dict):
            records.append(loaded)
    return records


def analyze_memory(base_dir: Path | None = None) -> dict[str, Any]:
    root = ensure_memory_dirs(base_dir)
    index_payload = _read_json(memory_index_path(root))
    index = [item for item in index_payload if isinstance(item, dict)] if isinstance(index_payload, list) else []
    recent_records = load_memory_records(root, limit=8)

    preset_counter: Counter[str] = Counter()
    bgm_counter: Counter[str] = Counter()
    title_counter: Counter[str] = Counter()
    shorts_counter: Counter[str] = Counter()
    project_counter: Counter[str] = Counter()

    for item in index:
        preset_counter.update(_as_list(item.get("preset")))
        bgm_counter.update(_as_list(item.get("bgm")))
        title_counter.update(_as_list(item.get("title")))
        shorts_counter.update(_as_list(item.get("shorts_direction")))
        project_counter.update(_as_list(item.get("package_name")))

    for record in recent_records:
        bridge = record.get("project_bridge") if isinstance(record.get("project_bridge"), dict) else {}
        project_counter.update(_as_list(bridge.get("project_name")))
        if not index:
            preset_counter.update(_as_list(bridge.get("preset")))
            bgm_counter.update(_as_list(record.get("bgm")))
            title_counter.update(_as_list(record.get("title")))
            shorts_counter.update(_as_list(bridge.get("shorts_direction")))

    return {
        "memory_dir": str(root),
        "index_path": str(memory_index_path(root)),
        "summary_path": str(memory_summary_path(root)),
        "entries": len(index),
        "project_records": len(recent_records),
        "index": index,
        "recent_records": recent_records,
        "summary_text": _read_text(memory_summary_path(root)),
        "preset_counts": _top(preset_counter),
        "bgm_counts": _top(bgm_counter),
        "title_counts": _top(title_counter),
        "shorts_direction_counts": _top(shorts_counter),
        "project_counts": _top(project_counter),
    }


def _counter_values(analysis: dict[str, Any], key: str, limit: int = 5) -> list[str]:
    values: list[str] = []
    for item in analysis.get(key, []):
        if isinstance(item, dict) and item.get("value"):
            values.append(str(item["value"]))
    return values[:limit]


def _current_direction(record: dict[str, Any]) -> list[str]:
    bridge = record.get("project_bridge") if isinstance(record.get("project_bridge"), dict) else {}
    parts = [
        bridge.get("editing_mood"),
        bridge.get("mood"),
        bridge.get("preset"),
        record.get("title"),
    ]
    parts.extend(_as_list(record.get("bgm"))[:2])
    return [str(part).strip() for part in parts if str(part or "").strip()][:6]


def _related_projects(analysis: dict[str, Any], current: dict[str, Any]) -> list[str]:
    current_bridge = current.get("project_bridge") if isinstance(current.get("project_bridge"), dict) else {}
    current_preset = str(current_bridge.get("preset") or "").lower()
    current_bgm = {item.lower() for item in _as_list(current.get("bgm"))}

    scored: list[tuple[int, str]] = []
    for record in analysis.get("recent_records", []):
        if not isinstance(record, dict):
            continue
        bridge = record.get("project_bridge") if isinstance(record.get("project_bridge"), dict) else {}
        name = str(bridge.get("project_name") or record.get("package_name") or "").strip()
        if not name:
            continue
        score = 1
        if current_preset and current_preset == str(bridge.get("preset") or "").lower():
            score += 3
        if current_bgm.intersection({item.lower() for item in _as_list(record.get("bgm"))}):
            score += 3
        scored.append((score, name))
    scored.sort(key=lambda item: (-item[0], item[1]))

    seen: set[str] = set()
    related: list[str] = []
    for _score, name in scored:
        if name not in seen:
            seen.add(name)
            related.append(name)
        if len(related) >= 5:
            break
    if related:
        return related
    return _counter_values(analysis, "project_counts", limit=5) or ["No related memory yet."]


def _template_recommendation(analysis: dict[str, Any], current: dict[str, Any]) -> str:
    bridge = current.get("project_bridge") if isinstance(current.get("project_bridge"), dict) else {}
    current_parts = _current_direction(current) or ["quiet workflow", "calm process"]
    related = _related_projects(analysis, current)
    presets = _counter_values(analysis, "preset_counts", limit=3)
    bgm = _counter_values(analysis, "bgm_counts", limit=3)
    titles = _counter_values(analysis, "title_counts", limit=4)
    shorts = _as_list(bridge.get("shorts_direction")) or _counter_values(analysis, "shorts_direction_counts", limit=5)
    if not shorts:
        shorts = ["まだ作ってる。", "深夜の机", "静かなコード作業", "机と光"]
    title_direction = _as_list(bridge.get("title_direction")) or titles or ["深夜、まだ作ってる。", "今日も少し進める。", "稼働中。"]

    similar: list[str] = []
    if presets:
        similar.append(f"{presets[0]}系の記録が目立ちます。")
    if bgm:
        similar.append(f"{bgm[0]} の使用履歴があります。")
    if titles:
        similar.append("短く余白のあるタイトルが多めです。")
    if not similar:
        similar.append("まだ履歴が少ないため、現在の素材を基準に整えています。")

    lines = [
        "# 補助脳リコメンド",
        "",
        "## Current Direction",
        *[f"- {item}" for item in current_parts],
        "",
        "## Related Past Projects",
        *[f"- {item}" for item in related],
        "",
        "## Similar Tendencies",
        *[f"- {item}" for item in similar],
        "",
        "## Suggested Shorts Direction",
        *[f"- {item}" for item in shorts[:5]],
        "",
        "## Suggested Title Direction",
        *[f"- {item}" for item in title_direction[:5]],
        "",
        "## Suggested Editing Mood",
        "- カットを急ぎすぎない",
        "- 間を残す",
        "- テロップ少なめ",
        "- 静かな余熱を残す",
        "",
        "## Assistant Note",
        "最近の制作は、“静かな作業の余熱”へ寄っています。",
        "",
        "最終判断はユーザーが行います。",
    ]
    return "\n".join(lines).rstrip() + "\n"


def _ollama_prompt(analysis: dict[str, Any], current: dict[str, Any]) -> str:
    compact = {
        "memory_summary": str(analysis.get("summary_text") or "")[:1600],
        "recent_index": analysis.get("index", [])[-8:],
        "preset_counts": analysis.get("preset_counts", [])[:5],
        "bgm_counts": analysis.get("bgm_counts", [])[:5],
        "title_counts": analysis.get("title_counts", [])[:5],
        "shorts_direction_counts": analysis.get("shorts_direction_counts", [])[:6],
        "current": {
            "package_name": current.get("package_name", ""),
            "title": current.get("title", ""),
            "bgm": current.get("bgm", []),
            "project_bridge": current.get("project_bridge", {}),
            "selected_short": current.get("selected_short", {}),
            "assistant_review_excerpt": str(current.get("assistant_review") or "")[:1000],
            "metadata_excerpt": str(current.get("metadata_draft") or "")[:1000],
        },
    }
    return (
        "You are the local assistant brain for a quiet GPU video production console.\n"
        "Create a recommendation Markdown in Japanese. You only suggest; the user decides.\n"
        "Do not mention automatic upload as available. Do not call external APIs.\n"
        "Keep it short and practical.\n"
        "Required sections: # 補助脳リコメンド, Current Direction, Related Past Projects, Similar Tendencies, "
        "Suggested Shorts Direction, Suggested Title Direction, Suggested Editing Mood, Assistant Note.\n"
        "End with: 最終判断はユーザーが行います。\n\n"
        f"Context JSON:\n{json.dumps(compact, ensure_ascii=False)[:6500]}"
    )


def generate_assistant_recommendation(
    package_dir: Path,
    ollama_ready: bool = False,
    memory_base_dir: Path | None = None,
) -> dict[str, Any]:
    analysis = analyze_memory(memory_base_dir)
    current = collect_package_memory(package_dir)
    used_ollama = False
    ollama_model = ""
    ollama_reason = ""
    text = ""

    if ollama_ready:
        response = generate_ollama_text(_ollama_prompt(analysis, current), timeout=45)
        used_ollama = bool(response.get("ok"))
        ollama_model = str(response.get("model") or "")
        ollama_reason = str(response.get("reason") or "")
        if used_ollama:
            text = str(response.get("text") or "").strip()

    if not text:
        text = _template_recommendation(analysis, current).strip()
    if "最終判断はユーザーが行います。" not in text:
        text = text.rstrip() + "\n\n最終判断はユーザーが行います。"

    output_path = package_dir / "assistant_recommendation.md"
    output_path.write_text(text.rstrip() + "\n", encoding="utf-8")
    return {
        "status": "COMPLETED",
        "package_dir": str(package_dir),
        "recommendation_path": str(output_path),
        "memory_dir": str(analysis.get("memory_dir", "")),
        "memory_entries": analysis.get("entries", 0),
        "project_records": analysis.get("project_records", 0),
        "current_direction": _current_direction(current),
        "related_projects": _related_projects(analysis, current),
        "preset_counts": analysis.get("preset_counts", []),
        "bgm_counts": analysis.get("bgm_counts", []),
        "title_counts": analysis.get("title_counts", []),
        "shorts_direction_counts": analysis.get("shorts_direction_counts", []),
        "used_ollama": used_ollama,
        "ollama_model": ollama_model,
        "ollama_reason": ollama_reason,
    }
