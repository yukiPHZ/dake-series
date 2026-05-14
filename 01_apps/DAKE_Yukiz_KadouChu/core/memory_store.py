from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any

from core.app_config import ensure_app_dirs, memory_dir
from core.ollama_client import generate_ollama_text


def memory_projects_dir(base_dir: Path | None = None) -> Path:
    root = base_dir or memory_dir()
    return root / "projects"


def memory_index_path(base_dir: Path | None = None) -> Path:
    return (base_dir or memory_dir()) / "memory_index.json"


def memory_summary_path(base_dir: Path | None = None) -> Path:
    return (base_dir or memory_dir()) / "memory_summary.md"


def ensure_memory_dirs(base_dir: Path | None = None) -> Path:
    if base_dir is None:
        ensure_app_dirs()
    root = base_dir or memory_dir()
    memory_projects_dir(root).mkdir(parents=True, exist_ok=True)
    return root


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


def _write_json(path: Path, payload: Any) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def _safe_name(value: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value)
    cleaned = re.sub(r"\s+", "_", cleaned).strip("._ ")
    return (cleaned or "package")[:80]


def _first_title(package_dir: Path) -> str:
    selected = _read_text(package_dir / "selected" / "selected_title.txt")
    if selected:
        return selected.splitlines()[0].strip()
    title_ideas = _read_text(package_dir / "metadata" / "title_ideas.txt")
    for raw in title_ideas.splitlines():
        line = raw.strip().lstrip("-").strip()
        if line:
            return line
    return ""


def _source_video_path(package_dir: Path) -> str:
    meta = _read_json(package_dir / "package_meta.json")
    if isinstance(meta, dict) and meta.get("source_video_path"):
        return str(meta.get("source_video_path"))
    log_text = _read_text(package_dir / "logs" / "package_log.txt")
    for line in log_text.splitlines():
        if line.lower().startswith("source_video_path:"):
            return line.split(":", 1)[1].strip()
    media = _read_json(package_dir / "media_info.json")
    if isinstance(media, dict) and media.get("source_video_path"):
        return str(media.get("source_video_path"))
    return ""


def _selected_short(package_dir: Path) -> dict[str, Any]:
    selected = _read_json(package_dir / "selected" / "selected_short.json")
    if isinstance(selected, dict):
        return selected
    shorts = _read_json(package_dir / "shorts_candidates.json")
    if isinstance(shorts, list) and shorts and isinstance(shorts[0], dict):
        return shorts[0]
    return {}


def _bgm_files(package_dir: Path) -> list[str]:
    bgm_dir = package_dir / "selected" / "bgm"
    if not bgm_dir.exists():
        return []
    return sorted(
        [
            path.name
            for path in bgm_dir.iterdir()
            if path.is_file() and path.suffix.lower() in {".mp3", ".wav"}
        ],
        key=str.lower,
    )


def _section_value(text: str, section: str) -> str:
    lines = text.splitlines()
    for index, raw in enumerate(lines):
        if raw.strip().lower() == f"{section.lower()}:":
            values: list[str] = []
            for next_line in lines[index + 1 :]:
                line = next_line.strip()
                if not line:
                    if values:
                        break
                    continue
                if line.endswith(":") and len(line) <= 60:
                    break
                values.append(line.lstrip("-").strip())
            return "\n".join(values).strip()
    return ""


def _label_value(text: str, label: str) -> str:
    pattern = re.compile(rf"^\s*{re.escape(label)}\s*:\s*(.+?)\s*$", re.IGNORECASE | re.MULTILINE)
    match = pattern.search(text)
    return match.group(1).strip() if match else ""


def _direction_list(value: str) -> list[str]:
    if not value:
        return []
    parts: list[str] = []
    for line in value.splitlines():
        line = line.strip().lstrip("-").strip()
        if not line:
            continue
        if "/" in line:
            parts.extend(part.strip() for part in line.split("/") if part.strip())
        else:
            parts.append(line)
    return parts[:8]


def _bridge_values(metadata_text: str) -> dict[str, Any]:
    shorts_section = _section_value(metadata_text, "Shorts Direction")
    project = _section_value(metadata_text, "Project")
    preset = _section_value(metadata_text, "Selected Preset")
    return {
        "project_name": project.splitlines()[0] if project else "",
        "preset": preset.splitlines()[0] if preset else "",
        "mood": _section_value(metadata_text, "Suggested Mood"),
        "editing_mood": _label_value(metadata_text, "editing_mood"),
        "suggested_scene": _label_value(metadata_text, "suggested_scene"),
        "shorts_direction": _direction_list(_label_value(metadata_text, "shorts_direction") or shorts_section),
        "title_direction": _direction_list(_label_value(metadata_text, "title_direction") or _section_value(metadata_text, "Suggested Title")),
    }


def collect_package_memory(package_dir: Path) -> dict[str, Any]:
    package = package_dir.resolve()
    selected_summary = _read_text(package / "selected" / "selected_summary.md")
    assistant_review = _read_text(package / "assistant_review.md")
    metadata_draft = _read_text(package / "selected" / "upload" / "metadata_draft.txt")
    bridge = _bridge_values(metadata_draft)
    created_at = datetime.now().isoformat(timespec="seconds")
    return {
        "created_at": created_at,
        "package_name": package.name,
        "source_video_path": _source_video_path(package),
        "title": _first_title(package),
        "selected_short": _selected_short(package),
        "selected_summary": selected_summary,
        "assistant_review": assistant_review,
        "metadata_draft": metadata_draft,
        "bgm": _bgm_files(package),
        "project_bridge": {
            "project_name": bridge["project_name"],
            "preset": bridge["preset"],
            "mood": bridge["mood"],
            "editing_mood": bridge["editing_mood"],
            "suggested_scene": bridge["suggested_scene"],
            "shorts_direction": bridge["shorts_direction"],
            "title_direction": bridge["title_direction"],
        },
        "source_package": str(package),
    }


def _load_index(base_dir: Path | None = None) -> list[dict[str, Any]]:
    path = memory_index_path(base_dir)
    loaded = _read_json(path)
    if isinstance(loaded, list):
        return [item for item in loaded if isinstance(item, dict)]
    return []


def _index_entry(record: dict[str, Any]) -> dict[str, Any]:
    bridge = record.get("project_bridge") if isinstance(record.get("project_bridge"), dict) else {}
    return {
        "created_at": record.get("created_at", ""),
        "package_name": record.get("package_name", ""),
        "title": record.get("title", ""),
        "preset": bridge.get("preset", ""),
        "mood": bridge.get("editing_mood") or bridge.get("mood", ""),
        "bgm": record.get("bgm", []),
        "shorts_direction": bridge.get("shorts_direction", []),
        "source_package": record.get("source_package", ""),
    }


def _record_markdown(record: dict[str, Any]) -> str:
    bridge = record.get("project_bridge") if isinstance(record.get("project_bridge"), dict) else {}
    selected_short = record.get("selected_short") if isinstance(record.get("selected_short"), dict) else {}
    lines = [
        "# Memory Record",
        "",
        f"- Created At: {record.get('created_at', '')}",
        f"- Package: {record.get('package_name', '')}",
        f"- Source Video: {record.get('source_video_path', '')}",
        f"- Title: {record.get('title', '')}",
        f"- Project Bridge: {bridge.get('project_name', '')}",
        f"- Preset: {bridge.get('preset', '')}",
        f"- Editing Mood: {bridge.get('editing_mood') or bridge.get('mood', '')}",
        f"- Suggested Scene: {bridge.get('suggested_scene', '')}",
        f"- BGM: {', '.join(str(item) for item in record.get('bgm', []))}",
        "",
        "## Selected Short",
        f"- start: {selected_short.get('start', '')}",
        f"- end: {selected_short.get('end', '')}",
        f"- reason: {selected_short.get('reason', '')}",
        "",
        "## Shorts Direction",
        *[f"- {item}" for item in bridge.get("shorts_direction", [])],
        "",
        "## Title Direction",
        *[f"- {item}" for item in bridge.get("title_direction", [])],
        "",
        "## Human Decision",
        "最終判断はユーザーが行います。",
        "自動投稿はしていません。",
    ]
    return "\n".join(lines).rstrip() + "\n"


def _word_counts(index: list[dict[str, Any]], key: str) -> list[tuple[str, int]]:
    counts: dict[str, int] = {}
    for item in index:
        value = item.get(key)
        values = value if isinstance(value, list) else [value]
        for raw in values:
            text = str(raw or "").strip()
            if not text:
                continue
            counts[text] = counts.get(text, 0) + 1
    return sorted(counts.items(), key=lambda pair: (-pair[1], pair[0]))[:8]


def _template_summary(index: list[dict[str, Any]]) -> str:
    presets = _word_counts(index, "preset")
    moods = _word_counts(index, "mood")
    bgm = _word_counts(index, "bgm")
    directions = _word_counts(index, "shorts_direction")
    titles = [str(item.get("title") or "").strip() for item in index if str(item.get("title") or "").strip()]
    recent_titles = titles[-5:]

    lines = [
        "# 補助脳メモリ",
        "",
        "## 最近の制作傾向",
    ]
    if presets:
        lines.append(f"- {presets[0][0]}系のBGMを使う流れがあります。")
    if moods:
        lines.append(f"- {moods[0][0]}の空気と相性が良いです。")
    if directions:
        lines.append(f"- {directions[0][0]}をShorts方向として使うことが多いです。")
    if not (presets or moods or directions):
        lines.append("- まだ履歴が少ないため、傾向はこれから育ちます。")

    lines.extend(["", "## よく使う言葉"])
    words = ["稼働中。", "まだ作ってる。", "静かな作業机。"]
    for title in recent_titles:
        if title not in words:
            words.append(title)
    lines.extend(f"- {word}" for word in words[:8])

    lines.extend(["", "## 最近のBGM"])
    if bgm:
        lines.extend(f"- {name}" for name, _count in bgm[:8])
    else:
        lines.append("- まだBGM履歴はありません。")

    lines.extend(
        [
            "",
            "## 補助脳メモ",
            "最後の判断はユーザーが行います。",
        ]
    )
    return "\n".join(lines).rstrip() + "\n"


def _ollama_summary_prompt(index: list[dict[str, Any]], latest_record: dict[str, Any] | None) -> str:
    payload = {
        "recent_index": index[-12:],
        "latest_record": latest_record or {},
    }
    return (
        "You are a local assistant memory layer for a quiet video production console.\n"
        "Summarize production tendencies in Japanese. Keep it short, calm, and practical.\n"
        "Do not decide automatically. The user makes final decisions.\n"
        "Return Markdown with these sections: # 補助脳メモリ, 最近の制作傾向, よく使う言葉, 最近のBGM, 補助脳メモ.\n\n"
        f"Memory JSON:\n{json.dumps(payload, ensure_ascii=False)[:6000]}"
    )


def generate_memory_summary(
    latest_record: dict[str, Any] | None = None,
    ollama_ready: bool = False,
    base_dir: Path | None = None,
) -> dict[str, Any]:
    root = ensure_memory_dirs(base_dir)
    index = _load_index(root)
    used_ollama = False
    ollama_model = ""
    text = ""
    reason = ""
    if ollama_ready:
        response = generate_ollama_text(_ollama_summary_prompt(index, latest_record), timeout=45)
        used_ollama = bool(response.get("ok"))
        ollama_model = str(response.get("model") or "")
        reason = str(response.get("reason") or "")
        if used_ollama:
            text = str(response.get("text") or "").strip()
    if not text:
        text = _template_summary(index)
    path = memory_summary_path(root)
    path.write_text(text.rstrip() + "\n", encoding="utf-8")
    return {
        "status": "UPDATED",
        "memory_dir": str(root),
        "summary_path": str(path),
        "entries": len(index),
        "used_ollama": used_ollama,
        "ollama_model": ollama_model,
        "ollama_reason": reason,
    }


def save_package_to_memory(
    package_dir: Path,
    ollama_ready: bool = False,
    base_dir: Path | None = None,
) -> dict[str, Any]:
    root = ensure_memory_dirs(base_dir)
    record = collect_package_memory(package_dir)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    basename = f"{stamp}_{_safe_name(record['package_name'])}"
    json_path = memory_projects_dir(root) / f"{basename}.json"
    md_path = memory_projects_dir(root) / f"{basename}.md"
    _write_json(json_path, record)
    md_path.write_text(_record_markdown(record), encoding="utf-8")

    index = _load_index(root)
    index.append(_index_entry(record))
    _write_json(memory_index_path(root), index)
    summary = generate_memory_summary(record, ollama_ready=ollama_ready, base_dir=root)
    return {
        "status": "SAVED",
        "memory_dir": str(root),
        "index_path": str(memory_index_path(root)),
        "summary_path": summary["summary_path"],
        "record_json": str(json_path),
        "record_md": str(md_path),
        "entry": index[-1],
        "entries": len(index),
        "used_ollama": summary["used_ollama"],
        "ollama_model": summary.get("ollama_model", ""),
        "ollama_reason": summary.get("ollama_reason", ""),
    }
