from __future__ import annotations

import json
import urllib.request
from pathlib import Path
from typing import Any

from core.app_config import load_config
from core.cli_checker import is_ollama_api_ready
from core.media_probe import MediaInfo

OLLAMA_BASE_URL = "http://localhost:11434"


def _post_json(path: str, payload: dict[str, Any], timeout: int = 30) -> dict[str, Any]:
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    request = urllib.request.Request(
        f"{OLLAMA_BASE_URL}{path}",
        data=data,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8", errors="replace"))


def _get_models(timeout: int = 3) -> list[str]:
    try:
        with urllib.request.urlopen(f"{OLLAMA_BASE_URL}/api/tags", timeout=timeout) as response:
            payload = json.loads(response.read().decode("utf-8", errors="replace"))
    except Exception:
        return []
    models = payload.get("models", [])
    return [str(model.get("name")) for model in models if model.get("name")]


def _read_transcript_excerpt(path: Path | None, limit: int = 1200) -> str:
    if not path or not path.exists():
        return ""
    text = path.read_text(encoding="utf-8", errors="replace").strip()
    return text[:limit]


def _fallback_metadata(source_name: str) -> dict[str, Any]:
    return {
        "used_ollama": False,
        "title_ideas": [
            "稼働中。夜に作る。",
            "止まらず作る。",
            "静かな作業机。",
        ],
        "description": (
            "制作記録です。\n\n"
            "補助脳で素材を整理し、YouTube投稿前のメタデータ雛形を作成しました。\n"
            "自動公開は行っていません。"
        ),
        "tags": ["制作記録", "作業ログ", "Dakeユキズ稼働中", "補助脳", "PEAKHEADZ"],
        "notes": [
            f"Source: {source_name}",
            "Review title, description, tags, and preview clip before upload.",
            "YouTube upload is not automated in Phase 1.",
        ],
    }


def build_metadata_draft(
    project_name: str,
    source_name: str,
    media_info: MediaInfo | None,
    transcript_path: Path | None,
) -> dict[str, Any]:
    fallback = _fallback_metadata(source_name)
    if not is_ollama_api_ready():
        return fallback

    config = load_config()
    models = _get_models()
    preferred = str(config.get("ollama_model") or "").strip()
    model = preferred if preferred else (models[0] if models else "")
    if not model:
        return fallback

    duration = f"{media_info.duration:.1f}s" if media_info else "unknown"
    excerpt = _read_transcript_excerpt(transcript_path)
    prompt = (
        "You are the local assistant brain for a quiet GPU production console.\n"
        "Create YouTube metadata draft in Japanese. Do not mention automatic upload.\n"
        "Tone: quiet, practical, production-base feeling.\n\n"
        f"Project: {project_name}\n"
        f"Source: {source_name}\n"
        f"Duration: {duration}\n"
        f"Transcript excerpt:\n{excerpt}\n\n"
        "Return JSON with keys: title_ideas (3 strings), description (string), tags (array), notes (array)."
    )
    try:
        response = _post_json("/api/generate", {"model": model, "prompt": prompt, "stream": False}, timeout=45)
        text = str(response.get("response") or "").strip()
        start = text.find("{")
        end = text.rfind("}")
        if start != -1 and end != -1 and end > start:
            parsed = json.loads(text[start : end + 1])
            fallback.update(parsed)
            fallback["used_ollama"] = True
            fallback["ollama_model"] = model
            return fallback
    except Exception:
        return fallback
    return fallback
