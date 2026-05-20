from __future__ import annotations

import hashlib
import json
import os
from dataclasses import asdict, dataclass, replace
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib import error, request

from core.config import app_dir, series_root


OLLAMA_BASE_URL = "http://127.0.0.1:11434"
OLLAMA_TAGS_URL = f"{OLLAMA_BASE_URL}/api/tags"
OLLAMA_GENERATE_URL = f"{OLLAMA_BASE_URL}/api/generate"
HEAT_HINTS_FILE_NAME = "qpsc_heat_hints.json"
OLLAMA_HEAT_SOURCE_LIMIT = 700
OLLAMA_HEAT_RESPONSE_LIMIT = 80
CONFIDENCE_VALUES = {"low", "medium", "high"}
MODEL_SKIP_MARKERS = ("embed", "embedding", "bert")


@dataclass(frozen=True)
class HeatHintInput:
    title: str
    message: str
    source: str
    related_path: str
    source_excerpt: str = ""


@dataclass(frozen=True)
class HeatHintResult:
    ok: bool
    key: str
    generated_at: str
    heat_hint: str
    reason: str
    confidence: str
    related_path: str
    title: str
    source: str
    model: str = ""
    status: str = ""
    cached: bool = False


def heat_hints_candidates() -> list[Path]:
    apps_dir = app_dir().parent
    return [
        apps_dir / "DAKE_Brainz_Search" / "data" / "config" / HEAT_HINTS_FILE_NAME,
        series_root() / "01_apps" / "DAKE_Brainz_Search" / "data" / "config" / HEAT_HINTS_FILE_NAME,
        app_dir() / "data" / "config" / HEAT_HINTS_FILE_NAME,
    ]


def heat_hints_path() -> Path:
    candidates = heat_hints_candidates()
    for path in candidates:
        if path.exists():
            return path
    return candidates[0]


def request_heat_hint(input_item: HeatHintInput, cache_path: Path | None = None, timeout: float = 20.0) -> HeatHintResult:
    cached = read_cached_heat_hint(input_item, cache_path=cache_path)
    if cached is not None:
        return cached

    model = resolve_ollama_model(timeout=min(timeout, 3.0))
    if not model:
        return _result(input_item, ok=False, status="ollama_unavailable")

    prompt = build_heat_prompt(input_item)
    response = _post_json(
        OLLAMA_GENERATE_URL,
        {"model": model, "prompt": prompt, "stream": False},
        timeout=timeout,
    )
    if not isinstance(response, dict):
        return _result(input_item, ok=False, status="ollama_unavailable", model=model)

    text = str(response.get("response", "") or "").strip()
    parsed = parse_heat_response(text)
    if parsed is None:
        return _result(input_item, ok=False, status="invalid_response", model=model)

    result = _result(
        input_item,
        ok=True,
        heat_hint=_short_text(parsed.get("heat_hint"), OLLAMA_HEAT_RESPONSE_LIMIT),
        reason=_short_text(parsed.get("reason"), OLLAMA_HEAT_RESPONSE_LIMIT),
        confidence=normalize_confidence(parsed.get("confidence")),
        model=model,
        status="ok",
    )
    save_heat_hint(result, cache_path=cache_path)
    return result


def resolve_ollama_model(timeout: float = 3.0) -> str:
    configured = os.environ.get("QPSC_OLLAMA_MODEL", "").strip() or os.environ.get("OLLAMA_MODEL", "").strip()
    if configured:
        return configured
    response = _get_json(OLLAMA_TAGS_URL, timeout=timeout)
    models = response.get("models") if isinstance(response, dict) else None
    if not isinstance(models, list):
        return ""
    names = [str(item.get("name", "") or "").strip() for item in models if isinstance(item, dict)]
    preferred = [
        name
        for name in names
        if name and not any(marker in name.lower() for marker in MODEL_SKIP_MARKERS)
    ]
    return next(iter(preferred), next((name for name in names if name), ""))


def build_heat_prompt(input_item: HeatHintInput) -> str:
    source_excerpt = clean_text(input_item.source_excerpt)[:OLLAMA_HEAT_SOURCE_LIMIT]
    title = clean_text(input_item.title)
    message = clean_text(input_item.message)
    source = clean_text(input_item.source)
    related_path = clean_text(input_item.related_path)
    return "\n".join(
        [
            "あなたはQPSCの静かな補助です。",
            "AIは判断を代行しません。命令も断定もしません。",
            "熾火候補に、まだ熱が残っている気配があるかを弱く補助してください。",
            "未完了の熱、再訪、作業の続き、気になっている断片だけを短く見ます。",
            "ユーザーに「やるべき」と言わないでください。",
            "日本語で、次のJSONだけを返してください。",
            '{"heat_hint":"まだ熱が残っているかも","reason":"最近の作業記録と関係しています","confidence":"low"}',
            "",
            f"title: {title}",
            f"message: {message}",
            f"source: {source}",
            f"related_path: {related_path}",
            "source_excerpt:",
            source_excerpt,
        ]
    )


def parse_heat_response(text: str) -> dict[str, str] | None:
    candidate = _json_text(text)
    try:
        payload = json.loads(candidate)
    except (TypeError, json.JSONDecodeError):
        return parse_freeform_heat_response(text)
    if not isinstance(payload, dict):
        return parse_freeform_heat_response(text)
    heat_hint = _short_text(payload.get("heat_hint"), OLLAMA_HEAT_RESPONSE_LIMIT)
    reason = _short_text(payload.get("reason"), OLLAMA_HEAT_RESPONSE_LIMIT)
    confidence = normalize_confidence(payload.get("confidence"))
    if not heat_hint or not reason:
        return None
    return {"heat_hint": heat_hint, "reason": reason, "confidence": confidence}


def parse_freeform_heat_response(text: str) -> dict[str, str] | None:
    lines = [clean_text(line.strip(" -*:：")) for line in str(text or "").splitlines()]
    lines = [line for line in lines if line]
    if not lines:
        clean = clean_text(text)
        if not clean:
            return None
        lines = [clean]
    confidence = "low"
    lowered = " ".join(lines).lower()
    if "high" in lowered:
        confidence = "high"
    elif "medium" in lowered:
        confidence = "medium"
    heat_hint = _short_text(lines[0], OLLAMA_HEAT_RESPONSE_LIMIT)
    reason = _short_text(lines[1] if len(lines) > 1 else "短い応答から補助しました", OLLAMA_HEAT_RESPONSE_LIMIT)
    if not heat_hint or not reason:
        return None
    return {"heat_hint": heat_hint, "reason": reason, "confidence": confidence}


def read_cached_heat_hint(input_item: HeatHintInput, cache_path: Path | None = None) -> HeatHintResult | None:
    key = heat_hint_key(input_item)
    for item in read_heat_hint_cache(cache_path=cache_path):
        if str(item.get("key", "") or "") != key:
            continue
        result = _result_from_cache(item)
        if result is not None:
            return replace(result, cached=True)
    return None


def save_heat_hint(result: HeatHintResult, cache_path: Path | None = None) -> Path:
    path = cache_path or heat_hints_path()
    records = [item for item in read_heat_hint_cache(cache_path=path) if str(item.get("key", "") or "") != result.key]
    records.append(asdict(result))
    records.sort(key=lambda item: str(item.get("generated_at", "") or ""), reverse=True)
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_suffix(path.suffix + ".tmp")
    temp_path.write_text(json.dumps(records[:200], ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    temp_path.replace(path)
    return path


def read_heat_hint_cache(cache_path: Path | None = None) -> list[dict[str, Any]]:
    path = cache_path or heat_hints_path()
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    if not isinstance(payload, list):
        return []
    return [item for item in payload if isinstance(item, dict)]


def heat_hint_key(input_item: HeatHintInput) -> str:
    related_path = clean_text(input_item.related_path)
    title = clean_text(input_item.title)
    source = clean_text(input_item.source)
    message = clean_text(input_item.message)
    seed = related_path or f"{source}\n{title}\n{message}"
    return hashlib.sha256(seed.encode("utf-8")).hexdigest()


def clean_text(value: object) -> str:
    return " ".join(str(value or "").replace("\r", "\n").split())


def normalize_confidence(value: object) -> str:
    confidence = str(value or "").strip().lower()
    return confidence if confidence in CONFIDENCE_VALUES else "low"


def _result(
    input_item: HeatHintInput,
    *,
    ok: bool,
    heat_hint: str = "",
    reason: str = "",
    confidence: str = "low",
    model: str = "",
    status: str = "",
) -> HeatHintResult:
    return HeatHintResult(
        ok=ok,
        key=heat_hint_key(input_item),
        generated_at=datetime.now().astimezone().isoformat(timespec="seconds"),
        heat_hint=heat_hint,
        reason=reason,
        confidence=normalize_confidence(confidence),
        related_path=clean_text(input_item.related_path),
        title=clean_text(input_item.title),
        source=clean_text(input_item.source),
        model=model,
        status=status,
    )


def _result_from_cache(item: dict[str, Any]) -> HeatHintResult | None:
    heat_hint = _short_text(item.get("heat_hint"), OLLAMA_HEAT_RESPONSE_LIMIT)
    reason = _short_text(item.get("reason"), OLLAMA_HEAT_RESPONSE_LIMIT)
    if not heat_hint or not reason:
        return None
    return HeatHintResult(
        ok=True,
        key=str(item.get("key", "") or ""),
        generated_at=str(item.get("generated_at", "") or ""),
        heat_hint=heat_hint,
        reason=reason,
        confidence=normalize_confidence(item.get("confidence")),
        related_path=str(item.get("related_path", "") or ""),
        title=str(item.get("title", "") or ""),
        source=str(item.get("source", "") or ""),
        model=str(item.get("model", "") or ""),
        status=str(item.get("status", "cached") or "cached"),
        cached=True,
    )


def _short_text(value: object, limit: int) -> str:
    text = clean_text(value)
    if len(text) <= limit:
        return text
    return text[:limit].rstrip()


def _json_text(text: str) -> str:
    clean = str(text or "").strip()
    if clean.startswith("```"):
        clean = clean.strip("`").strip()
        if clean.lower().startswith("json"):
            clean = clean[4:].strip()
    start = clean.find("{")
    end = clean.rfind("}")
    if start >= 0 and end > start:
        return clean[start : end + 1]
    return clean


def _get_json(url: str, timeout: float) -> dict[str, Any]:
    try:
        with request.urlopen(url, timeout=timeout) as response:
            payload = json.loads(response.read().decode("utf-8", errors="replace"))
    except (OSError, error.URLError, json.JSONDecodeError):
        return {}
    return payload if isinstance(payload, dict) else {}


def _post_json(url: str, payload: dict[str, Any], timeout: float) -> dict[str, Any]:
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = request.Request(
        url,
        data=data,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with request.urlopen(req, timeout=timeout) as response:
            decoded = json.loads(response.read().decode("utf-8", errors="replace"))
    except (OSError, error.URLError, json.JSONDecodeError):
        return {}
    return decoded if isinstance(decoded, dict) else {}
