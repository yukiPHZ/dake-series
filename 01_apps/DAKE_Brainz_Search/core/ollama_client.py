from __future__ import annotations

import json
from dataclasses import dataclass
from typing import Any
from urllib import error, request


OLLAMA_BASE_URL = "http://127.0.0.1:11434"
DEFAULT_EMBED_MODEL = "nomic-embed-text"


@dataclass(frozen=True)
class OllamaStatus:
    available: bool
    message: str
    models: list[str]


def check_ollama(timeout: float = 1.2) -> OllamaStatus:
    try:
        with request.urlopen(f"{OLLAMA_BASE_URL}/api/tags", timeout=timeout) as response:
            payload = json.loads(response.read().decode("utf-8", errors="replace"))
    except (OSError, error.URLError, json.JSONDecodeError) as exc:
        return OllamaStatus(available=False, message=str(exc), models=[])

    models = []
    for item in payload.get("models", []):
        if isinstance(item, dict) and item.get("name"):
            models.append(str(item["name"]))
    return OllamaStatus(available=True, message="ready", models=models)


def embed_text(text: str, model: str = DEFAULT_EMBED_MODEL, timeout: float = 8.0) -> list[float] | None:
    payload = json.dumps({"model": model, "prompt": text}).encode("utf-8")
    req = request.Request(
        f"{OLLAMA_BASE_URL}/api/embeddings",
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with request.urlopen(req, timeout=timeout) as response:
            data: dict[str, Any] = json.loads(response.read().decode("utf-8", errors="replace"))
    except (OSError, error.URLError, json.JSONDecodeError):
        return None

    embedding = data.get("embedding")
    if not isinstance(embedding, list):
        return None
    try:
        return [float(value) for value in embedding]
    except (TypeError, ValueError):
        return None
