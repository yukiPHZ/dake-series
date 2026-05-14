from __future__ import annotations

import json
import math
from dataclasses import dataclass
from pathlib import Path
from threading import Event
from typing import Callable, Protocol
from urllib import error, request

from core.app_config import logs_dir, now_iso


OLLAMA_BASE_URL = "http://127.0.0.1:11434"
OLLAMA_EMBEDDINGS_URL = f"{OLLAMA_BASE_URL}/api/embeddings"
DEFAULT_EMBED_MODEL = "nomic-embed-text"
MAX_EMBED_TEXT_CHARS = 8000


class EmbeddingDatabase(Protocol):
    def chunk_rows_for_document(self, document_id: int, missing_only: bool = False) -> list[dict[str, object]]:
        ...

    def upsert_embedding(self, chunk_id: int, model_name: str, vector: list[float]) -> None:
        ...

    def mark_embedding_status(self, chunk_id: int, status: str) -> None:
        ...


@dataclass(frozen=True)
class EmbeddingResult:
    available: bool
    vector: list[float]
    message: str
    model_name: str


@dataclass(frozen=True)
class EmbeddingStatus:
    available: bool
    message: str
    model_name: str


@dataclass(frozen=True)
class EmbeddingRunResult:
    total: int
    generated: int
    skipped: int
    failed: int
    available: bool
    message: str
    log_path: str = ""


ProgressCallback = Callable[[int, int], None]


class EmbeddingSession:
    def __init__(self, model_name: str = DEFAULT_EMBED_MODEL) -> None:
        self.model_name = model_name
        self.unavailable_message = ""

    @property
    def available(self) -> bool:
        return not self.unavailable_message

    def mark_unavailable(self, message: str) -> None:
        self.unavailable_message = message or "embedding unavailable"


def embed_text(text: str, model_name: str = DEFAULT_EMBED_MODEL, timeout: float = 45.0) -> EmbeddingResult:
    clean = " ".join((text or "").replace("\r", "\n").split())
    if not clean:
        return EmbeddingResult(False, [], "empty text", model_name)

    payload = json.dumps({"model": model_name, "prompt": clean[:MAX_EMBED_TEXT_CHARS]}).encode("utf-8")
    req = request.Request(
        OLLAMA_EMBEDDINGS_URL,
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with request.urlopen(req, timeout=timeout) as response:
            data = json.loads(response.read().decode("utf-8", errors="replace"))
    except error.HTTPError as exc:
        return EmbeddingResult(False, [], _http_error_message(exc), model_name)
    except (OSError, error.URLError, json.JSONDecodeError) as exc:
        return EmbeddingResult(False, [], str(exc), model_name)

    embedding = data.get("embedding") if isinstance(data, dict) else None
    if not isinstance(embedding, list):
        return EmbeddingResult(False, [], "embedding response missing", model_name)
    try:
        vector = [float(value) for value in embedding]
    except (TypeError, ValueError):
        return EmbeddingResult(False, [], "embedding response invalid", model_name)
    if not vector:
        return EmbeddingResult(False, [], "embedding response empty", model_name)
    return EmbeddingResult(True, vector, "embedding ready", model_name)


def check_embedding_status(model_name: str = DEFAULT_EMBED_MODEL, timeout: float = 45.0) -> EmbeddingStatus:
    result = embed_text("brainz semantic ping", model_name=model_name, timeout=timeout)
    return EmbeddingStatus(result.available, result.message, model_name)


def generate_embeddings_for_document(
    database: EmbeddingDatabase,
    document_id: int,
    session: EmbeddingSession | None = None,
    cancel_event: Event | None = None,
    progress_callback: ProgressCallback | None = None,
) -> EmbeddingRunResult:
    session = session or EmbeddingSession()
    chunks = database.chunk_rows_for_document(document_id, missing_only=True)
    total = len(chunks)
    if total == 0:
        return EmbeddingRunResult(0, 0, 0, 0, session.available, session.unavailable_message or "no pending chunks")

    if session.unavailable_message:
        _mark_unavailable(database, chunks)
        return EmbeddingRunResult(total, 0, total, 0, False, session.unavailable_message)

    generated = 0
    skipped = 0
    failed = 0
    for index, chunk in enumerate(chunks, start=1):
        if cancel_event is not None and cancel_event.is_set():
            skipped += total - index + 1
            break
        if progress_callback:
            progress_callback(index, total)

        chunk_id = int(chunk["id"])
        content = str(chunk["content"] or "")
        result = embed_text(content, model_name=session.model_name)
        if not result.available:
            session.mark_unavailable(result.message)
            database.mark_embedding_status(chunk_id, "unavailable")
            failed += 1
            remaining = chunks[index:]
            _mark_unavailable(database, remaining)
            skipped += len(remaining)
            break

        database.upsert_embedding(chunk_id, result.model_name, result.vector)
        generated += 1

    available = not session.unavailable_message
    message = session.unavailable_message or "embedding generated"
    log_path = ""
    if generated or failed or skipped:
        log_path = str(
            write_semantic_log(
                [
                    f"Semantic embedding run: document_id={document_id}",
                    f"model={session.model_name}",
                    f"total={total}",
                    f"generated={generated}",
                    f"skipped={skipped}",
                    f"failed={failed}",
                    f"available={available}",
                    f"message={message}",
                ]
            )
        )
    return EmbeddingRunResult(total, generated, skipped, failed, available, message, log_path)


def cosine_similarity(left: list[float], right: list[float]) -> float:
    if not left or not right or len(left) != len(right):
        return 0.0
    dot = sum(a * b for a, b in zip(left, right))
    left_norm = math.sqrt(sum(a * a for a in left))
    right_norm = math.sqrt(sum(b * b for b in right))
    if left_norm == 0.0 or right_norm == 0.0:
        return 0.0
    return dot / (left_norm * right_norm)


def write_semantic_log(lines: list[str]) -> Path:
    logs_dir().mkdir(parents=True, exist_ok=True)
    path = logs_dir() / f"semantic_{now_iso().replace(':', '').replace('-', '').replace('T', '_')}.log"
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path


def _mark_unavailable(database: EmbeddingDatabase, chunks: list[dict[str, object]]) -> None:
    for chunk in chunks:
        database.mark_embedding_status(int(chunk["id"]), "unavailable")


def _http_error_message(exc: error.HTTPError) -> str:
    try:
        body = exc.read().decode("utf-8", errors="replace")
    except Exception:
        body = ""
    if body:
        return f"{exc.code} {body[:240]}"
    return str(exc)
