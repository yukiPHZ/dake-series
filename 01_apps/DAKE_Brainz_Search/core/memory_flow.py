from __future__ import annotations

import json
import re
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

from core.app_config import logs_dir, now_iso
from core.db import BrainzDatabase, SearchResult
from core.ollama_embeddings import DEFAULT_EMBED_MODEL, embed_text


@dataclass(frozen=True)
class MemoryFlowItem:
    result: SearchResult
    flow_score: float
    relation: str


@dataclass(frozen=True)
class MemoryFlowResponse:
    anchor_id: int
    items: list[MemoryFlowItem]
    semantic_available: bool
    semantic_message: str
    log_path: str


def generate_memory_flow(
    database: BrainzDatabase,
    anchor: SearchResult,
    semantic_enabled: bool = True,
    ascending: bool = True,
    limit: int = 10,
) -> MemoryFlowResponse:
    candidates = database.list_documents(limit=900)
    semantic_by_id: dict[int, SearchResult] = {}
    semantic_available = False
    semantic_message = "semantic disabled"

    if semantic_enabled:
        embedding = embed_text(anchor_text(anchor), model_name=DEFAULT_EMBED_MODEL)
        semantic_available = embedding.available
        semantic_message = embedding.message
        if embedding.available:
            for result in database.semantic_search(
                embedding.vector,
                embedding.model_name,
                limit=40,
                exclude_document_ids={anchor.id},
            ):
                semantic_by_id[result.id] = result

    scored: dict[int, MemoryFlowItem] = {
        anchor.id: MemoryFlowItem(anchor, 999.0, "anchor"),
    }
    for candidate in candidates:
        if candidate.id == anchor.id:
            continue
        score, relation = score_candidate(anchor, candidate, semantic_by_id.get(candidate.id))
        if score <= 0:
            continue
        scored[candidate.id] = MemoryFlowItem(
            semantic_by_id.get(candidate.id, candidate),
            score,
            relation,
        )

    ranked = sorted(scored.values(), key=lambda item: item.flow_score, reverse=True)[:limit]
    items = sorted(
        ranked,
        key=lambda item: timeline_sort_key(item.result),
        reverse=not ascending,
    )
    log_path = write_memory_flow_log(anchor, items, semantic_available, semantic_message)
    return MemoryFlowResponse(anchor.id, items, semantic_available, semantic_message, str(log_path))


def score_candidate(
    anchor: SearchResult,
    candidate: SearchResult,
    semantic_result: SearchResult | None = None,
) -> tuple[float, str]:
    score = 0.0
    relations: list[str] = []

    if anchor.conversation_id and candidate.conversation_id == anchor.conversation_id:
        score += 8.0
        relations.append("conversation")
        if anchor.message_index >= 0 and candidate.message_index >= 0:
            distance = abs(anchor.message_index - candidate.message_index)
            score += max(0.0, 3.0 - min(distance, 3) * 0.8)

    if anchor.commit_hash and candidate.commit_hash and candidate.commit_hash == anchor.commit_hash:
        score += 6.0
        relations.append("commit")

    changed_overlap = file_overlap(anchor.changed_files_json, candidate.changed_files_json)
    if changed_overlap:
        score += min(3.0, changed_overlap * 0.8)
        relations.append("files")

    if semantic_result and semantic_result.semantic_score > 0:
        score += semantic_result.semantic_score * 5.0
        relations.append("semantic")

    date_score = date_proximity_score(anchor, candidate)
    if date_score > 0:
        score += date_score
        relations.append("near_time")

    title_score = title_similarity(anchor.title, candidate.title)
    if title_score > 0:
        score += title_score * 3.0
        relations.append("title")

    if anchor.source_type and anchor.source_type == candidate.source_type:
        score += 0.9
        relations.append("source")

    return score, ", ".join(unique(relations))


def date_proximity_score(anchor: SearchResult, candidate: SearchResult) -> float:
    anchor_date = result_datetime(anchor)
    candidate_date = result_datetime(candidate)
    if not anchor_date or not candidate_date:
        return 0.0
    days = abs((anchor_date - candidate_date).total_seconds()) / 86400.0
    if days <= 1:
        return 3.0
    if days <= 7:
        return 2.2
    if days <= 30:
        return 1.4
    if days <= 180:
        return 0.6
    return 0.0


def title_similarity(left: str, right: str) -> float:
    left_tokens = set(title_tokens(left))
    right_tokens = set(title_tokens(right))
    if not left_tokens or not right_tokens:
        return 0.0
    overlap = len(left_tokens & right_tokens)
    union = len(left_tokens | right_tokens)
    return overlap / union if union else 0.0


def title_tokens(value: str) -> list[str]:
    text = (value or "").lower()
    tokens = re.findall(r"[a-z0-9_][a-z0-9_\-]+", text)
    tokens.extend(part for part in re.split(r"\s+", text) if len(part.strip()) >= 2)
    return unique(token.strip(" -_/") for token in tokens if token.strip(" -_/"))


def file_overlap(left_json: str, right_json: str) -> int:
    left = set(json_list(left_json))
    right = set(json_list(right_json))
    if not left or not right:
        return 0
    return len(left & right)


def json_list(value: str) -> list[str]:
    try:
        data = json.loads(value or "[]")
    except json.JSONDecodeError:
        return []
    if not isinstance(data, list):
        return []
    return [str(item).replace("\\", "/") for item in data if str(item).strip()]


def anchor_text(result: SearchResult) -> str:
    return "\n".join(
        [
            result.title,
            result.conversation_title,
            result.codex_summary,
            result.snippet,
            result.content[:2200],
        ]
    )


def result_datetime(result: SearchResult) -> datetime | None:
    for value in (result.source_created_at, result.modified_at, result.indexed_at, result.source_updated_at):
        parsed = parse_datetime(value)
        if parsed:
            return parsed
    return None


def timeline_sort_key(result: SearchResult) -> tuple[datetime, int, str]:
    value = result_datetime(result) or datetime.max.replace(tzinfo=timezone.utc)
    return value, result.message_index if result.message_index >= 0 else 999999, result.title


def timeline_date(result: SearchResult) -> str:
    parsed = result_datetime(result)
    if not parsed:
        return ""
    return parsed.date().isoformat()


def short_summary(result: SearchResult, limit: int = 180) -> str:
    text = result.codex_summary or result.snippet or result.content
    clean = " ".join((text or "").replace("\r", "\n").split())
    if len(clean) <= limit:
        return clean
    return clean[: max(0, limit - 3)] + "..."


def parse_datetime(value: str) -> datetime | None:
    text = str(value or "").strip()
    if not text:
        return None
    try:
        if text.endswith("Z"):
            text = f"{text[:-1]}+00:00"
        parsed = datetime.fromisoformat(text)
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=timezone.utc)
        return parsed
    except ValueError:
        return None


def unique(values) -> list[str]:
    seen: list[str] = []
    for value in values:
        clean = str(value or "").strip()
        if clean and clean not in seen:
            seen.append(clean)
    return seen


def write_memory_flow_log(
    anchor: SearchResult,
    items: list[MemoryFlowItem],
    semantic_available: bool,
    semantic_message: str,
) -> Path:
    logs_dir().mkdir(parents=True, exist_ok=True)
    path = logs_dir() / f"memory_flow_{now_iso().replace(':', '').replace('-', '').replace('T', '_')}.log"
    lines = [
        "Memory Flow generated.",
        f"anchor_id={anchor.id}",
        f"anchor_title={anchor.title}",
        f"semantic_available={semantic_available}",
        f"semantic_message={semantic_message}",
        f"items={len(items)}",
    ]
    for item in items:
        lines.append(
            f"- {timeline_date(item.result)} | {item.result.source_type} | "
            f"{item.result.title} | score={item.flow_score:.2f} | relation={item.relation}"
        )
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path
