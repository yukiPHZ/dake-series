from __future__ import annotations

from dataclasses import dataclass, replace

from core.db import BrainzDatabase, SearchResult
from core.memory_flow import MemoryFlowResponse, generate_memory_flow
from core.ollama_embeddings import DEFAULT_EMBED_MODEL, embed_text, write_semantic_log


@dataclass(frozen=True)
class SearchResponse:
    results: list[SearchResult]
    related: list[SearchResult]
    semantic_available: bool
    semantic_message: str


class SearchEngine:
    def __init__(self, database: BrainzDatabase) -> None:
        self.database = database

    def search(self, query: str, limit: int = 40) -> list[SearchResult]:
        return self.database.search(query, limit=limit)

    def search_with_related(
        self,
        query: str,
        limit: int = 40,
        related_limit: int = 5,
        semantic_enabled: bool = True,
    ) -> SearchResponse:
        results = self.database.search(query, limit=limit)
        if not semantic_enabled:
            return SearchResponse(results, [], False, "semantic search disabled")

        embedding = embed_text(query, model_name=DEFAULT_EMBED_MODEL)
        if not embedding.available:
            write_semantic_log(
                [
                    "Semantic search initialized.",
                    "semantic_available=False",
                    f"message={embedding.message}",
                ]
            )
            return SearchResponse(results, [], False, embedding.message)

        semantic_results = self.database.semantic_search(
            embedding.vector,
            embedding.model_name,
            limit=max(limit, related_limit),
        )
        merged: dict[int, SearchResult] = {result.id: result for result in results}
        for result in semantic_results:
            current = merged.get(result.id)
            if current is None:
                merged[result.id] = result
            elif result.semantic_score > current.semantic_score:
                merged[result.id] = replace(current, semantic_score=result.semantic_score)

        combined = sorted(
            merged.values(),
            key=lambda item: (max(item.score, item.semantic_score * 100.0), item.semantic_score),
            reverse=True,
        )[:limit]
        related = semantic_results[:related_limit]
        write_semantic_log(
            [
                "Semantic search initialized.",
                "Related memory found.",
                f"semantic_available=True",
                f"related_count={len(related)}",
            ]
        )
        return SearchResponse(combined, related, True, "semantic search ready")

    def stats(self) -> dict[str, int]:
        return self.database.stats()

    def memory_flow(
        self,
        anchor: SearchResult,
        semantic_enabled: bool = True,
        ascending: bool = True,
        limit: int = 10,
    ) -> MemoryFlowResponse:
        return generate_memory_flow(
            self.database,
            anchor,
            semantic_enabled=semantic_enabled,
            ascending=ascending,
            limit=limit,
        )
