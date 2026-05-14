from __future__ import annotations

from core.db import BrainzDatabase, SearchResult


class SearchEngine:
    def __init__(self, database: BrainzDatabase) -> None:
        self.database = database

    def search(self, query: str, limit: int = 40) -> list[SearchResult]:
        return self.database.search(query, limit=limit)

    def stats(self) -> dict[str, int]:
        return self.database.stats()
