from __future__ import annotations

import json
import re
import sqlite3
from dataclasses import dataclass
from pathlib import Path
from threading import Lock
from typing import Iterable

from core.app_config import db_path, ensure_app_dirs, now_iso


@dataclass(frozen=True)
class DocumentRecord:
    path: str
    title: str
    source_type: str
    created_at: str
    modified_at: str
    indexed_at: str
    hash: str
    content: str
    source_label: str = ""
    conversation_id: str = ""
    conversation_title: str = ""
    role: str = ""
    message_index: int = -1
    source_created_at: str = ""
    source_updated_at: str = ""
    codex_summary: str = ""
    changed_files_json: str = ""
    created_files_json: str = ""
    test_results: str = ""
    build_results: str = ""
    commit_hash: str = ""
    push_result: str = ""
    git_status: str = ""
    phase_notes: str = ""


@dataclass(frozen=True)
class SearchResult:
    id: int
    path: str
    title: str
    source_type: str
    modified_at: str
    indexed_at: str
    content: str
    snippet: str
    score: float
    source_label: str = ""
    conversation_id: str = ""
    conversation_title: str = ""
    role: str = ""
    message_index: int = -1
    source_created_at: str = ""
    source_updated_at: str = ""
    codex_summary: str = ""
    changed_files_json: str = ""
    created_files_json: str = ""
    test_results: str = ""
    build_results: str = ""
    commit_hash: str = ""
    push_result: str = ""
    git_status: str = ""
    phase_notes: str = ""


DOCUMENT_METADATA_COLUMNS = {
    "source_label": "TEXT NOT NULL DEFAULT ''",
    "conversation_id": "TEXT NOT NULL DEFAULT ''",
    "conversation_title": "TEXT NOT NULL DEFAULT ''",
    "role": "TEXT NOT NULL DEFAULT ''",
    "message_index": "INTEGER NOT NULL DEFAULT -1",
    "source_created_at": "TEXT NOT NULL DEFAULT ''",
    "source_updated_at": "TEXT NOT NULL DEFAULT ''",
}

CODEX_METADATA_COLUMNS = {
    "codex_summary": "TEXT NOT NULL DEFAULT ''",
    "changed_files_json": "TEXT NOT NULL DEFAULT ''",
    "created_files_json": "TEXT NOT NULL DEFAULT ''",
    "test_results": "TEXT NOT NULL DEFAULT ''",
    "build_results": "TEXT NOT NULL DEFAULT ''",
    "commit_hash": "TEXT NOT NULL DEFAULT ''",
    "push_result": "TEXT NOT NULL DEFAULT ''",
    "git_status": "TEXT NOT NULL DEFAULT ''",
    "phase_notes": "TEXT NOT NULL DEFAULT ''",
}


class BrainzDatabase:
    def __init__(self, path: Path | None = None) -> None:
        ensure_app_dirs()
        self.path = path or db_path()
        self._lock = Lock()

    def connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.path, timeout=30)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA synchronous=NORMAL")
        conn.execute("PRAGMA foreign_keys=ON")
        return conn

    def ensure_schema(self) -> None:
        with self._lock, self.connect() as conn:
            conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS documents (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    path TEXT NOT NULL UNIQUE,
                    title TEXT NOT NULL,
                    source_type TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    modified_at TEXT NOT NULL,
                    indexed_at TEXT NOT NULL,
                    hash TEXT NOT NULL,
                    content TEXT NOT NULL,
                    source_label TEXT NOT NULL DEFAULT '',
                    conversation_id TEXT NOT NULL DEFAULT '',
                    conversation_title TEXT NOT NULL DEFAULT '',
                    role TEXT NOT NULL DEFAULT '',
                    message_index INTEGER NOT NULL DEFAULT -1,
                    source_created_at TEXT NOT NULL DEFAULT '',
                    source_updated_at TEXT NOT NULL DEFAULT '',
                    codex_summary TEXT NOT NULL DEFAULT '',
                    changed_files_json TEXT NOT NULL DEFAULT '',
                    created_files_json TEXT NOT NULL DEFAULT '',
                    test_results TEXT NOT NULL DEFAULT '',
                    build_results TEXT NOT NULL DEFAULT '',
                    commit_hash TEXT NOT NULL DEFAULT '',
                    push_result TEXT NOT NULL DEFAULT '',
                    git_status TEXT NOT NULL DEFAULT '',
                    phase_notes TEXT NOT NULL DEFAULT ''
                );

                CREATE TABLE IF NOT EXISTS chunks (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    document_id INTEGER NOT NULL,
                    chunk_index INTEGER NOT NULL,
                    content TEXT NOT NULL,
                    summary TEXT NOT NULL DEFAULT '',
                    embedding_status TEXT NOT NULL DEFAULT 'pending',
                    source_type TEXT NOT NULL DEFAULT '',
                    source_label TEXT NOT NULL DEFAULT '',
                    conversation_id TEXT NOT NULL DEFAULT '',
                    conversation_title TEXT NOT NULL DEFAULT '',
                    role TEXT NOT NULL DEFAULT '',
                    message_index INTEGER NOT NULL DEFAULT -1,
                    source_created_at TEXT NOT NULL DEFAULT '',
                    source_updated_at TEXT NOT NULL DEFAULT '',
                    codex_summary TEXT NOT NULL DEFAULT '',
                    changed_files_json TEXT NOT NULL DEFAULT '',
                    created_files_json TEXT NOT NULL DEFAULT '',
                    test_results TEXT NOT NULL DEFAULT '',
                    build_results TEXT NOT NULL DEFAULT '',
                    commit_hash TEXT NOT NULL DEFAULT '',
                    push_result TEXT NOT NULL DEFAULT '',
                    git_status TEXT NOT NULL DEFAULT '',
                    phase_notes TEXT NOT NULL DEFAULT '',
                    FOREIGN KEY(document_id) REFERENCES documents(id) ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS codex_results (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    title TEXT NOT NULL,
                    summary TEXT NOT NULL DEFAULT '',
                    changed_files_json TEXT NOT NULL DEFAULT '',
                    created_files_json TEXT NOT NULL DEFAULT '',
                    test_results TEXT NOT NULL DEFAULT '',
                    build_results TEXT NOT NULL DEFAULT '',
                    commit_hash TEXT NOT NULL DEFAULT '',
                    push_result TEXT NOT NULL DEFAULT '',
                    git_status TEXT NOT NULL DEFAULT '',
                    phase_notes TEXT NOT NULL DEFAULT '',
                    raw_text TEXT NOT NULL,
                    content_hash TEXT NOT NULL,
                    import_key TEXT NOT NULL UNIQUE,
                    document_path TEXT NOT NULL,
                    imported_at TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS search_logs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    query TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    result_count INTEGER NOT NULL
                );

                CREATE TABLE IF NOT EXISTS decisions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    title TEXT NOT NULL,
                    note TEXT NOT NULL,
                    related_document_ids TEXT NOT NULL DEFAULT '',
                    created_at TEXT NOT NULL
                );

                CREATE VIRTUAL TABLE IF NOT EXISTS documents_fts USING fts5(
                    title,
                    path,
                    content,
                    tokenize='unicode61'
                );
                """
            )
            self._ensure_columns(conn, "documents", {**DOCUMENT_METADATA_COLUMNS, **CODEX_METADATA_COLUMNS})
            chunk_columns = {"source_type": "TEXT NOT NULL DEFAULT ''", **DOCUMENT_METADATA_COLUMNS, **CODEX_METADATA_COLUMNS}
            self._ensure_columns(conn, "chunks", chunk_columns)
            conn.commit()

    def _ensure_columns(self, conn: sqlite3.Connection, table: str, columns: dict[str, str]) -> None:
        existing = {
            str(row["name"])
            for row in conn.execute(f"PRAGMA table_info({table})").fetchall()
        }
        for name, definition in columns.items():
            if name not in existing:
                conn.execute(f"ALTER TABLE {table} ADD COLUMN {name} {definition}")

    def upsert_document(self, record: DocumentRecord, chunks: Iterable[str]) -> tuple[int, bool]:
        self.ensure_schema()
        with self._lock, self.connect() as conn:
            old = conn.execute(
                "SELECT id, hash FROM documents WHERE path = ?",
                (record.path,),
            ).fetchone()
            if old and old["hash"] == record.hash:
                return int(old["id"]), False

            values = (
                record.title,
                record.source_type,
                record.created_at,
                record.modified_at,
                record.indexed_at,
                record.hash,
                record.content,
                record.source_label,
                record.conversation_id,
                record.conversation_title,
                record.role,
                int(record.message_index),
                record.source_created_at,
                record.source_updated_at,
                record.codex_summary,
                record.changed_files_json,
                record.created_files_json,
                record.test_results,
                record.build_results,
                record.commit_hash,
                record.push_result,
                record.git_status,
                record.phase_notes,
            )

            if old:
                document_id = int(old["id"])
                conn.execute(
                    """
                    UPDATE documents
                    SET title = ?, source_type = ?, created_at = ?, modified_at = ?,
                        indexed_at = ?, hash = ?, content = ?, source_label = ?,
                        conversation_id = ?, conversation_title = ?, role = ?,
                        message_index = ?, source_created_at = ?, source_updated_at = ?,
                        codex_summary = ?, changed_files_json = ?, created_files_json = ?,
                        test_results = ?, build_results = ?, commit_hash = ?,
                        push_result = ?, git_status = ?, phase_notes = ?
                    WHERE id = ?
                    """,
                    (*values, document_id),
                )
                conn.execute("DELETE FROM chunks WHERE document_id = ?", (document_id,))
                conn.execute("DELETE FROM documents_fts WHERE rowid = ?", (document_id,))
            else:
                cursor = conn.execute(
                    """
                    INSERT INTO documents
                    (
                        title, source_type, created_at, modified_at, indexed_at, hash,
                        content, source_label, conversation_id, conversation_title, role,
                        message_index, source_created_at, source_updated_at, codex_summary,
                        changed_files_json, created_files_json, test_results, build_results,
                        commit_hash, push_result, git_status, phase_notes, path
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (*values, record.path),
                )
                document_id = int(cursor.lastrowid)

            for index, chunk in enumerate(chunks):
                conn.execute(
                    """
                    INSERT INTO chunks
                    (
                        document_id, chunk_index, content, embedding_status, source_type,
                        source_label, conversation_id, conversation_title, role, message_index,
                        source_created_at, source_updated_at, codex_summary, changed_files_json,
                        created_files_json, test_results, build_results, commit_hash, push_result,
                        git_status, phase_notes
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        document_id,
                        index,
                        chunk,
                        "pending",
                        record.source_type,
                        record.source_label,
                        record.conversation_id,
                        record.conversation_title,
                        record.role,
                        int(record.message_index),
                        record.source_created_at,
                        record.source_updated_at,
                        record.codex_summary,
                        record.changed_files_json,
                        record.created_files_json,
                        record.test_results,
                        record.build_results,
                        record.commit_hash,
                        record.push_result,
                        record.git_status,
                        record.phase_notes,
                    ),
                )

            conn.execute(
                "INSERT INTO documents_fts(rowid, title, path, content) VALUES (?, ?, ?, ?)",
                (document_id, record.title, record.path, record.content),
            )
            conn.commit()
            return document_id, True

    def upsert_codex_result(self, parsed: object, document_path: str) -> tuple[int, bool]:
        self.ensure_schema()
        commit_hash = str(getattr(parsed, "commit_hash", "") or "")
        content_hash = str(getattr(parsed, "content_hash", "") or "")
        import_key = commit_hash or content_hash
        with self._lock, self.connect() as conn:
            old = conn.execute(
                "SELECT id, content_hash FROM codex_results WHERE import_key = ?",
                (import_key,),
            ).fetchone()
            values = (
                str(getattr(parsed, "title", "") or ""),
                str(getattr(parsed, "summary", "") or ""),
                json.dumps(getattr(parsed, "changed_files", []) or [], ensure_ascii=False),
                json.dumps(getattr(parsed, "created_files", []) or [], ensure_ascii=False),
                str(getattr(parsed, "test_results", "") or ""),
                str(getattr(parsed, "build_results", "") or ""),
                commit_hash,
                str(getattr(parsed, "push_result", "") or ""),
                str(getattr(parsed, "git_status", "") or ""),
                str(getattr(parsed, "phase_notes", "") or ""),
                str(getattr(parsed, "raw_text", "") or ""),
                content_hash,
                import_key,
                document_path,
                str(getattr(parsed, "imported_at", "") or now_iso()),
            )
            if old:
                result_id = int(old["id"])
                if old["content_hash"] == content_hash:
                    return result_id, False
                conn.execute(
                    """
                    UPDATE codex_results
                    SET title = ?, summary = ?, changed_files_json = ?, created_files_json = ?,
                        test_results = ?, build_results = ?, commit_hash = ?, push_result = ?,
                        git_status = ?, phase_notes = ?, raw_text = ?, content_hash = ?,
                        import_key = ?, document_path = ?, imported_at = ?
                    WHERE id = ?
                    """,
                    (*values, result_id),
                )
            else:
                cursor = conn.execute(
                    """
                    INSERT INTO codex_results
                    (
                        title, summary, changed_files_json, created_files_json, test_results,
                        build_results, commit_hash, push_result, git_status, phase_notes,
                        raw_text, content_hash, import_key, document_path, imported_at
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    values,
                )
                result_id = int(cursor.lastrowid)
            conn.commit()
            return result_id, True

    def log_search(self, query: str, result_count: int) -> None:
        self.ensure_schema()
        with self._lock, self.connect() as conn:
            conn.execute(
                "INSERT INTO search_logs (query, created_at, result_count) VALUES (?, ?, ?)",
                (query, now_iso(), result_count),
            )
            conn.commit()

    def stats(self) -> dict[str, int]:
        self.ensure_schema()
        with self._lock, self.connect() as conn:
            documents = conn.execute("SELECT COUNT(*) AS count FROM documents").fetchone()["count"]
            chunks = conn.execute("SELECT COUNT(*) AS count FROM chunks").fetchone()["count"]
            searches = conn.execute("SELECT COUNT(*) AS count FROM search_logs").fetchone()["count"]
        return {"documents": int(documents), "chunks": int(chunks), "searches": int(searches)}

    def get_document(self, document_id: int) -> SearchResult | None:
        self.ensure_schema()
        with self._lock, self.connect() as conn:
            row = conn.execute(
                """
                SELECT
                    id, path, title, source_type, modified_at, indexed_at, content,
                    source_label, conversation_id, conversation_title, role, message_index,
                    source_created_at, source_updated_at, codex_summary, changed_files_json,
                    created_files_json, test_results, build_results, commit_hash, push_result,
                    git_status, phase_notes
                FROM documents
                WHERE id = ?
                """,
                (document_id,),
            ).fetchone()
        if row is None:
            return None
        return row_to_result(row, snippet=make_snippet(row["content"], ""), score=0.0)

    def search(self, query: str, limit: int = 40) -> list[SearchResult]:
        self.ensure_schema()
        query = query.strip()
        if not query:
            return []

        merged: dict[int, SearchResult] = {}
        fts_query = build_fts_query(query)
        if fts_query:
            try:
                for result in self._search_fts(fts_query, query, limit):
                    merged[result.id] = result
            except sqlite3.OperationalError:
                pass

        for result in self._search_like(query, limit):
            current = merged.get(result.id)
            if current is None or result.score > current.score:
                merged[result.id] = result

        results = sorted(merged.values(), key=lambda item: item.score, reverse=True)[:limit]
        self.log_search(query, len(results))
        return results

    def _search_fts(self, fts_query: str, original_query: str, limit: int) -> list[SearchResult]:
        with self._lock, self.connect() as conn:
            rows = conn.execute(
                """
                SELECT
                    d.id,
                    d.path,
                    d.title,
                    d.source_type,
                    d.modified_at,
                    d.indexed_at,
                    d.content,
                    d.source_label,
                    d.conversation_id,
                    d.conversation_title,
                    d.role,
                    d.message_index,
                    d.source_created_at,
                    d.source_updated_at,
                    d.codex_summary,
                    d.changed_files_json,
                    d.created_files_json,
                    d.test_results,
                    d.build_results,
                    d.commit_hash,
                    d.push_result,
                    d.git_status,
                    d.phase_notes,
                    snippet(documents_fts, 2, '[', ']', ' ... ', 28) AS snippet,
                    bm25(documents_fts) AS rank
                FROM documents_fts
                JOIN documents d ON d.id = documents_fts.rowid
                WHERE documents_fts MATCH ?
                ORDER BY rank
                LIMIT ?
                """,
                (fts_query, limit),
            ).fetchall()

        results: list[SearchResult] = []
        for row in rows:
            rank = float(row["rank"] or 0.0)
            results.append(
                row_to_result(
                    row,
                    snippet=row["snippet"] or make_snippet(row["content"], original_query),
                    score=100.0 - rank,
                )
            )
        return results

    def _search_like(self, query: str, limit: int) -> list[SearchResult]:
        terms = like_terms(query)
        if not terms:
            return []

        where_parts: list[str] = []
        params: list[str] = []
        for term in terms[:10]:
            where_parts.append("(title LIKE ? OR path LIKE ? OR content LIKE ?)")
            pattern = f"%{term}%"
            params.extend([pattern, pattern, pattern])

        sql = f"""
            SELECT
                id, path, title, source_type, modified_at, indexed_at, content,
                source_label, conversation_id, conversation_title, role, message_index,
                source_created_at, source_updated_at, codex_summary, changed_files_json,
                created_files_json, test_results, build_results, commit_hash, push_result,
                git_status, phase_notes
            FROM documents
            WHERE {' OR '.join(where_parts)}
            ORDER BY indexed_at DESC
            LIMIT ?
        """
        params.append(str(limit))
        with self._lock, self.connect() as conn:
            rows = conn.execute(sql, params).fetchall()

        results: list[SearchResult] = []
        for row in rows:
            score = like_score(row["title"], row["path"], row["content"], terms)
            results.append(row_to_result(row, snippet=make_snippet(row["content"], query), score=score))
        return results


def row_to_result(row: sqlite3.Row, snippet: str, score: float) -> SearchResult:
    return SearchResult(
        id=int(row["id"]),
        path=row["path"],
        title=row["title"],
        source_type=row["source_type"],
        modified_at=row["modified_at"],
        indexed_at=row["indexed_at"],
        content=row["content"],
        snippet=snippet,
        score=score,
        source_label=row["source_label"] or "",
        conversation_id=row["conversation_id"] or "",
        conversation_title=row["conversation_title"] or "",
        role=row["role"] or "",
        message_index=int(row["message_index"] if row["message_index"] is not None else -1),
        source_created_at=row["source_created_at"] or "",
        source_updated_at=row["source_updated_at"] or "",
        codex_summary=row["codex_summary"] or "",
        changed_files_json=row["changed_files_json"] or "",
        created_files_json=row["created_files_json"] or "",
        test_results=row["test_results"] or "",
        build_results=row["build_results"] or "",
        commit_hash=row["commit_hash"] or "",
        push_result=row["push_result"] or "",
        git_status=row["git_status"] or "",
        phase_notes=row["phase_notes"] or "",
    )


def build_fts_query(query: str) -> str:
    terms = like_terms(query)
    fts_terms: list[str] = []
    for term in terms[:8]:
        clean = term.replace('"', '""')
        if re.fullmatch(r"[A-Za-z0-9_][A-Za-z0-9_\-]*", clean):
            fts_terms.append(f"{clean}*")
        else:
            fts_terms.append(f'"{clean}"')
    return " OR ".join(fts_terms)


def like_terms(query: str) -> list[str]:
    cleaned = query.strip()
    if not cleaned:
        return []

    terms: list[str] = [cleaned]
    terms.extend(re.findall(r"[A-Za-z0-9_][A-Za-z0-9_\-]+", cleaned))
    terms.extend(part for part in re.split(r"\s+", cleaned) if part and part not in terms)

    unique: list[str] = []
    for term in terms:
        normalized = term.strip().strip('"').strip("'")
        if len(normalized) < 2:
            continue
        if normalized not in unique:
            unique.append(normalized)
    return unique


def like_score(title: str, path: str, content: str, terms: list[str]) -> float:
    target_title = title.lower()
    target_path = path.lower()
    target_content = content.lower()
    score = 0.0
    for term in terms:
        needle = term.lower()
        if needle in target_title:
            score += 30.0
        if needle in target_path:
            score += 14.0
        occurrences = target_content.count(needle)
        if occurrences:
            score += 5.0 + min(occurrences, 20)
    return score


def make_snippet(content: str, query: str, radius: int = 160) -> str:
    clean = " ".join(content.replace("\r", "\n").split())
    if not clean:
        return ""

    index = -1
    for term in like_terms(query):
        index = clean.lower().find(term.lower())
        if index >= 0:
            break

    if index < 0:
        return clean[: radius * 2]

    start = max(0, index - radius)
    end = min(len(clean), index + radius)
    prefix = "..." if start > 0 else ""
    suffix = "..." if end < len(clean) else ""
    return f"{prefix}{clean[start:end]}{suffix}"
