from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from threading import Event
from typing import Callable

from core.app_config import logs_dir, now_iso
from core.db import BrainzDatabase, DocumentRecord
from core.file_scanner import iter_supported_files, scan_file
from core.ollama_embeddings import EmbeddingSession, generate_embeddings_for_document
from core.text_splitter import split_text


@dataclass(frozen=True)
class IndexProgress:
    total: int
    current: int
    indexed: int
    skipped: int
    errors: int
    message: str
    done: bool = False
    cancelled: bool = False
    log_path: str = ""


ProgressCallback = Callable[[IndexProgress], None]


class Indexer:
    def __init__(self, database: BrainzDatabase) -> None:
        self.database = database

    def run(
        self,
        memory_folder: Path,
        cancel_event: Event,
        progress_callback: ProgressCallback | None = None,
    ) -> IndexProgress:
        self.database.ensure_schema()
        started_at = now_iso()
        log_lines = [
            f"BRAINZ INDEX START {started_at}",
            f"memory_folder={memory_folder}",
        ]

        files = iter_supported_files(memory_folder, cancel_event=cancel_event)
        total = len(files)
        indexed = 0
        skipped = 0
        errors = 0
        embedding_session = EmbeddingSession()

        self._emit(progress_callback, total, 0, indexed, skipped, errors, "scan_ready")

        for current, path in enumerate(files, start=1):
            if cancel_event.is_set():
                log_lines.append(f"CANCELLED current={current - 1} total={total}")
                log_path = self._write_log(log_lines)
                progress = IndexProgress(
                    total=total,
                    current=current - 1,
                    indexed=indexed,
                    skipped=skipped,
                    errors=errors,
                    message="cancelled",
                    done=True,
                    cancelled=True,
                    log_path=str(log_path),
                )
                if progress_callback:
                    progress_callback(progress)
                return progress

            try:
                scanned = scan_file(path)
                chunks = split_text(scanned.content)
                record = DocumentRecord(
                    path=str(scanned.path),
                    title=scanned.title,
                    source_type=scanned.source_type,
                    created_at=scanned.created_at,
                    modified_at=scanned.modified_at,
                    indexed_at=now_iso(),
                    hash=scanned.content_hash,
                    content=scanned.content,
                )
                document_id, changed = self.database.upsert_document(record, chunks)
                if changed:
                    indexed += 1
                    log_lines.append(f"INDEXED {scanned.path}")
                    self._emit(progress_callback, total, current, indexed, skipped, errors, "embedding_chunks")
                    try:
                        embedding_result = generate_embeddings_for_document(
                            self.database,
                            document_id,
                            session=embedding_session,
                            cancel_event=cancel_event,
                        )
                        log_lines.append(
                            "EMBEDDING "
                            f"generated={embedding_result.generated} skipped={embedding_result.skipped} "
                            f"failed={embedding_result.failed} available={embedding_result.available}"
                        )
                    except Exception as exc:
                        log_lines.append(f"EMBEDDING_ERROR {scanned.path} :: {exc}")
                else:
                    skipped += 1
                    log_lines.append(f"SKIPPED_UNCHANGED {scanned.path}")
            except Exception as exc:
                errors += 1
                log_lines.append(f"ERROR {path} :: {exc}")

            self._emit(progress_callback, total, current, indexed, skipped, errors, str(path))

        finished_at = now_iso()
        log_lines.extend(
            [
                f"BRAINZ INDEX END {finished_at}",
                f"indexed={indexed}",
                f"skipped={skipped}",
                f"errors={errors}",
            ]
        )
        log_path = self._write_log(log_lines)
        progress = IndexProgress(
            total=total,
            current=total,
            indexed=indexed,
            skipped=skipped,
            errors=errors,
            message="done",
            done=True,
            cancelled=False,
            log_path=str(log_path),
        )
        if progress_callback:
            progress_callback(progress)
        return progress

    def _emit(
        self,
        callback: ProgressCallback | None,
        total: int,
        current: int,
        indexed: int,
        skipped: int,
        errors: int,
        message: str,
    ) -> None:
        if callback:
            callback(
                IndexProgress(
                    total=total,
                    current=current,
                    indexed=indexed,
                    skipped=skipped,
                    errors=errors,
                    message=message,
                )
            )

    def _write_log(self, lines: list[str]) -> Path:
        logs_dir().mkdir(parents=True, exist_ok=True)
        path = logs_dir() / f"index_{now_iso().replace(':', '').replace('-', '').replace('T', '_')}.log"
        path.write_text("\n".join(lines) + "\n", encoding="utf-8")
        return path
