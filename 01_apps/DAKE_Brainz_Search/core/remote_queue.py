from __future__ import annotations

import hashlib
import json
import re
from dataclasses import dataclass
from pathlib import Path

from core.app_config import now_iso, read_text_safe
from core.db import BrainzDatabase, DocumentRecord
from core.file_scanner import SUPPORTED_EXTENSIONS, scan_file
from core.ollama_embeddings import EmbeddingSession, generate_embeddings_for_document
from core.text_splitter import split_text


QUEUE_EXTENSIONS = {".md", ".txt", ".json"}
QUEUE_SYSTEM_DIRS = {"processed", "failed"}
TASK_TYPES = {"import", "search", "handoff_chatgpt", "handoff_codex", "note"}


@dataclass(frozen=True)
class RemoteQueueTask:
    task_type: str
    query: str
    note: str
    import_path: str
    source_file: str
    raw_text: str
    title: str


@dataclass(frozen=True)
class RemoteQueueTaskResult:
    task_type: str
    query: str
    note: str
    source_file: str
    destination_file: str
    status: str
    raw_text: str
    error: str = ""
    document_id: int = 0
    changed: bool = False


@dataclass(frozen=True)
class RemoteQueueBatchResult:
    detected: int
    processed: int
    failed: int
    pending: int
    results: list[RemoteQueueTaskResult]


def iter_queue_files(queue_folder: Path) -> list[Path]:
    if not queue_folder.exists() or not queue_folder.is_dir():
        return []

    files: list[Path] = []
    for path in sorted(queue_folder.rglob("*"), key=lambda item: str(item).lower()):
        if not path.is_file():
            continue
        try:
            relative_parts = tuple(part.lower() for part in path.relative_to(queue_folder).parts)
        except ValueError:
            relative_parts = tuple(part.lower() for part in path.parts)
        if any(part in QUEUE_SYSTEM_DIRS for part in relative_parts):
            continue
        if path.suffix.lower() in QUEUE_EXTENSIONS:
            files.append(path)
    return files


def count_pending_queue_files(queue_folder: Path) -> int:
    return len(iter_queue_files(queue_folder))


def parse_remote_queue_file(path: Path) -> RemoteQueueTask:
    raw_text = read_text_safe(path)
    if not raw_text.strip():
        raise ValueError("empty remote queue task")

    values: dict[str, str] = {}
    title = path.stem
    if path.suffix.lower() == ".json":
        try:
            data = json.loads(raw_text)
        except json.JSONDecodeError as exc:
            raise ValueError(f"invalid json: {exc}") from exc
        if not isinstance(data, dict):
            raise ValueError("json task must be an object")
        for key, value in data.items():
            if value is None:
                continue
            values[str(key).strip().lower()] = str(value).strip()
        title = values.get("title") or title
    else:
        for line in raw_text.splitlines():
            clean = line.strip().strip("-").strip()
            if not clean:
                continue
            if clean.startswith("#") and title == path.stem:
                title = clean.lstrip("#").strip() or title
                continue
            match = re.match(r"^([A-Za-z_][A-Za-z0-9_\- ]{0,48})\s*:\s*(.+)$", clean)
            if not match:
                continue
            key = match.group(1).strip().lower().replace("-", "_").replace(" ", "_")
            value = match.group(2).strip()
            if key in TASK_TYPES and "type" not in values:
                values["type"] = key
                if key in {"search", "handoff_chatgpt", "handoff_codex"} and "query" not in values:
                    values["query"] = value
                elif key == "note" and "note" not in values:
                    values["note"] = value
                elif key == "import" and "path" not in values:
                    values["path"] = value
                continue
            values[key] = value

    task_type = normalize_task_type(values.get("type") or values.get("task_type") or "")
    query = values.get("query", "")
    note = values.get("note", "")
    import_path = values.get("file_path") or values.get("import_path") or values.get("path") or values.get("file") or ""

    if not task_type:
        task_type = "search" if query else "note"
    if task_type not in TASK_TYPES:
        raise ValueError(f"unsupported task type: {task_type}")
    if task_type in {"search", "handoff_chatgpt", "handoff_codex"} and not query:
        raise ValueError("query is required")
    if task_type == "import" and not import_path:
        raise ValueError("import path is required")

    return RemoteQueueTask(
        task_type=task_type,
        query=query,
        note=note,
        import_path=import_path,
        source_file=str(path.resolve()),
        raw_text=raw_text,
        title=title,
    )


def normalize_task_type(value: str) -> str:
    clean = value.strip().lower().replace("-", "_").replace(" ", "_")
    aliases = {
        "chatgpt_handoff": "handoff_chatgpt",
        "handoff_gpt": "handoff_chatgpt",
        "codex_handoff": "handoff_codex",
    }
    return aliases.get(clean, clean)


def process_remote_queue_folder(database: BrainzDatabase, queue_folder: Path, limit: int = 20) -> RemoteQueueBatchResult:
    queue_folder.mkdir(parents=True, exist_ok=True)
    files = iter_queue_files(queue_folder)[:limit]
    results: list[RemoteQueueTaskResult] = []
    embedding_session = EmbeddingSession()

    for source_file in files:
        try:
            task = parse_remote_queue_file(source_file)
            processed_destination = destination_for(queue_folder, source_file, "processed")
            document_id, changed = execute_task(database, task, processed_destination, embedding_session)
            destination = move_task_file(queue_folder, source_file, "processed", processed_destination)
            result = RemoteQueueTaskResult(
                task_type=task.task_type,
                query=task.query,
                note=task.note,
                source_file=task.source_file,
                destination_file=str(destination),
                status="processed",
                raw_text=task.raw_text,
                document_id=document_id,
                changed=changed,
            )
        except Exception as exc:
            raw_text = read_text_safe(source_file)
            task_type, query, note = task_hint(raw_text)
            destination = move_task_file(queue_folder, source_file, "failed")
            result = RemoteQueueTaskResult(
                task_type=task_type,
                query=query,
                note=note,
                source_file=str(source_file.resolve()),
                destination_file=str(destination),
                status="failed",
                raw_text=raw_text,
                error=str(exc),
            )

        database.log_remote_queue_task(
            task_type=result.task_type,
            query=result.query,
            note=result.note,
            source_file=result.source_file,
            status=result.status,
            raw_text=result.raw_text,
        )
        results.append(result)

    processed = sum(1 for result in results if result.status == "processed")
    failed = sum(1 for result in results if result.status == "failed")
    return RemoteQueueBatchResult(
        detected=len(files),
        processed=processed,
        failed=failed,
        pending=count_pending_queue_files(queue_folder),
        results=results,
    )


def execute_task(
    database: BrainzDatabase,
    task: RemoteQueueTask,
    document_path: Path,
    embedding_session: EmbeddingSession,
) -> tuple[int, bool]:
    if task.task_type == "note":
        return index_note_task(database, task, document_path, embedding_session)
    if task.task_type == "import":
        return index_import_task(database, task, embedding_session)
    return 0, False


def index_note_task(
    database: BrainzDatabase,
    task: RemoteQueueTask,
    document_path: Path,
    embedding_session: EmbeddingSession,
) -> tuple[int, bool]:
    content = task.raw_text.strip()
    digest = hashlib.sha256(content.encode("utf-8", errors="replace")).hexdigest()
    indexed_at = now_iso()
    record = DocumentRecord(
        path=str(document_path.resolve()),
        title=task.title or "Remote Queue Note",
        source_type="remote_queue_note",
        created_at=indexed_at,
        modified_at=indexed_at,
        indexed_at=indexed_at,
        hash=digest,
        content=content,
        source_label="Remote Queue / note",
        source_created_at=indexed_at,
        source_updated_at=indexed_at,
    )
    document_id, changed = database.upsert_document(record, split_text(content))
    if changed:
        generate_embeddings_for_document(database, document_id, session=embedding_session)
    return document_id, changed


def index_import_task(
    database: BrainzDatabase,
    task: RemoteQueueTask,
    embedding_session: EmbeddingSession,
) -> tuple[int, bool]:
    if re.match(r"^https?://", task.import_path, flags=re.IGNORECASE):
        raise ValueError("remote import URLs are not allowed")
    source_path = Path(task.import_path)
    if not source_path.is_absolute():
        source_path = Path(task.source_file).parent / source_path
    source_path = source_path.resolve()
    if not source_path.exists() or not source_path.is_file():
        raise FileNotFoundError(str(source_path))
    if source_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
        raise ValueError("unsupported import file type")

    scanned = scan_file(source_path)
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
    document_id, changed = database.upsert_document(record, split_text(scanned.content))
    if changed:
        generate_embeddings_for_document(database, document_id, session=embedding_session)
    return document_id, changed


def task_hint(raw_text: str) -> tuple[str, str, str]:
    try:
        text = raw_text.strip()
        if not text:
            return "unknown", "", ""
        if text.startswith("{"):
            data = json.loads(text)
            if isinstance(data, dict):
                return (
                    normalize_task_type(str(data.get("type") or data.get("task_type") or "unknown")),
                    str(data.get("query") or ""),
                    str(data.get("note") or ""),
                )
        parsed = {}
        for line in text.splitlines():
            match = re.match(r"^([A-Za-z_][A-Za-z0-9_\- ]{0,48})\s*:\s*(.+)$", line.strip())
            if match:
                parsed[match.group(1).strip().lower().replace("-", "_").replace(" ", "_")] = match.group(2).strip()
        return (
            normalize_task_type(parsed.get("type", "unknown")),
            parsed.get("query", ""),
            parsed.get("note", ""),
        )
    except Exception:
        return "unknown", "", ""


def move_task_file(queue_folder: Path, source_file: Path, status: str, destination: Path | None = None) -> Path:
    destination = destination or destination_for(queue_folder, source_file, status)
    destination.parent.mkdir(parents=True, exist_ok=True)
    source_file.replace(destination)
    return destination


def destination_for(queue_folder: Path, source_file: Path, status: str) -> Path:
    target_dir = queue_folder / status
    target = target_dir / source_file.name
    if not target.exists():
        return target
    stamp = now_iso().replace(":", "").replace("-", "").replace("T", "_")
    return target_dir / f"{source_file.stem}_{stamp}{source_file.suffix}"
