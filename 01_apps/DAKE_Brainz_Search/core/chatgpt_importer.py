from __future__ import annotations

import hashlib
import json
import re
import shutil
import tempfile
import zipfile
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from core.app_config import logs_dir, now_iso
from core.db import BrainzDatabase, DocumentRecord
from core.ollama_embeddings import EmbeddingSession, generate_embeddings_for_document
from core.qpsc_notifications import UI_TEXT as QPSC_NOTIFICATION_TEXT
from core.qpsc_notifications import append_saved_count_notification
from core.text_splitter import split_text


SOURCE_TYPE_CHATGPT = "chatgpt_export"
SPLIT_CONVERSATIONS_PATTERN = re.compile(r"^conversations-\d+\.json$", re.IGNORECASE)


class ConversationsJsonNotFound(FileNotFoundError):
    pass


@dataclass(frozen=True)
class ChatGPTMessageRecord:
    conversation_id: str
    title: str
    create_time: str
    update_time: str
    role: str
    message_text: str
    message_index: int


@dataclass(frozen=True)
class ChatGPTImportResult:
    source_path: str
    conversations_json_path: str
    conversations_imported: int
    messages_seen: int
    messages_indexed: int
    skipped_duplicates: int
    errors: int
    log_path: str


def import_chatgpt_export(source_path: Path, database: BrainzDatabase) -> ChatGPTImportResult:
    source_path = source_path.expanduser().resolve()
    if not source_path.exists():
        raise FileNotFoundError(str(source_path))

    if source_path.is_file() and source_path.suffix.lower() == ".zip":
        with tempfile.TemporaryDirectory(prefix="brainz_chatgpt_export_") as temp_dir:
            root = Path(temp_dir)
            safe_extract_zip(source_path, root)
            return _import_from_root(source_path, root, database)

    return _import_from_root(source_path, source_path, database)


def _import_from_root(original_path: Path, root: Path, database: BrainzDatabase) -> ChatGPTImportResult:
    conversation_paths = find_conversation_json_files(root)
    if not conversation_paths:
        raise ConversationsJsonNotFound("chatgpt export conversations json not found")

    records = parse_conversations_json_files(conversation_paths)
    conversations = {record.conversation_id for record in records}
    messages_indexed = 0
    skipped_duplicates = 0
    errors = 0
    embedding_session = EmbeddingSession()

    for record in records:
        try:
            document = build_document_record(record)
            chunks = split_text(document.content)
            document_id, changed = database.upsert_document(document, chunks)
            if changed:
                messages_indexed += 1
                try:
                    generate_embeddings_for_document(database, document_id, session=embedding_session)
                except Exception:
                    pass
            else:
                skipped_duplicates += 1
        except Exception:
            errors += 1

    result_without_log = ChatGPTImportResult(
        source_path=str(original_path),
        conversations_json_path="; ".join(str(path) for path in conversation_paths),
        conversations_imported=len(conversations),
        messages_seen=len(records),
        messages_indexed=messages_indexed,
        skipped_duplicates=skipped_duplicates,
        errors=errors,
        log_path="",
    )
    log_path = write_import_log(result_without_log)
    append_saved_count_notification(
        source=SOURCE_TYPE_CHATGPT,
        title=QPSC_NOTIFICATION_TEXT["title_chatgpt_export"],
        count=result_without_log.messages_indexed,
        related_path=str(log_path),
    )
    return ChatGPTImportResult(
        source_path=result_without_log.source_path,
        conversations_json_path=result_without_log.conversations_json_path,
        conversations_imported=result_without_log.conversations_imported,
        messages_seen=result_without_log.messages_seen,
        messages_indexed=result_without_log.messages_indexed,
        skipped_duplicates=result_without_log.skipped_duplicates,
        errors=result_without_log.errors,
        log_path=str(log_path),
    )


def safe_extract_zip(zip_path: Path, destination: Path) -> None:
    destination.mkdir(parents=True, exist_ok=True)
    root = destination.resolve()
    with zipfile.ZipFile(zip_path) as archive:
        for member in archive.infolist():
            member_name = member.filename.replace("\\", "/")
            if not member_name or member_name.endswith("/"):
                continue
            target = (root / member_name).resolve()
            if root != target and root not in target.parents:
                continue
            target.parent.mkdir(parents=True, exist_ok=True)
            with archive.open(member) as source, target.open("wb") as output:
                shutil.copyfileobj(source, output)


def find_conversations_json(root: Path) -> Path | None:
    paths = find_conversation_json_files(root)
    return paths[0] if paths else None


def find_conversation_json_files(root: Path) -> list[Path]:
    if root.is_file():
        name = root.name.lower()
        if name == "conversations.json":
            return [root]
        if SPLIT_CONVERSATIONS_PATTERN.match(name):
            return sorted(
                [path for path in root.parent.iterdir() if path.is_file() and SPLIT_CONVERSATIONS_PATTERN.match(path.name)],
                key=lambda path: path.name.lower(),
            )
        return []
    if not root.exists() or not root.is_dir():
        return []

    exact_files = sorted(
        [path for path in root.rglob("*") if path.is_file() and path.name.lower() == "conversations.json"],
        key=lambda path: str(path).lower(),
    )
    if exact_files:
        return exact_files

    return sorted(
        [path for path in root.rglob("*") if path.is_file() and SPLIT_CONVERSATIONS_PATTERN.match(path.name)],
        key=lambda path: str(path).lower(),
    )


def parse_conversations_json_files(paths: list[Path]) -> list[ChatGPTMessageRecord]:
    records: list[ChatGPTMessageRecord] = []
    for path in paths:
        records.extend(parse_conversations_json(path))
    return records


def parse_conversations_json(path: Path) -> list[ChatGPTMessageRecord]:
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError("conversations.json could not be read") from exc

    conversations = normalize_conversations(payload)
    records: list[ChatGPTMessageRecord] = []
    for conversation_index, conversation in enumerate(conversations):
        if not isinstance(conversation, dict):
            continue
        records.extend(parse_conversation(conversation, conversation_index))
    return records


def normalize_conversations(payload: Any) -> list[Any]:
    if isinstance(payload, list):
        return payload
    if isinstance(payload, dict):
        for key in ("conversations", "items", "data"):
            value = payload.get(key)
            if isinstance(value, list):
                return value
        return [payload]
    return []


def parse_conversation(conversation: dict[str, Any], conversation_index: int) -> list[ChatGPTMessageRecord]:
    conversation_id = clean_identifier(
        conversation.get("id")
        or conversation.get("conversation_id")
        or conversation.get("conversationId")
        or f"conversation_{conversation_index}"
    )
    title = clean_title(conversation.get("title") or conversation_id)
    conversation_create = format_export_time(conversation.get("create_time") or conversation.get("created_at"))
    conversation_update = format_export_time(conversation.get("update_time") or conversation.get("updated_at"))

    raw_messages = collect_messages(conversation)
    raw_messages.sort(key=lambda item: (item[0] is None, item[0] or 0.0, item[1]))

    records: list[ChatGPTMessageRecord] = []
    for message_index, (_, _, message) in enumerate(raw_messages, start=1):
        role = extract_role(message)
        text = extract_message_text(message).strip()
        if not text:
            continue
        records.append(
            ChatGPTMessageRecord(
                conversation_id=conversation_id,
                title=title,
                create_time=format_export_time(message.get("create_time") or conversation_create),
                update_time=format_export_time(message.get("update_time") or conversation_update),
                role=role,
                message_text=text,
                message_index=message_index,
            )
        )
    return records


def collect_messages(conversation: dict[str, Any]) -> list[tuple[float | None, int, dict[str, Any]]]:
    messages: list[tuple[float | None, int, dict[str, Any]]] = []
    order = 0

    mapping = conversation.get("mapping")
    if isinstance(mapping, dict):
        for node in mapping.values():
            if not isinstance(node, dict):
                continue
            message = node.get("message")
            if not isinstance(message, dict):
                continue
            messages.append((numeric_time(message.get("create_time")), order, message))
            order += 1

    for key in ("messages", "conversation"):
        value = conversation.get(key)
        if not isinstance(value, list):
            continue
        for item in value:
            if isinstance(item, dict):
                message = item.get("message") if isinstance(item.get("message"), dict) else item
                messages.append((numeric_time(message.get("create_time")), order, message))
                order += 1

    return messages


def extract_role(message: dict[str, Any]) -> str:
    author = message.get("author")
    if isinstance(author, dict) and author.get("role"):
        return str(author["role"]).strip() or "unknown"
    for key in ("role", "sender", "from"):
        if message.get(key):
            return str(message[key]).strip() or "unknown"
    return "unknown"


def extract_message_text(message: dict[str, Any]) -> str:
    content = message.get("content")
    parts = extract_text_parts(content)
    if parts:
        return "\n".join(parts)
    for key in ("text", "message"):
        value = message.get(key)
        if isinstance(value, str) and value.strip():
            return value
    return ""


def extract_text_parts(value: Any) -> list[str]:
    if isinstance(value, str):
        return [value] if value.strip() else []
    if isinstance(value, list):
        parts: list[str] = []
        for item in value:
            parts.extend(extract_text_parts(item))
        return parts
    if isinstance(value, dict):
        if isinstance(value.get("parts"), list):
            return extract_text_parts(value["parts"])
        parts = []
        for key in ("text", "content", "value", "result"):
            if key in value:
                parts.extend(extract_text_parts(value[key]))
        return parts
    return []


def build_document_record(record: ChatGPTMessageRecord) -> DocumentRecord:
    content_hash = hashlib.sha256(
        f"{record.conversation_id}\n{record.message_index}\n{record.role}\n{record.message_text}".encode("utf-8")
    ).hexdigest()
    source_path = f"chatgpt_export://{record.conversation_id}/{record.message_index:06d}/{record.role}"
    source_label = f"ChatGPT / {record.title} / {record.role}"
    content = "\n".join(
        [
            f"Conversation: {record.title}",
            f"Conversation ID: {record.conversation_id}",
            f"Role: {record.role}",
            f"Message Index: {record.message_index}",
            "",
            record.message_text,
        ]
    )
    return DocumentRecord(
        path=source_path,
        title=record.title,
        source_type=SOURCE_TYPE_CHATGPT,
        source_label=source_label,
        conversation_id=record.conversation_id,
        conversation_title=record.title,
        role=record.role,
        message_index=record.message_index,
        source_created_at=record.create_time,
        source_updated_at=record.update_time,
        created_at=record.create_time or now_iso(),
        modified_at=record.update_time or record.create_time or now_iso(),
        indexed_at=now_iso(),
        hash=content_hash,
        content=content,
    )


def clean_identifier(value: Any) -> str:
    text = str(value or "").strip()
    return text or "unknown_conversation"


def clean_title(value: Any) -> str:
    text = str(value or "").strip()
    return text or "untitled"


def numeric_time(value: Any) -> float | None:
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value))
    except (TypeError, ValueError):
        return None


def format_export_time(value: Any) -> str:
    if value in (None, ""):
        return ""
    if isinstance(value, str) and not value.replace(".", "", 1).isdigit():
        return value
    numeric = numeric_time(value)
    if numeric is None:
        return str(value)
    try:
        return datetime.fromtimestamp(numeric, tz=timezone.utc).isoformat(timespec="seconds")
    except (OSError, OverflowError, ValueError):
        return str(value)


def write_import_log(result: ChatGPTImportResult) -> Path:
    logs_dir().mkdir(parents=True, exist_ok=True)
    path = logs_dir() / f"chatgpt_import_{now_iso().replace(':', '').replace('-', '').replace('T', '_')}.txt"
    lines = [
        f"ChatGPT export detected: {result.source_path}",
        f"conversations.json found: {result.conversations_json_path}",
        f"conversations_imported={result.conversations_imported}",
        f"messages_seen={result.messages_seen}",
        f"messages_indexed={result.messages_indexed}",
        f"skipped_duplicates={result.skipped_duplicates}",
        f"errors={result.errors}",
    ]
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path
