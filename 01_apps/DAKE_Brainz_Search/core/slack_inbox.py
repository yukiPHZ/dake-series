from __future__ import annotations

import hashlib
import json
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Protocol
from urllib.parse import urlsplit, urlunsplit

import requests

from core.app_config import logs_dir, now_iso
from core.db import BrainzDatabase, DocumentRecord
from core.embers import build_ember_metadata
from core.ollama_embeddings import EmbeddingSession, generate_embeddings_for_document
from core.qpsc_notifications import UI_TEXT as QPSC_NOTIFICATION_TEXT
from core.qpsc_notifications import append_import_notification
from core.qpsc_notifications import append_saved_count_notification
from core.remote_queue import (
    RemoteQueueTask,
    destination_for,
    execute_task,
    move_task_file,
    parse_remote_queue_file,
)
from core.text_splitter import split_text


SOURCE_TYPE_SLACK = "slack"
SOURCE_TYPE_SLACK_INBOX = SOURCE_TYPE_SLACK
SOURCE_TYPE_SLACK_TASK = "slack_task"
SOURCE_TYPE_ARU = "aru"
SOURCE_TYPE_BORINEF_NOTE = "borinef_note"
NOTE_URL_PREFIX = "https://note.com/"
SLACK_HISTORY_URL = "https://slack.com/api/conversations.history"
SLACK_PERMALINK_URL = "https://slack.com/api/chat.getPermalink"
DEFAULT_HISTORY_LIMIT = 200
MAX_HISTORY_PAGES = 10
SLACK_TASK_TYPES = {"search", "note", "handoff_chatgpt", "handoff_codex", "import"}
SLACK_NOTIFICATION_KIND_PRIORITY = {
    "handoff_codex": 0,
    "handoff_chatgpt": 1,
    "note": 2,
    "search": 3,
    "import": 4,
    "slack_memory": 5,
}


class SlackSession(Protocol):
    def get(
        self,
        url: str,
        headers: dict[str, str] | None = None,
        params: dict[str, str] | None = None,
        timeout: float | None = None,
    ) -> Any:
        ...


@dataclass(frozen=True)
class SlackInboxMessage:
    ts: str
    user: str
    text: str
    permalink: str = ""
    files: list[dict[str, str]] = field(default_factory=list)
    attachments: list[dict[str, str]] = field(default_factory=list)


@dataclass(frozen=True)
class SlackInboxImportItem:
    ts: str
    user: str
    saved_path: str
    changed: bool
    permalink: str = ""
    title: str = ""
    notification_kind: str = "slack_memory"
    task_result: SlackTaskResult | None = None


@dataclass(frozen=True)
class BorinefNoteImport:
    saved_path: Path
    markdown: str
    title: str
    url: str
    changed: bool


@dataclass(frozen=True)
class SlackInboxResult:
    status: str
    imported: int
    skipped: int
    failed: int
    latest_ts: str
    saved_files: list[str]
    channel_id: str
    channel_label: str
    message: str = ""
    log_path: str = ""
    items: list[SlackInboxImportItem] = field(default_factory=list)
    task_results: list[SlackTaskResult] = field(default_factory=list)


@dataclass(frozen=True)
class SlackTask:
    task_type: str
    query: str
    note: str
    import_path: str
    raw_text: str
    title: str
    task_hash: str


@dataclass(frozen=True)
class SlackTaskResult:
    task_type: str
    query: str
    note: str
    import_path: str
    status: str
    source_file: str
    destination_file: str
    task_hash: str
    skipped_duplicate: bool = False
    document_id: int = 0
    changed: bool = False
    slack_task_document_id: int = 0
    slack_task_changed: bool = False
    error: str = ""


def classify_slack_notification_text(text: str, task_result: SlackTaskResult | None = None) -> str:
    task_type = (task_result.task_type if task_result else "").strip().lower()
    if task_type == "import":
        return "import"
    if task_type in SLACK_NOTIFICATION_KIND_PRIORITY:
        return task_type

    normalized = text.strip().lower()
    if "handoff_codex" in normalized:
        return "handoff_codex"
    if "handoff_chatgpt" in normalized:
        return "handoff_chatgpt"
    if re.search(r"(?m)(^|\s)note\s*:", normalized):
        return "note"
    if re.search(r"(?m)(^|\s)search\s*:", normalized):
        return "search"
    if re.search(r"(?m)(^|\s)import\s*:", normalized):
        return "import"
    return "slack_memory"


def slack_notification_title_for_kind(kind: str) -> str:
    if kind == "import":
        return QPSC_NOTIFICATION_TEXT["title_slack_import_task"]
    return QPSC_NOTIFICATION_TEXT.get(f"title_slack_{kind}", QPSC_NOTIFICATION_TEXT["title_slack_memory"])


def select_slack_notification_kind(items: list[SlackInboxImportItem]) -> str:
    changed_kinds = [item.notification_kind for item in items if item.changed]
    if not changed_kinds:
        return "slack_memory"
    return min(changed_kinds, key=lambda kind: SLACK_NOTIFICATION_KIND_PRIORITY.get(kind, 99))


def append_slack_import_notification(items: list[SlackInboxImportItem], related_path: str, source: str) -> None:
    changed_items = [item for item in items if item.changed]
    if not changed_items:
        return
    kind = select_slack_notification_kind(changed_items)
    try:
        append_import_notification(
            source=source,
            title=slack_notification_title_for_kind(kind),
            message=QPSC_NOTIFICATION_TEXT["message_slack_saved"],
            related_path=related_path,
        )
    except OSError:
        return


def poll_slack_inbox(
    database: BrainzDatabase,
    memory_folder: Path,
    token: str,
    channel_id: str,
    last_ts: str = "",
    poll_timeout_seconds: float = 8.0,
    session: SlackSession | None = None,
    source_type: str = SOURCE_TYPE_SLACK,
    folder_name: str = "slack",
    inbox_label: str = "Slack Inbox",
    process_tasks: bool = True,
    save_folder: str = "",
) -> SlackInboxResult:
    token = token.strip()
    channel_id = channel_id.strip()
    if not token or not channel_id:
        return SlackInboxResult(
            status="config_missing",
            imported=0,
            skipped=0,
            failed=0,
            latest_ts=last_ts,
            saved_files=[],
            channel_id=channel_id,
            channel_label=channel_id,
            message="slack config missing",
        )

    session = session or requests.Session()
    try:
        messages, status, message = fetch_slack_history(
            session=session,
            token=token,
            channel_id=channel_id,
            last_ts=last_ts,
            timeout_seconds=poll_timeout_seconds,
        )
    except requests.Timeout:
        return _result("timeout", last_ts, channel_id, "slack timeout")
    except requests.RequestException as exc:
        return _result("error", last_ts, channel_id, str(exc))
    except ValueError as exc:
        return _result("error", last_ts, channel_id, str(exc))

    if status != "ok":
        return _result(status, last_ts, channel_id, message)

    slack_folder = memory_folder / safe_relative_folder(save_folder, folder_name)
    imported = 0
    skipped = 0
    failed = 0
    latest_ts = last_ts
    saved_files: list[str] = []
    items: list[SlackInboxImportItem] = []
    task_results: list[SlackTaskResult] = []
    embedding_session = EmbeddingSession()
    note_changed_paths: list[str] = []

    for message_item in sorted(messages, key=lambda item: slack_ts_float(item.ts)):
        if slack_ts_float(message_item.ts) <= slack_ts_float(last_ts):
            skipped += 1
            continue
        try:
            permalink = message_item.permalink or fetch_slack_permalink(
                session=session,
                token=token,
                channel_id=channel_id,
                message_ts=message_item.ts,
                timeout_seconds=poll_timeout_seconds,
            )
            message_with_permalink = SlackInboxMessage(
                ts=message_item.ts,
                user=message_item.user,
                text=message_item.text,
                permalink=permalink,
                files=message_item.files,
                attachments=message_item.attachments,
            )
            saved_path, markdown, title = save_slack_markdown(
                slack_folder,
                channel_id,
                message_with_permalink,
                source_type=source_type,
                inbox_label=inbox_label,
            )
            document_id, changed = index_slack_markdown(
                database=database,
                saved_path=saved_path,
                markdown=markdown,
                channel_id=channel_id,
                message=message_with_permalink,
                title=title,
                source_type=source_type,
                inbox_label=inbox_label,
            )
            if changed:
                generate_embeddings_for_document(database, document_id, session=embedding_session)
                imported += 1
            else:
                skipped += 1
            note_import = save_borinef_note_from_slack(
                memory_folder=memory_folder,
                message=message_with_permalink,
                timestamp=slack_timestamp(message_item.ts),
            )
            if note_import is not None:
                if note_import.changed:
                    note_document_id, note_index_changed = index_published_note_markdown(
                        database=database,
                        saved_path=note_import.saved_path,
                        markdown=note_import.markdown,
                        message=message_with_permalink,
                        timestamp=slack_timestamp(message_item.ts),
                        title=note_import.title,
                    )
                    if note_index_changed:
                        generate_embeddings_for_document(database, note_document_id, session=embedding_session)
                    note_changed_paths.append(str(note_import.saved_path))
                    saved_files.append(str(note_import.saved_path))
                    imported += 1
                else:
                    skipped += 1
            task_result = None
            if process_tasks:
                task_result = process_slack_task(
                    database=database,
                    memory_folder=memory_folder,
                    channel_id=channel_id,
                    message=message_with_permalink,
                    embedding_session=embedding_session,
                )
            if task_result is not None:
                task_results.append(task_result)
            notification_kind = classify_slack_notification_text(message_with_permalink.text, task_result)
            saved_files.append(str(saved_path))
            items.append(
                SlackInboxImportItem(
                    ts=message_item.ts,
                    user=message_item.user,
                    saved_path=str(saved_path),
                    changed=changed,
                    permalink=permalink,
                    title=title,
                    notification_kind=notification_kind,
                    task_result=task_result,
                )
            )
            if slack_ts_float(message_item.ts) > slack_ts_float(latest_ts):
                latest_ts = message_item.ts
        except Exception:
            failed += 1

    status = "imported" if imported else "connected"
    log_path = ""
    if items or failed:
        log_path = write_slack_inbox_log(
            channel_id=channel_id,
            inbox_label=inbox_label,
            source_type=source_type,
            status=status,
            imported=imported,
            skipped=skipped,
            failed=failed,
            latest_ts=latest_ts,
            items=items,
            task_results=task_results,
        )
    changed_paths = [item.saved_path for item in items if item.changed]
    is_paste_source = source_type == SOURCE_TYPE_ARU
    if is_paste_source:
        append_saved_count_notification(
            source="paste",
            title=QPSC_NOTIFICATION_TEXT["title_paste_import"],
            count=imported,
            related_path=changed_paths[0] if changed_paths else "",
        )
    elif note_changed_paths:
        append_borinef_note_notification(note_changed_paths[0])
    else:
        append_slack_import_notification(
            items=items,
            related_path=changed_paths[0] if changed_paths else "",
            source=source_type,
        )
    return SlackInboxResult(
        status=status,
        imported=imported,
        skipped=skipped,
        failed=failed,
        latest_ts=latest_ts,
        saved_files=saved_files,
        channel_id=channel_id,
        channel_label=inbox_label,
        message=f"{source_type} inbox imported" if imported else f"{source_type} connected",
        log_path=log_path,
        items=items,
        task_results=task_results,
    )


def fetch_slack_history(
    session: SlackSession,
    token: str,
    channel_id: str,
    last_ts: str,
    timeout_seconds: float,
) -> tuple[list[SlackInboxMessage], str, str]:
    base_params = {
        "channel": channel_id,
        "limit": str(DEFAULT_HISTORY_LIMIT),
    }
    if last_ts:
        base_params["oldest"] = last_ts
        base_params["inclusive"] = "false"

    messages: list[SlackInboxMessage] = []
    cursor = ""
    for _page in range(MAX_HISTORY_PAGES):
        params = dict(base_params)
        if cursor:
            params["cursor"] = cursor
        payload = slack_get_json(
            session=session,
            url=SLACK_HISTORY_URL,
            token=token,
            params=params,
            timeout_seconds=timeout_seconds,
        )
        if not payload.get("ok"):
            return [], map_slack_error(str(payload.get("error") or "")), str(payload.get("error") or "slack error")

        for raw_message in payload.get("messages") or []:
            if not isinstance(raw_message, dict):
                continue
            ts = str(raw_message.get("ts") or "").strip()
            text = str(raw_message.get("text") or "")
            files = normalize_files(raw_message.get("files"))
            attachments = normalize_attachments(raw_message.get("attachments"))
            if not ts or (not text.strip() and not files and not attachments):
                continue
            user = str(raw_message.get("user") or raw_message.get("username") or raw_message.get("bot_id") or "unknown")
            messages.append(SlackInboxMessage(ts=ts, user=user, text=text, files=files, attachments=attachments))

        metadata = payload.get("response_metadata") or {}
        cursor = str(metadata.get("next_cursor") or "").strip() if isinstance(metadata, dict) else ""
        if not cursor:
            break
    return messages, "ok", ""


def fetch_slack_permalink(
    session: SlackSession,
    token: str,
    channel_id: str,
    message_ts: str,
    timeout_seconds: float,
) -> str:
    try:
        payload = slack_get_json(
            session=session,
            url=SLACK_PERMALINK_URL,
            token=token,
            params={"channel": channel_id, "message_ts": message_ts},
            timeout_seconds=timeout_seconds,
        )
    except Exception:
        return ""
    if not payload.get("ok"):
        return ""
    return str(payload.get("permalink") or "")


def slack_get_json(
    session: SlackSession,
    url: str,
    token: str,
    params: dict[str, str],
    timeout_seconds: float,
) -> dict[str, Any]:
    response = session.get(
        url,
        headers={"Authorization": f"Bearer {token}"},
        params=params,
        timeout=timeout_seconds,
    )
    status_code = int(getattr(response, "status_code", 200) or 200)
    if status_code >= 500:
        raise requests.RequestException(f"slack http {status_code}")
    try:
        payload = response.json()
    except Exception as exc:
        raise ValueError(f"invalid slack response: {exc}") from exc
    if isinstance(payload, dict):
        return payload
    raise ValueError("invalid slack response")


def append_borinef_note_notification(related_path: str) -> None:
    try:
        append_import_notification(
            source=SOURCE_TYPE_BORINEF_NOTE,
            title=QPSC_NOTIFICATION_TEXT["title_borinef_note_published"],
            message=QPSC_NOTIFICATION_TEXT["message_borinef_note_returned"],
            related_path=related_path,
        )
    except OSError:
        return


def save_borinef_note_from_slack(
    memory_folder: Path,
    message: SlackInboxMessage,
    timestamp: datetime,
) -> BorinefNoteImport | None:
    note_url = extract_note_url(message.text)
    if not note_url:
        return None
    title = extract_borinef_note_title(message.text, note_url)
    existing = find_existing_borinef_note_path(memory_folder, note_url)
    if existing:
        return BorinefNoteImport(
            saved_path=existing,
            markdown="",
            title=title,
            url=note_url,
            changed=False,
        )

    published_date = timestamp.strftime("%Y-%m-%d")
    note_folder = memory_folder / "40_borinef" / "note" / "published" / timestamp.strftime("%Y")
    note_folder.mkdir(parents=True, exist_ok=True)
    filename = f"{published_date}_{safe_note_filename_part(title)}.md"
    path = next_available_path(note_folder / filename)
    markdown = build_borinef_note_markdown(title=title, url=note_url, published_at=published_date)
    path.write_text(markdown, encoding="utf-8")
    return BorinefNoteImport(
        saved_path=path.resolve(),
        markdown=markdown,
        title=title,
        url=note_url,
        changed=True,
    )


def extract_note_url(text: str) -> str:
    for url in extract_urls(text):
        normalized = normalize_note_url(url)
        if normalized.startswith(NOTE_URL_PREFIX):
            return normalized
    return ""


def normalize_note_url(url: str) -> str:
    cleaned = str(url or "").strip().rstrip(".,、。)")
    parsed = urlsplit(cleaned)
    if parsed.scheme != "https" or parsed.netloc != "note.com":
        return cleaned
    path = parsed.path.rstrip("/")
    return urlunsplit((parsed.scheme, parsed.netloc, path, "", ""))


def extract_borinef_note_title(text: str, note_url: str) -> str:
    lines = [clean_slack_title_line(line) for line in (text or "").replace("\r", "\n").split("\n")]
    note_index = next((index for index, line in enumerate(lines) if note_url in line or "note.com/" in line), -1)
    if note_index > 0:
        for line in reversed(lines[:note_index]):
            candidate = remove_urls_from_title(line)
            if candidate:
                return candidate[:120]
    for line in lines:
        candidate = remove_urls_from_title(line)
        if candidate:
            return candidate[:120]
    return "BORINEF note"


def clean_slack_title_line(line: str) -> str:
    text = re.sub(r"<(https?://[^>|]+)\|([^>]+)>", r"\2", line or "")
    text = re.sub(r"<(https?://[^>]+)>", r"\1", text)
    return " ".join(text.strip().split())


def remove_urls_from_title(line: str) -> str:
    candidate = re.sub(r"https?://\S+", "", line or "").strip()
    return candidate.strip(" -:：|")


def safe_note_filename_part(title: str) -> str:
    clean = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "", title or "").strip()
    clean = re.sub(r"\s+", "_", clean).strip("._ ")
    return clean[:80] or "note"


def next_available_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    for index in range(2, 100):
        candidate = path.with_name(f"{stem}_{index}{suffix}")
        if not candidate.exists():
            return candidate
    return path.with_name(f"{stem}_{safe_ts(str(datetime.now().timestamp()))}{suffix}")


def find_existing_borinef_note_path(memory_folder: Path, note_url: str) -> Path | None:
    roots = (
        memory_folder / "40_borinef" / "note" / "published",
        memory_folder / "BORINEF" / "note" / "published",
    )
    for root in roots:
        if not root.exists():
            continue
        for path in root.rglob("*.md"):
            try:
                if note_url in path.read_text(encoding="utf-8"):
                    return path.resolve()
            except OSError:
                continue
    return None


def build_borinef_note_markdown(title: str, url: str, published_at: str) -> str:
    lines = [
        "---",
        f"title: {json.dumps(title, ensure_ascii=False)}",
        "series: BORINEF",
        "status: published",
        f"published_at: {published_at}",
        "platform: note",
        f"url: {json.dumps(url, ensure_ascii=False)}",
        "tags:",
        "  - BORINEF",
        "  - 在る",
        "---",
        "",
        f"# {title}",
        "",
        "URL:",
        url,
        "",
        "memo:",
        "Slack自動保存",
    ]
    return "\n".join(lines).rstrip() + "\n"


def save_slack_markdown(
    slack_folder: Path,
    channel_id: str,
    message: SlackInboxMessage,
    source_type: str = SOURCE_TYPE_SLACK,
    inbox_label: str = "Slack Inbox",
) -> tuple[Path, str, str]:
    slack_folder.mkdir(parents=True, exist_ok=True)
    timestamp = slack_timestamp(message.ts)
    filename = f"{timestamp.strftime('%Y-%m-%d_%H%M%S')}_{safe_ts(message.ts)}_{safe_file_part(source_type)}.md"
    title = build_title(message.text, timestamp)
    markdown = build_slack_markdown(
        title=title,
        channel_id=channel_id,
        message=message,
        timestamp=timestamp,
        source_type=source_type,
        inbox_label=inbox_label,
    )
    path = slack_folder / filename
    path.write_text(markdown, encoding="utf-8")
    return path.resolve(), markdown, title


def build_slack_markdown(
    title: str,
    channel_id: str,
    message: SlackInboxMessage,
    timestamp: datetime,
    source_type: str = SOURCE_TYPE_SLACK,
    inbox_label: str = "Slack Inbox",
) -> str:
    urls = extract_urls(message.text)
    channel_name = clean_channel_name(inbox_label)
    body = message.text.strip() or title
    lines = [
        "---",
        f"source: {source_type}",
        f"title: {json.dumps(title, ensure_ascii=False)}",
        f"channel: {channel_name}",
        f"channel_id: {channel_id}",
        f"timestamp: {timestamp.strftime('%Y-%m-%d %H:%M:%S')}",
        f"slack_ts: {json.dumps(message.ts, ensure_ascii=False)}",
        "status: captured",
        "tags:",
    ]
    for tag in slack_tags_for_label(inbox_label, source_type):
        lines.append(f"  - {tag}")
    lines.extend(
        [
            "---",
            "",
            f"# {title}",
            "",
            body,
            "",
            "---",
        ]
    )
    if urls:
        lines.extend(["", "URL:"])
        lines.extend(urls)
    if message.permalink:
        lines.extend(["", "Slack permalink:", message.permalink])
    if message.files:
        lines.extend(["", "Files:"])
        for file_item in message.files:
            name = file_item.get("name") or file_item.get("title") or "file"
            url = file_item.get("url_private") or file_item.get("permalink") or file_item.get("mimetype") or ""
            suffix = f" {url}" if url else ""
            lines.append(f"- {name}{suffix}")
    if message.attachments:
        lines.extend(["", "Attachments / unfurl:"])
        for attachment in message.attachments:
            title_text = attachment.get("title") or attachment.get("fallback") or attachment.get("service_name") or "attachment"
            link = attachment.get("title_link") or attachment.get("from_url") or ""
            body = attachment.get("text") or attachment.get("pretext") or ""
            suffix = f" {link}" if link else ""
            lines.append(f"- {title_text}{suffix}")
            if body:
                lines.append(f"  {body}")
    return "\n".join(lines).rstrip() + "\n"


def clean_channel_name(value: str) -> str:
    clean = str(value or "").strip()
    if clean.lower() in {"slack inbox", "aru inbox"}:
        return clean
    return clean.lstrip("#") or "slack"


def slack_tags_for_label(inbox_label: str, source_type: str) -> list[str]:
    normalized = clean_channel_name(inbox_label).lower()
    if normalized == "brainz-inbox":
        return ["inbox"]
    if normalized == "brainz-aru" or source_type == SOURCE_TYPE_ARU:
        return ["aru"]
    if normalized == "brainz-note":
        return ["BORINEF", "note"]
    if normalized == "brainz-codex":
        return ["codex"]
    if normalized == "brainz-reaction":
        return ["reaction"]
    return ["slack"]


def index_slack_markdown(
    database: BrainzDatabase,
    saved_path: Path,
    markdown: str,
    channel_id: str,
    message: SlackInboxMessage,
    title: str,
    source_type: str = SOURCE_TYPE_SLACK,
    inbox_label: str = "Slack Inbox",
) -> tuple[int, bool]:
    timestamp = slack_timestamp(message.ts).isoformat(timespec="seconds")
    digest = hashlib.sha256(markdown.encode("utf-8", errors="replace")).hexdigest()
    indexed_at = now_iso()
    record = DocumentRecord(
        path=str(saved_path.resolve()),
        title=title,
        source_type=source_type,
        created_at=timestamp,
        modified_at=timestamp,
        indexed_at=indexed_at,
        hash=digest,
        content=markdown,
        source_label=f"{inbox_label} / {channel_id} / {message.user}",
        conversation_id=channel_id,
        conversation_title=channel_id,
        role=message.user,
        message_index=0,
        source_created_at=timestamp,
        source_updated_at=timestamp,
    )
    document_id, changed = database.upsert_document(record, split_text(markdown))
    metadata = build_ember_metadata(markdown)
    database.upsert_ember_index(
        document_id=document_id,
        heat_tags=metadata.heat_tags,
        temperature=metadata.temperature,
        unfinished_score=metadata.unfinished_score,
        reignition_score=metadata.reignition_score,
        related_terms=metadata.related_terms,
        source_path=str(saved_path.resolve()),
        created_at=timestamp,
        updated_at=timestamp,
        excerpt=metadata.excerpt,
    )
    return document_id, changed


def index_published_note_markdown(
    database: BrainzDatabase,
    saved_path: Path,
    markdown: str,
    message: SlackInboxMessage,
    timestamp: datetime,
    title: str,
) -> tuple[int, bool]:
    timestamp_text = timestamp.isoformat(timespec="seconds")
    digest = hashlib.sha256(markdown.encode("utf-8", errors="replace")).hexdigest()
    record = DocumentRecord(
        path=str(saved_path.resolve()),
        title=title,
        source_type=SOURCE_TYPE_BORINEF_NOTE,
        created_at=timestamp_text,
        modified_at=timestamp_text,
        indexed_at=now_iso(),
        hash=digest,
        content=markdown,
        source_label="BORINEF note / published",
        conversation_id="BORINEF/note/published",
        conversation_title="BORINEF published note",
        role=message.user,
        message_index=0,
        source_created_at=timestamp_text,
        source_updated_at=timestamp_text,
    )
    document_id, changed = database.upsert_document(record, split_text(markdown))
    metadata = build_ember_metadata(markdown)
    database.upsert_ember_index(
        document_id=document_id,
        heat_tags=metadata.heat_tags,
        temperature=metadata.temperature,
        unfinished_score=metadata.unfinished_score,
        reignition_score=metadata.reignition_score,
        related_terms=metadata.related_terms,
        source_path=str(saved_path.resolve()),
        created_at=timestamp_text,
        updated_at=timestamp_text,
        excerpt=metadata.excerpt,
    )
    return document_id, changed


def process_slack_task(
    database: BrainzDatabase,
    memory_folder: Path,
    channel_id: str,
    message: SlackInboxMessage,
    embedding_session: EmbeddingSession,
) -> SlackTaskResult | None:
    task = parse_slack_task_text(message.text, message.ts)
    if task is None:
        return None

    queue_folder = memory_folder / "remote_queue"
    queue_source = queue_folder / "slack_tasks" / f"slack_{safe_ts(message.ts)}_{task.task_hash[:12]}.md"
    processed_destination = destination_for(queue_folder, queue_source, "processed")
    failed_destination = destination_for(queue_folder, queue_source, "failed")
    duplicate_destination = existing_task_destination(queue_folder, queue_source.name)

    slack_task_document_id = 0
    slack_task_changed = False
    try:
        slack_task_path, slack_task_markdown = save_slack_task_markdown(
            memory_folder=memory_folder,
            channel_id=channel_id,
            message=message,
            task=task,
        )
        slack_task_document_id, slack_task_changed = index_slack_task_markdown(
            database=database,
            saved_path=slack_task_path,
            markdown=slack_task_markdown,
            channel_id=channel_id,
            message=message,
            task=task,
            embedding_session=embedding_session,
        )
    except Exception:
        slack_task_document_id = 0
        slack_task_changed = False

    if duplicate_destination is not None:
        return SlackTaskResult(
            task_type=task.task_type,
            query=task.query,
            note=task.note,
            import_path=task.import_path,
            status="duplicate",
            source_file=str(queue_source),
            destination_file=str(duplicate_destination),
            task_hash=task.task_hash,
            skipped_duplicate=True,
            slack_task_document_id=slack_task_document_id,
            slack_task_changed=slack_task_changed,
        )

    try:
        queue_source.parent.mkdir(parents=True, exist_ok=True)
        queue_source.write_text(build_remote_queue_task_text(task, channel_id, message), encoding="utf-8")
        parsed_task = parse_remote_queue_file(queue_source)
        document_id, changed = execute_task(database, parsed_task, processed_destination, embedding_session)
        destination = move_task_file(queue_folder, queue_source, "processed", processed_destination)
        result = SlackTaskResult(
            task_type=task.task_type,
            query=task.query,
            note=task.note,
            import_path=task.import_path,
            status="processed",
            source_file=str(queue_source),
            destination_file=str(destination),
            task_hash=task.task_hash,
            document_id=document_id,
            changed=changed,
            slack_task_document_id=slack_task_document_id,
            slack_task_changed=slack_task_changed,
        )
    except Exception as exc:
        destination = queue_source
        if queue_source.exists():
            destination = move_task_file(queue_folder, queue_source, "failed", failed_destination)
        result = SlackTaskResult(
            task_type=task.task_type,
            query=task.query,
            note=task.note,
            import_path=task.import_path,
            status="failed",
            source_file=str(queue_source),
            destination_file=str(destination),
            task_hash=task.task_hash,
            slack_task_document_id=slack_task_document_id,
            slack_task_changed=slack_task_changed,
            error=str(exc),
        )

    database.log_remote_queue_task(
        task_type=result.task_type,
        query=result.query,
        note=result.note,
        source_file=result.source_file,
        status=result.status,
        raw_text=task.raw_text,
    )
    return result


def parse_slack_task_text(text: str, slack_ts: str = "") -> SlackTask | None:
    raw_text = (text or "").strip()
    if not raw_text:
        return None
    lines = raw_text.splitlines()
    first_index = next((index for index, line in enumerate(lines) if line.strip()), -1)
    if first_index < 0:
        return None

    first_line = lines[first_index].strip()
    first_match = re.match(r"^(type|search|note|handoff_chatgpt|handoff_codex|import)\s*:\s*(.*)$", first_line, re.IGNORECASE)
    if not first_match:
        return None

    key = first_match.group(1).strip().lower()
    value = first_match.group(2).strip()
    remainder = "\n".join(lines[first_index + 1 :]).strip()
    values = key_value_lines(raw_text)

    if key == "type":
        task_type = normalize_slack_task_type(value)
        if task_type not in SLACK_TASK_TYPES:
            return None
        query = values.get("query", "")
        note = values.get("note", "")
        import_path = values.get("path") or values.get("file_path") or values.get("import_path") or values.get("file") or ""
        if task_type in {"search", "handoff_chatgpt", "handoff_codex"} and not query:
            query = values.get(task_type, "") or remainder
        if task_type == "note" and not note:
            note = remainder
        if task_type == "import" and not import_path:
            import_path = remainder
    else:
        task_type = normalize_slack_task_type(key)
        payload = value or remainder
        query = payload if task_type in {"search", "handoff_chatgpt", "handoff_codex"} else ""
        note = payload if task_type == "note" else values.get("note", "")
        import_path = payload if task_type == "import" else ""

    if task_type in {"search", "handoff_chatgpt", "handoff_codex"} and not query.strip():
        return None
    if task_type == "note" and not note.strip():
        return None
    if task_type == "import" and not import_path.strip():
        return None

    title = f"Slack Task {task_type}"
    task_hash = hashlib.sha256(
        "\n".join([slack_ts, task_type, query, note, import_path, raw_text]).encode("utf-8", errors="replace")
    ).hexdigest()
    return SlackTask(
        task_type=task_type,
        query=query.strip(),
        note=note.strip(),
        import_path=import_path.strip(),
        raw_text=raw_text,
        title=title,
        task_hash=task_hash,
    )


def key_value_lines(text: str) -> dict[str, str]:
    values: dict[str, str] = {}
    for line in (text or "").splitlines():
        match = re.match(r"^\s*([A-Za-z_][A-Za-z0-9_\- ]{0,48})\s*:\s*(.*)$", line.strip())
        if not match:
            continue
        key = match.group(1).strip().lower().replace("-", "_").replace(" ", "_")
        value = match.group(2).strip()
        if value:
            values[key] = value
    return values


def normalize_slack_task_type(value: str) -> str:
    clean = value.strip().lower().replace("-", "_").replace(" ", "_")
    aliases = {
        "chatgpt_handoff": "handoff_chatgpt",
        "handoff_gpt": "handoff_chatgpt",
        "codex_handoff": "handoff_codex",
    }
    return aliases.get(clean, clean)


def build_remote_queue_task_text(task: SlackTask, channel_id: str, message: SlackInboxMessage) -> str:
    lines = [
        "# Slack Task",
        "",
        f"type: {task.task_type}",
    ]
    if task.query:
        lines.append(f"query: {' '.join(task.query.split())}")
    if task.note:
        lines.extend(["note:", task.note])
    if task.import_path:
        lines.append(f"path: {task.import_path}")
    lines.extend(
        [
            f"source: slack",
            f"channel: {channel_id}",
            f"slack_ts: {message.ts}",
            "",
            "---",
            "",
            task.raw_text,
            "",
        ]
    )
    return "\n".join(lines).rstrip() + "\n"


def save_slack_task_markdown(
    memory_folder: Path,
    channel_id: str,
    message: SlackInboxMessage,
    task: SlackTask,
) -> tuple[Path, str]:
    task_folder = memory_folder / "slack_tasks"
    task_folder.mkdir(parents=True, exist_ok=True)
    timestamp = slack_timestamp(message.ts)
    path = task_folder / f"{timestamp.strftime('%Y-%m-%d_%H%M%S')}_{safe_ts(message.ts)}_{task.task_hash[:12]}_task.md"
    markdown = build_slack_task_markdown(channel_id=channel_id, message=message, task=task, timestamp=timestamp)
    path.write_text(markdown, encoding="utf-8")
    return path.resolve(), markdown


def build_slack_task_markdown(
    channel_id: str,
    message: SlackInboxMessage,
    task: SlackTask,
    timestamp: datetime,
) -> str:
    lines = [
        "# Slack Task",
        "",
        "source: slack_task",
        f"channel: {channel_id}",
        f"user: {message.user}",
        f"timestamp: {timestamp.strftime('%Y-%m-%d %H:%M:%S')}",
        f"slack_ts: {message.ts}",
        f"task_type: {task.task_type}",
        f"task_hash: {task.task_hash}",
        "",
        "---",
        "",
        task.raw_text,
        "",
    ]
    return "\n".join(lines).rstrip() + "\n"


def index_slack_task_markdown(
    database: BrainzDatabase,
    saved_path: Path,
    markdown: str,
    channel_id: str,
    message: SlackInboxMessage,
    task: SlackTask,
    embedding_session: EmbeddingSession,
) -> tuple[int, bool]:
    timestamp = slack_timestamp(message.ts).isoformat(timespec="seconds")
    digest = hashlib.sha256(markdown.encode("utf-8", errors="replace")).hexdigest()
    record = DocumentRecord(
        path=str(saved_path.resolve()),
        title=task.title,
        source_type=SOURCE_TYPE_SLACK_TASK,
        created_at=timestamp,
        modified_at=timestamp,
        indexed_at=now_iso(),
        hash=digest,
        content=markdown,
        source_label=f"Slack Task / {task.task_type}",
        conversation_id=channel_id,
        conversation_title=channel_id,
        role=message.user,
        source_created_at=timestamp,
        source_updated_at=timestamp,
    )
    document_id, changed = database.upsert_document(record, split_text(markdown))
    metadata = build_ember_metadata(markdown)
    database.upsert_ember_index(
        document_id=document_id,
        heat_tags=metadata.heat_tags,
        temperature=metadata.temperature,
        unfinished_score=metadata.unfinished_score,
        reignition_score=metadata.reignition_score,
        related_terms=metadata.related_terms,
        source_path=str(saved_path.resolve()),
        created_at=timestamp,
        updated_at=timestamp,
        excerpt=metadata.excerpt,
    )
    if changed:
        generate_embeddings_for_document(database, document_id, session=embedding_session)
    return document_id, changed


def existing_task_destination(queue_folder: Path, file_name: str) -> Path | None:
    for status in ("processed", "failed"):
        path = queue_folder / status / file_name
        if path.exists():
            return path.resolve()
    return None


def safe_folder_name(value: str) -> str:
    clean = re.sub(r"[^0-9A-Za-z_.-]+", "_", (value or "").strip()).strip("._")
    return clean or "slack"


def safe_relative_folder(value: str, fallback: str) -> Path:
    parts: list[str] = []
    for part in str(value or "").replace("\\", "/").split("/"):
        clean = safe_folder_name(part)
        if clean and clean not in {".", ".."}:
            parts.append(clean)
    if not parts:
        return Path(safe_folder_name(fallback))
    return Path(*parts)


def safe_file_part(value: str) -> str:
    clean = re.sub(r"[^0-9A-Za-z_-]+", "_", (value or "").strip()).strip("_")
    return clean or "slack"


def normalize_files(raw_files: Any) -> list[dict[str, str]]:
    if not isinstance(raw_files, list):
        return []
    files: list[dict[str, str]] = []
    for raw_file in raw_files:
        if not isinstance(raw_file, dict):
            continue
        files.append(
            {
                "name": str(raw_file.get("name") or ""),
                "title": str(raw_file.get("title") or ""),
                "url_private": str(raw_file.get("url_private") or ""),
                "permalink": str(raw_file.get("permalink") or ""),
                "mimetype": str(raw_file.get("mimetype") or ""),
            }
        )
    return files


def normalize_attachments(raw_attachments: Any) -> list[dict[str, str]]:
    if not isinstance(raw_attachments, list):
        return []
    attachments: list[dict[str, str]] = []
    for raw_attachment in raw_attachments:
        if not isinstance(raw_attachment, dict):
            continue
        attachments.append(
            {
                "fallback": str(raw_attachment.get("fallback") or ""),
                "pretext": str(raw_attachment.get("pretext") or ""),
                "title": str(raw_attachment.get("title") or ""),
                "title_link": str(raw_attachment.get("title_link") or ""),
                "text": str(raw_attachment.get("text") or ""),
                "from_url": str(raw_attachment.get("from_url") or ""),
                "service_name": str(raw_attachment.get("service_name") or ""),
            }
        )
    return attachments


def build_title(text: str, timestamp: datetime) -> str:
    clean = " ".join((text or "").replace("\r", "\n").split())
    clean = re.sub(r"<([^|>]+)\|([^>]+)>", r"\2", clean)
    clean = re.sub(r"<([^>]+)>", r"\1", clean)
    if not clean:
        return f"Slack Inbox {timestamp.strftime('%Y-%m-%d %H:%M:%S')}"
    return clean[:72]


def extract_urls(text: str) -> list[str]:
    urls: list[str] = []
    for match in re.finditer(r"<(https?://[^>|]+)(?:\|[^>]+)?>|(https?://[^\s<>()]+)", text or ""):
        url = match.group(1) or match.group(2) or ""
        if url and url not in urls:
            urls.append(url)
    return urls


def map_slack_error(error_code: str) -> str:
    if error_code in {"invalid_auth", "not_authed", "account_inactive", "token_revoked"}:
        return "auth_failed"
    if error_code in {"channel_not_found", "not_in_channel", "is_archived"}:
        return "channel_not_found"
    if error_code in {"ratelimited", "request_timeout"}:
        return "timeout"
    return "error"


def slack_timestamp(ts: str) -> datetime:
    value = slack_ts_float(ts)
    if value <= 0:
        return datetime.now()
    return datetime.fromtimestamp(value)


def slack_ts_float(ts: str) -> float:
    try:
        return float(str(ts or "0"))
    except ValueError:
        return 0.0


def safe_ts(ts: str) -> str:
    return re.sub(r"[^0-9A-Za-z_-]+", "_", ts).strip("_") or "0"


def write_slack_inbox_log(
    channel_id: str,
    inbox_label: str,
    source_type: str,
    status: str,
    imported: int,
    skipped: int,
    failed: int,
    latest_ts: str,
    items: list[SlackInboxImportItem],
    task_results: list[SlackTaskResult],
) -> str:
    logs_dir().mkdir(parents=True, exist_ok=True)
    log_prefix = safe_file_part(source_type)
    path = logs_dir() / f"{log_prefix}_inbox_{now_iso().replace(':', '').replace('-', '').replace('T', '_')}.log"
    lines = [
        f"{inbox_label.upper()}:",
        f"source_type={source_type}",
        f"channel={channel_id}",
        f"status={status}",
        f"imported={imported}",
        f"skipped={skipped}",
        f"failed={failed}",
        f"latest_ts={latest_ts}",
    ]
    for item in items:
        lines.extend(
            [
                f"- ts: {item.ts}",
                f"  user: {item.user}",
                f"  title: {item.title}",
                f"  changed: {item.changed}",
                f"  path: {item.saved_path}",
                f"  permalink: {item.permalink}",
            ]
        )
        if item.task_result is not None:
            lines.extend(
                [
                    f"  task_type: {item.task_result.task_type}",
                    f"  task_status: {item.task_result.status}",
                    f"  task_query: {item.task_result.query}",
                    f"  task_hash: {item.task_result.task_hash}",
                ]
            )
    if task_results:
        lines.append("SLACK TASKS:")
    for task_result in task_results:
        lines.extend(
            [
                f"- task_type: {task_result.task_type}",
                f"  status: {task_result.status}",
                f"  query: {task_result.query}",
                f"  note: {task_result.note[:160]}",
                f"  import_path: {task_result.import_path}",
                f"  skipped_duplicate: {task_result.skipped_duplicate}",
                f"  source: {task_result.source_file}",
                f"  destination: {task_result.destination_file}",
                f"  error: {task_result.error}",
            ]
        )
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return str(path)


def _result(status: str, latest_ts: str, channel_id: str, message: str) -> SlackInboxResult:
    return SlackInboxResult(
        status=status,
        imported=0,
        skipped=0,
        failed=0,
        latest_ts=latest_ts,
        saved_files=[],
        channel_id=channel_id,
        channel_label=channel_id,
        message=message,
    )
