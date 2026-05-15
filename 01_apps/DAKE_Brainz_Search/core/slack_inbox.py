from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Protocol

import requests

from core.app_config import logs_dir, now_iso
from core.db import BrainzDatabase, DocumentRecord
from core.ollama_embeddings import EmbeddingSession, generate_embeddings_for_document
from core.text_splitter import split_text


SOURCE_TYPE_SLACK_INBOX = "slack_inbox"
SLACK_HISTORY_URL = "https://slack.com/api/conversations.history"
SLACK_PERMALINK_URL = "https://slack.com/api/chat.getPermalink"
DEFAULT_HISTORY_LIMIT = 30


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


@dataclass(frozen=True)
class SlackInboxImportItem:
    ts: str
    user: str
    saved_path: str
    changed: bool
    permalink: str = ""
    title: str = ""


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


def poll_slack_inbox(
    database: BrainzDatabase,
    memory_folder: Path,
    token: str,
    channel_id: str,
    last_ts: str = "",
    poll_timeout_seconds: float = 8.0,
    session: SlackSession | None = None,
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

    slack_folder = memory_folder / "slack"
    imported = 0
    skipped = 0
    failed = 0
    latest_ts = last_ts
    saved_files: list[str] = []
    items: list[SlackInboxImportItem] = []
    embedding_session = EmbeddingSession()

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
            )
            saved_path, markdown, title = save_slack_markdown(slack_folder, channel_id, message_with_permalink)
            document_id, changed = index_slack_markdown(
                database=database,
                saved_path=saved_path,
                markdown=markdown,
                channel_id=channel_id,
                message=message_with_permalink,
                title=title,
            )
            if changed:
                generate_embeddings_for_document(database, document_id, session=embedding_session)
                imported += 1
            else:
                skipped += 1
            saved_files.append(str(saved_path))
            items.append(
                SlackInboxImportItem(
                    ts=message_item.ts,
                    user=message_item.user,
                    saved_path=str(saved_path),
                    changed=changed,
                    permalink=permalink,
                    title=title,
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
            status=status,
            imported=imported,
            skipped=skipped,
            failed=failed,
            latest_ts=latest_ts,
            items=items,
        )
    return SlackInboxResult(
        status=status,
        imported=imported,
        skipped=skipped,
        failed=failed,
        latest_ts=latest_ts,
        saved_files=saved_files,
        channel_id=channel_id,
        channel_label=channel_id,
        message="slack inbox imported" if imported else "slack connected",
        log_path=log_path,
        items=items,
    )


def fetch_slack_history(
    session: SlackSession,
    token: str,
    channel_id: str,
    last_ts: str,
    timeout_seconds: float,
) -> tuple[list[SlackInboxMessage], str, str]:
    params = {
        "channel": channel_id,
        "limit": str(DEFAULT_HISTORY_LIMIT),
    }
    if last_ts:
        params["oldest"] = last_ts
        params["inclusive"] = "false"

    payload = slack_get_json(
        session=session,
        url=SLACK_HISTORY_URL,
        token=token,
        params=params,
        timeout_seconds=timeout_seconds,
    )
    if not payload.get("ok"):
        return [], map_slack_error(str(payload.get("error") or "")), str(payload.get("error") or "slack error")

    messages: list[SlackInboxMessage] = []
    for raw_message in payload.get("messages") or []:
        if not isinstance(raw_message, dict):
            continue
        ts = str(raw_message.get("ts") or "").strip()
        text = str(raw_message.get("text") or "")
        files = normalize_files(raw_message.get("files"))
        if not ts or (not text.strip() and not files):
            continue
        user = str(raw_message.get("user") or raw_message.get("username") or raw_message.get("bot_id") or "unknown")
        messages.append(SlackInboxMessage(ts=ts, user=user, text=text, files=files))
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


def save_slack_markdown(slack_folder: Path, channel_id: str, message: SlackInboxMessage) -> tuple[Path, str, str]:
    slack_folder.mkdir(parents=True, exist_ok=True)
    timestamp = slack_timestamp(message.ts)
    filename = f"{timestamp.strftime('%Y-%m-%d_%H%M%S')}_{safe_ts(message.ts)}_slack.md"
    title = build_title(message.text, timestamp)
    markdown = build_slack_markdown(
        title=title,
        channel_id=channel_id,
        message=message,
        timestamp=timestamp,
    )
    path = slack_folder / filename
    path.write_text(markdown, encoding="utf-8")
    return path.resolve(), markdown, title


def build_slack_markdown(
    title: str,
    channel_id: str,
    message: SlackInboxMessage,
    timestamp: datetime,
) -> str:
    urls = extract_urls(message.text)
    lines = [
        "# Slack Inbox",
        "",
        "source: slack_inbox",
        f"channel: {channel_id}",
        f"user: {message.user}",
        f"timestamp: {timestamp.strftime('%Y-%m-%d %H:%M:%S')}",
        f"slack_ts: {message.ts}",
        f"title: {title}",
        "",
        "---",
        "",
        message.text.rstrip(),
        "",
        "---",
    ]
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
    return "\n".join(lines).rstrip() + "\n"


def index_slack_markdown(
    database: BrainzDatabase,
    saved_path: Path,
    markdown: str,
    channel_id: str,
    message: SlackInboxMessage,
    title: str,
) -> tuple[int, bool]:
    timestamp = slack_timestamp(message.ts).isoformat(timespec="seconds")
    digest = hashlib.sha256(markdown.encode("utf-8", errors="replace")).hexdigest()
    indexed_at = now_iso()
    record = DocumentRecord(
        path=str(saved_path.resolve()),
        title=title,
        source_type=SOURCE_TYPE_SLACK_INBOX,
        created_at=timestamp,
        modified_at=timestamp,
        indexed_at=indexed_at,
        hash=digest,
        content=markdown,
        source_label=f"Slack / {channel_id} / {message.user}",
        conversation_id=channel_id,
        conversation_title=channel_id,
        role=message.user,
        message_index=0,
        source_created_at=timestamp,
        source_updated_at=timestamp,
    )
    return database.upsert_document(record, split_text(markdown))


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
    status: str,
    imported: int,
    skipped: int,
    failed: int,
    latest_ts: str,
    items: list[SlackInboxImportItem],
) -> str:
    logs_dir().mkdir(parents=True, exist_ok=True)
    path = logs_dir() / f"slack_inbox_{now_iso().replace(':', '').replace('-', '').replace('T', '_')}.log"
    lines = [
        "SLACK INBOX:",
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
