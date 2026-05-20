# -*- coding: utf-8 -*-
from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Any, Callable
from urllib import error, request

from core.config import app_dir, brainz_config_candidates, read_json_file, write_json_file


HISTORY_FILE_NAME = "qpsc_slack_notify_history.json"
DEFAULT_QUIET_HOURS = "22:00-04:00"
DEFAULT_MAX_PER_DAY = 3
DUPLICATE_SUPPRESS_HOURS = 12

UI_TEXT = {
    "message_return_again": "最近また戻っています。",
    "message_side_memory": "側に残っている記憶です。",
    "message_heat_hint": "熱の気配があります。",
    "message_quiet_float": "静かに浮いています。",
    "message_quiet_memory": "しばらく静かでした。",
    "message_oikawa_return": "OIKAWAから原本へ戻れます。",
    "label_title": "記憶",
    "label_path": "原本",
}


@dataclass(frozen=True)
class SlackNotifyConfig:
    enabled: bool = False
    webhook_url: str = ""
    max_per_day: int = DEFAULT_MAX_PER_DAY
    quiet_hours: str = DEFAULT_QUIET_HOURS


@dataclass(frozen=True)
class SlackNotifyCandidate:
    type: str
    title: str
    message: str
    related_path: str
    reason: str = ""
    score: int = 0
    opened_count: int = 0
    has_heat_hint: bool = False


@dataclass(frozen=True)
class SlackNotifyResult:
    status: str
    sent: bool = False
    title: str = ""
    related_path: str = ""
    message: str = ""


SlackSender = Callable[[str, str], None]


def read_slack_notify_config() -> SlackNotifyConfig:
    for path in brainz_config_candidates():
        data = read_json_file(path)
        if data:
            return slack_notify_config_from_dict(data)
    return SlackNotifyConfig()


def slack_notify_config_from_dict(data: dict[str, Any]) -> SlackNotifyConfig:
    return SlackNotifyConfig(
        enabled=bool(data.get("slack_notify_enabled", False)),
        webhook_url=str(data.get("slack_webhook_url", "") or "").strip(),
        max_per_day=_parse_int(data.get("slack_notify_max_per_day", DEFAULT_MAX_PER_DAY), DEFAULT_MAX_PER_DAY, 1, 5),
        quiet_hours=str(data.get("slack_notify_quiet_hours", DEFAULT_QUIET_HOURS) or DEFAULT_QUIET_HOURS),
    )


def slack_notify_history_path() -> Path:
    for config_path in brainz_config_candidates():
        if config_path.parent.exists():
            return config_path.parent / HISTORY_FILE_NAME
    return app_dir() / "data" / "config" / HISTORY_FILE_NAME


def maybe_send_slack_notification(
    candidates: list[SlackNotifyCandidate],
    *,
    config: SlackNotifyConfig | None = None,
    history_path: Path | None = None,
    sender: SlackSender | None = None,
    now: datetime | None = None,
) -> SlackNotifyResult:
    current = now or datetime.now().astimezone()
    notify_config = config or read_slack_notify_config()
    if not notify_config.enabled or not notify_config.webhook_url:
        return SlackNotifyResult(status="disabled")

    path = history_path or slack_notify_history_path()
    history = read_slack_notify_history(path)
    if _sent_count_today(history, current.date()) >= notify_config.max_per_day:
        return SlackNotifyResult(status="daily_limit")

    eligible = [
        candidate
        for candidate in candidates
        if _candidate_is_eligible(candidate, history, current, notify_config.quiet_hours)
    ]
    if not eligible:
        return SlackNotifyResult(status="no_candidate")

    selected = sorted(eligible, key=lambda item: (item.score, item.opened_count, item.title), reverse=True)[0]
    text = build_slack_notify_text(selected)
    try:
        (sender or send_slack_webhook)(notify_config.webhook_url, text)
    except (OSError, ValueError, error.URLError) as exc:
        return SlackNotifyResult(
            status="failed",
            sent=False,
            title=selected.title,
            related_path=selected.related_path,
            message=str(exc),
        )

    append_slack_notify_history(path, selected, current)
    return SlackNotifyResult(
        status="sent",
        sent=True,
        title=selected.title,
        related_path=selected.related_path,
        message=text,
    )


def build_slack_notify_text(candidate: SlackNotifyCandidate) -> str:
    lead = _lead_message(candidate)
    title = candidate.title.strip() or Path(candidate.related_path).stem
    path_text = _short_path(candidate.related_path)
    lines = [
        lead,
        "",
        title,
        UI_TEXT["message_oikawa_return"],
    ]
    if path_text:
        lines.append(f"{UI_TEXT['label_path']}: {path_text}")
    return "\n".join(lines)


def send_slack_webhook(webhook_url: str, text: str) -> None:
    payload = json.dumps({"text": text}, ensure_ascii=False).encode("utf-8")
    req = request.Request(
        webhook_url,
        data=payload,
        headers={"Content-Type": "application/json; charset=utf-8"},
        method="POST",
    )
    with request.urlopen(req, timeout=8) as response:
        if response.status >= 400:
            raise ValueError(f"slack webhook status {response.status}")


def read_slack_notify_history(path: Path | None = None) -> list[dict[str, Any]]:
    history_path = path or slack_notify_history_path()
    try:
        payload = json.loads(history_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    if not isinstance(payload, list):
        return []
    return [item for item in payload if isinstance(item, dict)]


def append_slack_notify_history(path: Path, candidate: SlackNotifyCandidate, sent_at: datetime) -> None:
    history = read_slack_notify_history(path)
    history.insert(
        0,
        {
            "related_path": candidate.related_path,
            "sent_at": sent_at.isoformat(timespec="seconds"),
            "type": candidate.type,
            "title": candidate.title,
        },
    )
    write_json_file(path, history[:200])


def _candidate_is_eligible(
    candidate: SlackNotifyCandidate,
    history: list[dict[str, Any]],
    now: datetime,
    quiet_hours: str,
) -> bool:
    if not candidate.related_path.strip():
        return False
    if candidate.type not in {"side_memory", "revisit", "quiet_memory", "heat_candidate", "heat_hint"}:
        return False
    if _recently_sent(candidate, history, now):
        return False

    quiet_now = is_quiet_hour(now, quiet_hours)
    has_revisit = candidate.opened_count > 1 or candidate.reason in {"recent", "night", "long_gap", "continuing"}
    has_heat = candidate.has_heat_hint or candidate.reason == "heat" or candidate.type in {"heat_candidate", "heat_hint"}
    has_time_flow = candidate.type == "quiet_memory" and candidate.reason in {"long_gap", "night", "recent", "continuing"}
    if quiet_now:
        return has_revisit or has_heat or has_time_flow or candidate.score >= 4
    return has_heat or has_time_flow or candidate.opened_count > 1 or candidate.score >= 6


def is_quiet_hour(now: datetime, quiet_hours: str) -> bool:
    start_time, end_time = _parse_quiet_hours(quiet_hours)
    current = now.time()
    if start_time <= end_time:
        return start_time <= current <= end_time
    return current >= start_time or current <= end_time


def _parse_quiet_hours(value: str) -> tuple[time, time]:
    try:
        start_text, end_text = str(value).split("-", 1)
        return _parse_clock(start_text), _parse_clock(end_text)
    except (TypeError, ValueError):
        return time(22, 0), time(4, 0)


def _parse_clock(value: str) -> time:
    hour_text, minute_text = value.strip().split(":", 1)
    return time(max(0, min(23, int(hour_text))), max(0, min(59, int(minute_text))))


def _recently_sent(candidate: SlackNotifyCandidate, history: list[dict[str, Any]], now: datetime) -> bool:
    threshold = now - timedelta(hours=DUPLICATE_SUPPRESS_HOURS)
    related_path = candidate.related_path.strip()
    for item in history:
        if str(item.get("related_path", "") or "").strip() != related_path:
            continue
        sent_at = _parse_datetime(str(item.get("sent_at", "") or ""))
        if sent_at and sent_at >= threshold:
            return True
    return False


def _sent_count_today(history: list[dict[str, Any]], today: date) -> int:
    count = 0
    for item in history:
        sent_at = _parse_datetime(str(item.get("sent_at", "") or ""))
        if sent_at and sent_at.date() == today:
            count += 1
    return count


def _lead_message(candidate: SlackNotifyCandidate) -> str:
    if candidate.reason == "long_gap":
        return f"{UI_TEXT['message_quiet_memory']}\n\n{UI_TEXT['message_return_again']}"
    if candidate.reason == "continuing":
        return f"{candidate.message.strip()}\n\n{UI_TEXT['message_quiet_float']}"
    if candidate.has_heat_hint or candidate.reason == "heat":
        return f"{UI_TEXT['message_heat_hint']}\n\n{UI_TEXT['message_quiet_float']}"
    if candidate.reason == "night":
        return f"{candidate.message.strip()}\n\n{UI_TEXT['message_side_memory']}"
    if candidate.opened_count > 1 or candidate.reason == "recent":
        return f"{UI_TEXT['message_return_again']}\n\n{UI_TEXT['message_side_memory']}"
    return f"{candidate.message.strip() or UI_TEXT['message_side_memory']}\n\n{UI_TEXT['message_quiet_float']}"


def _short_path(value: str) -> str:
    text = value.strip()
    if not text:
        return ""
    path = Path(text)
    if path.name:
        return path.name
    return text[-80:]


def _parse_datetime(value: str) -> datetime | None:
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None
    if parsed.tzinfo:
        return parsed.astimezone()
    return parsed.astimezone()


def _parse_int(value: Any, default: int, minimum: int, maximum: int) -> int:
    try:
        parsed = int(str(value).strip())
    except (TypeError, ValueError):
        return default
    return max(minimum, min(maximum, parsed))
