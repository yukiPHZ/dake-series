# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import json
import math
import queue
import random
import threading
import time
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk

from core.config import ConfigStore, app_dir, existing_folder, open_path, resolve_memory_folder, write_json_file
from core.heat_engine import AnalysisResult, analyze_documents
from core.markdown_writer import write_suggestion
from core.qpsc_notifications import (
    QpscNotification,
    brainz_notification_candidates,
    mark_qpsc_notification_read,
    read_qpsc_notification_events,
)
from core.qpsc_status import brainz_status_candidates, is_brainz_awake, read_brainz_status
from core.scanner import MemoryDocument, scan_memory


APP_NAME = "DakeBrainzOIKAWA"
WINDOW_TITLE = "OIKAWA"
COPYRIGHT = "QPSC — Quiet Personal Cognitive System by Yukihiko Kikuta"

UI_TEXT = {
    "app_title": "OIKAWA",
    "copyright": COPYRIGHT,
    "app_subcopy": "補助脳BRAINZ",
    "app_subtitle": "検索 / 原本 / 熾火 / 通知",
    "button_search": "記憶検索",
    "button_searching": "検索中",
    "button_heat_search": "熱検索",
    "button_heat_searching": "熱検索中",
    "button_scan": "巡回する",
    "button_scanning": "巡回中",
    "button_open_output": "保存先を開く",
    "button_choose_folder": "記憶フォルダを選ぶ",
    "button_preview_source": "原本を読む",
    "button_qpsc_state_check": "状態確認",
    "button_qpsc_state_checking": "確認中",
    "status_idle": "眠っています",
    "status_searching": "記憶を検索中",
    "status_search_complete": "検索完了",
    "status_no_search_results": "検索結果はありません",
    "status_heat_searching": "熱検索中",
    "status_heat_search_complete": "熱検索完了",
    "status_no_heat_results": "熱検索の候補はありません",
    "status_scanning": "記憶層を巡回中",
    "status_complete": "観測完了",
    "status_no_trace": "強い熱の痕跡は見つかりませんでした",
    "status_error": "巡回に失敗しました",
    "section_traces": "浮上した痕跡",
    "section_related": "関連断片",
    "section_suggestion": "OIKAWA提案",
    "section_source_preview": "原本プレビュー",
    "section_heat_candidates": "熾火",
    "section_orbit_today": "今日の整理",
    "section_orbit_flow": "今日の流れ",
    "section_orbit_next": "次の候補",
    "section_qpsc_state": "QPSC状態",
    "label_memory_folder": "記憶フォルダ",
    "memory_folder_missing": "記憶フォルダ未検出",
    "dialog_choose_memory": "記憶フォルダを選択",
    "dialog_title": "OIKAWA",
    "summary_idle": "呼ばれるまで、記憶層の外側で待機します",
    "summary_scan": "files {files} / skipped {skipped} / traces {traces}",
    "summary_search": "results {results} / skipped {skipped}",
    "summary_heat_search": "heat {results} / skipped {skipped}",
    "summary_saved": "提案Markdownを保存しました",
    "card_file": "該当ファイル",
    "card_excerpt": "抜粋",
    "card_score": "score {score}",
    "card_empty": "まだ浮上したカードはありません",
    "source_preview_empty": "通知または検索結果から原本を選んでください。",
    "source_preview_loading": "原本を読み込んでいます。",
    "source_preview_loaded": "原本を表示しました。",
    "source_preview_missing": "ファイルが見つかりません。",
    "source_preview_failed": "原本を表示できませんでした。",
    "source_preview_not_file": "原本を表示できませんでした。",
    "source_preview_path": "PATH: {path}",
    "source_preview_truncated": "\n\n---\n先頭のみ表示しています。",
    "search_result_meta": "{path} / score {score}",
    "heat_candidate_empty": "まだ熾火候補はありません。",
    "heat_candidate_title": "まだ熱が残っている記憶",
    "heat_candidate_message": "{source}から取り込まれた記憶があります。",
    "heat_candidate_no_related": "原本への道筋はまだありません。",
    "heat_reason_unread_notification": "未読通知",
    "heat_reason_recent_import": "最近の取り込み",
    "heat_reason_related_path": "原本あり",
    "heat_reason_query_match": "検索語に近い",
    "orbit_status_loading": "今日の流れを整理しています。",
    "orbit_status_unavailable": "今日の整理を表示できませんでした。",
    "orbit_metric_today": "今日取り込まれた記憶: {count}件",
    "orbit_metric_unread": "未読通知: {count}件",
    "orbit_metric_quieted": "今日静かになった通知: {count}件",
    "orbit_metric_ember": "熾火候補: {count}件",
    "orbit_metric_recent_returns": "最近戻った原本: {count}件",
    "orbit_metric_awake": "記憶庫は起きています",
    "orbit_metric_quiet": "記憶庫は静かです",
    "orbit_flow_import_source": "{source}から記憶が入りました",
    "orbit_flow_slack": "Slackから記憶が入りました",
    "orbit_flow_codex": "Codex報告がBRAINZに入りました",
    "orbit_flow_unread": "未読通知が{count}件あります",
    "orbit_flow_quieted": "静かになった通知が{count}件あります",
    "orbit_flow_ember": "熾火候補が{count}件あります",
    "orbit_flow_recent_returns": "最近戻った原本が{count}件あります",
    "orbit_flow_query": "検索語: {query}",
    "orbit_flow_awake": "記憶庫は起きています",
    "orbit_flow_quiet": "今日は静かです",
    "orbit_bullet": "- {text}",
    "orbit_next_empty": "次の候補はまだありません。",
    "orbit_next_unread": "この記憶に戻れます",
    "orbit_next_heat": "まだ熱が残っていそうです",
    "orbit_next_today": "今日取り込まれた記憶です",
    "orbit_next_related": "原本へ戻れます",
    "orbit_next_query": "検索語に近い記憶です",
    "orbit_next_no_related": "原本への道筋はまだありません。",
    "orbit_next_template": "{message}: {title}",
    "section_qpsc_notifications": "QPSC通知",
    "qpsc_brainz_awake": "BRAINZ is awake.",
    "qpsc_memory_awake": "記憶庫は起きています。",
    "qpsc_waiting": "BRAINZの起床通知を待っています。",
    "qpsc_memory_waiting": "記憶庫の状態を確認中です。",
    "qpsc_heartbeat_quiet": "BRAINZ heartbeat is quiet.",
    "qpsc_memory_heartbeat_quiet": "記憶庫の鼓動を待っています。",
    "qpsc_no_suggestion": "今日の提案はまだありません。",
    "qpsc_notice_template": "{line1}\n{line2}\n{line3}",
    "qpsc_notification_empty": "取り込み通知はありません。",
    "qpsc_notification_sedimented": "沈んだ通知: {count}件",
    "qpsc_notification_open_related": "原本",
    "qpsc_notification_read": "既読",
    "qpsc_slack_title": "Slackから記憶が入りました",
    "qpsc_slack_message": "記憶庫に保存しました。OIKAWAから戻れます。",
    "qpsc_slack_legacy_title": "Slackから取り込みました",
    "qpsc_saved_count_marker": "件の記憶を保存しました",
    "qpsc_codex_title": "Codex報告を正本として保存しました",
    "qpsc_codex_message": "原本をそのまま保存しました。OIKAWAから戻れます。",
    "qpsc_codex_legacy_report_title": "Codex報告を保存しました",
    "qpsc_codex_legacy_result_title": "Codex結果を保存しました",
    "qpsc_codex_legacy_notify_title": "Codex報告を記憶しました",
    "qpsc_related_missing": "原本が見つかりません。",
    "qpsc_check_brainz_ok": "BRAINZ awake: OK",
    "qpsc_check_brainz_unconfirmed": "BRAINZ awake: 未確認",
    "qpsc_check_notifications_ok": "通知ファイル: OK",
    "qpsc_check_notifications_missing": "通知ファイル: 未作成",
    "qpsc_check_preview_ok": "原本プレビュー: OK",
    "qpsc_check_preview_waiting": "原本プレビュー: 待機中",
    "qpsc_check_orbit_ok": "ORBIT: OK",
    "qpsc_check_orbit_missing": "ORBIT: 未生成",
    "qpsc_check_recent_returns": "最近戻った原本: {count}件",
    "qpsc_check_related_ok": "原本パス: OK",
    "qpsc_check_related_empty": "原本パス: まだ記録がありません",
    "qpsc_check_related_missing": "原本パス: ファイルが見つかりません",
    "footer_source": "記憶はBRAINZに在り、OIKAWAが静かに呼び戻します。",
    "launch_check_ok": "LAUNCH CHECK OK",
    "gui_smoke_ok": "GUI SMOKE OK",
    "scan_check_ok": "SCAN CHECK OK",
    "ghost_words": ["熾火", "巡り", "側に", "在る", "余白", "記憶", "痕跡"],
}

MAX_PREVIEW_CHARS = 240_000
ORBIT_FILE_NAME = "qpsc_orbit_today.json"
RECENT_RETURNS_FILE_NAME = "qpsc_recent_returns.json"


@dataclass(frozen=True)
class SourcePreviewResult:
    request_id: int
    path: Path
    title: str
    text: str
    truncated: bool


@dataclass(frozen=True)
class SearchHit:
    title: str
    path: Path
    relative_path: str
    excerpt: str
    score: int


@dataclass(frozen=True)
class HeatCandidate:
    id: str
    title: str
    message: str
    reason: str
    related_path: str
    score: int


@dataclass(frozen=True)
class NotificationView:
    notification: QpscNotification
    priority: int
    sedimented: bool
    heat_related: bool


@dataclass(frozen=True)
class OrbitCandidate:
    id: str
    title: str
    message: str
    reason: str
    related_path: str
    score: int


@dataclass(frozen=True)
class OrbitToday:
    generated_at: str
    date: str
    brainz_awake: bool
    notification_count_today: int
    unread_count: int
    quieted_count: int
    ember_count: int
    recent_return_count: int
    summary_lines: list[str]
    next_candidates: list[OrbitCandidate]


@dataclass(frozen=True)
class RecentReturn:
    opened_at: str
    title: str
    related_path: str


@dataclass(frozen=True)
class QpscSelfCheckResult:
    lines: list[str]


def _is_slack_notification(notification: QpscNotification) -> bool:
    return notification.source.strip().lower() == "slack"


def _is_codex_notification(notification: QpscNotification) -> bool:
    source = notification.source.strip().lower()
    title = notification.title.strip().lower()
    return source in {"codex_result", "codex_report_auto", "handoff_codex"} or "codex" in title


def _notification_display_title(notification: QpscNotification) -> str:
    title = notification.title.strip()
    if _is_slack_notification(notification):
        if not title or title == UI_TEXT["qpsc_slack_legacy_title"]:
            return UI_TEXT["qpsc_slack_title"]
        return title
    if _is_codex_notification(notification):
        legacy_titles = {
            UI_TEXT["qpsc_codex_legacy_report_title"],
            UI_TEXT["qpsc_codex_legacy_result_title"],
            UI_TEXT["qpsc_codex_legacy_notify_title"],
        }
        if not title or title in legacy_titles:
            return UI_TEXT["qpsc_codex_title"]
        return title
    return title


def _notification_display_message(notification: QpscNotification) -> str:
    message = notification.message.strip()
    if _is_slack_notification(notification):
        if not message or UI_TEXT["qpsc_saved_count_marker"] in message:
            return UI_TEXT["qpsc_slack_message"]
        return message
    if _is_codex_notification(notification):
        if not message or UI_TEXT["qpsc_saved_count_marker"] in message:
            return UI_TEXT["qpsc_codex_message"]
        return message
    return message


def _orbit_flow_line_for_source(source: str) -> str:
    source_key = source.strip().lower()
    if source_key == "slack":
        return UI_TEXT["orbit_flow_slack"]
    if source_key in {"codex_result", "codex_report_auto", "handoff_codex"}:
        return UI_TEXT["orbit_flow_codex"]
    return UI_TEXT["orbit_flow_import_source"].format(source=source)


def build_search_hits(documents: list[MemoryDocument], query: str, memory_root: Path, limit: int = 20) -> list[SearchHit]:
    terms = [term for term in query.lower().split() if term]
    if not terms:
        return []
    hits: list[SearchHit] = []
    for document in documents:
        score = _document_search_score(document, terms)
        if score <= 0:
            continue
        hits.append(
            SearchHit(
                title=document.title,
                path=document.path,
                relative_path=_relative_path(document.path, memory_root),
                excerpt=_search_excerpt(document.text, terms),
                score=score,
            )
        )
    hits.sort(key=lambda item: (item.score, item.path.stat().st_mtime if item.path.exists() else 0), reverse=True)
    return hits[:limit]


def build_heat_candidates(
    notifications: list[QpscNotification],
    query: str = "",
    limit: int = 3,
) -> list[HeatCandidate]:
    terms = [term for term in query.lower().split() if term]
    candidates: list[HeatCandidate] = []
    seen: set[str] = set()
    for index, notification in enumerate(notifications):
        key = notification.related_path.strip() or notification.id
        if key and key in seen:
            continue
        if key:
            seen.add(key)
        score = _notification_heat_score(notification, index, terms)
        if score <= 0:
            continue
        candidates.append(
            HeatCandidate(
                id=notification.id or f"auto-{index}",
                title=_notification_display_title(notification) or UI_TEXT["heat_candidate_title"],
                message=_notification_display_message(notification) or UI_TEXT["heat_candidate_message"].format(
                    source=notification.source.strip() or UI_TEXT["section_qpsc_notifications"],
                ),
                reason=_notification_heat_reason(notification, terms),
                related_path=notification.related_path.strip(),
                score=score,
            )
        )
    candidates.sort(key=lambda item: item.score, reverse=True)
    return candidates[:limit]


def build_heat_search_hits(
    documents: list[MemoryDocument],
    query: str,
    memory_root: Path,
    notifications: list[QpscNotification],
    limit: int = 20,
) -> list[SearchHit]:
    terms = [term for term in query.lower().split() if term]
    if not terms:
        return []
    notification_boosts = _notification_boosts(notifications, memory_root, terms)
    hits: list[SearchHit] = []
    seen_paths: set[Path] = set()
    for document in documents:
        base_score = _document_search_score(document, terms)
        resolved_path = document.path.resolve()
        boost = notification_boosts.get(resolved_path, 0)
        if base_score <= 0 and boost <= 0:
            continue
        seen_paths.add(resolved_path)
        hits.append(
            SearchHit(
                title=document.title,
                path=document.path,
                relative_path=_relative_path(document.path, memory_root),
                excerpt=_search_excerpt(document.text, terms),
                score=base_score + boost,
            )
        )
    for index, notification in enumerate(notifications):
        related_path = notification.related_path.strip()
        if not related_path:
            continue
        path = _resolve_related_source_path(related_path, memory_root)
        resolved_path = path.resolve()
        if resolved_path in seen_paths:
            continue
        notification_text_match = any(term in _notification_search_text(notification) for term in terms)
        preview_text = ""
        file_score = 0
        if path.exists() and path.is_file() and path.suffix.lower() in {".md", ".txt"}:
            preview_text, _truncated = read_source_preview_text(path, limit=24_000)
            file_score = _text_search_score(preview_text, terms)
        if not notification_text_match and file_score <= 0:
            continue
        score = file_score + _notification_heat_score(notification, index, terms)
        excerpt_source = preview_text or _notification_display_message(notification) or _notification_display_title(notification)
        hits.append(
            SearchHit(
                title=_notification_display_title(notification) or path.stem,
                path=path,
                relative_path=_relative_path(path, memory_root),
                excerpt=_search_excerpt(excerpt_source, terms),
                score=score,
            )
        )
        seen_paths.add(resolved_path)
    hits.sort(key=lambda item: (item.score, item.path.stat().st_mtime if item.path.exists() else 0), reverse=True)
    return hits[:limit]


def build_notification_view(
    notifications: list[QpscNotification],
    heat_candidates: list[HeatCandidate],
    limit: int = 3,
    today: date | None = None,
) -> tuple[list[NotificationView], int]:
    target_date = today or datetime.now().astimezone().date()
    heat_paths = {candidate.related_path.strip() for candidate in heat_candidates if candidate.related_path.strip()}
    view_items: list[NotificationView] = []
    sedimented_count = 0
    for notification in notifications:
        heat_related = notification.related_path.strip() in heat_paths
        sedimented = _notification_is_sedimented(notification, target_date)
        priority = _notification_view_priority(notification, heat_related, target_date)
        item = NotificationView(
            notification=notification,
            priority=priority,
            sedimented=sedimented,
            heat_related=heat_related,
        )
        if sedimented:
            sedimented_count += 1
            continue
        view_items.append(item)
    view_items.sort(key=lambda item: (item.priority, -_notification_timestamp(item.notification)))
    return view_items[:limit], sedimented_count


def build_orbit_today(
    notifications: list[QpscNotification],
    heat_candidates: list[HeatCandidate],
    brainz_awake: bool,
    recent_returns: list[RecentReturn] | None = None,
    query: str = "",
    today: date | None = None,
) -> OrbitToday:
    generated = datetime.now().astimezone()
    orbit_date = today or generated.date()
    today_notifications = [item for item in notifications if _notification_is_today(item, orbit_date)]
    unread_count = sum(1 for item in notifications if item.status == "unread")
    quieted_count = sum(1 for item in notifications if _notification_quieted_today(item, orbit_date))
    recent_return_count = _count_recent_returns(recent_returns or [], orbit_date)
    summary_lines = _build_orbit_summary_lines(
        today_notifications,
        unread_count,
        quieted_count,
        len(heat_candidates),
        recent_return_count,
        brainz_awake,
        query,
    )
    next_candidates = build_orbit_next_candidates(notifications, heat_candidates, query, orbit_date, limit=3)
    return OrbitToday(
        generated_at=generated.isoformat(timespec="seconds"),
        date=orbit_date.isoformat(),
        brainz_awake=brainz_awake,
        notification_count_today=len(today_notifications),
        unread_count=unread_count,
        quieted_count=quieted_count,
        ember_count=len(heat_candidates),
        recent_return_count=recent_return_count,
        summary_lines=summary_lines,
        next_candidates=next_candidates,
    )


def build_orbit_next_candidates(
    notifications: list[QpscNotification],
    heat_candidates: list[HeatCandidate],
    query: str = "",
    today: date | None = None,
    limit: int = 3,
) -> list[OrbitCandidate]:
    orbit_date = today or datetime.now().astimezone().date()
    terms = [term for term in query.lower().split() if term]
    candidates: list[OrbitCandidate] = []
    seen: set[str] = set()

    for notification in notifications:
        if notification.status == "unread":
            _append_orbit_candidate(
                candidates,
                seen,
                _orbit_candidate_from_notification(notification, "unread_notification", UI_TEXT["orbit_next_unread"], 5000),
            )

    for heat_candidate in heat_candidates:
        _append_orbit_candidate(
            candidates,
            seen,
            OrbitCandidate(
                id=heat_candidate.id,
                title=heat_candidate.title,
                message=UI_TEXT["orbit_next_heat"],
                reason="heat_candidate",
                related_path=heat_candidate.related_path,
                score=4000 + heat_candidate.score,
            ),
        )

    for notification in notifications:
        if _notification_is_today(notification, orbit_date):
            _append_orbit_candidate(
                candidates,
                seen,
                _orbit_candidate_from_notification(notification, "today_import", UI_TEXT["orbit_next_today"], 3000),
            )

    for notification in notifications:
        if notification.related_path.strip():
            _append_orbit_candidate(
                candidates,
                seen,
                _orbit_candidate_from_notification(notification, "related_path", UI_TEXT["orbit_next_related"], 2000),
            )

    if terms:
        for notification in notifications:
            if any(term in _notification_search_text(notification) for term in terms):
                _append_orbit_candidate(
                    candidates,
                    seen,
                    _orbit_candidate_from_notification(notification, "query_match", UI_TEXT["orbit_next_query"], 1000),
                )

    candidates.sort(key=lambda item: item.score, reverse=True)
    return candidates[:limit]


def save_orbit_today(orbit: OrbitToday) -> Path:
    path = _orbit_today_path()
    write_json_file(
        path,
        {
            "generated_at": orbit.generated_at,
            "date": orbit.date,
            "brainz_awake": orbit.brainz_awake,
            "notification_count_today": orbit.notification_count_today,
            "unread_count": orbit.unread_count,
            "quieted_count": orbit.quieted_count,
            "ember_count": orbit.ember_count,
            "recent_return_count": orbit.recent_return_count,
            "summary_lines": orbit.summary_lines,
            "next_candidates": [
                {
                    "id": candidate.id,
                    "title": candidate.title,
                    "message": candidate.message,
                    "reason": candidate.reason,
                    "related_path": candidate.related_path,
                }
                for candidate in orbit.next_candidates
            ],
        },
    )
    return path


def _build_orbit_summary_lines(
    today_notifications: list[QpscNotification],
    unread_count: int,
    quieted_count: int,
    ember_count: int,
    recent_return_count: int,
    brainz_awake: bool,
    query: str,
) -> list[str]:
    lines: list[str] = []
    source_counts: dict[str, int] = {}
    for notification in today_notifications:
        source = notification.source.strip() or UI_TEXT["section_qpsc_notifications"]
        source_counts[source] = source_counts.get(source, 0) + 1
    for source, _count in sorted(source_counts.items(), key=lambda item: (-item[1], item[0]))[:2]:
        lines.append(_orbit_flow_line_for_source(source))
    if unread_count:
        lines.append(UI_TEXT["orbit_flow_unread"].format(count=unread_count))
    if quieted_count:
        lines.append(UI_TEXT["orbit_flow_quieted"].format(count=quieted_count))
    if ember_count:
        lines.append(UI_TEXT["orbit_flow_ember"].format(count=ember_count))
    if recent_return_count:
        lines.append(UI_TEXT["orbit_flow_recent_returns"].format(count=recent_return_count))
    if query.strip():
        lines.append(UI_TEXT["orbit_flow_query"].format(query=query.strip()))
    if brainz_awake:
        lines.append(UI_TEXT["orbit_flow_awake"])
    if not lines:
        lines.append(UI_TEXT["orbit_flow_quiet"])
    return lines[:5]


def _orbit_candidate_from_notification(
    notification: QpscNotification,
    reason: str,
    message: str,
    score: int,
) -> OrbitCandidate:
    return OrbitCandidate(
        id=notification.id,
        title=_notification_display_title(notification) or UI_TEXT["heat_candidate_title"],
        message=message,
        reason=reason,
        related_path=notification.related_path.strip(),
        score=score + (80 if notification.related_path.strip() else 0),
    )


def _append_orbit_candidate(candidates: list[OrbitCandidate], seen: set[str], candidate: OrbitCandidate) -> None:
    key = candidate.related_path.strip() or candidate.id or candidate.title
    if key in seen:
        return
    seen.add(key)
    candidates.append(candidate)


def _notification_view_priority(notification: QpscNotification, heat_related: bool, target_date: date) -> int:
    if notification.status == "unread":
        return 0
    if heat_related:
        return 1
    if _notification_is_today(notification, target_date):
        return 2
    if notification.related_path.strip():
        return 3
    return 4


def _notification_is_sedimented(notification: QpscNotification, target_date: date) -> bool:
    if notification.status != "read":
        return False
    age = _notification_age_days(notification, target_date)
    return age >= 3 or not notification.related_path.strip()


def _notification_age_days(notification: QpscNotification, target_date: date) -> int:
    created_date = _notification_created_date(notification)
    if not created_date:
        return 0
    return max(0, (target_date - created_date).days)


def _notification_timestamp(notification: QpscNotification) -> float:
    parsed = _parse_notification_datetime(notification.created_at)
    if not parsed:
        return 0.0
    return parsed.timestamp()


def _notification_quieted_today(notification: QpscNotification, target_date: date) -> bool:
    return notification.status == "read" and _notification_is_today(notification, target_date)


def _notification_is_today(notification: QpscNotification, target_date: date) -> bool:
    return _notification_created_date(notification) == target_date


def _notification_created_date(notification: QpscNotification) -> date | None:
    parsed = _parse_notification_datetime(notification.created_at)
    if parsed:
        return parsed.date()
    return None


def _parse_notification_datetime(value: str) -> datetime | None:
    text = value.strip()
    if not text:
        return None
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError:
        try:
            parsed_date = date.fromisoformat(text[:10])
        except ValueError:
            return None
        return datetime.combine(parsed_date, datetime.min.time()).astimezone()
    if parsed.tzinfo:
        parsed = parsed.astimezone()
    return parsed


def _orbit_today_path() -> Path:
    return app_dir() / "data" / "config" / ORBIT_FILE_NAME


def _recent_returns_path() -> Path:
    return app_dir() / "data" / "config" / RECENT_RETURNS_FILE_NAME


def read_recent_returns(limit: int = 30) -> list[RecentReturn]:
    path = _recent_returns_path()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    if not isinstance(data, list):
        return []
    returns: list[RecentReturn] = []
    for item in data:
        if not isinstance(item, dict):
            continue
        returns.append(
            RecentReturn(
                opened_at=str(item.get("opened_at", "") or ""),
                title=str(item.get("title", "") or ""),
                related_path=str(item.get("related_path", "") or ""),
            )
        )
    return returns[:limit]


def record_recent_return(path: Path, title: str) -> None:
    opened_at = datetime.now().astimezone().isoformat(timespec="seconds")
    related_path = str(path)
    current = read_recent_returns(limit=50)
    next_items = [item for item in current if item.related_path != related_path]
    next_items.insert(0, RecentReturn(opened_at=opened_at, title=title.strip() or path.stem, related_path=related_path))
    _write_recent_returns(next_items[:50])


def _write_recent_returns(items: list[RecentReturn]) -> None:
    path = _recent_returns_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = [
        {
            "opened_at": item.opened_at,
            "title": item.title,
            "related_path": item.related_path,
        }
        for item in items
    ]
    temp_path = path.with_suffix(path.suffix + ".tmp")
    temp_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    temp_path.replace(path)


def build_qpsc_self_check(
    memory_root: Path | None,
    preview_loaded: bool,
    status_paths: list[Path] | None = None,
    notification_paths: list[Path] | None = None,
    orbit_path: Path | None = None,
    recent_returns_path: Path | None = None,
) -> QpscSelfCheckResult:
    status_exists, status_data = _read_first_json_dict(status_paths or brainz_status_candidates())
    brainz_line = (
        UI_TEXT["qpsc_check_brainz_ok"]
        if status_exists and is_brainz_awake(status_data)
        else UI_TEXT["qpsc_check_brainz_unconfirmed"]
    )

    notifications_exists, notification_items = _read_first_json_list(notification_paths or brainz_notification_candidates())
    notification_line = (
        UI_TEXT["qpsc_check_notifications_ok"]
        if notifications_exists
        else UI_TEXT["qpsc_check_notifications_missing"]
    )

    orbit_exists, _orbit_data = _read_json_dict(orbit_path or _orbit_today_path())
    orbit_line = UI_TEXT["qpsc_check_orbit_ok"] if orbit_exists else UI_TEXT["qpsc_check_orbit_missing"]

    recent_returns = _read_recent_returns_from_path(recent_returns_path or _recent_returns_path())
    recent_line = UI_TEXT["qpsc_check_recent_returns"].format(count=len(recent_returns))

    related_line = _self_check_related_path_line(notification_items, memory_root)
    preview_line = UI_TEXT["qpsc_check_preview_ok"] if preview_loaded else UI_TEXT["qpsc_check_preview_waiting"]
    return QpscSelfCheckResult(
        lines=[
            brainz_line,
            notification_line,
            preview_line,
            orbit_line,
            recent_line,
            related_line,
        ]
    )


def _read_first_json_dict(paths: list[Path]) -> tuple[bool, dict[str, object]]:
    for path in paths:
        exists, data = _read_json_dict(path)
        if exists:
            return exists, data
    return False, {}


def _read_json_dict(path: Path) -> tuple[bool, dict[str, object]]:
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return False, {}
    return (True, raw) if isinstance(raw, dict) else (False, {})


def _read_first_json_list(paths: list[Path]) -> tuple[bool, list[dict[str, object]]]:
    for path in paths:
        exists, data = _read_json_list(path)
        if exists:
            return exists, data
    return False, []


def _read_json_list(path: Path) -> tuple[bool, list[dict[str, object]]]:
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return False, []
    if not isinstance(raw, list):
        return False, []
    return True, [item for item in raw if isinstance(item, dict)]


def _read_recent_returns_from_path(path: Path) -> list[RecentReturn]:
    exists, data = _read_json_list(path)
    if not exists:
        return []
    returns: list[RecentReturn] = []
    for item in data:
        returns.append(
            RecentReturn(
                opened_at=str(item.get("opened_at", "") or ""),
                title=str(item.get("title", "") or ""),
                related_path=str(item.get("related_path", "") or ""),
            )
        )
    return returns


def _self_check_related_path_line(notification_items: list[dict[str, object]], memory_root: Path | None) -> str:
    related_paths = [str(item.get("related_path", "") or "").strip() for item in notification_items]
    related_paths = [path for path in related_paths if path]
    if not related_paths:
        return UI_TEXT["qpsc_check_related_empty"]
    for related_path in related_paths:
        resolved = _resolve_related_source_path(related_path, memory_root)
        if resolved.exists() and resolved.is_file():
            return UI_TEXT["qpsc_check_related_ok"]
    return UI_TEXT["qpsc_check_related_missing"]


def _count_recent_returns(recent_returns: list[RecentReturn], target_date: date) -> int:
    return sum(1 for item in recent_returns if _recent_return_is_today(item, target_date))


def _recent_return_is_today(item: RecentReturn, target_date: date) -> bool:
    parsed = _parse_notification_datetime(item.opened_at)
    return bool(parsed and parsed.date() == target_date)


def _document_search_score(document: MemoryDocument, terms: list[str]) -> int:
    title_text = document.title.lower()
    path_text = document.relative_path.lower()
    body_text = document.text.lower()
    score = _text_search_score(body_text, terms)
    for term in terms:
        if term in title_text:
            score += 6
        if term in path_text:
            score += 3
    return score


def _text_search_score(text: str, terms: list[str]) -> int:
    lowered = text.lower()
    return sum(lowered.count(term) for term in terms)


def _notification_heat_score(notification: QpscNotification, index: int, terms: list[str]) -> int:
    score = max(0, 120 - index)
    if notification.status == "unread":
        score += 600
    if notification.related_path.strip():
        score += 80
    if _is_codex_notification(notification) and notification.related_path.strip():
        score += 160
    text = _notification_search_text(notification)
    for term in terms:
        if term in text:
            score += 180
    return score


def _notification_heat_reason(notification: QpscNotification, terms: list[str]) -> str:
    if notification.status == "unread":
        return "unread_notification"
    text = _notification_search_text(notification)
    if terms and any(term in text for term in terms):
        return "query_match"
    if notification.related_path.strip():
        return "related_path"
    return "recent_import"


def _notification_boosts(
    notifications: list[QpscNotification],
    memory_root: Path,
    terms: list[str],
) -> dict[Path, int]:
    boosts: dict[Path, int] = {}
    for index, notification in enumerate(notifications):
        related_path = notification.related_path.strip()
        if not related_path:
            continue
        path = _resolve_related_source_path(related_path, memory_root)
        score = 0
        if notification.status == "unread":
            score += 55
        score += max(0, 35 - index)
        score += 10
        text = _notification_search_text(notification)
        for term in terms:
            if term in text:
                score += 90
        if score > 0:
            resolved = path.resolve()
            boosts[resolved] = max(boosts.get(resolved, 0), score)
    return boosts


def _notification_search_text(notification: QpscNotification) -> str:
    return " ".join(
        [
            _notification_display_title(notification),
            _notification_display_message(notification),
            notification.source,
            notification.kind,
        ]
    ).lower()


def _resolve_related_source_path(related_path: str | Path, memory_root: Path | None) -> Path:
    path = Path(str(related_path).strip())
    if path.is_absolute() or not memory_root:
        return path
    return memory_root / path


def _relative_path(path: Path, memory_root: Path) -> str:
    try:
        return str(path.resolve().relative_to(memory_root.resolve()))
    except ValueError:
        return str(path)


def _search_excerpt(text: str, terms: list[str], width: int = 160) -> str:
    compact = " ".join(text.split())
    lowered = compact.lower()
    index = -1
    for term in terms:
        index = lowered.find(term)
        if index >= 0:
            break
    if index < 0:
        return compact[:width]
    left = max(0, index - width // 2)
    right = min(len(compact), index + width // 2)
    excerpt = compact[left:right].strip()
    if left > 0:
        excerpt = f"...{excerpt}"
    if right < len(compact):
        excerpt = f"{excerpt}..."
    return excerpt


def read_source_preview_text(path: Path, limit: int = MAX_PREVIEW_CHARS) -> tuple[str, bool]:
    for encoding in ("utf-8-sig", "utf-8", "cp932", "utf-16"):
        try:
            with path.open("r", encoding=encoding) as handle:
                text = handle.read(limit + 1)
            return text[:limit], len(text) > limit
        except UnicodeDecodeError:
            continue
        except OSError:
            return "", False
    try:
        with path.open("r", encoding="utf-8", errors="replace") as handle:
            text = handle.read(limit + 1)
        return text[:limit], len(text) > limit
    except OSError:
        return "", False

COLORS = {
    "background": "#080A0F",
    "panel": "#11151C",
    "panel_light": "#151A22",
    "text": "#E6EAF0",
    "muted": "#9AA4B2",
    "glow": "#263247",
    "heat": "#7AA7FF",
    "line": "#252B36",
    "line_soft": "#1A202B",
}

FONT_JP = ("Yu Gothic UI", 12)
FONT_JP_SMALL = ("Yu Gothic UI", 10)
FONT_TITLE = ("Yu Gothic UI", 20)
FONT_LABEL = ("Yu Gothic UI", 10)
FONT_MONO = ("Cascadia Mono", 10)


@dataclass
class Particle:
    x: float
    y: float
    vx: float
    vy: float
    size: float
    phase: float


class OikawaApp(tk.Tk):
    def __init__(
        self,
        launch_check: bool = False,
        gui_smoke_seconds: float = 0.0,
        memory_folder_override: str = "",
    ) -> None:
        super().__init__()
        self.title(WINDOW_TITLE)
        self.geometry("1240x760")
        self.minsize(1120, 680)
        self.configure(bg=COLORS["background"])

        self.config_store = ConfigStore()
        self.config_data = self.config_store.load()
        self.memory_folder = self._resolve_initial_memory_folder(memory_folder_override)
        self.output_path: Path | None = None
        self.events: queue.Queue[tuple[str, object]] = queue.Queue()
        self.scan_thread: threading.Thread | None = None
        self.search_thread: threading.Thread | None = None
        self.orbit_thread: threading.Thread | None = None
        self.source_preview_thread: threading.Thread | None = None
        self.self_check_thread: threading.Thread | None = None
        self.source_preview_request_id = 0
        self.orbit_request_id = 0
        self.source_preview_loaded = False
        self.qpsc_notifications_all: list[QpscNotification] = []
        self.heat_candidates: list[HeatCandidate] = []
        self.last_search_query = ""
        self.scanning = False
        self.particles: list[Particle] = []
        self.last_animation = time.monotonic()

        self._build_canvas()
        self._build_overlay()
        self._init_particles()
        self._render_memory_state()
        self._render_empty_cards()
        self._refresh_qpsc_notifications()

        self.after(80, self._animate)
        self.after(100, self._poll_events)

        if launch_check:
            self.after(500, self._finish_launch_check)
        elif gui_smoke_seconds > 0:
            self.after(max(500, int(gui_smoke_seconds * 1000)), self._finish_gui_smoke)

    def _resolve_initial_memory_folder(self, override: str) -> Path | None:
        if override:
            folder = existing_folder(override)
            if folder:
                self.config_data.memory_folder = str(folder)
                self.config_store.save(self.config_data)
            return folder

        folder = resolve_memory_folder(self.config_data)
        if folder and self.config_data.memory_folder != str(folder):
            self.config_data.memory_folder = str(folder)
            self.config_store.save(self.config_data)
        return folder

    def _build_canvas(self) -> None:
        self.canvas = tk.Canvas(
            self,
            bg=COLORS["background"],
            highlightthickness=0,
            bd=0,
        )
        self.canvas.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.canvas.bind("<Configure>", lambda _event: self._init_particles())

    def _build_overlay(self) -> None:
        self.status_var = tk.StringVar(value=UI_TEXT["status_idle"])
        self.memory_var = tk.StringVar(value="")
        self.summary_var = tk.StringVar(value=UI_TEXT["summary_idle"])
        self.qpsc_notification_var = tk.StringVar(value=UI_TEXT["qpsc_waiting"])
        self.orbit_metrics_var = tk.StringVar(value=UI_TEXT["orbit_status_loading"])
        self.orbit_flow_var = tk.StringVar(value=UI_TEXT["orbit_status_loading"])
        self.qpsc_self_check_var = tk.StringVar(value=self._initial_qpsc_self_check_text())

        left_panel = tk.Frame(self, bg=COLORS["background"])
        left_panel.place(relx=0.03, rely=0.055, relwidth=0.21, relheight=0.84)
        tk.Label(
            left_panel,
            text=UI_TEXT["app_title"],
            fg=COLORS["text"],
            bg=COLORS["background"],
            font=FONT_TITLE,
        ).pack(anchor="w")
        tk.Label(
            left_panel,
            text=UI_TEXT["app_subcopy"],
            fg=COLORS["heat"],
            bg=COLORS["background"],
            font=FONT_JP,
        ).pack(anchor="w", pady=(5, 0))
        tk.Label(
            left_panel,
            text=UI_TEXT["app_subtitle"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
            font=FONT_JP_SMALL,
        ).pack(anchor="w", pady=(4, 14))
        tk.Label(
            left_panel,
            textvariable=self.status_var,
            fg=COLORS["heat"],
            bg=COLORS["background"],
            font=FONT_JP_SMALL,
        ).pack(anchor="w")
        tk.Label(
            left_panel,
            textvariable=self.memory_var,
            fg=COLORS["muted"],
            bg=COLORS["background"],
            font=FONT_MONO,
            wraplength=240,
            justify="left",
        ).pack(anchor="w", pady=(8, 0))
        search_row = tk.Frame(left_panel, bg=COLORS["background"])
        search_row.pack(fill="x", pady=(14, 0))
        self.search_entry = tk.Entry(
            search_row,
            bg=COLORS["panel"],
            fg=COLORS["text"],
            insertbackground=COLORS["text"],
            relief="flat",
            bd=0,
            font=FONT_JP_SMALL,
        )
        self.search_entry.pack(fill="x", ipady=7)
        self.search_entry.bind("<Return>", lambda _event: self._start_memory_search())
        search_buttons = tk.Frame(left_panel, bg=COLORS["background"])
        search_buttons.pack(fill="x", pady=(8, 0))
        self.search_button = tk.Button(
            search_buttons,
            text=UI_TEXT["button_search"],
            command=self._start_memory_search,
            **self._button_style(COLORS["panel_light"]),
        )
        self.search_button.pack(side="left")
        self.heat_search_button = tk.Button(
            search_buttons,
            text=UI_TEXT["button_heat_search"],
            command=self._start_heat_search,
            **self._button_style(COLORS["heat"]),
        )
        self.heat_search_button.pack(side="left", padx=(8, 0))

        self.missing_frame = tk.Frame(
            self,
            bg=COLORS["panel"],
            highlightbackground=COLORS["line"],
            highlightthickness=1,
            padx=18,
            pady=16,
        )
        tk.Label(
            self.missing_frame,
            text=UI_TEXT["memory_folder_missing"],
            fg=COLORS["text"],
            bg=COLORS["panel"],
            font=FONT_JP,
        ).pack(anchor="w")
        tk.Button(
            self.missing_frame,
            text=UI_TEXT["button_choose_folder"],
            command=self._choose_memory_folder,
            **self._button_style(COLORS["panel_light"]),
        ).pack(anchor="w", pady=(12, 0))

        self.results_frame = tk.Frame(left_panel, bg=COLORS["background"])
        self.results_frame.pack(fill="both", expand=True, pady=(18, 0))
        tk.Label(
            self.results_frame,
            text=UI_TEXT["section_related"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
            font=FONT_JP_SMALL,
        ).pack(anchor="w")
        self.cards_frame = tk.Frame(self.results_frame, bg=COLORS["background"])
        self.cards_frame.pack(fill="both", expand=True, pady=(8, 0))

        actions = tk.Frame(left_panel, bg=COLORS["background"])
        actions.pack(fill="x", pady=(12, 0))
        self.open_output_button = tk.Button(
            actions,
            text=UI_TEXT["button_open_output"],
            command=self._open_output,
            state="disabled",
            **self._small_button_style(COLORS["panel_light"]),
        )
        self.open_output_button.pack(side="left")
        self.scan_button = tk.Button(
            actions,
            text=UI_TEXT["button_scan"],
            command=self._start_scan,
            **self._small_button_style(COLORS["heat"]),
        )
        self.scan_button.pack(side="left", padx=(8, 0))

        qpsc_state = tk.Frame(
            left_panel,
            bg=COLORS["background"],
            highlightbackground=COLORS["line_soft"],
            highlightthickness=1,
            padx=10,
            pady=8,
        )
        qpsc_state.pack(fill="x", pady=(10, 0))
        qpsc_state_header = tk.Frame(qpsc_state, bg=COLORS["background"])
        qpsc_state_header.pack(fill="x")
        tk.Label(
            qpsc_state_header,
            text=UI_TEXT["section_qpsc_state"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
            font=FONT_LABEL,
        ).pack(side="left", anchor="w")
        self.qpsc_self_check_button = tk.Button(
            qpsc_state_header,
            text=UI_TEXT["button_qpsc_state_check"],
            command=self._start_qpsc_self_check,
            **self._small_button_style(COLORS["panel_light"]),
        )
        self.qpsc_self_check_button.pack(side="right")
        tk.Label(
            qpsc_state,
            textvariable=self.qpsc_self_check_var,
            fg=COLORS["text"],
            bg=COLORS["background"],
            font=FONT_JP_SMALL,
            wraplength=240,
            justify="left",
        ).pack(anchor="w", pady=(6, 0))

        self.source_preview_status_var = tk.StringVar(value=UI_TEXT["source_preview_empty"])
        source_preview = tk.Frame(
            self,
            bg=COLORS["panel"],
            highlightbackground=COLORS["line"],
            highlightthickness=1,
            padx=16,
            pady=14,
        )
        source_preview.place(relx=0.27, rely=0.055, relwidth=0.44, relheight=0.84)
        tk.Label(
            source_preview,
            text=UI_TEXT["section_source_preview"],
            fg=COLORS["muted"],
            bg=COLORS["panel"],
            font=FONT_LABEL,
        ).pack(anchor="w")
        tk.Label(
            source_preview,
            textvariable=self.source_preview_status_var,
            fg=COLORS["text"],
            bg=COLORS["panel"],
            font=FONT_JP_SMALL,
            wraplength=500,
            justify="left",
        ).pack(anchor="w", pady=(6, 10))
        preview_body = tk.Frame(source_preview, bg=COLORS["panel"])
        preview_body.pack(fill="both", expand=True)
        preview_scrollbar = tk.Scrollbar(preview_body)
        preview_scrollbar.pack(side="right", fill="y")
        self.source_preview_box = tk.Text(
            preview_body,
            bg=COLORS["background"],
            fg=COLORS["text"],
            insertbackground=COLORS["text"],
            relief="flat",
            bd=0,
            wrap="word",
            font=FONT_JP,
            yscrollcommand=preview_scrollbar.set,
        )
        self.source_preview_box.pack(side="left", fill="both", expand=True)
        preview_scrollbar.configure(command=self.source_preview_box.yview)
        self._set_source_preview_text(UI_TEXT["source_preview_empty"])

        right_panel = tk.Frame(self, bg=COLORS["background"])
        right_panel.place(relx=0.74, rely=0.055, relwidth=0.23, relheight=0.84)

        notice = tk.Frame(
            right_panel,
            bg=COLORS["background"],
            highlightbackground=COLORS["line_soft"],
            highlightthickness=1,
            padx=12,
            pady=9,
        )
        notice.pack(fill="x")
        tk.Label(
            notice,
            text=UI_TEXT["section_qpsc_notifications"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
            font=FONT_LABEL,
        ).pack(anchor="w")
        self.qpsc_status_label = tk.Label(
            notice,
            textvariable=self.qpsc_notification_var,
            fg=COLORS["text"],
            bg=COLORS["background"],
            font=FONT_JP_SMALL,
            justify="left",
            wraplength=250,
        )
        self.qpsc_status_label.pack(anchor="w", pady=(6, 0))
        self.qpsc_import_frame = tk.Frame(notice, bg=COLORS["background"])
        self.qpsc_import_frame.pack(fill="x", pady=(8, 0))

        heat = tk.Frame(
            right_panel,
            bg=COLORS["background"],
            highlightbackground=COLORS["line"],
            highlightthickness=1,
            padx=12,
            pady=8,
        )
        heat.pack(fill="x", pady=(0, 10), before=notice)
        tk.Label(
            heat,
            text=UI_TEXT["section_heat_candidates"],
            fg=COLORS["heat"],
            bg=COLORS["background"],
            font=FONT_LABEL,
        ).pack(anchor="w")
        self.heat_candidates_frame = tk.Frame(heat, bg=COLORS["background"])
        self.heat_candidates_frame.pack(fill="both", expand=True, pady=(5, 0))

        orbit = tk.Frame(
            right_panel,
            bg=COLORS["background"],
            highlightbackground=COLORS["line_soft"],
            highlightthickness=1,
            padx=12,
            pady=9,
        )
        orbit.pack(fill="both", expand=True, pady=(0, 10), before=notice)
        orbit_top = tk.Frame(orbit, bg=COLORS["background"])
        orbit_top.pack(fill="x")
        tk.Label(
            orbit_top,
            text=UI_TEXT["section_orbit_today"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
            font=FONT_LABEL,
        ).pack(anchor="w")
        tk.Label(
            orbit,
            textvariable=self.orbit_metrics_var,
            fg=COLORS["text"],
            bg=COLORS["background"],
            font=FONT_JP_SMALL,
            justify="left",
            wraplength=250,
        ).pack(anchor="w", pady=(5, 0))
        orbit_body = tk.Frame(orbit, bg=COLORS["background"])
        orbit_body.pack(fill="both", expand=True, pady=(6, 0))
        orbit_flow = tk.Frame(orbit_body, bg=COLORS["background"])
        orbit_flow.pack(fill="x")
        tk.Label(
            orbit_flow,
            text=UI_TEXT["section_orbit_flow"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
            font=FONT_LABEL,
        ).pack(anchor="w")
        tk.Label(
            orbit_flow,
            textvariable=self.orbit_flow_var,
            fg=COLORS["text"],
            bg=COLORS["background"],
            font=FONT_JP_SMALL,
            justify="left",
            wraplength=250,
        ).pack(anchor="w", pady=(3, 0))
        orbit_next = tk.Frame(orbit_body, bg=COLORS["background"])
        orbit_next.pack(fill="both", expand=True, pady=(8, 0))
        tk.Label(
            orbit_next,
            text=UI_TEXT["section_orbit_next"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
            font=FONT_LABEL,
        ).pack(anchor="w")
        self.orbit_next_frame = tk.Frame(orbit_next, bg=COLORS["background"])
        self.orbit_next_frame.pack(fill="both", expand=True, pady=(3, 0))

        footer = tk.Frame(self, bg=COLORS["background"])
        footer.place(relx=0.03, rely=0.93, relwidth=0.94, relheight=0.055)
        tk.Label(
            footer,
            text=UI_TEXT["copyright"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
            font=FONT_LABEL,
        ).pack(side="left", anchor="w")
        tk.Label(
            footer,
            text=UI_TEXT["footer_source"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
            font=FONT_JP_SMALL,
        ).pack(side="left", padx=(18, 0))
        tk.Label(
            footer,
            textvariable=self.summary_var,
            fg=COLORS["muted"],
            bg=COLORS["background"],
            font=FONT_MONO,
        ).pack(side="right", anchor="e")

    def _initial_qpsc_self_check_text(self) -> str:
        return "\n".join(
            [
                UI_TEXT["qpsc_check_brainz_unconfirmed"],
                UI_TEXT["qpsc_check_notifications_missing"],
                UI_TEXT["qpsc_check_preview_waiting"],
                UI_TEXT["qpsc_check_orbit_missing"],
                UI_TEXT["qpsc_check_recent_returns"].format(count=0),
                UI_TEXT["qpsc_check_related_empty"],
            ]
        )

    def _start_qpsc_self_check(self) -> None:
        if self.self_check_thread and self.self_check_thread.is_alive():
            return
        self.qpsc_self_check_button.configure(text=UI_TEXT["button_qpsc_state_checking"], state="disabled")
        memory_folder = self.memory_folder
        preview_loaded = self.source_preview_loaded
        self.self_check_thread = threading.Thread(
            target=self._qpsc_self_check_worker,
            args=(memory_folder, preview_loaded),
            daemon=True,
        )
        self.self_check_thread.start()

    def _qpsc_self_check_worker(self, memory_folder: Path | None, preview_loaded: bool) -> None:
        result = build_qpsc_self_check(memory_folder, preview_loaded)
        self.events.put(("self_check_done", result))

    def _handle_self_check_done(self, payload: object) -> None:
        assert isinstance(payload, QpscSelfCheckResult)
        self.qpsc_self_check_var.set("\n".join(payload.lines))
        self.qpsc_self_check_button.configure(text=UI_TEXT["button_qpsc_state_check"], state="normal")

    def _refresh_qpsc_notifications(self, schedule: bool = True) -> None:
        status = read_brainz_status()
        brainz_awake = is_brainz_awake(status)
        if brainz_awake:
            line1 = UI_TEXT["qpsc_brainz_awake"]
            line2 = UI_TEXT["qpsc_memory_awake"]
        elif status:
            line1 = UI_TEXT["qpsc_heartbeat_quiet"]
            line2 = UI_TEXT["qpsc_memory_heartbeat_quiet"]
        else:
            line1 = UI_TEXT["qpsc_waiting"]
            line2 = UI_TEXT["qpsc_memory_waiting"]
        self.qpsc_notification_var.set(
            UI_TEXT["qpsc_notice_template"].format(
                line1=line1,
                line2=line2,
                line3=UI_TEXT["qpsc_no_suggestion"],
            )
        )
        self.qpsc_notifications_all = read_qpsc_notification_events(limit=30)
        self._refresh_heat_candidates()
        notifications, sedimented_count = build_notification_view(self.qpsc_notifications_all, self.heat_candidates, limit=3)
        self._render_qpsc_import_notifications(notifications, sedimented_count)
        self._start_orbit_refresh(brainz_awake=brainz_awake)
        if schedule:
            self.after(10000, self._refresh_qpsc_notifications)

    def _refresh_heat_candidates(self, query: str = "") -> None:
        self.heat_candidates = build_heat_candidates(self.qpsc_notifications_all, query=query, limit=3)
        self._render_heat_candidates(self.heat_candidates)

    def _start_orbit_refresh(self, query: str = "", brainz_awake: bool | None = None) -> None:
        self.orbit_request_id += 1
        request_id = self.orbit_request_id
        if brainz_awake is None:
            brainz_awake = is_brainz_awake(read_brainz_status())
        notifications = list(self.qpsc_notifications_all)
        heat_candidates = list(self.heat_candidates)
        if not query and hasattr(self, "search_entry"):
            query = self.search_entry.get().strip() or self.last_search_query
        self.orbit_thread = threading.Thread(
            target=self._orbit_worker,
            args=(request_id, notifications, heat_candidates, brainz_awake, query),
            daemon=True,
        )
        self.orbit_thread.start()

    def _orbit_worker(
        self,
        request_id: int,
        notifications: list[QpscNotification],
        heat_candidates: list[HeatCandidate],
        brainz_awake: bool,
        query: str,
    ) -> None:
        try:
            orbit = build_orbit_today(notifications, heat_candidates, brainz_awake, recent_returns=read_recent_returns(), query=query)
            save_orbit_today(orbit)
            self.events.put(("orbit_done", (request_id, orbit)))
        except Exception as exc:  # noqa: BLE001
            self.events.put(("orbit_error", (request_id, exc)))

    def _render_qpsc_import_notifications(self, notifications: list[NotificationView], sedimented_count: int) -> None:
        if not hasattr(self, "qpsc_import_frame"):
            return
        for child in self.qpsc_import_frame.winfo_children():
            child.destroy()
        if not notifications:
            tk.Label(
                self.qpsc_import_frame,
                text=UI_TEXT["qpsc_notification_empty"],
                fg=COLORS["muted"],
                bg=COLORS["background"],
                font=FONT_JP_SMALL,
                wraplength=260,
                justify="left",
            ).pack(anchor="w")

        for view_item in notifications:
            notification = view_item.notification
            display_title = _notification_display_title(notification)
            display_message = _notification_display_message(notification)
            event_frame = tk.Frame(self.qpsc_import_frame, bg=COLORS["background"])
            event_frame.pack(fill="x", pady=(0, 8))
            title_color = COLORS["heat"] if view_item.heat_related else COLORS["text"]
            if notification.status == "read":
                title_color = COLORS["muted"]
            tk.Label(
                event_frame,
                text=display_title,
                fg=title_color,
                bg=COLORS["background"],
                font=FONT_JP_SMALL,
                wraplength=250,
                justify="left",
            ).pack(anchor="w")
            if display_message:
                tk.Label(
                    event_frame,
                    text=display_message,
                    fg=COLORS["muted"],
                    bg=COLORS["background"],
                    font=FONT_JP_SMALL,
                    wraplength=250,
                    justify="left",
                ).pack(anchor="w", pady=(2, 0))
            controls = tk.Frame(event_frame, bg=COLORS["background"])
            controls.pack(anchor="w", pady=(4, 0))
            if notification.related_path:
                tk.Button(
                    controls,
                    text=UI_TEXT["qpsc_notification_open_related"],
                    command=lambda item=notification: self._open_qpsc_related_path(item),
                    **self._small_button_style(COLORS["panel_light"]),
                ).pack(side="left", padx=(0, 6))
            tk.Button(
                controls,
                text=UI_TEXT["qpsc_notification_read"],
                command=lambda item=notification: self._mark_qpsc_notification_read(item),
                **self._small_button_style(COLORS["panel_light"]),
            ).pack(side="left")
        if sedimented_count:
            tk.Label(
                self.qpsc_import_frame,
                text=UI_TEXT["qpsc_notification_sedimented"].format(count=sedimented_count),
                fg=COLORS["muted"],
                bg=COLORS["background"],
                font=FONT_LABEL,
                wraplength=250,
                justify="left",
            ).pack(anchor="w", pady=(2, 0))

    def _render_heat_candidates(self, candidates: list[HeatCandidate]) -> None:
        if not hasattr(self, "heat_candidates_frame"):
            return
        for child in self.heat_candidates_frame.winfo_children():
            child.destroy()
        if not candidates:
            tk.Label(
                self.heat_candidates_frame,
                text=UI_TEXT["heat_candidate_empty"],
                fg=COLORS["muted"],
                bg=COLORS["background"],
                font=FONT_JP_SMALL,
                wraplength=300,
                justify="left",
            ).pack(anchor="w")
            return
        for candidate in candidates[:3]:
            row = tk.Frame(self.heat_candidates_frame, bg=COLORS["background"], cursor="hand2")
            row.pack(fill="x", pady=(0, 4))
            row.bind("<Button-1>", lambda _event, selected=candidate: self._open_heat_candidate(selected))
            title_label = tk.Label(
                row,
                text=candidate.title,
                fg=COLORS["text"],
                bg=COLORS["background"],
                font=FONT_JP_SMALL,
                wraplength=250,
                justify="left",
                cursor="hand2",
            )
            title_label.pack(side="left", fill="x", expand=True)
            title_label.bind("<Button-1>", lambda _event, selected=candidate: self._open_heat_candidate(selected))
            reason_text = UI_TEXT.get(f"heat_reason_{candidate.reason}", UI_TEXT["section_heat_candidates"])
            reason_label = tk.Label(
                row,
                text=reason_text,
                fg=COLORS["muted"],
                bg=COLORS["background"],
                font=FONT_LABEL,
                cursor="hand2",
            )
            reason_label.pack(side="right", padx=(8, 0))
            reason_label.bind("<Button-1>", lambda _event, selected=candidate: self._open_heat_candidate(selected))

    def _handle_orbit_done(self, payload: object) -> None:
        request_id, orbit = payload
        assert isinstance(request_id, int)
        assert isinstance(orbit, OrbitToday)
        if request_id != self.orbit_request_id:
            return
        metrics = [
            UI_TEXT["orbit_metric_today"].format(count=orbit.notification_count_today),
            UI_TEXT["orbit_metric_unread"].format(count=orbit.unread_count),
            UI_TEXT["orbit_metric_quieted"].format(count=orbit.quieted_count),
            UI_TEXT["orbit_metric_ember"].format(count=orbit.ember_count),
            UI_TEXT["orbit_metric_recent_returns"].format(count=orbit.recent_return_count),
            UI_TEXT["orbit_metric_awake"] if orbit.brainz_awake else UI_TEXT["orbit_metric_quiet"],
        ]
        self.orbit_metrics_var.set("\n".join(metrics))
        flow_text = "\n".join(UI_TEXT["orbit_bullet"].format(text=line) for line in orbit.summary_lines[:5])
        self.orbit_flow_var.set(flow_text)
        self._render_orbit_next_candidates(orbit.next_candidates)

    def _handle_orbit_error(self, payload: object) -> None:
        request_id, _error = payload
        assert isinstance(request_id, int)
        if request_id != self.orbit_request_id:
            return
        self.orbit_metrics_var.set(UI_TEXT["orbit_status_unavailable"])
        self.orbit_flow_var.set(UI_TEXT["orbit_status_unavailable"])
        self._render_orbit_next_candidates([])

    def _render_orbit_next_candidates(self, candidates: list[OrbitCandidate]) -> None:
        if not hasattr(self, "orbit_next_frame"):
            return
        for child in self.orbit_next_frame.winfo_children():
            child.destroy()
        if not candidates:
            tk.Label(
                self.orbit_next_frame,
                text=UI_TEXT["orbit_next_empty"],
                fg=COLORS["muted"],
                bg=COLORS["background"],
                font=FONT_JP_SMALL,
                wraplength=220,
                justify="left",
            ).pack(anchor="w")
            return
        for candidate in candidates[:3]:
            row = tk.Frame(self.orbit_next_frame, bg=COLORS["background"], cursor="hand2")
            row.pack(fill="x", pady=(0, 3))
            row.bind("<Button-1>", lambda _event, selected=candidate: self._open_orbit_candidate(selected))
            tk.Label(
                row,
                text=UI_TEXT["orbit_next_template"].format(message=candidate.message, title=candidate.title),
                fg=COLORS["text"],
                bg=COLORS["background"],
                font=FONT_JP_SMALL,
                wraplength=220,
                justify="left",
                cursor="hand2",
            ).pack(anchor="w")
            for child in row.winfo_children():
                child.bind("<Button-1>", lambda _event, selected=candidate: self._open_orbit_candidate(selected))

    def _small_button_style(self, background: str) -> dict[str, object]:
        style = self._button_style(background)
        style.update({"padx": 8, "pady": 4})
        return style

    def _mark_qpsc_notification_read(self, notification: QpscNotification) -> None:
        if notification.id:
            mark_qpsc_notification_read(notification.id)
        self._refresh_qpsc_notifications(schedule=False)

    def _open_qpsc_related_path(self, notification: QpscNotification) -> None:
        path_text = notification.related_path.strip()
        if not path_text:
            return
        self._show_source_path(path_text, notification.title)

    def _open_heat_candidate(self, candidate: HeatCandidate) -> None:
        if candidate.related_path:
            self._show_source_path(candidate.related_path, candidate.title)
            return
        self.source_preview_status_var.set(UI_TEXT["heat_candidate_no_related"])
        self._set_source_preview_text(f"{candidate.title}\n{UI_TEXT['heat_candidate_no_related']}")

    def _open_orbit_candidate(self, candidate: OrbitCandidate) -> None:
        if candidate.related_path:
            self._show_source_path(candidate.related_path, candidate.title)
            return
        self.source_preview_status_var.set(UI_TEXT["orbit_next_no_related"])
        self._set_source_preview_text(f"{candidate.title}\n{UI_TEXT['orbit_next_no_related']}")

    def _button_style(self, background: str) -> dict[str, object]:
        return {
            "bg": background,
            "fg": COLORS["text"],
            "activebackground": COLORS["glow"],
            "activeforeground": COLORS["text"],
            "disabledforeground": COLORS["muted"],
            "font": FONT_JP_SMALL,
            "relief": "flat",
            "bd": 0,
            "padx": 18,
            "pady": 9,
            "highlightthickness": 1,
            "highlightbackground": COLORS["line"],
            "cursor": "hand2",
        }

    def _render_memory_state(self) -> None:
        if self.memory_folder:
            self.memory_var.set(f"{UI_TEXT['label_memory_folder']}: {self.memory_folder}")
            self.missing_frame.place_forget()
            return

        self.memory_var.set(UI_TEXT["memory_folder_missing"])
        self.missing_frame.place(relx=0.03, rely=0.38, relwidth=0.21)

    def _init_particles(self) -> None:
        width = max(1, self.canvas.winfo_width() or 1120)
        height = max(1, self.canvas.winfo_height() or 720)
        random.seed(20260518)
        count = 32
        self.particles = [
            Particle(
                x=random.uniform(0, width),
                y=random.uniform(0, height),
                vx=random.uniform(-0.18, 0.18),
                vy=random.uniform(-0.14, 0.14),
                size=random.uniform(1.0, 2.2),
                phase=random.uniform(0, math.tau),
            )
            for _ in range(count)
        ]

    def _animate(self) -> None:
        now = time.monotonic()
        delta = min(0.2, now - self.last_animation)
        self.last_animation = now
        width = max(1, self.canvas.winfo_width())
        height = max(1, self.canvas.winfo_height())
        speed = 2.2 if self.scanning else 1.0

        self.canvas.delete("field")
        self.canvas.create_rectangle(0, 0, width, height, fill=COLORS["background"], outline="", tags="field")
        self._draw_ghost_words(width, height)

        for particle in self.particles:
            particle.phase += delta * 0.6
            particle.x += (particle.vx + math.sin(particle.phase) * 0.05) * speed * delta * 60
            particle.y += (particle.vy + math.cos(particle.phase) * 0.04) * speed * delta * 60
            if particle.x < 0:
                particle.x += width
            elif particle.x > width:
                particle.x -= width
            if particle.y < 0:
                particle.y += height
            elif particle.y > height:
                particle.y -= height

        for index, current in enumerate(self.particles):
            for other in self.particles[index + 1 :]:
                distance = math.hypot(current.x - other.x, current.y - other.y)
                if distance < 138:
                    color = COLORS["line"] if self.scanning and distance < 96 else COLORS["line_soft"]
                    self.canvas.create_line(current.x, current.y, other.x, other.y, fill=color, width=1, tags="field")

        for particle in self.particles:
            radius = particle.size * (1.25 if self.scanning else 1.0)
            self.canvas.create_oval(
                particle.x - radius,
                particle.y - radius,
                particle.x + radius,
                particle.y + radius,
                fill=COLORS["glow"],
                outline="",
                tags="field",
            )

        self.after(90, self._animate)

    def _draw_ghost_words(self, width: int, height: int) -> None:
        positions = [(0.19, 0.22), (0.70, 0.19), (0.52, 0.37), (0.82, 0.58), (0.28, 0.76), (0.64, 0.82), (0.42, 0.16)]
        for word, (x_ratio, y_ratio) in zip(UI_TEXT["ghost_words"], positions):
            self.canvas.create_text(
                width * x_ratio,
                height * y_ratio,
                text=word,
                fill="#0F1219",
                font=("BIZ UDPGothic", 18),
                tags="field",
            )

    def _start_memory_search(self) -> None:
        if self.search_thread and self.search_thread.is_alive():
            return
        if not self.memory_folder:
            self.status_var.set(UI_TEXT["memory_folder_missing"])
            self._render_memory_state()
            return
        query = self.search_entry.get().strip()
        if not query:
            return
        self.last_search_query = query

        self.status_var.set(UI_TEXT["status_searching"])
        self.summary_var.set(UI_TEXT["status_searching"])
        self.search_button.configure(text=UI_TEXT["button_searching"], state="disabled")
        self.heat_search_button.configure(state="disabled")
        self._render_loading_card(UI_TEXT["status_searching"])
        self.search_thread = threading.Thread(
            target=self._memory_search_worker,
            args=(self.memory_folder, query, False, []),
            daemon=True,
        )
        self.search_thread.start()

    def _start_heat_search(self) -> None:
        if self.search_thread and self.search_thread.is_alive():
            return
        if not self.memory_folder:
            self.status_var.set(UI_TEXT["memory_folder_missing"])
            self._render_memory_state()
            return
        query = self.search_entry.get().strip()
        if not query:
            self._refresh_heat_candidates()
            self._start_orbit_refresh()
            return
        self.last_search_query = query

        notifications = read_qpsc_notification_events(limit=30)
        self.qpsc_notifications_all = notifications
        self._refresh_heat_candidates(query)
        self.status_var.set(UI_TEXT["status_heat_searching"])
        self.summary_var.set(UI_TEXT["status_heat_searching"])
        self.search_button.configure(state="disabled")
        self.heat_search_button.configure(text=UI_TEXT["button_heat_searching"], state="disabled")
        self._render_loading_card(UI_TEXT["status_heat_searching"])
        self.search_thread = threading.Thread(
            target=self._memory_search_worker,
            args=(self.memory_folder, query, True, notifications),
            daemon=True,
        )
        self.search_thread.start()

    def _memory_search_worker(
        self,
        memory_folder: Path,
        query: str,
        heat_search: bool,
        notifications: list[QpscNotification],
    ) -> None:
        try:
            documents, skipped = scan_memory(memory_folder)
            if heat_search:
                hits = build_heat_search_hits(documents, query, memory_folder, notifications)
            else:
                hits = build_search_hits(documents, query, memory_folder)
            self.events.put(("search_done", (heat_search, query, hits, skipped)))
        except Exception as exc:  # noqa: BLE001
            self.events.put(("search_error", (heat_search, exc)))

    def _handle_search_done(self, payload: object) -> None:
        heat_search, query, hits, skipped = payload
        assert isinstance(heat_search, bool)
        assert isinstance(query, str)
        assert isinstance(hits, list)
        self.search_button.configure(text=UI_TEXT["button_search"], state="normal")
        self.heat_search_button.configure(text=UI_TEXT["button_heat_search"], state="normal")
        if heat_search:
            self.status_var.set(UI_TEXT["status_heat_search_complete"] if hits else UI_TEXT["status_no_heat_results"])
            summary_text = UI_TEXT["summary_heat_search"]
        else:
            self.status_var.set(UI_TEXT["status_search_complete"] if hits else UI_TEXT["status_no_search_results"])
            summary_text = UI_TEXT["summary_search"]
        self.summary_var.set(
            summary_text.format(
                results=len(hits),
                skipped=skipped,
            )
        )
        self._render_search_hits(hits)
        self._start_orbit_refresh(query=query)

    def _handle_search_error(self, payload: object) -> None:
        _heat_search, error = payload
        self.search_button.configure(text=UI_TEXT["button_search"], state="normal")
        self.heat_search_button.configure(text=UI_TEXT["button_heat_search"], state="normal")
        self.status_var.set(UI_TEXT["status_error"])
        self.summary_var.set(str(error))
        self._render_loading_card(UI_TEXT["status_error"])

    def _start_scan(self) -> None:
        if self.scanning:
            return
        if not self.memory_folder:
            self.status_var.set(UI_TEXT["memory_folder_missing"])
            self._render_memory_state()
            return

        self.scanning = True
        self.status_var.set(UI_TEXT["status_scanning"])
        self.summary_var.set(UI_TEXT["status_scanning"])
        self.scan_button.configure(text=UI_TEXT["button_scanning"], state="disabled")
        self.open_output_button.configure(state="disabled")
        self._render_empty_cards()

        self.scan_thread = threading.Thread(target=self._scan_worker, args=(self.memory_folder,), daemon=True)
        self.scan_thread.start()

    def _scan_worker(self, memory_folder: Path) -> None:
        try:
            documents, skipped = scan_memory(memory_folder)
            result = analyze_documents(documents, memory_folder, skipped_files=skipped)
            output_path = write_suggestion(memory_folder, result)
            self.events.put(("scan_done", (result, output_path)))
        except Exception as exc:  # noqa: BLE001
            self.events.put(("scan_error", exc))

    def _poll_events(self) -> None:
        try:
            while True:
                event_name, payload = self.events.get_nowait()
                if event_name == "scan_done":
                    result, output_path = payload
                    assert isinstance(result, AnalysisResult)
                    assert isinstance(output_path, Path)
                    self._handle_scan_done(result, output_path)
                elif event_name == "scan_error":
                    self._handle_scan_error(payload)
                elif event_name == "search_done":
                    self._handle_search_done(payload)
                elif event_name == "search_error":
                    self._handle_search_error(payload)
                elif event_name == "orbit_done":
                    self._handle_orbit_done(payload)
                elif event_name == "orbit_error":
                    self._handle_orbit_error(payload)
                elif event_name == "source_preview_done":
                    self._handle_source_preview_done(payload)
                elif event_name == "source_preview_error":
                    self._handle_source_preview_error(payload)
                elif event_name == "self_check_done":
                    self._handle_self_check_done(payload)
        except queue.Empty:
            pass
        self.after(100, self._poll_events)

    def _handle_scan_done(self, result: AnalysisResult, output_path: Path) -> None:
        self.scanning = False
        self.output_path = output_path
        self.scan_button.configure(text=UI_TEXT["button_scan"], state="normal")
        self.open_output_button.configure(state="normal")
        self.status_var.set(UI_TEXT["status_complete"] if result.traces else UI_TEXT["status_no_trace"])
        self.summary_var.set(
            UI_TEXT["summary_scan"].format(
                files=result.scanned_files,
                skipped=result.skipped_files,
                traces=len(result.traces),
            )
        )
        self._render_cards(result)

    def _handle_scan_error(self, payload: object) -> None:
        self.scanning = False
        self.scan_button.configure(text=UI_TEXT["button_scan"], state="normal")
        self.status_var.set(UI_TEXT["status_error"])
        self.summary_var.set(str(payload))
        messagebox.showerror(UI_TEXT["dialog_title"], f"{UI_TEXT['status_error']}\n{payload}")

    def _render_empty_cards(self) -> None:
        self._clear_cards()
        card = self._card_frame()
        card.pack(fill="x", pady=(0, 10))
        tk.Label(
            card,
            text=UI_TEXT["card_empty"],
            fg=COLORS["muted"],
            bg=COLORS["panel"],
            font=FONT_JP_SMALL,
        ).pack(anchor="w")

    def _render_loading_card(self, message: str) -> None:
        self._clear_cards()
        card = self._card_frame()
        card.pack(fill="x", pady=(0, 10))
        tk.Label(
            card,
            text=message,
            fg=COLORS["muted"],
            bg=COLORS["panel"],
            font=FONT_JP_SMALL,
        ).pack(anchor="w")

    def _render_search_hits(self, hits: list[SearchHit]) -> None:
        self._clear_cards()
        if not hits:
            self._render_loading_card(UI_TEXT["status_no_search_results"])
            return
        for hit in hits[:8]:
            card = self._card_frame()
            card.pack(fill="x", pady=(0, 10))
            card.bind("<Button-1>", lambda _event, selected=hit: self._show_source_path(selected.path, selected.title))
            tk.Label(
                card,
                text=hit.title,
                fg=COLORS["heat"],
                bg=COLORS["panel"],
                font=FONT_JP,
                wraplength=230,
                justify="left",
            ).pack(anchor="w")
            tk.Label(
                card,
                text=UI_TEXT["search_result_meta"].format(path=hit.relative_path, score=hit.score),
                fg=COLORS["muted"],
                bg=COLORS["panel"],
                font=FONT_MONO,
                wraplength=230,
                justify="left",
            ).pack(anchor="w", pady=(5, 0))
            tk.Label(
                card,
                text=hit.excerpt,
                fg=COLORS["muted"],
                bg=COLORS["panel"],
                font=FONT_JP_SMALL,
                wraplength=230,
                justify="left",
            ).pack(anchor="w", pady=(5, 0))
            tk.Button(
                card,
                text=UI_TEXT["button_preview_source"],
                command=lambda selected=hit: self._show_source_path(selected.path, selected.title),
                **self._small_button_style(COLORS["panel_light"]),
            ).pack(anchor="w", pady=(8, 0))

    def _render_cards(self, result: AnalysisResult) -> None:
        self._clear_cards()
        if not result.fragments:
            card = self._card_frame()
            card.pack(fill="x", pady=(0, 10))
            tk.Label(
                card,
                text=UI_TEXT["status_no_trace"],
                fg=COLORS["text"],
                bg=COLORS["panel"],
                font=FONT_JP,
            ).pack(anchor="w")
            tk.Label(
                card,
                text=result.suggestion,
                fg=COLORS["muted"],
                bg=COLORS["panel"],
                font=FONT_JP_SMALL,
                wraplength=230,
                justify="left",
            ).pack(anchor="w", pady=(8, 0))
            return

        for fragment in result.fragments[:5]:
            card = self._card_frame()
            card.pack(fill="x", pady=(0, 10))
            card.bind("<Button-1>", lambda _event, selected=fragment: self._show_source_path(selected.path, selected.title))
            top = tk.Frame(card, bg=COLORS["panel"])
            top.pack(fill="x")
            tk.Label(
                top,
                text=fragment.heat_word,
                fg=COLORS["heat"],
                bg=COLORS["panel"],
                font=FONT_JP,
            ).pack(side="left")
            tk.Label(
                top,
                text=UI_TEXT["card_score"].format(score=fragment.score),
                fg=COLORS["muted"],
                bg=COLORS["panel"],
                font=FONT_MONO,
            ).pack(side="right")
            tk.Label(
                card,
                text=f"{UI_TEXT['card_file']}: {Path(fragment.relative_path).name}",
                fg=COLORS["text"],
                bg=COLORS["panel"],
                font=FONT_JP_SMALL,
                wraplength=230,
                justify="left",
            ).pack(anchor="w", pady=(6, 0))
            tk.Label(
                card,
                text=f"{UI_TEXT['card_excerpt']}: {fragment.excerpt}",
                fg=COLORS["muted"],
                bg=COLORS["panel"],
                font=FONT_JP_SMALL,
                wraplength=230,
                justify="left",
            ).pack(anchor="w", pady=(4, 0))
            tk.Label(
                card,
                text=f"{UI_TEXT['section_suggestion']}: {result.suggestion.splitlines()[0]}",
                fg=COLORS["glow"],
                bg=COLORS["panel"],
                font=FONT_JP_SMALL,
                wraplength=230,
                justify="left",
            ).pack(anchor="w", pady=(6, 0))
            tk.Button(
                card,
                text=UI_TEXT["button_preview_source"],
                command=lambda selected=fragment: self._show_source_path(selected.path, selected.title),
                **self._small_button_style(COLORS["panel_light"]),
            ).pack(anchor="w", pady=(8, 0))

    def _card_frame(self) -> tk.Frame:
        return tk.Frame(
            self.cards_frame,
            bg=COLORS["panel"],
            highlightbackground=COLORS["line"],
            highlightthickness=1,
            padx=14,
            pady=12,
        )

    def _clear_cards(self) -> None:
        for child in self.cards_frame.winfo_children():
            child.destroy()

    def _show_source_path(self, source_path: str | Path, title: str = "") -> None:
        path = self._resolve_source_path(source_path)
        self.source_preview_request_id += 1
        request_id = self.source_preview_request_id
        self.source_preview_loaded = False
        self.source_preview_status_var.set(UI_TEXT["source_preview_loading"])
        self._set_source_preview_text(UI_TEXT["source_preview_loading"])
        self.source_preview_thread = threading.Thread(
            target=self._source_preview_worker,
            args=(request_id, path, title),
            daemon=True,
        )
        self.source_preview_thread.start()

    def _resolve_source_path(self, source_path: str | Path) -> Path:
        return _resolve_related_source_path(source_path, self.memory_folder)

    def _source_preview_worker(self, request_id: int, path: Path, title: str) -> None:
        try:
            if not path.exists():
                self.events.put(("source_preview_error", (request_id, path, UI_TEXT["source_preview_missing"])))
                return
            if not path.is_file() or path.suffix.lower() not in {".md", ".txt"}:
                self.events.put(("source_preview_error", (request_id, path, UI_TEXT["source_preview_not_file"])))
                return
            text, truncated = read_source_preview_text(path)
            try:
                record_recent_return(path, title)
            except OSError:
                pass
            self.events.put(("source_preview_done", SourcePreviewResult(request_id, path, title, text, truncated)))
        except Exception:  # noqa: BLE001
            self.events.put(("source_preview_error", (request_id, path, UI_TEXT["source_preview_failed"])))

    def _handle_source_preview_done(self, payload: object) -> None:
        assert isinstance(payload, SourcePreviewResult)
        if payload.request_id != self.source_preview_request_id:
            return
        header = UI_TEXT["source_preview_path"].format(path=payload.path)
        title = payload.title.strip()
        content = payload.text
        if payload.truncated:
            content = f"{content}{UI_TEXT['source_preview_truncated']}"
        if title:
            content = f"{title}\n{header}\n\n{content}"
        else:
            content = f"{header}\n\n{content}"
        self.source_preview_loaded = True
        self.source_preview_status_var.set(UI_TEXT["source_preview_loaded"])
        self._set_source_preview_text(content)

    def _handle_source_preview_error(self, payload: object) -> None:
        request_id, path, message = payload
        assert isinstance(request_id, int)
        assert isinstance(path, Path)
        assert isinstance(message, str)
        if request_id != self.source_preview_request_id:
            return
        self.source_preview_loaded = False
        self.source_preview_status_var.set(message)
        self._set_source_preview_text(f"{message}\n{UI_TEXT['source_preview_path'].format(path=path)}")

    def _set_source_preview_text(self, text: str) -> None:
        if not hasattr(self, "source_preview_box"):
            return
        self.source_preview_box.configure(state="normal")
        self.source_preview_box.delete("1.0", "end")
        self.source_preview_box.insert("1.0", text)
        self.source_preview_box.configure(state="disabled")

    def _choose_memory_folder(self) -> None:
        selected = filedialog.askdirectory(title=UI_TEXT["dialog_choose_memory"])
        if not selected:
            return
        folder = existing_folder(selected)
        if not folder:
            self.status_var.set(UI_TEXT["memory_folder_missing"])
            return
        self.memory_folder = folder
        self.config_data.memory_folder = str(folder)
        self.config_store.save(self.config_data)
        self.status_var.set(UI_TEXT["status_idle"])
        self._render_memory_state()

    def _open_output(self) -> None:
        if self.output_path:
            open_path(self.output_path.parent)
            return
        if self.memory_folder:
            open_path(self.memory_folder / "OIKAWA" / "suggestions")

    def _finish_launch_check(self) -> None:
        print(UI_TEXT["launch_check_ok"])
        self.destroy()

    def _finish_gui_smoke(self) -> None:
        print(UI_TEXT["gui_smoke_ok"])
        self.destroy()


def run_gui(launch_check: bool = False, gui_smoke_seconds: float = 0.0, memory_folder_override: str = "") -> int:
    app = OikawaApp(
        launch_check=launch_check,
        gui_smoke_seconds=gui_smoke_seconds,
        memory_folder_override=memory_folder_override,
    )
    app.mainloop()
    return 0


def run_scan_check(memory_folder: str) -> int:
    root = existing_folder(memory_folder)
    if not root:
        raise RuntimeError(UI_TEXT["memory_folder_missing"])
    documents, skipped = scan_memory(root)
    result = analyze_documents(documents, root, skipped_files=skipped)
    output_path = write_suggestion(root, result)
    print(UI_TEXT["scan_check_ok"])
    print(output_path)
    return 0


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--launch-check", action="store_true")
    parser.add_argument("--gui-smoke-seconds", type=float, default=0.0)
    parser.add_argument("--memory-folder", default="")
    parser.add_argument("--scan-check", default="")
    args = parser.parse_args()

    if args.scan_check:
        return run_scan_check(args.scan_check)

    return run_gui(
        launch_check=args.launch_check,
        gui_smoke_seconds=args.gui_smoke_seconds,
        memory_folder_override=args.memory_folder,
    )


if __name__ == "__main__":
    raise SystemExit(main())
