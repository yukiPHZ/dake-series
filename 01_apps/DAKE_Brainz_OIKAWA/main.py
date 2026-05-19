# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import math
import queue
import random
import threading
import time
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk

from core.config import AppConfig, ConfigStore, existing_folder, open_path, resolve_memory_folder
from core.heat_engine import AnalysisResult, analyze_documents
from core.markdown_writer import write_suggestion
from core.qpsc_notifications import (
    QpscNotification,
    mark_qpsc_notification_read,
    read_qpsc_notification_events,
    read_qpsc_notifications,
)
from core.qpsc_status import is_brainz_awake, read_brainz_status
from core.scanner import MemoryDocument, scan_memory


APP_NAME = "DakeBrainzOIKAWA"
WINDOW_TITLE = "OIKAWA"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "app_title": "OIKAWA",
    "copyright": COPYRIGHT,
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
    "qpsc_notification_open_related": "原本",
    "qpsc_notification_read": "既読",
    "qpsc_related_missing": "原本が見つかりません。",
    "footer_source": "local scan / no cloud",
    "launch_check_ok": "LAUNCH CHECK OK",
    "gui_smoke_ok": "GUI SMOKE OK",
    "scan_check_ok": "SCAN CHECK OK",
    "ghost_words": ["熾火", "巡り", "側に", "在る", "余白", "記憶", "痕跡"],
}

MAX_PREVIEW_CHARS = 240_000


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
                title=notification.title.strip() or UI_TEXT["heat_candidate_title"],
                message=notification.message.strip() or UI_TEXT["heat_candidate_message"].format(
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
        excerpt_source = preview_text or notification.message or notification.title
        hits.append(
            SearchHit(
                title=notification.title.strip() or path.stem,
                path=path,
                relative_path=_relative_path(path, memory_root),
                excerpt=_search_excerpt(excerpt_source, terms),
                score=score,
            )
        )
        seen_paths.add(resolved_path)
    hits.sort(key=lambda item: (item.score, item.path.stat().st_mtime if item.path.exists() else 0), reverse=True)
    return hits[:limit]


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
            notification.title,
            notification.message,
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
    "background": "#06070A",
    "panel": "#101218",
    "panel_light": "#151821",
    "text": "#D7DAE0",
    "muted": "#7B8190",
    "glow": "#6E7FA8",
    "heat": "#C47A3A",
    "line": "#242B3A",
    "line_soft": "#161B26",
}

FONT_JP = ("BIZ UDPGothic", 11)
FONT_JP_SMALL = ("BIZ UDPGothic", 9)
FONT_TITLE = ("Segoe UI", 18, "bold")
FONT_LABEL = ("Segoe UI", 9)
FONT_MONO = ("JetBrains Mono", 9)


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
        self.geometry("1120x720")
        self.minsize(980, 620)
        self.configure(bg=COLORS["background"])

        self.config_store = ConfigStore()
        self.config_data = self.config_store.load()
        self.memory_folder = self._resolve_initial_memory_folder(memory_folder_override)
        self.output_path: Path | None = None
        self.events: queue.Queue[tuple[str, object]] = queue.Queue()
        self.scan_thread: threading.Thread | None = None
        self.search_thread: threading.Thread | None = None
        self.source_preview_thread: threading.Thread | None = None
        self.source_preview_request_id = 0
        self.qpsc_notifications_all: list[QpscNotification] = []
        self.heat_candidates: list[HeatCandidate] = []
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

        header = tk.Frame(self, bg=COLORS["background"])
        header.place(x=30, y=24)
        tk.Label(
            header,
            text=UI_TEXT["app_title"],
            fg=COLORS["text"],
            bg=COLORS["background"],
            font=FONT_TITLE,
        ).pack(anchor="w")
        tk.Label(
            header,
            text=UI_TEXT["app_subtitle"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
            font=FONT_JP_SMALL,
        ).pack(anchor="w", pady=(4, 12))
        tk.Label(
            header,
            textvariable=self.status_var,
            fg=COLORS["heat"],
            bg=COLORS["background"],
            font=FONT_JP_SMALL,
        ).pack(anchor="w")
        tk.Label(
            header,
            textvariable=self.memory_var,
            fg=COLORS["muted"],
            bg=COLORS["background"],
            font=FONT_MONO,
            wraplength=520,
            justify="left",
        ).pack(anchor="w", pady=(8, 0))
        search_row = tk.Frame(header, bg=COLORS["background"])
        search_row.pack(anchor="w", pady=(10, 0))
        self.search_entry = tk.Entry(
            search_row,
            width=34,
            bg=COLORS["panel"],
            fg=COLORS["text"],
            insertbackground=COLORS["text"],
            relief="flat",
            bd=0,
            font=FONT_JP_SMALL,
        )
        self.search_entry.pack(side="left", ipady=7)
        self.search_entry.bind("<Return>", lambda _event: self._start_memory_search())
        self.search_button = tk.Button(
            search_row,
            text=UI_TEXT["button_search"],
            command=self._start_memory_search,
            **self._button_style(COLORS["panel_light"]),
        )
        self.search_button.pack(side="left", padx=(8, 0))
        self.heat_search_button = tk.Button(
            search_row,
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

        self.results_frame = tk.Frame(self, bg=COLORS["background"])
        self.results_frame.place(relx=0.04, rely=0.50, relwidth=0.58, relheight=0.44)
        tk.Label(
            self.results_frame,
            text=UI_TEXT["section_related"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
            font=FONT_JP_SMALL,
        ).pack(anchor="w")
        self.cards_frame = tk.Frame(self.results_frame, bg=COLORS["background"])
        self.cards_frame.pack(fill="both", expand=True, pady=(10, 0))

        actions = tk.Frame(self, bg=COLORS["background"])
        actions.place(relx=0.97, rely=0.94, anchor="se")
        tk.Label(
            actions,
            text=UI_TEXT["footer_source"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
            font=FONT_LABEL,
        ).pack(anchor="e", pady=(0, 10))
        self.open_output_button = tk.Button(
            actions,
            text=UI_TEXT["button_open_output"],
            command=self._open_output,
            state="disabled",
            **self._button_style(COLORS["panel_light"]),
        )
        self.open_output_button.pack(anchor="e", pady=(0, 10))
        self.scan_button = tk.Button(
            actions,
            text=UI_TEXT["button_scan"],
            command=self._start_scan,
            **self._button_style(COLORS["heat"]),
        )
        self.scan_button.pack(anchor="e")

        notice = tk.Frame(
            self,
            bg=COLORS["background"],
            highlightbackground=COLORS["line_soft"],
            highlightthickness=1,
            padx=14,
            pady=10,
        )
        notice.place(relx=0.97, rely=0.08, anchor="ne")
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
            wraplength=260,
        )
        self.qpsc_status_label.pack(anchor="w", pady=(6, 0))
        self.qpsc_import_frame = tk.Frame(notice, bg=COLORS["background"])
        self.qpsc_import_frame.pack(fill="x", pady=(8, 0))

        heat = tk.Frame(
            self,
            bg=COLORS["background"],
            highlightbackground=COLORS["line_soft"],
            highlightthickness=1,
            padx=12,
            pady=8,
        )
        heat.place(relx=0.66, rely=0.34, relwidth=0.31, relheight=0.13)
        tk.Label(
            heat,
            text=UI_TEXT["section_heat_candidates"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
            font=FONT_LABEL,
        ).pack(anchor="w")
        self.heat_candidates_frame = tk.Frame(heat, bg=COLORS["background"])
        self.heat_candidates_frame.pack(fill="both", expand=True, pady=(5, 0))

        self.source_preview_status_var = tk.StringVar(value=UI_TEXT["source_preview_empty"])
        source_preview = tk.Frame(
            self,
            bg=COLORS["panel"],
            highlightbackground=COLORS["line"],
            highlightthickness=1,
            padx=12,
            pady=10,
        )
        source_preview.place(relx=0.66, rely=0.50, relwidth=0.31, relheight=0.35)
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
            wraplength=300,
            justify="left",
        ).pack(anchor="w", pady=(5, 8))
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
            font=FONT_JP_SMALL,
            yscrollcommand=preview_scrollbar.set,
        )
        self.source_preview_box.pack(side="left", fill="both", expand=True)
        preview_scrollbar.configure(command=self.source_preview_box.yview)
        self._set_source_preview_text(UI_TEXT["source_preview_empty"])

        footer = tk.Frame(self, bg=COLORS["background"])
        footer.place(relx=0.04, rely=0.96, anchor="sw")
        tk.Label(
            footer,
            text=UI_TEXT["copyright"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
            font=FONT_LABEL,
        ).pack(anchor="w")
        tk.Label(
            footer,
            textvariable=self.summary_var,
            fg=COLORS["muted"],
            bg=COLORS["background"],
            font=FONT_MONO,
        ).pack(anchor="w", pady=(4, 0))

    def _refresh_qpsc_notifications(self, schedule: bool = True) -> None:
        status = read_brainz_status()
        if is_brainz_awake(status):
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
        self._render_qpsc_import_notifications(read_qpsc_notifications(limit=3))
        self._refresh_heat_candidates()
        if schedule:
            self.after(10000, self._refresh_qpsc_notifications)

    def _refresh_heat_candidates(self, query: str = "") -> None:
        self.heat_candidates = build_heat_candidates(self.qpsc_notifications_all, query=query, limit=3)
        self._render_heat_candidates(self.heat_candidates)

    def _render_qpsc_import_notifications(self, notifications: list[QpscNotification]) -> None:
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
            return

        for notification in notifications:
            event_frame = tk.Frame(self.qpsc_import_frame, bg=COLORS["background"])
            event_frame.pack(fill="x", pady=(0, 8))
            tk.Label(
                event_frame,
                text=notification.title,
                fg=COLORS["text"],
                bg=COLORS["background"],
                font=FONT_JP_SMALL,
                wraplength=250,
                justify="left",
            ).pack(anchor="w")
            if notification.message:
                tk.Label(
                    event_frame,
                    text=notification.message,
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
        self.missing_frame.place(relx=0.04, rely=0.30, relwidth=0.28)

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
            return

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
                elif event_name == "source_preview_done":
                    self._handle_source_preview_done(payload)
                elif event_name == "source_preview_error":
                    self._handle_source_preview_error(payload)
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
                wraplength=560,
                justify="left",
            ).pack(anchor="w")
            tk.Label(
                card,
                text=UI_TEXT["search_result_meta"].format(path=hit.relative_path, score=hit.score),
                fg=COLORS["muted"],
                bg=COLORS["panel"],
                font=FONT_MONO,
                wraplength=560,
                justify="left",
            ).pack(anchor="w", pady=(5, 0))
            tk.Label(
                card,
                text=hit.excerpt,
                fg=COLORS["muted"],
                bg=COLORS["panel"],
                font=FONT_JP_SMALL,
                wraplength=560,
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
                wraplength=560,
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
                wraplength=560,
                justify="left",
            ).pack(anchor="w", pady=(6, 0))
            tk.Label(
                card,
                text=f"{UI_TEXT['card_excerpt']}: {fragment.excerpt}",
                fg=COLORS["muted"],
                bg=COLORS["panel"],
                font=FONT_JP_SMALL,
                wraplength=560,
                justify="left",
            ).pack(anchor="w", pady=(4, 0))
            tk.Label(
                card,
                text=f"{UI_TEXT['section_suggestion']}: {result.suggestion.splitlines()[0]}",
                fg=COLORS["glow"],
                bg=COLORS["panel"],
                font=FONT_JP_SMALL,
                wraplength=560,
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
        self.source_preview_status_var.set(UI_TEXT["source_preview_loaded"])
        self._set_source_preview_text(content)

    def _handle_source_preview_error(self, payload: object) -> None:
        request_id, path, message = payload
        assert isinstance(request_id, int)
        assert isinstance(path, Path)
        assert isinstance(message, str)
        if request_id != self.source_preview_request_id:
            return
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
