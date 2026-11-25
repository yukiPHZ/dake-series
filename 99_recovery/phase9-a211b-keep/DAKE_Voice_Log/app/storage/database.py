# -*- coding: utf-8 -*-
from __future__ import annotations

import sqlite3
import threading
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Iterable

from app.core.app_state import TodayStats


TRANSCRIPTION_RECORDED = "RECORDED"
TRANSCRIPTION_QUEUED = "QUEUED"
TRANSCRIPTION_PROCESSING = "TRANSCRIBING"
TRANSCRIPTION_DONE = "TRANSCRIBED"
TRANSCRIPTION_NO_SPEECH = "NO_SPEECH"
TRANSCRIPTION_FAILED = "FAILED"
DAILY_WAITING = "WAITING"
DAILY_PROCESSING = "PROCESSING"
DAILY_COMPLETE = "COMPLETE"
DAILY_FAILED = "FAILED"


@dataclass(frozen=True)
class ClipRecord:
    id: int
    session_id: str
    record_date: str
    started_at: str
    ended_at: str
    duration_sec: float
    audio_path: str
    transcript_path: str
    transcription_status: str
    transcription_retries: int
    last_error: str


class VoiceLogDatabase:
    def __init__(self, path: Path) -> None:
        self.path = path
        self._lock = threading.RLock()
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.initialize()

    def connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.path)
        connection.row_factory = sqlite3.Row
        return connection

    @contextmanager
    def connection(self):
        connection = self.connect()
        try:
            yield connection
            connection.commit()
        except Exception:
            connection.rollback()
            raise
        finally:
            connection.close()

    def initialize(self) -> None:
        with self._lock, self.connection() as conn:
            conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS clips (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    session_id TEXT NOT NULL,
                    record_date TEXT NOT NULL,
                    started_at TEXT NOT NULL,
                    ended_at TEXT NOT NULL,
                    duration_sec REAL NOT NULL DEFAULT 0,
                    audio_path TEXT NOT NULL,
                    transcript_path TEXT NOT NULL DEFAULT '',
                    transcription_status TEXT NOT NULL DEFAULT 'RECORDED',
                    transcription_retries INTEGER NOT NULL DEFAULT 0,
                    next_retry_at TEXT NOT NULL DEFAULT '',
                    last_error TEXT NOT NULL DEFAULT '',
                    speaker_status TEXT NOT NULL DEFAULT 'unused',
                    speaker_label TEXT,
                    speaker_profile_id TEXT,
                    speaker_segments_path TEXT,
                    classification_tag TEXT,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );

                CREATE INDEX IF NOT EXISTS idx_clips_record_date_started
                    ON clips(record_date, started_at);
                CREATE INDEX IF NOT EXISTS idx_clips_transcription_status
                    ON clips(transcription_status, next_retry_at);

                CREATE TABLE IF NOT EXISTS daily_jobs (
                    record_date TEXT PRIMARY KEY,
                    daily_transcript_status TEXT NOT NULL DEFAULT 'WAITING',
                    daily_transcript_path TEXT NOT NULL DEFAULT '',
                    summary_status TEXT NOT NULL DEFAULT 'WAITING',
                    summary_path TEXT NOT NULL DEFAULT '',
                    summary_model TEXT NOT NULL DEFAULT '',
                    last_error TEXT NOT NULL DEFAULT '',
                    updated_at TEXT NOT NULL
                );
                """
            )

    def create_clip(
        self,
        session_id: str,
        started_at: datetime,
        ended_at: datetime,
        audio_path: Path,
        transcript_path: Path,
    ) -> int:
        now = datetime.now().isoformat(timespec="seconds")
        duration = max(0.0, (ended_at - started_at).total_seconds())
        record_date = started_at.date().isoformat()
        with self._lock, self.connection() as conn:
            cursor = conn.execute(
                """
                INSERT INTO clips (
                    session_id, record_date, started_at, ended_at, duration_sec,
                    audio_path, transcript_path, transcription_status,
                    created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    session_id,
                    record_date,
                    started_at.isoformat(timespec="seconds"),
                    ended_at.isoformat(timespec="seconds"),
                    duration,
                    str(audio_path),
                    str(transcript_path),
                    TRANSCRIPTION_RECORDED,
                    now,
                    now,
                ),
            )
            self.ensure_daily_job(record_date, conn)
            return int(cursor.lastrowid)

    def ensure_daily_job(self, record_date: str, conn: sqlite3.Connection | None = None) -> None:
        now = datetime.now().isoformat(timespec="seconds")
        if conn is None:
            with self._lock, self.connection() as own_conn:
                self.ensure_daily_job(record_date, own_conn)
            return
        conn.execute(
            """
            INSERT OR IGNORE INTO daily_jobs (record_date, updated_at)
            VALUES (?, ?)
            """,
            (record_date, now),
        )

    def update_clip_status(
        self,
        clip_id: int,
        status: str,
        transcript_path: Path | None = None,
        error: str = "",
        retries: int | None = None,
        next_retry_at: str = "",
    ) -> None:
        fields = ["transcription_status = ?", "last_error = ?", "updated_at = ?"]
        params: list[object] = [status, error, datetime.now().isoformat(timespec="seconds")]
        if transcript_path is not None:
            fields.append("transcript_path = ?")
            params.append(str(transcript_path))
        if retries is not None:
            fields.append("transcription_retries = ?")
            params.append(retries)
        fields.append("next_retry_at = ?")
        params.append(next_retry_at)
        params.append(clip_id)
        with self._lock, self.connection() as conn:
            conn.execute(f"UPDATE clips SET {', '.join(fields)} WHERE id = ?", params)

    def increment_retry(self, clip_id: int, error: str, next_retry_at: str, max_retries: int) -> None:
        with self._lock, self.connection() as conn:
            row = conn.execute("SELECT transcription_retries FROM clips WHERE id = ?", (clip_id,)).fetchone()
            retries = int(row["transcription_retries"]) + 1 if row else 1
            status = TRANSCRIPTION_QUEUED if retries < max_retries else TRANSCRIPTION_FAILED
            conn.execute(
                """
                UPDATE clips
                SET transcription_status = ?, transcription_retries = ?, last_error = ?,
                    next_retry_at = ?, updated_at = ?
                WHERE id = ?
                """,
                (
                    status,
                    retries,
                    error,
                    next_retry_at if status == TRANSCRIPTION_QUEUED else "",
                    datetime.now().isoformat(timespec="seconds"),
                    clip_id,
                ),
            )

    def queue_recorded_clips(self) -> int:
        with self._lock, self.connection() as conn:
            cursor = conn.execute(
                """
                UPDATE clips
                SET transcription_status = ?, updated_at = ?
                WHERE transcription_status = ?
                """,
                (TRANSCRIPTION_QUEUED, datetime.now().isoformat(timespec="seconds"), TRANSCRIPTION_RECORDED),
            )
            return int(cursor.rowcount)

    def next_transcription_clip(self) -> ClipRecord | None:
        now = datetime.now().isoformat(timespec="seconds")
        with self._lock, self.connection() as conn:
            row = conn.execute(
                """
                SELECT * FROM clips
                WHERE transcription_status IN (?, ?)
                  AND (next_retry_at = '' OR next_retry_at <= ?)
                ORDER BY started_at ASC
                LIMIT 1
                """,
                (TRANSCRIPTION_RECORDED, TRANSCRIPTION_QUEUED, now),
            ).fetchone()
            if row is None:
                return None
            return self._clip_from_row(row)

    def backlog_count(self) -> int:
        with self._lock, self.connection() as conn:
            row = conn.execute(
                """
                SELECT COUNT(*) AS value FROM clips
                WHERE transcription_status IN (?, ?, ?)
                """,
                (TRANSCRIPTION_RECORDED, TRANSCRIPTION_QUEUED, TRANSCRIPTION_PROCESSING),
            ).fetchone()
            return int(row["value"] if row else 0)

    def today_stats(self, today: date) -> TodayStats:
        record_date = today.isoformat()
        with self._lock, self.connection() as conn:
            row = conn.execute(
                """
                SELECT
                  COUNT(*) AS recorded,
                  SUM(CASE WHEN transcription_status = ? THEN 1 ELSE 0 END) AS transcribed,
                  SUM(CASE WHEN transcription_status IN (?, ?, ?) THEN 1 ELSE 0 END) AS pending
                FROM clips
                WHERE record_date = ?
                """,
                (
                    TRANSCRIPTION_DONE,
                    TRANSCRIPTION_RECORDED,
                    TRANSCRIPTION_QUEUED,
                    TRANSCRIPTION_PROCESSING,
                    record_date,
                ),
            ).fetchone()
            daily = conn.execute("SELECT summary_status FROM daily_jobs WHERE record_date = ?", (record_date,)).fetchone()
            return TodayStats(
                recorded=int(row["recorded"] or 0),
                transcribed=int(row["transcribed"] or 0),
                pending=int(row["pending"] or 0),
                summary_status=str(daily["summary_status"]) if daily else DAILY_WAITING,
            )

    def clips_for_date(self, record_date: date | str, include_no_speech: bool = True) -> list[ClipRecord]:
        day = record_date if isinstance(record_date, str) else record_date.isoformat()
        statuses: tuple[str, ...]
        if include_no_speech:
            statuses = (
                TRANSCRIPTION_DONE,
                TRANSCRIPTION_NO_SPEECH,
                TRANSCRIPTION_FAILED,
                TRANSCRIPTION_RECORDED,
                TRANSCRIPTION_QUEUED,
                TRANSCRIPTION_PROCESSING,
            )
        else:
            statuses = (TRANSCRIPTION_DONE,)
        placeholders = ",".join("?" for _ in statuses)
        with self._lock, self.connection() as conn:
            rows = conn.execute(
                f"""
                SELECT * FROM clips
                WHERE record_date = ? AND transcription_status IN ({placeholders})
                ORDER BY started_at ASC
                """,
                (day, *statuses),
            ).fetchall()
            return [self._clip_from_row(row) for row in rows]

    def recent_clips(self, limit: int = 5) -> list[ClipRecord]:
        with self._lock, self.connection() as conn:
            rows = conn.execute(
                "SELECT * FROM clips ORDER BY started_at DESC LIMIT ?",
                (limit,),
            ).fetchall()
            return [self._clip_from_row(row) for row in rows]

    def dates_needing_daily_transcript(self) -> list[str]:
        with self._lock, self.connection() as conn:
            rows = conn.execute(
                """
                SELECT DISTINCT c.record_date
                FROM clips c
                LEFT JOIN daily_jobs d ON d.record_date = c.record_date
                WHERE c.transcription_status = ?
                  AND (d.daily_transcript_status IS NULL OR d.daily_transcript_status != ?)
                ORDER BY c.record_date ASC
                """,
                (TRANSCRIPTION_DONE, DAILY_COMPLETE),
            ).fetchall()
            return [str(row["record_date"]) for row in rows]

    def summary_pending_dates(self, before_date: date, model: str) -> list[str]:
        with self._lock, self.connection() as conn:
            rows = conn.execute(
                """
                SELECT d.record_date
                FROM daily_jobs d
                WHERE d.record_date < ?
                  AND d.daily_transcript_status = ?
                  AND d.summary_status != ?
                ORDER BY d.record_date ASC
                """,
                (before_date.isoformat(), DAILY_COMPLETE, DAILY_COMPLETE),
            ).fetchall()
        return [str(row["record_date"]) for row in rows]

    def update_daily_transcript(self, record_date: str, path: Path, status: str, error: str = "") -> None:
        self.ensure_daily_job(record_date)
        with self._lock, self.connection() as conn:
            conn.execute(
                """
                UPDATE daily_jobs
                SET daily_transcript_status = ?, daily_transcript_path = ?,
                    last_error = ?, updated_at = ?
                WHERE record_date = ?
                """,
                (status, str(path), error, datetime.now().isoformat(timespec="seconds"), record_date),
            )

    def update_summary(self, record_date: str, path: Path, status: str, model: str, error: str = "") -> None:
        self.ensure_daily_job(record_date)
        with self._lock, self.connection() as conn:
            conn.execute(
                """
                UPDATE daily_jobs
                SET summary_status = ?, summary_path = ?, summary_model = ?,
                    last_error = ?, updated_at = ?
                WHERE record_date = ?
                """,
                (status, str(path), model, error, datetime.now().isoformat(timespec="seconds"), record_date),
            )

    def reset_processing_jobs(self) -> None:
        with self._lock, self.connection() as conn:
            conn.execute(
                """
                UPDATE clips
                SET transcription_status = ?, updated_at = ?
                WHERE transcription_status = ?
                """,
                (TRANSCRIPTION_QUEUED, datetime.now().isoformat(timespec="seconds"), TRANSCRIPTION_PROCESSING),
            )
            conn.execute(
                """
                UPDATE daily_jobs
                SET summary_status = ?, updated_at = ?
                WHERE summary_status = ?
                """,
                (DAILY_WAITING, datetime.now().isoformat(timespec="seconds"), DAILY_PROCESSING),
            )

    def iter_clip_paths(self) -> Iterable[Path]:
        with self._lock, self.connection() as conn:
            rows = conn.execute("SELECT audio_path FROM clips").fetchall()
        for row in rows:
            yield Path(str(row["audio_path"]))

    @staticmethod
    def _clip_from_row(row: sqlite3.Row) -> ClipRecord:
        return ClipRecord(
            id=int(row["id"]),
            session_id=str(row["session_id"]),
            record_date=str(row["record_date"]),
            started_at=str(row["started_at"]),
            ended_at=str(row["ended_at"]),
            duration_sec=float(row["duration_sec"]),
            audio_path=str(row["audio_path"]),
            transcript_path=str(row["transcript_path"]),
            transcription_status=str(row["transcription_status"]),
            transcription_retries=int(row["transcription_retries"]),
            last_error=str(row["last_error"]),
        )
