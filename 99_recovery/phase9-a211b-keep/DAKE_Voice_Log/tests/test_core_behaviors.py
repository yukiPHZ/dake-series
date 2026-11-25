# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
import tempfile
import unittest
from datetime import datetime, timedelta
from pathlib import Path

import numpy as np

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app.audio.recorder_service import RotationPolicy
from app.audio.vad_service import EnergyVad, VadConfig
from app.core.job_queue import RetryPolicy
from app.storage.daily_builder import build_daily_transcript, write_clip_transcript
from app.storage.database import TRANSCRIPTION_DONE, TRANSCRIPTION_NO_SPEECH, VoiceLogDatabase
from app.storage.paths import transcript_path_for_audio, unique_audio_path
from app.storage.recovery import recover_spool_files
from app.summary.summary_service import split_daily_text


def frame(amplitude: int, samples: int = 100) -> np.ndarray:
    return np.full((samples, 1), amplitude, dtype=np.int16)


class VoiceLogCoreTests(unittest.TestCase):
    def test_vad_start_end_and_pre_roll(self) -> None:
        vad = EnergyVad(
            VadConfig(
                sample_rate=1000,
                frame_ms=100,
                pre_roll_sec=0.3,
                min_speech_sec=0.2,
                silence_timeout_sec=0.3,
                sensitivity="standard",
            )
        )
        for _ in range(20):
            decision = vad.consume(frame(80))
            self.assertFalse(decision.start_recording)

        starts = []
        for _ in range(3):
            starts.append(vad.consume(frame(5000)).start_recording)
        self.assertIn(True, starts)
        self.assertGreaterEqual(len(vad.pre_roll_blocks()), 3)

        stops = []
        for _ in range(5):
            stops.append(vad.consume(frame(0)).stop_recording)
        self.assertIn(True, stops)

    def test_rotation_and_date_rollover(self) -> None:
        policy = RotationPolicy(max_clip_sec=1800)
        started = datetime(2026, 7, 8, 9, 0, 0)
        self.assertFalse(policy.should_rotate(started, started + timedelta(seconds=1799)))
        self.assertTrue(policy.should_rotate(started, started + timedelta(seconds=1800)))
        self.assertTrue(policy.should_roll_date(datetime(2026, 7, 8, 23, 59), datetime(2026, 7, 9, 0, 0)))

    def test_unique_audio_path_adds_suffix(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            started = datetime(2026, 7, 8, 9, 5, 12)
            first = unique_audio_path(root, started)
            first.parent.mkdir(parents=True, exist_ok=True)
            first.write_bytes(b"")
            second = unique_audio_path(root, started)
            self.assertNotEqual(first, second)
            self.assertTrue(second.name.startswith("20260708_090512"))

    def test_db_transition_and_daily_transcript(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            db = VoiceLogDatabase(root / "state.db")
            started = datetime(2026, 7, 8, 9, 5, 12)
            ended = started + timedelta(minutes=8)
            audio_path = unique_audio_path(root / "logs", started)
            audio_path.parent.mkdir(parents=True, exist_ok=True)
            audio_path.write_bytes(b"RIFF")
            transcript_path = transcript_path_for_audio(root / "logs", audio_path, started.date())
            clip_id = db.create_clip("20260708_090512", started, ended, audio_path, transcript_path)
            db.queue_recorded_clips()
            clip = db.next_transcription_clip()
            self.assertIsNotNone(clip)
            assert clip is not None
            self.assertEqual(clip.id, clip_id)
            write_clip_transcript(transcript_path, clip, "おはようございます。\n書類の確認をお願いします。")
            db.update_clip_status(clip_id, TRANSCRIPTION_DONE, transcript_path=transcript_path)
            self.assertIn("2026-07-08", db.dates_needing_daily_transcript())
            output = build_daily_transcript(db, root / "logs", "2026-07-08")
            text = output.read_text(encoding="utf-8")
            self.assertIn("# 2026-07-08 声ログ", text)
            self.assertIn("09:05", text)
            self.assertIn("書類の確認", text)
            self.assertNotIn("2026-07-08", db.dates_needing_daily_transcript())

    def test_no_speech_is_skipped_in_daily_transcript(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            db = VoiceLogDatabase(root / "state.db")
            started = datetime(2026, 7, 8, 10, 0, 0)
            audio_path = unique_audio_path(root / "logs", started)
            transcript_path = transcript_path_for_audio(root / "logs", audio_path, started.date())
            clip_id = db.create_clip("s", started, started + timedelta(seconds=20), audio_path, transcript_path)
            db.update_clip_status(clip_id, TRANSCRIPTION_NO_SPEECH, transcript_path=transcript_path)
            output = build_daily_transcript(db, root / "logs", "2026-07-08")
            text = output.read_text(encoding="utf-8")
            self.assertNotIn("10:00", text)

    def test_recovery_finalizes_spool_pcm(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            spool = root / "spool"
            spool.mkdir()
            pcm = spool / "sample.pcm"
            pcm.write_bytes((np.zeros((160,), dtype=np.int16)).tobytes())
            meta = spool / "sample.json"
            meta.write_text(
                """{
  "pcm_path": "%s",
  "sample_rate": 16000,
  "channels": 1,
  "sample_width": 2,
  "start_time": "2026-07-08T09:05:12",
  "session_id": "20260708_090512"
}
"""
                % str(pcm).replace("\\", "\\\\"),
                encoding="utf-8",
            )
            results = recover_spool_files(spool, root / "logs")
            self.assertEqual(len(results), 1)
            self.assertTrue(results[0].wav_path.is_file())
            self.assertFalse(pcm.exists())
            self.assertFalse(meta.exists())

    def test_retry_policy_stops_after_max(self) -> None:
        policy = RetryPolicy(max_retries=3, base_delay_sec=10)
        self.assertTrue(policy.can_retry(2))
        self.assertFalse(policy.can_retry(3))
        self.assertGreater(policy.next_retry_at(2), datetime.now())

    def test_summary_chunk_split_preserves_order(self) -> None:
        text = "# 2026-07-08 声ログ\n\n" + "\n\n".join(f"## {hour:02d}:00–{hour:02d}:10\n内容{hour}" for hour in range(9, 18))
        chunks = split_daily_text(text, max_chars=120)
        self.assertGreater(len(chunks), 1)
        joined = "\n".join(chunks)
        self.assertLess(joined.index("内容9"), joined.index("内容17"))


if __name__ == "__main__":
    unittest.main()
