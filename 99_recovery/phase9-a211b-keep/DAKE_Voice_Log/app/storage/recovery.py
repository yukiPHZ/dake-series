# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import wave
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from app.storage.paths import transcript_path_for_audio, unique_audio_path


@dataclass(frozen=True)
class RecoveryResult:
    wav_path: Path
    started_at: datetime
    ended_at: datetime
    session_id: str


def finalize_pcm_to_wav(
    pcm_path: Path,
    wav_path: Path,
    sample_rate: int,
    channels: int,
    sample_width: int,
) -> None:
    wav_path.parent.mkdir(parents=True, exist_ok=True)
    with pcm_path.open("rb") as source, wave.open(str(wav_path), "wb") as target:
        target.setnchannels(channels)
        target.setsampwidth(sample_width)
        target.setframerate(sample_rate)
        while True:
            chunk = source.read(1024 * 1024)
            if not chunk:
                break
            target.writeframes(chunk)


def recover_spool_files(spool_dir: Path, save_root: Path) -> list[RecoveryResult]:
    results: list[RecoveryResult] = []
    for meta_path in sorted(spool_dir.glob("*.json")):
        try:
            payload = json.loads(meta_path.read_text(encoding="utf-8"))
            pcm_path = Path(payload["pcm_path"])
            if not pcm_path.is_file():
                continue
            started_at = datetime.fromisoformat(payload["start_time"])
            ended_at = datetime.now()
            wav_path = unique_audio_path(save_root, started_at)
            finalize_pcm_to_wav(
                pcm_path,
                wav_path,
                int(payload.get("sample_rate", 16000)),
                int(payload.get("channels", 1)),
                int(payload.get("sample_width", 2)),
            )
            results.append(
                RecoveryResult(
                    wav_path=wav_path,
                    started_at=started_at,
                    ended_at=ended_at,
                    session_id=str(payload.get("session_id") or started_at.strftime("%Y%m%d_%H%M%S")),
                )
            )
            pcm_path.unlink(missing_ok=True)
            meta_path.unlink(missing_ok=True)
        except Exception:
            continue
    return results


def recovered_transcript_path(save_root: Path, wav_path: Path, started_at: datetime) -> Path:
    return transcript_path_for_audio(save_root, wav_path, started_at.date())

