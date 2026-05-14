from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from core.app_config import load_config

try:
    from faster_whisper import WhisperModel
except ImportError:  # pragma: no cover - optional dependency
    WhisperModel = None  # type: ignore[assignment]


LogCallback = Callable[[str], None]
ProgressCallback = Callable[[float], None]


@dataclass(frozen=True)
class TranscriptionResult:
    available: bool
    transcript_path: Path | None
    srt_path: Path | None
    unavailable_path: Path | None
    reason: str
    segment_count: int = 0


def is_faster_whisper_available() -> bool:
    return WhisperModel is not None


def _srt_timestamp(seconds: float) -> str:
    total_ms = max(0, int(seconds * 1000))
    hours, remainder = divmod(total_ms, 3_600_000)
    minutes, remainder = divmod(remainder, 60_000)
    secs, ms = divmod(remainder, 1000)
    return f"{hours:02d}:{minutes:02d}:{secs:02d},{ms:03d}"


def _write_unavailable(project_dir: Path, reason: str) -> TranscriptionResult:
    path = project_dir / "transcript_unavailable.txt"
    path.write_text(
        "Transcription unavailable\n"
        f"Reason: {reason}\n\n"
        "Install faster-whisper and FFmpeg, then run the project again.\n",
        encoding="utf-8",
    )
    return TranscriptionResult(False, None, None, path, reason)


def transcribe_media(
    video_path: Path,
    project_dir: Path,
    log: LogCallback | None = None,
    progress: ProgressCallback | None = None,
) -> TranscriptionResult:
    if progress:
        progress(0.0)
    if WhisperModel is None:
        if log:
            log("補助脳：faster-whisperが見つかりません。文字起こしをスキップします。")
        return _write_unavailable(project_dir, "faster-whisper is not installed.")

    config = load_config()
    model_name = str(config.get("whisper_model") or "base")
    transcript_path = project_dir / "transcript.txt"
    srt_path = project_dir / "transcript.srt"

    try:
        if log:
            log(f"補助脳：文字起こしモデルを読み込んでいます。({model_name})")
        model = WhisperModel(model_name, device="auto")
        segments_iter, _info = model.transcribe(str(video_path), beam_size=5, vad_filter=True)

        lines: list[str] = []
        srt_blocks: list[str] = []
        count = 0
        for count, segment in enumerate(segments_iter, start=1):
            text = segment.text.strip()
            if not text:
                continue
            lines.append(f"[{_srt_timestamp(segment.start)[:-4]}] {text}")
            srt_blocks.append(
                f"{count}\n"
                f"{_srt_timestamp(segment.start)} --> {_srt_timestamp(segment.end)}\n"
                f"{text}\n"
            )
            if progress:
                progress(min(0.95, count / 120))

        if not lines:
            lines.append("No speech segment was detected.")
        transcript_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
        srt_path.write_text("\n".join(srt_blocks) + "\n", encoding="utf-8")
        if progress:
            progress(1.0)
        if log:
            log("補助脳：文字起こしを保存しました。")
        return TranscriptionResult(True, transcript_path, srt_path, None, "", count)
    except Exception as exc:
        reason = str(exc)
        if log:
            log("補助脳：文字起こしが利用できませんでした。")
        return _write_unavailable(project_dir, reason)
