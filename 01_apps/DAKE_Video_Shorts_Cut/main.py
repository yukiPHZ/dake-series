# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import html
import json
import mimetypes
import os
import queue
import re
import shutil
import socket
import subprocess
import sys
import tempfile
import threading
import webbrowser
from dataclasses import dataclass
from datetime import datetime
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from tkinter import filedialog, font as tkfont, messagebox, ttk
import tkinter as tk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore
except Exception:
    DND_FILES = None
    TkinterDnD = None


APP_NAME = "Dakeショート切り出し"
WINDOW_TITLE = "Dakeショート切り出し"
COPYRIGHT = "© 2026 しまりす不動産 / Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "main_title": "動画からショートを作る",
    "main_description": "MP4を入れるだけで、ショート候補とサムネを作成します。",
    "drop_title": "MP4をここへドロップ",
    "drop_description": "または下のボタンから1本だけ選びます。",
    "button_select_file": "MP4を選ぶ",
    "button_create": "ショートを作成",
    "button_open_output": "保存先を開く",
    "button_open_transfer": "転送ページを開く",
    "selected_file_label": "選択中",
    "selected_file_empty": "まだ動画が選ばれていません",
    "video_info_label": "動画情報",
    "video_info_empty": "MP4を選ぶと表示されます。",
    "video_info_template": "長さ: {duration} / 解像度: {width}x{height} / fps: {fps} / 音声: {audio}",
    "audio_yes": "あり",
    "audio_no": "なし",
    "status_label": "ステータス",
    "status_ready": "MP4を選んでください。",
    "status_checking": "動画を確認しています",
    "status_creating_shorts": "ショートを作成しています",
    "status_creating_thumbs": "サムネを作成しています",
    "status_transcribing": "文字起こししています",
    "status_transcript_done": "文字起こしを作成しました",
    "status_transcript_skipped": "文字起こしはスキップしました",
    "status_transcript_failed": "文字起こしに失敗しました",
    "status_reviewing": "候補を評価しています",
    "status_review_done": "候補評価を作成しました",
    "status_ollama_fallback": "Ollama失敗。通常評価で完了しました",
    "status_preparing_transfer": "スマホ転送を準備しています",
    "status_complete": "完了しました",
    "status_error": "エラーが発生しました",
    "result_title": "生成結果",
    "result_waiting": "作成後に結果が表示されます。",
    "candidate_label": "候補 {number}",
    "result_short": "short",
    "result_thumb": "thumb",
    "result_title_file": "title",
    "result_thumb_text": "thumb_text",
    "result_transcript": "transcript",
    "result_transcript_missing": "未作成",
    "result_score": "score",
    "recommend_waiting": "おすすめ: 未作成",
    "recommend_template": "おすすめ: 候補 {number}（{score}点） / {method}",
    "recommend_card_template": "おすすめ：候補{number}（{score}点） / {method}\nタイトル案：{title}\nサムネ文言：{thumbnail_text}\n理由：{reason}",
    "checkbox_transcribe": "ショートごとに文字起こし",
    "checkbox_ollama_review": "Ollamaで候補評価",
    "ollama_review_note": "Ollama起動中のみ使用できます。失敗時は通常評価に戻ります。",
    "review_method_label": "評価方法",
    "review_method_local": "通常評価",
    "review_method_ollama": "Ollama",
    "review_method_local_fallback": "通常評価（Ollama失敗）",
    "result_exists": "あり",
    "result_missing": "なし",
    "qr_title": "スマホ転送",
    "qr_waiting": "生成後にQRコードを表示します。",
    "qr_dependency_missing": "QR表示に qrcode / Pillow が必要です。URLから開けます。",
    "transfer_url_label": "URL",
    "phase_short_item": "{index}/3 ショートを作成中",
    "phase_thumb_item": "{index}/3 サムネを作成中",
    "title_candidate": "ショート候補 {number:02d}",
    "filetype_mp4": "MP4動画",
    "filetype_all": "すべてのファイル",
    "dialog_error_title": "確認してください",
    "dialog_done_title": "作成しました",
    "dialog_done_message": "ショート候補を作成しました。\nスマホ転送用QRコードを確認してください。",
    "error_not_mp4": "MP4ファイルを選んでください。",
    "error_file_missing": "選んだファイルが見つかりません。",
    "error_no_file": "先にMP4を選んでください。",
    "error_ffmpeg_missing": "ffmpeg が見つかりません。動画処理に必要です。",
    "error_ffprobe_missing": "ffprobe が見つかりません。動画確認に必要です。",
    "error_ffprobe_failed": "動画情報を取得できませんでした。",
    "error_no_video_stream": "動画ストリームが見つかりませんでした。",
    "error_process_failed": "動画処理に失敗しました。",
    "error_faster_whisper_missing": "faster-whisper が見つかりません。文字起こしには追加インストールが必要です。",
    "error_output_missing": "保存先がまだありません。",
    "error_open_output": "保存先フォルダを開けませんでした。",
    "error_open_transfer": "転送ページを開けませんでした。",
    "mobile_title": "ショート切り出し",
    "mobile_description": "PCで作成したショート動画・サムネ・タイトル案を保存できます。",
    "mobile_download": "ダウンロード",
    "mobile_title_preview": "タイトル案",
    "mobile_review_section": "候補評価",
    "mobile_recommend_section": "おすすめ候補",
    "mobile_development_section": "開発用ファイル",
    "mobile_score": "score",
    "mobile_thumbnail_text": "サムネ文言",
    "mobile_reason": "理由",
    "mobile_empty": "ファイルが見つかりません。",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_tagline": "止まらない、迷わない、すぐ終わる。",
    "footer_link_1": "戸建買取査定",
    "footer_link_2": "Instagram",
    "footer_separator": " ｜ ",
    "footer_copyright": COPYRIGHT,
    "launch_check_ok": "Dakeショート切り出し launch-check OK",
    "launch_check_ffmpeg": "ffmpeg={status}",
    "launch_check_ffprobe": "ffprobe={status}",
    "launch_check_segments": "segments=OK",
    "launch_check_html": "transfer_html=OK",
    "launch_check_transcript": "transcript_format=OK",
    "launch_check_review": "review=OK",
    "launch_check_qr": "qr_dependency={status}",
    "process_check_done": "process-check OK: {output_dir}",
    "process_check_transcript": "transcript={status}",
    "process_check_review": "review={status}",
}

THEME = {
    "background": "#F6F7F9",
    "panel": "#FFFFFF",
    "subtle": "#EEF2F7",
    "border": "#E2E8F0",
    "text": "#1E2430",
    "muted": "#667085",
    "accent": "#2F6FED",
    "accent_hover": "#2458BF",
    "success": "#118A4E",
    "success_bg": "#EAFBF3",
    "error": "#C92A2A",
    "error_bg": "#FDECEC",
    "link": "#58677D",
    "link_hover": "#2F6FED",
}

LINKS = {
    "assessment": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "instagram": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

WINDOW_SIZE = "940x820"
WINDOW_MIN_SIZE = (820, 720)
FONT_CANDIDATES = ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo", "Segoe UI")
OUTPUT_FOLDER_PREFIX = "dake_shorts_output"
REVIEW_TEXT_FILE = "shorts_review.txt"
REVIEW_JSON_FILE = "shorts_review.json"
APP_CONFIG_FILE = "DakeVideo_Shorts_Cut_config.json"
DEFAULT_OLLAMA_MODEL = "qwen2.5:7b"
DEFAULT_OLLAMA_URL = "http://localhost:11434/api/generate"
TRANSFER_START_PORT = 8765
TRANSFER_PORT_SCAN = 80
SHORT_COUNT = 3
SHORT_MIN_SECONDS = 45.0
SHORT_MAX_SECONDS = 60.0
QUEUE_POLL_MS = 100
COMMON_ICON_RELATIVE = Path("..") / ".." / "02_assets" / "dake_icon.ico"
WHISPER_MODEL_SIZE = "base"
WHISPER_DEVICE = "cpu"
WHISPER_COMPUTE_TYPE = "int8"
WHISPER_LANGUAGE = "ja"
REVIEW_KEYWORDS = (
    "作る",
    "AI",
    "DAKE",
    "稼働中",
    "お金",
    "遊び",
    "人格",
    "情報商材",
    "いいね",
    "これだ",
)
REVIEW_STRONG_WORDS = (
    "いい",
    "すごい",
    "やばい",
    "強い",
    "大事",
    "必要",
    "本当",
    "絶対",
    "できる",
    "したい",
    "なりたい",
    "なりたくない",
)


class UserFacingError(RuntimeError):
    pass


@dataclass(frozen=True)
class VideoInfo:
    duration: float
    width: int
    height: int
    fps: float
    has_audio: bool


@dataclass(frozen=True)
class ShortSegment:
    index: int
    start: float
    duration: float


@dataclass(frozen=True)
class CandidateReview:
    candidate: int
    score: int
    title: str
    thumbnail_text: str
    reason: str
    transcript_file: str


@dataclass
class GeneratedCandidate:
    index: int
    short_path: Path
    thumb_path: Path
    title_path: Path
    thumb_text_path: Path
    transcript_path: Path
    segments_path: Path
    review: CandidateReview | None = None
    transcript_error: str = ""


@dataclass(frozen=True)
class TranscriptOutput:
    text_path: Path
    segments_path: Path


@dataclass(frozen=True)
class ReviewOutput:
    text_path: Path
    json_path: Path
    reviews: list[CandidateReview]
    review_method: str
    best_candidate: int | None = None
    ollama_error: str = ""


@dataclass(frozen=True)
class ProcessResult:
    output_dir: Path
    video_info: VideoInfo
    candidates: list[GeneratedCandidate]
    transfer_url: str
    review: ReviewOutput | None = None
    transcript_error: str = ""


def get_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def get_config_path() -> Path:
    return get_app_dir() / APP_CONFIG_FILE


def default_app_config() -> dict[str, object]:
    return {
        "use_ollama_review": False,
        "ollama_model": DEFAULT_OLLAMA_MODEL,
        "ollama_url": DEFAULT_OLLAMA_URL,
    }


def load_app_config() -> dict[str, object]:
    config = default_app_config()
    path = get_config_path()
    if not path.exists():
        return config
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return config
    if not isinstance(payload, dict):
        return config
    config["use_ollama_review"] = bool(payload.get("use_ollama_review", config["use_ollama_review"]))
    model = str(payload.get("ollama_model") or "").strip()
    url = str(payload.get("ollama_url") or "").strip()
    if model:
        config["ollama_model"] = model
    if url:
        config["ollama_url"] = url
    return config


def save_app_config(config: dict[str, object]) -> None:
    path = get_config_path()
    payload = {
        "use_ollama_review": bool(config.get("use_ollama_review", False)),
        "ollama_model": str(config.get("ollama_model") or DEFAULT_OLLAMA_MODEL),
        "ollama_url": str(config.get("ollama_url") or DEFAULT_OLLAMA_URL),
    }
    try:
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    except OSError:
        return


def get_common_icon_candidates() -> list[Path]:
    base = get_app_dir()
    source = Path(__file__).resolve().parent
    return [
        source / COMMON_ICON_RELATIVE,
        base / COMMON_ICON_RELATIVE,
        base.parent.parent / "02_assets" / "dake_icon.ico",
        base / "dake_icon.ico",
    ]


def choose_font_family(root: tk.Tk) -> str:
    available = set(tkfont.families(root))
    for family in FONT_CANDIDATES:
        if family in available:
            return family
    return "TkDefaultFont"


def find_tool(name: str) -> str | None:
    return shutil.which(name)


def ensure_required_tools() -> tuple[str, str]:
    ffmpeg = find_tool("ffmpeg")
    if not ffmpeg:
        raise UserFacingError(UI_TEXT["error_ffmpeg_missing"])
    ffprobe = find_tool("ffprobe")
    if not ffprobe:
        raise UserFacingError(UI_TEXT["error_ffprobe_missing"])
    return ffmpeg, ffprobe


def format_seconds(seconds: float) -> str:
    seconds = max(0.0, seconds)
    minutes = int(seconds // 60)
    remain = int(round(seconds - minutes * 60))
    if remain == 60:
        minutes += 1
        remain = 0
    return f"{minutes}:{remain:02d}"


def format_timestamp(seconds: float) -> str:
    seconds = max(0.0, seconds)
    total_seconds = int(seconds)
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    remain = total_seconds % 60
    if hours <= 0:
        return f"{minutes:02d}:{remain:02d}"
    return f"{hours:02d}:{minutes:02d}:{remain:02d}"


def parse_fps(value: str) -> float:
    if not value or value == "0/0":
        return 0.0
    if "/" in value:
        numerator, denominator = value.split("/", 1)
        try:
            den = float(denominator)
            return float(numerator) / den if den else 0.0
        except ValueError:
            return 0.0
    try:
        return float(value)
    except ValueError:
        return 0.0


def is_mp4(path: Path) -> bool:
    return path.suffix.lower() == ".mp4"


def validate_mp4_path(path: Path) -> Path:
    if not path.exists():
        raise UserFacingError(UI_TEXT["error_file_missing"])
    if not is_mp4(path):
        raise UserFacingError(UI_TEXT["error_not_mp4"])
    return path


def get_subprocess_creationflags() -> int:
    if os.name == "nt":
        return getattr(subprocess, "CREATE_NO_WINDOW", 0)
    return 0


def run_subprocess(command: list[str], error_message: str) -> subprocess.CompletedProcess[str]:
    try:
        completed = subprocess.run(
            command,
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=get_subprocess_creationflags(),
        )
    except OSError as exc:
        raise UserFacingError(error_message) from exc
    if completed.returncode != 0:
        detail = (completed.stderr or completed.stdout or "").strip()
        if detail:
            detail = detail.splitlines()[0][:220]
            raise UserFacingError(f"{error_message}\n{detail}")
        raise UserFacingError(error_message)
    return completed


def probe_video(input_path: Path, ffprobe_path: str) -> VideoInfo:
    completed = run_subprocess(
        [
            ffprobe_path,
            "-v",
            "error",
            "-print_format",
            "json",
            "-show_format",
            "-show_streams",
            str(input_path),
        ],
        UI_TEXT["error_ffprobe_failed"],
    )
    try:
        payload = json.loads(completed.stdout)
    except json.JSONDecodeError as exc:
        raise UserFacingError(UI_TEXT["error_ffprobe_failed"]) from exc

    streams = payload.get("streams")
    if not isinstance(streams, list):
        raise UserFacingError(UI_TEXT["error_ffprobe_failed"])

    video_stream = next((item for item in streams if item.get("codec_type") == "video"), None)
    if not isinstance(video_stream, dict):
        raise UserFacingError(UI_TEXT["error_no_video_stream"])

    audio_stream = next((item for item in streams if item.get("codec_type") == "audio"), None)
    duration_value = video_stream.get("duration") or payload.get("format", {}).get("duration") or 0
    try:
        duration = float(duration_value)
    except (TypeError, ValueError):
        duration = 0.0

    width = int(video_stream.get("width") or 0)
    height = int(video_stream.get("height") or 0)
    fps = parse_fps(str(video_stream.get("avg_frame_rate") or video_stream.get("r_frame_rate") or "0/0"))
    if duration <= 0 or width <= 0 or height <= 0:
        raise UserFacingError(UI_TEXT["error_ffprobe_failed"])
    return VideoInfo(duration=duration, width=width, height=height, fps=fps, has_audio=audio_stream is not None)


def build_segments(duration: float) -> list[ShortSegment]:
    if duration <= 0:
        return []
    third = duration / SHORT_COUNT
    if third >= SHORT_MIN_SECONDS:
        clip_duration = min(SHORT_MAX_SECONDS, third)
    else:
        clip_duration = max(1.0, third)
    clip_duration = min(clip_duration, duration)
    max_start = max(0.0, duration - clip_duration)
    segments: list[ShortSegment] = []
    for index in range(SHORT_COUNT):
        start = min(index * third, max_start)
        segments.append(ShortSegment(index=index + 1, start=start, duration=clip_duration))
    return segments


def create_output_dir(input_path: Path, now: datetime | None = None) -> Path:
    timestamp = (now or datetime.now()).strftime("%Y%m%d_%H%M%S")
    output_parent = choose_output_parent(input_path)
    base = output_parent / f"{OUTPUT_FOLDER_PREFIX}_{timestamp}"
    candidate = base
    serial = 2
    while candidate.exists():
        candidate = output_parent / f"{OUTPUT_FOLDER_PREFIX}_{timestamp}_{serial:02d}"
        serial += 1
    candidate.mkdir(parents=True, exist_ok=False)
    return candidate


def choose_output_parent(input_path: Path) -> Path:
    parent = input_path.parent
    while parent.name.startswith(OUTPUT_FOLDER_PREFIX) and parent.parent != parent:
        parent = parent.parent
    return parent


def video_filter() -> str:
    return "scale=1080:1920:force_original_aspect_ratio=increase,crop=1080:1920"


def ffmpeg_time(value: float) -> str:
    return f"{max(0.0, value):.3f}"


def create_short_video(
    ffmpeg_path: str,
    input_path: Path,
    output_path: Path,
    segment: ShortSegment,
    has_audio: bool,
) -> None:
    command = [
        ffmpeg_path,
        "-y",
        "-ss",
        ffmpeg_time(segment.start),
        "-i",
        str(input_path),
        "-t",
        ffmpeg_time(segment.duration),
        "-vf",
        video_filter(),
        "-c:v",
        "libx264",
        "-preset",
        "veryfast",
        "-crf",
        "23",
        "-pix_fmt",
        "yuv420p",
    ]
    if has_audio:
        command.extend(["-c:a", "aac", "-b:a", "160k"])
    else:
        command.append("-an")
    command.extend(["-movflags", "+faststart", str(output_path)])
    run_subprocess(command, UI_TEXT["error_process_failed"])


def create_thumbnail(ffmpeg_path: str, input_path: Path, output_path: Path, segment: ShortSegment) -> None:
    thumb_time = segment.start + min(3.0, max(0.0, segment.duration / 2))
    command = [
        ffmpeg_path,
        "-y",
        "-ss",
        ffmpeg_time(thumb_time),
        "-i",
        str(input_path),
        "-frames:v",
        "1",
        "-vf",
        video_filter(),
        "-q:v",
        "2",
        str(output_path),
    ]
    run_subprocess(command, UI_TEXT["error_process_failed"])


def write_title_file(path: Path, number: int) -> None:
    path.write_text(UI_TEXT["title_candidate"].format(number=number) + "\n", encoding="utf-8")


def short_transcript_text_name(index: int) -> str:
    return f"short_{index:02d}_transcript.txt"


def thumb_text_name(index: int) -> str:
    return f"thumb_text_{index:02d}.txt"


def short_segments_json_name(index: int) -> str:
    return f"short_{index:02d}_segments.json"


def write_transcript_files(text_path: Path, segments_path: Path, segments: list[dict[str, object]]) -> TranscriptOutput:
    text_blocks: list[str] = []
    for segment in segments:
        text = str(segment.get("text") or "").strip()
        if not text:
            continue
        start_text = str(segment.get("start_text") or "")
        end_text = str(segment.get("end_text") or "")
        text_blocks.append(f"[{start_text} - {end_text}]\n{text}")
    text_path.write_text("\n\n".join(text_blocks) + ("\n" if text_blocks else ""), encoding="utf-8")
    segments_path.write_text(json.dumps(segments, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return TranscriptOutput(text_path=text_path, segments_path=segments_path)


def transcribe_video(input_path: Path, text_path: Path, segments_path: Path) -> TranscriptOutput:
    try:
        from faster_whisper import WhisperModel
    except Exception as exc:
        raise UserFacingError(UI_TEXT["error_faster_whisper_missing"]) from exc

    try:
        model = WhisperModel(WHISPER_MODEL_SIZE, device=WHISPER_DEVICE, compute_type=WHISPER_COMPUTE_TYPE)
        raw_segments, _info = model.transcribe(str(input_path), language=WHISPER_LANGUAGE)
        transcript_segments: list[dict[str, object]] = []
        for segment in raw_segments:
            text = str(getattr(segment, "text", "") or "").strip()
            start = float(getattr(segment, "start", 0.0) or 0.0)
            end = float(getattr(segment, "end", start) or start)
            if not text:
                continue
            transcript_segments.append(
                {
                    "start": start,
                    "end": end,
                    "start_text": format_timestamp(start),
                    "end_text": format_timestamp(end),
                    "text": text,
                }
            )
    except UserFacingError:
        raise
    except Exception as exc:
        raise UserFacingError(UI_TEXT["status_transcript_failed"]) from exc

    return write_transcript_files(text_path, segments_path, transcript_segments)


def clean_transcript_for_review(raw_text: str) -> str:
    lines: list[str] = []
    for line in raw_text.splitlines():
        value = line.strip()
        if not value:
            continue
        if value.startswith("[") and "]" in value:
            continue
        lines.append(value)
    return " ".join(lines).strip()


def read_candidate_transcript(candidate: GeneratedCandidate) -> str:
    if not candidate.transcript_path.exists():
        return ""
    try:
        return clean_transcript_for_review(candidate.transcript_path.read_text(encoding="utf-8"))
    except OSError:
        return ""


def count_review_hits(text: str, words: tuple[str, ...]) -> list[str]:
    lowered = text.casefold()
    hits: list[str] = []
    for word in words:
        needle = word.casefold()
        if needle in lowered:
            hits.append(word)
    return hits


def clamp_score(value: int) -> int:
    return max(0, min(100, value))


def pick_short_phrase(text: str, fallback: str) -> str:
    for word in ("これだ", "いいね", "人格", "お金", "AI", "DAKE", "作る", "遊び", "稼働中", "情報商材"):
        if word.casefold() in text.casefold():
            return word if len(word) <= 10 else word[:10]
    parts = [part.strip() for part in re.split(r"[。！？!?\n]", text) if part.strip()]
    for part in parts:
        compact = re.sub(r"\s+", "", part)
        if 4 <= len(compact) <= 14:
            return compact
    if text:
        compact = re.sub(r"\s+", "", text)
        return compact[:12] + ("…" if len(compact) > 12 else "")
    return fallback


def suggest_review_title(candidate: GeneratedCandidate, text: str) -> str:
    lowered = text.casefold()
    if "情報商材" in text:
        return "情報商材にしない話"
    if "人格" in text:
        return "人格で伝わるショート"
    if "お金" in text:
        return "お金の話を切り出す"
    if "ai" in lowered or "AI" in text:
        return "AIで変えるショート"
    if "dake" in lowered or "DAKE" in text:
        return "DAKEらしさが出る候補"
    if "作る" in text:
        return "作る話のショート"
    return UI_TEXT["title_candidate"].format(number=candidate.index)


def build_review_reason(length: int, keyword_hits: list[str], strong_hits: list[str], score: int) -> str:
    if length <= 0:
        return "文字起こしが未作成または空のため、判断材料が少ない"
    reasons: list[str] = []
    if 35 <= length <= 220:
        reasons.append("会話量があり、短く切り出しやすい")
    elif length < 35:
        reasons.append("文字量が少なく、単独では伝わりにくい")
    else:
        reasons.append("情報量が多めなので、さらに短く整える余地がある")
    if keyword_hits:
        reasons.append("使いやすい語が含まれる")
    if strong_hits:
        reasons.append("感情や断定のある言葉がある")
    if score >= 75:
        reasons.append("ショート候補として優先しやすい")
    return "、".join(reasons)


def evaluate_candidate(candidate: GeneratedCandidate) -> CandidateReview:
    text = read_candidate_transcript(candidate)
    length = len(re.sub(r"\s+", "", text))
    keyword_hits = count_review_hits(text, REVIEW_KEYWORDS)
    strong_hits = count_review_hits(text, REVIEW_STRONG_WORDS)

    if length <= 0:
        score = 5
    elif length < 15:
        score = 18
    elif length < 35:
        score = 38
    elif length <= 140:
        score = 55
    elif length <= 260:
        score = 48
    else:
        score = 34

    score += min(30, len(keyword_hits) * 7)
    score += min(20, len(strong_hits) * 5)
    if re.search(r"[。！？!?]", text):
        score += 4
    if any(phrase in text for phrase in ("これだ", "いいね", "なりたくない", "実は")):
        score += 8
    if length > 320:
        score -= 12
    score = clamp_score(score)

    fallback = UI_TEXT["title_candidate"].format(number=candidate.index)
    return CandidateReview(
        candidate=candidate.index,
        score=score,
        title=suggest_review_title(candidate, text),
        thumbnail_text=pick_short_phrase(text, fallback),
        reason=build_review_reason(length, keyword_hits, strong_hits, score),
        transcript_file=candidate.transcript_path.name,
    )


def review_to_dict(review: CandidateReview) -> dict[str, object]:
    return {
        "candidate": review.candidate,
        "score": review.score,
        "title": review.title,
        "thumbnail_text": review.thumbnail_text,
        "thumbnail_text_file": thumb_text_name(review.candidate),
        "reason": review.reason,
        "transcript_file": review.transcript_file,
    }


def review_method_label(method: str) -> str:
    if method == "ollama":
        return UI_TEXT["review_method_ollama"]
    if method == "local_fallback":
        return UI_TEXT["review_method_local_fallback"]
    return UI_TEXT["review_method_local"]


def short_error_message(exc: Exception) -> str:
    message = str(exc).strip() or exc.__class__.__name__
    message = " ".join(message.splitlines())
    return message[:180]


def best_candidate_from_reviews(reviews: list[CandidateReview]) -> int | None:
    if not reviews:
        return None
    return max(reviews, key=lambda review: (review.score, -review.candidate)).candidate


def build_ollama_review_prompt(candidates: list[GeneratedCandidate]) -> str:
    blocks: list[str] = []
    for candidate in candidates:
        text = read_candidate_transcript(candidate) or "（文字起こしなし）"
        blocks.append(f"候補{candidate.index}:\n{text[:4000]}")
    return """あなたは「ユキズ稼働中」向けのショート動画候補を選ぶ編集者です。
以下の3候補を、ショート動画として切り出す価値で評価してください。

評価基準:
- 一言で伝わるか
- 見た人が続きを見たくなるか
- ユキズ稼働中らしさがあるか
- 静かな制作感があるか
- 人格・思想・作る姿勢が出ているか
- 誇張したタイトルにしない
- サムネ文言は短く、強く、静かに
- 弱い候補でも、タイトルとサムネ文言で候補を貶さない
- 理由は率直に、ただし攻撃的にしない

必ずJSONだけを返してください。説明文やMarkdownは不要です。
形式:
{
  "reviews": [
    {
      "candidate": 1,
      "score": 72,
      "title": "静かに作る時間",
      "thumbnail_text": "ただ作る。",
      "reason": "作業の空気が伝わり、短尺でも雰囲気がある"
    }
  ],
  "best_candidate": 1
}

""" + "\n\n".join(blocks)


def extract_json_payload(text: str) -> dict[str, object]:
    value = text.strip()
    try:
        payload = json.loads(value)
        if isinstance(payload, dict):
            return payload
    except json.JSONDecodeError:
        pass
    start = value.find("{")
    end = value.rfind("}")
    if start < 0 or end <= start:
        raise ValueError("Ollama response did not contain JSON")
    payload = json.loads(value[start : end + 1])
    if not isinstance(payload, dict):
        raise ValueError("Ollama response JSON was not an object")
    return payload


def coerce_score(value: object) -> int:
    try:
        return clamp_score(int(round(float(value))))
    except (TypeError, ValueError):
        raise ValueError("Ollama score was invalid")


def sanitize_ollama_field(value: object, fallback: str, max_length: int) -> str:
    text = str(value or "").strip()
    if not text:
        return fallback
    noisy_words = (
        "弱すぎ",
        "無駄",
        "意味不明",
        "言ってること何",
        "ではない",
        "欠け",
        "違う",
        "低め",
        "単調",
        "FQUE",
        "ENERGY",
    )
    if any(word.casefold() in text.casefold() for word in noisy_words):
        return fallback
    latin_words = re.findall(r"[A-Za-z]{4,}", text)
    allowed = {"dake"}
    if any(word.casefold() not in allowed for word in latin_words):
        return fallback
    return text.replace(".", "。")[:max_length] or fallback


def parse_ollama_reviews(payload: dict[str, object], candidates: list[GeneratedCandidate]) -> tuple[list[CandidateReview], int | None]:
    raw_reviews = payload.get("reviews")
    if not isinstance(raw_reviews, list):
        raise ValueError("Ollama response did not include reviews")

    by_candidate: dict[int, dict[str, object]] = {}
    for item in raw_reviews:
        if not isinstance(item, dict):
            continue
        try:
            candidate_number = int(item.get("candidate", 0) or 0)
        except (TypeError, ValueError):
            continue
        by_candidate[candidate_number] = item

    reviews: list[CandidateReview] = []
    for candidate in candidates:
        raw = by_candidate.get(candidate.index)
        if raw is None:
            raise ValueError("Ollama response missed a candidate")
        local = evaluate_candidate(candidate)
        title = sanitize_ollama_field(raw.get("title"), local.title, 80)
        thumbnail_text = sanitize_ollama_field(raw.get("thumbnail_text"), local.thumbnail_text, 40)
        reason = sanitize_ollama_field(raw.get("reason"), local.reason, 180)
        review = CandidateReview(
            candidate=candidate.index,
            score=coerce_score(raw.get("score")),
            title=title or local.title,
            thumbnail_text=thumbnail_text or local.thumbnail_text,
            reason=reason or local.reason,
            transcript_file=candidate.transcript_path.name,
        )
        reviews.append(review)

    best_candidate = payload.get("best_candidate")
    try:
        best_number = int(best_candidate) if best_candidate is not None else None
    except (TypeError, ValueError):
        best_number = None
    valid_numbers = {candidate.index for candidate in candidates}
    if best_number not in valid_numbers:
        best_number = best_candidate_from_reviews(reviews)
    return reviews, best_number


def request_ollama_reviews(
    candidates: list[GeneratedCandidate],
    model: str = DEFAULT_OLLAMA_MODEL,
    url: str = DEFAULT_OLLAMA_URL,
) -> tuple[list[CandidateReview], int | None]:
    import urllib.error
    import urllib.request

    prompt = build_ollama_review_prompt(candidates)
    request_body = json.dumps(
        {
            "model": model,
            "prompt": prompt,
            "stream": False,
            "format": "json",
        },
        ensure_ascii=False,
    ).encode("utf-8")
    request = urllib.request.Request(
        url,
        data=request_body,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=75) as response:
            response_text = response.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")[:180]
        raise RuntimeError(detail or f"Ollama HTTP {exc.code}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(str(exc.reason)) from exc
    response_payload = json.loads(response_text)
    if not isinstance(response_payload, dict):
        raise ValueError("Ollama API response was not an object")
    if "error" in response_payload:
        raise RuntimeError(str(response_payload.get("error") or "Ollama error"))
    if "reviews" in response_payload:
        review_payload = response_payload
    else:
        raw_review = str(response_payload.get("response") or "")
        review_payload = extract_json_payload(raw_review)
    return parse_ollama_reviews(review_payload, candidates)


def write_review_files(
    output_dir: Path,
    candidates: list[GeneratedCandidate],
    use_ollama_review: bool = False,
    ollama_model: str = DEFAULT_OLLAMA_MODEL,
    ollama_url: str = DEFAULT_OLLAMA_URL,
) -> ReviewOutput:
    review_method = "local"
    ollama_error = ""
    best_candidate: int | None = None
    if use_ollama_review:
        try:
            reviews, best_candidate = request_ollama_reviews(candidates, model=ollama_model, url=ollama_url)
            review_method = "ollama"
        except Exception as exc:
            ollama_error = short_error_message(exc)
            review_method = "local_fallback"
            reviews = [evaluate_candidate(candidate) for candidate in candidates]
    else:
        reviews = [evaluate_candidate(candidate) for candidate in candidates]

    for candidate, review in zip(candidates, reviews):
        candidate.review = review
        candidate.title_path.write_text(review.title + "\n", encoding="utf-8")
        candidate.thumb_text_path.write_text(review.thumbnail_text + "\n", encoding="utf-8")

    if best_candidate is None:
        best_candidate = best_candidate_from_reviews(reviews)

    text_blocks = []
    method_label = review_method_label(review_method)
    if best_candidate is None:
        text_blocks.append(f"{UI_TEXT['review_method_label']}：{method_label}")
    else:
        best_review = next((review for review in reviews if review.candidate == best_candidate), None)
        best_score = best_review.score if best_review is not None else 0
        text_blocks.append(
            "\n".join(
                [
                    f"{UI_TEXT['review_method_label']}：{method_label}",
                    f"おすすめ：候補{best_candidate}（{best_score}点）",
                ]
            )
        )
    if ollama_error:
        text_blocks.append(f"Ollamaエラー：{ollama_error}")
    for review in reviews:
        text_blocks.append(
            "\n".join(
                [
                    f"候補{review.candidate}：{review.score}点",
                    f"タイトル案：{review.title}",
                    f"サムネ文言：{review.thumbnail_text}",
                    f"理由：{review.reason}",
                ]
            )
        )
    text_path = output_dir / REVIEW_TEXT_FILE
    json_path = output_dir / REVIEW_JSON_FILE
    text_path.write_text("\n\n".join(text_blocks) + ("\n" if text_blocks else ""), encoding="utf-8")
    payload: dict[str, object] = {
        "review_method": review_method,
        "best_candidate": best_candidate,
        "reviews": [review_to_dict(review) for review in reviews],
    }
    if ollama_error:
        payload["ollama_error"] = ollama_error
    json_path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return ReviewOutput(
        text_path=text_path,
        json_path=json_path,
        reviews=reviews,
        review_method=review_method,
        best_candidate=best_candidate,
        ollama_error=ollama_error,
    )


def load_review_payload(output_dir: Path) -> dict[str, object]:
    path = output_dir / REVIEW_JSON_FILE
    if not path.exists():
        return {}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    if isinstance(payload, list):
        reviews = [item for item in payload if isinstance(item, dict)]
        return {
            "review_method": "local",
            "best_candidate": best_candidate_from_review_dicts(reviews),
            "reviews": reviews,
        }
    if isinstance(payload, dict):
        return payload
    return {}


def best_candidate_from_review_dicts(reviews: list[dict[str, object]]) -> int | None:
    best_number: int | None = None
    best_score = -1
    for item in reviews:
        try:
            candidate_number = int(item.get("candidate", 0) or 0)
            score = int(item.get("score", 0) or 0)
        except (TypeError, ValueError):
            continue
        if score > best_score:
            best_number = candidate_number
            best_score = score
    return best_number


def load_review_dicts(output_dir: Path) -> list[dict[str, object]]:
    payload = load_review_payload(output_dir)
    raw_reviews = payload.get("reviews")
    if not isinstance(raw_reviews, list):
        return []
    return [item for item in raw_reviews if isinstance(item, dict)]


def get_best_review(candidates: list[GeneratedCandidate]) -> CandidateReview | None:
    reviews = [candidate.review for candidate in candidates if candidate.review is not None]
    if not reviews:
        return None
    return max(reviews, key=lambda review: (review.score, -review.candidate))


def get_recommended_review(candidates: list[GeneratedCandidate], review_output: ReviewOutput | None = None) -> CandidateReview | None:
    if review_output is not None and review_output.best_candidate is not None:
        for candidate in candidates:
            if candidate.review is not None and candidate.index == review_output.best_candidate:
                return candidate.review
    return get_best_review(candidates)


def get_local_ip() -> str:
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        sock.connect(("8.8.8.8", 80))
        return sock.getsockname()[0]
    except OSError:
        try:
            return socket.gethostbyname(socket.gethostname())
        except OSError:
            return "127.0.0.1"
    finally:
        sock.close()


def find_available_port(start_port: int = TRANSFER_START_PORT) -> int:
    for offset in range(TRANSFER_PORT_SCAN + 1):
        port = start_port + offset
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            try:
                sock.bind(("0.0.0.0", port))
            except OSError:
                continue
            return port
    raise UserFacingError(UI_TEXT["error_open_transfer"])


def mobile_file_link(item: Path) -> str:
    if not item.exists():
        return f"<span class=\"missing\">{html.escape(item.name)} / {html.escape(UI_TEXT['mobile_empty'])}</span>"
    href = f"/files/{html.escape(item.name)}"
    label = html.escape(item.name)
    return f"<a href=\"{href}\" download>{label}<span>{html.escape(UI_TEXT['mobile_download'])}</span></a>"


def mobile_text_preview(item: Path, label: str) -> str:
    if not item.exists() or item.suffix.lower() != ".txt":
        return ""
    try:
        text = item.read_text(encoding="utf-8").strip()
    except OSError:
        return ""
    if not text:
        return ""
    return "<p class=\"title-preview\">" + html.escape(label) + ": " + html.escape(text) + "</p>"


def review_summary_html(review: dict[str, object]) -> str:
    return (
        "<div class=\"review-summary\">"
        + "<p><strong>"
        + html.escape(UI_TEXT["mobile_score"])
        + f": {html.escape(str(review.get('score', '')))}"
        + "</strong></p>"
        + "<p>"
        + html.escape(UI_TEXT["mobile_title_preview"])
        + ": "
        + html.escape(str(review.get("title", "")))
        + "</p>"
        + "<p>"
        + html.escape(UI_TEXT["mobile_thumbnail_text"])
        + ": "
        + html.escape(str(review.get("thumbnail_text", "")))
        + "</p>"
        + "<p>"
        + html.escape(UI_TEXT["mobile_reason"])
        + ": "
        + html.escape(str(review.get("reason", "")))
        + "</p>"
        + "</div>"
    )


def build_transfer_html(output_dir: Path) -> str:
    rows: list[str] = []
    review_payload = load_review_payload(output_dir)
    method = str(review_payload.get("review_method") or "local")
    method_label = review_method_label(method)
    review_header = (
        "<p class=\"review-method\">"
        + html.escape(UI_TEXT["review_method_label"])
        + "："
        + html.escape(method_label)
        + "</p>"
    )
    reviews = load_review_dicts(output_dir)
    review_map: dict[int, dict[str, object]] = {}
    for item in reviews:
        try:
            candidate_number = int(item.get("candidate", 0) or 0)
        except (TypeError, ValueError):
            continue
        review_map[candidate_number] = item

    try:
        best_candidate = int(review_payload.get("best_candidate", 0) or 0)
    except (TypeError, ValueError):
        best_candidate = 0
    if best_candidate <= 0:
        best_candidate = best_candidate_from_review_dicts(reviews) or 0
    best_review = review_map.get(best_candidate)
    if best_candidate > 0:
        best_items = [
            output_dir / f"short_{best_candidate:02d}.mp4",
            output_dir / f"thumb_{best_candidate:02d}.jpg",
            output_dir / f"title_{best_candidate:02d}.txt",
            output_dir / thumb_text_name(best_candidate),
            output_dir / short_transcript_text_name(best_candidate),
        ]
        best_links = [mobile_file_link(item) for item in best_items]
        review_text = output_dir / REVIEW_TEXT_FILE
        if review_text.exists():
            best_links.append(mobile_file_link(review_text))
        best_summary = review_summary_html(best_review) if best_review else ""
        rows.append(
            "<section class=\"recommend-section\"><h2>"
            + html.escape(UI_TEXT["mobile_recommend_section"])
            + "</h2>"
            + review_header
            + best_summary
            + "\n".join(best_links)
            + "</section>"
        )

    for index in range(1, SHORT_COUNT + 1):
        items = [
            output_dir / f"short_{index:02d}.mp4",
            output_dir / f"thumb_{index:02d}.jpg",
            output_dir / f"title_{index:02d}.txt",
            output_dir / thumb_text_name(index),
            output_dir / short_transcript_text_name(index),
        ]
        links = []
        review = review_map.get(index)
        if review:
            links.append(review_summary_html(review))
        for item in items:
            links.append(mobile_file_link(item))
            if item.name.startswith("title_") and item.suffix.lower() == ".txt":
                links.append(mobile_text_preview(item, UI_TEXT["mobile_title_preview"]))
            if item.name.startswith("thumb_text_") and item.suffix.lower() == ".txt":
                links.append(mobile_text_preview(item, UI_TEXT["mobile_thumbnail_text"]))
        rows.append(
            "<section><h2>"
            + html.escape(UI_TEXT["candidate_label"].format(number=index))
            + "</h2>"
            + "\n".join(links)
            + "</section>"
        )

    development_items = [output_dir / REVIEW_JSON_FILE] + [
        output_dir / short_segments_json_name(index) for index in range(1, SHORT_COUNT + 1)
    ]
    development_links = [mobile_file_link(item) for item in development_items if item.exists()]
    if development_links:
        rows.append(
            "<section><h2>"
            + html.escape(UI_TEXT["mobile_development_section"])
            + "</h2>"
            + "\n".join(development_links)
            + "</section>"
        )

    return f"""<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{html.escape(UI_TEXT["mobile_title"])}</title>
<style>
:root {{
  color-scheme: light;
  --bg: {THEME["background"]};
  --panel: {THEME["panel"]};
  --text: {THEME["text"]};
  --muted: {THEME["muted"]};
  --border: {THEME["border"]};
  --accent: {THEME["accent"]};
}}
* {{ box-sizing: border-box; }}
body {{
  margin: 0;
  min-height: 100vh;
  background: var(--bg);
  color: var(--text);
  font-family: "BIZ UDPGothic", "Yu Gothic UI", Meiryo, sans-serif;
  padding: 22px;
}}
main {{
  width: min(100%, 520px);
  margin: 0 auto;
}}
h1 {{
  margin: 0 0 8px;
  font-size: 25px;
  letter-spacing: 0;
}}
.lead {{
  margin: 0 0 18px;
  color: var(--muted);
  line-height: 1.7;
  font-size: 14px;
}}
section {{
  background: var(--panel);
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 16px;
  margin: 0 0 14px;
}}
.recommend-section {{
  background: #F8FAFD;
  border-left: 4px solid var(--accent);
}}
h2 {{
  margin: 0 0 12px;
  font-size: 17px;
  letter-spacing: 0;
}}
a {{
  min-height: 46px;
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 12px;
  margin: 8px 0;
  color: var(--text);
  display: flex;
  justify-content: space-between;
  align-items: center;
  gap: 12px;
  text-decoration: none;
  font-weight: 700;
}}
a span {{
  color: var(--accent);
  font-size: 13px;
}}
.title-preview,
.review-summary,
.review-method,
.missing {{
  display: block;
  color: var(--muted);
  font-size: 13px;
  line-height: 1.7;
  margin: 8px 0 0;
}}
.review-method {{
  margin: 0 0 10px;
  font-weight: 700;
}}
.review-summary {{
  background: #F8FAFD;
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 10px 12px;
  margin-bottom: 10px;
}}
.review-summary p {{
  margin: 0 0 4px;
}}
.review-summary p:last-child {{
  margin-bottom: 0;
}}
</style>
</head>
<body>
<main>
<h1>{html.escape(UI_TEXT["mobile_title"])}</h1>
<p class="lead">{html.escape(UI_TEXT["mobile_description"])}</p>
{''.join(rows)}
</main>
</body>
</html>"""


class TransferRequestHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, output_dir: Path, **kwargs) -> None:
        self.output_dir = output_dir
        super().__init__(*args, directory=str(output_dir), **kwargs)

    def log_message(self, _format: str, *args) -> None:
        return

    def do_GET(self) -> None:
        if self.path in ("/", "/index.html"):
            self._send_html(build_transfer_html(self.output_dir))
            return
        if self.path.startswith("/files/"):
            self._send_file(self.path.removeprefix("/files/"))
            return
        self.send_error(404)

    def _send_html(self, content: str) -> None:
        payload = content.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def _send_file(self, raw_name: str) -> None:
        filename = Path(raw_name).name
        path = self.output_dir / filename
        if not path.exists() or not path.is_file():
            self.send_error(404)
            return
        content_type = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        try:
            payload = path.read_bytes()
        except OSError:
            self.send_error(404)
            return
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(payload)))
        self.send_header("Content-Disposition", f'attachment; filename="{path.name}"')
        self.end_headers()
        self.wfile.write(payload)


class TransferServer:
    def __init__(self) -> None:
        self.server: ThreadingHTTPServer | None = None
        self.thread: threading.Thread | None = None
        self.url = ""

    def start(self, output_dir: Path) -> str:
        self.shutdown()
        port = find_available_port()
        handler = partial(TransferRequestHandler, output_dir=output_dir)
        self.server = ThreadingHTTPServer(("0.0.0.0", port), handler)
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()
        self.url = f"http://{get_local_ip()}:{port}/"
        return self.url

    def shutdown(self) -> None:
        if self.server is None:
            return
        self.server.shutdown()
        self.server.server_close()
        if self.thread is not None:
            self.thread.join(timeout=2.0)
        self.server = None
        self.thread = None
        self.url = ""


def make_qr_photo(url: str, size: int = 180):
    try:
        import qrcode
        from PIL import Image, ImageTk
    except Exception as exc:
        raise UserFacingError(UI_TEXT["qr_dependency_missing"]) from exc

    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10,
        border=2,
    )
    qr.add_data(url)
    qr.make(fit=True)
    image = qr.make_image(fill_color=THEME["text"], back_color=THEME["panel"]).convert("RGB")
    resample = Image.Resampling.NEAREST if hasattr(Image, "Resampling") else Image.NEAREST
    image = image.resize((size, size), resample)
    return ImageTk.PhotoImage(image)


def process_video(
    input_path: Path,
    status_callback=None,
    transfer_server: TransferServer | None = None,
    create_transcript: bool = True,
    use_ollama_review: bool = False,
    ollama_model: str = DEFAULT_OLLAMA_MODEL,
    ollama_url: str = DEFAULT_OLLAMA_URL,
) -> ProcessResult:
    input_path = validate_mp4_path(input_path)
    ffmpeg_path, ffprobe_path = ensure_required_tools()
    if status_callback:
        status_callback(UI_TEXT["status_checking"])
    video_info = probe_video(input_path, ffprobe_path)
    segments = build_segments(video_info.duration)
    output_dir = create_output_dir(input_path)
    candidates: list[GeneratedCandidate] = []

    for segment in segments:
        if status_callback:
            status_callback(UI_TEXT["phase_short_item"].format(index=segment.index))
        short_path = output_dir / f"short_{segment.index:02d}.mp4"
        create_short_video(ffmpeg_path, input_path, short_path, segment, video_info.has_audio)
        candidates.append(
            GeneratedCandidate(
                index=segment.index,
                short_path=short_path,
                thumb_path=output_dir / f"thumb_{segment.index:02d}.jpg",
                title_path=output_dir / f"title_{segment.index:02d}.txt",
                thumb_text_path=output_dir / thumb_text_name(segment.index),
                transcript_path=output_dir / short_transcript_text_name(segment.index),
                segments_path=output_dir / short_segments_json_name(segment.index),
            )
        )

    if status_callback:
        status_callback(UI_TEXT["status_creating_thumbs"])
    for segment, candidate in zip(segments, candidates):
        if status_callback:
            status_callback(UI_TEXT["phase_thumb_item"].format(index=segment.index))
        create_thumbnail(ffmpeg_path, input_path, candidate.thumb_path, segment)
        write_title_file(candidate.title_path, segment.index)

    transcript_error = ""
    if create_transcript:
        transcript_failures: list[str] = []
        for candidate in candidates:
            if status_callback:
                status_callback(f"{candidate.index}/3 {UI_TEXT['status_transcribing']}")
            try:
                transcribe_video(candidate.short_path, candidate.transcript_path, candidate.segments_path)
            except Exception as exc:
                candidate.transcript_error = str(exc) or UI_TEXT["status_transcript_failed"]
                transcript_failures.append(candidate.transcript_error)
                if status_callback:
                    status_callback(UI_TEXT["status_transcript_failed"])
        if transcript_failures:
            transcript_error = transcript_failures[0]
        elif status_callback:
            status_callback(UI_TEXT["status_transcript_done"])
    elif status_callback:
        status_callback(UI_TEXT["status_transcript_skipped"])

    if status_callback:
        status_callback(UI_TEXT["status_reviewing"])
    review = write_review_files(
        output_dir,
        candidates,
        use_ollama_review=use_ollama_review,
        ollama_model=ollama_model,
        ollama_url=ollama_url,
    )
    if status_callback:
        status_callback(UI_TEXT["status_review_done"])

    if status_callback:
        status_callback(UI_TEXT["status_preparing_transfer"])
    server = transfer_server or TransferServer()
    transfer_url = server.start(output_dir)
    return ProcessResult(
        output_dir=output_dir,
        video_info=video_info,
        candidates=candidates,
        transfer_url=transfer_url,
        review=review,
        transcript_error=transcript_error,
    )


def open_folder(path: Path) -> None:
    if os.name == "nt":
        os.startfile(str(path))  # type: ignore[attr-defined]
        return
    webbrowser.open(path.resolve().as_uri())


class DakeVideoShortsCutApp:
    def __init__(self) -> None:
        root_class = TkinterDnD.Tk if TkinterDnD is not None else tk.Tk
        self.root = root_class()
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(*WINDOW_MIN_SIZE)
        self.root.configure(bg=THEME["background"])

        self.font_family = choose_font_family(self.root)
        self.root.option_add("*Font", (self.font_family, 10))
        self.transfer_server = TransferServer()
        self.work_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.input_path: Path | None = None
        self.video_info: VideoInfo | None = None
        self.result: ProcessResult | None = None
        self.busy = False
        self.qr_photo = None
        self.footer_compact: bool | None = None
        self.app_config = load_app_config()

        self.status_var = tk.StringVar(value=UI_TEXT["status_ready"])
        self.selected_var = tk.StringVar(value=UI_TEXT["selected_file_empty"])
        self.video_info_var = tk.StringVar(value=UI_TEXT["video_info_empty"])
        self.transfer_url_var = tk.StringVar(value="")
        self.transcript_enabled_var = tk.BooleanVar(value=True)
        self.ollama_review_var = tk.BooleanVar(value=bool(self.app_config.get("use_ollama_review", False)))
        self.recommend_var = tk.StringVar(value=UI_TEXT["recommend_waiting"])
        self.result_vars = [
            tk.StringVar(value=UI_TEXT["result_waiting"]),
            tk.StringVar(value=UI_TEXT["result_waiting"]),
            tk.StringVar(value=UI_TEXT["result_waiting"]),
        ]

        self._apply_window_icon()
        self._configure_styles()
        self._build_ui()
        self._register_drop_targets()
        self._set_status(UI_TEXT["status_ready"])
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(QUEUE_POLL_MS, self._poll_queue)

    def run(self) -> None:
        self.root.mainloop()

    def _configure_styles(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Primary.TButton", padding=(18, 10), font=(self.font_family, 10, "bold"))
        style.configure("Secondary.TButton", padding=(14, 9), font=(self.font_family, 9, "bold"))

    def _build_ui(self) -> None:
        outer = tk.Frame(self.root, bg=THEME["background"])
        outer.pack(fill="both", expand=True, padx=26, pady=(22, 12))

        header = tk.Frame(outer, bg=THEME["background"])
        header.pack(fill="x")
        tk.Label(
            header,
            text=UI_TEXT["main_title"],
            bg=THEME["background"],
            fg=THEME["text"],
            font=(self.font_family, 21, "bold"),
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            header,
            text=UI_TEXT["main_description"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 10),
            anchor="w",
            justify="left",
            wraplength=820,
        ).pack(fill="x", pady=(5, 0))

        body = tk.Frame(outer, bg=THEME["background"])
        body.pack(fill="both", expand=True, pady=(16, 12))
        body.grid_columnconfigure(0, weight=3)
        body.grid_columnconfigure(1, weight=2)
        body.grid_rowconfigure(0, weight=1)

        left = tk.Frame(body, bg=THEME["background"])
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        right = tk.Frame(body, bg=THEME["background"])
        right.grid(row=0, column=1, sticky="nsew")

        self._build_input_panel(left)
        self._build_actions(left)
        self._build_results(left)
        self._build_qr_panel(right)
        self._build_footer(outer)

    def _build_input_panel(self, parent: tk.Frame) -> None:
        panel = self._panel(parent)
        panel.pack(fill="x")

        self.drop_area = tk.Frame(panel, bg="#F8FAFD", highlightbackground=THEME["border"], highlightthickness=1)
        self.drop_area.pack(fill="x", padx=18, pady=(18, 12), ipady=22)
        tk.Label(
            self.drop_area,
            text=UI_TEXT["drop_title"],
            bg="#F8FAFD",
            fg=THEME["text"],
            font=(self.font_family, 16, "bold"),
        ).pack()
        tk.Label(
            self.drop_area,
            text=UI_TEXT["drop_description"],
            bg="#F8FAFD",
            fg=THEME["muted"],
            font=(self.font_family, 9),
        ).pack(pady=(6, 0))
        tk.Button(
            self.drop_area,
            text=UI_TEXT["button_select_file"],
            command=self.select_file,
            bg=THEME["panel"],
            fg=THEME["accent"],
            activebackground="#EEF4FF",
            activeforeground=THEME["accent"],
            relief="solid",
            bd=1,
            padx=18,
            pady=8,
            font=(self.font_family, 9, "bold"),
            cursor="hand2",
        ).pack(pady=(14, 0))

        info = tk.Frame(panel, bg=THEME["panel"])
        info.pack(fill="x", padx=18, pady=(0, 18))
        self._info_row(info, UI_TEXT["selected_file_label"], self.selected_var)
        self._info_row(info, UI_TEXT["video_info_label"], self.video_info_var)
        self.transcript_check = tk.Checkbutton(
            info,
            text=UI_TEXT["checkbox_transcribe"],
            variable=self.transcript_enabled_var,
            bg=THEME["panel"],
            fg=THEME["text"],
            activebackground=THEME["panel"],
            activeforeground=THEME["text"],
            selectcolor=THEME["panel"],
            font=(self.font_family, 9, "bold"),
            anchor="w",
            cursor="hand2",
        )
        self.transcript_check.pack(anchor="w", pady=(6, 0))
        self.ollama_review_check = tk.Checkbutton(
            info,
            text=UI_TEXT["checkbox_ollama_review"],
            variable=self.ollama_review_var,
            bg=THEME["panel"],
            fg=THEME["text"],
            activebackground=THEME["panel"],
            activeforeground=THEME["text"],
            selectcolor=THEME["panel"],
            font=(self.font_family, 9, "bold"),
            anchor="w",
            cursor="hand2",
        )
        self.ollama_review_check.pack(anchor="w", pady=(8, 0))
        tk.Label(
            info,
            text=UI_TEXT["ollama_review_note"],
            bg=THEME["panel"],
            fg=THEME["muted"],
            font=(self.font_family, 8),
            anchor="w",
            justify="left",
            wraplength=560,
        ).pack(fill="x", pady=(3, 0))

    def _info_row(self, parent: tk.Frame, label: str, value_var: tk.StringVar) -> None:
        row = tk.Frame(parent, bg=THEME["panel"])
        row.pack(fill="x", pady=4)
        tk.Label(
            row,
            text=label,
            bg=THEME["panel"],
            fg=THEME["muted"],
            font=(self.font_family, 9, "bold"),
            width=10,
            anchor="w",
        ).pack(side="left")
        tk.Label(
            row,
            textvariable=value_var,
            bg=THEME["panel"],
            fg=THEME["text"],
            font=(self.font_family, 9),
            anchor="w",
            justify="left",
            wraplength=560,
        ).pack(side="left", fill="x", expand=True)

    def _build_actions(self, parent: tk.Frame) -> None:
        actions = tk.Frame(parent, bg=THEME["background"])
        actions.pack(fill="x", pady=(12, 0))
        self.create_button = tk.Button(
            actions,
            text=UI_TEXT["button_create"],
            command=self.create_shorts,
            bg=THEME["accent"],
            fg="#FFFFFF",
            activebackground=THEME["accent_hover"],
            activeforeground="#FFFFFF",
            disabledforeground="#EAF2FF",
            relief="flat",
            bd=0,
            padx=24,
            pady=12,
            font=(self.font_family, 11, "bold"),
            cursor="hand2",
        )
        self.create_button.pack(side="left")
        self.open_output_button = tk.Button(
            actions,
            text=UI_TEXT["button_open_output"],
            command=self.open_output,
            bg=THEME["panel"],
            fg=THEME["text"],
            activebackground="#EEF4FF",
            activeforeground=THEME["text"],
            relief="solid",
            bd=1,
            padx=16,
            pady=10,
            font=(self.font_family, 9, "bold"),
            cursor="hand2",
        )
        self.open_output_button.pack(side="left", padx=(10, 0))

        status_row = tk.Frame(parent, bg=THEME["background"])
        status_row.pack(fill="x", pady=(10, 0))
        tk.Label(
            status_row,
            text=UI_TEXT["status_label"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 9, "bold"),
        ).pack(side="left", padx=(0, 8))
        self.status_badge = tk.Label(
            status_row,
            textvariable=self.status_var,
            bg=THEME["subtle"],
            fg=THEME["muted"],
            font=(self.font_family, 9, "bold"),
            padx=10,
            pady=4,
        )
        self.status_badge.pack(side="left", fill="x")

    def _build_results(self, parent: tk.Frame) -> None:
        panel = self._panel(parent)
        panel.pack(fill="both", expand=True, pady=(12, 0))
        tk.Label(
            panel,
            text=UI_TEXT["result_title"],
            bg=THEME["panel"],
            fg=THEME["text"],
            font=(self.font_family, 13, "bold"),
            anchor="w",
        ).pack(fill="x", padx=18, pady=(16, 8))
        recommend_box = tk.Frame(panel, bg="#F8FAFD", highlightbackground=THEME["border"], highlightthickness=1)
        recommend_box.pack(fill="x", padx=18, pady=(0, 10))
        tk.Label(
            recommend_box,
            textvariable=self.recommend_var,
            bg="#F8FAFD",
            fg=THEME["text"],
            font=(self.font_family, 9, "bold"),
            anchor="w",
            justify="left",
            wraplength=620,
            padx=10,
            pady=9,
        ).pack(fill="x")
        for index, var in enumerate(self.result_vars, start=1):
            row = tk.Frame(panel, bg=THEME["panel"], highlightbackground=THEME["border"], highlightthickness=1)
            row.pack(fill="x", padx=18, pady=(0, 8))
            tk.Label(
                row,
                text=UI_TEXT["candidate_label"].format(number=index),
                bg=THEME["panel"],
                fg=THEME["text"],
                font=(self.font_family, 10, "bold"),
                width=10,
                anchor="w",
                padx=10,
                pady=9,
            ).pack(side="left")
            tk.Label(
                row,
                textvariable=var,
                bg=THEME["panel"],
                fg=THEME["muted"],
                font=(self.font_family, 9),
                anchor="w",
                justify="left",
                wraplength=520,
            ).pack(side="left", fill="x", expand=True, padx=(0, 10))

    def _build_qr_panel(self, parent: tk.Frame) -> None:
        panel = self._panel(parent)
        panel.pack(fill="both", expand=True)
        tk.Label(
            panel,
            text=UI_TEXT["qr_title"],
            bg=THEME["panel"],
            fg=THEME["text"],
            font=(self.font_family, 13, "bold"),
            anchor="w",
        ).pack(fill="x", padx=18, pady=(16, 8))

        self.qr_area = tk.Frame(panel, bg="#F8FAFD", width=220, height=220, highlightbackground=THEME["border"], highlightthickness=1)
        self.qr_area.pack(padx=18, pady=(0, 12))
        self.qr_area.pack_propagate(False)
        self.qr_label = tk.Label(
            self.qr_area,
            text=UI_TEXT["qr_waiting"],
            bg="#F8FAFD",
            fg=THEME["muted"],
            font=(self.font_family, 10, "bold"),
            wraplength=180,
            justify="center",
        )
        self.qr_label.pack(expand=True)

        tk.Label(
            panel,
            text=UI_TEXT["transfer_url_label"],
            bg=THEME["panel"],
            fg=THEME["muted"],
            font=(self.font_family, 9, "bold"),
            anchor="w",
        ).pack(fill="x", padx=18)
        self.url_label = tk.Label(
            panel,
            textvariable=self.transfer_url_var,
            bg=THEME["panel"],
            fg=THEME["accent"],
            font=(self.font_family, 9, "bold"),
            wraplength=300,
            justify="left",
            anchor="w",
            cursor="hand2",
        )
        self.url_label.pack(fill="x", padx=18, pady=(4, 12))
        self.url_label.bind("<Button-1>", lambda _event: self.open_transfer_page())

        self.open_transfer_button = tk.Button(
            panel,
            text=UI_TEXT["button_open_transfer"],
            command=self.open_transfer_page,
            bg=THEME["panel"],
            fg=THEME["text"],
            activebackground="#EEF4FF",
            activeforeground=THEME["text"],
            relief="solid",
            bd=1,
            padx=16,
            pady=9,
            font=(self.font_family, 9, "bold"),
            cursor="hand2",
        )
        self.open_transfer_button.pack(anchor="w", padx=18)

    def _build_footer(self, parent: tk.Frame) -> None:
        self.footer = tk.Frame(parent, bg=THEME["background"])
        self.footer.pack(fill="x", pady=(2, 0))
        self.footer_left = tk.Frame(self.footer, bg=THEME["background"])
        self._make_footer_text(self.footer_left, UI_TEXT["footer_left"], bold=True)
        self._make_footer_text(self.footer_left, UI_TEXT["footer_separator"])
        self._make_footer_text(self.footer_left, UI_TEXT["footer_tagline"])

        self.footer_right = tk.Frame(self.footer, bg=THEME["background"])
        self._make_footer_link(self.footer_right, UI_TEXT["footer_link_1"], LINKS["assessment"])
        self._make_footer_text(self.footer_right, UI_TEXT["footer_separator"])
        self._make_footer_link(self.footer_right, UI_TEXT["footer_link_2"], LINKS["instagram"])
        self._make_footer_text(self.footer_right, UI_TEXT["footer_separator"])
        self._make_footer_text(self.footer_right, UI_TEXT["footer_copyright"])
        self.root.bind("<Configure>", self._update_footer_layout, add="+")
        self._update_footer_layout()

    def _make_footer_text(self, parent: tk.Frame, label: str, bold: bool = False) -> None:
        tk.Label(
            parent,
            text=label,
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 8, "bold" if bold else "normal"),
        ).pack(side="left")

    def _make_footer_link(self, parent: tk.Frame, label: str, url: str) -> None:
        link = tk.Label(
            parent,
            text=label,
            bg=THEME["background"],
            fg=THEME["link"],
            font=(self.font_family, 8, "bold"),
            cursor="hand2",
        )
        link.pack(side="left")
        link.bind("<Button-1>", lambda _event: webbrowser.open(url, new=2))
        link.bind("<Enter>", lambda _event: link.configure(fg=THEME["link_hover"]))
        link.bind("<Leave>", lambda _event: link.configure(fg=THEME["link"]))

    def _update_footer_layout(self, _event=None) -> None:
        compact = self.root.winfo_width() < 900
        if compact == self.footer_compact:
            return
        self.footer_compact = compact
        self.footer_left.pack_forget()
        self.footer_right.pack_forget()
        if compact:
            self.footer_left.pack(anchor="center", pady=(0, 2))
            self.footer_right.pack(anchor="center")
            return
        self.footer_left.pack(side="left")
        self.footer_right.pack(side="right")

    def _panel(self, parent: tk.Frame) -> tk.Frame:
        return tk.Frame(
            parent,
            bg=THEME["panel"],
            highlightbackground=THEME["border"],
            highlightthickness=1,
        )

    def _apply_window_icon(self) -> None:
        for icon_path in get_common_icon_candidates():
            try:
                resolved = icon_path.resolve()
            except OSError:
                resolved = icon_path
            if not resolved.exists():
                continue
            try:
                self.root.iconbitmap(str(resolved))
            except tk.TclError:
                pass
            return

    def _register_drop_targets(self) -> None:
        if DND_FILES is None:
            return
        targets = [self.root, self.drop_area]
        for target in targets:
            if not hasattr(target, "drop_target_register"):
                continue
            try:
                target.drop_target_register(DND_FILES)
                target.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                continue

    def _on_drop(self, event) -> None:
        try:
            paths = self.root.tk.splitlist(event.data)
        except Exception:
            paths = [event.data]
        if not paths:
            return
        self.load_file(Path(paths[0]))

    def select_file(self) -> None:
        filename = filedialog.askopenfilename(
            title=UI_TEXT["button_select_file"],
            filetypes=[
                (UI_TEXT["filetype_mp4"], "*.mp4"),
                (UI_TEXT["filetype_all"], "*.*"),
            ],
        )
        if filename:
            self.load_file(Path(filename))

    def load_file(self, path: Path) -> None:
        try:
            path = validate_mp4_path(path)
        except UserFacingError as exc:
            self._show_error(str(exc))
            return
        self.input_path = path
        self.result = None
        self.video_info = None
        self.selected_var.set(str(path))
        self.video_info_var.set(UI_TEXT["status_checking"])
        self.transfer_url_var.set("")
        self._clear_results()
        self._set_status(UI_TEXT["status_checking"])
        worker = threading.Thread(target=self._probe_worker, args=(path,), daemon=True)
        worker.start()

    def _probe_worker(self, path: Path) -> None:
        try:
            _ffmpeg_path, ffprobe_path = ensure_required_tools()
            info = probe_video(path, ffprobe_path)
            self.work_queue.put(("probe_done", info))
        except Exception as exc:
            self.work_queue.put(("error", str(exc)))

    def create_shorts(self) -> None:
        if self.busy:
            return
        if self.input_path is None:
            self._show_error(UI_TEXT["error_no_file"])
            return
        self.busy = True
        self._set_button_states()
        self._clear_results()
        self._set_status(UI_TEXT["status_checking"])
        create_transcript = bool(self.transcript_enabled_var.get())
        use_ollama_review = bool(self.ollama_review_var.get())
        self.app_config["use_ollama_review"] = use_ollama_review
        save_app_config(self.app_config)
        worker = threading.Thread(
            target=self._process_worker,
            args=(self.input_path, create_transcript, use_ollama_review),
            daemon=True,
        )
        worker.start()

    def _process_worker(self, path: Path, create_transcript: bool, use_ollama_review: bool) -> None:
        try:
            result = process_video(
                path,
                self._queue_status,
                self.transfer_server,
                create_transcript=create_transcript,
                use_ollama_review=use_ollama_review,
                ollama_model=str(self.app_config.get("ollama_model") or DEFAULT_OLLAMA_MODEL),
                ollama_url=str(self.app_config.get("ollama_url") or DEFAULT_OLLAMA_URL),
            )
            self.work_queue.put(("process_done", result))
        except Exception as exc:
            self.work_queue.put(("process_error", str(exc)))

    def _queue_status(self, message: str) -> None:
        self.work_queue.put(("status", message))

    def _poll_queue(self) -> None:
        try:
            while True:
                event_type, payload = self.work_queue.get_nowait()
                if event_type == "status":
                    self._set_status(str(payload))
                elif event_type == "probe_done" and isinstance(payload, VideoInfo):
                    self.video_info = payload
                    self.video_info_var.set(self._format_video_info(payload))
                    self._set_status(UI_TEXT["status_ready"])
                elif event_type == "process_done" and isinstance(payload, ProcessResult):
                    self._handle_process_done(payload)
                elif event_type == "process_error":
                    self.busy = False
                    self._set_button_states()
                    self._show_error(str(payload))
                elif event_type == "error":
                    self.video_info_var.set(UI_TEXT["video_info_empty"])
                    self._show_error(str(payload))
        except queue.Empty:
            pass
        self.root.after(QUEUE_POLL_MS, self._poll_queue)

    def _handle_process_done(self, result: ProcessResult) -> None:
        self.busy = False
        self.result = result
        self.video_info = result.video_info
        self.video_info_var.set(self._format_video_info(result.video_info))
        self.transfer_url_var.set(result.transfer_url)
        self._refresh_results(result)
        self._show_qr(result.transfer_url)
        if result.transcript_error:
            self._set_status(UI_TEXT["status_transcript_failed"], error=True)
        elif result.review is not None and result.review.review_method == "local_fallback":
            self._set_status(UI_TEXT["status_ollama_fallback"], success=True)
        else:
            self._set_status(UI_TEXT["status_complete"], success=True)
        self._set_button_states()
        messagebox.showinfo(UI_TEXT["dialog_done_title"], UI_TEXT["dialog_done_message"])

    def _format_video_info(self, info: VideoInfo) -> str:
        return UI_TEXT["video_info_template"].format(
            duration=format_seconds(info.duration),
            width=info.width,
            height=info.height,
            fps=f"{info.fps:.2f}" if info.fps else "0",
            audio=UI_TEXT["audio_yes"] if info.has_audio else UI_TEXT["audio_no"],
        )

    def _clear_results(self) -> None:
        self.recommend_var.set(UI_TEXT["recommend_waiting"])
        for var in self.result_vars:
            var.set(UI_TEXT["result_waiting"])
        self.qr_photo = None
        self.qr_label.configure(image="", text=UI_TEXT["qr_waiting"], fg=THEME["muted"])

    def _refresh_results(self, result: ProcessResult) -> None:
        best = get_recommended_review(result.candidates, result.review)
        if best is None:
            self.recommend_var.set(UI_TEXT["recommend_waiting"])
        else:
            method = result.review.review_method if result.review is not None else "local"
            self.recommend_var.set(
                UI_TEXT["recommend_card_template"].format(
                    number=best.candidate,
                    score=best.score,
                    method=review_method_label(method),
                    title=best.title,
                    thumbnail_text=best.thumbnail_text,
                    reason=best.reason,
                )
            )
        for index, candidate in enumerate(result.candidates):
            score = candidate.review.score if candidate.review is not None else "--"
            values = [
                f"{UI_TEXT['result_score']}: {score}",
                f"{UI_TEXT['result_short']}: {self._exists_text(candidate.short_path)}",
                f"{UI_TEXT['result_thumb']}: {self._exists_text(candidate.thumb_path)}",
                f"{UI_TEXT['result_title_file']}: {self._exists_text(candidate.title_path)}",
                f"{UI_TEXT['result_thumb_text']}: {self._exists_text(candidate.thumb_text_path)}",
                f"{UI_TEXT['result_transcript']}: {self._transcript_exists_text(candidate)}",
            ]
            self.result_vars[index].set(" / ".join(values))

    def _exists_text(self, path: Path) -> str:
        return UI_TEXT["result_exists"] if path.exists() else UI_TEXT["result_missing"]

    def _transcript_exists_text(self, candidate: GeneratedCandidate) -> str:
        if candidate.transcript_path.exists() and candidate.segments_path.exists():
            return UI_TEXT["result_exists"]
        return UI_TEXT["result_transcript_missing"]

    def _show_qr(self, url: str) -> None:
        try:
            self.qr_photo = make_qr_photo(url)
        except UserFacingError:
            self.qr_photo = None
            self.qr_label.configure(image="", text=UI_TEXT["qr_dependency_missing"], fg=THEME["muted"])
            return
        self.qr_label.configure(image=self.qr_photo, text="")

    def open_output(self) -> None:
        if self.result is None:
            self._show_error(UI_TEXT["error_output_missing"])
            return
        try:
            open_folder(self.result.output_dir)
        except Exception:
            self._show_error(UI_TEXT["error_open_output"])

    def open_transfer_page(self) -> None:
        url = self.transfer_url_var.get().strip()
        if not url:
            self._show_error(UI_TEXT["error_open_transfer"])
            return
        try:
            webbrowser.open(url, new=2)
        except Exception:
            self._show_error(UI_TEXT["error_open_transfer"])

    def _set_status(self, message: str, error: bool = False, success: bool = False) -> None:
        self.status_var.set(message)
        if error:
            self.status_badge.configure(bg=THEME["error_bg"], fg=THEME["error"])
        elif success:
            self.status_badge.configure(bg=THEME["success_bg"], fg=THEME["success"])
        else:
            self.status_badge.configure(bg=THEME["subtle"], fg=THEME["muted"])

    def _show_error(self, message: str) -> None:
        self._set_status(UI_TEXT["status_error"], error=True)
        messagebox.showerror(UI_TEXT["dialog_error_title"], message)

    def _set_button_states(self) -> None:
        state = tk.DISABLED if self.busy else tk.NORMAL
        self.create_button.configure(state=state, bg="#A9C0F7" if self.busy else THEME["accent"])
        self.transcript_check.configure(state=state)
        self.ollama_review_check.configure(state=state)

    def _on_close(self) -> None:
        self.app_config["use_ollama_review"] = bool(self.ollama_review_var.get())
        save_app_config(self.app_config)
        self.transfer_server.shutdown()
        self.root.destroy()


def check_qr_dependency() -> str:
    try:
        import qrcode  # noqa: F401
        from PIL import Image, ImageTk  # noqa: F401
    except Exception:
        return "missing"
    return "available"


def run_launch_check() -> int:
    long_segments = build_segments(180.0)
    if len(long_segments) != SHORT_COUNT or any(segment.duration > SHORT_MAX_SECONDS for segment in long_segments):
        raise RuntimeError("segment fixture failed")
    short_segments = build_segments(30.0)
    if len(short_segments) != SHORT_COUNT or any(segment.duration <= 0 for segment in short_segments):
        raise RuntimeError("short segment fixture failed")
    with tempfile.TemporaryDirectory() as temp_dir:
        output_dir = Path(temp_dir)
        candidates: list[GeneratedCandidate] = []
        for index in range(1, SHORT_COUNT + 1):
            candidate = GeneratedCandidate(
                index=index,
                short_path=output_dir / f"short_{index:02d}.mp4",
                thumb_path=output_dir / f"thumb_{index:02d}.jpg",
                title_path=output_dir / f"title_{index:02d}.txt",
                thumb_text_path=output_dir / thumb_text_name(index),
                transcript_path=output_dir / short_transcript_text_name(index),
                segments_path=output_dir / short_segments_json_name(index),
            )
            candidates.append(candidate)
            candidate.short_path.write_bytes(b"dummy")
            candidate.thumb_path.write_bytes(b"dummy")
            candidate.title_path.write_text(
                UI_TEXT["title_candidate"].format(number=index) + "\n",
                encoding="utf-8",
            )
            transcript = write_transcript_files(
                candidate.transcript_path,
                candidate.segments_path,
                [
                    {
                        "start": 12.4,
                        "end": 18.9,
                        "start_text": "00:12",
                        "end_text": "00:18",
                        "text": "AIでDAKEを作る。これはいいね。",
                    }
                ],
            )
            if not transcript.text_path.exists() or not transcript.segments_path.exists():
                raise RuntimeError("transcript fixture failed")
            if "00:12" not in transcript.text_path.read_text(encoding="utf-8"):
                raise RuntimeError("transcript text fixture failed")
        if (output_dir / "transcript.txt").exists() or (output_dir / "transcript_segments.json").exists():
            raise RuntimeError("global transcript fixture failed")
        review = write_review_files(output_dir, candidates)
        if not review.text_path.exists() or not review.json_path.exists() or not get_best_review(candidates):
            raise RuntimeError("review fixture failed")
        if not all(candidate.thumb_text_path.exists() for candidate in candidates):
            raise RuntimeError("thumb text fixture failed")
        review_payload = load_review_payload(output_dir)
        if review_payload.get("review_method") != "local" or "reviews" not in review_payload:
            raise RuntimeError("review payload fixture failed")
        page = build_transfer_html(output_dir)
        if (
            "short_01.mp4" not in page
            or thumb_text_name(1) not in page
            or short_transcript_text_name(1) not in page
            or REVIEW_TEXT_FILE not in page
            or UI_TEXT["review_method_label"] not in page
            or UI_TEXT["mobile_recommend_section"] not in page
            or UI_TEXT["mobile_development_section"] not in page
            or UI_TEXT["mobile_score"] not in page
            or UI_TEXT["mobile_title"] not in page
        ):
            raise RuntimeError("transfer html fixture failed")

        nested_parent = output_dir / f"{OUTPUT_FOLDER_PREFIX}_fixture"
        nested_parent.mkdir()
        nested_output = create_output_dir(nested_parent / "short_01.mp4", now=datetime(2026, 1, 2, 3, 4, 5))
        if nested_output.parent != output_dir:
            raise RuntimeError("nested output fixture failed")

    ffmpeg_status = "available" if find_tool("ffmpeg") else "missing"
    ffprobe_status = "available" if find_tool("ffprobe") else "missing"
    print(UI_TEXT["launch_check_ok"])
    print(UI_TEXT["launch_check_ffmpeg"].format(status=ffmpeg_status))
    print(UI_TEXT["launch_check_ffprobe"].format(status=ffprobe_status))
    print(UI_TEXT["launch_check_segments"])
    print(UI_TEXT["launch_check_html"])
    print(UI_TEXT["launch_check_transcript"])
    print(UI_TEXT["launch_check_review"])
    print(UI_TEXT["launch_check_qr"].format(status=check_qr_dependency()))
    return 0


def run_process_check(input_file: str, create_transcript: bool = True, use_ollama_review: bool = False) -> int:
    server = TransferServer()
    try:
        result = process_video(
            Path(input_file),
            transfer_server=server,
            create_transcript=create_transcript,
            use_ollama_review=use_ollama_review,
        )
        print(UI_TEXT["process_check_done"].format(output_dir=result.output_dir))
        print(result.transfer_url)
        try:
            from urllib.parse import urlparse
            from urllib.request import urlopen

            parsed_url = urlparse(result.transfer_url)
            local_url = f"http://127.0.0.1:{parsed_url.port or TRANSFER_START_PORT}/"
            with urlopen(local_url, timeout=5) as response:
                body = response.read(24000).decode("utf-8", errors="replace")
            print(f"transfer_page={'short_01.mp4' in body}")
            transfer_has_transcript = all(
                candidate.transcript_path.exists() and candidate.transcript_path.name in body
                for candidate in result.candidates
            )
            print(f"transfer_transcript={transfer_has_transcript}")
            transfer_has_review = REVIEW_TEXT_FILE in body and UI_TEXT["mobile_score"] in body
            print(f"transfer_review={transfer_has_review}")
            print(f"transfer_thumb_text={thumb_text_name(1) in body}")
        except Exception:
            print("transfer_page=False")
            print("transfer_transcript=False")
            print("transfer_review=False")
            print("transfer_thumb_text=False")
        for candidate in result.candidates:
            print(candidate.short_path.name, candidate.short_path.exists())
            print(candidate.thumb_path.name, candidate.thumb_path.exists())
            print(candidate.title_path.name, candidate.title_path.exists())
            print(candidate.thumb_text_path.name, candidate.thumb_text_path.exists())
            print(candidate.transcript_path.name, candidate.transcript_path.exists())
            print(candidate.segments_path.name, candidate.segments_path.exists())
            print(f"review_score_{candidate.index}={candidate.review.score if candidate.review else 'missing'}")
        transcript_ok = all(
            candidate.transcript_path.exists() and candidate.segments_path.exists()
            for candidate in result.candidates
        )
        review_ok = result.review is not None and result.review.text_path.exists() and result.review.json_path.exists()
        thumb_text_ok = all(candidate.thumb_text_path.exists() for candidate in result.candidates)
        best = get_recommended_review(result.candidates, result.review)
        print(f"global_transcript={(result.output_dir / 'transcript.txt').exists()}")
        print(UI_TEXT["process_check_transcript"].format(status="created" if transcript_ok else "missing"))
        print(UI_TEXT["process_check_review"].format(status="created" if review_ok else "missing"))
        print(f"review_method={result.review.review_method if result.review else 'missing'}")
        print(f"ollama_error={result.review.ollama_error if result.review and result.review.ollama_error else ''}")
        print(f"shorts_review.txt={(result.output_dir / REVIEW_TEXT_FILE).exists()}")
        print(f"shorts_review.json={(result.output_dir / REVIEW_JSON_FILE).exists()}")
        print(f"thumb_text={thumb_text_ok}")
        print(f"review_best={best.candidate if best else 'missing'}")
        print(f"nested_output={not result.output_dir.parent.name.startswith(OUTPUT_FOLDER_PREFIX)}")
        if result.transcript_error:
            print(f"transcript_error={result.transcript_error}")
    finally:
        server.shutdown()
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=APP_NAME)
    parser.add_argument("--launch-check", action="store_true")
    parser.add_argument("--process-check", metavar="MP4")
    parser.add_argument("--skip-transcript", action="store_true")
    parser.add_argument("--ollama-review", action="store_true")
    args = parser.parse_args(argv)
    if args.launch_check:
        return run_launch_check()
    if args.process_check:
        return run_process_check(
            args.process_check,
            create_transcript=not args.skip_transcript,
            use_ollama_review=args.ollama_review,
        )
    app = DakeVideoShortsCutApp()
    app.run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
