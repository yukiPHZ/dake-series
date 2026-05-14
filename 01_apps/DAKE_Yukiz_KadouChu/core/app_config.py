from __future__ import annotations

import json
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

APP_KEY = "DAKE_Yukiz_KadouChu"
APP_NAME = "Dakeユキズ稼働中"
WINDOW_TITLE = "ユキズ稼働中"
EXE_NAME = "DakeYukiz_KadouChu.exe"
COPYRIGHT = "© 2026 しまリス不動産 / Vibe-Coded by Yukihiko Kikuta"

DEFAULT_CONFIG = {
    "whisper_model": "base",
    "ollama_model": "",
    "preview_clip_seconds": 45,
    "prefer_nvenc": True,
}

LOG_TEXT = {
    "startup": "補助脳：起動しました。",
    "running": "稼働中。",
    "complete": "整っています。",
    "beside": "側に。",
    "source_detected": "補助脳：素材を検出しました。",
    "system_check_start": "補助脳：制作基地を確認しています。",
    "system_check_complete": "補助脳：制作基地を確認しました。",
    "tools_missing": "補助脳：不足している道具があります。",
    "cli_disconnected": "補助脳：外部CLIが未接続です。",
    "install_candidates_ready": "補助脳：インストール候補を作成しました。",
    "install_commands_copied": "補助脳：インストール候補をクリップボードへ送信しました。",
    "install_guide_opened": "補助脳：CLI導入補助を開きました。",
    "first_test_start": "補助脳：初回動画テストを開始します。",
    "first_test_probe_ready": "補助脳：ffprobeで動画情報を取得しました。",
    "first_test_probe_failed": "補助脳：ffprobe情報取得に失敗しました。",
    "first_test_ffprobe_skipped": "補助脳：FFprobeがないため動画情報取得をスキップしました。",
    "first_test_ffmpeg_required": "補助脳：FFmpegが必要です。CLI導入補助を確認してください。",
    "first_test_nvenc_try": "補助脳：NVENCでテストエンコードを試します。",
    "first_test_nvenc_ready": "補助脳：NVENCでテストクリップを作成しました。",
    "first_test_nvenc_fallback": "補助脳：NVENCで試しましたが、CPUへ切り替えました。",
    "first_test_cpu_try": "補助脳：NVENCは未使用でCPUエンコードを試します。",
    "first_test_clip_ready": "補助脳：テストクリップを作成しました。",
    "first_test_clip_failed": "補助脳：テストクリップ作成に失敗しました。",
    "gpu_encode_ready": "補助脳：GPUエンコードの準備ができています。",
    "github_unauthorized": "補助脳：GitHub CLIはありますが、認証が必要です。",
    "wrangler_unauthorized": "補助脳：Wranglerはありますが、認証が必要です。",
    "ollama_sleeping": "補助脳：ローカル補助脳はまだ眠っています。",
    "package_prepare": "補助脳：制作パッケージを準備しています。",
    "media_probe": "補助脳：メディア情報を取得しています。",
    "ffprobe_missing": "補助脳：FFprobeが見つかりません。解析を一部スキップします。",
    "quiet_scene_search": "補助脳：静かな場面を探しています。",
    "shorts_extracted": "補助脳：Shorts候補を抽出しました。",
    "preview_created": "補助脳：プレビュークリップを作成しました。",
    "preview_skipped": "補助脳：プレビュー作成をスキップしました。",
    "metadata_ready": "補助脳：投稿用メタデータ雛形を整えました。",
    "process_stopped": "補助脳：処理を停止しました。",
    "whisper_missing": "補助脳：faster-whisperが見つかりません。文字起こしをスキップします。",
    "whisper_loading": "補助脳：文字起こしモデルを読み込んでいます。",
    "whisper_saved": "補助脳：文字起こしを保存しました。",
    "whisper_unavailable": "補助脳：文字起こしが利用できませんでした。",
    "posting_package_start": "補助脳：投稿用パッケージを作成しています。",
    "posting_media_ready": "補助脳：投稿用の動画情報を保存しました。",
    "posting_media_unavailable": "補助脳：動画情報は取得できませんでしたが、処理を続けます。",
    "posting_shorts_ready": "補助脳：Shorts候補を投稿用に整えました。",
    "posting_titles_ready": "補助脳：タイトル案を整えました。",
    "posting_ollama_fallback": "補助脳：ローカル補助脳の応答がないため、固定テンプレートで整えました。",
    "posting_ready": "補助脳：出せる形になりました。",
    "review_package_read": "補助脳：投稿パッケージを読みました。",
    "review_atmosphere": "補助脳：動画の空気を整理しています。",
    "review_shorts": "補助脳：Shorts候補を見直しました。",
    "review_created": "補助脳：レビューを作成しました。",
    "review_template_fallback": "補助脳：ローカル補助脳の応答がないため、固定レビューで整えました。",
}

SHORTS_REASON_TEXT = {
    "speech": "補助脳：発話と作業音が安定しています。",
    "even": "補助脳：動画尺から均等に抽出しました。",
    "duration_unknown": "補助脳：動画尺未取得のため、確認用の仮候補です。",
}


def app_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def data_dir() -> Path:
    return app_root() / "data"


def outputs_dir() -> Path:
    return data_dir() / "outputs"


def logs_dir() -> Path:
    return data_dir() / "logs"


def config_path() -> Path:
    return data_dir() / "config.json"


def ensure_app_dirs() -> None:
    directories = [
        app_root() / "assets",
        data_dir(),
        data_dir() / "inbox",
        data_dir() / "bgm",
        data_dir() / "templates",
        data_dir() / "templates" / "thumbnails",
        data_dir() / "templates" / "logos",
        outputs_dir(),
        logs_dir(),
    ]
    for directory in directories:
        directory.mkdir(parents=True, exist_ok=True)


def load_config() -> dict[str, Any]:
    ensure_app_dirs()
    path = config_path()
    if not path.exists():
        return DEFAULT_CONFIG.copy()
    try:
        loaded = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return DEFAULT_CONFIG.copy()
    config = DEFAULT_CONFIG.copy()
    if isinstance(loaded, dict):
        for key, value in loaded.items():
            if key in DEFAULT_CONFIG and key not in {"api_key", "token", "secret"}:
                config[key] = value
    return config


def save_config(config: dict[str, Any]) -> None:
    ensure_app_dirs()
    safe_config = {key: value for key, value in config.items() if key in DEFAULT_CONFIG}
    config_path().write_text(json.dumps(safe_config, ensure_ascii=False, indent=2), encoding="utf-8")


def safe_project_name(source_name: str) -> str:
    stem = Path(source_name).stem or "project"
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", stem)
    cleaned = re.sub(r"\s+", "_", cleaned).strip("._ ")
    if not cleaned:
        cleaned = "project"
    cleaned = cleaned[:64]
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{stamp}_{cleaned}"


def format_duration(seconds: float | int | None) -> str:
    total = max(0, int(seconds or 0))
    hours, remainder = divmod(total, 3600)
    minutes, secs = divmod(remainder, 60)
    if hours:
        return f"{hours:02d}:{minutes:02d}:{secs:02d}"
    return f"{minutes:02d}m {secs:02d}s"


def seconds_to_timecode(seconds: float | int | None) -> str:
    total = max(0, int(seconds or 0))
    hours, remainder = divmod(total, 3600)
    minutes, secs = divmod(remainder, 60)
    return f"{hours:02d}:{minutes:02d}:{secs:02d}"


def human_size(size: int | float | None) -> str:
    value = float(size or 0)
    units = ["B", "KB", "MB", "GB", "TB"]
    for unit in units:
        if value < 1024 or unit == units[-1]:
            if unit == "B":
                return f"{int(value)} {unit}"
            return f"{value:.2f} {unit}"
        value /= 1024
    return f"{value:.2f} TB"


def estimate_processing_seconds(duration: float, whisper_available: bool) -> float:
    if duration <= 0:
        return 90
    transcription_factor = 0.65 if whisper_available else 0.08
    packaging_factor = 0.18
    base = 45
    return max(60, base + duration * (transcription_factor + packaging_factor))
