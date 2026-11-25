# -*- coding: utf-8 -*-
from __future__ import annotations

from pathlib import Path


APP_ID = "dake_voice_log"
APP_NAME = "Dake声ログ"
WINDOW_TITLE = "Dake声ログ"
DISPLAY_NAME = "声ログ"
EXE_NAME = "DakeVoiceLog.exe"
APP_DATA_DIR_NAME = "DakeVoiceLog"
USER_DATA_DIR_NAME = "Dake声ログ"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"
SUMMARY_MODEL = "qwen3:1.7b"
SUMMARY_MODEL_SIZE_LABEL = "約1.4GB"
OLLAMA_DOWNLOAD_URL = "https://ollama.com/download/windows"
OLLAMA_BASE_URL = "http://localhost:11434"
TRANSCRIPTION_MODELS = {
    "light": "base",
    "standard": "small",
}

UI_TEXT = {
    "brand_series": "シンプルそれDAKEシリーズ",
    "header_title": APP_NAME,
    "header_subtitle": "声を、あとから振り返れる記録に。",
    "privacy_title": "Dake声ログの準備",
    "privacy_message": "Dake声ログは、マイクの音声を検知して、このPC内に記録を保存します。\n\n録音を行う場所のルールや、必要な同意・許可については、利用環境に応じて確認してください。\n\n音声待機はいつでも一時停止できます。",
    "privacy_start": "使い始める",
    "privacy_later_summary": "日次要約はあとで設定できます。",
    "setup_step_mic": "マイク",
    "setup_step_save": "保存先",
    "setup_step_transcription": "文字起こしAI",
    "setup_step_summary": "日次要約",
    "setup_ready": "準備を進めています",
    "setup_summary_question": "日次要約を使いますか？",
    "setup_summary_use": "日次要約を使う",
    "setup_summary_later": "あとで設定する",
    "capture_title": "音声待機",
    "status_initializing": "準備中",
    "status_waiting": "待機中",
    "status_recording": "録音中",
    "status_paused": "一時停止中",
    "status_mic_error": "マイクを確認しています",
    "status_stopped": "停止中",
    "capture_since": "{time}から記録しています",
    "capture_mic": "マイク：{name}",
    "capture_no_mic": "マイク：未確認",
    "processing_title": "バックグラウンド処理",
    "processing_idle": "処理待ちなし",
    "processing_transcribing": "文字起こし中",
    "processing_summarizing": "日次要約を作成中",
    "processing_error": "確認が必要です",
    "today_title": "今日のログ",
    "today_recorded": "録音",
    "today_transcribed": "文字起こし",
    "today_pending": "処理待ち",
    "today_summary": "日次要約",
    "summary_not_created": "未作成",
    "summary_ready": "作成済み",
    "recent_title": "最近の記録",
    "recent_empty": "まだ記録はありません",
    "recent_processing": "処理中",
    "recent_done": "文字起こし済み",
    "recent_failed": "未完了",
    "button_pause": "一時停止",
    "button_resume": "再開",
    "button_today_log": "今日のログを見る",
    "button_open_folder": "保存フォルダを開く",
    "button_settings": "設定",
    "button_close": "閉じる",
    "button_save": "保存",
    "button_cancel": "キャンセル",
    "button_select_folder": "選択",
    "button_retry": "再確認",
    "button_summarize_pending": "未処理日を要約",
    "button_open_text": "テキストを見る",
    "button_open_audio": "音声を開く",
    "button_open_ollama": "公式ページを開く",
    "button_prepare_summary_ai": "要約AIを準備",
    "button_install_check": "インストール後に再確認",
    "dialog_error_title": "確認してください",
    "dialog_info_title": "お知らせ",
    "dialog_mic_error": "マイクに接続できませんでした。\n接続状態を確認するか、設定から別のマイクを選んでください。",
    "dialog_save_error": "音声を保存できませんでした。\n保存先の空き容量とアクセス権を確認してください。",
    "dialog_transcribe_error": "文字起こしを完了できませんでした。\n音声ファイルは保存されています。あとから再処理できます。",
    "dialog_ollama_missing": "日次要約AIがまだ準備されていません。\n録音と文字起こしはそのまま利用できます。",
    "dialog_tray_notice": "Dake声ログは通知領域で音声待機を続けています。\n終了する場合は、トレイメニューの「終了」を選んでください。",
    "settings_title": "設定",
    "settings_audio": "音声",
    "settings_mic": "使用マイク",
    "settings_default_mic": "既定のマイク",
    "settings_level": "入力レベル",
    "settings_sensitivity": "感度",
    "sensitivity_low": "低い",
    "sensitivity_standard": "標準",
    "sensitivity_high": "高い",
    "settings_silence": "録音終了までの無音",
    "silence_10": "10秒",
    "silence_30": "30秒",
    "silence_60": "60秒",
    "settings_transcription": "文字起こし",
    "settings_quality": "品質",
    "quality_light": "軽量",
    "quality_standard": "標準",
    "settings_model_status": "モデル状態",
    "model_ready": "準備済み",
    "model_missing": "未準備",
    "model_checking": "確認中",
    "settings_summary": "日次要約",
    "settings_summary_enabled": "使用する",
    "settings_summary_disabled": "使用しない",
    "settings_ai_status": "AI状態",
    "settings_save": "保存",
    "settings_save_root": "保存先",
    "settings_auto_start": "Windows起動時にDake声ログを開始する",
    "settings_auto_delete": "音声自動削除",
    "auto_delete_off": "OFF",
    "auto_delete_30": "30日",
    "auto_delete_90": "90日",
    "log_viewer_title": "ログを見る",
    "log_viewer_date": "日付",
    "log_viewer_empty": "この日の記録はありません",
    "tray_open": "Dake声ログを開く",
    "tray_pause": "音声待機を一時停止",
    "tray_resume": "音声待機を再開",
    "tray_today": "今日のログを開く",
    "tray_folder": "保存フォルダを開く",
    "tray_settings": "設定",
    "tray_exit": "終了",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_link_1": "戸建買取査定",
    "footer_link_2": "Instagram",
    "footer_separator": " ｜ ",
    "footer_copyright": COPYRIGHT,
    "ollama_setup_title": "日次要約AIの準備",
    "ollama_setup_message": "日次要約には、ローカルAI環境が必要です。\n\n音声と文字データをこのPC内で処理するため、Ollamaのセットアップを行います。",
    "ollama_model_message": "要約AIを準備します。\n約1.4GBのデータをダウンロードします。",
    "ollama_model_pull_confirm": "要約AIを準備します。\n約1.4GBのデータをダウンロードします。\n\n開始しますか？",
    "ollama_model_pulling": "要約AIを準備中",
    "status_ollama_missing": "Ollama未設定",
    "status_ollama_ready": "要約AI 準備済み",
    "status_ollama_model_missing": "要約AI 未準備",
    "window_minimum_size": "860x700",
}

THEME = {
    "background": "#F6F7F9",
    "card": "#FFFFFF",
    "text": "#1E2430",
    "muted": "#667085",
    "border": "#E6EAF0",
    "accent": "#2F6FED",
    "accent_hover": "#2458BF",
    "selection_bg": "#EAF2FF",
    "selection_border": "#7AA7FF",
    "success": "#12B76A",
    "success_bg": "#EAFBF3",
    "warning": "#F79009",
    "warning_bg": "#FFF6E8",
    "error": "#D92D20",
    "error_bg": "#FDECEC",
    "soft": "#EEF2F7",
    "link": "#58677D",
}

LINKS = {
    "assessment": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "instagram": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

FONT_CANDIDATES = ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo")


def app_source_dir() -> Path:
    return Path(__file__).resolve().parents[1]
