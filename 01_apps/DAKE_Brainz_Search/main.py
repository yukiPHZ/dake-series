# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import json
import queue
import subprocess
import sys
import tempfile
import threading
import time
import zipfile
from pathlib import Path


APP_NAME = "BRAINZ"
WINDOW_TITLE = "BRAINZ - 記憶庫 / 取り込み母艦"
COPYRIGHT = "QPSC — Quiet Personal Cognitive System by Yukihiko Kikuta"

UI_TEXT = {
    "app_title": APP_NAME,
    "window_title": WINDOW_TITLE,
    "copyright": COPYRIGHT,
    "subtitle": "記憶庫 / 取り込み母艦",
    "role_summary": "保存 / 取り込み / 設定 / 状態",
    "brainz_status_title": "BRAINZ状態",
    "status_brainz_standby": "待機しています",
    "status_brainz_awake": "起きています",
    "status_brainz_awake_detail": "記憶庫は起きています",
    "status_memory_path": "記憶フォルダ: {path}",
    "status_last_import_none": "最終取り込み: まだありません",
    "status_last_import": "最終取り込み: {source}",
    "last_import_chatgpt_export": "ChatGPT export",
    "last_import_codex_result": "Codex結果",
    "last_import_codex_report": "Codex報告",
    "last_import_slack": "Slack Inbox",
    "last_import_aru": "Aru Inbox",
    "button_open_oikawa": "OIKAWAを開く",
    "button_settings_entry": "接続 / 設定",
    "section_import_entry": "取り込み入口",
    "section_settings_entry": "接続 / 設定",
    "settings_entry_hint": "Slack / Queue / Codex報告の入口を下にまとめています。",
    "section_oikawa_bridge": "OIKAWAへ",
    "oikawa_bridge_message": "検索・原本表示・熾火・ORBITはOIKAWAで扱います。",
    "search_bridge_title": "検索入口（補助）",
    "search_bridge_helper": "読む場所はOIKAWAへ移しています。",
    "search_placeholder": "うろ覚えで検索",
    "embers_search_placeholder": "何を探したいかわからなくても大丈夫です。",
    "embers_search_helper": "忘れていた熱を探します。",
    "tab_search": "Search",
    "tab_embers": "熾火",
    "button_search": "Search",
    "button_searching": "Searching...",
    "button_read": "読む",
    "checkbox_semantic_search": "Semantic Search",
    "button_index": "Index",
    "button_cancel": "Cancel",
    "button_choose": "Choose",
    "button_import_chatgpt_export": "ChatGPT export取込",
    "chatgpt_import_title": "ChatGPT exportを取り込む",
    "chatgpt_import_helper": "zip / フォルダ / conversations.json に対応",
    "chatgpt_import_note": "取り込むとBRAINZに保存され、OIKAWAで読めます。",
    "chatgpt_import_waiting": "選んだexportを静かに取り込みます。",
    "chatgpt_import_importing": "ChatGPT exportを取り込んでいます。",
    "button_import_chatgpt_zip": "zip",
    "button_import_chatgpt_folder": "フォルダ",
    "button_import_chatgpt_json": "conversations.json",
    "button_import_codex_result": "Codex結果取込",
    "button_chatgpt": "ChatGPTまとめ",
    "button_codex": "Codex素材",
    "memory_title": "記憶庫 / 取り込み入口",
    "label_watch_folder": "Watch Folder",
    "checkbox_auto_index": "Auto Index",
    "index_title": "Index Status",
    "system_title": "System Status",
    "section_notifications": "Notifications",
    "checkbox_enable_notifications": "Enable Notifications",
    "section_remote_queue": "Remote Queue",
    "checkbox_enable_remote_queue": "Enable Remote Queue",
    "button_choose_remote_queue": "Choose Queue",
    "section_codex_reports": "Codex Reports",
    "section_slack_inbox": "Slack Inbox",
    "checkbox_enable_slack_inbox": "Enable Slack Inbox",
    "label_slack_token": "Bot Token",
    "label_slack_channel": "Channel ID",
    "label_slack_interval": "Poll sec 5-15",
    "button_save_slack": "Save Slack",
    "section_aru_inbox": "Aru Inbox",
    "checkbox_enable_aru_inbox": "Enable Aru Inbox",
    "label_aru_token": "Aru Slack Token",
    "label_aru_channel": "Aru Channel ID",
    "button_save_aru": "Save Aru",
    "results_title": "検索結果（補助）",
    "embers_results_title": "熾火",
    "source_view_title": "原本確認（補助）",
    "preview_title": "詳細（補助）",
    "tags_title": "Related Tags",
    "section_related_memory": "Related Memory",
    "section_memory_flow": "Memory Flow（OIKAWAへ）",
    "label_memory_flow": "関連する記憶を辿っています。",
    "handoff_title": "Handoff Summary",
    "log_title": "BRAINZ Log",
    "empty_memory": "未選択",
    "empty_watch_folder": "未選択",
    "empty_remote_queue_folder": "未選択",
    "empty_results": "検索結果はまだありません",
    "empty_preview": "結果を選ぶと詳細が表示されます",
    "empty_source_view": "読む記憶を選んでください",
    "empty_tags": "タグ候補なし",
    "related_memory_empty": "関連記憶なし",
    "memory_flow_empty": "記憶の流れはまだありません",
    "empty_handoff": "検索結果から引き継ぎ素材を生成できます",
    "choose_memory_title": "記憶フォルダを選択",
    "dialog_title": "補助脳BRAINZ",
    "dialog_memory_missing": "記憶フォルダを選択してください",
    "dialog_no_results": "検索結果がありません",
    "dialog_export_done": "ファイルを生成しました",
    "dialog_select_chatgpt_export": "ChatGPT export zipを選択（キャンセルでフォルダ選択）",
    "dialog_select_chatgpt_export_folder": "展開済みChatGPT exportフォルダを選択",
    "dialog_select_chatgpt_conversations_json": "conversations.jsonを選択",
    "dialog_import_codex_result": "Codex結果を取り込む",
    "dialog_select_codex_result_file": "Codex結果ファイルを選択",
    "label_codex_result_input": "Codexの完了報告・修正結果・commit/push結果を貼り付け",
    "button_import_pasted_codex": "貼り付けを取り込む",
    "button_import_codex_file": "txt / md を選択",
    "button_close": "閉じる",
    "button_timeline_ascending": "Old -> New",
    "button_timeline_descending": "New -> Old",
    "filetype_zip": "zipファイル",
    "filetype_json": "jsonファイル",
    "filetype_text_markdown": "txt / md",
    "filetype_all": "すべてのファイル",
    "status_importing_chatgpt": "CHATGPT EXPORT IMPORTING...",
    "status_chatgpt_import_complete": "CHATGPT EXPORT IMPORT COMPLETE",
    "chatgpt_import_result": "ChatGPT exportを取り込みました。{conversations}件の会話 / {messages}件の記憶を保存しました。",
    "chatgpt_import_result_no_new": "ChatGPT exportを確認しました。新しく保存する記憶はありませんでした。",
    "chatgpt_import_failed_short": "取り込みできませんでした。",
    "status_importing_codex": "CODEX RESULT IMPORTING...",
    "status_codex_import_complete": "CODEX RESULT IMPORT COMPLETE",
    "error_conversations_json_not_found": "conversations.json が見つかりませんでした。zipまたは展開済みフォルダを確認してください。",
    "error_chatgpt_import_failed": "ChatGPT exportを取り込めませんでした。",
    "error_codex_result_empty": "Codex結果が空です。テキストを貼り付けるか、txt / md ファイルを選択してください。",
    "error_codex_import_failed": "Codex結果を取り込めませんでした。",
    "log_chatgpt_export_detected": "ChatGPT export detected.",
    "log_conversations_json_found": "conversations.json found: {path}",
    "log_chatgpt_import_complete": "{conversations} conversations imported. {messages} messages indexed. Skipped duplicates: {skipped}. Errors: {errors}.",
    "log_chatgpt_memory_imported": "補助脳：ChatGPTの記憶を取り込みました。",
    "log_chatgpt_import_file": "IMPORT LOG: {path}",
    "log_codex_result_detected": "Codex result detected.",
    "log_codex_commit_found": "Commit hash found: {commit_hash}",
    "log_codex_import_complete": "Changed files: {changed_files}. Skipped duplicate: {skipped}.",
    "log_codex_memory_imported": "補助脳：Codexの実装結果を記憶しました。",
    "log_codex_import_file": "IMPORT LOG: {path}",
    "status_codex_report_watching": "CODEX REPORTS: WATCHING",
    "status_codex_report_imported": "CODEX REPORT IMPORTED",
    "status_codex_report_failed": "CODEX REPORT FAILED",
    "status_codex_report_idle": "CODEX REPORTS: WAITING",
    "log_codex_report_detected": "CODEX REPORT AUTO IMPORT: {path}",
    "log_codex_report_imported": "Codex report imported: commit={commit_hash} changed_files={changed_files} skipped={skipped}",
    "log_codex_report_failed": "Codex report failed: {path} :: {error}",
    "phrase_codex_report_saved": "補助脳：Codex報告を記憶しました。",
    "label_last_commit": "LAST COMMIT: {commit_hash}",
    "label_last_report": "LAST REPORT: {path}",
    "status_slack_connected": "SLACK: CONNECTED",
    "status_slack_auth_failed": "SLACK AUTH FAILED",
    "status_slack_timeout": "SLACK TIMEOUT",
    "status_slack_import_complete": "SLACK IMPORT COMPLETE",
    "status_slack_channel_not_found": "CHANNEL NOT FOUND",
    "status_slack_config_missing": "SLACK CONFIG MISSING",
    "status_slack_disabled": "SLACK: OFF",
    "status_slack_ready": "SLACK: READY",
    "status_slack_error": "SLACK ERROR",
    "status_slack_config_saved": "SLACK CONFIG SAVED",
    "status_slack_config_save_failed": "SLACK CONFIG SAVE FAILED",
    "status_slack_task_detected": "SLACK TASK DETECTED",
    "status_slack_task_processed": "SLACK TASK PROCESSED",
    "status_slack_task_failed": "SLACK TASK FAILED",
    "status_aru_connected": "ARU: CONNECTED",
    "status_aru_import_complete": "ARU IMPORT COMPLETE",
    "status_aru_config_missing": "ARU CONFIG MISSING",
    "status_aru_auth_failed": "ARU AUTH FAILED",
    "status_aru_channel_not_found": "ARU CHANNEL NOT FOUND",
    "status_aru_timeout": "ARU TIMEOUT",
    "status_aru_disabled": "ARU: OFF",
    "status_aru_ready": "ARU: READY",
    "status_aru_error": "ARU ERROR",
    "status_aru_config_saved": "ARU CONFIG SAVED",
    "status_aru_config_save_failed": "ARU CONFIG SAVE FAILED",
    "label_slack_last_import": "LAST IMPORT: {time}",
    "label_slack_channel_status": "CHANNEL: {channel}",
    "label_aru_last_import": "ARU LAST IMPORT: {time}",
    "label_aru_channel_status": "ARU CHANNEL: {channel}",
    "label_last_task": "LAST TASK: {task}",
    "log_slack_import": "SLACK INBOX: New message imported. imported={imported} skipped={skipped} failed={failed}",
    "log_slack_status": "SLACK INBOX: {status} {message}",
    "log_slack_file": "SLACK LOG: {path}",
    "log_slack_config_saved": "SLACK CONFIG SAVED: enabled={enabled} channel={channel} interval={interval}s",
    "log_slack_config_save_failed": "SLACK CONFIG SAVE FAILED: missing={missing}",
    "log_slack_config_reloaded": "SLACK CONFIG RELOADED: enabled={enabled} channel={channel} interval={interval}s",
    "log_aru_import": "ARU INBOX: New fragment imported. imported={imported} skipped={skipped} failed={failed}",
    "log_aru_status": "ARU INBOX: {status} {message}",
    "log_aru_file": "ARU LOG: {path}",
    "log_aru_config_saved": "ARU CONFIG SAVED: enabled={enabled} channel={channel}",
    "log_aru_config_save_failed": "ARU CONFIG SAVE FAILED: missing={missing}",
    "log_aru_config_reloaded": "ARU CONFIG RELOADED: enabled={enabled} channel={channel}",
    "log_slack_task_detected": "SLACK TASK: {task_type} / {query}",
    "log_slack_task_processed": "Slack task processed: status={status} changed={changed} duplicate={duplicate}",
    "log_slack_task_failed": "Slack task failed: {task_type} / {error}",
    "phrase_slack_memory_saved": "\u88dc\u52a9\u8133\uff1aSlack Inbox\u3092\u53d6\u308a\u8fbc\u307f\u307e\u3057\u305f\u3002",
    "phrase_slack_task_received": "\u88dc\u52a9\u8133\uff1aSlack task\u3092\u53d7\u4fe1\u3057\u307e\u3057\u305f\u3002",
    "phrase_aru_memory_saved": "補助脳：在る断片を記憶しました。",
    "log_embers_search": "EMBERS: {query}",
    "status_markdown_loading": "READING SOURCE...",
    "status_markdown_loaded": "SOURCE LOADED",
    "embers_card_meta": "{updated_at} | {source_type} | {terms}",
    "log_semantic_search_initialized": "Semantic search initialized.",
    "log_embedding_generated": "Embedding generated.",
    "log_related_memory_found": "Related memory found: {count}",
    "log_semantic_disabled": "FTS only mode: {message}",
    "log_memory_flow_generated": "Memory Flow generated.",
    "log_memory_flow_updated": "Related timeline updated: {count}",
    "log_memory_flow_file": "MEMORY FLOW LOG: {path}",
    "phrase_memory_reconnected": "補助脳：記憶の流れを接続しました。",
    "codex_source_paste": "Codex result paste",
    "smoke_markdown": "# quiet workflow\n\n補助脳BRAINZは静かな青の検索脳。Codexに投げたやつを忘れない。",
    "smoke_text": "DAKEのGitルール: git status, add, commit, push。UIを止めない。",
    "smoke_json_name": "補助脳BRAINZ",
    "smoke_query_memory": "静かな青 Codexに投げたやつ",
    "smoke_query_git": "DAKEのGitルール",
    "smoke_json_role": "local memory bridge",
    "smoke_chatgpt_zip_title": "補助脳BRAINZの話",
    "smoke_chatgpt_folder_title": "Cloudflare Pages整理",
    "smoke_chatgpt_user_text": "ChatGPT export zipから補助脳BRAINZの記憶を取り込みたい。",
    "smoke_chatgpt_assistant_text": "BRAINZは会話タイトル、role、本文をsource_type=chatgpt_exportとして検索できます。",
    "smoke_chatgpt_folder_text": "展開済みフォルダのconversations.jsonもローカルで解析します。",
    "smoke_chatgpt_query": "source_type chatgpt_export 補助脳BRAINZ",
    "smoke_codex_title": "Add Codex result import to Brainz",
    "smoke_codex_text": "# Add Codex result import to Brainz\n\n完了しました。\n\n主な変更:\n- 追加: core/codex_importer.py\n- 更新: main.py, core/db.py, core/handoff_writer.py, README.md\n\n確認結果:\n- python main.py --launch-check: OK\n- python main.py --smoke-test: OK\n- build.bat: OK\n- dist/DakeBrainz_Search.exe --launch-check: OK\n\nGit:\n- commit: `abc1234def567890abc1234def567890abc1234d`\n- push: `origin/main` 成功\n- git status clean\n\nPhase 4候補: 関連タグ自動生成",
    "smoke_codex_file_text": "Codex実装結果\n\n変更済みファイル: core/search_engine.py README.md\n確認済みテスト: smoke-test OK\ncommit: `def5678abc1234def5678abc1234def5678abc12`\npush: origin/main 成功",
    "smoke_codex_query": "Add Codex result import Brainz abc1234",
    "smoke_codex_chatgpt_heading": "過去のCodex実装結果",
    "smoke_codex_codex_heading": "既存実装履歴",
    "log_ready": "READY: SQLite memory layer initialized",
    "log_folder_set": "MEMORY FOLDER: {folder}",
    "log_index_started": "INDEX START: {folder}",
    "log_index_cancel_requested": "INDEX CANCEL REQUESTED",
    "log_index_done": "INDEX DONE: indexed={indexed}, skipped={skipped}, errors={errors}",
    "log_index_cancelled": "INDEX CANCELLED: indexed={indexed}, skipped={skipped}, errors={errors}",
    "log_index_error": "INDEX ERROR: {error}",
    "log_search": "SEARCH: {query} -> {count} results",
    "log_searching": "SEARCH: searching... {query}",
    "log_search_complete": "SEARCH COMPLETE: {count} results",
    "log_no_results": "NO RESULTS: {query}",
    "phrase_searching_memory": "補助脳：関連する記憶を探しています。",
    "phrase_no_results": "補助脳：関連する記憶はまだ見つかりません。",
    "log_watch_initialized": "Watch folder initialized.",
    "log_watching_folder": "Watching: {folder}",
    "log_new_memory_detected": "New memory detected: {path}",
    "log_auto_index_complete": "Auto indexing complete.",
    "log_watch_file": "WATCH LOG: {path}",
    "log_auto_index_queued": "AUTO INDEX QUEUED: index is already running",
    "phrase_memory_updated": "補助脳：記憶を更新しました。",
    "log_notification_sent": "NOTIFY: {message}",
    "log_remote_queue_detected": "REMOTE QUEUE: Task detected: {task_type} / {query}",
    "log_remote_queue_processed": "Remote queue processed: processed={processed}, failed={failed}",
    "log_remote_queue_failed": "Remote queue failed: {path} :: {error}",
    "phrase_remote_queue_received": "補助脳：遠隔キューを受け取りました。",
    "log_export": "EXPORT: {path}",
    "log_oikawa_open": "OIKAWA OPEN: {mode} / {path}",
    "log_oikawa_missing": "OIKAWA NOT FOUND",
    "oikawa_mode_exe": "dist exe",
    "oikawa_mode_python": "python main.py",
    "log_settings_entry": "SETTINGS ENTRY: Slack / Queue / Codex reports",
    "notify_title": "BRAINZ",
    "notify_oikawa_opened": "OIKAWAを開きます。",
    "notify_oikawa_missing": "OIKAWAが見つかりません。buildまたは配置を確認してください。",
    "notify_memory_detected": "新しい記憶を検出しました。",
    "notify_auto_index_complete": "記憶を更新しました。",
    "notify_semantic_updated": "Semantic Searchを更新しました。",
    "notify_chatgpt_import_complete": "ChatGPTの記憶を取り込みました。",
    "notify_codex_import_complete": "Codex結果を記憶しました。",
    "notify_codex_report_imported": "Codex報告を記憶しました。",
    "notify_codex_report_failed": "Codex報告の取り込みに失敗しました。",
    "notify_slack_import_complete": "Slack Inboxを取り込みました。",
    "notify_slack_import_failed": "Slack Inboxの取得に失敗しました。",
    "notify_slack_task_received": "Slack taskを受信しました。",
    "notify_aru_import_complete": "Aru Inboxを取り込みました。",
    "notify_aru_import_failed": "Aru Inboxの取得に失敗しました。",
    "notify_handoff_complete": "{kind} handoffを生成しました。",
    "notify_remote_queue_detected": "Remote Queueを受け取りました。",
    "notify_remote_queue_processed": "Remote Queueを処理しました。",
    "notify_remote_queue_failed": "Remote Queueの処理に失敗しました。",
    "index_idle": "IDLE",
    "index_running": "RUNNING {current}/{total}",
    "index_embedding": "EMBEDDING CHUNKS...",
    "index_done": "DONE {indexed} indexed / {skipped} skipped / {errors} errors",
    "searching": "SEARCHING...",
    "status_searching": "SEARCHING...",
    "status_search_complete": "SEARCH COMPLETE / {count} results",
    "status_no_results": "NO RESULTS",
    "status_watching": "WATCHING...",
    "status_auto_index_on": "AUTO INDEX: ON",
    "status_auto_index_off": "AUTO INDEX: OFF",
    "status_new_memory_detected": "NEW MEMORY DETECTED",
    "status_last_memory": "LAST MEMORY: {path}",
    "status_last_index": "LAST INDEX: {time}",
    "status_watch_folder_missing": "WATCH FOLDER NOT SET",
    "status_remote_queue_off": "REMOTE QUEUE: OFF",
    "status_remote_queue_watching": "REMOTE QUEUE: WATCHING",
    "status_remote_queue_detected": "REMOTE QUEUE: DETECTED",
    "status_remote_queue_processed": "REMOTE QUEUE: PROCESSED",
    "status_remote_queue_failed": "REMOTE QUEUE: FAILED",
    "status_remote_queue_folder_missing": "QUEUE FOLDER NOT SET",
    "status_remote_queue_counts": "Pending {pending} / Processed {processed} / Failed {failed}",
    "status_sqlite_ready": "SQLITE READY",
    "status_sqlite_error": "SQLITE ERROR",
    "status_cuda_online": "CUDA ONLINE",
    "status_cuda_unavailable": "CUDA UNAVAILABLE",
    "status_gpu_detected": "GPU DETECTED: {name}",
    "status_gpu_missing": "GPU NOT DETECTED",
    "status_ollama_ready": "OLLAMA LOCAL READY",
    "status_ollama_not_running": "OLLAMA NOT RUNNING",
    "status_embedding_ready": "EMBEDDING READY",
    "status_embedding_unavailable": "EMBEDDING UNAVAILABLE",
    "status_semantic_search_ready": "SEMANTIC SEARCH READY",
    "status_semantic_search_disabled": "SEMANTIC SEARCH DISABLED",
    "status_memory_flow_ready": "MEMORY FLOW READY",
    "status_memory_flow_loading": "MEMORY FLOW...",
    "status_docs": "DOCS {documents} / CHUNKS {chunks}",
    "preview_template": "PATH: {path}\nSOURCE: {source_type}\nLABEL: {source_label}\nCONVERSATION: {conversation_title}\nROLE: {role}\nMESSAGE INDEX: {message_index}\nCOMMIT: {commit_hash}\nCHANGED FILES: {changed_files}\nTEST RESULTS: {test_results}\nBUILD RESULTS: {build_results}\nPUSH: {push_result}\nGIT STATUS: {git_status}\nMODIFIED: {modified_at}\nINDEXED: {indexed_at}\nSCORE: {score:.2f}\nSEMANTIC: {semantic_score:.2f}\n\n{content}",
    "source_view_template": "PATH: {path}\nSOURCE: {source_type}\nTS: {source_created_at}\n\n---\n\n{content}",
    "result_meta": "{source_type} | score {score:.1f}",
    "result_meta_chatgpt": "{source_label} | score {score:.1f}",
    "result_meta_codex": "commit {commit_hash} | score {score:.1f}",
    "result_meta_semantic": "semantic {semantic_score:.2f}",
    "timeline_meta": "{date} [{source_type}] flow {flow_score:.1f}{semantic_text}",
    "timeline_semantic": " / semantic {semantic_score:.2f}",
    "handoff_preview": "query: {query}\nresults: {count}\n\n{items}",
    "launch_check_ok": "LAUNCH CHECK OK",
}


def resolve_oikawa_launch_target(brainz_app_dir: Path | None = None) -> tuple[Path | None, str]:
    app_dir = brainz_app_dir or Path(__file__).resolve().parent
    oikawa_dir = app_dir.parent / "DAKE_Brainz_OIKAWA"
    exe_path = oikawa_dir / "dist" / "DakeBrainz_OIKAWA.exe"
    if exe_path.exists():
        return exe_path, UI_TEXT["oikawa_mode_exe"]
    main_path = oikawa_dir / "main.py"
    if main_path.exists():
        return main_path, UI_TEXT["oikawa_mode_python"]
    return None, ""


def resolve_oikawa_launch_command(brainz_app_dir: Path | None = None) -> tuple[list[str], Path | None, str]:
    target, mode = resolve_oikawa_launch_target(brainz_app_dir)
    if target is None:
        return [], None, ""
    if target.suffix.lower() == ".py":
        return [sys.executable, str(target)], target, mode
    return [str(target)], target, mode


def run_smoke_test() -> int:
    import requests

    from core.app_config import AppConfig, ConfigStore, ensure_app_dirs
    from core.chatgpt_importer import import_chatgpt_export
    from core.codex_importer import import_codex_file, import_codex_text
    from core.codex_report_auto import count_pending_reports, process_codex_reports_folder
    from core.db import BrainzDatabase
    from core.gpu_checker import check_gpu
    from core.handoff_writer import write_chatgpt_handoff, write_codex_handoff
    from core.indexer import Indexer
    from core.notifications import NotificationQueue
    from core.ollama_client import check_ollama
    from core.ollama_embeddings import DEFAULT_EMBED_MODEL, check_embedding_status
    from core.qpsc_status import AWAKE_STATUS_MESSAGE, write_brainz_awake_status
    from core.remote_queue import count_pending_queue_files, process_remote_queue_folder
    from core.search_engine import SearchEngine
    from core.slack_inbox import poll_slack_inbox, slack_ts_float
    from core.watch_folder import detect_changed_files

    ensure_app_dirs()
    database = BrainzDatabase()
    database.ensure_schema()

    with tempfile.TemporaryDirectory(prefix="brainz_memory_") as tmp:
        root = Path(tmp)
        default_config = ConfigStore(root / "default_config.json").load()
        if not default_config.enable_notifications:
            raise RuntimeError("notifications should default to enabled")
        config_store = ConfigStore(root / "config.json")
        config_store.save(
            AppConfig(
                memory_folder=str(root),
                watch_folder=str(root),
                auto_index_enabled=True,
                remote_queue_folder=str(root / "remote_queue"),
                enable_remote_queue=True,
                auto_run_remote_search=False,
                codex_reports_folder=str(root / "codex_reports"),
                enable_slack_inbox=True,
                slack_bot_token="xoxb-smoke-token",
                slack_channel_id="C123SMOKE",
                slack_poll_interval_seconds=5,
                slack_last_ts="0",
                enable_aru_inbox=True,
                aru_slack_token="xoxb-aru-token",
                aru_channel_id="CARUSMOKE",
                aru_poll_interval_seconds=5,
                aru_last_ts="0",
                enable_notifications=False,
            )
        )
        config_data = config_store.load()
        if config_data.enable_notifications or not config_data.auto_index_enabled:
            raise RuntimeError("notification config did not roundtrip")
        if not config_data.enable_remote_queue or config_data.auto_run_remote_search:
            raise RuntimeError("remote queue config did not roundtrip")
        if config_data.codex_reports_folder != str(root / "codex_reports"):
            raise RuntimeError("codex reports config did not roundtrip")
        if (
            not config_data.enable_slack_inbox
            or config_data.slack_bot_token != "xoxb-smoke-token"
            or config_data.slack_channel_id != "C123SMOKE"
            or config_data.slack_poll_interval_seconds != 5
        ):
            raise RuntimeError("slack inbox config did not roundtrip")
        if (
            not config_data.enable_aru_inbox
            or config_data.aru_slack_token != "xoxb-aru-token"
            or config_data.aru_channel_id != "CARUSMOKE"
            or config_data.aru_poll_interval_seconds != 5
        ):
            raise RuntimeError("aru inbox config did not roundtrip")
        qpsc_status = write_brainz_awake_status(
            path=root / "qpsc_brainz_status.json",
            started_at="2026-05-19T00:00:00",
        )
        if (
            qpsc_status.get("brainz_awake") is not True
            or qpsc_status.get("status_message") != AWAKE_STATUS_MESSAGE
            or not qpsc_status.get("last_heartbeat_at")
        ):
            raise RuntimeError("qpsc awake status did not roundtrip")
        oikawa_command, oikawa_target, oikawa_mode = resolve_oikawa_launch_command()
        if not oikawa_command or oikawa_target is None or not oikawa_target.exists():
            raise RuntimeError("oikawa launch target was not resolved")
        if oikawa_mode not in {UI_TEXT["oikawa_mode_exe"], UI_TEXT["oikawa_mode_python"]}:
            raise RuntimeError("oikawa launch mode was not resolved")
        fake_apps = root / "fake_apps"
        fake_brainz = fake_apps / "DAKE_Brainz_Search"
        fake_oikawa = fake_apps / "DAKE_Brainz_OIKAWA"
        fake_dist = fake_oikawa / "dist"
        fake_brainz.mkdir(parents=True)
        command, target, mode = resolve_oikawa_launch_command(fake_brainz)
        if command or target is not None or mode:
            raise RuntimeError("missing oikawa fallback did not return quiet missing state")
        fake_oikawa.mkdir(parents=True)
        fake_main = fake_oikawa / "main.py"
        fake_main.write_text("print('ok')\n", encoding="utf-8")
        command, target, mode = resolve_oikawa_launch_command(fake_brainz)
        if not command or target != fake_main or mode != UI_TEXT["oikawa_mode_python"]:
            raise RuntimeError("oikawa python fallback was not selected")
        fake_dist.mkdir(parents=True)
        fake_exe = fake_dist / "DakeBrainz_OIKAWA.exe"
        fake_exe.write_text("", encoding="utf-8")
        command, target, mode = resolve_oikawa_launch_command(fake_brainz)
        if not command or target != fake_exe or mode != UI_TEXT["oikawa_mode_exe"]:
            raise RuntimeError("oikawa dist exe was not prioritized")
        parsed_interval, interval_valid = parse_slack_poll_interval("9")
        if parsed_interval != 9 or not interval_valid:
            raise RuntimeError("slack interval parser rejected a valid value")
        parsed_interval, interval_valid = parse_slack_poll_interval("bad")
        if parsed_interval != 10 or interval_valid:
            raise RuntimeError("slack interval parser did not flag invalid input")
        missing_fields = slack_config_missing_fields("", "", "", False)
        if set(missing_fields) != {"memory_folder", "bot_token", "channel_id", "poll_interval"}:
            raise RuntimeError("slack missing field detection failed")

        notification_queue = NotificationQueue(history_limit=3)
        for index in range(4):
            notification_queue.push("BRAINZ", f"notification {index}")
        popped = []
        while notification_queue.pending_count:
            item = notification_queue.pop()
            if item:
                notification_queue.remember(item)
                popped.append(item.message)
        if len(popped) != 4 or len(notification_queue.history) != 3:
            raise RuntimeError("notification queue failed")

        (root / "ideas").mkdir(parents=True, exist_ok=True)
        (root / "ideas" / "quiet_blue.md").write_text(
            UI_TEXT["smoke_markdown"],
            encoding="utf-8",
        )
        (root / "codex_result.txt").write_text(
            UI_TEXT["smoke_text"],
            encoding="utf-8",
        )
        (root / "spec.json").write_text(
            json.dumps({"name": UI_TEXT["smoke_json_name"], "role": UI_TEXT["smoke_json_role"]}, ensure_ascii=False),
            encoding="utf-8",
        )

        cancel_event = threading.Event()
        final_progress = Indexer(database).run(root, cancel_event)
        if final_progress.errors:
            raise RuntimeError(f"index errors: {final_progress.errors}")

        watch_new_md = root / "ideas" / "auto_watch.md"
        watch_new_txt = root / "auto_watch.txt"
        watch_new_json = root / "auto_watch.json"
        watch_new_md.write_text("Watch Folder auto index memory flow", encoding="utf-8")
        watch_new_txt.write_text("Auto Index detects text memory", encoding="utf-8")
        watch_new_json.write_text(json.dumps({"watch": "auto index"}, ensure_ascii=False), encoding="utf-8")
        watch_detection = detect_changed_files(database, root)
        if len(watch_detection.changed_files) < 3:
            raise RuntimeError("watch folder did not detect new txt/md/json files")
        auto_progress = Indexer(database).run(root, threading.Event())
        if auto_progress.errors:
            raise RuntimeError(f"auto index errors: {auto_progress.errors}")
        watch_new_md.write_text("Watch Folder auto index memory flow updated", encoding="utf-8")
        watch_update = detect_changed_files(database, root)
        if str(watch_new_md.resolve()) not in watch_update.changed_files:
            raise RuntimeError("watch folder did not detect updated md file")
        auto_update_progress = Indexer(database).run(root, threading.Event())
        if auto_update_progress.errors:
            raise RuntimeError(f"auto update index errors: {auto_update_progress.errors}")

        remote_queue = root / "remote_queue"
        remote_queue.mkdir()
        import_target = root / "remote_import.md"
        import_target.write_text("Remote Queue import target quiet workflow", encoding="utf-8")
        (remote_queue / "search.md").write_text(
            "# BRAINZ TASK\n\ntype: search\nquery: quiet workflow\nnote: ChatGPT handoff用",
            encoding="utf-8",
        )
        (remote_queue / "note.txt").write_text(
            "type: note\nnote: スマホから置いた補助脳メモ\n\nRemote Queue note memory",
            encoding="utf-8",
        )
        (remote_queue / "handoff_chatgpt.json").write_text(
            json.dumps({"type": "handoff_chatgpt", "query": "quiet workflow", "note": "ChatGPT handoff用"}, ensure_ascii=False),
            encoding="utf-8",
        )
        (remote_queue / "handoff_codex.txt").write_text(
            "handoff_codex: DAKE Git rules\nnote: Codex素材候補",
            encoding="utf-8",
        )
        (remote_queue / "import.json").write_text(
            json.dumps({"type": "import", "path": str(import_target)}, ensure_ascii=False),
            encoding="utf-8",
        )
        (remote_queue / "missing_import.json").write_text(
            json.dumps({"type": "import", "path": str(root / "missing.md")}, ensure_ascii=False),
            encoding="utf-8",
        )
        if count_pending_queue_files(remote_queue) != 6:
            raise RuntimeError("remote queue pending count failed")
        remote_result = process_remote_queue_folder(database, remote_queue)
        if remote_result.processed < 5 or remote_result.failed < 1:
            raise RuntimeError("remote queue processing failed")
        if not (remote_queue / "processed").exists() or not (remote_queue / "failed").exists():
            raise RuntimeError("remote queue did not create processed/failed folders")
        queue_stats = database.remote_queue_stats()
        if queue_stats["processed"] < 5 or queue_stats["failed"] < 1:
            raise RuntimeError("remote queue history was not saved")
        (remote_queue / "search.md").write_text("search: quiet workflow collision", encoding="utf-8")
        collision_result = process_remote_queue_folder(database, remote_queue)
        collision_names = [Path(item.destination_file).name for item in collision_result.results]
        if not any(name.startswith("search_") and name.endswith(".md") for name in collision_names):
            raise RuntimeError("remote queue collision timestamp failed")

        unique_suffix = str(time.time_ns())
        zip_export = root / "chatgpt_export_zip"
        folder_export = root / "chatgpt_export_folder"
        zip_export.mkdir()
        folder_export.mkdir()

        zip_conversations = sample_conversations(
            f"brainz_zip_{unique_suffix}",
            UI_TEXT["smoke_chatgpt_zip_title"],
            UI_TEXT["smoke_chatgpt_user_text"],
            UI_TEXT["smoke_chatgpt_assistant_text"],
        )
        folder_conversations = sample_conversations(
            f"brainz_folder_{unique_suffix}",
            UI_TEXT["smoke_chatgpt_folder_title"],
            UI_TEXT["smoke_chatgpt_folder_text"],
            UI_TEXT["smoke_chatgpt_assistant_text"],
        )
        (zip_export / "conversations.json").write_text(json.dumps(zip_conversations, ensure_ascii=False), encoding="utf-8")
        (folder_export / "conversations.json").write_text(json.dumps(folder_conversations, ensure_ascii=False), encoding="utf-8")

        zip_path = root / "chatgpt_export.zip"
        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.write(zip_export / "conversations.json", "chatgpt_export/conversations.json")

        zip_result = import_chatgpt_export(zip_path, database)
        folder_result = import_chatgpt_export(folder_export, database)
        duplicate_result = import_chatgpt_export(zip_path, database)

        if zip_result.messages_indexed < 2:
            raise RuntimeError("zip import did not index messages")
        if folder_result.messages_indexed < 2:
            raise RuntimeError("folder import did not index messages")
        if duplicate_result.skipped_duplicates < 2:
            raise RuntimeError("duplicate import was not skipped")

        unique_commit_a = (f"{time.time_ns():x}" * 4)[:40]
        unique_commit_b = (f"{time.time_ns() + 1:x}" * 4)[:40]
        codex_text = UI_TEXT["smoke_codex_text"].replace(
            "abc1234def567890abc1234def567890abc1234d",
            unique_commit_a,
        )
        codex_file_text = UI_TEXT["smoke_codex_file_text"].replace(
            "def5678abc1234def5678abc1234def5678abc12",
            unique_commit_b,
        )
        codex_query = f"{UI_TEXT['smoke_codex_query']} {unique_commit_a[:7]}"
        codex_result = import_codex_text(codex_text, database, UI_TEXT["codex_source_paste"])
        codex_duplicate = import_codex_text(codex_text, database, UI_TEXT["codex_source_paste"])
        codex_file = root / "codex_result.md"
        codex_file.write_text(codex_file_text, encoding="utf-8")
        codex_file_result = import_codex_file(codex_file, database)
        if not codex_result.changed:
            raise RuntimeError("codex paste import did not index")
        if not codex_duplicate.skipped_duplicate:
            raise RuntimeError("codex duplicate import was not skipped")
        if not codex_file_result.changed:
            raise RuntimeError("codex file import did not index")

        unique_commit_c = (f"{time.time_ns() + 2:x}" * 4)[:40]
        unique_commit_d = (f"{time.time_ns() + 3:x}" * 4)[:40]
        codex_reports = root / "codex_reports"
        codex_reports.mkdir()
        auto_report_md = codex_text.replace(unique_commit_a, unique_commit_c).replace(
            "Add Codex result import to Brainz",
            "Add codex report auto import to Brainz",
        )
        auto_report_txt = codex_file_text.replace(unique_commit_b, unique_commit_d)
        (codex_reports / "auto_report.md").write_text(auto_report_md, encoding="utf-8")
        (codex_reports / "auto_report.txt").write_text(auto_report_txt, encoding="utf-8")
        (codex_reports / "empty.md").write_text("", encoding="utf-8")
        if count_pending_reports(codex_reports) != 3:
            raise RuntimeError("codex report pending count failed")
        codex_report_result = process_codex_reports_folder(database, codex_reports)
        if codex_report_result.imported < 2 or codex_report_result.failed < 1:
            raise RuntimeError("codex report auto import failed")
        if not (codex_reports / "processed").exists() or not (codex_reports / "failed").exists():
            raise RuntimeError("codex report processed/failed folders missing")
        (codex_reports / "auto_report.md").write_text(auto_report_md, encoding="utf-8")
        codex_report_duplicate = process_codex_reports_folder(database, codex_reports)
        duplicate_item = next((item for item in codex_report_duplicate.items if item.status == "processed"), None)
        if duplicate_item is None or not duplicate_item.skipped_duplicate:
            raise RuntimeError("codex report duplicate was not skipped")

        class FakeSlackResponse:
            def __init__(self, payload: dict[str, object], status_code: int = 200) -> None:
                self.payload = payload
                self.status_code = status_code

            def json(self) -> dict[str, object]:
                return self.payload

        class FakeSlackSession:
            def __init__(self, mode: str = "ok", messages: list[dict[str, object]] | None = None) -> None:
                self.mode = mode
                self.messages = messages

            def get(self, url: str, headers=None, params=None, timeout=None):  # noqa: ANN001
                if self.mode == "timeout":
                    raise requests.Timeout("slack timeout")
                if self.mode == "auth_failed":
                    return FakeSlackResponse({"ok": False, "error": "invalid_auth"})
                if "chat.getPermalink" in url:
                    return FakeSlackResponse({"ok": True, "permalink": "https://example.slack.com/archives/C123SMOKE/p123"})
                messages = self.messages
                if messages is None:
                    messages = [
                        {
                            "ts": slack_ts,
                            "user": "U123",
                            "text": slack_message_text,
                            "files": [{"name": "brainz-note.md", "mimetype": "text/markdown"}],
                            "attachments": [{"title": "Preview title", "text": "Preview body"}],
                        }
                    ]
                return FakeSlackResponse(
                    {
                        "ok": True,
                        "messages": messages,
                    }
                )

        slack_ts = f"{int(time.time())}.{str(time.time_ns())[-6:]}"
        slack_query = f"Slack Inbox quiet workflow {unique_suffix}"
        slack_message_text = f"{slack_query}\n<https://example.com/brainz|reference>"
        slack_result = poll_slack_inbox(
            database=database,
            memory_folder=root,
            token="xoxb-smoke-token",
            channel_id="C123SMOKE",
            last_ts="",
            session=FakeSlackSession(),
        )
        if slack_result.imported < 1 or slack_result.status != "imported":
            raise RuntimeError("slack inbox import failed")
        if not slack_result.saved_files or not Path(slack_result.saved_files[0]).exists():
            raise RuntimeError("slack markdown was not saved")
        slack_markdown_text = Path(slack_result.saved_files[0]).read_text(encoding="utf-8")
        if "source: slack" not in slack_markdown_text:
            raise RuntimeError("slack markdown source marker missing")
        if slack_query not in slack_markdown_text or "Preview title" not in slack_markdown_text:
            raise RuntimeError("slack canonical text was not preserved")
        slack_duplicate = poll_slack_inbox(
            database=database,
            memory_folder=root,
            token="xoxb-smoke-token",
            channel_id="C123SMOKE",
            last_ts="",
            session=FakeSlackSession(),
        )
        if slack_duplicate.skipped < 1:
            raise RuntimeError("slack duplicate was not skipped")
        backlog_messages = [
            {"ts": f"{int(slack_ts_float(slack_ts)) - 1}.000001", "user": "U123", "text": f"older slack memory {unique_suffix}"},
            {"ts": f"{int(slack_ts_float(slack_ts)) + 1}.000002", "user": "U123", "text": f"#embers backlog slack memory {unique_suffix}"},
        ]
        backlog_result = poll_slack_inbox(
            database=database,
            memory_folder=root,
            token="xoxb-smoke-token",
            channel_id="C123SMOKE",
            last_ts=slack_ts,
            session=FakeSlackSession(messages=backlog_messages),
        )
        if backlog_result.imported != 1 or slack_ts_float(backlog_result.latest_ts) <= slack_ts_float(slack_ts):
            raise RuntimeError("slack backlog sync did not import only newer messages")
        aru_ts = f"{int(time.time()) + 2}.{str(time.time_ns())[-6:]}"
        aru_text = f"#aru 火照りが解けた 在る断片 {unique_suffix}\nnote:\nこれはtaskではなく正本です。"
        aru_result = poll_slack_inbox(
            database=database,
            memory_folder=root,
            token="xoxb-aru-token",
            channel_id="CARUSMOKE",
            last_ts="",
            session=FakeSlackSession(messages=[{"ts": aru_ts, "user": "UARU", "text": aru_text}]),
            source_type="aru",
            folder_name="aru",
            inbox_label="Aru Inbox",
            process_tasks=False,
        )
        if aru_result.imported != 1 or aru_result.task_results:
            raise RuntimeError("aru inbox import failed")
        if not aru_result.saved_files or "\\aru\\" not in str(aru_result.saved_files[0]).lower():
            raise RuntimeError("aru markdown was not saved under aru folder")
        aru_backlog_result = poll_slack_inbox(
            database=database,
            memory_folder=root,
            token="xoxb-aru-token",
            channel_id="CARUSMOKE",
            last_ts=aru_ts,
            session=FakeSlackSession(
                messages=[
                    {"ts": f"{int(slack_ts_float(aru_ts)) - 1}.000001", "user": "UARU", "text": f"older aru {unique_suffix}"},
                    {"ts": f"{int(slack_ts_float(aru_ts)) + 1}.000002", "user": "UARU", "text": f"#embers 在る backlog {unique_suffix}"},
                ]
            ),
            source_type="aru",
            folder_name="aru",
            inbox_label="Aru Inbox",
            process_tasks=False,
        )
        if aru_backlog_result.imported != 1:
            raise RuntimeError("aru backlog sync did not import newer message")
        slack_auth = poll_slack_inbox(
            database=database,
            memory_folder=root,
            token="xoxb-bad-token",
            channel_id="C123SMOKE",
            session=FakeSlackSession("auth_failed"),
        )
        if slack_auth.status != "auth_failed":
            raise RuntimeError("slack auth failure branch failed")
        slack_timeout = poll_slack_inbox(
            database=database,
            memory_folder=root,
            token="xoxb-smoke-token",
            channel_id="C123SMOKE",
            session=FakeSlackSession("timeout"),
        )
        if slack_timeout.status != "timeout":
            raise RuntimeError("slack timeout branch failed")

        slack_import_file = root / "slack_import_target.md"
        slack_import_file.write_text("Slack import task target memory", encoding="utf-8")
        task_base = int(time.time())
        slack_task_query = f"slack task quiet workflow {unique_suffix}"
        slack_task_messages = [
            {"ts": f"{task_base + 1}.000001", "user": "U123", "text": f"search: {slack_task_query}"},
            {"ts": f"{task_base + 2}.000002", "user": "U123", "text": f"note:\nSlack task note memory {unique_suffix}"},
            {"ts": f"{task_base + 3}.000003", "user": "U123", "text": f"handoff_chatgpt:\n{slack_task_query}"},
            {"ts": f"{task_base + 4}.000004", "user": "U123", "text": f"handoff_codex:\n{slack_task_query}"},
            {"ts": f"{task_base + 5}.000005", "user": "U123", "text": f"import:\n{slack_import_file}"},
        ]
        slack_task_result = poll_slack_inbox(
            database=database,
            memory_folder=root,
            token="xoxb-smoke-token",
            channel_id="C123SMOKE",
            last_ts="",
            session=FakeSlackSession(messages=slack_task_messages),
        )
        task_types = {task_result.task_type for task_result in slack_task_result.task_results}
        if task_types != {"search", "note", "handoff_chatgpt", "handoff_codex", "import"}:
            raise RuntimeError("slack task parser did not detect all task types")
        if any(task_result.status != "processed" for task_result in slack_task_result.task_results):
            raise RuntimeError("slack tasks were not processed")
        if not (root / "remote_queue" / "processed").exists():
            raise RuntimeError("slack tasks were not converted to remote queue processed files")
        slack_task_duplicate = poll_slack_inbox(
            database=database,
            memory_folder=root,
            token="xoxb-smoke-token",
            channel_id="C123SMOKE",
            last_ts="",
            session=FakeSlackSession(messages=slack_task_messages),
        )
        if not slack_task_duplicate.task_results or not all(task.skipped_duplicate for task in slack_task_duplicate.task_results):
            raise RuntimeError("slack task duplicate prevention failed")

    engine = SearchEngine(database)
    remote_note_results = engine.search("Remote Queue note memory", limit=10)
    if not any(result.source_type == "remote_queue_note" for result in remote_note_results):
        raise RuntimeError("remote queue note was not indexed")
    remote_import_results = engine.search("Remote Queue import target", limit=10)
    if not remote_import_results:
        raise RuntimeError("remote queue import was not indexed")
    slack_results = engine.search(slack_query, limit=10)
    slack_match = next((result for result in slack_results if result.source_type in {"slack", "slack_inbox"}), None)
    if slack_match is None:
        raise RuntimeError("slack search returned no result")
    if "Slack permalink:" not in slack_match.content or slack_match.conversation_id != "C123SMOKE":
        raise RuntimeError("slack inbox metadata was not stored")
    slack_task_search_results = engine.search(slack_task_query, limit=20)
    if not any(result.source_type == "slack_task" for result in slack_task_search_results):
        raise RuntimeError("slack_task source was not indexed")
    if not any(result.source_type == "remote_queue_note" for result in engine.search(f"Slack task note memory {unique_suffix}", limit=20)):
        raise RuntimeError("slack note task was not indexed as remote_queue_note")
    if not engine.search("Slack import task target memory", limit=10):
        raise RuntimeError("slack import task did not index target file")
    aru_results = engine.search(f"在る断片 {unique_suffix}", limit=10)
    aru_match = next((result for result in aru_results if result.source_type == "aru"), None)
    if aru_match is None:
        raise RuntimeError("aru source search returned no result")
    if "これはtaskではなく正本です。" not in aru_match.content:
        raise RuntimeError("aru canonical text was not preserved")

    file_results = engine.search(UI_TEXT["smoke_query_memory"], limit=10)
    if not file_results:
        raise RuntimeError("file search returned no results")
    non_file_sources = {"chatgpt_export", "codex_result", "codex_report_auto", "slack", "slack_inbox", "slack_task", "aru"}
    file_match = next((result for result in file_results if result.source_type not in non_file_sources), None)
    if file_match is None:
        file_match_results = engine.search(UI_TEXT["smoke_query_git"], limit=10)
        file_match = next((result for result in file_match_results if result.source_type not in non_file_sources), None)
    if file_match is None:
        raise RuntimeError("file search returned no file source result")

    chatgpt_query = f"brainz_zip_{unique_suffix}"
    chatgpt_results = engine.search(chatgpt_query, limit=10)
    chatgpt_match = next((result for result in chatgpt_results if result.source_type == "chatgpt_export"), None)
    if chatgpt_match is None:
        raise RuntimeError("chatgpt_export search returned no result")
    if not chatgpt_match.conversation_title or not chatgpt_match.role:
        raise RuntimeError("chatgpt metadata was not stored")

    chatgpt_path = write_chatgpt_handoff(chatgpt_query, chatgpt_results)
    codex_path = write_codex_handoff(UI_TEXT["smoke_query_git"], chatgpt_results + file_results)
    chatgpt_text = chatgpt_path.read_text(encoding="utf-8")
    if chatgpt_match.conversation_title not in chatgpt_text or chatgpt_match.role not in chatgpt_text:
        raise RuntimeError("chatgpt handoff did not include conversation metadata")
    if not chatgpt_path.exists() or not codex_path.exists():
        raise RuntimeError("handoff export failed")

    codex_results = engine.search(codex_query, limit=10)
    codex_match = next((result for result in codex_results if result.source_type == "codex_result"), None)
    if codex_match is None:
        raise RuntimeError("codex_result search returned no result")
    if not codex_match.commit_hash or not codex_match.changed_files_json:
        raise RuntimeError("codex metadata was not stored")
    codex_chatgpt_path = write_chatgpt_handoff(codex_query, codex_results)
    codex_codex_path = write_codex_handoff(codex_query, codex_results)
    if UI_TEXT["smoke_codex_chatgpt_heading"] not in codex_chatgpt_path.read_text(encoding="utf-8"):
        raise RuntimeError("chatgpt handoff did not include codex history")
    if UI_TEXT["smoke_codex_codex_heading"] not in codex_codex_path.read_text(encoding="utf-8"):
        raise RuntimeError("codex handoff did not include implementation history")

    codex_report_query = f"codex report auto import {unique_commit_c[:7]}"
    codex_report_results = engine.search(codex_report_query, limit=10)
    codex_report_match = next((result for result in codex_report_results if result.source_type == "codex_report_auto"), None)
    if codex_report_match is None:
        raise RuntimeError("codex_report_auto search returned no result")
    if not codex_report_match.commit_hash or not codex_report_match.changed_files_json:
        raise RuntimeError("codex_report_auto metadata was not stored")
    if "Source Type:" in codex_report_match.content or "Summary:" in codex_report_match.content:
        raise RuntimeError("codex_report_auto raw markdown was altered")
    codex_report_handoff = write_codex_handoff(codex_report_query, codex_report_results)
    if UI_TEXT["smoke_codex_codex_heading"] not in codex_report_handoff.read_text(encoding="utf-8"):
        raise RuntimeError("codex_report_auto handoff was not treated as codex history")

    semantic_off = engine.search_with_related(UI_TEXT["smoke_query_memory"], limit=10, semantic_enabled=False)
    if not semantic_off.results or semantic_off.semantic_available:
        raise RuntimeError("semantic off fallback failed")
    embedding_status = check_embedding_status(model_name=DEFAULT_EMBED_MODEL)
    semantic_response = engine.search_with_related(UI_TEXT["smoke_query_memory"], limit=10, semantic_enabled=True)
    embedding_stats = database.embedding_stats()
    if embedding_status.available:
        if not semantic_response.semantic_available:
            raise RuntimeError("semantic search was unavailable despite embedding readiness")
        if embedding_stats["ready"] > 0 and not semantic_response.related:
            raise RuntimeError("semantic related memory returned no results")
    elif semantic_response.semantic_available:
        raise RuntimeError("semantic unavailable branch failed")

    flow_chatgpt = engine.memory_flow(chatgpt_match, semantic_enabled=False, ascending=True)
    if len(flow_chatgpt.items) < 2:
        raise RuntimeError("chatgpt memory flow returned too few items")
    if sum(1 for item in flow_chatgpt.items if item.result.conversation_id == chatgpt_match.conversation_id) < 2:
        raise RuntimeError("chatgpt memory flow did not include same conversation")

    flow_codex = engine.memory_flow(codex_match, semantic_enabled=False, ascending=True)
    if not any(item.result.source_type == "codex_result" for item in flow_codex.items):
        raise RuntimeError("codex memory flow did not include codex_result")
    flow_codex_report = engine.memory_flow(codex_report_match, semantic_enabled=False, ascending=True)
    if not any(item.result.source_type == "codex_report_auto" for item in flow_codex_report.items):
        raise RuntimeError("codex report memory flow did not include codex_report_auto")
    flow_slack = engine.memory_flow(slack_match, semantic_enabled=False, ascending=True)
    if not any(item.result.source_type in {"slack", "slack_inbox"} for item in flow_slack.items):
        raise RuntimeError("slack memory flow did not include slack")
    flow_aru = engine.memory_flow(aru_match, semantic_enabled=False, ascending=True)
    if not any(item.result.source_type == "aru" for item in flow_aru.items):
        raise RuntimeError("aru memory flow did not include aru")

    flow_file = engine.memory_flow(file_match, semantic_enabled=False, ascending=True)
    if not any(item.result.source_type not in non_file_sources for item in flow_file.items):
        raise RuntimeError("file memory flow did not include file result")

    flow_semantic = engine.memory_flow(chatgpt_match, semantic_enabled=True, ascending=True)
    if embedding_status.available and not flow_semantic.semantic_available:
        raise RuntimeError("semantic memory flow was unavailable despite embedding readiness")
    flow_desc = engine.memory_flow(chatgpt_match, semantic_enabled=False, ascending=False)
    if not flow_desc.items:
        raise RuntimeError("descending memory flow returned no items")

    ollama_status = check_ollama()
    gpu_status = check_gpu()
    stats = database.stats()
    print("SMOKE OK")
    print(f"documents={stats['documents']} chunks={stats['chunks']}")
    print(f"file_results={len(file_results)}")
    print(f"chatgpt_results={len(chatgpt_results)}")
    print(f"codex_results={len(codex_results)}")
    print(f"codex_report_results={len(codex_report_results)}")
    print(f"chatgpt_handoff={chatgpt_path}")
    print(f"codex_handoff={codex_path}")
    print(f"codex_chatgpt_handoff={codex_chatgpt_path}")
    print(f"codex_codex_handoff={codex_codex_path}")
    print(f"codex_report_handoff={codex_report_handoff}")
    print(f"ollama_available={ollama_status.available}")
    print(f"embedding_available={embedding_status.available}")
    print(f"embeddings_ready={embedding_stats['ready']}")
    print(f"semantic_available={semantic_response.semantic_available}")
    print(f"related_memory={len(semantic_response.related)}")
    print(f"memory_flow_chatgpt={len(flow_chatgpt.items)}")
    print(f"memory_flow_codex={len(flow_codex.items)}")
    print(f"memory_flow_codex_report={len(flow_codex_report.items)}")
    print(f"memory_flow_slack={len(flow_slack.items)}")
    print(f"memory_flow_aru={len(flow_aru.items)}")
    print(f"memory_flow_file={len(flow_file.items)}")
    print(f"memory_flow_semantic={flow_semantic.semantic_available}")
    print(f"watch_new_detected={len(watch_detection.changed_files)}")
    print(f"watch_update_detected={len(watch_update.changed_files)}")
    print(f"notification_queue_history={len(notification_queue.history)}")
    print(f"notifications_default={default_config.enable_notifications}")
    print(f"oikawa_launch_target={oikawa_target}")
    print(f"oikawa_launch_mode={oikawa_mode}")
    print(f"remote_queue_processed={remote_result.processed}")
    print(f"remote_queue_failed={remote_result.failed}")
    print(f"remote_queue_note_results={len(remote_note_results)}")
    print(f"codex_report_imported={codex_report_result.imported}")
    print(f"codex_report_failed={codex_report_result.failed}")
    print(f"slack_imported={slack_result.imported}")
    print(f"slack_duplicate_skipped={slack_duplicate.skipped}")
    print(f"slack_tasks={len(slack_task_result.task_results)}")
    print(f"slack_task_duplicates={sum(1 for task in slack_task_duplicate.task_results if task.skipped_duplicate)}")
    print(f"aru_imported={aru_result.imported}")
    print(f"gpu_detected={gpu_status.gpu_detected}")
    return 0


def sample_conversations(conversation_id: str, title: str, user_text: str, assistant_text: str) -> list[dict[str, object]]:
    return [
        {
            "id": conversation_id,
            "title": title,
            "create_time": 1767225600.0,
            "update_time": 1767225660.0,
            "mapping": {
                "user-node": {
                    "message": {
                        "author": {"role": "user"},
                        "create_time": 1767225601.0,
                        "content": {"content_type": "text", "parts": [user_text]},
                    }
                },
                "assistant-node": {
                    "message": {
                        "author": {"role": "assistant"},
                        "create_time": 1767225602.0,
                        "content": {"content_type": "text", "parts": [assistant_text]},
                    }
                },
            },
        }
    ]


def fit_logo_size(image_size: tuple[int, int], max_width: int = 118, max_height: int = 36) -> tuple[int, int]:
    width, height = image_size
    if width <= 0 or height <= 0:
        return max_height, max_height
    scale = min(max_width / width, max_height / height)
    return max(1, round(width * scale)), max(1, round(height * scale))


def parse_slack_poll_interval(value: str) -> tuple[int, bool]:
    try:
        interval = int(str(value).strip())
    except ValueError:
        return 10, False
    if interval < 5 or interval > 15:
        return max(5, min(15, interval)), False
    return interval, True


def slack_config_missing_fields(
    memory_folder: str,
    token: str,
    channel_id: str,
    interval_valid: bool,
) -> list[str]:
    missing: list[str] = []
    if not memory_folder:
        missing.append("memory_folder")
    if not token:
        missing.append("bot_token")
    if not channel_id:
        missing.append("channel_id")
    if not interval_valid:
        missing.append("poll_interval")
    return missing


def run_gui(launch_check: bool = False) -> int:
    import customtkinter as ctk
    from PIL import Image
    from tkinter import filedialog, messagebox

    from core.app_config import (
        ConfigStore,
        ensure_app_dirs,
        now_iso,
        open_path,
        peakheadz_icon_path,
        peakheadz_logo_path,
    )
    from core.chatgpt_importer import ConversationsJsonNotFound, ChatGPTImportResult, import_chatgpt_export
    from core.codex_importer import CodexImportResult, import_codex_file, import_codex_text
    from core.codex_report_auto import CodexReportAutoResult, count_pending_reports, process_codex_reports_folder
    from core.db import BrainzDatabase, SearchResult
    from core.embers import HEAT_TERMS, default_ember_query, detect_heat_terms
    from core.gpu_checker import check_gpu
    from core.handoff_writer import write_chatgpt_handoff, write_codex_handoff
    from core.indexer import IndexProgress, Indexer
    from core.memory_flow import MemoryFlowItem, MemoryFlowResponse, short_summary, timeline_date
    from core.notifications import NotificationItem, NotificationQueue
    from core.ollama_client import check_ollama
    from core.ollama_embeddings import check_embedding_status
    from core.qpsc_status import write_brainz_awake_status
    from core.remote_queue import RemoteQueueBatchResult, count_pending_queue_files, process_remote_queue_folder
    from core.search_engine import SearchEngine, SearchResponse
    from core.slack_inbox import SlackInboxResult, SlackTaskResult, poll_slack_inbox
    from core.watch_folder import WatchScanResult, detect_changed_files, write_watch_log
    from ui.components import choose_font_family, set_textbox_text
    from ui.theme import (
        COLORS,
        FONT_CANDIDATES,
        FONT_SIZES,
        MONO_FONT_CANDIDATES,
        READING_FONT_CANDIDATES,
        STATUS_FONT_CANDIDATES,
    )

    ensure_app_dirs()
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")

    class BrainzApp(ctk.CTk):
        def __init__(self) -> None:
            super().__init__()
            self.title(UI_TEXT["window_title"])
            self.geometry("1360x820")
            self.minsize(1120, 680)
            self.configure(fg_color=COLORS["bg"])
            self.font_family = choose_font_family(self, FONT_CANDIDATES)
            self.reading_font_family = choose_font_family(self, READING_FONT_CANDIDATES)
            self.status_font_family = choose_font_family(self, STATUS_FONT_CANDIDATES)
            self.mono_font_family = choose_font_family(self, MONO_FONT_CANDIDATES)
            self.config_store = ConfigStore()
            self.config_data = self.config_store.load()
            if not self.config_data.watch_folder and self.config_data.memory_folder:
                self.config_data.watch_folder = self.config_data.memory_folder
            if not self.config_data.codex_reports_folder and self.config_data.memory_folder:
                self.config_data.codex_reports_folder = str(Path(self.config_data.memory_folder) / "codex_reports")
            self.database = BrainzDatabase()
            self.database.ensure_schema()
            self.search_engine = SearchEngine(self.database)
            self.indexer = Indexer(self.database)
            self.events: queue.Queue[tuple[str, object]] = queue.Queue()
            self.cancel_event = threading.Event()
            self.index_thread: threading.Thread | None = None
            self.search_thread: threading.Thread | None = None
            self.import_thread: threading.Thread | None = None
            self.flow_thread: threading.Thread | None = None
            self.watch_scan_thread: threading.Thread | None = None
            self.remote_queue_thread: threading.Thread | None = None
            self.codex_report_thread: threading.Thread | None = None
            self.slack_thread: threading.Thread | None = None
            self.aru_thread: threading.Thread | None = None
            self.current_results: list[SearchResult] = []
            self.current_related_results: list[SearchResult] = []
            self.current_flow_items: list[MemoryFlowItem] = []
            self.selected_result: SearchResult | None = None
            self.current_query = self.config_data.last_query
            self.semantic_available = True
            self.view_mode_var = ctk.StringVar(value=UI_TEXT["tab_search"])
            self.embers_mode = False
            self.semantic_search_var = ctk.BooleanVar(value=True)
            self.auto_index_var = ctk.BooleanVar(value=self.config_data.auto_index_enabled)
            self.notifications_var = ctk.BooleanVar(value=self.config_data.enable_notifications)
            self.remote_queue_var = ctk.BooleanVar(value=self.config_data.enable_remote_queue)
            self.slack_inbox_var = ctk.BooleanVar(value=self.config_data.enable_slack_inbox)
            self.aru_inbox_var = ctk.BooleanVar(value=self.config_data.enable_aru_inbox)
            self.flow_sort_ascending = True
            self.flow_request_id = 0
            self.flow_cache: dict[tuple[int, bool, bool], MemoryFlowResponse] = {}
            self.logo_image = None
            self.auto_index_active = False
            self.pending_auto_index_folder: Path | None = None
            self.watch_poll_interval_ms = 8000
            self.remote_queue_poll_interval_ms = 8000
            self.codex_report_poll_interval_ms = 8000
            self.slack_poll_interval_ms = max(5000, int(self.config_data.slack_poll_interval_seconds) * 1000)
            self.aru_poll_interval_ms = max(5000, int(self.config_data.aru_poll_interval_seconds) * 1000)
            self.notification_queue = NotificationQueue(history_limit=3)
            self.notification_visible = False
            self.last_slack_status = ""
            self.last_aru_status = ""
            self.qpsc_started_at = now_iso()

            self._apply_icon()
            self._build_ui()
            self._set_memory_folder(self.config_data.memory_folder, persist=False)
            self._append_log(UI_TEXT["log_ready"])
            self._write_qpsc_awake_status()
            self._update_watch_status()
            if self.config_data.watch_folder:
                self._append_log(UI_TEXT["log_watch_initialized"])
                self._append_log(UI_TEXT["log_watching_folder"].format(folder=self.config_data.watch_folder))
            self._refresh_stats()
            self._refresh_system_status()
            self._update_remote_queue_status()
            self._update_codex_report_status()
            self._update_slack_status()
            self._update_aru_status()
            self.after(100, self._poll_events)
            if not launch_check:
                self.after(1500, self._poll_watch_folder)
                self.after(2200, self._poll_remote_queue)
                self.after(2800, self._poll_codex_reports)
                self.after(3400, self._poll_slack_inbox)
                self.after(3800, self._poll_aru_inbox)
            if launch_check:
                self.after(1200, self._launch_check_finish)
            else:
                self.after(30000, self._heartbeat_qpsc_status)

        def _apply_icon(self) -> None:
            try:
                icon = peakheadz_icon_path()
                if icon.exists():
                    self.iconbitmap(str(icon))
            except Exception:
                pass

        def _build_ui(self) -> None:
            self.grid_columnconfigure(0, weight=0, minsize=300)
            self.grid_columnconfigure(1, weight=1, minsize=470)
            self.grid_columnconfigure(2, weight=0, minsize=370)
            self.grid_rowconfigure(1, weight=1)
            self.grid_rowconfigure(2, weight=0)

            header = ctk.CTkFrame(self, fg_color=COLORS["bg"], corner_radius=0)
            header.grid(row=0, column=0, columnspan=3, sticky="ew", padx=20, pady=(14, 8))
            header.grid_columnconfigure(1, weight=1)

            logo_path = peakheadz_logo_path()
            if logo_path.exists():
                try:
                    image = Image.open(logo_path)
                    self.logo_image = ctk.CTkImage(light_image=image, dark_image=image, size=fit_logo_size(image.size))
                    ctk.CTkLabel(header, image=self.logo_image, text="").grid(row=0, column=0, rowspan=3, padx=(0, 10))
                except Exception:
                    pass

            ctk.CTkLabel(
                header,
                text=UI_TEXT["app_title"],
                text_color=COLORS["text"],
                font=(self.font_family, FONT_SIZES["title"]),
            ).grid(row=0, column=1, sticky="w")
            ctk.CTkLabel(
                header,
                text=UI_TEXT["subtitle"],
                text_color=COLORS["muted"],
                font=(self.status_font_family, FONT_SIZES["subtitle"]),
            ).grid(row=1, column=1, sticky="w")
            ctk.CTkLabel(
                header,
                text=UI_TEXT["role_summary"],
                text_color=COLORS["quiet"],
                font=(self.status_font_family, FONT_SIZES["micro"]),
            ).grid(row=2, column=1, sticky="w", pady=(2, 0))

            header_action = ctk.CTkFrame(header, fg_color=COLORS["bg"], corner_radius=0)
            header_action.grid(row=0, column=2, rowspan=3, sticky="e")
            header_action.grid_columnconfigure(0, weight=1)
            self.awake_status_var = ctk.StringVar(value=UI_TEXT["status_brainz_standby"])
            ctk.CTkLabel(
                header_action,
                textvariable=self.awake_status_var,
                text_color=COLORS["section"],
                font=(self.status_font_family, FONT_SIZES["small"]),
                anchor="e",
            ).grid(row=0, column=0, sticky="e", pady=(0, 7))
            ctk.CTkButton(
                header_action,
                text=UI_TEXT["button_open_oikawa"],
                width=168,
                height=38,
                fg_color=COLORS["accent"],
                hover_color=COLORS["accent_hover"],
                font=(self.status_font_family, FONT_SIZES["button"]),
                command=self._open_oikawa,
            ).grid(row=1, column=0, sticky="e")

            left = ctk.CTkScrollableFrame(
                self,
                fg_color=COLORS["panel"],
                border_color=COLORS["border"],
                border_width=1,
                corner_radius=0,
                scrollbar_button_color=COLORS["panel_soft"],
                scrollbar_button_hover_color=COLORS["accent_soft"],
            )
            left.grid(row=1, column=0, sticky="nsew", padx=(18, 8), pady=(0, 10))
            left.grid_columnconfigure(0, weight=1)
            self._section_title(left, UI_TEXT["memory_title"], 0)
            self.memory_var = ctk.StringVar(value=UI_TEXT["empty_memory"])
            ctk.CTkLabel(
                left,
                textvariable=self.memory_var,
                text_color=COLORS["muted"],
                font=(self.reading_font_family, FONT_SIZES["small"]),
                wraplength=238,
                justify="left",
            ).grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 10))
            ctk.CTkButton(
                left,
                text=UI_TEXT["button_choose"],
                height=36,
                fg_color=COLORS["panel_soft"],
                hover_color=COLORS["accent_soft"],
                font=(self.font_family, FONT_SIZES["button"]),
                command=self._choose_memory_folder,
            ).grid(row=2, column=0, sticky="ew", padx=16, pady=(0, 9))
            chatgpt_card = ctk.CTkFrame(
                left,
                fg_color=COLORS["panel_alt"],
                border_color=COLORS["border"],
                border_width=1,
                corner_radius=6,
            )
            chatgpt_card.grid(row=3, column=0, sticky="ew", padx=16, pady=(0, 10))
            chatgpt_card.grid_columnconfigure((0, 1), weight=1)
            ctk.CTkLabel(
                chatgpt_card,
                text=UI_TEXT["chatgpt_import_title"],
                text_color=COLORS["section"],
                font=(self.font_family, FONT_SIZES["section"]),
                anchor="w",
            ).grid(row=0, column=0, columnspan=2, sticky="ew", padx=12, pady=(10, 2))
            ctk.CTkLabel(
                chatgpt_card,
                text=UI_TEXT["chatgpt_import_helper"],
                text_color=COLORS["muted"],
                font=(self.reading_font_family, FONT_SIZES["small"]),
                anchor="w",
            ).grid(row=1, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 2))
            ctk.CTkLabel(
                chatgpt_card,
                text=UI_TEXT["chatgpt_import_note"],
                text_color=COLORS["muted"],
                font=(self.reading_font_family, FONT_SIZES["small"]),
                wraplength=218,
                justify="left",
                anchor="w",
            ).grid(row=2, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 8))
            self.import_button = ctk.CTkButton(
                chatgpt_card,
                text=UI_TEXT["button_import_chatgpt_zip"],
                height=32,
                fg_color=COLORS["panel_soft"],
                hover_color=COLORS["accent_soft"],
                font=(self.font_family, FONT_SIZES["button"]),
                command=self._choose_chatgpt_export_zip,
            )
            self.import_button.grid(row=3, column=0, sticky="ew", padx=(12, 5), pady=(0, 6))
            self.chatgpt_folder_button = ctk.CTkButton(
                chatgpt_card,
                text=UI_TEXT["button_import_chatgpt_folder"],
                height=32,
                fg_color=COLORS["panel_soft"],
                hover_color=COLORS["accent_soft"],
                font=(self.font_family, FONT_SIZES["button"]),
                command=self._choose_chatgpt_export_folder,
            )
            self.chatgpt_folder_button.grid(row=3, column=1, sticky="ew", padx=(5, 12), pady=(0, 6))
            self.chatgpt_json_button = ctk.CTkButton(
                chatgpt_card,
                text=UI_TEXT["button_import_chatgpt_json"],
                height=32,
                fg_color=COLORS["panel_soft"],
                hover_color=COLORS["accent_soft"],
                font=(self.font_family, FONT_SIZES["button"]),
                command=self._choose_chatgpt_conversations_json,
            )
            self.chatgpt_json_button.grid(row=4, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 8))
            self.chatgpt_import_status_var = ctk.StringVar(value=UI_TEXT["chatgpt_import_waiting"])
            ctk.CTkLabel(
                chatgpt_card,
                textvariable=self.chatgpt_import_status_var,
                text_color=COLORS["muted"],
                font=(self.reading_font_family, FONT_SIZES["small"]),
                wraplength=218,
                justify="left",
                anchor="w",
            ).grid(row=5, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 10))
            self.codex_import_button = ctk.CTkButton(
                left,
                text=UI_TEXT["button_import_codex_result"],
                height=36,
                fg_color=COLORS["panel_soft"],
                hover_color=COLORS["accent_soft"],
                font=(self.font_family, FONT_SIZES["button"]),
                command=self._open_codex_import_dialog,
            )
            self.codex_import_button.grid(row=4, column=0, sticky="ew", padx=16, pady=(0, 16))

            self._section_title(left, UI_TEXT["label_watch_folder"], 5)
            self.watch_folder_var = ctk.StringVar(value=UI_TEXT["empty_watch_folder"])
            ctk.CTkLabel(
                left,
                textvariable=self.watch_folder_var,
                text_color=COLORS["muted"],
                font=(self.reading_font_family, FONT_SIZES["small"]),
                wraplength=238,
                justify="left",
            ).grid(row=6, column=0, sticky="ew", padx=16, pady=(0, 8))
            ctk.CTkButton(
                left,
                text=UI_TEXT["button_choose"],
                height=34,
                fg_color=COLORS["panel_soft"],
                hover_color=COLORS["accent_soft"],
                font=(self.font_family, FONT_SIZES["button"]),
                command=self._choose_watch_folder,
            ).grid(row=7, column=0, sticky="ew", padx=16, pady=(0, 8))
            self.auto_index_checkbox = ctk.CTkCheckBox(
                left,
                text=UI_TEXT["checkbox_auto_index"],
                variable=self.auto_index_var,
                text_color=COLORS["muted"],
                fg_color=COLORS["accent"],
                hover_color=COLORS["accent_hover"],
                border_color=COLORS["border"],
                font=(self.status_font_family, FONT_SIZES["small"]),
                command=self._toggle_auto_index,
            )
            self.auto_index_checkbox.grid(row=8, column=0, sticky="w", padx=16, pady=(0, 8))
            self.watch_status_var = ctk.StringVar(value=UI_TEXT["status_auto_index_off"])
            self.last_memory_var = ctk.StringVar(value=UI_TEXT["status_last_memory"].format(path="-"))
            self.last_index_var = ctk.StringVar(value=UI_TEXT["status_last_index"].format(time="-"))
            for watch_index, variable in enumerate(
                (self.watch_status_var, self.last_memory_var, self.last_index_var),
                start=9,
            ):
                ctk.CTkLabel(
                    left,
                    textvariable=variable,
                    text_color=COLORS["muted"],
                    font=(self.status_font_family, FONT_SIZES["micro"]),
                    anchor="w",
                ).grid(row=watch_index, column=0, sticky="ew", padx=16, pady=2)

            self._section_title(left, UI_TEXT["section_remote_queue"], 12)
            self.remote_queue_folder_var = ctk.StringVar(value=UI_TEXT["empty_remote_queue_folder"])
            ctk.CTkLabel(
                left,
                textvariable=self.remote_queue_folder_var,
                text_color=COLORS["muted"],
                font=(self.reading_font_family, FONT_SIZES["small"]),
                wraplength=238,
                justify="left",
            ).grid(row=13, column=0, sticky="ew", padx=16, pady=(0, 8))
            remote_row = ctk.CTkFrame(left, fg_color="transparent")
            remote_row.grid(row=14, column=0, sticky="ew", padx=16, pady=(0, 8))
            remote_row.grid_columnconfigure(0, weight=1)
            remote_row.grid_columnconfigure(1, weight=1)
            ctk.CTkButton(
                remote_row,
                text=UI_TEXT["button_choose_remote_queue"],
                height=34,
                fg_color=COLORS["panel_soft"],
                hover_color=COLORS["accent_soft"],
                font=(self.font_family, FONT_SIZES["button"]),
                command=self._choose_remote_queue_folder,
            ).grid(row=0, column=0, sticky="ew", padx=(0, 6))
            self.remote_queue_checkbox = ctk.CTkCheckBox(
                remote_row,
                text=UI_TEXT["checkbox_enable_remote_queue"],
                variable=self.remote_queue_var,
                text_color=COLORS["muted"],
                fg_color=COLORS["accent"],
                hover_color=COLORS["accent_hover"],
                border_color=COLORS["border"],
                font=(self.status_font_family, FONT_SIZES["micro"]),
                command=self._toggle_remote_queue,
            )
            self.remote_queue_checkbox.grid(row=0, column=1, sticky="w", padx=(6, 0))
            self.remote_queue_status_var = ctk.StringVar(value=UI_TEXT["status_remote_queue_off"])
            self.remote_queue_counts_var = ctk.StringVar(
                value=UI_TEXT["status_remote_queue_counts"].format(pending=0, processed=0, failed=0)
            )
            for remote_index, variable in enumerate((self.remote_queue_status_var, self.remote_queue_counts_var), start=15):
                ctk.CTkLabel(
                    left,
                    textvariable=variable,
                    text_color=COLORS["muted"],
                    font=(self.status_font_family, FONT_SIZES["micro"]),
                    anchor="w",
                ).grid(row=remote_index, column=0, sticky="ew", padx=16, pady=2)

            self._section_title(left, UI_TEXT["section_codex_reports"], 17)
            self.codex_report_status_var = ctk.StringVar(value=UI_TEXT["status_codex_report_idle"])
            self.last_report_var = ctk.StringVar(value=UI_TEXT["label_last_report"].format(path="-"))
            self.last_commit_var = ctk.StringVar(value=UI_TEXT["label_last_commit"].format(commit_hash="-"))
            for report_index, variable in enumerate(
                (self.codex_report_status_var, self.last_report_var, self.last_commit_var),
                start=18,
            ):
                ctk.CTkLabel(
                    left,
                    textvariable=variable,
                    text_color=COLORS["muted"],
                    font=(self.status_font_family, FONT_SIZES["micro"]),
                    anchor="w",
                ).grid(row=report_index, column=0, sticky="ew", padx=16, pady=2)

            self._section_title(left, UI_TEXT["section_slack_inbox"], 21)
            self.slack_token_entry = ctk.CTkEntry(
                left,
                height=28,
                placeholder_text=UI_TEXT["label_slack_token"],
                show="*",
                fg_color=COLORS["input"],
                border_color=COLORS["border"],
                text_color=COLORS["text"],
                font=(self.status_font_family, FONT_SIZES["micro"]),
            )
            self.slack_token_entry.grid(row=22, column=0, sticky="ew", padx=16, pady=(0, 5))
            self.slack_token_entry.insert(0, self.config_data.slack_bot_token)
            slack_row = ctk.CTkFrame(left, fg_color="transparent")
            slack_row.grid(row=23, column=0, sticky="ew", padx=16, pady=(0, 5))
            slack_row.grid_columnconfigure(0, weight=2)
            slack_row.grid_columnconfigure(1, weight=1)
            self.slack_channel_entry = ctk.CTkEntry(
                slack_row,
                height=28,
                placeholder_text=UI_TEXT["label_slack_channel"],
                fg_color=COLORS["input"],
                border_color=COLORS["border"],
                text_color=COLORS["text"],
                font=(self.status_font_family, FONT_SIZES["micro"]),
            )
            self.slack_channel_entry.grid(row=0, column=0, sticky="ew", padx=(0, 5))
            self.slack_channel_entry.insert(0, self.config_data.slack_channel_id)
            self.slack_interval_entry = ctk.CTkEntry(
                slack_row,
                height=28,
                placeholder_text=UI_TEXT["label_slack_interval"],
                fg_color=COLORS["input"],
                border_color=COLORS["border"],
                text_color=COLORS["text"],
                font=(self.status_font_family, FONT_SIZES["micro"]),
            )
            self.slack_interval_entry.grid(row=0, column=1, sticky="ew", padx=(5, 0))
            self.slack_interval_entry.insert(0, str(self.config_data.slack_poll_interval_seconds))
            slack_control_row = ctk.CTkFrame(left, fg_color="transparent")
            slack_control_row.grid(row=24, column=0, sticky="ew", padx=16, pady=(0, 5))
            slack_control_row.grid_columnconfigure((0, 1), weight=1)
            self.slack_inbox_checkbox = ctk.CTkCheckBox(
                slack_control_row,
                text=UI_TEXT["checkbox_enable_slack_inbox"],
                variable=self.slack_inbox_var,
                text_color=COLORS["muted"],
                fg_color=COLORS["accent"],
                hover_color=COLORS["accent_hover"],
                border_color=COLORS["border"],
                font=(self.status_font_family, FONT_SIZES["micro"]),
                command=self._toggle_slack_inbox,
            )
            self.slack_inbox_checkbox.grid(row=0, column=0, sticky="w", padx=(0, 6))
            ctk.CTkButton(
                slack_control_row,
                text=UI_TEXT["button_save_slack"],
                height=28,
                fg_color=COLORS["panel_soft"],
                hover_color=COLORS["accent_soft"],
                font=(self.status_font_family, FONT_SIZES["micro"]),
                command=self._save_slack_settings,
            ).grid(row=0, column=1, sticky="ew", padx=(6, 0))
            self.slack_status_var = ctk.StringVar(value=UI_TEXT["status_slack_disabled"])
            self.slack_last_import_var = ctk.StringVar(value=UI_TEXT["label_slack_last_import"].format(time="-"))
            self.slack_channel_var = ctk.StringVar(value=UI_TEXT["label_slack_channel_status"].format(channel="-"))
            self.last_slack_task_var = ctk.StringVar(value=UI_TEXT["label_last_task"].format(task="-"))
            for slack_index, variable in enumerate(
                (self.slack_status_var, self.slack_last_import_var, self.slack_channel_var, self.last_slack_task_var),
                start=25,
            ):
                ctk.CTkLabel(
                    left,
                    textvariable=variable,
                    text_color=COLORS["muted"],
                    font=(self.status_font_family, FONT_SIZES["micro"]),
                    anchor="w",
                ).grid(row=slack_index, column=0, sticky="ew", padx=16, pady=1)

            self._section_title(left, UI_TEXT["section_aru_inbox"], 29)
            self.aru_token_entry = ctk.CTkEntry(
                left,
                height=28,
                placeholder_text=UI_TEXT["label_aru_token"],
                show="*",
                fg_color=COLORS["input"],
                border_color=COLORS["border"],
                text_color=COLORS["text"],
                font=(self.status_font_family, FONT_SIZES["micro"]),
            )
            self.aru_token_entry.grid(row=30, column=0, sticky="ew", padx=16, pady=(0, 5))
            self.aru_token_entry.insert(0, self.config_data.aru_slack_token)
            self.aru_channel_entry = ctk.CTkEntry(
                left,
                height=28,
                placeholder_text=UI_TEXT["label_aru_channel"],
                fg_color=COLORS["input"],
                border_color=COLORS["border"],
                text_color=COLORS["text"],
                font=(self.status_font_family, FONT_SIZES["micro"]),
            )
            self.aru_channel_entry.grid(row=31, column=0, sticky="ew", padx=16, pady=(0, 5))
            self.aru_channel_entry.insert(0, self.config_data.aru_channel_id)
            aru_control_row = ctk.CTkFrame(left, fg_color="transparent")
            aru_control_row.grid(row=32, column=0, sticky="ew", padx=16, pady=(0, 5))
            aru_control_row.grid_columnconfigure((0, 1), weight=1)
            self.aru_inbox_checkbox = ctk.CTkCheckBox(
                aru_control_row,
                text=UI_TEXT["checkbox_enable_aru_inbox"],
                variable=self.aru_inbox_var,
                text_color=COLORS["muted"],
                fg_color=COLORS["accent"],
                hover_color=COLORS["accent_hover"],
                border_color=COLORS["border"],
                font=(self.status_font_family, FONT_SIZES["micro"]),
                command=self._toggle_aru_inbox,
            )
            self.aru_inbox_checkbox.grid(row=0, column=0, sticky="w", padx=(0, 6))
            ctk.CTkButton(
                aru_control_row,
                text=UI_TEXT["button_save_aru"],
                height=28,
                fg_color=COLORS["panel_soft"],
                hover_color=COLORS["accent_soft"],
                font=(self.status_font_family, FONT_SIZES["micro"]),
                command=self._save_aru_settings,
            ).grid(row=0, column=1, sticky="ew", padx=(6, 0))
            self.aru_status_var = ctk.StringVar(value=UI_TEXT["status_aru_disabled"])
            self.aru_last_import_var = ctk.StringVar(value=UI_TEXT["label_aru_last_import"].format(time="-"))
            self.aru_channel_var = ctk.StringVar(value=UI_TEXT["label_aru_channel_status"].format(channel="-"))
            for aru_index, variable in enumerate(
                (self.aru_status_var, self.aru_last_import_var, self.aru_channel_var),
                start=33,
            ):
                ctk.CTkLabel(
                    left,
                    textvariable=variable,
                    text_color=COLORS["muted"],
                    font=(self.status_font_family, FONT_SIZES["micro"]),
                    anchor="w",
                ).grid(row=aru_index, column=0, sticky="ew", padx=16, pady=1)

            self._section_title(left, UI_TEXT["index_title"], 36)
            self.index_status_var = ctk.StringVar(value=UI_TEXT["index_idle"])
            ctk.CTkLabel(
                left,
                textvariable=self.index_status_var,
                text_color=COLORS["section"],
                font=(self.status_font_family, FONT_SIZES["body"]),
            ).grid(row=37, column=0, sticky="w", padx=16, pady=(0, 8))
            self.progress = ctk.CTkProgressBar(left, height=10, progress_color=COLORS["accent"])
            self.progress.set(0)
            self.progress.grid(row=38, column=0, sticky="ew", padx=16, pady=(0, 12))

            button_row = ctk.CTkFrame(left, fg_color="transparent")
            button_row.grid(row=39, column=0, sticky="ew", padx=16, pady=(0, 18))
            button_row.grid_columnconfigure((0, 1), weight=1)
            self.index_button = ctk.CTkButton(
                button_row,
                text=UI_TEXT["button_index"],
                height=36,
                fg_color=COLORS["accent"],
                hover_color=COLORS["accent_hover"],
                font=(self.status_font_family, FONT_SIZES["button"]),
                command=self._start_index,
            )
            self.index_button.grid(row=0, column=0, sticky="ew", padx=(0, 6))
            self.cancel_button = ctk.CTkButton(
                button_row,
                text=UI_TEXT["button_cancel"],
                height=36,
                fg_color=COLORS["panel_soft"],
                hover_color=COLORS["accent_soft"],
                font=(self.status_font_family, FONT_SIZES["button"]),
                state="disabled",
                command=self._cancel_index,
            )
            self.cancel_button.grid(row=0, column=1, sticky="ew", padx=(6, 0))

            self._section_title(left, UI_TEXT["system_title"], 40)
            self.sqlite_var = ctk.StringVar(value=UI_TEXT["status_sqlite_ready"])
            self.cuda_var = ctk.StringVar(value=UI_TEXT["status_cuda_unavailable"])
            self.gpu_var = ctk.StringVar(value=UI_TEXT["status_gpu_missing"])
            self.ollama_var = ctk.StringVar(value=UI_TEXT["status_ollama_not_running"])
            self.embedding_var = ctk.StringVar(value=UI_TEXT["status_embedding_unavailable"])
            self.semantic_status_var = ctk.StringVar(value=UI_TEXT["status_semantic_search_disabled"])
            self.docs_var = ctk.StringVar(value=UI_TEXT["status_docs"].format(documents=0, chunks=0))
            for index, variable in enumerate(
                (
                    self.sqlite_var,
                    self.cuda_var,
                    self.gpu_var,
                    self.ollama_var,
                    self.embedding_var,
                    self.semantic_status_var,
                    self.docs_var,
                ),
                start=41,
            ):
                ctk.CTkLabel(
                    left,
                    textvariable=variable,
                    text_color=COLORS["muted"],
                    font=(self.status_font_family, FONT_SIZES["small"]),
                    anchor="w",
                ).grid(row=index, column=0, sticky="ew", padx=16, pady=2)

            self._section_title(left, UI_TEXT["section_notifications"], 48)
            self.notifications_checkbox = ctk.CTkCheckBox(
                left,
                text=UI_TEXT["checkbox_enable_notifications"],
                variable=self.notifications_var,
                text_color=COLORS["muted"],
                fg_color=COLORS["accent"],
                hover_color=COLORS["accent_hover"],
                border_color=COLORS["border"],
                font=(self.status_font_family, FONT_SIZES["small"]),
                command=self._toggle_notifications,
            )
            self.notifications_checkbox.grid(row=49, column=0, sticky="w", padx=16, pady=(0, 10))

            center = self._panel(self, 0)
            center.grid(row=1, column=1, sticky="nsew", padx=8, pady=(0, 12))
            center.grid_columnconfigure(0, weight=1)
            center.grid_rowconfigure(1, weight=0)
            center.grid_rowconfigure(5, weight=1)
            center.grid_rowconfigure(7, weight=1)
            self._section_title(center, UI_TEXT["brainz_status_title"], 0)
            mothership = ctk.CTkFrame(
                center,
                fg_color=COLORS["panel_alt"],
                border_color=COLORS["border"],
                border_width=1,
                corner_radius=6,
            )
            mothership.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 12))
            mothership.grid_columnconfigure(0, weight=1)
            mothership.grid_columnconfigure(1, weight=1)
            ctk.CTkLabel(
                mothership,
                textvariable=self.awake_status_var,
                text_color=COLORS["section"],
                font=(self.status_font_family, FONT_SIZES["body"]),
                anchor="w",
            ).grid(row=0, column=0, sticky="ew", padx=14, pady=(12, 2))
            self.memory_path_summary_var = ctk.StringVar(
                value=UI_TEXT["status_memory_path"].format(path=self.memory_var.get())
            )
            ctk.CTkLabel(
                mothership,
                textvariable=self.memory_path_summary_var,
                text_color=COLORS["muted"],
                font=(self.reading_font_family, FONT_SIZES["small"]),
                wraplength=330,
                justify="left",
                anchor="w",
            ).grid(row=1, column=0, sticky="ew", padx=14, pady=2)
            self.last_import_summary_var = ctk.StringVar(value=UI_TEXT["status_last_import_none"])
            ctk.CTkLabel(
                mothership,
                textvariable=self.last_import_summary_var,
                text_color=COLORS["muted"],
                font=(self.status_font_family, FONT_SIZES["small"]),
                anchor="w",
            ).grid(row=2, column=0, sticky="ew", padx=14, pady=(2, 12))
            ctk.CTkLabel(
                mothership,
                text=UI_TEXT["oikawa_bridge_message"],
                text_color=COLORS["quiet"],
                font=(self.reading_font_family, FONT_SIZES["small"]),
                wraplength=290,
                justify="left",
                anchor="w",
            ).grid(row=0, column=1, sticky="ew", padx=14, pady=(12, 4))
            mission_buttons = ctk.CTkFrame(mothership, fg_color="transparent")
            mission_buttons.grid(row=1, column=1, rowspan=2, sticky="ew", padx=14, pady=(0, 12))
            mission_buttons.grid_columnconfigure((0, 1), weight=1)
            ctk.CTkButton(
                mission_buttons,
                text=UI_TEXT["button_open_oikawa"],
                height=34,
                fg_color=COLORS["accent"],
                hover_color=COLORS["accent_hover"],
                font=(self.status_font_family, FONT_SIZES["button"]),
                command=self._open_oikawa,
            ).grid(row=0, column=0, sticky="ew", padx=(0, 6))
            ctk.CTkButton(
                mission_buttons,
                text=UI_TEXT["button_settings_entry"],
                height=34,
                fg_color=COLORS["panel_soft"],
                hover_color=COLORS["accent_soft"],
                font=(self.status_font_family, FONT_SIZES["button"]),
                command=self._mark_settings_entry,
            ).grid(row=0, column=1, sticky="ew", padx=(6, 0))

            self._section_title(center, UI_TEXT["search_bridge_title"], 2)
            search_box = ctk.CTkFrame(center, fg_color="transparent")
            search_box.grid(row=3, column=0, sticky="ew", padx=12, pady=(0, 10))
            search_box.grid_columnconfigure(0, weight=1)
            self.search_entry = ctk.CTkEntry(
                search_box,
                height=34,
                placeholder_text=UI_TEXT["search_placeholder"],
                fg_color=COLORS["input"],
                border_color=COLORS["border"],
                text_color=COLORS["text"],
                font=(self.reading_font_family, FONT_SIZES["body"]),
            )
            self.search_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
            self.search_entry.insert(0, self.current_query)
            self.search_entry.bind("<Return>", lambda _event: self._start_search())
            self.search_button = ctk.CTkButton(
                search_box,
                text=UI_TEXT["button_search"],
                width=96,
                height=34,
                fg_color=COLORS["panel_soft"],
                hover_color=COLORS["accent_soft"],
                font=(self.status_font_family, FONT_SIZES["button"]),
                command=self._start_search,
            )
            self.search_button.grid(row=0, column=1, padx=(0, 8))
            self.view_switch = ctk.CTkSegmentedButton(
                search_box,
                values=[UI_TEXT["tab_search"], UI_TEXT["tab_embers"]],
                variable=self.view_mode_var,
                fg_color=COLORS["panel_soft"],
                selected_color=COLORS["accent"],
                selected_hover_color=COLORS["accent_hover"],
                unselected_color=COLORS["panel_soft"],
                unselected_hover_color=COLORS["accent_soft"],
                text_color=COLORS["text"],
                font=(self.status_font_family, FONT_SIZES["micro"]),
                command=self._change_view_mode,
            )
            self.view_switch.grid(row=0, column=2)
            self.semantic_checkbox = ctk.CTkCheckBox(
                search_box,
                text=UI_TEXT["checkbox_semantic_search"],
                variable=self.semantic_search_var,
                text_color=COLORS["quiet"],
                fg_color=COLORS["accent"],
                hover_color=COLORS["accent_hover"],
                border_color=COLORS["border"],
                font=(self.status_font_family, FONT_SIZES["micro"]),
            )
            self.semantic_checkbox.grid(row=1, column=0, sticky="w", pady=(5, 0))
            self.search_helper_var = ctk.StringVar(value=UI_TEXT["search_bridge_helper"])
            ctk.CTkLabel(
                search_box,
                textvariable=self.search_helper_var,
                text_color=COLORS["quiet"],
                font=(self.reading_font_family, FONT_SIZES["micro"]),
                anchor="w",
            ).grid(row=1, column=1, columnspan=2, sticky="ew", pady=(5, 0))

            self._section_title(center, UI_TEXT["source_view_title"], 4)
            self.source_view_box = ctk.CTkTextbox(
                center,
                height=120,
                fg_color=COLORS["input"],
                border_color=COLORS["border"],
                border_width=1,
                text_color=COLORS["text"],
                font=(self.reading_font_family, FONT_SIZES["preview"]),
                wrap="word",
            )
            self.source_view_box.grid(row=5, column=0, sticky="nsew", padx=12, pady=(0, 10))
            self._relax_textbox_spacing(self.source_view_box, spacing1=3, spacing3=6)
            set_textbox_text(self.source_view_box, UI_TEXT["empty_source_view"])
            self.results_title_var = ctk.StringVar(value=UI_TEXT["results_title"])
            ctk.CTkLabel(
                center,
                textvariable=self.results_title_var,
                text_color=COLORS["section"],
                font=(self.font_family, FONT_SIZES["section"]),
            ).grid(row=6, column=0, sticky="w", padx=14, pady=(8, 7))
            self.results_frame = ctk.CTkScrollableFrame(
                center,
                fg_color=COLORS["panel"],
                corner_radius=0,
                scrollbar_button_color=COLORS["panel_soft"],
                scrollbar_button_hover_color=COLORS["accent_soft"],
            )
            self.results_frame.grid(row=7, column=0, sticky="nsew", padx=12, pady=(0, 12))
            self._render_empty_results()

            right = self._panel(self, 0)
            right.grid(row=1, column=2, sticky="nsew", padx=(8, 20), pady=(0, 12))
            right.grid_columnconfigure(0, weight=1)
            right.grid_rowconfigure(1, weight=2)
            right.grid_rowconfigure(8, weight=1)
            right.grid_rowconfigure(10, weight=1)
            self._section_title(right, UI_TEXT["preview_title"], 0)
            self.preview_box = ctk.CTkTextbox(
                right,
                fg_color=COLORS["input"],
                border_color=COLORS["border"],
                border_width=1,
                text_color=COLORS["text"],
                font=(self.reading_font_family, FONT_SIZES["preview"]),
                wrap="word",
            )
            self.preview_box.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))
            self._relax_textbox_spacing(self.preview_box, spacing1=3, spacing3=5)
            set_textbox_text(self.preview_box, UI_TEXT["empty_preview"])

            self._section_title(right, UI_TEXT["tags_title"], 2)
            self.tags_var = ctk.StringVar(value=UI_TEXT["empty_tags"])
            ctk.CTkLabel(
                right,
                textvariable=self.tags_var,
                text_color=COLORS["muted"],
                font=(self.reading_font_family, FONT_SIZES["small"]),
                wraplength=320,
                justify="left",
            ).grid(row=3, column=0, sticky="ew", padx=14, pady=(0, 12))

            self._section_title(right, UI_TEXT["section_related_memory"], 4)
            self.related_box = ctk.CTkTextbox(
                right,
                height=92,
                fg_color=COLORS["input"],
                border_color=COLORS["border"],
                border_width=1,
                text_color=COLORS["muted"],
                font=(self.reading_font_family, FONT_SIZES["body"]),
                wrap="word",
            )
            self.related_box.grid(row=5, column=0, sticky="ew", padx=12, pady=(0, 12))
            self._relax_textbox_spacing(self.related_box, spacing1=2, spacing3=5)
            set_textbox_text(self.related_box, UI_TEXT["related_memory_empty"])

            self._section_title(right, UI_TEXT["section_memory_flow"], 6)
            flow_header = ctk.CTkFrame(right, fg_color="transparent")
            flow_header.grid(row=7, column=0, sticky="ew", padx=10, pady=(0, 8))
            flow_header.grid_columnconfigure(0, weight=1)
            self.flow_status_var = ctk.StringVar(value=UI_TEXT["label_memory_flow"])
            ctk.CTkLabel(
                flow_header,
                textvariable=self.flow_status_var,
                text_color=COLORS["muted"],
                font=(self.reading_font_family, FONT_SIZES["small"]),
                anchor="w",
            ).grid(row=0, column=0, sticky="ew", padx=(2, 8))
            self.flow_sort_button = ctk.CTkButton(
                flow_header,
                text=UI_TEXT["button_timeline_ascending"],
                width=92,
                height=26,
                fg_color=COLORS["panel_soft"],
                hover_color=COLORS["accent_soft"],
                text_color=COLORS["muted"],
                font=(self.status_font_family, FONT_SIZES["micro"]),
                command=self._toggle_flow_sort,
            )
            self.flow_sort_button.grid(row=0, column=1, sticky="e")
            self.flow_frame = ctk.CTkScrollableFrame(
                right,
                height=150,
                fg_color=COLORS["input"],
                border_color=COLORS["border"],
                border_width=1,
                corner_radius=6,
                scrollbar_button_color=COLORS["panel_soft"],
                scrollbar_button_hover_color=COLORS["accent_soft"],
            )
            self.flow_frame.grid(row=8, column=0, sticky="nsew", padx=12, pady=(0, 12))
            self._render_empty_memory_flow()

            self._section_title(right, UI_TEXT["handoff_title"], 9)
            self.handoff_box = ctk.CTkTextbox(
                right,
                height=96,
                fg_color=COLORS["input"],
                border_color=COLORS["border"],
                border_width=1,
                text_color=COLORS["muted"],
                font=(self.reading_font_family, FONT_SIZES["body"]),
                wrap="word",
            )
            self.handoff_box.grid(row=10, column=0, sticky="nsew", padx=12, pady=(0, 12))
            self._relax_textbox_spacing(self.handoff_box, spacing1=2, spacing3=5)
            set_textbox_text(self.handoff_box, UI_TEXT["empty_handoff"])

            export_row = ctk.CTkFrame(right, fg_color="transparent")
            export_row.grid(row=11, column=0, sticky="ew", padx=12, pady=(0, 12))
            export_row.grid_columnconfigure((0, 1), weight=1)
            ctk.CTkButton(
                export_row,
                text=UI_TEXT["button_chatgpt"],
                height=36,
                fg_color=COLORS["panel_soft"],
                hover_color=COLORS["accent_soft"],
                font=(self.font_family, FONT_SIZES["button"]),
                command=lambda: self._export_handoff("chatgpt"),
            ).grid(row=0, column=0, sticky="ew", padx=(0, 6))
            ctk.CTkButton(
                export_row,
                text=UI_TEXT["button_codex"],
                height=36,
                fg_color=COLORS["panel_soft"],
                hover_color=COLORS["accent_soft"],
                font=(self.font_family, FONT_SIZES["button"]),
                command=lambda: self._export_handoff("codex"),
            ).grid(row=0, column=1, sticky="ew", padx=(6, 0))

            bottom = self._panel(self, 0)
            bottom.grid(row=2, column=0, columnspan=3, sticky="ew", padx=20, pady=(0, 16))
            bottom.grid_columnconfigure(0, weight=1)
            bottom.grid_columnconfigure(1, weight=0, minsize=300)
            self._section_title(bottom, UI_TEXT["log_title"], 0)
            self.log_box = ctk.CTkTextbox(
                bottom,
                height=96,
                fg_color=COLORS["input"],
                text_color=COLORS["log_text"],
                font=(self.mono_font_family, FONT_SIZES["log"]),
                wrap="word",
            )
            self.log_box.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 12))
            self._relax_textbox_spacing(self.log_box, spacing1=2, spacing3=4)
            self.log_box.configure(state="disabled")

            self.notification_frame = ctk.CTkFrame(
                bottom,
                fg_color=COLORS["notification"],
                border_color=COLORS["notification_border"],
                border_width=1,
                corner_radius=6,
            )
            self.notification_frame.grid(row=0, column=1, rowspan=2, sticky="sew", padx=(0, 12), pady=(14, 12))
            self.notification_frame.grid_columnconfigure(0, weight=1)
            self.notification_title_var = ctk.StringVar(value=UI_TEXT["notify_title"])
            self.notification_body_var = ctk.StringVar(value="")
            ctk.CTkLabel(
                self.notification_frame,
                textvariable=self.notification_title_var,
                text_color=COLORS["section"],
                font=(self.font_family, FONT_SIZES["small"]),
                anchor="w",
            ).grid(row=0, column=0, sticky="ew", padx=14, pady=(10, 2))
            ctk.CTkLabel(
                self.notification_frame,
                textvariable=self.notification_body_var,
                text_color=COLORS["text"],
                font=(self.reading_font_family, FONT_SIZES["body"]),
                wraplength=260,
                justify="left",
                anchor="w",
            ).grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 12))
            self.notification_frame.grid_remove()

        def _panel(self, parent, corner_radius: int):
            return ctk.CTkFrame(
                parent,
                fg_color=COLORS["panel"],
                border_color=COLORS["border"],
                border_width=1,
                corner_radius=corner_radius,
            )

        def _section_title(self, parent, text: str, row: int) -> None:
            ctk.CTkLabel(
                parent,
                text=text,
                text_color=COLORS["section"],
                font=(self.font_family, FONT_SIZES["section"]),
            ).grid(row=row, column=0, sticky="w", padx=14, pady=(14, 7))

        def _relax_textbox_spacing(self, textbox, spacing1: int = 2, spacing3: int = 4) -> None:
            try:
                textbox._textbox.configure(spacing1=spacing1, spacing2=1, spacing3=spacing3)
            except Exception:
                pass

        def _render_empty_results(self) -> None:
            for child in self.results_frame.winfo_children():
                child.destroy()
            ctk.CTkLabel(
                self.results_frame,
                text=UI_TEXT["empty_results"],
                text_color=COLORS["quiet"],
                font=(self.reading_font_family, FONT_SIZES["body"]),
            ).pack(fill="x", padx=12, pady=18)

        def _render_empty_memory_flow(self) -> None:
            for child in self.flow_frame.winfo_children():
                child.destroy()
            ctk.CTkLabel(
                self.flow_frame,
                text=UI_TEXT["memory_flow_empty"],
                text_color=COLORS["quiet"],
                font=(self.reading_font_family, FONT_SIZES["small"]),
                wraplength=310,
            ).pack(fill="x", padx=10, pady=12)

        def _append_log(self, message: str) -> None:
            self.log_box.configure(state="normal")
            self.log_box.insert("end", f"{message}\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")

        def _set_memory_folder(self, folder: str, persist: bool = True) -> None:
            clean = str(folder or "")
            self.config_data.memory_folder = clean
            self.memory_var.set(clean if clean else UI_TEXT["empty_memory"])
            if hasattr(self, "memory_path_summary_var"):
                path_text = clean if clean else UI_TEXT["empty_memory"]
                self.memory_path_summary_var.set(UI_TEXT["status_memory_path"].format(path=path_text))
            if clean and not self.config_data.watch_folder:
                self._set_watch_folder(clean, persist=False)
            if clean and not self.config_data.codex_reports_folder:
                self.config_data.codex_reports_folder = str(Path(clean) / "codex_reports")
                self._update_codex_report_status()
            self._update_slack_status()
            self._update_aru_status()
            if persist:
                self.config_store.save(self.config_data)
                if clean:
                    self._append_log(UI_TEXT["log_folder_set"].format(folder=clean))

        def _choose_memory_folder(self) -> None:
            folder = filedialog.askdirectory(title=UI_TEXT["choose_memory_title"])
            if folder:
                self._set_memory_folder(folder)

        def _set_watch_folder(self, folder: str, persist: bool = True) -> None:
            clean = str(folder or "")
            self.config_data.watch_folder = clean
            self.watch_folder_var.set(clean if clean else UI_TEXT["empty_watch_folder"])
            self._update_watch_status()
            if persist:
                self.config_store.save(self.config_data)
                if clean:
                    self._append_log(UI_TEXT["log_watch_initialized"])
                    self._append_log(UI_TEXT["log_watching_folder"].format(folder=clean))

        def _choose_watch_folder(self) -> None:
            folder = filedialog.askdirectory(title=UI_TEXT["label_watch_folder"])
            if folder:
                self._set_watch_folder(folder)

        def _set_remote_queue_folder(self, folder: str, persist: bool = True) -> None:
            clean = str(folder or "")
            self.config_data.remote_queue_folder = clean
            self.remote_queue_folder_var.set(clean if clean else UI_TEXT["empty_remote_queue_folder"])
            self._update_remote_queue_status()
            if persist:
                self.config_store.save(self.config_data)

        def _choose_remote_queue_folder(self) -> None:
            folder = filedialog.askdirectory(title=UI_TEXT["section_remote_queue"])
            if folder:
                self._set_remote_queue_folder(folder)

        def _toggle_remote_queue(self) -> None:
            self.config_data.enable_remote_queue = bool(self.remote_queue_var.get())
            self.config_store.save(self.config_data)
            self._update_remote_queue_status()

        def _save_slack_settings(self) -> None:
            token = self.slack_token_entry.get().strip() if hasattr(self, "slack_token_entry") else ""
            channel_id = self.slack_channel_entry.get().strip() if hasattr(self, "slack_channel_entry") else ""
            interval_text = self.slack_interval_entry.get().strip() if hasattr(self, "slack_interval_entry") else "10"
            interval, interval_valid = parse_slack_poll_interval(interval_text)
            enabled = bool(self.slack_inbox_var.get())
            if enabled:
                missing = slack_config_missing_fields(
                    memory_folder=self.config_data.memory_folder,
                    token=token,
                    channel_id=channel_id,
                    interval_valid=interval_valid,
                )
            else:
                missing = [] if interval_valid else ["poll_interval"]
            try:
                self.config_data.enable_slack_inbox = enabled
                self.config_data.slack_bot_token = token
                self.config_data.slack_channel_id = channel_id
                self.config_data.slack_poll_interval_seconds = interval
                self.slack_poll_interval_ms = interval * 1000
                if hasattr(self, "slack_interval_entry"):
                    self.slack_interval_entry.delete(0, "end")
                    self.slack_interval_entry.insert(0, str(interval))
                self.config_store.save(self.config_data)
                self.config_data = self.config_store.load()
                self.slack_inbox_var.set(self.config_data.enable_slack_inbox)
                self.slack_poll_interval_ms = max(5000, int(self.config_data.slack_poll_interval_seconds) * 1000)
                self._append_log(
                    UI_TEXT["log_slack_config_reloaded"].format(
                        enabled=self.config_data.enable_slack_inbox,
                        channel=self.config_data.slack_channel_id or "-",
                        interval=self.config_data.slack_poll_interval_seconds,
                    )
                )
                self._update_slack_status()
                if missing:
                    self.slack_status_var.set(UI_TEXT["status_slack_config_save_failed"])
                    self._append_log(UI_TEXT["log_slack_config_save_failed"].format(missing=", ".join(missing)))
                else:
                    self.slack_status_var.set(UI_TEXT["status_slack_config_saved"])
                    self._append_log(
                        UI_TEXT["log_slack_config_saved"].format(
                            enabled=self.config_data.enable_slack_inbox,
                            channel=self.config_data.slack_channel_id,
                            interval=self.config_data.slack_poll_interval_seconds,
                        )
                    )
                if self.config_data.enable_slack_inbox and not missing:
                    self._start_slack_poll_worker_if_ready()
            except Exception as exc:
                self.slack_status_var.set(UI_TEXT["status_slack_config_save_failed"])
                self._append_log(UI_TEXT["log_slack_config_save_failed"].format(missing=str(exc)))

        def _toggle_slack_inbox(self) -> None:
            self._save_slack_settings()

        def _save_aru_settings(self) -> None:
            token = self.aru_token_entry.get().strip() if hasattr(self, "aru_token_entry") else ""
            channel_id = self.aru_channel_entry.get().strip() if hasattr(self, "aru_channel_entry") else ""
            enabled = bool(self.aru_inbox_var.get())
            missing = []
            if enabled:
                missing = slack_config_missing_fields(
                    memory_folder=self.config_data.memory_folder,
                    token=token,
                    channel_id=channel_id,
                    interval_valid=True,
                )
            try:
                self.config_data.enable_aru_inbox = enabled
                self.config_data.aru_slack_token = token
                self.config_data.aru_channel_id = channel_id
                self.config_data.aru_poll_interval_seconds = max(5, int(self.config_data.aru_poll_interval_seconds or 10))
                self.config_store.save(self.config_data)
                self.config_data = self.config_store.load()
                self.aru_inbox_var.set(self.config_data.enable_aru_inbox)
                self.aru_poll_interval_ms = max(5000, int(self.config_data.aru_poll_interval_seconds) * 1000)
                self._append_log(
                    UI_TEXT["log_aru_config_reloaded"].format(
                        enabled=self.config_data.enable_aru_inbox,
                        channel=self.config_data.aru_channel_id or "-",
                    )
                )
                self._update_aru_status()
                if missing:
                    self.aru_status_var.set(UI_TEXT["status_aru_config_save_failed"])
                    self._append_log(UI_TEXT["log_aru_config_save_failed"].format(missing=", ".join(missing)))
                else:
                    self.aru_status_var.set(UI_TEXT["status_aru_config_saved"])
                    self._append_log(
                        UI_TEXT["log_aru_config_saved"].format(
                            enabled=self.config_data.enable_aru_inbox,
                            channel=self.config_data.aru_channel_id,
                        )
                    )
                if self.config_data.enable_aru_inbox and not missing:
                    self._start_aru_poll_worker_if_ready()
            except Exception as exc:
                self.aru_status_var.set(UI_TEXT["status_aru_config_save_failed"])
                self._append_log(UI_TEXT["log_aru_config_save_failed"].format(missing=str(exc)))

        def _toggle_aru_inbox(self) -> None:
            self._save_aru_settings()

        def _toggle_auto_index(self) -> None:
            self.config_data.auto_index_enabled = bool(self.auto_index_var.get())
            self.config_store.save(self.config_data)
            self._update_watch_status()
            if self.config_data.auto_index_enabled and self.config_data.watch_folder:
                self._append_log(UI_TEXT["log_watch_initialized"])
                self._append_log(UI_TEXT["log_watching_folder"].format(folder=self.config_data.watch_folder))

        def _toggle_notifications(self) -> None:
            self.config_data.enable_notifications = bool(self.notifications_var.get())
            self.config_store.save(self.config_data)
            if not self.config_data.enable_notifications:
                self.notification_frame.grid_remove()
                self.notification_visible = False

        def _mark_settings_entry(self) -> None:
            self._append_log(UI_TEXT["log_settings_entry"])
            self._notify(UI_TEXT["settings_entry_hint"])

        def _open_oikawa(self) -> None:
            command, target, mode = resolve_oikawa_launch_command()
            if not command or target is None:
                self._append_log(UI_TEXT["log_oikawa_missing"])
                self._notify(UI_TEXT["notify_oikawa_missing"])
                return
            try:
                cwd = target.parent.parent if target.parent.name.lower() == "dist" else target.parent
                subprocess.Popen(
                    command,
                    cwd=str(cwd),
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    stdin=subprocess.DEVNULL,
                    close_fds=True,
                )
            except Exception:
                self._append_log(UI_TEXT["log_oikawa_missing"])
                self._notify(UI_TEXT["notify_oikawa_missing"])
                return
            self._append_log(UI_TEXT["log_oikawa_open"].format(mode=mode, path=target))
            self._notify(UI_TEXT["notify_oikawa_opened"])

        def _set_last_import_summary(self, source_key: str) -> None:
            if hasattr(self, "last_import_summary_var"):
                source = UI_TEXT[source_key]
                self.last_import_summary_var.set(UI_TEXT["status_last_import"].format(source=source))

        def _notify(self, message: str) -> None:
            if not self.notifications_var.get():
                return
            self.notification_queue.push(UI_TEXT["notify_title"], message)
            self._append_log(UI_TEXT["log_notification_sent"].format(message=message))
            if not self.notification_visible:
                self._show_next_notification()

        def _show_next_notification(self) -> None:
            if not self.notifications_var.get():
                self.notification_visible = False
                self.notification_frame.grid_remove()
                return
            item = self.notification_queue.pop()
            if item is None:
                self.notification_visible = False
                self.notification_frame.grid_remove()
                return
            self.notification_visible = True
            self.notification_title_var.set(item.title)
            self.notification_body_var.set(item.message)
            self.notification_frame.grid()
            self.after(4200, lambda current=item: self._finish_notification(current))

        def _finish_notification(self, item: NotificationItem) -> None:
            self.notification_queue.remember(item)
            self.notification_frame.grid_remove()
            self.notification_visible = False
            if self.notification_queue.pending_count:
                self.after(250, self._show_next_notification)

        def _update_watch_status(self) -> None:
            if not hasattr(self, "watch_folder_var"):
                return
            folder = self.config_data.watch_folder
            self.watch_folder_var.set(folder if folder else UI_TEXT["empty_watch_folder"])
            if not folder:
                self.watch_status_var.set(UI_TEXT["status_watch_folder_missing"])
            elif self.auto_index_var.get():
                self.watch_status_var.set(f"{UI_TEXT['status_auto_index_on']} / {UI_TEXT['status_watching']}")
            else:
                self.watch_status_var.set(UI_TEXT["status_auto_index_off"])
            last_indexed_at = self.config_data.last_indexed_at or "-"
            self.last_index_var.set(UI_TEXT["status_last_index"].format(time=last_indexed_at))

        def _update_remote_queue_status(self) -> None:
            if not hasattr(self, "remote_queue_folder_var"):
                return
            folder = self.config_data.remote_queue_folder
            self.remote_queue_folder_var.set(folder if folder else UI_TEXT["empty_remote_queue_folder"])
            if not folder:
                self.remote_queue_status_var.set(UI_TEXT["status_remote_queue_folder_missing"])
                pending = 0
            elif self.remote_queue_var.get():
                self.remote_queue_status_var.set(UI_TEXT["status_remote_queue_watching"])
                pending = count_pending_queue_files(Path(folder))
            else:
                self.remote_queue_status_var.set(UI_TEXT["status_remote_queue_off"])
                pending = count_pending_queue_files(Path(folder)) if Path(folder).exists() else 0
            stats = self.database.remote_queue_stats()
            self.remote_queue_counts_var.set(
                UI_TEXT["status_remote_queue_counts"].format(
                    pending=pending,
                    processed=stats["processed"],
                    failed=stats["failed"],
                )
            )

        def _codex_reports_folder(self) -> Path | None:
            folder = self.config_data.codex_reports_folder
            if not folder and self.config_data.memory_folder:
                folder = str(Path(self.config_data.memory_folder) / "codex_reports")
                self.config_data.codex_reports_folder = folder
            return Path(folder) if folder else None

        def _update_codex_report_status(self) -> None:
            if not hasattr(self, "codex_report_status_var"):
                return
            folder = self._codex_reports_folder()
            if folder:
                self.codex_report_status_var.set(UI_TEXT["status_codex_report_watching"])
                pending = count_pending_reports(folder) if folder.exists() else 0
                if pending:
                    self.codex_report_status_var.set(f"{UI_TEXT['status_codex_report_watching']} / {pending}")
            else:
                self.codex_report_status_var.set(UI_TEXT["status_codex_report_idle"])

        def _update_slack_status(self) -> None:
            if not hasattr(self, "slack_status_var"):
                return
            channel_id = self.config_data.slack_channel_id or "-"
            self.slack_channel_var.set(UI_TEXT["label_slack_channel_status"].format(channel=channel_id))
            if not self.slack_inbox_var.get():
                self.slack_status_var.set(UI_TEXT["status_slack_disabled"])
            elif not self.config_data.memory_folder or not self.config_data.slack_bot_token or not self.config_data.slack_channel_id:
                self.slack_status_var.set(UI_TEXT["status_slack_config_missing"])
            elif self.slack_status_var.get() in {UI_TEXT["status_slack_disabled"], UI_TEXT["status_slack_config_missing"]}:
                self.slack_status_var.set(UI_TEXT["status_slack_ready"])

        def _update_aru_status(self) -> None:
            if not hasattr(self, "aru_status_var"):
                return
            channel_id = self.config_data.aru_channel_id or "-"
            self.aru_channel_var.set(UI_TEXT["label_aru_channel_status"].format(channel=channel_id))
            if not self.aru_inbox_var.get():
                self.aru_status_var.set(UI_TEXT["status_aru_disabled"])
            elif not self.config_data.memory_folder or not self.config_data.aru_slack_token or not self.config_data.aru_channel_id:
                self.aru_status_var.set(UI_TEXT["status_aru_config_missing"])
            elif self.aru_status_var.get() in {UI_TEXT["status_aru_disabled"], UI_TEXT["status_aru_config_missing"]}:
                self.aru_status_var.set(UI_TEXT["status_aru_ready"])

        def _choose_chatgpt_export(self) -> None:
            self._choose_chatgpt_export_zip(fallback_to_folder=True)

        def _choose_chatgpt_export_zip(self, fallback_to_folder: bool = False) -> None:
            file_path = filedialog.askopenfilename(
                title=UI_TEXT["dialog_select_chatgpt_export"],
                filetypes=[
                    (UI_TEXT["filetype_zip"], "*.zip"),
                    (UI_TEXT["filetype_all"], "*.*"),
                ],
            )
            if file_path:
                self._start_chatgpt_import(Path(file_path))
            elif fallback_to_folder:
                self._choose_chatgpt_export_folder()

        def _choose_chatgpt_export_folder(self) -> None:
            folder_path = filedialog.askdirectory(title=UI_TEXT["dialog_select_chatgpt_export_folder"])
            if folder_path:
                self._start_chatgpt_import(Path(folder_path))

        def _choose_chatgpt_conversations_json(self) -> None:
            file_path = filedialog.askopenfilename(
                title=UI_TEXT["dialog_select_chatgpt_conversations_json"],
                filetypes=[
                    (UI_TEXT["filetype_json"], "*.json"),
                    (UI_TEXT["filetype_all"], "*.*"),
                ],
            )
            if file_path:
                self._start_chatgpt_import(Path(file_path))

        def _start_chatgpt_import(self, source_path: Path) -> None:
            if self.import_thread and self.import_thread.is_alive():
                return
            self._set_import_buttons_state("disabled")
            self.index_status_var.set(UI_TEXT["status_importing_chatgpt"])
            if hasattr(self, "chatgpt_import_status_var"):
                self.chatgpt_import_status_var.set(UI_TEXT["chatgpt_import_importing"])
            self.progress.set(0)
            self._append_log(UI_TEXT["log_chatgpt_export_detected"])
            self.import_thread = threading.Thread(target=self._chatgpt_import_worker, args=(source_path,), daemon=True)
            self.import_thread.start()

        def _chatgpt_import_worker(self, source_path: Path) -> None:
            try:
                result = import_chatgpt_export(source_path, self.database)
                self.events.put(("chatgpt_import_done", result))
            except ConversationsJsonNotFound as exc:
                self.events.put(("chatgpt_import_missing", str(exc)))
            except Exception as exc:
                self.events.put(("chatgpt_import_error", str(exc)))

        def _open_codex_import_dialog(self) -> None:
            dialog = ctk.CTkToplevel(self)
            dialog.title(UI_TEXT["dialog_import_codex_result"])
            dialog.geometry("720x520")
            dialog.transient(self)
            dialog.grab_set()
            dialog.grid_columnconfigure(0, weight=1)
            dialog.grid_rowconfigure(1, weight=1)

            ctk.CTkLabel(
                dialog,
                text=UI_TEXT["label_codex_result_input"],
                text_color=COLORS["section"],
                font=(self.font_family, FONT_SIZES["section"]),
            ).grid(row=0, column=0, sticky="w", padx=14, pady=(14, 8))
            input_box = ctk.CTkTextbox(
                dialog,
                fg_color=COLORS["input"],
                border_color=COLORS["border"],
                border_width=1,
                text_color=COLORS["text"],
                font=(self.reading_font_family, FONT_SIZES["body"]),
                wrap="word",
            )
            input_box.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 12))
            self._relax_textbox_spacing(input_box)

            button_row = ctk.CTkFrame(dialog, fg_color="transparent")
            button_row.grid(row=2, column=0, sticky="ew", padx=14, pady=(0, 14))
            button_row.grid_columnconfigure((0, 1, 2), weight=1)

            def import_paste() -> None:
                text_value = input_box.get("1.0", "end").strip()
                if not text_value:
                    messagebox.showwarning(UI_TEXT["dialog_title"], UI_TEXT["error_codex_result_empty"])
                    return
                dialog.destroy()
                self._start_codex_import_text(text_value)

            def import_file() -> None:
                path = filedialog.askopenfilename(
                    title=UI_TEXT["dialog_select_codex_result_file"],
                    filetypes=[
                        (UI_TEXT["filetype_text_markdown"], "*.txt *.md"),
                        (UI_TEXT["filetype_all"], "*.*"),
                    ],
                )
                if path:
                    dialog.destroy()
                    self._start_codex_import_file(Path(path))

            ctk.CTkButton(
                button_row,
                text=UI_TEXT["button_import_pasted_codex"],
                height=36,
                fg_color=COLORS["accent"],
                hover_color=COLORS["accent_hover"],
                font=(self.font_family, FONT_SIZES["button"]),
                command=import_paste,
            ).grid(row=0, column=0, sticky="ew", padx=(0, 6))
            ctk.CTkButton(
                button_row,
                text=UI_TEXT["button_import_codex_file"],
                height=36,
                fg_color=COLORS["panel_soft"],
                hover_color=COLORS["accent_soft"],
                font=(self.font_family, FONT_SIZES["button"]),
                command=import_file,
            ).grid(row=0, column=1, sticky="ew", padx=6)
            ctk.CTkButton(
                button_row,
                text=UI_TEXT["button_close"],
                height=36,
                fg_color=COLORS["panel_soft"],
                hover_color=COLORS["accent_soft"],
                font=(self.font_family, FONT_SIZES["button"]),
                command=dialog.destroy,
            ).grid(row=0, column=2, sticky="ew", padx=(6, 0))

        def _set_import_buttons_state(self, state: str) -> None:
            for button_name in (
                "import_button",
                "chatgpt_folder_button",
                "chatgpt_json_button",
                "codex_import_button",
            ):
                button = getattr(self, button_name, None)
                if button is not None:
                    button.configure(state=state)

        def _start_codex_import_text(self, text_value: str) -> None:
            if self.import_thread and self.import_thread.is_alive():
                return
            self._set_import_buttons_state("disabled")
            self.index_status_var.set(UI_TEXT["status_importing_codex"])
            self.progress.set(0)
            self._append_log(UI_TEXT["log_codex_result_detected"])
            self.import_thread = threading.Thread(target=self._codex_import_text_worker, args=(text_value,), daemon=True)
            self.import_thread.start()

        def _start_codex_import_file(self, source_path: Path) -> None:
            if self.import_thread and self.import_thread.is_alive():
                return
            self._set_import_buttons_state("disabled")
            self.index_status_var.set(UI_TEXT["status_importing_codex"])
            self.progress.set(0)
            self._append_log(UI_TEXT["log_codex_result_detected"])
            self.import_thread = threading.Thread(target=self._codex_import_file_worker, args=(source_path,), daemon=True)
            self.import_thread.start()

        def _codex_import_text_worker(self, text_value: str) -> None:
            try:
                result = import_codex_text(text_value, self.database, UI_TEXT["codex_source_paste"])
                self.events.put(("codex_import_done", result))
            except ValueError as exc:
                self.events.put(("codex_import_empty", str(exc)))
            except Exception as exc:
                self.events.put(("codex_import_error", str(exc)))

        def _codex_import_file_worker(self, source_path: Path) -> None:
            try:
                result = import_codex_file(source_path, self.database)
                self.events.put(("codex_import_done", result))
            except ValueError as exc:
                self.events.put(("codex_import_empty", str(exc)))
            except Exception as exc:
                self.events.put(("codex_import_error", str(exc)))

        def _start_index(self) -> None:
            folder = self.config_data.memory_folder
            if not folder:
                messagebox.showwarning(UI_TEXT["dialog_title"], UI_TEXT["dialog_memory_missing"])
                return
            memory_folder = Path(folder)
            if not memory_folder.exists():
                messagebox.showwarning(UI_TEXT["dialog_title"], UI_TEXT["dialog_memory_missing"])
                return
            if self.index_thread and self.index_thread.is_alive():
                return

            self.cancel_event.clear()
            self.progress.set(0)
            self.index_button.configure(state="disabled")
            self.cancel_button.configure(state="normal")
            self.index_status_var.set(UI_TEXT["index_running"].format(current=0, total=0))
            self._append_log(UI_TEXT["log_index_started"].format(folder=folder))
            self.index_thread = threading.Thread(target=self._index_worker, args=(memory_folder,), daemon=True)
            self.index_thread.start()

        def _index_worker(self, memory_folder: Path) -> None:
            try:
                self.indexer.run(memory_folder, self.cancel_event, lambda progress: self.events.put(("index_progress", progress)))
            except Exception as exc:
                self.events.put(("index_error", str(exc)))

        def _poll_watch_folder(self) -> None:
            try:
                if self.auto_index_var.get() and self.config_data.watch_folder:
                    watch_folder = Path(self.config_data.watch_folder)
                    if watch_folder.exists() and watch_folder.is_dir():
                        if not self.watch_scan_thread or not self.watch_scan_thread.is_alive():
                            self.watch_scan_thread = threading.Thread(
                                target=self._watch_scan_worker,
                                args=(watch_folder,),
                                daemon=True,
                            )
                            self.watch_scan_thread.start()
            finally:
                self.after(self.watch_poll_interval_ms, self._poll_watch_folder)

        def _watch_scan_worker(self, watch_folder: Path) -> None:
            try:
                result = detect_changed_files(self.database, watch_folder)
                self.events.put(("watch_scan_done", result))
            except Exception as exc:
                self.events.put(("watch_scan_error", str(exc)))

        def _poll_remote_queue(self) -> None:
            try:
                if self.remote_queue_var.get() and self.config_data.remote_queue_folder:
                    queue_folder = Path(self.config_data.remote_queue_folder)
                    if queue_folder.exists() and queue_folder.is_dir():
                        if not self.remote_queue_thread or not self.remote_queue_thread.is_alive():
                            self.remote_queue_thread = threading.Thread(
                                target=self._remote_queue_worker,
                                args=(queue_folder,),
                                daemon=True,
                            )
                            self.remote_queue_thread.start()
            finally:
                self.after(self.remote_queue_poll_interval_ms, self._poll_remote_queue)

        def _remote_queue_worker(self, queue_folder: Path) -> None:
            try:
                result = process_remote_queue_folder(self.database, queue_folder)
                self.events.put(("remote_queue_done", result))
            except Exception as exc:
                self.events.put(("remote_queue_error", str(exc)))

        def _poll_codex_reports(self) -> None:
            try:
                report_folder = self._codex_reports_folder()
                if report_folder:
                    if not self.codex_report_thread or not self.codex_report_thread.is_alive():
                        self.codex_report_thread = threading.Thread(
                            target=self._codex_report_worker,
                            args=(report_folder,),
                            daemon=True,
                        )
                        self.codex_report_thread.start()
            finally:
                self.after(self.codex_report_poll_interval_ms, self._poll_codex_reports)

        def _codex_report_worker(self, report_folder: Path) -> None:
            try:
                result = process_codex_reports_folder(self.database, report_folder)
                self.events.put(("codex_report_done", result))
            except Exception as exc:
                self.events.put(("codex_report_error", str(exc)))

        def _poll_slack_inbox(self) -> None:
            try:
                self._start_slack_poll_worker_if_ready()
            finally:
                self.after(self.slack_poll_interval_ms, self._poll_slack_inbox)

        def _start_slack_poll_worker_if_ready(self) -> None:
            if not self.config_data.enable_slack_inbox:
                return
            if not self.config_data.memory_folder:
                return
            if not self.config_data.slack_bot_token or not self.config_data.slack_channel_id:
                return
            if self.slack_thread and self.slack_thread.is_alive():
                return
            self.slack_thread = threading.Thread(
                target=self._slack_inbox_worker,
                args=(
                    Path(self.config_data.memory_folder),
                    self.config_data.slack_bot_token,
                    self.config_data.slack_channel_id,
                    self.config_data.slack_last_ts,
                ),
                daemon=True,
            )
            self.slack_thread.start()

        def _slack_inbox_worker(self, memory_folder: Path, token: str, channel_id: str, last_ts: str) -> None:
            try:
                result = poll_slack_inbox(
                    database=self.database,
                    memory_folder=memory_folder,
                    token=token,
                    channel_id=channel_id,
                    last_ts=last_ts,
                    poll_timeout_seconds=8.0,
                )
                self.events.put(("slack_inbox_done", result))
            except Exception as exc:
                self.events.put(("slack_inbox_error", str(exc)))

        def _poll_aru_inbox(self) -> None:
            try:
                self._start_aru_poll_worker_if_ready()
            finally:
                self.after(self.aru_poll_interval_ms, self._poll_aru_inbox)

        def _start_aru_poll_worker_if_ready(self) -> None:
            if not self.config_data.enable_aru_inbox:
                return
            if not self.config_data.memory_folder:
                return
            if not self.config_data.aru_slack_token or not self.config_data.aru_channel_id:
                return
            if self.aru_thread and self.aru_thread.is_alive():
                return
            self.aru_thread = threading.Thread(
                target=self._aru_inbox_worker,
                args=(
                    Path(self.config_data.memory_folder),
                    self.config_data.aru_slack_token,
                    self.config_data.aru_channel_id,
                    self.config_data.aru_last_ts,
                ),
                daemon=True,
            )
            self.aru_thread.start()

        def _aru_inbox_worker(self, memory_folder: Path, token: str, channel_id: str, last_ts: str) -> None:
            try:
                result = poll_slack_inbox(
                    database=self.database,
                    memory_folder=memory_folder,
                    token=token,
                    channel_id=channel_id,
                    last_ts=last_ts,
                    poll_timeout_seconds=8.0,
                    source_type="aru",
                    folder_name="aru",
                    inbox_label="Aru Inbox",
                    process_tasks=False,
                )
                self.events.put(("aru_inbox_done", result))
            except Exception as exc:
                self.events.put(("aru_inbox_error", str(exc)))

        def _start_auto_index(self, memory_folder: Path) -> None:
            if self.index_thread and self.index_thread.is_alive():
                self.pending_auto_index_folder = memory_folder
                self._append_log(UI_TEXT["log_auto_index_queued"])
                return
            if not memory_folder.exists():
                return
            self.auto_index_active = True
            self.cancel_event.clear()
            self.progress.set(0)
            self.index_button.configure(state="disabled")
            self.cancel_button.configure(state="normal")
            self.index_status_var.set(UI_TEXT["index_running"].format(current=0, total=0))
            self.index_thread = threading.Thread(target=self._index_worker, args=(memory_folder,), daemon=True)
            self.index_thread.start()

        def _write_watch_event_log(self, lines: list[str]) -> None:
            try:
                path = write_watch_log(lines)
            except OSError:
                return
            self._append_log(UI_TEXT["log_watch_file"].format(path=path))

        def _cancel_index(self) -> None:
            self.cancel_event.set()
            self._append_log(UI_TEXT["log_index_cancel_requested"])

        def _start_search(self) -> None:
            query_text = self.search_entry.get().strip()
            if not query_text and self.embers_mode:
                query_text = default_ember_query()
            if not query_text:
                return
            if self.search_thread and self.search_thread.is_alive():
                return
            self.current_query = query_text
            self.config_data.last_query = query_text
            self.config_store.save(self.config_data)
            self._set_search_running(True)
            self.index_status_var.set(UI_TEXT["status_searching"])
            self._append_log(UI_TEXT["log_searching"].format(query=query_text))
            self._append_log(UI_TEXT["phrase_searching_memory"])
            if self.embers_mode:
                self._append_log(UI_TEXT["log_embers_search"].format(query=query_text))
            self.update_idletasks()
            semantic_enabled = bool(self.semantic_search_var.get()) and self.semantic_available
            self.search_thread = threading.Thread(
                target=self._search_worker,
                args=(query_text, semantic_enabled),
                daemon=True,
            )
            self.search_thread.start()

        def _search_worker(self, query_text: str, semantic_enabled: bool) -> None:
            try:
                response = self.search_engine.search_with_related(query_text, semantic_enabled=semantic_enabled)
                self.events.put(("search_done", (query_text, response, semantic_enabled)))
            except Exception as exc:
                self.events.put(("index_error", str(exc)))

        def _set_search_running(self, running: bool) -> None:
            if running:
                self.search_button.configure(text=UI_TEXT["button_searching"], state="disabled")
            else:
                self.search_button.configure(text=UI_TEXT["button_search"], state="normal")

        def _change_view_mode(self, value: str) -> None:
            self.embers_mode = value == UI_TEXT["tab_embers"]
            if hasattr(self, "results_title_var"):
                title = UI_TEXT["embers_results_title"] if self.embers_mode else UI_TEXT["results_title"]
                self.results_title_var.set(title)
            if hasattr(self, "search_helper_var"):
                helper = UI_TEXT["embers_search_helper"] if self.embers_mode else UI_TEXT["search_bridge_helper"]
                self.search_helper_var.set(helper)
            if hasattr(self, "search_entry"):
                placeholder = UI_TEXT["embers_search_placeholder"] if self.embers_mode else UI_TEXT["search_placeholder"]
                self.search_entry.configure(placeholder_text=placeholder)

        def _render_results(self, results: list[SearchResult]) -> None:
            for child in self.results_frame.winfo_children():
                child.destroy()
            if not results:
                self._render_empty_results()
                return

            for result in results:
                item = ctk.CTkFrame(self.results_frame, fg_color=COLORS["panel_alt"], corner_radius=6)
                item.pack(fill="x", padx=5, pady=6)
                item.grid_columnconfigure(0, weight=1)
                title = ctk.CTkButton(
                    item,
                    text=self._result_title(result),
                    anchor="w",
                    fg_color="transparent",
                    hover_color=COLORS["accent_soft"],
                    text_color=COLORS["text"],
                    font=(self.reading_font_family, FONT_SIZES["result_title"]),
                    command=lambda selected=result: self._select_result(selected),
                )
                title.grid(row=0, column=0, sticky="ew", padx=10, pady=(9, 2))
                ctk.CTkLabel(
                    item,
                    text=self._result_meta(result),
                    text_color=COLORS["quiet"],
                    font=(self.status_font_family, FONT_SIZES["result_meta"]),
                    anchor="w",
                ).grid(row=1, column=0, sticky="ew", padx=14)
                ctk.CTkLabel(
                    item,
                    text=result.snippet,
                    text_color=COLORS["muted"],
                    font=(self.reading_font_family, FONT_SIZES["result_body"]),
                    wraplength=510,
                    justify="left",
                    anchor="w",
                ).grid(row=2, column=0, sticky="ew", padx=14, pady=(7, 8 if self.embers_mode else 14))
                if self.embers_mode:
                    terms = ", ".join(detect_heat_terms(result.content)) or "-"
                    ctk.CTkLabel(
                        item,
                        text=UI_TEXT["embers_card_meta"].format(
                            updated_at=result.source_updated_at or result.modified_at,
                            source_type=result.source_type,
                            terms=terms,
                        ),
                        text_color=COLORS["quiet"],
                        font=(self.status_font_family, FONT_SIZES["result_meta"]),
                        anchor="w",
                    ).grid(row=3, column=0, sticky="ew", padx=14, pady=(0, 6))
                    ctk.CTkButton(
                        item,
                        text=UI_TEXT["button_read"],
                        width=82,
                        height=28,
                        fg_color=COLORS["panel_soft"],
                        hover_color=COLORS["accent_soft"],
                        text_color=COLORS["section"],
                        font=(self.status_font_family, FONT_SIZES["small"]),
                        command=lambda selected=result: self._select_result(selected),
                    ).grid(row=4, column=0, sticky="w", padx=14, pady=(0, 12))

        def _result_title(self, result: SearchResult) -> str:
            if result.source_type == "chatgpt_export":
                return f"[chatgpt_export] {result.conversation_title or result.title}"
            if result.source_type in {"slack", "slack_inbox"}:
                return f"[slack] {result.title}"
            if result.source_type == "aru":
                return f"[aru] {result.title}"
            if result.source_type == "slack_task":
                return f"[slack_task] {result.title}"
            if result.source_type in {"codex_result", "codex_report_auto"}:
                suffix = f" / {result.commit_hash[:7]}" if result.commit_hash else ""
                return f"[{result.source_type}] {result.title}{suffix}"
            return f"[file] {result.title}"

        def _result_meta(self, result: SearchResult) -> str:
            if result.source_type == "chatgpt_export":
                label = result.source_label or f"ChatGPT / {result.conversation_title or result.title} / {result.role}"
                base = UI_TEXT["result_meta_chatgpt"].format(source_label=label, score=result.score)
                return self._append_semantic_meta(base, result)
            if result.source_type in {"codex_result", "codex_report_auto"}:
                commit_hash = result.commit_hash[:12] if result.commit_hash else "-"
                base = UI_TEXT["result_meta_codex"].format(commit_hash=commit_hash, score=result.score)
                return self._append_semantic_meta(base, result)
            base = UI_TEXT["result_meta"].format(source_type=result.source_type, score=result.score)
            return self._append_semantic_meta(base, result)

        def _append_semantic_meta(self, base: str, result: SearchResult) -> str:
            if result.semantic_score <= 0:
                return base
            return f"{base} | {UI_TEXT['result_meta_semantic'].format(semantic_score=result.semantic_score)}"

        def _select_result(self, result: SearchResult) -> None:
            self.index_status_var.set(UI_TEXT["status_markdown_loading"])
            preview = UI_TEXT["preview_template"].format(
                path=result.path,
                source_type=result.source_type,
                source_label=result.source_label,
                conversation_title=result.conversation_title,
                role=result.role,
                message_index=result.message_index,
                commit_hash=result.commit_hash,
                changed_files=result.changed_files_json,
                test_results=result.test_results,
                build_results=result.build_results,
                push_result=result.push_result,
                git_status=result.git_status,
                modified_at=result.modified_at,
                indexed_at=result.indexed_at,
                score=result.score,
                semantic_score=result.semantic_score,
                content=result.content[:9000],
            )
            source_view = UI_TEXT["source_view_template"].format(
                path=result.path,
                source_type=result.source_type,
                source_created_at=result.source_created_at or result.modified_at,
                content=result.content[:18000],
            )
            if hasattr(self, "source_view_box"):
                set_textbox_text(self.source_view_box, source_view)
            set_textbox_text(self.preview_box, preview)
            self.tags_var.set(self._tags_for_result(result))
            self.selected_result = result
            self.index_status_var.set(UI_TEXT["status_markdown_loaded"])
            self._start_memory_flow(result)

        def _tags_for_result(self, result: SearchResult) -> str:
            tags = [
                result.source_type,
                result.role,
                result.conversation_title,
                result.commit_hash,
                Path(result.path).parent.name,
                Path(result.path).stem,
            ]
            for part in self.current_query.split():
                if part not in tags:
                    tags.append(part)
            return " / ".join(tag for tag in tags if tag)

        def _update_handoff_preview(self) -> None:
            items = "\n".join(f"- {self._result_title(result)}" for result in self.current_results[:6])
            text = UI_TEXT["handoff_preview"].format(
                query=self.current_query,
                count=len(self.current_results),
                items=items or UI_TEXT["empty_results"],
            )
            set_textbox_text(self.handoff_box, text)

        def _render_related_memory(self, results: list[SearchResult]) -> None:
            if not results:
                set_textbox_text(self.related_box, UI_TEXT["related_memory_empty"])
                return
            lines = []
            for result in results[:5]:
                lines.append(f"- {self._result_title(result)}")
                if result.semantic_score > 0:
                    lines.append(f"  {UI_TEXT['result_meta_semantic'].format(semantic_score=result.semantic_score)}")
            set_textbox_text(self.related_box, "\n".join(lines))

        def _toggle_flow_sort(self) -> None:
            self.flow_sort_ascending = not self.flow_sort_ascending
            label = UI_TEXT["button_timeline_ascending"] if self.flow_sort_ascending else UI_TEXT["button_timeline_descending"]
            self.flow_sort_button.configure(text=label)
            if self.selected_result:
                self._start_memory_flow(self.selected_result, force=True)

        def _start_memory_flow(self, result: SearchResult, force: bool = False) -> None:
            semantic_enabled = bool(self.semantic_search_var.get()) and self.semantic_available
            cache_key = (result.id, semantic_enabled, self.flow_sort_ascending)
            cached = self.flow_cache.get(cache_key)
            if cached and not force:
                self._render_memory_flow(cached.items)
                self.flow_status_var.set(UI_TEXT["status_memory_flow_ready"])
                return
            self.flow_request_id += 1
            request_id = self.flow_request_id
            self.flow_status_var.set(UI_TEXT["status_memory_flow_loading"])
            self.flow_thread = threading.Thread(
                target=self._memory_flow_worker,
                args=(request_id, result, semantic_enabled, self.flow_sort_ascending),
                daemon=True,
            )
            self.flow_thread.start()

        def _memory_flow_worker(
            self,
            request_id: int,
            result: SearchResult,
            semantic_enabled: bool,
            ascending: bool,
        ) -> None:
            try:
                response = self.search_engine.memory_flow(
                    result,
                    semantic_enabled=semantic_enabled,
                    ascending=ascending,
                    limit=10,
                )
                self.events.put(("memory_flow_done", (request_id, semantic_enabled, ascending, response)))
            except Exception as exc:
                self.events.put(("memory_flow_error", (request_id, str(exc))))

        def _render_memory_flow(self, items: list[MemoryFlowItem]) -> None:
            for child in self.flow_frame.winfo_children():
                child.destroy()
            if not items:
                self._render_empty_memory_flow()
                return
            for item in items:
                result = item.result
                card = ctk.CTkFrame(self.flow_frame, fg_color=COLORS["panel_alt"], corner_radius=6)
                card.pack(fill="x", padx=5, pady=6)
                card.grid_columnconfigure(0, weight=1)
                title = ctk.CTkButton(
                    card,
                    text=result.conversation_title or result.title,
                    anchor="w",
                    fg_color="transparent",
                    hover_color=COLORS["accent_soft"],
                    text_color=COLORS["text"],
                    font=(self.reading_font_family, FONT_SIZES["small"]),
                    command=lambda selected=result: self._select_result(selected),
                )
                title.grid(row=0, column=0, sticky="ew", padx=8, pady=(7, 1))
                semantic_text = ""
                if result.semantic_score > 0:
                    semantic_text = UI_TEXT["timeline_semantic"].format(semantic_score=result.semantic_score)
                meta = UI_TEXT["timeline_meta"].format(
                    date=timeline_date(result) or "-",
                    source_type=result.source_type,
                    flow_score=item.flow_score,
                    semantic_text=semantic_text,
                )
                ctk.CTkLabel(
                    card,
                    text=meta,
                    text_color=COLORS["quiet"],
                    font=(self.status_font_family, FONT_SIZES["micro"]),
                    anchor="w",
                ).grid(row=1, column=0, sticky="ew", padx=10)
                ctk.CTkLabel(
                    card,
                    text=short_summary(result, 140),
                    text_color=COLORS["muted"],
                    font=(self.reading_font_family, FONT_SIZES["small"]),
                    wraplength=305,
                    justify="left",
                    anchor="w",
                ).grid(row=2, column=0, sticky="ew", padx=10, pady=(4, 9))

        def _export_handoff(self, kind: str) -> None:
            if not self.current_results:
                messagebox.showwarning(UI_TEXT["dialog_title"], UI_TEXT["dialog_no_results"])
                return
            if kind == "chatgpt":
                path = write_chatgpt_handoff(self.current_query, self.current_results)
            else:
                path = write_codex_handoff(self.current_query, self.current_results)
            self._append_log(UI_TEXT["log_export"].format(path=path))
            self._notify(UI_TEXT["notify_handoff_complete"].format(kind=kind.upper()))
            messagebox.showinfo(UI_TEXT["dialog_title"], f"{UI_TEXT['dialog_export_done']}\n{path}")
            try:
                open_path(path)
            except Exception:
                pass

        def _refresh_system_status(self) -> None:
            threading.Thread(target=self._system_worker, daemon=True).start()

        def _system_worker(self) -> None:
            gpu = check_gpu()
            ollama = check_ollama()
            embedding = check_embedding_status()
            self.events.put(("system_status", (gpu, ollama, embedding)))

        def _refresh_stats(self) -> None:
            try:
                stats = self.database.stats()
                self.sqlite_var.set(UI_TEXT["status_sqlite_ready"])
                self.docs_var.set(UI_TEXT["status_docs"].format(documents=stats["documents"], chunks=stats["chunks"]))
            except Exception:
                self.sqlite_var.set(UI_TEXT["status_sqlite_error"])

        def _poll_events(self) -> None:
            while True:
                try:
                    event, payload = self.events.get_nowait()
                except queue.Empty:
                    break

                if event == "index_progress":
                    self._handle_index_progress(payload)
                elif event == "index_error":
                    self._handle_index_error(str(payload))
                elif event == "search_done":
                    query_text, response, semantic_enabled = payload
                    self._handle_search_done(query_text, response, semantic_enabled)
                elif event == "system_status":
                    gpu, ollama, embedding = payload
                    self._handle_system_status(gpu, ollama, embedding)
                elif event == "memory_flow_done":
                    request_id, semantic_enabled, ascending, response = payload
                    self._handle_memory_flow_done(request_id, semantic_enabled, ascending, response)
                elif event == "memory_flow_error":
                    request_id, error_text = payload
                    self._handle_memory_flow_error(request_id, str(error_text))
                elif event == "chatgpt_import_done":
                    self._handle_chatgpt_import_done(payload)
                elif event == "chatgpt_import_missing":
                    self._handle_chatgpt_import_missing()
                elif event == "chatgpt_import_error":
                    self._handle_chatgpt_import_error(str(payload))
                elif event == "codex_import_done":
                    self._handle_codex_import_done(payload)
                elif event == "codex_import_empty":
                    self._handle_codex_import_empty()
                elif event == "codex_import_error":
                    self._handle_codex_import_error(str(payload))
                elif event == "watch_scan_done":
                    self._handle_watch_scan_done(payload)
                elif event == "watch_scan_error":
                    self._handle_watch_scan_error(str(payload))
                elif event == "remote_queue_done":
                    self._handle_remote_queue_done(payload)
                elif event == "remote_queue_error":
                    self._handle_remote_queue_error(str(payload))
                elif event == "codex_report_done":
                    self._handle_codex_report_done(payload)
                elif event == "codex_report_error":
                    self._handle_codex_report_error(str(payload))
                elif event == "slack_inbox_done":
                    self._handle_slack_inbox_done(payload)
                elif event == "slack_inbox_error":
                    self._handle_slack_inbox_error(str(payload))
                elif event == "aru_inbox_done":
                    self._handle_aru_inbox_done(payload)
                elif event == "aru_inbox_error":
                    self._handle_aru_inbox_error(str(payload))

            self.after(100, self._poll_events)

        def _handle_remote_queue_done(self, result: RemoteQueueBatchResult) -> None:
            if result.detected <= 0:
                self._update_remote_queue_status()
                return
            if result.failed:
                self.remote_queue_status_var.set(UI_TEXT["status_remote_queue_failed"])
            else:
                self.remote_queue_status_var.set(UI_TEXT["status_remote_queue_processed"])
            for item in result.results:
                query_label = item.query or item.note or Path(item.source_file).name
                self._append_log(
                    UI_TEXT["log_remote_queue_detected"].format(task_type=item.task_type, query=query_label)
                )
                if item.status == "failed":
                    self._append_log(
                        UI_TEXT["log_remote_queue_failed"].format(path=item.source_file, error=item.error)
                    )
                elif item.task_type in {"search", "handoff_chatgpt", "handoff_codex"} and item.query:
                    self.search_entry.delete(0, "end")
                    self.search_entry.insert(0, item.query)
                    self.current_query = item.query
                    self.config_data.last_query = item.query
                    self.config_store.save(self.config_data)
                    if self.config_data.auto_run_remote_search:
                        self._start_search()
            self._append_log(
                UI_TEXT["log_remote_queue_processed"].format(processed=result.processed, failed=result.failed)
            )
            self._append_log(UI_TEXT["phrase_remote_queue_received"])
            self._refresh_stats()
            self._update_remote_queue_status()
            self._notify(UI_TEXT["notify_remote_queue_detected"])
            if result.processed:
                self._notify(UI_TEXT["notify_remote_queue_processed"])
            if result.failed:
                self._notify(UI_TEXT["notify_remote_queue_failed"])

        def _handle_remote_queue_error(self, error_text: str) -> None:
            self.remote_queue_status_var.set(UI_TEXT["status_remote_queue_failed"])
            self._append_log(UI_TEXT["log_remote_queue_failed"].format(path="-", error=error_text))
            self._notify(UI_TEXT["notify_remote_queue_failed"])
            self._update_remote_queue_status()

        def _handle_codex_report_done(self, result: CodexReportAutoResult) -> None:
            if result.detected <= 0:
                self._update_codex_report_status()
                return
            if result.failed:
                self.codex_report_status_var.set(UI_TEXT["status_codex_report_failed"])
            else:
                self.codex_report_status_var.set(UI_TEXT["status_codex_report_imported"])
            for item in result.items:
                self._append_log(UI_TEXT["log_codex_report_detected"].format(path=item.source_file))
                if item.status == "failed":
                    self._append_log(UI_TEXT["log_codex_report_failed"].format(path=item.source_file, error=item.error))
                else:
                    commit_hash = item.commit_hash[:12] if item.commit_hash else "-"
                    self.last_report_var.set(UI_TEXT["label_last_report"].format(path=Path(item.destination_file).name))
                    self.last_commit_var.set(UI_TEXT["label_last_commit"].format(commit_hash=commit_hash))
                    self._append_log(
                        UI_TEXT["log_codex_report_imported"].format(
                            commit_hash=item.commit_hash,
                            changed_files=item.changed_files_count,
                            skipped=item.skipped_duplicate,
                        )
                    )
            self._append_log(UI_TEXT["phrase_codex_report_saved"])
            if result.log_path:
                self._append_log(UI_TEXT["log_codex_import_file"].format(path=result.log_path))
            self._refresh_stats()
            self._update_codex_report_status()
            if result.failed:
                self.codex_report_status_var.set(UI_TEXT["status_codex_report_failed"])
            else:
                self.codex_report_status_var.set(UI_TEXT["status_codex_report_imported"])
            if result.imported:
                self._set_last_import_summary("last_import_codex_report")
                self._notify(UI_TEXT["notify_codex_report_imported"])
                if self.semantic_available:
                    self._notify(UI_TEXT["notify_semantic_updated"])
            if result.failed:
                self._notify(UI_TEXT["notify_codex_report_failed"])

        def _handle_codex_report_error(self, error_text: str) -> None:
            self.codex_report_status_var.set(UI_TEXT["status_codex_report_failed"])
            self._append_log(UI_TEXT["log_codex_report_failed"].format(path="-", error=error_text))
            self._notify(UI_TEXT["notify_codex_report_failed"])
            self._update_codex_report_status()
            self.codex_report_status_var.set(UI_TEXT["status_codex_report_failed"])

        def _handle_slack_inbox_done(self, result: SlackInboxResult) -> None:
            status_text = self._slack_status_text(result.status)
            self.slack_status_var.set(status_text)
            self.slack_channel_var.set(UI_TEXT["label_slack_channel_status"].format(channel=result.channel_label or result.channel_id or "-"))
            if result.latest_ts and result.latest_ts != self.config_data.slack_last_ts:
                self.config_data.slack_last_ts = result.latest_ts
                self.config_store.save(self.config_data)
            if result.imported:
                self.slack_last_import_var.set(UI_TEXT["label_slack_last_import"].format(time=now_iso()))
                self._append_log(
                    UI_TEXT["log_slack_import"].format(
                        imported=result.imported,
                        skipped=result.skipped,
                        failed=result.failed,
                    )
                )
                self._append_log(UI_TEXT["phrase_slack_memory_saved"])
                if result.log_path:
                    self._append_log(UI_TEXT["log_slack_file"].format(path=result.log_path))
                self._refresh_stats()
                self._set_last_import_summary("last_import_slack")
                self._notify(UI_TEXT["notify_slack_import_complete"])
                if self.semantic_available:
                    self._notify(UI_TEXT["notify_semantic_updated"])
            if result.task_results:
                self._handle_slack_task_results(result.task_results)
            elif result.status not in {"connected"} and result.status != self.last_slack_status:
                self._append_log(UI_TEXT["log_slack_status"].format(status=result.status, message=result.message))
                if result.status in {"auth_failed", "channel_not_found", "timeout", "error"}:
                    self._notify(UI_TEXT["notify_slack_import_failed"])
            self.last_slack_status = result.status

        def _handle_slack_task_results(self, task_results: list[SlackTaskResult]) -> None:
            processed = 0
            failed = 0
            for task_result in task_results:
                label = self._slack_task_label(task_result)
                self.last_slack_task_var.set(UI_TEXT["label_last_task"].format(task=label))
                self._append_log(
                    UI_TEXT["log_slack_task_detected"].format(
                        task_type=task_result.task_type,
                        query=task_result.query or task_result.note or task_result.import_path,
                    )
                )
                if task_result.status == "failed":
                    failed += 1
                    self._append_log(
                        UI_TEXT["log_slack_task_failed"].format(
                            task_type=task_result.task_type,
                            error=task_result.error,
                        )
                    )
                    continue
                processed += 1
                self._append_log(
                    UI_TEXT["log_slack_task_processed"].format(
                        status=task_result.status,
                        changed=task_result.changed or task_result.slack_task_changed,
                        duplicate=task_result.skipped_duplicate,
                    )
                )
                if task_result.task_type in {"search", "handoff_chatgpt", "handoff_codex"} and task_result.query:
                    self.search_entry.delete(0, "end")
                    self.search_entry.insert(0, task_result.query)
                    self.current_query = task_result.query
                    self.config_data.last_query = task_result.query
                    self.config_store.save(self.config_data)
                    if self.config_data.auto_run_remote_search:
                        self._start_search()
            if processed:
                self.slack_status_var.set(UI_TEXT["status_slack_task_processed"])
                self._append_log(UI_TEXT["phrase_slack_task_received"])
                self._notify(UI_TEXT["notify_slack_task_received"])
            if failed:
                self.slack_status_var.set(UI_TEXT["status_slack_task_failed"])
                self._notify(UI_TEXT["notify_remote_queue_failed"])
            self._refresh_stats()
            self._update_remote_queue_status()

        def _slack_task_label(self, task_result: SlackTaskResult) -> str:
            value = task_result.query or task_result.note or task_result.import_path or "-"
            value = " ".join(value.split())
            if len(value) > 48:
                value = f"{value[:45]}..."
            return f"{task_result.task_type} / {value}"

        def _handle_slack_inbox_error(self, error_text: str) -> None:
            self.slack_status_var.set(UI_TEXT["status_slack_error"])
            self._append_log(UI_TEXT["log_slack_status"].format(status="error", message=error_text))
            self._notify(UI_TEXT["notify_slack_import_failed"])
            self.last_slack_status = "error"

        def _handle_aru_inbox_done(self, result: SlackInboxResult) -> None:
            status_text = self._aru_status_text(result.status)
            self.aru_status_var.set(status_text)
            self.aru_channel_var.set(UI_TEXT["label_aru_channel_status"].format(channel=result.channel_label or result.channel_id or "-"))
            if result.latest_ts and result.latest_ts != self.config_data.aru_last_ts:
                self.config_data.aru_last_ts = result.latest_ts
                self.config_store.save(self.config_data)
            if result.imported:
                self.aru_last_import_var.set(UI_TEXT["label_aru_last_import"].format(time=now_iso()))
                self._append_log(
                    UI_TEXT["log_aru_import"].format(
                        imported=result.imported,
                        skipped=result.skipped,
                        failed=result.failed,
                    )
                )
                self._append_log(UI_TEXT["phrase_aru_memory_saved"])
                if result.log_path:
                    self._append_log(UI_TEXT["log_aru_file"].format(path=result.log_path))
                self._refresh_stats()
                self._set_last_import_summary("last_import_aru")
                self._notify(UI_TEXT["notify_aru_import_complete"])
                if self.semantic_available:
                    self._notify(UI_TEXT["notify_semantic_updated"])
            elif result.status not in {"connected"} and result.status != self.last_aru_status:
                self._append_log(UI_TEXT["log_aru_status"].format(status=result.status, message=result.message))
                if result.status in {"auth_failed", "channel_not_found", "timeout", "error"}:
                    self._notify(UI_TEXT["notify_aru_import_failed"])
            self.last_aru_status = result.status

        def _handle_aru_inbox_error(self, error_text: str) -> None:
            self.aru_status_var.set(UI_TEXT["status_aru_error"])
            self._append_log(UI_TEXT["log_aru_status"].format(status="error", message=error_text))
            self._notify(UI_TEXT["notify_aru_import_failed"])
            self.last_aru_status = "error"

        def _slack_status_text(self, status: str) -> str:
            return {
                "imported": UI_TEXT["status_slack_import_complete"],
                "connected": UI_TEXT["status_slack_connected"],
                "auth_failed": UI_TEXT["status_slack_auth_failed"],
                "channel_not_found": UI_TEXT["status_slack_channel_not_found"],
                "timeout": UI_TEXT["status_slack_timeout"],
                "config_missing": UI_TEXT["status_slack_config_missing"],
                "error": UI_TEXT["status_slack_error"],
            }.get(status, UI_TEXT["status_slack_error"])

        def _aru_status_text(self, status: str) -> str:
            return {
                "imported": UI_TEXT["status_aru_import_complete"],
                "connected": UI_TEXT["status_aru_connected"],
                "auth_failed": UI_TEXT["status_aru_auth_failed"],
                "channel_not_found": UI_TEXT["status_aru_channel_not_found"],
                "timeout": UI_TEXT["status_aru_timeout"],
                "config_missing": UI_TEXT["status_aru_config_missing"],
                "error": UI_TEXT["status_aru_error"],
            }.get(status, UI_TEXT["status_aru_error"])

        def _handle_watch_scan_done(self, result: WatchScanResult) -> None:
            self._update_watch_status()
            if not self.auto_index_var.get() or not result.changed_files:
                return
            first_path = Path(result.changed_files[0])
            self.index_status_var.set(UI_TEXT["status_new_memory_detected"])
            self.watch_status_var.set(f"{UI_TEXT['status_auto_index_on']} / {UI_TEXT['status_new_memory_detected']}")
            self.last_memory_var.set(UI_TEXT["status_last_memory"].format(path=first_path.name))
            for path_text in result.changed_files[:8]:
                self._append_log(UI_TEXT["log_new_memory_detected"].format(path=path_text))
            self._notify(UI_TEXT["notify_memory_detected"])
            lines = [
                UI_TEXT["log_watch_initialized"],
                UI_TEXT["log_watching_folder"].format(folder=result.folder),
                f"checked={result.checked}",
                f"changed={len(result.changed_files)}",
                *[UI_TEXT["log_new_memory_detected"].format(path=path) for path in result.changed_files],
            ]
            self._write_watch_event_log(lines)
            self._start_auto_index(Path(result.folder))

        def _handle_watch_scan_error(self, error_text: str) -> None:
            self._append_log(UI_TEXT["log_index_error"].format(error=error_text))

        def _handle_index_progress(self, progress: IndexProgress) -> None:
            if progress.total:
                self.progress.set(progress.current / progress.total)
            else:
                self.progress.set(0)
            if progress.done:
                was_auto_index = self.auto_index_active
                self.auto_index_active = False
                self.index_button.configure(state="normal")
                self.cancel_button.configure(state="disabled")
                self.config_data.last_indexed_at = now_iso()
                self.config_store.save(self.config_data)
                self._update_watch_status()
                self._refresh_stats()
                if progress.cancelled:
                    self.index_status_var.set(
                        UI_TEXT["index_done"].format(indexed=progress.indexed, skipped=progress.skipped, errors=progress.errors)
                    )
                    self._append_log(
                        UI_TEXT["log_index_cancelled"].format(indexed=progress.indexed, skipped=progress.skipped, errors=progress.errors)
                    )
                else:
                    self.index_status_var.set(
                        UI_TEXT["index_done"].format(indexed=progress.indexed, skipped=progress.skipped, errors=progress.errors)
                    )
                    self._append_log(
                        UI_TEXT["log_index_done"].format(indexed=progress.indexed, skipped=progress.skipped, errors=progress.errors)
                    )
                    if was_auto_index:
                        self._append_log(UI_TEXT["log_auto_index_complete"])
                        self._append_log(UI_TEXT["phrase_memory_updated"])
                        self._notify(UI_TEXT["notify_auto_index_complete"])
                        self._write_watch_event_log(
                            [
                                UI_TEXT["log_auto_index_complete"],
                                UI_TEXT["log_index_done"].format(
                                    indexed=progress.indexed,
                                    skipped=progress.skipped,
                                    errors=progress.errors,
                                ),
                            ]
                        )
                    if progress.indexed > 0 and self.semantic_available:
                        self._notify(UI_TEXT["notify_semantic_updated"])
                if self.pending_auto_index_folder and not self.cancel_event.is_set():
                    next_folder = self.pending_auto_index_folder
                    self.pending_auto_index_folder = None
                    self.after(200, lambda folder=next_folder: self._start_auto_index(folder))
            else:
                if progress.message == "embedding_chunks":
                    self.index_status_var.set(UI_TEXT["index_embedding"])
                else:
                    self.index_status_var.set(
                        UI_TEXT["index_running"].format(current=progress.current, total=progress.total)
                    )

        def _handle_index_error(self, error_text: str) -> None:
            self.auto_index_active = False
            self.index_button.configure(state="normal")
            self.cancel_button.configure(state="disabled")
            self._set_import_buttons_state("normal")
            self._set_search_running(False)
            self.index_status_var.set(UI_TEXT["index_idle"])
            self._update_watch_status()
            self._append_log(UI_TEXT["log_index_error"].format(error=error_text))

        def _handle_chatgpt_import_done(self, result: ChatGPTImportResult) -> None:
            self._set_import_buttons_state("normal")
            self.index_status_var.set(UI_TEXT["status_chatgpt_import_complete"])
            self.progress.set(1)
            if hasattr(self, "chatgpt_import_status_var"):
                if result.messages_indexed:
                    self.chatgpt_import_status_var.set(
                        UI_TEXT["chatgpt_import_result"].format(
                            conversations=result.conversations_imported,
                            messages=result.messages_indexed,
                        )
                    )
                else:
                    self.chatgpt_import_status_var.set(UI_TEXT["chatgpt_import_result_no_new"])
            self._refresh_stats()
            self._append_log(UI_TEXT["log_conversations_json_found"].format(path=result.conversations_json_path))
            self._append_log(
                UI_TEXT["log_chatgpt_import_complete"].format(
                    conversations=result.conversations_imported,
                    messages=result.messages_indexed,
                    skipped=result.skipped_duplicates,
                    errors=result.errors,
                )
            )
            self._append_log(UI_TEXT["log_chatgpt_memory_imported"])
            self._append_log(UI_TEXT["log_chatgpt_import_file"].format(path=result.log_path))
            self._set_last_import_summary("last_import_chatgpt_export")
            self._notify(UI_TEXT["notify_chatgpt_import_complete"])

        def _handle_chatgpt_import_missing(self) -> None:
            self._set_import_buttons_state("normal")
            self.index_status_var.set(UI_TEXT["index_idle"])
            self.progress.set(0)
            if hasattr(self, "chatgpt_import_status_var"):
                self.chatgpt_import_status_var.set(UI_TEXT["error_conversations_json_not_found"])
            self._append_log(UI_TEXT["error_conversations_json_not_found"])

        def _handle_chatgpt_import_error(self, error_text: str) -> None:
            self._set_import_buttons_state("normal")
            self.index_status_var.set(UI_TEXT["index_idle"])
            self.progress.set(0)
            if hasattr(self, "chatgpt_import_status_var"):
                self.chatgpt_import_status_var.set(UI_TEXT["chatgpt_import_failed_short"])
            self._append_log(f"{UI_TEXT['error_chatgpt_import_failed']} {error_text}")

        def _handle_codex_import_done(self, result: CodexImportResult) -> None:
            self._set_import_buttons_state("normal")
            self.index_status_var.set(UI_TEXT["status_codex_import_complete"])
            self.progress.set(1)
            self._refresh_stats()
            self._append_log(UI_TEXT["log_codex_commit_found"].format(commit_hash=result.commit_hash))
            self._append_log(
                UI_TEXT["log_codex_import_complete"].format(
                    changed_files=result.changed_files_count,
                    skipped=result.skipped_duplicate,
                )
            )
            self._append_log(UI_TEXT["log_codex_memory_imported"])
            self._append_log(UI_TEXT["log_codex_import_file"].format(path=result.log_path))
            self._set_last_import_summary("last_import_codex_result")
            self._notify(UI_TEXT["notify_codex_import_complete"])
            messagebox.showinfo(UI_TEXT["dialog_title"], UI_TEXT["status_codex_import_complete"])

        def _handle_codex_import_empty(self) -> None:
            self._set_import_buttons_state("normal")
            self.index_status_var.set(UI_TEXT["index_idle"])
            self.progress.set(0)
            self._append_log(UI_TEXT["error_codex_result_empty"])
            messagebox.showwarning(UI_TEXT["dialog_title"], UI_TEXT["error_codex_result_empty"])

        def _handle_codex_import_error(self, error_text: str) -> None:
            self._set_import_buttons_state("normal")
            self.index_status_var.set(UI_TEXT["index_idle"])
            self.progress.set(0)
            self._append_log(f"{UI_TEXT['error_codex_import_failed']} {error_text}")
            messagebox.showwarning(UI_TEXT["dialog_title"], f"{UI_TEXT['error_codex_import_failed']}\n{error_text}")

        def _handle_memory_flow_done(
            self,
            request_id: int,
            semantic_enabled: bool,
            ascending: bool,
            response: MemoryFlowResponse,
        ) -> None:
            if request_id != self.flow_request_id:
                return
            self.current_flow_items = response.items
            self.flow_cache[(response.anchor_id, semantic_enabled, ascending)] = response
            self.flow_status_var.set(UI_TEXT["status_memory_flow_ready"])
            self._render_memory_flow(response.items)
            self._append_log(UI_TEXT["log_memory_flow_generated"])
            self._append_log(UI_TEXT["log_memory_flow_updated"].format(count=len(response.items)))
            self._append_log(UI_TEXT["phrase_memory_reconnected"])
            self._append_log(UI_TEXT["log_memory_flow_file"].format(path=response.log_path))

        def _handle_memory_flow_error(self, request_id: int, error_text: str) -> None:
            if request_id != self.flow_request_id:
                return
            self.flow_status_var.set(UI_TEXT["index_idle"])
            self._render_empty_memory_flow()
            self._append_log(UI_TEXT["log_index_error"].format(error=error_text))

        def _handle_search_done(self, query_text: str, response: SearchResponse, semantic_enabled: bool) -> None:
            results = response.results
            if self.embers_mode:
                results = sorted(results, key=lambda item: (item.source_type != "aru", -item.score, item.title.lower()))
            self.current_results = results
            self.current_related_results = response.related
            self._set_search_running(False)
            if results:
                self.index_status_var.set(UI_TEXT["status_search_complete"].format(count=len(results)))
            else:
                self.index_status_var.set(UI_TEXT["status_no_results"])
            self._render_results(results)
            self._render_related_memory(response.related)
            self._update_handoff_preview()
            if results:
                self._select_result(results[0])
            else:
                self.selected_result = None
                self.current_flow_items = []
                self._render_empty_memory_flow()
            self._refresh_stats()
            if results:
                self._append_log(UI_TEXT["log_search_complete"].format(count=len(results)))
            else:
                self._append_log(UI_TEXT["log_no_results"].format(query=query_text))
                self._append_log(UI_TEXT["phrase_no_results"])
            if semantic_enabled:
                self._append_log(UI_TEXT["log_semantic_search_initialized"])
                if response.semantic_available:
                    self.semantic_status_var.set(UI_TEXT["status_semantic_search_ready"])
                    self._append_log(UI_TEXT["log_related_memory_found"].format(count=len(response.related)))
                else:
                    self.semantic_available = False
                    self.semantic_search_var.set(False)
                    self.embedding_var.set(UI_TEXT["status_embedding_unavailable"])
                    self.semantic_status_var.set(UI_TEXT["status_semantic_search_disabled"])
                    self._append_log(UI_TEXT["log_semantic_disabled"].format(message=response.semantic_message))
            self._append_log(UI_TEXT["log_search"].format(query=query_text, count=len(results)))

        def _handle_system_status(self, gpu, ollama, embedding) -> None:
            if gpu.cuda_available:
                self.cuda_var.set(UI_TEXT["status_cuda_online"])
            else:
                self.cuda_var.set(UI_TEXT["status_cuda_unavailable"])

            if gpu.gpu_detected:
                self.gpu_var.set(UI_TEXT["status_gpu_detected"].format(name=gpu.gpu_name))
            else:
                self.gpu_var.set(UI_TEXT["status_gpu_missing"])

            if ollama.available:
                model_count = len(ollama.models)
                self.ollama_var.set(f"{UI_TEXT['status_ollama_ready']} ({model_count})")
            else:
                self.ollama_var.set(UI_TEXT["status_ollama_not_running"])

            if embedding.available:
                self.semantic_available = True
                self.embedding_var.set(UI_TEXT["status_embedding_ready"])
                self.semantic_status_var.set(UI_TEXT["status_semantic_search_ready"])
                self.semantic_checkbox.configure(state="normal")
            else:
                self.semantic_available = False
                self.semantic_search_var.set(False)
                self.embedding_var.set(UI_TEXT["status_embedding_unavailable"])
                self.semantic_status_var.set(UI_TEXT["status_semantic_search_disabled"])
                self.semantic_checkbox.configure(state="disabled")

        def _write_qpsc_awake_status(self) -> None:
            try:
                write_brainz_awake_status(started_at=self.qpsc_started_at)
                if hasattr(self, "awake_status_var"):
                    self.awake_status_var.set(UI_TEXT["status_brainz_awake_detail"])
            except Exception:
                if hasattr(self, "awake_status_var"):
                    self.awake_status_var.set(UI_TEXT["status_brainz_standby"])
                pass

        def _heartbeat_qpsc_status(self) -> None:
            self._write_qpsc_awake_status()
            self.after(30000, self._heartbeat_qpsc_status)

        def _launch_check_finish(self) -> None:
            print(UI_TEXT["launch_check_ok"])
            self.destroy()

    app = BrainzApp()
    app.mainloop()
    return 0


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--smoke-test", action="store_true")
    parser.add_argument("--launch-check", action="store_true")
    args = parser.parse_args()

    if args.smoke_test:
        return run_smoke_test()
    return run_gui(launch_check=args.launch_check)


if __name__ == "__main__":
    raise SystemExit(main())
