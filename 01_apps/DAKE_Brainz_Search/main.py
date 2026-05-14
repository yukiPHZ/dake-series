# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import json
import queue
import tempfile
import threading
import time
import zipfile
from pathlib import Path


APP_NAME = "補助脳BRAINZ"
WINDOW_TITLE = "補助脳BRAINZ - Local Memory Search"
COPYRIGHT = "© 2026 PEAKHEADZ / DAKE series"

UI_TEXT = {
    "app_title": APP_NAME,
    "window_title": WINDOW_TITLE,
    "copyright": COPYRIGHT,
    "subtitle": "Local Memory Search",
    "search_placeholder": "うろ覚えで検索",
    "button_search": "Search",
    "checkbox_semantic_search": "Semantic Search",
    "button_index": "Index",
    "button_cancel": "Cancel",
    "button_choose": "Choose",
    "button_import_chatgpt_export": "ChatGPT exportを取り込む",
    "button_import_codex_result": "Codex結果を取り込む",
    "button_chatgpt": "ChatGPTまとめ",
    "button_codex": "Codex素材",
    "memory_title": "Memory Folder",
    "index_title": "Index Status",
    "system_title": "System Status",
    "results_title": "Search Results",
    "preview_title": "Preview",
    "tags_title": "Related Tags",
    "section_related_memory": "Related Memory",
    "handoff_title": "Handoff Summary",
    "log_title": "BRAINZ Log",
    "empty_memory": "未選択",
    "empty_results": "検索結果はまだありません",
    "empty_preview": "結果を選ぶと詳細が表示されます",
    "empty_tags": "タグ候補なし",
    "related_memory_empty": "関連記憶なし",
    "empty_handoff": "検索結果から引き継ぎ素材を生成できます",
    "choose_memory_title": "記憶フォルダを選択",
    "dialog_title": "補助脳BRAINZ",
    "dialog_memory_missing": "記憶フォルダを選択してください",
    "dialog_no_results": "検索結果がありません",
    "dialog_export_done": "ファイルを生成しました",
    "dialog_select_chatgpt_export": "ChatGPT export zipを選択（キャンセルでフォルダ選択）",
    "dialog_select_chatgpt_export_folder": "展開済みChatGPT exportフォルダを選択",
    "dialog_import_codex_result": "Codex結果を取り込む",
    "dialog_select_codex_result_file": "Codex結果ファイルを選択",
    "label_codex_result_input": "Codexの完了報告・修正結果・commit/push結果を貼り付け",
    "button_import_pasted_codex": "貼り付けを取り込む",
    "button_import_codex_file": "txt / md を選択",
    "button_close": "閉じる",
    "filetype_zip": "zipファイル",
    "filetype_text_markdown": "txt / md",
    "filetype_all": "すべてのファイル",
    "status_importing_chatgpt": "CHATGPT EXPORT IMPORTING...",
    "status_chatgpt_import_complete": "CHATGPT EXPORT IMPORT COMPLETE",
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
    "log_semantic_search_initialized": "Semantic search initialized.",
    "log_embedding_generated": "Embedding generated.",
    "log_related_memory_found": "Related memory found: {count}",
    "log_semantic_disabled": "FTS only mode: {message}",
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
    "log_export": "EXPORT: {path}",
    "index_idle": "IDLE",
    "index_running": "RUNNING {current}/{total}",
    "index_embedding": "EMBEDDING CHUNKS...",
    "index_done": "DONE {indexed} indexed / {skipped} skipped / {errors} errors",
    "searching": "SEARCHING...",
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
    "status_docs": "DOCS {documents} / CHUNKS {chunks}",
    "preview_template": "PATH: {path}\nSOURCE: {source_type}\nLABEL: {source_label}\nCONVERSATION: {conversation_title}\nROLE: {role}\nMESSAGE INDEX: {message_index}\nCOMMIT: {commit_hash}\nCHANGED FILES: {changed_files}\nTEST RESULTS: {test_results}\nBUILD RESULTS: {build_results}\nPUSH: {push_result}\nGIT STATUS: {git_status}\nMODIFIED: {modified_at}\nINDEXED: {indexed_at}\nSCORE: {score:.2f}\nSEMANTIC: {semantic_score:.2f}\n\n{content}",
    "result_meta": "{source_type} | score {score:.1f}",
    "result_meta_chatgpt": "{source_label} | score {score:.1f}",
    "result_meta_codex": "commit {commit_hash} | score {score:.1f}",
    "result_meta_semantic": "semantic {semantic_score:.2f}",
    "handoff_preview": "query: {query}\nresults: {count}\n\n{items}",
    "launch_check_ok": "LAUNCH CHECK OK",
}


def run_smoke_test() -> int:
    from core.app_config import ensure_app_dirs
    from core.chatgpt_importer import import_chatgpt_export
    from core.codex_importer import import_codex_file, import_codex_text
    from core.db import BrainzDatabase
    from core.gpu_checker import check_gpu
    from core.handoff_writer import write_chatgpt_handoff, write_codex_handoff
    from core.indexer import Indexer
    from core.ollama_client import check_ollama
    from core.ollama_embeddings import DEFAULT_EMBED_MODEL, check_embedding_status
    from core.search_engine import SearchEngine

    ensure_app_dirs()
    database = BrainzDatabase()
    database.ensure_schema()

    with tempfile.TemporaryDirectory(prefix="brainz_memory_") as tmp:
        root = Path(tmp)
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

    engine = SearchEngine(database)
    file_results = engine.search(UI_TEXT["smoke_query_memory"], limit=10)
    if not file_results:
        raise RuntimeError("file search returned no results")

    chatgpt_results = engine.search(UI_TEXT["smoke_chatgpt_query"], limit=10)
    chatgpt_match = next((result for result in chatgpt_results if result.source_type == "chatgpt_export"), None)
    if chatgpt_match is None:
        raise RuntimeError("chatgpt_export search returned no result")
    if not chatgpt_match.conversation_title or not chatgpt_match.role:
        raise RuntimeError("chatgpt metadata was not stored")

    chatgpt_path = write_chatgpt_handoff(UI_TEXT["smoke_chatgpt_query"], chatgpt_results)
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

    ollama_status = check_ollama()
    gpu_status = check_gpu()
    stats = database.stats()
    print("SMOKE OK")
    print(f"documents={stats['documents']} chunks={stats['chunks']}")
    print(f"file_results={len(file_results)}")
    print(f"chatgpt_results={len(chatgpt_results)}")
    print(f"codex_results={len(codex_results)}")
    print(f"chatgpt_handoff={chatgpt_path}")
    print(f"codex_handoff={codex_path}")
    print(f"codex_chatgpt_handoff={codex_chatgpt_path}")
    print(f"codex_codex_handoff={codex_codex_path}")
    print(f"ollama_available={ollama_status.available}")
    print(f"embedding_available={embedding_status.available}")
    print(f"embeddings_ready={embedding_stats['ready']}")
    print(f"semantic_available={semantic_response.semantic_available}")
    print(f"related_memory={len(semantic_response.related)}")
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


def run_gui(launch_check: bool = False) -> int:
    import customtkinter as ctk
    from PIL import Image
    from tkinter import filedialog, messagebox

    from core.app_config import (
        ConfigStore,
        common_icon_path,
        ensure_app_dirs,
        open_path,
        peakheadz_logo_path,
    )
    from core.chatgpt_importer import ConversationsJsonNotFound, ChatGPTImportResult, import_chatgpt_export
    from core.codex_importer import CodexImportResult, import_codex_file, import_codex_text
    from core.db import BrainzDatabase, SearchResult
    from core.gpu_checker import check_gpu
    from core.handoff_writer import write_chatgpt_handoff, write_codex_handoff
    from core.indexer import IndexProgress, Indexer
    from core.ollama_client import check_ollama
    from core.ollama_embeddings import check_embedding_status
    from core.search_engine import SearchEngine, SearchResponse
    from ui.components import choose_font_family, set_textbox_text
    from ui.theme import COLORS, FONT_CANDIDATES

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
            self.config_store = ConfigStore()
            self.config_data = self.config_store.load()
            self.database = BrainzDatabase()
            self.database.ensure_schema()
            self.search_engine = SearchEngine(self.database)
            self.indexer = Indexer(self.database)
            self.events: queue.Queue[tuple[str, object]] = queue.Queue()
            self.cancel_event = threading.Event()
            self.index_thread: threading.Thread | None = None
            self.search_thread: threading.Thread | None = None
            self.import_thread: threading.Thread | None = None
            self.current_results: list[SearchResult] = []
            self.current_related_results: list[SearchResult] = []
            self.current_query = self.config_data.last_query
            self.semantic_available = True
            self.semantic_search_var = ctk.BooleanVar(value=True)
            self.logo_image = None

            self._apply_icon()
            self._build_ui()
            self._set_memory_folder(self.config_data.memory_folder, persist=False)
            self._append_log(UI_TEXT["log_ready"])
            self._refresh_stats()
            self._refresh_system_status()
            self.after(100, self._poll_events)
            if launch_check:
                self.after(1200, self._launch_check_finish)

        def _apply_icon(self) -> None:
            try:
                icon = common_icon_path()
                if icon.exists():
                    self.iconbitmap(str(icon))
            except Exception:
                pass

        def _build_ui(self) -> None:
            self.grid_columnconfigure(0, weight=0, minsize=290)
            self.grid_columnconfigure(1, weight=1, minsize=470)
            self.grid_columnconfigure(2, weight=0, minsize=360)
            self.grid_rowconfigure(1, weight=1)
            self.grid_rowconfigure(2, weight=0)

            header = ctk.CTkFrame(self, fg_color=COLORS["bg"], corner_radius=0)
            header.grid(row=0, column=0, columnspan=3, sticky="ew", padx=18, pady=(16, 10))
            header.grid_columnconfigure(1, weight=1)

            logo_path = peakheadz_logo_path()
            if logo_path.exists():
                try:
                    image = Image.open(logo_path)
                    self.logo_image = ctk.CTkImage(light_image=image, dark_image=image, size=(42, 42))
                    ctk.CTkLabel(header, image=self.logo_image, text="").grid(row=0, column=0, rowspan=2, padx=(0, 12))
                except Exception:
                    pass

            ctk.CTkLabel(
                header,
                text=UI_TEXT["app_title"],
                text_color=COLORS["text"],
                font=(self.font_family, 26, "bold"),
            ).grid(row=0, column=1, sticky="w")
            ctk.CTkLabel(
                header,
                text=UI_TEXT["subtitle"],
                text_color=COLORS["muted"],
                font=(self.font_family, 13),
            ).grid(row=1, column=1, sticky="w")

            search_box = ctk.CTkFrame(header, fg_color=COLORS["bg"], corner_radius=0)
            search_box.grid(row=0, column=2, rowspan=2, sticky="e")
            search_box.grid_columnconfigure(0, weight=1)

            self.search_entry = ctk.CTkEntry(
                search_box,
                width=430,
                height=42,
                placeholder_text=UI_TEXT["search_placeholder"],
                fg_color=COLORS["input"],
                border_color=COLORS["border"],
                text_color=COLORS["text"],
                font=(self.font_family, 15),
            )
            self.search_entry.grid(row=0, column=0, padx=(0, 10))
            self.search_entry.insert(0, self.current_query)
            self.search_entry.bind("<Return>", lambda _event: self._start_search())
            ctk.CTkButton(
                search_box,
                text=UI_TEXT["button_search"],
                width=92,
                height=42,
                fg_color=COLORS["accent"],
                hover_color=COLORS["accent_hover"],
                command=self._start_search,
            ).grid(row=0, column=1)
            self.semantic_checkbox = ctk.CTkCheckBox(
                search_box,
                text=UI_TEXT["checkbox_semantic_search"],
                variable=self.semantic_search_var,
                text_color=COLORS["muted"],
                fg_color=COLORS["accent"],
                hover_color=COLORS["accent_hover"],
                border_color=COLORS["border"],
                font=(self.font_family, 12),
            )
            self.semantic_checkbox.grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))

            left = self._panel(self, 0)
            left.grid(row=1, column=0, sticky="nsew", padx=(18, 8), pady=(0, 10))
            left.grid_columnconfigure(0, weight=1)
            self._section_title(left, UI_TEXT["memory_title"], 0)
            self.memory_var = ctk.StringVar(value=UI_TEXT["empty_memory"])
            ctk.CTkLabel(
                left,
                textvariable=self.memory_var,
                text_color=COLORS["muted"],
                font=(self.font_family, 12),
                wraplength=238,
                justify="left",
            ).grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 10))
            ctk.CTkButton(
                left,
                text=UI_TEXT["button_choose"],
                height=34,
                fg_color=COLORS["panel_soft"],
                hover_color=COLORS["accent_soft"],
                command=self._choose_memory_folder,
            ).grid(row=2, column=0, sticky="ew", padx=14, pady=(0, 10))
            self.import_button = ctk.CTkButton(
                left,
                text=UI_TEXT["button_import_chatgpt_export"],
                height=34,
                fg_color=COLORS["panel_soft"],
                hover_color=COLORS["accent_soft"],
                command=self._choose_chatgpt_export,
            )
            self.import_button.grid(row=3, column=0, sticky="ew", padx=14, pady=(0, 16))
            self.codex_import_button = ctk.CTkButton(
                left,
                text=UI_TEXT["button_import_codex_result"],
                height=34,
                fg_color=COLORS["panel_soft"],
                hover_color=COLORS["accent_soft"],
                command=self._open_codex_import_dialog,
            )
            self.codex_import_button.grid(row=4, column=0, sticky="ew", padx=14, pady=(0, 16))

            self._section_title(left, UI_TEXT["index_title"], 5)
            self.index_status_var = ctk.StringVar(value=UI_TEXT["index_idle"])
            ctk.CTkLabel(
                left,
                textvariable=self.index_status_var,
                text_color=COLORS["text"],
                font=(self.font_family, 13, "bold"),
            ).grid(row=6, column=0, sticky="w", padx=14, pady=(0, 8))
            self.progress = ctk.CTkProgressBar(left, height=10, progress_color=COLORS["accent"])
            self.progress.set(0)
            self.progress.grid(row=7, column=0, sticky="ew", padx=14, pady=(0, 12))

            button_row = ctk.CTkFrame(left, fg_color="transparent")
            button_row.grid(row=8, column=0, sticky="ew", padx=14, pady=(0, 18))
            button_row.grid_columnconfigure((0, 1), weight=1)
            self.index_button = ctk.CTkButton(
                button_row,
                text=UI_TEXT["button_index"],
                height=34,
                fg_color=COLORS["accent"],
                hover_color=COLORS["accent_hover"],
                command=self._start_index,
            )
            self.index_button.grid(row=0, column=0, sticky="ew", padx=(0, 6))
            self.cancel_button = ctk.CTkButton(
                button_row,
                text=UI_TEXT["button_cancel"],
                height=34,
                fg_color=COLORS["panel_soft"],
                hover_color=COLORS["accent_soft"],
                state="disabled",
                command=self._cancel_index,
            )
            self.cancel_button.grid(row=0, column=1, sticky="ew", padx=(6, 0))

            self._section_title(left, UI_TEXT["system_title"], 9)
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
                start=10,
            ):
                ctk.CTkLabel(
                    left,
                    textvariable=variable,
                    text_color=COLORS["muted"],
                    font=(self.font_family, 12),
                    anchor="w",
                ).grid(row=index, column=0, sticky="ew", padx=14, pady=2)

            center = self._panel(self, 0)
            center.grid(row=1, column=1, sticky="nsew", padx=8, pady=(0, 10))
            center.grid_columnconfigure(0, weight=1)
            center.grid_rowconfigure(1, weight=1)
            self._section_title(center, UI_TEXT["results_title"], 0)
            self.results_frame = ctk.CTkScrollableFrame(center, fg_color=COLORS["panel"], corner_radius=0)
            self.results_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
            self._render_empty_results()

            right = self._panel(self, 0)
            right.grid(row=1, column=2, sticky="nsew", padx=(8, 18), pady=(0, 10))
            right.grid_columnconfigure(0, weight=1)
            right.grid_rowconfigure(1, weight=2)
            right.grid_rowconfigure(7, weight=1)
            self._section_title(right, UI_TEXT["preview_title"], 0)
            self.preview_box = ctk.CTkTextbox(
                right,
                fg_color=COLORS["input"],
                border_color=COLORS["border"],
                border_width=1,
                text_color=COLORS["text"],
                font=(self.font_family, 12),
                wrap="word",
            )
            self.preview_box.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
            set_textbox_text(self.preview_box, UI_TEXT["empty_preview"])

            self._section_title(right, UI_TEXT["tags_title"], 2)
            self.tags_var = ctk.StringVar(value=UI_TEXT["empty_tags"])
            ctk.CTkLabel(
                right,
                textvariable=self.tags_var,
                text_color=COLORS["muted"],
                font=(self.font_family, 12),
                wraplength=320,
                justify="left",
            ).grid(row=3, column=0, sticky="ew", padx=12, pady=(0, 10))

            self._section_title(right, UI_TEXT["section_related_memory"], 4)
            self.related_box = ctk.CTkTextbox(
                right,
                height=92,
                fg_color=COLORS["input"],
                border_color=COLORS["border"],
                border_width=1,
                text_color=COLORS["muted"],
                font=(self.font_family, 12),
                wrap="word",
            )
            self.related_box.grid(row=5, column=0, sticky="ew", padx=10, pady=(0, 10))
            set_textbox_text(self.related_box, UI_TEXT["related_memory_empty"])

            self._section_title(right, UI_TEXT["handoff_title"], 6)
            self.handoff_box = ctk.CTkTextbox(
                right,
                height=118,
                fg_color=COLORS["input"],
                border_color=COLORS["border"],
                border_width=1,
                text_color=COLORS["muted"],
                font=(self.font_family, 12),
                wrap="word",
            )
            self.handoff_box.grid(row=7, column=0, sticky="nsew", padx=10, pady=(0, 10))
            set_textbox_text(self.handoff_box, UI_TEXT["empty_handoff"])

            export_row = ctk.CTkFrame(right, fg_color="transparent")
            export_row.grid(row=8, column=0, sticky="ew", padx=10, pady=(0, 10))
            export_row.grid_columnconfigure((0, 1), weight=1)
            ctk.CTkButton(
                export_row,
                text=UI_TEXT["button_chatgpt"],
                height=34,
                fg_color=COLORS["panel_soft"],
                hover_color=COLORS["accent_soft"],
                command=lambda: self._export_handoff("chatgpt"),
            ).grid(row=0, column=0, sticky="ew", padx=(0, 6))
            ctk.CTkButton(
                export_row,
                text=UI_TEXT["button_codex"],
                height=34,
                fg_color=COLORS["panel_soft"],
                hover_color=COLORS["accent_soft"],
                command=lambda: self._export_handoff("codex"),
            ).grid(row=0, column=1, sticky="ew", padx=(6, 0))

            bottom = self._panel(self, 0)
            bottom.grid(row=2, column=0, columnspan=3, sticky="ew", padx=18, pady=(0, 16))
            bottom.grid_columnconfigure(0, weight=1)
            self._section_title(bottom, UI_TEXT["log_title"], 0)
            self.log_box = ctk.CTkTextbox(
                bottom,
                height=96,
                fg_color=COLORS["input"],
                text_color=COLORS["muted"],
                font=(self.font_family, 12),
                wrap="word",
            )
            self.log_box.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 10))
            self.log_box.configure(state="disabled")

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
                text_color=COLORS["text"],
                font=(self.font_family, 14, "bold"),
            ).grid(row=row, column=0, sticky="w", padx=12, pady=(12, 8))

        def _render_empty_results(self) -> None:
            for child in self.results_frame.winfo_children():
                child.destroy()
            ctk.CTkLabel(
                self.results_frame,
                text=UI_TEXT["empty_results"],
                text_color=COLORS["quiet"],
                font=(self.font_family, 13),
            ).pack(fill="x", padx=12, pady=18)

        def _append_log(self, message: str) -> None:
            self.log_box.configure(state="normal")
            self.log_box.insert("end", f"{message}\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")

        def _set_memory_folder(self, folder: str, persist: bool = True) -> None:
            clean = str(folder or "")
            self.config_data.memory_folder = clean
            self.memory_var.set(clean if clean else UI_TEXT["empty_memory"])
            if persist:
                self.config_store.save(self.config_data)
                if clean:
                    self._append_log(UI_TEXT["log_folder_set"].format(folder=clean))

        def _choose_memory_folder(self) -> None:
            folder = filedialog.askdirectory(title=UI_TEXT["choose_memory_title"])
            if folder:
                self._set_memory_folder(folder)

        def _choose_chatgpt_export(self) -> None:
            file_path = filedialog.askopenfilename(
                title=UI_TEXT["dialog_select_chatgpt_export"],
                filetypes=[
                    (UI_TEXT["filetype_zip"], "*.zip"),
                    (UI_TEXT["filetype_all"], "*.*"),
                ],
            )
            selected = file_path
            if not selected:
                selected = filedialog.askdirectory(title=UI_TEXT["dialog_select_chatgpt_export_folder"])
            if selected:
                self._start_chatgpt_import(Path(selected))

        def _start_chatgpt_import(self, source_path: Path) -> None:
            if self.import_thread and self.import_thread.is_alive():
                return
            self._set_import_buttons_state("disabled")
            self.index_status_var.set(UI_TEXT["status_importing_chatgpt"])
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
                text_color=COLORS["text"],
                font=(self.font_family, 14, "bold"),
            ).grid(row=0, column=0, sticky="w", padx=14, pady=(14, 8))
            input_box = ctk.CTkTextbox(
                dialog,
                fg_color=COLORS["input"],
                border_color=COLORS["border"],
                border_width=1,
                text_color=COLORS["text"],
                font=(self.font_family, 12),
                wrap="word",
            )
            input_box.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 12))

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
                height=34,
                fg_color=COLORS["accent"],
                hover_color=COLORS["accent_hover"],
                command=import_paste,
            ).grid(row=0, column=0, sticky="ew", padx=(0, 6))
            ctk.CTkButton(
                button_row,
                text=UI_TEXT["button_import_codex_file"],
                height=34,
                fg_color=COLORS["panel_soft"],
                hover_color=COLORS["accent_soft"],
                command=import_file,
            ).grid(row=0, column=1, sticky="ew", padx=6)
            ctk.CTkButton(
                button_row,
                text=UI_TEXT["button_close"],
                height=34,
                fg_color=COLORS["panel_soft"],
                hover_color=COLORS["accent_soft"],
                command=dialog.destroy,
            ).grid(row=0, column=2, sticky="ew", padx=(6, 0))

        def _set_import_buttons_state(self, state: str) -> None:
            self.import_button.configure(state=state)
            self.codex_import_button.configure(state=state)

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

        def _cancel_index(self) -> None:
            self.cancel_event.set()
            self._append_log(UI_TEXT["log_index_cancel_requested"])

        def _start_search(self) -> None:
            query_text = self.search_entry.get().strip()
            if not query_text:
                return
            self.current_query = query_text
            self.config_data.last_query = query_text
            self.config_store.save(self.config_data)
            self.index_status_var.set(UI_TEXT["searching"])
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

        def _render_results(self, results: list[SearchResult]) -> None:
            for child in self.results_frame.winfo_children():
                child.destroy()
            if not results:
                self._render_empty_results()
                return

            for result in results:
                item = ctk.CTkFrame(self.results_frame, fg_color=COLORS["panel_alt"], corner_radius=6)
                item.pack(fill="x", padx=4, pady=5)
                item.grid_columnconfigure(0, weight=1)
                title = ctk.CTkButton(
                    item,
                    text=self._result_title(result),
                    anchor="w",
                    fg_color="transparent",
                    hover_color=COLORS["accent_soft"],
                    text_color=COLORS["text"],
                    font=(self.font_family, 14, "bold"),
                    command=lambda selected=result: self._select_result(selected),
                )
                title.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 2))
                ctk.CTkLabel(
                    item,
                    text=self._result_meta(result),
                    text_color=COLORS["muted"],
                    font=(self.font_family, 11),
                    anchor="w",
                ).grid(row=1, column=0, sticky="ew", padx=12)
                ctk.CTkLabel(
                    item,
                    text=result.snippet,
                    text_color=COLORS["muted"],
                    font=(self.font_family, 12),
                    wraplength=510,
                    justify="left",
                    anchor="w",
                ).grid(row=2, column=0, sticky="ew", padx=12, pady=(3, 10))

        def _result_title(self, result: SearchResult) -> str:
            if result.source_type == "chatgpt_export":
                return f"[chatgpt_export] {result.conversation_title or result.title}"
            if result.source_type == "codex_result":
                suffix = f" / {result.commit_hash[:7]}" if result.commit_hash else ""
                return f"[codex_result] {result.title}{suffix}"
            return f"[file] {result.title}"

        def _result_meta(self, result: SearchResult) -> str:
            if result.source_type == "chatgpt_export":
                label = result.source_label or f"ChatGPT / {result.conversation_title or result.title} / {result.role}"
                base = UI_TEXT["result_meta_chatgpt"].format(source_label=label, score=result.score)
                return self._append_semantic_meta(base, result)
            if result.source_type == "codex_result":
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
            set_textbox_text(self.preview_box, preview)
            self.tags_var.set(self._tags_for_result(result))

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

        def _export_handoff(self, kind: str) -> None:
            if not self.current_results:
                messagebox.showwarning(UI_TEXT["dialog_title"], UI_TEXT["dialog_no_results"])
                return
            if kind == "chatgpt":
                path = write_chatgpt_handoff(self.current_query, self.current_results)
            else:
                path = write_codex_handoff(self.current_query, self.current_results)
            self._append_log(UI_TEXT["log_export"].format(path=path))
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

            self.after(100, self._poll_events)

        def _handle_index_progress(self, progress: IndexProgress) -> None:
            if progress.total:
                self.progress.set(progress.current / progress.total)
            else:
                self.progress.set(0)
            if progress.done:
                self.index_button.configure(state="normal")
                self.cancel_button.configure(state="disabled")
                self.config_data.last_indexed_at = progress.message
                self.config_store.save(self.config_data)
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
            else:
                if progress.message == "embedding_chunks":
                    self.index_status_var.set(UI_TEXT["index_embedding"])
                else:
                    self.index_status_var.set(
                        UI_TEXT["index_running"].format(current=progress.current, total=progress.total)
                    )

        def _handle_index_error(self, error_text: str) -> None:
            self.index_button.configure(state="normal")
            self.cancel_button.configure(state="disabled")
            self._set_import_buttons_state("normal")
            self.index_status_var.set(UI_TEXT["index_idle"])
            self._append_log(UI_TEXT["log_index_error"].format(error=error_text))

        def _handle_chatgpt_import_done(self, result: ChatGPTImportResult) -> None:
            self._set_import_buttons_state("normal")
            self.index_status_var.set(UI_TEXT["status_chatgpt_import_complete"])
            self.progress.set(1)
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
            messagebox.showinfo(UI_TEXT["dialog_title"], UI_TEXT["status_chatgpt_import_complete"])

        def _handle_chatgpt_import_missing(self) -> None:
            self._set_import_buttons_state("normal")
            self.index_status_var.set(UI_TEXT["index_idle"])
            self.progress.set(0)
            self._append_log(UI_TEXT["error_conversations_json_not_found"])
            messagebox.showwarning(UI_TEXT["dialog_title"], UI_TEXT["error_conversations_json_not_found"])

        def _handle_chatgpt_import_error(self, error_text: str) -> None:
            self._set_import_buttons_state("normal")
            self.index_status_var.set(UI_TEXT["index_idle"])
            self.progress.set(0)
            self._append_log(f"{UI_TEXT['error_chatgpt_import_failed']} {error_text}")
            messagebox.showwarning(UI_TEXT["dialog_title"], f"{UI_TEXT['error_chatgpt_import_failed']}\n{error_text}")

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

        def _handle_search_done(self, query_text: str, response: SearchResponse, semantic_enabled: bool) -> None:
            results = response.results
            self.current_results = results
            self.current_related_results = response.related
            self.index_status_var.set(UI_TEXT["index_idle"])
            self._render_results(results)
            self._render_related_memory(response.related)
            self._update_handoff_preview()
            if results:
                self._select_result(results[0])
            self._refresh_stats()
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
