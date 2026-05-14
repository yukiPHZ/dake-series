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
    "button_index": "Index",
    "button_cancel": "Cancel",
    "button_choose": "Choose",
    "button_import_chatgpt_export": "ChatGPT exportを取り込む",
    "button_chatgpt": "ChatGPTまとめ",
    "button_codex": "Codex素材",
    "memory_title": "Memory Folder",
    "index_title": "Index Status",
    "system_title": "System Status",
    "results_title": "Search Results",
    "preview_title": "Preview",
    "tags_title": "Related Tags",
    "handoff_title": "Handoff Summary",
    "log_title": "BRAINZ Log",
    "empty_memory": "未選択",
    "empty_results": "検索結果はまだありません",
    "empty_preview": "結果を選ぶと詳細が表示されます",
    "empty_tags": "タグ候補なし",
    "empty_handoff": "検索結果から引き継ぎ素材を生成できます",
    "choose_memory_title": "記憶フォルダを選択",
    "dialog_title": "補助脳BRAINZ",
    "dialog_memory_missing": "記憶フォルダを選択してください",
    "dialog_no_results": "検索結果がありません",
    "dialog_export_done": "ファイルを生成しました",
    "dialog_select_chatgpt_export": "ChatGPT export zipを選択（キャンセルでフォルダ選択）",
    "dialog_select_chatgpt_export_folder": "展開済みChatGPT exportフォルダを選択",
    "filetype_zip": "zipファイル",
    "filetype_all": "すべてのファイル",
    "status_importing_chatgpt": "CHATGPT EXPORT IMPORTING...",
    "status_chatgpt_import_complete": "CHATGPT EXPORT IMPORT COMPLETE",
    "error_conversations_json_not_found": "conversations.json が見つかりませんでした。zipまたは展開済みフォルダを確認してください。",
    "error_chatgpt_import_failed": "ChatGPT exportを取り込めませんでした。",
    "log_chatgpt_export_detected": "ChatGPT export detected.",
    "log_conversations_json_found": "conversations.json found: {path}",
    "log_chatgpt_import_complete": "{conversations} conversations imported. {messages} messages indexed. Skipped duplicates: {skipped}. Errors: {errors}.",
    "log_chatgpt_memory_imported": "補助脳：ChatGPTの記憶を取り込みました。",
    "log_chatgpt_import_file": "IMPORT LOG: {path}",
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
    "status_docs": "DOCS {documents} / CHUNKS {chunks}",
    "preview_template": "PATH: {path}\nSOURCE: {source_type}\nLABEL: {source_label}\nCONVERSATION: {conversation_title}\nROLE: {role}\nMESSAGE INDEX: {message_index}\nMODIFIED: {modified_at}\nINDEXED: {indexed_at}\nSCORE: {score:.2f}\n\n{content}",
    "result_meta": "{source_type} | score {score:.1f}",
    "result_meta_chatgpt": "{source_label} | score {score:.1f}",
    "handoff_preview": "query: {query}\nresults: {count}\n\n{items}",
    "launch_check_ok": "LAUNCH CHECK OK",
}


def run_smoke_test() -> int:
    from core.app_config import ensure_app_dirs
    from core.chatgpt_importer import import_chatgpt_export
    from core.db import BrainzDatabase
    from core.gpu_checker import check_gpu
    from core.handoff_writer import write_chatgpt_handoff, write_codex_handoff
    from core.indexer import Indexer
    from core.ollama_client import check_ollama
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

    ollama_status = check_ollama()
    gpu_status = check_gpu()
    stats = database.stats()
    print("SMOKE OK")
    print(f"documents={stats['documents']} chunks={stats['chunks']}")
    print(f"file_results={len(file_results)}")
    print(f"chatgpt_results={len(chatgpt_results)}")
    print(f"chatgpt_handoff={chatgpt_path}")
    print(f"codex_handoff={codex_path}")
    print(f"ollama_available={ollama_status.available}")
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
    from core.db import BrainzDatabase, SearchResult
    from core.gpu_checker import check_gpu
    from core.handoff_writer import write_chatgpt_handoff, write_codex_handoff
    from core.indexer import IndexProgress, Indexer
    from core.ollama_client import check_ollama
    from core.search_engine import SearchEngine
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
            self.current_query = self.config_data.last_query
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

            self._section_title(left, UI_TEXT["index_title"], 4)
            self.index_status_var = ctk.StringVar(value=UI_TEXT["index_idle"])
            ctk.CTkLabel(
                left,
                textvariable=self.index_status_var,
                text_color=COLORS["text"],
                font=(self.font_family, 13, "bold"),
            ).grid(row=5, column=0, sticky="w", padx=14, pady=(0, 8))
            self.progress = ctk.CTkProgressBar(left, height=10, progress_color=COLORS["accent"])
            self.progress.set(0)
            self.progress.grid(row=6, column=0, sticky="ew", padx=14, pady=(0, 12))

            button_row = ctk.CTkFrame(left, fg_color="transparent")
            button_row.grid(row=7, column=0, sticky="ew", padx=14, pady=(0, 18))
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

            self._section_title(left, UI_TEXT["system_title"], 8)
            self.sqlite_var = ctk.StringVar(value=UI_TEXT["status_sqlite_ready"])
            self.cuda_var = ctk.StringVar(value=UI_TEXT["status_cuda_unavailable"])
            self.gpu_var = ctk.StringVar(value=UI_TEXT["status_gpu_missing"])
            self.ollama_var = ctk.StringVar(value=UI_TEXT["status_ollama_not_running"])
            self.docs_var = ctk.StringVar(value=UI_TEXT["status_docs"].format(documents=0, chunks=0))
            for index, variable in enumerate((self.sqlite_var, self.cuda_var, self.gpu_var, self.ollama_var, self.docs_var), start=9):
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
            right.grid_rowconfigure(5, weight=1)
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

            self._section_title(right, UI_TEXT["handoff_title"], 4)
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
            self.handoff_box.grid(row=5, column=0, sticky="nsew", padx=10, pady=(0, 10))
            set_textbox_text(self.handoff_box, UI_TEXT["empty_handoff"])

            export_row = ctk.CTkFrame(right, fg_color="transparent")
            export_row.grid(row=6, column=0, sticky="ew", padx=10, pady=(0, 10))
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
            self.import_button.configure(state="disabled")
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
            self.search_thread = threading.Thread(target=self._search_worker, args=(query_text,), daemon=True)
            self.search_thread.start()

        def _search_worker(self, query_text: str) -> None:
            try:
                results = self.search_engine.search(query_text)
                self.events.put(("search_done", (query_text, results)))
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
            badge = "chatgpt_export" if result.source_type == "chatgpt_export" else "file"
            return f"[{badge}] {result.conversation_title or result.title}"

        def _result_meta(self, result: SearchResult) -> str:
            if result.source_type == "chatgpt_export":
                label = result.source_label or f"ChatGPT / {result.conversation_title or result.title} / {result.role}"
                return UI_TEXT["result_meta_chatgpt"].format(source_label=label, score=result.score)
            return UI_TEXT["result_meta"].format(source_type=result.source_type, score=result.score)

        def _select_result(self, result: SearchResult) -> None:
            preview = UI_TEXT["preview_template"].format(
                path=result.path,
                source_type=result.source_type,
                source_label=result.source_label,
                conversation_title=result.conversation_title,
                role=result.role,
                message_index=result.message_index,
                modified_at=result.modified_at,
                indexed_at=result.indexed_at,
                score=result.score,
                content=result.content[:9000],
            )
            set_textbox_text(self.preview_box, preview)
            self.tags_var.set(self._tags_for_result(result))

        def _tags_for_result(self, result: SearchResult) -> str:
            tags = [
                result.source_type,
                result.role,
                result.conversation_title,
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
            self.events.put(("system_status", (gpu, ollama)))

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
                    query_text, results = payload
                    self._handle_search_done(query_text, results)
                elif event == "system_status":
                    gpu, ollama = payload
                    self._handle_system_status(gpu, ollama)
                elif event == "chatgpt_import_done":
                    self._handle_chatgpt_import_done(payload)
                elif event == "chatgpt_import_missing":
                    self._handle_chatgpt_import_missing()
                elif event == "chatgpt_import_error":
                    self._handle_chatgpt_import_error(str(payload))

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
                self.index_status_var.set(
                    UI_TEXT["index_running"].format(current=progress.current, total=progress.total)
                )

        def _handle_index_error(self, error_text: str) -> None:
            self.index_button.configure(state="normal")
            self.cancel_button.configure(state="disabled")
            self.import_button.configure(state="normal")
            self.index_status_var.set(UI_TEXT["index_idle"])
            self._append_log(UI_TEXT["log_index_error"].format(error=error_text))

        def _handle_chatgpt_import_done(self, result: ChatGPTImportResult) -> None:
            self.import_button.configure(state="normal")
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
            self.import_button.configure(state="normal")
            self.index_status_var.set(UI_TEXT["index_idle"])
            self.progress.set(0)
            self._append_log(UI_TEXT["error_conversations_json_not_found"])
            messagebox.showwarning(UI_TEXT["dialog_title"], UI_TEXT["error_conversations_json_not_found"])

        def _handle_chatgpt_import_error(self, error_text: str) -> None:
            self.import_button.configure(state="normal")
            self.index_status_var.set(UI_TEXT["index_idle"])
            self.progress.set(0)
            self._append_log(f"{UI_TEXT['error_chatgpt_import_failed']} {error_text}")
            messagebox.showwarning(UI_TEXT["dialog_title"], f"{UI_TEXT['error_chatgpt_import_failed']}\n{error_text}")

        def _handle_search_done(self, query_text: str, results: list[SearchResult]) -> None:
            self.current_results = results
            self.index_status_var.set(UI_TEXT["index_idle"])
            self._render_results(results)
            self._update_handoff_preview()
            if results:
                self._select_result(results[0])
            self._refresh_stats()
            self._append_log(UI_TEXT["log_search"].format(query=query_text, count=len(results)))

        def _handle_system_status(self, gpu, ollama) -> None:
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
