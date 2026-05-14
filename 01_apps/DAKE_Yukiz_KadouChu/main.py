from __future__ import annotations

import argparse
import json
import os
import queue
import sys
import tempfile
import threading
import traceback
import wave
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox

try:
    import customtkinter as ctk
except ImportError:  # pragma: no cover - exercised on machines without GUI deps
    ctk = None

try:
    from PIL import Image
except ImportError:  # pragma: no cover - logo is optional
    Image = None

from core.app_config import (
    APP_NAME,
    COPYRIGHT,
    EXE_NAME,
    LOG_TEXT,
    WINDOW_TITLE,
    app_root,
    ensure_app_dirs,
    estimate_processing_seconds,
    format_duration,
    human_size,
    seconds_to_timecode,
)
from core.audio_preview import AudioPreviewPlayer
from core.project_bridge import (
    BRIDGE_TEXT,
    add_bgm_to_video_box,
    generate_bridge_metadata_draft,
    list_project_boxes,
    read_project_box,
)
from core.memory_store import (
    ensure_memory_dirs,
    generate_memory_summary,
    memory_index_path,
    memory_summary_path,
    save_package_to_memory,
)
from core.memory_analyzer import (
    analyze_memory,
    generate_assistant_recommendation,
)
from core.cli_checker import (
    CLI_TOOLS,
    check_cli_environment,
    fetch_youtube_metadata,
    run_system_check,
)
from core.assistant_review import find_latest_package_dir, run_assistant_review
from core.ffmpeg_runner import create_preview_clip
from core.first_video_test import first_video_test_dir, run_first_video_test
from core.media_probe import MediaInfo, probe_media
from core.ollama_client import build_metadata_draft
from core.posting_package import generate_posting_package, packages_dir
from core.project_writer import (
    ProjectPaths,
    create_project,
    write_log_files,
    write_media_info,
    write_metadata_files,
    write_preview_note,
    write_source_manifest,
)
from core.selected_outputs import export_selected_draft, read_selected_candidates
from core.selected_preview import find_source_video_path, generate_selected_short_preview, generate_vertical_short
from core.sequence_builder import (
    generate_horizontal_edit,
    read_sequence,
    sequence_total_duration,
    write_sequence,
)
from core.shorts_analyzer import create_shorts_candidates, write_shorts_candidates
from core.transcription import is_faster_whisper_available, transcribe_media
from ui.theme import COLORS, FONT_FAMILY, setup_theme

if ctk is not None:
    from ui.components import StatusPill, make_panel, set_textbox


APP_VERSION = "0.1.0"
VIDEO_EXTENSIONS = (".mp4", ".mov", ".mkv", ".webm")

UI_TEXT = {
    "subtitle": "GPU Assisted Production Console",
    "soul_phrase": "稼働中。",
    "input": "INPUT",
    "drop_hint": "Drop video files here",
    "select_video": "Select Video File",
    "youtube_url": "YouTube LIVE URL",
    "fetch_metadata": "Fetch Metadata",
    "status_strip": "STATUS",
    "next_action": "NEXT ACTION",
    "dashboard_waiting": "WAITING",
    "dashboard_connected": "CONNECTED",
    "dashboard_package_ready": "READY",
    "dashboard_entries": "entries",
    "dashboard_gpu": "GPU",
    "dashboard_ollama": "Ollama",
    "dashboard_ffmpeg": "FFmpeg",
    "dashboard_package": "Package",
    "dashboard_memory": "Memory",
    "dashboard_bridge": "Bridge",
    "all_set": "整っています。",
    "process": "PROCESS",
    "video": "VIDEO",
    "start": "Start Production Run",
    "run_system_check": "Run System Check",
    "recheck_system": "Recheck System",
    "open_install_guide": "Open Install Guide",
    "copy_install_commands": "Copy Install Commands",
    "checking": "Checking...",
    "system_check_done": "System check completed.\n整っています。",
    "no_install_guide": "Install guide is not ready yet.",
    "install_guide_open_failed": "Could not open install guide.",
    "install_commands_copied": "Install commands copied.",
    "no_install_commands": "No missing CLI install commands.",
    "output": "OUTPUT",
    "open_output": "Open Output Folder",
    "posting_package": "POSTING PACKAGE",
    "generate_posting_package": "Generate Posting Package",
    "open_package_folder": "Open Package Folder",
    "package_status": "Status",
    "package": "Package",
    "package_ready_hint": "Select a video, then generate the YouTube posting package.",
    "package_requires_video": "Select a video file first.",
    "package_running": "Status: RUNNING",
    "package_completed": "Completed",
    "package_failed": "FAILED",
    "package_output_unavailable": "Package output is not ready yet.",
    "generated": "Generated",
    "assistant_review": "ASSISTANT REVIEW",
    "select_package_folder": "Select Package Folder",
    "run_assistant_review": "Run Assistant Review",
    "open_review_file": "Open Review File",
    "review_ready_hint": "Generate a posting package, then run the assistant review.",
    "review_requires_package": "Posting package is not ready yet.",
    "review_running": "Status: RUNNING",
    "review_completed": "COMPLETED",
    "review_failed": "FAILED",
    "review": "Review",
    "review_file_unavailable": "Review file is not ready yet.",
    "open_review_failed": "Could not open review file.",
    "files_read": "Files read",
    "selected_outputs": "SELECTED OUTPUTS",
    "refresh_candidates": "Refresh Candidates",
    "export_selected_draft": "Export Selected Draft",
    "open_selected_folder": "Open Selected Folder",
    "selected_ready_hint": "Refresh candidates from the posting package.",
    "selected_requires_package": "Posting package is not ready yet.",
    "selected_short": "Short",
    "selected_title": "Title",
    "selected_none": "No candidates loaded",
    "selected_description": "Description",
    "selected_tags": "Tags",
    "selected_notes": "Upload Notes",
    "selected_review": "Assistant Review",
    "selected_available": "available",
    "selected_missing": "missing",
    "selected_completed": "Selected draft exported.",
    "selected_folder_unavailable": "Selected folder is not ready yet.",
    "open_selected_failed": "Could not open selected folder.",
    "short_preview": "SHORT PREVIEW",
    "generate_short_preview": "Generate Selected Short Preview",
    "open_short_preview": "Open Short Preview",
    "short_preview_ready_hint": "Generate a short preview from selected_short.json.",
    "short_preview_running": "Status: RUNNING",
    "short_preview_completed": "Completed",
    "short_preview_failed": "FAILED",
    "short_preview_unavailable": "Short preview is not ready yet.",
    "open_short_preview_failed": "Could not open short preview.",
    "preview_source_dialog": "Select source video for short preview",
    "encoder": "Encoder",
    "preview_output": "Output",
    "vertical_short": "VERTICAL SHORT",
    "generate_vertical_short": "Generate 9:16 Short",
    "open_vertical_short": "Open 9:16 Short",
    "vertical_short_ready_hint": "Generate a 1080x1920 vertical short from selected_short.json.",
    "vertical_short_running": "Status: RUNNING",
    "vertical_short_completed": "Completed",
    "vertical_short_failed": "FAILED",
    "vertical_short_unavailable": "9:16 short is not ready yet.",
    "open_vertical_short_failed": "Could not open 9:16 short.",
    "output_size": "Size",
    "sequence_builder": "SEQUENCE BUILDER",
    "sequence": "SEQUENCE",
    "add_sequence_video": "Add Sequence Video",
    "remove_selected": "Remove Selected",
    "move_up": "Move Up",
    "move_down": "Move Down",
    "generate_horizontal_edit": "Generate Horizontal Edit",
    "open_horizontal_edit": "Open Horizontal Edit",
    "horizontal_edit": "HORIZONTAL EDIT",
    "horizontal_edit_ready_hint": "Add videos, then generate a quiet horizontal edit.",
    "horizontal_edit_running": "Status: RUNNING",
    "horizontal_edit_completed": "Completed",
    "horizontal_edit_failed": "FAILED",
    "horizontal_edit_unavailable": "Horizontal edit is not ready yet.",
    "open_horizontal_edit_failed": "Could not open horizontal edit.",
    "sequence_requires_package": "Select or generate a posting package first.",
    "sequence_empty": "No sequence videos",
    "sequence_no_selection": "Select a sequence video first.",
    "sequence_duration": "Duration",
    "sequence_videos": "Videos",
    "sequence_recommendation": "Recommendation",
    "sequence_log_arranging": "補助脳：素材を並べています。",
    "project_bridge": "PROJECT BRIDGE",
    "refresh_projects": "Refresh Projects",
    "no_project_boxes": "No Project Boxes Found",
    "selected_project": "Selected Project",
    "preset": "Preset",
    "suggested_use": "Suggested Use",
    "bgm": "BGM",
    "bgm_none": "No BGM found",
    "preview_start": "Preview Start",
    "stop_preview": "Stop Preview",
    "add_to_video_box": "Add to Current Video Box",
    "generate_upload_metadata": "Generate Upload Metadata",
    "project_bridge_ready_hint": "Refresh Project Boxes from DAKE_Music_Otooku.",
    "project_notes_unavailable": "Project notes unavailable.",
    "project_bridge_requires_package": "Select or generate a posting package first.",
    "project_bridge_requires_project": "Select a Project Box first.",
    "project_bridge_requires_bgm": "Select a BGM file first.",
    "project_bridge_copy_completed": "BGM added to selected/bgm.",
    "project_bridge_metadata_completed": "metadata_draft.txt generated.",
    "preview_failed": "Preview failed.",
    "preview_stopped": "Preview stopped.",
    "metadata": "Metadata",
    "memory": "MEMORY",
    "save_to_memory": "Save to Memory",
    "open_memory_folder": "Open Memory Folder",
    "generate_memory_summary": "Generate Memory Summary",
    "memory_ready_hint": "Save the current package into the assistant memory.",
    "memory_status": "Status",
    "memory_saved": "SAVED",
    "memory_updated": "UPDATED",
    "memory_requires_package": "Select or generate a posting package first.",
    "memory_folder_unavailable": "Memory folder is not ready yet.",
    "open_memory_failed": "Could not open memory folder.",
    "memory_entries": "Entries",
    "memory_record": "Record",
    "memory_summary": "Summary",
    "projects_loaded": "projects loaded",
    "assistant_recommend": "ASSISTANT RECOMMEND",
    "generate_recommendation": "Generate Recommendation",
    "open_recommendation": "Open Recommendation",
    "refresh_memory": "Refresh Memory",
    "recommend_ready_hint": "Generate a recommendation from memory and the current package.",
    "recommend_requires_package": "Select or generate a posting package first.",
    "recommend_file_unavailable": "Recommendation file is not ready yet.",
    "open_recommend_failed": "Could not open recommendation.",
    "memory_loaded": "Memory",
    "current_mood": "Current Mood",
    "related_projects": "Related Projects",
    "transcription_short": "Transcription",
    "shorts": "Shorts",
    "ollama": "Ollama",
    "used": "used",
    "template_fallback": "template fallback",
    "open_output_failed": "Could not open output folder.",
    "system": "SYSTEM / CLI STATUS",
    "assistant_log": "補助脳 LOG",
    "ready": "READY",
    "running": "RUNNING",
    "complete": "整っています。",
    "error": "ERROR",
    "no_file": "Select a video file first.",
    "file_types": "Video files",
    "select_video_dialog": "Select video file",
    "supported_files": "Supported files: mp4, mov, mkv, webm",
    "no_video_selected": "No video selected",
    "metadata_fetch_optional": "Metadata fetch is optional in Phase 1.",
    "system_check_not_run": "System check has not run yet.",
    "eta_empty": "ETA --",
    "eta_calculating": "ETA calculating...",
    "finish_empty": "Expected Finish --",
    "no_output": "No output yet",
    "media_placeholder": "Media information will appear here.",
    "transcription": "Transcription",
    "scene_analysis": "Scene Analysis",
    "shorts_candidates": "Shorts Candidates",
    "thumbnail_base": "Thumbnail Base",
    "export_package": "Export Package",
    "first_video_test": "FIRST VIDEO TEST",
    "select_test_video": "Select Test Video",
    "run_first_video_test": "Run First Video Test",
    "open_test_output": "Open Test Output",
    "test_selected": "Selected",
    "test_not_selected": "No test video selected",
    "test_clip": "Test Clip",
    "nvenc": "NVENC",
    "codec": "Codec",
    "resolution": "Resolution",
    "duration": "Duration",
    "first_test_ready_hint": "Select a video to run the first 10-second test.",
    "first_test_requires_video": "Select a test video first.",
    "first_test_ffmpeg_required": "FFmpeg is required for first video test.",
    "first_test_completed": "Completed",
    "open_test_output_unavailable": "Test output is not ready yet.",
    "phase2": "PHASE 2",
    "waiting": "WAITING",
    "done": "DONE",
    "working": "WORKING",
    "skipped": "SKIPPED",
    "beside": "側に。",
}


class KadouChuApp(ctk.CTk if ctk is not None else object):  # type: ignore[misc]
    def __init__(self) -> None:
        super().__init__()
        ensure_app_dirs()

        self.title(WINDOW_TITLE)
        self.geometry("1280x780")
        self.minsize(1120, 700)
        self.configure(fg_color=COLORS["bg"])
        self._apply_icon()

        self.events: queue.Queue[dict[str, object]] = queue.Queue()
        self.selected_file: Path | None = None
        self.current_project: ProjectPaths | None = None
        self.current_media_info: MediaInfo | None = None
        self.cli_status: dict[str, dict[str, str | None]] = {}
        self.nvenc_status: dict[str, str] = {"state": "CHECKING", "detail": ""}
        self.gpu_status: dict[str, str] = {"state": "CHECKING", "detail": ""}
        self.install_guide_path: Path | None = None
        self.install_commands: list[str] = []
        self.test_video_path: Path | None = None
        self.test_output_dir: Path | None = None
        self.package_output_dir: Path | None = None
        self.review_file_path: Path | None = None
        self.selected_output_dir: Path | None = None
        self.short_preview_path: Path | None = None
        self.vertical_short_path: Path | None = None
        self.horizontal_edit_path: Path | None = None
        self.preview_source_video_path: Path | None = None
        self.sequence_items: list[dict[str, object]] = []
        self.project_bridge_boxes: list[dict[str, str]] = []
        self.project_bridge_data: dict[str, object] = {}
        self.project_bridge_metadata_path: Path | None = None
        self.audio_preview = AudioPreviewPlayer()
        self.memory_output_dir: Path | None = None
        self.memory_summary_file_path: Path | None = None
        self.recommendation_path: Path | None = None
        self.candidate_data: dict[str, object] = {}
        self.short_choice_touched = False
        self.title_choice_touched = False
        self.first_test_running = False
        self.package_running = False
        self.review_running = False
        self.selected_export_running = False
        self.short_preview_running = False
        self.vertical_short_running = False
        self.sequence_running = False
        self.project_bridge_running = False
        self.memory_running = False
        self.recommendation_running = False
        self.worker_running = False
        self.system_check_completed = False
        self.log_lines: list[str] = []

        self.file_var = ctk.StringVar(value=UI_TEXT["no_video_selected"])
        self.test_file_var = ctk.StringVar(value=UI_TEXT["test_not_selected"])
        self.first_test_summary_var = ctk.StringVar(value=UI_TEXT["first_test_ready_hint"])
        self.package_summary_var = ctk.StringVar(value=UI_TEXT["package_ready_hint"])
        self.review_summary_var = ctk.StringVar(value=UI_TEXT["review_ready_hint"])
        self.selected_summary_var = ctk.StringVar(value=UI_TEXT["selected_ready_hint"])
        self.short_preview_summary_var = ctk.StringVar(value=UI_TEXT["short_preview_ready_hint"])
        self.vertical_short_summary_var = ctk.StringVar(value=UI_TEXT["vertical_short_ready_hint"])
        self.sequence_summary_var = ctk.StringVar(value=UI_TEXT["horizontal_edit_ready_hint"])
        self.project_bridge_summary_var = ctk.StringVar(value=UI_TEXT["project_bridge_ready_hint"])
        self.memory_summary_var = ctk.StringVar(value=UI_TEXT["memory_ready_hint"])
        self.recommendation_summary_var = ctk.StringVar(value=UI_TEXT["recommend_ready_hint"])
        self.short_choice_var = ctk.StringVar(value=UI_TEXT["selected_none"])
        self.title_choice_var = ctk.StringVar(value=UI_TEXT["selected_none"])
        self.sequence_choice_var = ctk.StringVar(value=UI_TEXT["sequence_empty"])
        self.project_choice_var = ctk.StringVar(value=UI_TEXT["no_project_boxes"])
        self.project_bgm_choice_var = ctk.StringVar(value=UI_TEXT["bgm_none"])
        self.youtube_var = ctk.StringVar(value="")
        self.youtube_status_var = ctk.StringVar(value=UI_TEXT["metadata_fetch_optional"])
        self.system_check_status_var = ctk.StringVar(value=UI_TEXT["system_check_not_run"])
        self.eta_var = ctk.StringVar(value=UI_TEXT["eta_empty"])
        self.finish_var = ctk.StringVar(value=UI_TEXT["finish_empty"])
        self.output_var = ctk.StringVar(value=UI_TEXT["no_output"])
        self.status_var = ctk.StringVar(value=UI_TEXT["ready"])
        self.status_strip_var = ctk.StringVar(value="")
        self.next_action_var = ctk.StringVar(value=UI_TEXT["run_system_check"])
        self.progress_var = ctk.DoubleVar(value=0.0)

        self.step_vars: dict[str, ctk.StringVar] = {}
        self.cli_pills: dict[str, StatusPill] = {}

        self._build_ui()
        self._update_dashboard()
        self._log(LOG_TEXT["startup"])
        self._log(LOG_TEXT["running"])
        self._refresh_project_bridge(silent=True)
        self._start_cli_check()
        self.after(100, self._drain_events)

    def _apply_icon(self) -> None:
        icon_candidates = [
            app_root() / ".." / ".." / "02_assets" / "dake_icon.ico",
            app_root() / "02_assets" / "dake_icon.ico",
        ]
        for icon_path in icon_candidates:
            try:
                resolved = icon_path.resolve()
                if resolved.exists():
                    self.iconbitmap(str(resolved))
                    return
            except Exception:
                continue

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)
        self.grid_rowconfigure(3, weight=0)

        header = ctk.CTkFrame(self, fg_color="transparent", height=86)
        header.grid(row=0, column=0, sticky="ew", padx=20, pady=(16, 8))
        header.grid_columnconfigure(1, weight=1)

        logo_label = self._make_logo_label(header)
        logo_label.grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 14))

        ctk.CTkLabel(
            header,
            text=APP_NAME,
            font=ctk.CTkFont(family=FONT_FAMILY, size=26, weight="bold"),
            text_color=COLORS["text"],
        ).grid(row=0, column=1, sticky="w")
        ctk.CTkLabel(
            header,
            text=UI_TEXT["subtitle"],
            font=ctk.CTkFont(family=FONT_FAMILY, size=13),
            text_color=COLORS["muted"],
        ).grid(row=1, column=1, sticky="w", pady=(2, 0))

        phrase = ctk.CTkLabel(
            header,
            text=UI_TEXT["soul_phrase"],
            font=ctk.CTkFont(family=FONT_FAMILY, size=16, weight="bold"),
            text_color=COLORS["accent"],
            fg_color=COLORS["panel"],
            corner_radius=6,
            padx=14,
            pady=8,
        )
        phrase.grid(row=0, column=2, rowspan=2, sticky="e")

        self._build_dashboard().grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 10))

        main = ctk.CTkFrame(self, fg_color="transparent")
        main.grid(row=2, column=0, sticky="nsew", padx=20, pady=(0, 10))
        main.grid_columnconfigure(0, weight=1, uniform="main")
        main.grid_columnconfigure(1, weight=1, uniform="main")
        main.grid_columnconfigure(2, weight=1, uniform="main")
        main.grid_rowconfigure(0, weight=1)

        self._build_input_panel(main).grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self._build_process_panel(main).grid(row=0, column=1, sticky="nsew", padx=8)
        self._build_output_panel(main).grid(row=0, column=2, sticky="nsew", padx=(8, 0))

        bottom = ctk.CTkFrame(self, fg_color="transparent")
        bottom.grid(row=3, column=0, sticky="ew", padx=20, pady=(0, 16))
        bottom.grid_columnconfigure(0, weight=2)
        bottom.grid_columnconfigure(1, weight=3)

        self._build_system_panel(bottom).grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self._build_log_panel(bottom).grid(row=0, column=1, sticky="nsew", padx=(8, 0))

    def _build_dashboard(self) -> ctk.CTkFrame:
        dashboard = ctk.CTkFrame(self, fg_color="transparent")
        dashboard.grid_columnconfigure(0, weight=3)
        dashboard.grid_columnconfigure(1, weight=1)

        status_card = ctk.CTkFrame(
            dashboard,
            fg_color=COLORS["panel"],
            border_width=1,
            border_color=COLORS["line"],
            corner_radius=8,
        )
        status_card.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        status_card.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            status_card,
            text=UI_TEXT["status_strip"],
            font=ctk.CTkFont(family=FONT_FAMILY, size=10, weight="bold"),
            text_color=COLORS["accent_soft"],
        ).grid(row=0, column=0, sticky="w", padx=12, pady=(8, 0))
        ctk.CTkLabel(
            status_card,
            textvariable=self.status_strip_var,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            text_color=COLORS["text"],
            wraplength=840,
            justify="left",
        ).grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 8))

        next_card = ctk.CTkFrame(
            dashboard,
            fg_color=COLORS["panel"],
            border_width=1,
            border_color=COLORS["line"],
            corner_radius=8,
        )
        next_card.grid(row=0, column=1, sticky="ew", padx=(8, 0))
        next_card.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            next_card,
            text=UI_TEXT["next_action"],
            font=ctk.CTkFont(family=FONT_FAMILY, size=10, weight="bold"),
            text_color=COLORS["accent_soft"],
        ).grid(row=0, column=0, sticky="w", padx=12, pady=(8, 0))
        ctk.CTkLabel(
            next_card,
            textvariable=self.next_action_var,
            font=ctk.CTkFont(family=FONT_FAMILY, size=14, weight="bold"),
            text_color=COLORS["accent"],
            wraplength=300,
            justify="left",
        ).grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 8))
        return dashboard

    def _make_logo_label(self, parent: ctk.CTkFrame) -> ctk.CTkLabel:
        logo_path = app_root() / "assets" / "peakheadz_logo.png"
        if logo_path.exists() and Image is not None:
            try:
                image = Image.open(logo_path)
                self.logo_image = ctk.CTkImage(light_image=image, dark_image=image, size=(52, 52))
                return ctk.CTkLabel(parent, text="", image=self.logo_image)
            except Exception:
                pass
        return ctk.CTkLabel(
            parent,
            text="PEAKHEADZ",
            width=86,
            height=52,
            font=ctk.CTkFont(family=FONT_FAMILY, size=10, weight="bold"),
            text_color=COLORS["muted"],
            fg_color=COLORS["panel"],
            corner_radius=6,
        )

    def _build_input_panel(self, parent: ctk.CTkFrame) -> ctk.CTkFrame:
        panel, body = make_panel(parent, UI_TEXT["input"])
        body.grid_columnconfigure(0, weight=1)

        drop = ctk.CTkFrame(body, fg_color=COLORS["panel_alt"], border_width=1, border_color=COLORS["line"], corner_radius=8)
        drop.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        drop.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            drop,
            text=UI_TEXT["drop_hint"],
            font=ctk.CTkFont(family=FONT_FAMILY, size=13),
            text_color=COLORS["muted"],
        ).grid(row=0, column=0, pady=(18, 4), padx=14)
        ctk.CTkButton(
            drop,
            text=UI_TEXT["select_video"],
            command=self._select_video,
            height=34,
            fg_color=COLORS["button"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        ).grid(row=1, column=0, pady=(4, 18), padx=14)

        ctk.CTkLabel(
            body,
            textvariable=self.file_var,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=COLORS["muted"],
            wraplength=330,
            justify="left",
        ).grid(row=1, column=0, sticky="ew", pady=(0, 18))

        ctk.CTkLabel(
            body,
            text=UI_TEXT["youtube_url"],
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLORS["text"],
        ).grid(row=2, column=0, sticky="w", pady=(4, 6))
        ctk.CTkEntry(
            body,
            textvariable=self.youtube_var,
            placeholder_text="https://www.youtube.com/watch?v=...",
            fg_color=COLORS["field"],
            border_color=COLORS["line"],
            text_color=COLORS["text"],
        ).grid(row=3, column=0, sticky="ew")
        ctk.CTkButton(
            body,
            text=UI_TEXT["fetch_metadata"],
            command=self._fetch_metadata,
            height=32,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        ).grid(row=4, column=0, sticky="ew", pady=(10, 8))
        ctk.CTkLabel(
            body,
            textvariable=self.youtube_status_var,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=COLORS["muted"],
            wraplength=330,
            justify="left",
        ).grid(row=5, column=0, sticky="ew")

        ctk.CTkLabel(
            body,
            text=UI_TEXT["beside"],
            font=ctk.CTkFont(family=FONT_FAMILY, size=16, weight="bold"),
            text_color=COLORS["accent_soft"],
        ).grid(row=6, column=0, sticky="w", pady=(24, 0))
        return panel

    def _build_process_panel(self, parent: ctk.CTkFrame) -> ctk.CTkFrame:
        panel, body = make_panel(parent, UI_TEXT["video"])
        body.grid_columnconfigure(0, weight=1)

        for index, label in enumerate(
            [
                UI_TEXT["transcription"],
                UI_TEXT["scene_analysis"],
                UI_TEXT["shorts_candidates"],
                UI_TEXT["thumbnail_base"],
                UI_TEXT["export_package"],
            ]
        ):
            status = UI_TEXT["phase2"] if label == UI_TEXT["thumbnail_base"] else UI_TEXT["waiting"]
            self.step_vars[label] = ctk.StringVar(value=status)
            row = ctk.CTkFrame(body, fg_color=COLORS["panel_alt"], corner_radius=6)
            row.grid(row=index, column=0, sticky="ew", pady=4)
            row.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(
                row,
                text=label,
                font=ctk.CTkFont(family=FONT_FAMILY, size=12),
                text_color=COLORS["text"],
            ).grid(row=0, column=0, sticky="w", padx=12, pady=9)
            ctk.CTkLabel(
                row,
                textvariable=self.step_vars[label],
                font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"),
                text_color=COLORS["accent_soft"],
            ).grid(row=0, column=1, sticky="e", padx=12, pady=9)

        self.progress = ctk.CTkProgressBar(body, variable=self.progress_var, progress_color=COLORS["accent"])
        self.progress.grid(row=6, column=0, sticky="ew", pady=(20, 8))
        self.progress.set(0)

        time_box = ctk.CTkFrame(body, fg_color="transparent")
        time_box.grid(row=7, column=0, sticky="ew")
        time_box.grid_columnconfigure(0, weight=1)
        time_box.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(time_box, textvariable=self.eta_var, text_color=COLORS["muted"]).grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(time_box, textvariable=self.finish_var, text_color=COLORS["muted"]).grid(row=0, column=1, sticky="e")

        self.start_button = ctk.CTkButton(
            body,
            text=UI_TEXT["start"],
            command=self._start_processing,
            height=42,
            fg_color=COLORS["button"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.start_button.grid(row=8, column=0, sticky="ew", pady=(22, 8))
        self.process_system_check_button = ctk.CTkButton(
            body,
            text=UI_TEXT["run_system_check"],
            command=self._start_cli_check,
            height=32,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.process_system_check_button.grid(row=9, column=0, sticky="ew")

        ctk.CTkLabel(
            body,
            textvariable=self.status_var,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLORS["accent"],
        ).grid(row=10, column=0, sticky="w", pady=(18, 0))

        test_box = ctk.CTkFrame(body, fg_color=COLORS["panel_alt"], border_width=1, border_color=COLORS["line"], corner_radius=8)
        test_box.grid(row=11, column=0, sticky="ew", pady=(16, 0))
        test_box.grid_columnconfigure(0, weight=1)
        test_box.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            test_box,
            text=UI_TEXT["first_video_test"],
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLORS["accent_soft"],
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=12, pady=(10, 4))
        ctk.CTkLabel(
            test_box,
            textvariable=self.test_file_var,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=COLORS["muted"],
            wraplength=330,
            justify="left",
        ).grid(row=1, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 6))
        ctk.CTkButton(
            test_box,
            text=UI_TEXT["select_test_video"],
            command=self._select_test_video,
            height=30,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        ).grid(row=2, column=0, sticky="ew", padx=(12, 4), pady=4)
        self.first_test_button = ctk.CTkButton(
            test_box,
            text=UI_TEXT["run_first_video_test"],
            command=self._start_first_video_test,
            height=30,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.first_test_button.grid(row=2, column=1, sticky="ew", padx=(4, 12), pady=4)
        self.generate_package_video_button = ctk.CTkButton(
            test_box,
            text=UI_TEXT["generate_posting_package"],
            command=self._start_posting_package,
            height=30,
            fg_color=COLORS["button"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.generate_package_video_button.grid(row=3, column=0, columnspan=2, sticky="ew", padx=12, pady=4)
        self.open_test_output_button = ctk.CTkButton(
            test_box,
            text=UI_TEXT["open_test_output"],
            command=self._open_test_output,
            height=30,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
            state="disabled",
        )
        self.open_test_output_button.grid(row=4, column=0, columnspan=2, sticky="ew", padx=12, pady=(4, 8))
        ctk.CTkLabel(
            test_box,
            textvariable=self.first_test_summary_var,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=COLORS["text"],
            wraplength=330,
            justify="left",
        ).grid(row=5, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 12))
        return panel

    def _build_output_panel(self, parent: ctk.CTkFrame) -> ctk.CTkFrame:
        panel, body = make_panel(parent, UI_TEXT["output"])
        body.grid_columnconfigure(0, weight=1)
        body.grid_rowconfigure(0, weight=1)
        scroll_body = ctk.CTkScrollableFrame(body, fg_color="transparent")
        scroll_body.grid(row=0, column=0, sticky="nsew")
        body = scroll_body
        body.grid_columnconfigure(0, weight=1)
        body.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            body,
            textvariable=self.output_var,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=COLORS["muted"],
            wraplength=330,
            justify="left",
        ).grid(row=0, column=0, sticky="ew", pady=(0, 10))

        self.media_box = ctk.CTkTextbox(
            body,
            height=90,
            fg_color=COLORS["field"],
            border_width=1,
            border_color=COLORS["line"],
            text_color=COLORS["text"],
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            wrap="word",
        )
        self.media_box.grid(row=1, column=0, sticky="nsew")
        set_textbox(self.media_box, UI_TEXT["media_placeholder"])

        package_box = ctk.CTkFrame(body, fg_color=COLORS["panel_alt"], border_width=1, border_color=COLORS["line"], corner_radius=8)
        package_box.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        package_box.grid_columnconfigure(0, weight=1)
        package_box.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            package_box,
            text=UI_TEXT["posting_package"],
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLORS["accent_soft"],
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=12, pady=(10, 4))
        ctk.CTkLabel(
            package_box,
            textvariable=self.package_summary_var,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=COLORS["text"],
            wraplength=330,
            justify="left",
        ).grid(row=1, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 8))
        self.generate_package_button = ctk.CTkButton(
            package_box,
            text=UI_TEXT["generate_posting_package"],
            command=self._start_posting_package,
            height=30,
            fg_color=COLORS["button"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.generate_package_button.grid(row=2, column=0, sticky="ew", padx=(12, 4), pady=(0, 10))
        self.open_package_button = ctk.CTkButton(
            package_box,
            text=UI_TEXT["open_package_folder"],
            command=self._open_package_folder,
            height=30,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
            state="disabled",
        )
        self.open_package_button.grid(row=2, column=1, sticky="ew", padx=(4, 12), pady=(0, 10))

        review_box = ctk.CTkFrame(body, fg_color=COLORS["panel_alt"], border_width=1, border_color=COLORS["line"], corner_radius=8)
        review_box.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        review_box.grid_columnconfigure(0, weight=1)
        review_box.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            review_box,
            text=UI_TEXT["assistant_review"],
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLORS["accent_soft"],
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=12, pady=(10, 4))
        ctk.CTkLabel(
            review_box,
            textvariable=self.review_summary_var,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=COLORS["text"],
            wraplength=330,
            justify="left",
        ).grid(row=1, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 8))
        self.select_package_button = ctk.CTkButton(
            review_box,
            text=UI_TEXT["select_package_folder"],
            command=self._select_package_folder,
            height=30,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.select_package_button.grid(row=2, column=0, sticky="ew", padx=(12, 4), pady=4)
        self.assistant_review_button = ctk.CTkButton(
            review_box,
            text=UI_TEXT["run_assistant_review"],
            command=self._start_assistant_review,
            height=30,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.assistant_review_button.grid(row=2, column=1, sticky="ew", padx=(4, 12), pady=4)
        self.open_review_button = ctk.CTkButton(
            review_box,
            text=UI_TEXT["open_review_file"],
            command=self._open_review_file,
            height=30,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
            state="disabled",
        )
        self.open_review_button.grid(row=3, column=0, columnspan=2, sticky="ew", padx=12, pady=(4, 10))

        selected_box = ctk.CTkFrame(body, fg_color=COLORS["panel_alt"], border_width=1, border_color=COLORS["line"], corner_radius=8)
        selected_box.grid(row=4, column=0, sticky="ew", pady=(12, 0))
        selected_box.grid_columnconfigure(0, weight=1)
        selected_box.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            selected_box,
            text=UI_TEXT["selected_outputs"],
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLORS["accent_soft"],
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=12, pady=(10, 4))
        ctk.CTkLabel(
            selected_box,
            textvariable=self.selected_summary_var,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=COLORS["text"],
            wraplength=330,
            justify="left",
        ).grid(row=1, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 8))
        self.short_choice_menu = ctk.CTkOptionMenu(
            selected_box,
            variable=self.short_choice_var,
            values=[UI_TEXT["selected_none"]],
            command=self._on_short_choice,
            height=30,
            fg_color=COLORS["button_secondary"],
            button_color=COLORS["button"],
            button_hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.short_choice_menu.grid(row=2, column=0, sticky="ew", padx=(12, 4), pady=4)
        self.title_choice_menu = ctk.CTkOptionMenu(
            selected_box,
            variable=self.title_choice_var,
            values=[UI_TEXT["selected_none"]],
            command=self._on_title_choice,
            height=30,
            fg_color=COLORS["button_secondary"],
            button_color=COLORS["button"],
            button_hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.title_choice_menu.grid(row=2, column=1, sticky="ew", padx=(4, 12), pady=4)
        self.refresh_candidates_button = ctk.CTkButton(
            selected_box,
            text=UI_TEXT["refresh_candidates"],
            command=self._refresh_selected_candidates,
            height=30,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.refresh_candidates_button.grid(row=3, column=0, sticky="ew", padx=(12, 4), pady=4)
        self.export_selected_button = ctk.CTkButton(
            selected_box,
            text=UI_TEXT["export_selected_draft"],
            command=self._start_export_selected_draft,
            height=30,
            fg_color=COLORS["button"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.export_selected_button.grid(row=3, column=1, sticky="ew", padx=(4, 12), pady=4)
        self.open_selected_button = ctk.CTkButton(
            selected_box,
            text=UI_TEXT["open_selected_folder"],
            command=self._open_selected_folder,
            height=30,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
            state="disabled",
        )
        self.open_selected_button.grid(row=4, column=0, columnspan=2, sticky="ew", padx=12, pady=(4, 10))
        ctk.CTkLabel(
            selected_box,
            textvariable=self.short_preview_summary_var,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=COLORS["muted"],
            wraplength=330,
            justify="left",
        ).grid(row=5, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 8))
        self.generate_short_preview_button = ctk.CTkButton(
            selected_box,
            text=UI_TEXT["generate_short_preview"],
            command=self._start_generate_short_preview,
            height=30,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.generate_short_preview_button.grid(row=6, column=0, sticky="ew", padx=(12, 4), pady=(0, 10))
        self.open_short_preview_button = ctk.CTkButton(
            selected_box,
            text=UI_TEXT["open_short_preview"],
            command=self._open_short_preview,
            height=30,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
            state="disabled",
        )
        self.open_short_preview_button.grid(row=6, column=1, sticky="ew", padx=(4, 12), pady=(0, 10))
        ctk.CTkLabel(
            selected_box,
            textvariable=self.vertical_short_summary_var,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=COLORS["muted"],
            wraplength=330,
            justify="left",
        ).grid(row=7, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 8))
        self.generate_vertical_short_button = ctk.CTkButton(
            selected_box,
            text=UI_TEXT["generate_vertical_short"],
            command=self._start_generate_vertical_short,
            height=30,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.generate_vertical_short_button.grid(row=8, column=0, sticky="ew", padx=(12, 4), pady=(0, 10))
        self.open_vertical_short_button = ctk.CTkButton(
            selected_box,
            text=UI_TEXT["open_vertical_short"],
            command=self._open_vertical_short,
            height=30,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
            state="disabled",
        )
        self.open_vertical_short_button.grid(row=8, column=1, sticky="ew", padx=(4, 12), pady=(0, 10))

        sequence_box = ctk.CTkFrame(body, fg_color=COLORS["panel_alt"], border_width=1, border_color=COLORS["line"], corner_radius=8)
        sequence_box.grid(row=5, column=0, sticky="ew", pady=(12, 0))
        sequence_box.grid_columnconfigure(0, weight=1)
        sequence_box.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            sequence_box,
            text=UI_TEXT["sequence_builder"],
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLORS["accent_soft"],
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=12, pady=(10, 4))
        ctk.CTkLabel(
            sequence_box,
            textvariable=self.sequence_summary_var,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=COLORS["text"],
            wraplength=330,
            justify="left",
        ).grid(row=1, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 8))
        self.sequence_choice_menu = ctk.CTkOptionMenu(
            sequence_box,
            variable=self.sequence_choice_var,
            values=[UI_TEXT["sequence_empty"]],
            height=30,
            fg_color=COLORS["button_secondary"],
            button_color=COLORS["button"],
            button_hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.sequence_choice_menu.grid(row=2, column=0, columnspan=2, sticky="ew", padx=12, pady=4)
        self.add_sequence_button = ctk.CTkButton(
            sequence_box,
            text=UI_TEXT["add_sequence_video"],
            command=self._add_sequence_videos,
            height=30,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.add_sequence_button.grid(row=3, column=0, sticky="ew", padx=(12, 4), pady=4)
        self.remove_sequence_button = ctk.CTkButton(
            sequence_box,
            text=UI_TEXT["remove_selected"],
            command=self._remove_selected_sequence_video,
            height=30,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.remove_sequence_button.grid(row=3, column=1, sticky="ew", padx=(4, 12), pady=4)
        self.move_sequence_up_button = ctk.CTkButton(
            sequence_box,
            text=UI_TEXT["move_up"],
            command=self._move_sequence_up,
            height=30,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.move_sequence_up_button.grid(row=4, column=0, sticky="ew", padx=(12, 4), pady=4)
        self.move_sequence_down_button = ctk.CTkButton(
            sequence_box,
            text=UI_TEXT["move_down"],
            command=self._move_sequence_down,
            height=30,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.move_sequence_down_button.grid(row=4, column=1, sticky="ew", padx=(4, 12), pady=4)
        self.generate_horizontal_edit_button = ctk.CTkButton(
            sequence_box,
            text=UI_TEXT["generate_horizontal_edit"],
            command=self._start_generate_horizontal_edit,
            height=30,
            fg_color=COLORS["button"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.generate_horizontal_edit_button.grid(row=5, column=0, sticky="ew", padx=(12, 4), pady=(4, 10))
        self.open_horizontal_edit_button = ctk.CTkButton(
            sequence_box,
            text=UI_TEXT["open_horizontal_edit"],
            command=self._open_horizontal_edit,
            height=30,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
            state="disabled",
        )
        self.open_horizontal_edit_button.grid(row=5, column=1, sticky="ew", padx=(4, 12), pady=(4, 10))

        bridge_box = ctk.CTkFrame(body, fg_color=COLORS["panel_alt"], border_width=1, border_color=COLORS["line"], corner_radius=8)
        bridge_box.grid(row=6, column=0, sticky="ew", pady=(12, 0))
        bridge_box.grid_columnconfigure(0, weight=1)
        bridge_box.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            bridge_box,
            text=UI_TEXT["project_bridge"],
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLORS["accent_soft"],
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=12, pady=(10, 4))
        ctk.CTkLabel(
            bridge_box,
            textvariable=self.project_bridge_summary_var,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=COLORS["text"],
            wraplength=330,
            justify="left",
        ).grid(row=1, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 8))
        self.project_choice_menu = ctk.CTkOptionMenu(
            bridge_box,
            variable=self.project_choice_var,
            values=[UI_TEXT["no_project_boxes"]],
            command=self._on_project_bridge_choice,
            height=30,
            fg_color=COLORS["button_secondary"],
            button_color=COLORS["button"],
            button_hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.project_choice_menu.grid(row=2, column=0, sticky="ew", padx=(12, 4), pady=4)
        self.project_bgm_menu = ctk.CTkOptionMenu(
            bridge_box,
            variable=self.project_bgm_choice_var,
            values=[UI_TEXT["bgm_none"]],
            height=30,
            fg_color=COLORS["button_secondary"],
            button_color=COLORS["button"],
            button_hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.project_bgm_menu.grid(row=2, column=1, sticky="ew", padx=(4, 12), pady=4)
        self.refresh_projects_button = ctk.CTkButton(
            bridge_box,
            text=UI_TEXT["refresh_projects"],
            command=self._refresh_project_bridge,
            height=30,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.refresh_projects_button.grid(row=3, column=0, sticky="ew", padx=(12, 4), pady=4)
        self.preview_project_bgm_button = ctk.CTkButton(
            bridge_box,
            text=UI_TEXT["preview_start"],
            command=self._start_project_bgm_preview,
            height=30,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.preview_project_bgm_button.grid(row=3, column=1, sticky="ew", padx=(4, 12), pady=4)
        self.stop_project_preview_button = ctk.CTkButton(
            bridge_box,
            text=UI_TEXT["stop_preview"],
            command=self._stop_project_bgm_preview,
            height=30,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.stop_project_preview_button.grid(row=4, column=0, sticky="ew", padx=(12, 4), pady=4)
        self.add_project_bgm_button = ctk.CTkButton(
            bridge_box,
            text=UI_TEXT["add_to_video_box"],
            command=self._start_add_project_bgm,
            height=30,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.add_project_bgm_button.grid(row=4, column=1, sticky="ew", padx=(4, 12), pady=4)
        self.generate_bridge_metadata_button = ctk.CTkButton(
            bridge_box,
            text=UI_TEXT["generate_upload_metadata"],
            command=self._start_generate_bridge_metadata,
            height=30,
            fg_color=COLORS["button"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.generate_bridge_metadata_button.grid(row=5, column=0, columnspan=2, sticky="ew", padx=12, pady=(4, 10))

        memory_box = ctk.CTkFrame(body, fg_color=COLORS["panel_alt"], border_width=1, border_color=COLORS["line"], corner_radius=8)
        memory_box.grid(row=7, column=0, sticky="ew", pady=(12, 0))
        memory_box.grid_columnconfigure(0, weight=1)
        memory_box.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            memory_box,
            text=UI_TEXT["memory"],
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLORS["accent_soft"],
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=12, pady=(10, 4))
        ctk.CTkLabel(
            memory_box,
            textvariable=self.memory_summary_var,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=COLORS["text"],
            wraplength=330,
            justify="left",
        ).grid(row=1, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 8))
        self.save_memory_button = ctk.CTkButton(
            memory_box,
            text=UI_TEXT["save_to_memory"],
            command=self._start_save_memory,
            height=30,
            fg_color=COLORS["button"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.save_memory_button.grid(row=2, column=0, sticky="ew", padx=(12, 4), pady=4)
        self.open_memory_button = ctk.CTkButton(
            memory_box,
            text=UI_TEXT["open_memory_folder"],
            command=self._open_memory_folder,
            height=30,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.open_memory_button.grid(row=2, column=1, sticky="ew", padx=(4, 12), pady=4)
        self.generate_memory_summary_button = ctk.CTkButton(
            memory_box,
            text=UI_TEXT["generate_memory_summary"],
            command=self._start_generate_memory_summary,
            height=30,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.generate_memory_summary_button.grid(row=3, column=0, columnspan=2, sticky="ew", padx=12, pady=(4, 10))

        recommend_box = ctk.CTkFrame(body, fg_color=COLORS["panel_alt"], border_width=1, border_color=COLORS["line"], corner_radius=8)
        recommend_box.grid(row=8, column=0, sticky="ew", pady=(12, 0))
        recommend_box.grid_columnconfigure(0, weight=1)
        recommend_box.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            recommend_box,
            text=UI_TEXT["assistant_recommend"],
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLORS["accent_soft"],
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=12, pady=(10, 4))
        ctk.CTkLabel(
            recommend_box,
            textvariable=self.recommendation_summary_var,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=COLORS["text"],
            wraplength=330,
            justify="left",
        ).grid(row=1, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 8))
        self.generate_recommendation_button = ctk.CTkButton(
            recommend_box,
            text=UI_TEXT["generate_recommendation"],
            command=self._start_generate_recommendation,
            height=30,
            fg_color=COLORS["button"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.generate_recommendation_button.grid(row=2, column=0, sticky="ew", padx=(12, 4), pady=4)
        self.open_recommendation_button = ctk.CTkButton(
            recommend_box,
            text=UI_TEXT["open_recommendation"],
            command=self._open_recommendation,
            height=30,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
            state="disabled",
        )
        self.open_recommendation_button.grid(row=2, column=1, sticky="ew", padx=(4, 12), pady=4)
        self.refresh_memory_button = ctk.CTkButton(
            recommend_box,
            text=UI_TEXT["refresh_memory"],
            command=self._refresh_recommend_memory,
            height=30,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.refresh_memory_button.grid(row=3, column=0, columnspan=2, sticky="ew", padx=12, pady=(4, 10))

        self.open_button = ctk.CTkButton(
            body,
            text=UI_TEXT["open_output"],
            command=self._open_output_folder,
            height=36,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
            state="disabled",
        )
        self.open_button.grid(row=9, column=0, sticky="ew", pady=(12, 0))
        return panel

    def _build_system_panel(self, parent: ctk.CTkFrame) -> ctk.CTkFrame:
        panel, body = make_panel(parent, UI_TEXT["system"])
        body.grid_columnconfigure(0, weight=1)
        body.grid_columnconfigure(1, weight=1)

        labels = ["FFMPEG", "FFPROBE", "YT-DLP", "GH", "WRANGLER", "NPM", "OLLAMA", "NVENC", "GPU"]
        for index, label in enumerate(labels):
            pill = StatusPill(body, label)
            pill.grid(row=index // 2, column=index % 2, sticky="ew", padx=4, pady=4)
            self.cli_pills[label] = pill
        self.system_check_button = ctk.CTkButton(
            body,
            text=UI_TEXT["recheck_system"],
            command=self._start_cli_check,
            height=32,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.system_check_button.grid(row=5, column=0, columnspan=2, sticky="ew", padx=4, pady=(10, 4))
        self.open_install_guide_button = ctk.CTkButton(
            body,
            text=UI_TEXT["open_install_guide"],
            command=self._open_install_guide,
            height=32,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
            state="disabled",
        )
        self.open_install_guide_button.grid(row=6, column=0, sticky="ew", padx=4, pady=4)
        self.copy_install_commands_button = ctk.CTkButton(
            body,
            text=UI_TEXT["copy_install_commands"],
            command=self._copy_install_commands,
            height=32,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
            state="disabled",
        )
        self.copy_install_commands_button.grid(row=6, column=1, sticky="ew", padx=4, pady=4)
        ctk.CTkLabel(
            body,
            textvariable=self.system_check_status_var,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=COLORS["muted"],
            justify="left",
        ).grid(row=7, column=0, columnspan=2, sticky="w", padx=4, pady=(2, 0))
        return panel

    def _build_log_panel(self, parent: ctk.CTkFrame) -> ctk.CTkFrame:
        panel, body = make_panel(parent, UI_TEXT["assistant_log"])
        body.grid_rowconfigure(0, weight=1)
        body.grid_columnconfigure(0, weight=1)
        self.log_box = ctk.CTkTextbox(
            body,
            height=150,
            fg_color=COLORS["field"],
            border_width=1,
            border_color=COLORS["line"],
            text_color=COLORS["text"],
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            wrap="word",
        )
        self.log_box.grid(row=0, column=0, sticky="nsew")
        set_textbox(self.log_box, "")
        return panel

    def _tool_state(self, key: str, default: str = "CHECKING") -> str:
        status = self.cli_status.get(key, {})
        if isinstance(status, dict):
            return str(status.get("state") or default)
        return default

    def _dashboard_gpu_name(self) -> str:
        detail = str(self.gpu_status.get("detail") or "")
        name = str(self.gpu_status.get("name") or detail.split(" / ", 1)[0] or self.gpu_status.get("state") or "--")
        for prefix in ("NVIDIA GeForce ", "NVIDIA "):
            if name.startswith(prefix):
                name = name[len(prefix) :]
        return name or "--"

    def _dashboard_memory_count(self) -> int:
        path = memory_index_path()
        if not path.exists():
            return 0
        try:
            data = json.loads(path.read_text(encoding="utf-8", errors="replace"))
        except Exception:
            return 0
        return len(data) if isinstance(data, list) else 0

    def _dashboard_package_dir(self) -> Path | None:
        if self.package_output_dir and self.package_output_dir.exists():
            return self.package_output_dir
        return None

    def _dashboard_file_ready(self, package_dir: Path | None, relative_path: str, cached_path: Path | None = None) -> bool:
        if cached_path is not None and cached_path.exists():
            return True
        if package_dir is None:
            return False
        return (package_dir / relative_path).exists()

    def _dashboard_memory_saved(self, package_dir: Path | None) -> bool:
        if package_dir is None:
            return False
        path = memory_index_path()
        if not path.exists():
            return False
        try:
            index = json.loads(path.read_text(encoding="utf-8", errors="replace"))
        except Exception:
            return False
        if not isinstance(index, list):
            return False
        try:
            target = str(package_dir.resolve(strict=False)).lower()
        except Exception:
            target = str(package_dir).lower()
        for item in index:
            if not isinstance(item, dict):
                continue
            source = str(item.get("source_package") or "")
            try:
                source = str(Path(source).resolve(strict=False))
            except Exception:
                pass
            if source.lower() == target:
                return True
        return False

    def _dashboard_next_action(self) -> str:
        package_dir = self._dashboard_package_dir()
        if not self.system_check_completed:
            return UI_TEXT["run_system_check"]
        if self.selected_file is None and self.test_video_path is None:
            return UI_TEXT["select_test_video"]
        if package_dir is None:
            return UI_TEXT["generate_posting_package"]
        if not self._dashboard_file_ready(package_dir, "assistant_review.md", self.review_file_path):
            return UI_TEXT["run_assistant_review"]
        if not self._dashboard_file_ready(package_dir, "selected/selected_summary.md", None):
            return UI_TEXT["export_selected_draft"]
        if not self._dashboard_file_ready(package_dir, "selected/short_vertical_1080x1920.mp4", self.vertical_short_path):
            return UI_TEXT["generate_vertical_short"]
        if not self._dashboard_memory_saved(package_dir):
            return UI_TEXT["save_to_memory"]
        if not self._dashboard_file_ready(package_dir, "assistant_recommendation.md", self.recommendation_path):
            return UI_TEXT["generate_recommendation"]
        return UI_TEXT["all_set"]

    def _update_dashboard(self) -> None:
        if not hasattr(self, "status_strip_var"):
            return
        package_dir = self._dashboard_package_dir()
        package_state = UI_TEXT["dashboard_package_ready"] if package_dir else UI_TEXT["dashboard_waiting"]
        bridge_state = UI_TEXT["dashboard_connected"] if self.project_bridge_data else UI_TEXT["dashboard_waiting"]
        memory_count = self._dashboard_memory_count()
        status_parts = [
            f"{UI_TEXT['dashboard_gpu']}: {self._dashboard_gpu_name()}",
            f"{UI_TEXT['dashboard_ollama']}: {self._tool_state('ollama')}",
            f"{UI_TEXT['dashboard_ffmpeg']}: {self._tool_state('ffmpeg')}",
            f"{UI_TEXT['dashboard_package']}: {package_state}",
            f"{UI_TEXT['dashboard_memory']}: {memory_count} {UI_TEXT['dashboard_entries']}",
            f"{UI_TEXT['dashboard_bridge']}: {bridge_state}",
        ]
        self.status_strip_var.set("  |  ".join(status_parts))
        self.next_action_var.set(self._dashboard_next_action())

    def _select_video(self) -> None:
        file_path = filedialog.askopenfilename(
            title=UI_TEXT["select_video_dialog"],
            filetypes=[(UI_TEXT["file_types"], "*.mp4 *.mov *.mkv *.webm"), ("All files", "*.*")],
        )
        if not file_path:
            return
        path = Path(file_path)
        if path.suffix.lower() not in VIDEO_EXTENSIONS:
            messagebox.showwarning(APP_NAME, UI_TEXT["supported_files"])
            return
        self.selected_file = path
        self.file_var.set(str(path))
        self.output_var.set(UI_TEXT["no_output"])
        self.current_project = None
        self.current_media_info = None
        self.package_output_dir = None
        self.review_file_path = None
        self.open_button.configure(state="disabled")
        self.open_package_button.configure(state="disabled")
        self.open_review_button.configure(state="disabled")
        self.open_selected_button.configure(state="disabled")
        self.open_short_preview_button.configure(state="disabled")
        self.open_vertical_short_button.configure(state="disabled")
        self.open_recommendation_button.configure(state="disabled")
        self.selected_output_dir = None
        self.short_preview_path = None
        self.vertical_short_path = None
        self.horizontal_edit_path = None
        self.recommendation_path = None
        self.preview_source_video_path = None
        self.sequence_items = []
        self.candidate_data = {}
        self.short_choice_touched = False
        self.title_choice_touched = False
        self.short_choice_var.set(UI_TEXT["selected_none"])
        self.title_choice_var.set(UI_TEXT["selected_none"])
        self.package_summary_var.set(UI_TEXT["package_ready_hint"])
        self.review_summary_var.set(UI_TEXT["review_ready_hint"])
        self.selected_summary_var.set(UI_TEXT["selected_ready_hint"])
        self.short_preview_summary_var.set(UI_TEXT["short_preview_ready_hint"])
        self.vertical_short_summary_var.set(UI_TEXT["vertical_short_ready_hint"])
        self.sequence_summary_var.set(UI_TEXT["horizontal_edit_ready_hint"])
        self.sequence_choice_var.set(UI_TEXT["sequence_empty"])
        self.sequence_choice_menu.configure(values=[UI_TEXT["sequence_empty"]])
        self.open_horizontal_edit_button.configure(state="disabled")
        self.recommendation_summary_var.set(UI_TEXT["recommend_ready_hint"])
        self._log(LOG_TEXT["source_detected"])
        self._update_dashboard()
        self._probe_selected_video()

    def _select_test_video(self) -> None:
        file_path = filedialog.askopenfilename(
            title=UI_TEXT["select_test_video"],
            filetypes=[(UI_TEXT["file_types"], "*.mp4 *.mov *.mkv *.webm"), ("All files", "*.*")],
        )
        if not file_path:
            return
        path = Path(file_path)
        if path.suffix.lower() not in VIDEO_EXTENSIONS:
            messagebox.showwarning(APP_NAME, UI_TEXT["supported_files"])
            return
        self.test_video_path = path
        self.test_file_var.set(f"{UI_TEXT['test_selected']}: {path.name}")
        self.first_test_summary_var.set(UI_TEXT["first_test_ready_hint"])
        self._log(LOG_TEXT["source_detected"])
        self._update_dashboard()

    def _select_package_folder(self) -> None:
        initial_dir = packages_dir()
        initial_dir.mkdir(parents=True, exist_ok=True)
        folder = filedialog.askdirectory(
            title=UI_TEXT["select_package_folder"],
            initialdir=str(initial_dir),
        )
        if not folder:
            return
        package_dir = Path(folder)
        try:
            package_dir.resolve().relative_to(packages_dir().resolve())
        except ValueError:
            messagebox.showinfo(APP_NAME, UI_TEXT["review_requires_package"])
            return
        self.package_output_dir = package_dir
        self.output_var.set(str(package_dir))
        self.current_project = ProjectPaths.from_root(package_dir)
        self.open_package_button.configure(state="normal")
        self.open_button.configure(state="normal")
        review_path = package_dir / "assistant_review.md"
        self.review_file_path = review_path if review_path.exists() else None
        self.open_review_button.configure(state="normal" if self.review_file_path else "disabled")
        recommendation_path = package_dir / "assistant_recommendation.md"
        self.recommendation_path = recommendation_path if recommendation_path.exists() else None
        self.open_recommendation_button.configure(state="normal" if self.recommendation_path else "disabled")
        self.review_summary_var.set(
            "\n".join(
                [
                    UI_TEXT["assistant_review"],
                    f"{UI_TEXT['package']}: {package_dir}",
                    f"{UI_TEXT['review']}: {review_path if review_path.exists() else '--'}",
                ]
            )
        )
        self._refresh_selected_candidates()
        self._load_sequence_for_package(package_dir)
        self._update_dashboard()

    def _start_first_video_test(self) -> None:
        if self.first_test_running:
            return
        if self.test_video_path is None:
            messagebox.showinfo(APP_NAME, UI_TEXT["first_test_requires_video"])
            return
        self.first_test_running = True
        self.first_test_button.configure(state="disabled")
        self.open_test_output_button.configure(state="disabled")
        self.first_test_summary_var.set(UI_TEXT["checking"])
        self.status_var.set(UI_TEXT["working"])
        self.progress_var.set(0.05)
        self._set_eta(30)

        def worker() -> None:
            try:
                system = run_system_check()
                self.events.put(
                    {
                        "type": "cli",
                        "statuses": system["cli"],
                        "nvenc": system["nvenc"],
                        "gpu": system["gpu"],
                        "install_guide": system["install_guide"],
                        "install_commands": system["install_commands"],
                    }
                )
                statuses = system["cli"]
                result = run_first_video_test(
                    video_path=self.test_video_path,  # type: ignore[arg-type]
                    ffprobe_path=statuses.get("ffprobe", {}).get("path"),
                    ffmpeg_path=statuses.get("ffmpeg", {}).get("path"),
                    nvenc_online=system["nvenc"].get("state") == "ONLINE",
                    log=lambda message: self.events.put({"type": "log", "message": message}),
                )
                self.events.put({"type": "first_test_result", "result": result})
            except Exception as exc:
                self.events.put({"type": "first_test_error", "message": str(exc)})

        threading.Thread(target=worker, daemon=True).start()

    def _start_posting_package(self) -> None:
        if self.package_running:
            return
        if self.selected_file is None:
            messagebox.showinfo(APP_NAME, UI_TEXT["package_requires_video"])
            return

        self.package_running = True
        self._set_package_buttons("disabled")
        self.open_package_button.configure(state="disabled")
        self.status_var.set(UI_TEXT["running"])
        self.package_summary_var.set(
            "\n".join(
                [
                    UI_TEXT["posting_package"],
                    UI_TEXT["package_running"],
                ]
            )
        )
        self.progress_var.set(0.03)
        self.step_vars[UI_TEXT["export_package"]].set(UI_TEXT["working"])
        if self.current_media_info and self.current_media_info.duration > 0:
            self._set_eta(estimate_processing_seconds(self.current_media_info.duration, is_faster_whisper_available()))
        else:
            self.eta_var.set(UI_TEXT["eta_calculating"])
            self.finish_var.set(UI_TEXT["finish_empty"])

        video_path = self.selected_file

        def worker() -> None:
            try:
                system = run_system_check()
                self.events.put(
                    {
                        "type": "cli",
                        "statuses": system["cli"],
                        "nvenc": system["nvenc"],
                        "gpu": system["gpu"],
                        "install_guide": system["install_guide"],
                        "install_commands": system["install_commands"],
                    }
                )
                statuses = system["cli"]
                result = generate_posting_package(
                    video_path=video_path,
                    ffprobe_path=statuses.get("ffprobe", {}).get("path"),
                    ollama_ready=statuses.get("ollama", {}).get("state") == "READY",
                    log=lambda message: self.events.put({"type": "log", "message": message}),
                    progress=lambda value: self.events.put({"type": "progress", "value": value}),
                )
                self.events.put({"type": "posting_package_result", "result": result})
            except Exception as exc:
                self.events.put({"type": "posting_package_error", "message": str(exc)})

        threading.Thread(target=worker, daemon=True).start()

    def _resolve_package_for_review(self) -> Path | None:
        if self.package_output_dir and self.package_output_dir.exists():
            return self.package_output_dir
        latest = find_latest_package_dir()
        if latest is not None:
            self.package_output_dir = latest
            return latest
        return None

    def _choice_index(self, value: str) -> int | None:
        value = value.strip()
        if not value.startswith("#"):
            return None
        number = value.split(" ", 1)[0].lstrip("#")
        try:
            return max(0, int(number) - 1)
        except ValueError:
            return None

    def _on_short_choice(self, _value: str) -> None:
        self.short_choice_touched = True

    def _on_title_choice(self, _value: str) -> None:
        self.title_choice_touched = True

    def _compact_options(self, labels: list[str]) -> list[str]:
        compact: list[str] = []
        for label in labels:
            compact.append(label if len(label) <= 78 else label[:75].rstrip() + "...")
        return compact

    def _refresh_selected_candidates(self) -> None:
        package_dir = self._resolve_package_for_review()
        if package_dir is None:
            messagebox.showinfo(APP_NAME, UI_TEXT["selected_requires_package"])
            return
        try:
            data = read_selected_candidates(package_dir)
        except Exception as exc:
            self.selected_summary_var.set(str(exc))
            return

        self.candidate_data = data
        short_labels = self._compact_options([str(item) for item in data.get("short_labels", []) if str(item)])
        title_labels = self._compact_options([str(item) for item in data.get("title_labels", []) if str(item)])
        if not short_labels:
            short_labels = [UI_TEXT["selected_none"]]
        if not title_labels:
            title_labels = [UI_TEXT["selected_none"]]
        self.short_choice_menu.configure(values=short_labels)
        self.title_choice_menu.configure(values=title_labels)
        self.short_choice_var.set(short_labels[0])
        self.title_choice_var.set(title_labels[0])
        self.short_choice_touched = False
        self.title_choice_touched = False

        selected_dir = package_dir / "selected"
        self.selected_output_dir = selected_dir if selected_dir.exists() else None
        self.open_selected_button.configure(state="normal" if self.selected_output_dir else "disabled")
        self.selected_summary_var.set(self._format_selected_candidates_summary(data))
        self._log(LOG_TEXT["selected_candidates_refreshed"])
        self._update_dashboard()

    def _sequence_labels(self) -> list[str]:
        labels: list[str] = []
        for index, item in enumerate(self.sequence_items, start=1):
            name = Path(str(item.get("path") or "")).name or "--"
            labels.append(f"#{index} {name}")
        return labels

    def _sequence_choice_index(self) -> int | None:
        value = self.sequence_choice_var.get().strip()
        if not value.startswith("#"):
            return None
        try:
            return int(value.split(" ", 1)[0].lstrip("#")) - 1
        except ValueError:
            return None

    def _format_sequence_summary(self, result: dict[str, object] | None = None) -> str:
        labels = self._sequence_labels()
        lines = [UI_TEXT["sequence"]]
        if labels:
            lines.extend(f"[{index}] {Path(str(item.get('path') or '')).name}" for index, item in enumerate(self.sequence_items, start=1))
        else:
            lines.append(UI_TEXT["sequence_empty"])
        lines.extend(["", f"{UI_TEXT['sequence_duration']}: {format_duration(sequence_total_duration(self.sequence_items))}"])
        if result:
            lines.extend(
                [
                    "",
                    UI_TEXT["horizontal_edit"],
                    f"{UI_TEXT['package_status']}: {result.get('status', UI_TEXT['horizontal_edit_failed'])}",
                    f"{UI_TEXT['encoder']}: {result.get('encoder', 'unavailable')}",
                    f"{UI_TEXT['preview_output']}: {result.get('output_path', '--')}",
                ]
            )
            recommendation = str(result.get("recommendation") or "")
            if recommendation:
                lines.extend(["", f"{UI_TEXT['sequence_recommendation']}: {recommendation}"])
        return "\n".join(lines)

    def _refresh_sequence_ui(self, result: dict[str, object] | None = None) -> None:
        labels = self._sequence_labels()
        values = labels if labels else [UI_TEXT["sequence_empty"]]
        self.sequence_choice_menu.configure(values=values)
        current = self.sequence_choice_var.get()
        self.sequence_choice_var.set(current if current in values else values[0])
        self.sequence_summary_var.set(self._format_sequence_summary(result))

    def _load_sequence_for_package(self, package_dir: Path | None = None) -> None:
        package = package_dir or self._dashboard_package_dir()
        if package is None:
            self.sequence_items = []
            self.horizontal_edit_path = None
            self.open_horizontal_edit_button.configure(state="disabled")
            self._refresh_sequence_ui()
            return
        self.sequence_items = read_sequence(package)
        candidate = package / "selected" / "horizontal_edit.mp4"
        self.horizontal_edit_path = candidate if candidate.exists() else None
        self.open_horizontal_edit_button.configure(state="normal" if self.horizontal_edit_path else "disabled")
        self._refresh_sequence_ui()

    def _save_sequence_for_package(self, package_dir: Path | None = None) -> Path | None:
        package = package_dir or self._resolve_package_for_review()
        if package is None:
            messagebox.showinfo(APP_NAME, UI_TEXT["sequence_requires_package"])
            return None
        path = write_sequence(package, self.sequence_items)
        self.selected_output_dir = path.parent
        self.open_selected_button.configure(state="normal")
        self._refresh_sequence_ui()
        self._update_dashboard()
        return path

    def _add_sequence_videos(self) -> None:
        package_dir = self._resolve_package_for_review()
        if package_dir is None:
            messagebox.showinfo(APP_NAME, UI_TEXT["sequence_requires_package"])
            return
        file_paths = filedialog.askopenfilenames(
            title=UI_TEXT["add_sequence_video"],
            filetypes=[(UI_TEXT["file_types"], "*.mp4 *.mov *.mkv *.webm"), ("All files", "*.*")],
        )
        if not file_paths:
            return
        new_indexes: list[int] = []
        for file_path in file_paths:
            path = Path(file_path)
            if path.suffix.lower() not in VIDEO_EXTENSIONS:
                continue
            self.sequence_items.append({"path": str(path), "duration": 0.0, "audio_present": True})
            new_indexes.append(len(self.sequence_items) - 1)
        if not new_indexes:
            messagebox.showwarning(APP_NAME, UI_TEXT["supported_files"])
            return
        self._save_sequence_for_package(package_dir)
        self._log(UI_TEXT["sequence_log_arranging"])
        self._probe_sequence_items(package_dir, new_indexes)

    def _probe_sequence_items(self, package_dir: Path, indexes: list[int]) -> None:
        sequence = [dict(item) for item in self.sequence_items]

        def worker() -> None:
            try:
                system = run_system_check()
                self.events.put(
                    {
                        "type": "cli",
                        "statuses": system["cli"],
                        "nvenc": system["nvenc"],
                        "gpu": system["gpu"],
                        "install_guide": system["install_guide"],
                        "install_commands": system["install_commands"],
                    }
                )
                ffprobe_path = system["cli"].get("ffprobe", {}).get("path")
                if not ffprobe_path:
                    return
                for index in indexes:
                    if index >= len(sequence):
                        continue
                    path = Path(str(sequence[index].get("path") or ""))
                    try:
                        info = probe_media(path, ffprobe_path)
                    except Exception:
                        continue
                    sequence[index]["duration"] = info.duration
                    sequence[index]["audio_present"] = info.audio_present
                self.events.put({"type": "sequence_probe_result", "package_dir": str(package_dir), "sequence": sequence})
            except Exception as exc:
                self.events.put({"type": "log", "message": str(exc)})

        threading.Thread(target=worker, daemon=True).start()

    def _remove_selected_sequence_video(self) -> None:
        index = self._sequence_choice_index()
        if index is None or not (0 <= index < len(self.sequence_items)):
            messagebox.showinfo(APP_NAME, UI_TEXT["sequence_no_selection"])
            return
        del self.sequence_items[index]
        self._save_sequence_for_package()

    def _move_sequence_up(self) -> None:
        index = self._sequence_choice_index()
        if index is None or index <= 0 or index >= len(self.sequence_items):
            return
        self.sequence_items[index - 1], self.sequence_items[index] = self.sequence_items[index], self.sequence_items[index - 1]
        self._save_sequence_for_package()
        labels = self._sequence_labels()
        if index - 1 < len(labels):
            self.sequence_choice_var.set(labels[index - 1])

    def _move_sequence_down(self) -> None:
        index = self._sequence_choice_index()
        if index is None or index < 0 or index >= len(self.sequence_items) - 1:
            return
        self.sequence_items[index + 1], self.sequence_items[index] = self.sequence_items[index], self.sequence_items[index + 1]
        self._save_sequence_for_package()
        labels = self._sequence_labels()
        if index + 1 < len(labels):
            self.sequence_choice_var.set(labels[index + 1])

    def _start_generate_horizontal_edit(self) -> None:
        if self.sequence_running:
            return
        package_dir = self._resolve_package_for_review()
        if package_dir is None:
            messagebox.showinfo(APP_NAME, UI_TEXT["sequence_requires_package"])
            return
        if not self.sequence_items:
            self._load_sequence_for_package(package_dir)
        if not self.sequence_items:
            messagebox.showinfo(APP_NAME, UI_TEXT["sequence_empty"])
            return

        self.sequence_running = True
        self.generate_horizontal_edit_button.configure(state="disabled")
        self.open_horizontal_edit_button.configure(state="disabled")
        self.status_var.set(UI_TEXT["running"])
        self.progress_var.set(0.12)
        self._set_eta(max(45, sequence_total_duration(self.sequence_items) * 0.35 + 45))
        self.sequence_summary_var.set(
            "\n".join(
                [
                    UI_TEXT["sequence_builder"],
                    UI_TEXT["horizontal_edit_running"],
                    f"{UI_TEXT['sequence_videos']}: {len(self.sequence_items)}",
                    f"{UI_TEXT['sequence_duration']}: {format_duration(sequence_total_duration(self.sequence_items))}",
                ]
            )
        )
        sequence = [dict(item) for item in self.sequence_items]

        def worker() -> None:
            try:
                system = run_system_check()
                self.events.put(
                    {
                        "type": "cli",
                        "statuses": system["cli"],
                        "nvenc": system["nvenc"],
                        "gpu": system["gpu"],
                        "install_guide": system["install_guide"],
                        "install_commands": system["install_commands"],
                    }
                )
                statuses = system["cli"]
                result = generate_horizontal_edit(
                    package_dir=package_dir,
                    sequence=sequence,
                    ffmpeg_path=statuses.get("ffmpeg", {}).get("path"),
                    nvenc_online=system["nvenc"].get("state") == "ONLINE",
                    ollama_ready=statuses.get("ollama", {}).get("state") == "READY",
                    log=lambda message: self.events.put({"type": "log", "message": message}),
                )
                self.events.put({"type": "horizontal_edit_result", "result": result})
            except Exception as exc:
                self.events.put({"type": "horizontal_edit_error", "message": str(exc)})

        threading.Thread(target=worker, daemon=True).start()

    def _project_bridge_root_for_name(self, project_name: str) -> Path | None:
        for box in self.project_bridge_boxes:
            if box.get("name") == project_name:
                root = Path(str(box.get("root") or ""))
                return root if root.exists() else None
        return None

    def _refresh_project_bridge(self, silent: bool = False) -> None:
        try:
            self.project_bridge_boxes = list_project_boxes()
        except Exception as exc:
            self.project_bridge_boxes = []
            self.project_bridge_data = {}
            self.project_bridge_summary_var.set(str(exc))
            return

        if not self.project_bridge_boxes:
            self.project_bridge_data = {}
            self.project_choice_menu.configure(values=[UI_TEXT["no_project_boxes"]])
            self.project_bgm_menu.configure(values=[UI_TEXT["bgm_none"]])
            self.project_choice_var.set(UI_TEXT["no_project_boxes"])
            self.project_bgm_choice_var.set(UI_TEXT["bgm_none"])
            self.project_bridge_summary_var.set(UI_TEXT["no_project_boxes"])
            if not silent:
                self._log(LOG_TEXT["project_bridge_missing"])
            self._update_dashboard()
            return

        values = [str(box["name"]) for box in self.project_bridge_boxes]
        self.project_choice_menu.configure(values=values)
        current = self.project_choice_var.get()
        chosen = current if current in values else values[0]
        self.project_choice_var.set(chosen)
        self._load_project_bridge(chosen, silent=silent)
        self._update_dashboard()

    def _on_project_bridge_choice(self, value: str) -> None:
        self._load_project_bridge(value)

    def _load_project_bridge(self, project_name: str, silent: bool = False) -> None:
        if project_name == UI_TEXT["no_project_boxes"]:
            self.project_bridge_data = {}
            self.project_bridge_summary_var.set(UI_TEXT["no_project_boxes"])
            self._update_dashboard()
            return
        project_root = self._project_bridge_root_for_name(project_name)
        if project_root is None:
            self.project_bridge_data = {}
            self.project_bridge_summary_var.set(UI_TEXT["project_bridge_requires_project"])
            self._update_dashboard()
            return
        try:
            data = read_project_box(project_root)
        except Exception as exc:
            self.project_bridge_data = {}
            self.project_bridge_summary_var.set(str(exc))
            self._update_dashboard()
            return

        self.project_bridge_data = data
        bgm_names = [str(item.get("name")) for item in data.get("bgm_files", []) if isinstance(item, dict) and item.get("name")]
        if not bgm_names:
            bgm_names = [UI_TEXT["bgm_none"]]
        self.project_bgm_menu.configure(values=bgm_names)
        selected_bgm = str(data.get("selected_bgm") or "")
        self.project_bgm_choice_var.set(selected_bgm if selected_bgm in bgm_names else bgm_names[0])
        self.project_bridge_summary_var.set(self._format_project_bridge_summary(data))
        if not silent:
            self._log(LOG_TEXT["project_box_loaded"])
        self._update_dashboard()

    def _selected_project_bgm_path(self) -> Path | None:
        if not self.project_bridge_data:
            self._load_project_bridge(self.project_choice_var.get(), silent=True)
        selected = self.project_bgm_choice_var.get()
        if selected == UI_TEXT["bgm_none"]:
            return None
        for item in self.project_bridge_data.get("bgm_files", []):
            if isinstance(item, dict) and item.get("name") == selected:
                path = Path(str(item.get("path") or ""))
                return path if path.exists() else None
        return None

    def _start_project_bgm_preview(self) -> None:
        bgm_path = self._selected_project_bgm_path()
        if bgm_path is None:
            messagebox.showinfo(APP_NAME, UI_TEXT["project_bridge_requires_bgm"])
            return
        self._log(LOG_TEXT["project_bgm_bridge"])
        result = self.audio_preview.play(bgm_path)
        if result.success:
            self._log(LOG_TEXT["project_preview_started"])
            self.project_bridge_summary_var.set(
                self._format_project_bridge_summary(self.project_bridge_data, extra=f"Preview: {result.backend}")
            )
        else:
            self.project_bridge_summary_var.set(f"{UI_TEXT['preview_failed']}\n{result.message}")

    def _stop_project_bgm_preview(self) -> None:
        result = self.audio_preview.stop()
        if result.success:
            self._log(LOG_TEXT["project_preview_stopped"])
            if self.project_bridge_data:
                self.project_bridge_summary_var.set(
                    self._format_project_bridge_summary(self.project_bridge_data, extra=UI_TEXT["preview_stopped"])
                )
        else:
            self.project_bridge_summary_var.set(f"{UI_TEXT['preview_failed']}\n{result.message}")

    def _start_add_project_bgm(self) -> None:
        if self.project_bridge_running:
            return
        package_dir = self._resolve_package_for_review()
        if package_dir is None:
            messagebox.showinfo(APP_NAME, UI_TEXT["project_bridge_requires_package"])
            return
        bgm_path = self._selected_project_bgm_path()
        if bgm_path is None:
            messagebox.showinfo(APP_NAME, UI_TEXT["project_bridge_requires_bgm"])
            return

        self.project_bridge_running = True
        self.add_project_bgm_button.configure(state="disabled")
        self.status_var.set(UI_TEXT["running"])
        self.progress_var.set(0.18)

        def worker() -> None:
            try:
                self.events.put({"type": "log", "message": LOG_TEXT["project_bgm_bridge"]})
                result = add_bgm_to_video_box(
                    package_dir=package_dir,
                    bgm_path=bgm_path,
                    log=lambda _message: self.events.put({"type": "log", "message": LOG_TEXT["project_bgm_added"]}),
                )
                self.events.put({"type": "project_bridge_copy_result", "result": result})
            except Exception as exc:
                self.events.put({"type": "project_bridge_error", "message": str(exc)})

        threading.Thread(target=worker, daemon=True).start()

    def _start_generate_bridge_metadata(self) -> None:
        if self.project_bridge_running:
            return
        package_dir = self._resolve_package_for_review()
        if package_dir is None:
            messagebox.showinfo(APP_NAME, UI_TEXT["project_bridge_requires_package"])
            return
        if not self.project_bridge_data:
            self._load_project_bridge(self.project_choice_var.get(), silent=True)
        if not self.project_bridge_data:
            messagebox.showinfo(APP_NAME, UI_TEXT["project_bridge_requires_project"])
            return
        bgm_path = self._selected_project_bgm_path()

        self.project_bridge_running = True
        self.generate_bridge_metadata_button.configure(state="disabled")
        self.status_var.set(UI_TEXT["running"])
        self.progress_var.set(0.20)
        self._set_eta(30)

        project = dict(self.project_bridge_data)

        def worker() -> None:
            try:
                system = run_system_check()
                self.events.put(
                    {
                        "type": "cli",
                        "statuses": system["cli"],
                        "nvenc": system["nvenc"],
                        "gpu": system["gpu"],
                        "install_guide": system["install_guide"],
                        "install_commands": system["install_commands"],
                    }
                )
                statuses = system["cli"]
                result = generate_bridge_metadata_draft(
                    package_dir=package_dir,
                    project=project,
                    bgm_path=bgm_path,
                    ollama_ready=statuses.get("ollama", {}).get("state") == "READY",
                )
                self.events.put({"type": "project_bridge_metadata_result", "result": result})
            except Exception as exc:
                self.events.put({"type": "project_bridge_error", "message": str(exc)})

        threading.Thread(target=worker, daemon=True).start()

    def _set_memory_buttons(self, state: str) -> None:
        self.save_memory_button.configure(state=state)
        self.generate_memory_summary_button.configure(state=state)

    def _set_package_buttons(self, state: str) -> None:
        for attr in ("generate_package_button", "generate_package_video_button"):
            if hasattr(self, attr):
                getattr(self, attr).configure(state=state)

    def _set_recommendation_buttons(self, state: str) -> None:
        self.generate_recommendation_button.configure(state=state)
        self.refresh_memory_button.configure(state=state)

    def _refresh_recommend_memory(self) -> None:
        try:
            analysis = analyze_memory()
        except Exception as exc:
            self.recommendation_summary_var.set(str(exc))
            return
        self.recommendation_summary_var.set(self._format_recommend_memory_summary(analysis))
        self._log(LOG_TEXT["recommend_memory_loading"])
        self._update_dashboard()

    def _start_save_memory(self) -> None:
        if self.memory_running:
            return
        package_dir = self._resolve_package_for_review()
        if package_dir is None:
            messagebox.showinfo(APP_NAME, UI_TEXT["memory_requires_package"])
            return

        self.memory_running = True
        self._set_memory_buttons("disabled")
        self.status_var.set(UI_TEXT["running"])
        self.progress_var.set(0.18)
        self._set_eta(30)
        self.memory_summary_var.set(
            "\n".join(
                [
                    UI_TEXT["memory"],
                    f"{UI_TEXT['memory_status']}: {UI_TEXT['running']}",
                    f"{UI_TEXT['package']}: {package_dir}",
                ]
            )
        )

        def worker() -> None:
            try:
                self.events.put({"type": "log", "message": LOG_TEXT["memory_organizing"]})
                system = run_system_check()
                self.events.put(
                    {
                        "type": "cli",
                        "statuses": system["cli"],
                        "nvenc": system["nvenc"],
                        "gpu": system["gpu"],
                        "install_guide": system["install_guide"],
                        "install_commands": system["install_commands"],
                    }
                )
                statuses = system["cli"]
                result = save_package_to_memory(
                    package_dir=package_dir,
                    ollama_ready=statuses.get("ollama", {}).get("state") == "READY",
                )
                self.events.put({"type": "memory_save_result", "result": result})
            except Exception as exc:
                self.events.put({"type": "memory_error", "message": str(exc)})

        threading.Thread(target=worker, daemon=True).start()

    def _start_generate_memory_summary(self) -> None:
        if self.memory_running:
            return
        self.memory_running = True
        self._set_memory_buttons("disabled")
        self.status_var.set(UI_TEXT["running"])
        self.progress_var.set(0.20)
        self._set_eta(25)
        self.memory_summary_var.set(
            "\n".join([UI_TEXT["memory"], f"{UI_TEXT['memory_status']}: {UI_TEXT['running']}"])
        )

        def worker() -> None:
            try:
                self.events.put({"type": "log", "message": LOG_TEXT["memory_organizing"]})
                system = run_system_check()
                self.events.put(
                    {
                        "type": "cli",
                        "statuses": system["cli"],
                        "nvenc": system["nvenc"],
                        "gpu": system["gpu"],
                        "install_guide": system["install_guide"],
                        "install_commands": system["install_commands"],
                    }
                )
                statuses = system["cli"]
                result = generate_memory_summary(
                    ollama_ready=statuses.get("ollama", {}).get("state") == "READY",
                )
                self.events.put({"type": "memory_summary_result", "result": result})
            except Exception as exc:
                self.events.put({"type": "memory_error", "message": str(exc)})

        threading.Thread(target=worker, daemon=True).start()

    def _start_generate_recommendation(self) -> None:
        if self.recommendation_running:
            return
        package_dir = self._resolve_package_for_review()
        if package_dir is None:
            messagebox.showinfo(APP_NAME, UI_TEXT["recommend_requires_package"])
            return

        self.recommendation_running = True
        self._set_recommendation_buttons("disabled")
        self.open_recommendation_button.configure(state="disabled")
        self.status_var.set(UI_TEXT["running"])
        self.progress_var.set(0.18)
        self._set_eta(30)
        self.recommendation_summary_var.set(
            "\n".join(
                [
                    UI_TEXT["assistant_recommend"],
                    f"{UI_TEXT['memory_status']}: {UI_TEXT['running']}",
                    f"{UI_TEXT['package']}: {package_dir}",
                ]
            )
        )

        def worker() -> None:
            try:
                self.events.put({"type": "log", "message": LOG_TEXT["recommend_memory_loading"]})
                system = run_system_check()
                self.events.put(
                    {
                        "type": "cli",
                        "statuses": system["cli"],
                        "nvenc": system["nvenc"],
                        "gpu": system["gpu"],
                        "install_guide": system["install_guide"],
                        "install_commands": system["install_commands"],
                    }
                )
                statuses = system["cli"]
                result = generate_assistant_recommendation(
                    package_dir=package_dir,
                    ollama_ready=statuses.get("ollama", {}).get("state") == "READY",
                )
                self.events.put({"type": "recommendation_result", "result": result})
            except Exception as exc:
                self.events.put({"type": "recommendation_error", "message": str(exc)})

        threading.Thread(target=worker, daemon=True).start()

    def _start_export_selected_draft(self) -> None:
        if self.selected_export_running:
            return
        package_dir = self._resolve_package_for_review()
        if package_dir is None:
            messagebox.showinfo(APP_NAME, UI_TEXT["selected_requires_package"])
            return
        self.selected_export_running = True
        self.export_selected_button.configure(state="disabled")
        self.status_var.set(UI_TEXT["running"])
        self.progress_var.set(0.16)

        parsed_short_index = self._choice_index(self.short_choice_var.get())
        parsed_title_index = self._choice_index(self.title_choice_var.get())
        short_index = parsed_short_index if self.short_choice_touched or (parsed_short_index not in {None, 0}) else None
        title_index = parsed_title_index if self.title_choice_touched or (parsed_title_index not in {None, 0}) else None

        def worker() -> None:
            try:
                result = export_selected_draft(
                    package_dir=package_dir,
                    short_index=short_index,
                    title_index=title_index,
                    log=lambda message: self.events.put({"type": "log", "message": message}),
                )
                self.events.put({"type": "selected_export_result", "result": result})
            except Exception as exc:
                self.events.put({"type": "selected_export_error", "message": str(exc)})

        threading.Thread(target=worker, daemon=True).start()

    def _resolve_preview_source_video(self, package_dir: Path) -> Path | None:
        source = self.preview_source_video_path
        if source is not None and source.exists():
            return source
        source = find_source_video_path(package_dir)
        if source is not None:
            self.preview_source_video_path = source
            return source
        if self.selected_file is not None and self.selected_file.exists():
            self.preview_source_video_path = self.selected_file
            return self.selected_file
        file_path = filedialog.askopenfilename(
            title=UI_TEXT["preview_source_dialog"],
            filetypes=[(UI_TEXT["file_types"], "*.mp4 *.mov *.mkv *.webm"), ("All files", "*.*")],
        )
        if not file_path:
            return None
        path = Path(file_path)
        if path.suffix.lower() not in VIDEO_EXTENSIONS:
            messagebox.showwarning(APP_NAME, UI_TEXT["supported_files"])
            return None
        self.preview_source_video_path = path
        return path

    def _start_generate_short_preview(self) -> None:
        if self.short_preview_running:
            return
        package_dir = self._resolve_package_for_review()
        if package_dir is None:
            messagebox.showinfo(APP_NAME, UI_TEXT["selected_requires_package"])
            return
        source_video = self._resolve_preview_source_video(package_dir)

        self.short_preview_running = True
        self.generate_short_preview_button.configure(state="disabled")
        self.open_short_preview_button.configure(state="disabled")
        self.status_var.set(UI_TEXT["running"])
        self.progress_var.set(0.12)
        self._set_eta(45)
        self.short_preview_summary_var.set(
            "\n".join(
                [
                    UI_TEXT["short_preview"],
                    UI_TEXT["short_preview_running"],
                    f"{UI_TEXT['package']}: {package_dir}",
                ]
            )
        )

        def worker() -> None:
            try:
                system = run_system_check()
                self.events.put(
                    {
                        "type": "cli",
                        "statuses": system["cli"],
                        "nvenc": system["nvenc"],
                        "gpu": system["gpu"],
                        "install_guide": system["install_guide"],
                        "install_commands": system["install_commands"],
                    }
                )
                statuses = system["cli"]
                result = generate_selected_short_preview(
                    package_dir=package_dir,
                    ffmpeg_path=statuses.get("ffmpeg", {}).get("path"),
                    nvenc_online=system["nvenc"].get("state") == "ONLINE",
                    source_video_path=source_video,
                    log=lambda message: self.events.put({"type": "log", "message": message}),
                )
                self.events.put({"type": "short_preview_result", "result": result})
            except Exception as exc:
                self.events.put({"type": "short_preview_error", "message": str(exc)})

        threading.Thread(target=worker, daemon=True).start()

    def _start_generate_vertical_short(self) -> None:
        if self.vertical_short_running:
            return
        package_dir = self._resolve_package_for_review()
        if package_dir is None:
            messagebox.showinfo(APP_NAME, UI_TEXT["selected_requires_package"])
            return
        source_video = self._resolve_preview_source_video(package_dir)

        self.vertical_short_running = True
        self.generate_vertical_short_button.configure(state="disabled")
        self.open_vertical_short_button.configure(state="disabled")
        self.status_var.set(UI_TEXT["running"])
        self.progress_var.set(0.10)
        self._set_eta(90)
        self.vertical_short_summary_var.set(
            "\n".join(
                [
                    UI_TEXT["vertical_short"],
                    UI_TEXT["vertical_short_running"],
                    f"{UI_TEXT['output_size']}: 1080x1920",
                    f"{UI_TEXT['package']}: {package_dir}",
                ]
            )
        )

        def worker() -> None:
            try:
                system = run_system_check()
                self.events.put(
                    {
                        "type": "cli",
                        "statuses": system["cli"],
                        "nvenc": system["nvenc"],
                        "gpu": system["gpu"],
                        "install_guide": system["install_guide"],
                        "install_commands": system["install_commands"],
                    }
                )
                statuses = system["cli"]
                result = generate_vertical_short(
                    package_dir=package_dir,
                    ffmpeg_path=statuses.get("ffmpeg", {}).get("path"),
                    nvenc_online=system["nvenc"].get("state") == "ONLINE",
                    source_video_path=source_video,
                    log=lambda message: self.events.put({"type": "log", "message": message}),
                )
                self.events.put({"type": "vertical_short_result", "result": result})
            except Exception as exc:
                self.events.put({"type": "vertical_short_error", "message": str(exc)})

        threading.Thread(target=worker, daemon=True).start()

    def _start_assistant_review(self) -> None:
        if self.review_running:
            return
        package_dir = self._resolve_package_for_review()
        if package_dir is None:
            messagebox.showinfo(APP_NAME, UI_TEXT["review_requires_package"])
            return

        self.review_running = True
        self.assistant_review_button.configure(state="disabled")
        self.open_review_button.configure(state="disabled")
        self.review_summary_var.set(
            "\n".join(
                [
                    UI_TEXT["assistant_review"],
                    UI_TEXT["review_running"],
                    f"{UI_TEXT['package']}: {package_dir}",
                ]
            )
        )
        self.status_var.set(UI_TEXT["running"])
        self.progress_var.set(0.08)
        self._set_eta(45)

        def worker() -> None:
            try:
                system = run_system_check()
                self.events.put(
                    {
                        "type": "cli",
                        "statuses": system["cli"],
                        "nvenc": system["nvenc"],
                        "gpu": system["gpu"],
                        "install_guide": system["install_guide"],
                        "install_commands": system["install_commands"],
                    }
                )
                statuses = system["cli"]
                result = run_assistant_review(
                    package_dir=package_dir,
                    ollama_ready=statuses.get("ollama", {}).get("state") == "READY",
                    log=lambda message: self.events.put({"type": "log", "message": message}),
                )
                self.events.put({"type": "assistant_review_result", "result": result})
            except Exception as exc:
                self.events.put({"type": "assistant_review_error", "message": str(exc)})

        threading.Thread(target=worker, daemon=True).start()

    def _probe_selected_video(self) -> None:
        if self.selected_file is None:
            return

        def worker() -> None:
            try:
                ffprobe_path = self._tool_path("ffprobe")
                if not ffprobe_path:
                    self.events.put({"type": "media_error", "message": "FFPROBE MISSING"})
                    return
                info = probe_media(self.selected_file, ffprobe_path)
                self.events.put({"type": "media", "info": info})
            except Exception as exc:
                self.events.put({"type": "media_error", "message": str(exc)})

        threading.Thread(target=worker, daemon=True).start()

    def _start_cli_check(self) -> None:
        self._log(LOG_TEXT["system_check_start"])
        self.system_check_completed = False
        self.system_check_status_var.set(UI_TEXT["checking"])
        self._set_system_action_buttons("disabled")
        for pill in self.cli_pills.values():
            pill.set_state("CHECKING")
        self._update_dashboard()

        def worker() -> None:
            try:
                system = run_system_check()
                self.events.put(
                    {
                        "type": "cli",
                        "statuses": system["cli"],
                        "nvenc": system["nvenc"],
                        "gpu": system["gpu"],
                        "install_guide": system["install_guide"],
                        "install_commands": system["install_commands"],
                    }
                )
            except Exception as exc:
                self.events.put({"type": "cli_error", "message": str(exc)})

        threading.Thread(target=worker, daemon=True).start()

    def _fetch_metadata(self) -> None:
        url = self.youtube_var.get().strip()
        if not url:
            self.youtube_status_var.set("Paste a YouTube LIVE URL first.")
            return

        def worker() -> None:
            ytdlp_path = self._tool_path("yt-dlp")
            if not ytdlp_path:
                self.events.put({"type": "youtube", "message": "YT-DLP MISSING"})
                return
            try:
                metadata = fetch_youtube_metadata(url, ytdlp_path)
                title = metadata.get("title") or "Untitled"
                duration = metadata.get("duration")
                if isinstance(duration, (int, float)):
                    duration_text = format_duration(float(duration))
                    message = f"Metadata ready: {title} / {duration_text}"
                else:
                    message = f"Metadata ready: {title}"
                self.events.put({"type": "youtube", "message": message})
            except Exception as exc:
                self.events.put({"type": "youtube", "message": f"Fetch failed: {exc}"})

        self.youtube_status_var.set("Fetching metadata without download...")
        threading.Thread(target=worker, daemon=True).start()

    def _start_processing(self) -> None:
        if self.worker_running:
            return
        if self.selected_file is None:
            messagebox.showinfo(APP_NAME, UI_TEXT["no_file"])
            return

        self.worker_running = True
        self.start_button.configure(state="disabled")
        self.status_var.set(UI_TEXT["running"])
        self.progress_var.set(0.02)
        for label, var in self.step_vars.items():
            var.set(UI_TEXT["phase2"] if label == UI_TEXT["thumbnail_base"] else UI_TEXT["waiting"])

        estimate = estimate_processing_seconds(
            self.current_media_info.duration if self.current_media_info else 0,
            is_faster_whisper_available(),
        )
        self._set_eta(estimate)
        threading.Thread(target=self._process_worker, args=(self.selected_file,), daemon=True).start()

    def _process_worker(self, video_path: Path) -> None:
        project: ProjectPaths | None = None
        log_entries: list[str] = []

        def emit_log(message: str) -> None:
            line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} {message}"
            log_entries.append(line)
            self.events.put({"type": "log", "message": message})

        def step(label: str, state: str) -> None:
            self.events.put({"type": "step", "label": label, "state": state})

        try:
            emit_log(LOG_TEXT["package_prepare"])
            project = create_project(video_path)
            self.events.put({"type": "output", "path": project.root})
            write_source_manifest(project, video_path)
            self.events.put({"type": "progress", "value": 0.08})

            statuses = self.cli_status or check_cli_environment()
            ffmpeg_path = statuses.get("ffmpeg", {}).get("path")
            ffprobe_path = statuses.get("ffprobe", {}).get("path")
            nvenc_online = self.nvenc_status.get("state") == "ONLINE"

            media_info: MediaInfo | None = None
            if ffprobe_path:
                emit_log(LOG_TEXT["media_probe"])
                media_info = probe_media(video_path, ffprobe_path)
                write_media_info(project, media_info)
                self.events.put({"type": "media", "info": media_info})
                self.events.put({"type": "eta", "seconds": estimate_processing_seconds(media_info.duration, is_faster_whisper_available())})
            else:
                write_media_info(project, None, "FFprobe is missing. Media information is unavailable.")
                emit_log(LOG_TEXT["ffprobe_missing"])
            self.events.put({"type": "progress", "value": 0.18})

            step(UI_TEXT["transcription"], UI_TEXT["working"])
            transcript_result = transcribe_media(
                video_path=video_path,
                project_dir=project.root,
                log=emit_log,
                progress=lambda value: self.events.put({"type": "progress", "value": 0.18 + (value * 0.28)}),
            )
            step(UI_TEXT["transcription"], UI_TEXT["done"] if transcript_result.available else UI_TEXT["skipped"])

            step(UI_TEXT["scene_analysis"], UI_TEXT["working"])
            emit_log(LOG_TEXT["quiet_scene_search"])
            self.events.put({"type": "progress", "value": 0.52})

            step(UI_TEXT["shorts_candidates"], UI_TEXT["working"])
            duration = media_info.duration if media_info else 0
            candidates = create_shorts_candidates(duration, transcript_result.srt_path)
            write_shorts_candidates(project, candidates)
            emit_log(LOG_TEXT["shorts_extracted"])
            step(UI_TEXT["shorts_candidates"], UI_TEXT["done"])
            self.events.put({"type": "progress", "value": 0.66})

            preview_created = False
            if ffmpeg_path and candidates:
                first = candidates[0]
                preview_path = project.shorts_dir / "short_01_preview.mp4"
                result = create_preview_clip(
                    video_path=video_path,
                    output_path=preview_path,
                    start_seconds=float(first["start_seconds"]),
                    end_seconds=float(first["end_seconds"]),
                    ffmpeg_path=ffmpeg_path,
                    use_nvenc=nvenc_online,
                )
                preview_created = result.created
                if result.created:
                    emit_log(LOG_TEXT["preview_created"])
                else:
                    write_preview_note(project, result.message)
                    emit_log(f"{LOG_TEXT['preview_skipped']} {result.message}")
            else:
                reason = "FFmpeg is required for preview clip generation." if not ffmpeg_path else "No candidate range was available."
                write_preview_note(project, reason)
                emit_log(f"{LOG_TEXT['preview_skipped']} {reason}")
            self.events.put({"type": "progress", "value": 0.78})

            step(UI_TEXT["export_package"], UI_TEXT["working"])
            metadata = build_metadata_draft(
                project_name=project.root.name,
                source_name=video_path.name,
                media_info=media_info,
                transcript_path=transcript_result.transcript_path,
            )
            write_metadata_files(project, metadata, preview_created)
            emit_log(LOG_TEXT["metadata_ready"])
            step(UI_TEXT["export_package"], UI_TEXT["done"])

            write_log_files(project, log_entries)
            self.events.put({"type": "progress", "value": 1.0})
            emit_log(LOG_TEXT["complete"])
            emit_log(LOG_TEXT["running"])
            self.events.put({"type": "complete", "path": project.root})
        except Exception as exc:
            emit_log(LOG_TEXT["process_stopped"])
            emit_log(str(exc))
            if project is not None:
                write_log_files(project, log_entries + [traceback.format_exc()])
            self.events.put({"type": "error", "message": str(exc)})

    def _tool_path(self, key: str) -> str | None:
        status = self.cli_status.get(key)
        if status:
            return status.get("path")
        return None

    def _set_eta(self, seconds: float) -> None:
        finish = datetime.now() + timedelta(seconds=max(0, seconds))
        self.eta_var.set(f"ETA {format_duration(seconds)}")
        self.finish_var.set(f"Expected Finish {finish.strftime('%I:%M %p')}")

    def _open_output_folder(self) -> None:
        if self.current_project is None:
            return
        try:
            os.startfile(str(self.current_project.root))  # type: ignore[attr-defined]
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"{UI_TEXT['open_output_failed']}\n{exc}")

    def _open_package_folder(self) -> None:
        output_dir = self.package_output_dir or packages_dir()
        if not output_dir.exists():
            messagebox.showinfo(APP_NAME, UI_TEXT["package_output_unavailable"])
            return
        try:
            os.startfile(str(output_dir))  # type: ignore[attr-defined]
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"{UI_TEXT['package_output_unavailable']}\n{exc}")

    def _open_review_file(self) -> None:
        review_path = self.review_file_path
        if review_path is None:
            package_dir = self._resolve_package_for_review()
            if package_dir is not None:
                candidate = package_dir / "assistant_review.md"
                review_path = candidate if candidate.exists() else None
        if review_path is None or not review_path.exists():
            messagebox.showinfo(APP_NAME, UI_TEXT["review_file_unavailable"])
            return
        try:
            os.startfile(str(review_path))  # type: ignore[attr-defined]
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"{UI_TEXT['open_review_failed']}\n{exc}")

    def _open_selected_folder(self) -> None:
        selected_dir = self.selected_output_dir
        if selected_dir is None:
            package_dir = self._resolve_package_for_review()
            if package_dir is not None:
                candidate = package_dir / "selected"
                selected_dir = candidate if candidate.exists() else None
        if selected_dir is None or not selected_dir.exists():
            messagebox.showinfo(APP_NAME, UI_TEXT["selected_folder_unavailable"])
            return
        try:
            os.startfile(str(selected_dir))  # type: ignore[attr-defined]
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"{UI_TEXT['open_selected_failed']}\n{exc}")

    def _open_short_preview(self) -> None:
        preview_path = self.short_preview_path
        if preview_path is None:
            package_dir = self._resolve_package_for_review()
            if package_dir is not None:
                candidate = package_dir / "selected" / "short_preview.mp4"
                preview_path = candidate if candidate.exists() else None
        if preview_path is None or not preview_path.exists():
            messagebox.showinfo(APP_NAME, UI_TEXT["short_preview_unavailable"])
            return
        try:
            os.startfile(str(preview_path))  # type: ignore[attr-defined]
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"{UI_TEXT['open_short_preview_failed']}\n{exc}")

    def _open_vertical_short(self) -> None:
        vertical_path = self.vertical_short_path
        if vertical_path is None:
            package_dir = self._resolve_package_for_review()
            if package_dir is not None:
                candidate = package_dir / "selected" / "short_vertical_1080x1920.mp4"
                vertical_path = candidate if candidate.exists() else None
        if vertical_path is None or not vertical_path.exists():
            messagebox.showinfo(APP_NAME, UI_TEXT["vertical_short_unavailable"])
            return
        try:
            os.startfile(str(vertical_path))  # type: ignore[attr-defined]
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"{UI_TEXT['open_vertical_short_failed']}\n{exc}")

    def _open_horizontal_edit(self) -> None:
        horizontal_path = self.horizontal_edit_path
        if horizontal_path is None:
            package_dir = self._resolve_package_for_review()
            if package_dir is not None:
                candidate = package_dir / "selected" / "horizontal_edit.mp4"
                horizontal_path = candidate if candidate.exists() else None
        if horizontal_path is None or not horizontal_path.exists():
            messagebox.showinfo(APP_NAME, UI_TEXT["horizontal_edit_unavailable"])
            return
        try:
            os.startfile(str(horizontal_path))  # type: ignore[attr-defined]
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"{UI_TEXT['open_horizontal_edit_failed']}\n{exc}")

    def _open_memory_folder(self) -> None:
        try:
            output_dir = ensure_memory_dirs()
            self.memory_output_dir = output_dir
            os.startfile(str(output_dir))  # type: ignore[attr-defined]
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"{UI_TEXT['open_memory_failed']}\n{exc}")

    def _open_recommendation(self) -> None:
        recommendation_path = self.recommendation_path
        if recommendation_path is None:
            package_dir = self._resolve_package_for_review()
            if package_dir is not None:
                candidate = package_dir / "assistant_recommendation.md"
                recommendation_path = candidate if candidate.exists() else None
        if recommendation_path is None or not recommendation_path.exists():
            messagebox.showinfo(APP_NAME, UI_TEXT["recommend_file_unavailable"])
            return
        try:
            os.startfile(str(recommendation_path))  # type: ignore[attr-defined]
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"{UI_TEXT['open_recommend_failed']}\n{exc}")

    def _open_test_output(self) -> None:
        output_dir = self.test_output_dir or first_video_test_dir()
        if not output_dir.exists():
            messagebox.showinfo(APP_NAME, UI_TEXT["open_test_output_unavailable"])
            return
        try:
            os.startfile(str(output_dir))  # type: ignore[attr-defined]
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Could not open test output.\n{exc}")

    def _open_install_guide(self) -> None:
        if self.install_guide_path is None or not self.install_guide_path.exists():
            messagebox.showinfo(APP_NAME, UI_TEXT["no_install_guide"])
            return
        try:
            os.startfile(str(self.install_guide_path))  # type: ignore[attr-defined]
            self._log(LOG_TEXT["install_guide_opened"])
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"{UI_TEXT['install_guide_open_failed']}\n{exc}")

    def _copy_install_commands(self) -> None:
        if not self.install_commands:
            self.system_check_status_var.set(UI_TEXT["no_install_commands"])
            return
        text = "\n".join(self.install_commands)
        self.clipboard_clear()
        self.clipboard_append(text)
        self.system_check_status_var.set(UI_TEXT["install_commands_copied"])
        self._log(LOG_TEXT["install_commands_copied"])

    def _set_system_action_buttons(self, state: str) -> None:
        for name in ["system_check_button", "process_system_check_button"]:
            button = getattr(self, name, None)
            if button is not None:
                button.configure(state=state)

    def _format_first_test_summary(self, result: dict[str, object]) -> str:
        media = result.get("media_info")
        selected = Path(str(result.get("selected_file", ""))).name if result.get("selected_file") else "--"
        lines = [
            UI_TEXT["first_video_test"],
            f"{UI_TEXT['test_selected']}: {selected}",
        ]
        if isinstance(media, dict):
            duration = media.get("duration")
            width = media.get("width") or 0
            height = media.get("height") or 0
            lines.extend(
                [
                    f"{UI_TEXT['duration']}: {seconds_to_timecode(float(duration or 0))}",
                    f"{UI_TEXT['codec']}: {media.get('video_codec') or 'unknown'}",
                    f"{UI_TEXT['resolution']}: {width}x{height}",
                ]
            )
        else:
            lines.extend(
                [
                    f"{UI_TEXT['duration']}: --",
                    f"{UI_TEXT['codec']}: --",
                    f"{UI_TEXT['resolution']}: --",
                ]
            )
        lines.extend(
            [
                "",
                f"{UI_TEXT['test_clip']}: {result.get('test_clip', 'UNKNOWN')}",
                f"{UI_TEXT['nvenc']}: {result.get('nvenc', 'UNKNOWN')}",
                f"Encoder: {result.get('encoder', 'none')}",
            ]
        )
        message = str(result.get("message") or "")
        if message:
            lines.append(message)
        return "\n".join(lines)

    def _format_posting_package_summary(self, result: dict[str, object]) -> str:
        package_dir = str(result.get("package_dir") or "")
        generated = result.get("generated")
        generated_items = [str(item) for item in generated] if isinstance(generated, list) else []
        status = str(result.get("status") or UI_TEXT["package_failed"])
        lines = [
            UI_TEXT["posting_package"],
            f"{UI_TEXT['package_status']}: {status}",
            f"{UI_TEXT['package']}: {package_dir or '--'}",
            f"{UI_TEXT['transcription_short']}: {result.get('transcription', 'UNKNOWN')}",
            f"{UI_TEXT['shorts']}: {result.get('shorts_count', 0)}",
            f"{UI_TEXT['ollama']}: {UI_TEXT['used'] if result.get('used_ollama') else UI_TEXT['template_fallback']}",
            "",
            f"{UI_TEXT['generated']}:",
        ]
        if generated_items:
            lines.extend(f"- {Path(item).name}" for item in generated_items[:10])
        else:
            lines.append("- --")
        message = str(result.get("message") or "")
        if message:
            lines.extend(["", message])
        return "\n".join(lines)

    def _format_assistant_review_summary(self, result: dict[str, object]) -> str:
        package_dir = str(result.get("package_dir") or "")
        review_path = str(result.get("review_path") or "")
        files_read = result.get("files_read")
        read_count = len(files_read) if isinstance(files_read, list) else 0
        return "\n".join(
            [
                UI_TEXT["assistant_review"],
                f"{UI_TEXT['package_status']}: {result.get('status', UI_TEXT['review_failed'])}",
                f"{UI_TEXT['package']}: {package_dir or '--'}",
                f"{UI_TEXT['review']}: {review_path or '--'}",
                f"{UI_TEXT['ollama']}: {UI_TEXT['used'] if result.get('used_ollama') else UI_TEXT['template_fallback']}",
                f"{UI_TEXT['files_read']}: {read_count}",
            ]
        )

    def _format_selected_candidates_summary(self, data: dict[str, object]) -> str:
        short_labels = [str(item) for item in data.get("short_labels", []) if str(item)]
        title_labels = [str(item) for item in data.get("title_labels", []) if str(item)]
        meta_parts = [
            f"{UI_TEXT['selected_description']}: {UI_TEXT['selected_available'] if data.get('has_description') else UI_TEXT['selected_missing']}",
            f"{UI_TEXT['selected_tags']}: {UI_TEXT['selected_available'] if data.get('has_tags') else UI_TEXT['selected_missing']}",
            f"{UI_TEXT['selected_notes']}: {UI_TEXT['selected_available'] if data.get('has_notes') else UI_TEXT['selected_missing']}",
            f"{UI_TEXT['selected_review']}: {UI_TEXT['selected_available'] if data.get('has_review') else UI_TEXT['selected_missing']}",
        ]
        lines = [
            UI_TEXT["selected_outputs"],
            f"{UI_TEXT['shorts']}: {len(short_labels)} / {UI_TEXT['selected_title']}: {len(title_labels)}",
            " | ".join(meta_parts),
        ]
        if short_labels:
            lines.append(short_labels[0] if len(short_labels[0]) <= 120 else short_labels[0][:117] + "...")
        if title_labels:
            lines.append(title_labels[0] if len(title_labels[0]) <= 120 else title_labels[0][:117] + "...")
        return "\n".join(lines)

    def _format_selected_export_summary(self, result: dict[str, object]) -> str:
        selected_dir = str(result.get("selected_dir") or "")
        selected_title = str(result.get("selected_title") or "--")
        selected_short = result.get("selected_short")
        short_text = "--"
        if isinstance(selected_short, dict):
            short_text = f"{selected_short.get('start', '--')} - {selected_short.get('end', '--')}"
        return "\n".join(
            [
                UI_TEXT["selected_outputs"],
                UI_TEXT["selected_completed"],
                f"{UI_TEXT['selected_short']}: {short_text}",
                f"{UI_TEXT['selected_title']}: {selected_title}",
                f"{UI_TEXT['package']}: {selected_dir}",
            ]
        )

    def _format_short_preview_summary(self, result: dict[str, object]) -> str:
        output_path = str(result.get("output_path") or "")
        return "\n".join(
            [
                UI_TEXT["short_preview"],
                f"{UI_TEXT['package_status']}: {result.get('status', UI_TEXT['short_preview_failed'])}",
                f"{UI_TEXT['encoder']}: {result.get('encoder', 'unavailable')}",
                f"{UI_TEXT['preview_output']}: {output_path or '--'}",
            ]
        )

    def _format_vertical_short_summary(self, result: dict[str, object]) -> str:
        output_path = str(result.get("output_path") or "")
        return "\n".join(
            [
                UI_TEXT["vertical_short"],
                f"{UI_TEXT['package_status']}: {result.get('status', UI_TEXT['vertical_short_failed'])}",
                f"{UI_TEXT['output_size']}: {result.get('size', '1080x1920')}",
                f"{UI_TEXT['encoder']}: {result.get('encoder', 'unavailable')}",
                f"{UI_TEXT['preview_output']}: {output_path or '--'}",
            ]
        )

    def _format_project_bridge_summary(self, data: dict[str, object], extra: str = "") -> str:
        suggested_use = [str(item) for item in data.get("suggested_use", []) if str(item)]
        bgm_files = [item for item in data.get("bgm_files", []) if isinstance(item, dict)]
        notes_line = "" if data.get("notes_available") else UI_TEXT["project_notes_unavailable"]
        lines = [
            UI_TEXT["project_bridge"],
            f"{UI_TEXT['selected_project']}: {data.get('name', '--')}",
            f"{UI_TEXT['preset']}: {data.get('preset', '--')}",
            f"{UI_TEXT['suggested_use']}: {', '.join(suggested_use[:3]) if suggested_use else '--'}",
            f"{UI_TEXT['bgm']}: {len(bgm_files)}",
        ]
        if notes_line:
            lines.append(notes_line)
        if extra:
            lines.append(extra)
        return "\n".join(lines)

    def _format_project_bridge_result_summary(self, result: dict[str, object]) -> str:
        copied_bgm = str(result.get("copied_bgm") or "")
        metadata_path = str(result.get("metadata_path") or "")
        lines = [
            UI_TEXT["project_bridge"],
            f"{UI_TEXT['package_status']}: {result.get('status', 'UNKNOWN')}",
            f"{UI_TEXT['selected_project']}: {result.get('project', self.project_choice_var.get())}",
        ]
        if result.get("preset"):
            lines.append(f"{UI_TEXT['preset']}: {result.get('preset')}")
        if result.get("bgm") or copied_bgm:
            lines.append(f"{UI_TEXT['bgm']}: {result.get('bgm') or Path(copied_bgm).name}")
        if copied_bgm:
            lines.append(f"{UI_TEXT['preview_output']}: {copied_bgm}")
        if metadata_path:
            lines.append(f"{UI_TEXT['metadata']}: {metadata_path}")
            lines.append(f"{UI_TEXT['ollama']}: {UI_TEXT['used'] if result.get('used_ollama') else UI_TEXT['template_fallback']}")
        message = str(result.get("message") or "")
        if message:
            lines.append(message)
        return "\n".join(lines)

    def _format_memory_summary(self, result: dict[str, object]) -> str:
        record = str(result.get("record_json") or "")
        summary = str(result.get("summary_path") or "")
        lines = [
            UI_TEXT["memory"],
            f"{UI_TEXT['memory_status']}: {result.get('status', UI_TEXT['ready'])}",
            f"{UI_TEXT['memory_entries']}: {result.get('entries', 0)}",
        ]
        if record:
            lines.append(f"{UI_TEXT['memory_record']}: {record}")
        if summary:
            lines.append(f"{UI_TEXT['memory_summary']}: {summary}")
        if "used_ollama" in result:
            lines.append(f"{UI_TEXT['ollama']}: {UI_TEXT['used'] if result.get('used_ollama') else UI_TEXT['template_fallback']}")
        return "\n".join(lines)

    def _format_recommend_memory_summary(self, analysis: dict[str, object]) -> str:
        presets = analysis.get("preset_counts", [])
        bgm = analysis.get("bgm_counts", [])
        preset = ""
        mood = ""
        if isinstance(presets, list) and presets:
            first = presets[0]
            if isinstance(first, dict):
                preset = str(first.get("value") or "")
        if isinstance(bgm, list) and bgm:
            first_bgm = bgm[0]
            if isinstance(first_bgm, dict):
                mood = str(first_bgm.get("value") or "")
        return "\n".join(
            [
                UI_TEXT["assistant_recommend"],
                f"{UI_TEXT['memory_loaded']}: {analysis.get('entries', 0)} {UI_TEXT['projects_loaded']}",
                f"{UI_TEXT['current_mood']}: {' / '.join(part for part in [preset, mood] if part) or '--'}",
            ]
        )

    def _format_recommendation_summary(self, result: dict[str, object]) -> str:
        current = result.get("current_direction", [])
        current_items = [str(item) for item in current if str(item)] if isinstance(current, list) else []
        related = result.get("related_projects", [])
        related_items = [str(item) for item in related if str(item)] if isinstance(related, list) else []
        return "\n".join(
            [
                UI_TEXT["assistant_recommend"],
                f"{UI_TEXT['memory_status']}: {result.get('status', UI_TEXT['ready'])}",
                f"{UI_TEXT['memory_loaded']}: {result.get('memory_entries', 0)} {UI_TEXT['projects_loaded']}",
                f"{UI_TEXT['current_mood']}: {', '.join(current_items[:3]) if current_items else '--'}",
                f"{UI_TEXT['related_projects']}: {', '.join(related_items[:3]) if related_items else '--'}",
                f"{UI_TEXT['ollama']}: {UI_TEXT['used'] if result.get('used_ollama') else UI_TEXT['template_fallback']}",
            ]
        )

    def _drain_events(self) -> None:
        while True:
            try:
                event = self.events.get_nowait()
            except queue.Empty:
                break
            self._handle_event(event)
        self.after(100, self._drain_events)

    def _handle_event(self, event: dict[str, object]) -> None:
        event_type = event.get("type")
        if event_type == "log":
            self._log(str(event.get("message", "")))
        elif event_type == "cli":
            self._apply_cli_status(
                event.get("statuses", {}),
                event.get("nvenc", {}),
                event.get("gpu", {}),
                str(event.get("install_guide", "")),
                event.get("install_commands", []),
            )
        elif event_type == "cli_error":
            self.system_check_status_var.set(f"System check failed: {event.get('message', '')}")
            self._set_system_action_buttons("normal")
            self._log(str(event.get("message", "")))
        elif event_type == "media":
            info = event.get("info")
            if isinstance(info, MediaInfo):
                self.current_media_info = info
                set_textbox(self.media_box, self._media_info_text(info))
                self._set_eta(estimate_processing_seconds(info.duration, is_faster_whisper_available()))
        elif event_type == "media_error":
            set_textbox(self.media_box, str(event.get("message", "Media information unavailable.")))
        elif event_type == "first_test_result":
            result = event.get("result", {})
            if isinstance(result, dict):
                self.test_output_dir = Path(str(result.get("output_dir") or first_video_test_dir()))
                self.first_test_summary_var.set(self._format_first_test_summary(result))
                self.output_var.set(str(self.test_output_dir))
                self.open_test_output_button.configure(state="normal")
                if result.get("test_clip") == "SKIPPED" and result.get("message") == UI_TEXT["first_test_ffmpeg_required"]:
                    self.system_check_status_var.set(UI_TEXT["first_test_ffmpeg_required"])
                self.progress_var.set(1.0)
                self.eta_var.set(UI_TEXT["first_test_completed"])
                self.finish_var.set(UI_TEXT["complete"])
                self.status_var.set(UI_TEXT["complete"])
                self._log(LOG_TEXT["complete"])
            self.first_test_running = False
            self.first_test_button.configure(state="normal")
        elif event_type == "first_test_error":
            self.first_test_running = False
            self.first_test_button.configure(state="normal")
            self.status_var.set(UI_TEXT["error"])
            self.first_test_summary_var.set(str(event.get("message", "")))
        elif event_type == "posting_package_result":
            result = event.get("result", {})
            if isinstance(result, dict):
                package_dir = Path(str(result.get("package_dir") or packages_dir()))
                self.package_output_dir = package_dir
                self.review_file_path = None
                self.package_summary_var.set(self._format_posting_package_summary(result))
                self.review_summary_var.set(
                    "\n".join(
                        [
                            UI_TEXT["assistant_review"],
                            UI_TEXT["review_ready_hint"],
                            f"{UI_TEXT['package']}: {package_dir}",
                        ]
                    )
                )
                self.output_var.set(str(package_dir))
                self.current_project = ProjectPaths.from_root(package_dir)
                self.open_package_button.configure(state="normal")
                self.open_review_button.configure(state="disabled")
                self.open_selected_button.configure(state="disabled")
                self.selected_output_dir = None
                self.short_preview_path = None
                self.vertical_short_path = None
                self.recommendation_path = None
                self.candidate_data = {}
                self.short_choice_touched = False
                self.title_choice_touched = False
                self.short_choice_var.set(UI_TEXT["selected_none"])
                self.title_choice_var.set(UI_TEXT["selected_none"])
                self.short_choice_menu.configure(values=[UI_TEXT["selected_none"]])
                self.title_choice_menu.configure(values=[UI_TEXT["selected_none"]])
                self.selected_summary_var.set(UI_TEXT["selected_ready_hint"])
                self.short_preview_summary_var.set(UI_TEXT["short_preview_ready_hint"])
                self.open_short_preview_button.configure(state="disabled")
                self.vertical_short_summary_var.set(UI_TEXT["vertical_short_ready_hint"])
                self.open_vertical_short_button.configure(state="disabled")
                self.sequence_items = []
                self.horizontal_edit_path = None
                self.sequence_summary_var.set(UI_TEXT["horizontal_edit_ready_hint"])
                self.sequence_choice_var.set(UI_TEXT["sequence_empty"])
                self.sequence_choice_menu.configure(values=[UI_TEXT["sequence_empty"]])
                self.open_horizontal_edit_button.configure(state="disabled")
                self.recommendation_summary_var.set(UI_TEXT["recommend_ready_hint"])
                self.open_recommendation_button.configure(state="disabled")
                self.open_button.configure(state="normal")
                self.progress_var.set(1.0)
                status = str(result.get("status") or "")
                if status == "COMPLETED":
                    self.status_var.set(UI_TEXT["complete"])
                    self.step_vars[UI_TEXT["export_package"]].set(UI_TEXT["done"])
                    self.eta_var.set(UI_TEXT["package_completed"])
                    self.finish_var.set(UI_TEXT["complete"])
                    self._log(LOG_TEXT["complete"])
                else:
                    self.status_var.set(UI_TEXT["error"])
                    self.step_vars[UI_TEXT["export_package"]].set(UI_TEXT["error"])
            self.package_running = False
            self._set_package_buttons("normal")
        elif event_type == "posting_package_error":
            self.package_running = False
            self._set_package_buttons("normal")
            self.status_var.set(UI_TEXT["error"])
            self.step_vars[UI_TEXT["export_package"]].set(UI_TEXT["error"])
            self.package_summary_var.set(str(event.get("message", "")))
        elif event_type == "assistant_review_result":
            result = event.get("result", {})
            if isinstance(result, dict):
                review_path = Path(str(result.get("review_path") or ""))
                package_dir = Path(str(result.get("package_dir") or packages_dir()))
                self.package_output_dir = package_dir
                self.review_file_path = review_path if review_path.exists() else None
                self.review_summary_var.set(self._format_assistant_review_summary(result))
                self.output_var.set(str(package_dir))
                self.current_project = ProjectPaths.from_root(package_dir)
                self.open_package_button.configure(state="normal")
                self.open_review_button.configure(state="normal" if self.review_file_path else "disabled")
                self.open_button.configure(state="normal")
                self.progress_var.set(1.0)
                self.status_var.set(UI_TEXT["complete"])
                self.eta_var.set(UI_TEXT["review_completed"])
                self.finish_var.set(UI_TEXT["complete"])
                self._log(LOG_TEXT["complete"])
            self.review_running = False
            self.assistant_review_button.configure(state="normal")
        elif event_type == "assistant_review_error":
            self.review_running = False
            self.assistant_review_button.configure(state="normal")
            self.status_var.set(UI_TEXT["error"])
            self.review_summary_var.set(str(event.get("message", "")))
        elif event_type == "selected_export_result":
            result = event.get("result", {})
            if isinstance(result, dict):
                selected_dir = Path(str(result.get("selected_dir") or ""))
                package_dir = Path(str(result.get("package_dir") or packages_dir()))
                self.package_output_dir = package_dir
                self.selected_output_dir = selected_dir if selected_dir.exists() else None
                self.selected_summary_var.set(self._format_selected_export_summary(result))
                self.open_selected_button.configure(state="normal" if self.selected_output_dir else "disabled")
                self.short_preview_path = None
                self.open_short_preview_button.configure(state="disabled")
                self.short_preview_summary_var.set(UI_TEXT["short_preview_ready_hint"])
                self.vertical_short_path = None
                self.open_vertical_short_button.configure(state="disabled")
                self.vertical_short_summary_var.set(UI_TEXT["vertical_short_ready_hint"])
                self.recommendation_path = None
                self.open_recommendation_button.configure(state="disabled")
                self.recommendation_summary_var.set(UI_TEXT["recommend_ready_hint"])
                self.output_var.set(str(selected_dir))
                self.progress_var.set(1.0)
                self.status_var.set(UI_TEXT["complete"])
                self.eta_var.set(UI_TEXT["selected_completed"])
                self.finish_var.set(UI_TEXT["complete"])
            self.selected_export_running = False
            self.export_selected_button.configure(state="normal")
        elif event_type == "selected_export_error":
            self.selected_export_running = False
            self.export_selected_button.configure(state="normal")
            self.status_var.set(UI_TEXT["error"])
            self.selected_summary_var.set(str(event.get("message", "")))
        elif event_type == "short_preview_result":
            result = event.get("result", {})
            if isinstance(result, dict):
                output_path = Path(str(result.get("output_path") or ""))
                selected_dir = Path(str(result.get("selected_dir") or ""))
                package_dir = Path(str(result.get("package_dir") or packages_dir()))
                self.package_output_dir = package_dir
                self.selected_output_dir = selected_dir if selected_dir.exists() else None
                self.short_preview_path = output_path if output_path.exists() else None
                self.short_preview_summary_var.set(self._format_short_preview_summary(result))
                self.open_selected_button.configure(state="normal" if self.selected_output_dir else "disabled")
                self.open_short_preview_button.configure(state="normal" if self.short_preview_path else "disabled")
                self.progress_var.set(1.0)
                if result.get("status") == "COMPLETED":
                    self.status_var.set(UI_TEXT["complete"])
                    self.eta_var.set(UI_TEXT["short_preview_completed"])
                    self.finish_var.set(UI_TEXT["complete"])
                else:
                    self.status_var.set(UI_TEXT["error"])
            self.short_preview_running = False
            self.generate_short_preview_button.configure(state="normal")
        elif event_type == "short_preview_error":
            self.short_preview_running = False
            self.generate_short_preview_button.configure(state="normal")
            self.status_var.set(UI_TEXT["error"])
            self.short_preview_summary_var.set(str(event.get("message", "")))
        elif event_type == "vertical_short_result":
            result = event.get("result", {})
            if isinstance(result, dict):
                output_path = Path(str(result.get("output_path") or ""))
                selected_dir = Path(str(result.get("selected_dir") or ""))
                package_dir = Path(str(result.get("package_dir") or packages_dir()))
                self.package_output_dir = package_dir
                self.selected_output_dir = selected_dir if selected_dir.exists() else None
                self.vertical_short_path = output_path if output_path.exists() else None
                self.vertical_short_summary_var.set(self._format_vertical_short_summary(result))
                self.open_selected_button.configure(state="normal" if self.selected_output_dir else "disabled")
                self.open_vertical_short_button.configure(state="normal" if self.vertical_short_path else "disabled")
                self.progress_var.set(1.0)
                if result.get("status") == "COMPLETED":
                    self.status_var.set(UI_TEXT["complete"])
                    self.eta_var.set(UI_TEXT["vertical_short_completed"])
                    self.finish_var.set(UI_TEXT["complete"])
                else:
                    self.status_var.set(UI_TEXT["error"])
            self.vertical_short_running = False
            self.generate_vertical_short_button.configure(state="normal")
        elif event_type == "vertical_short_error":
            self.vertical_short_running = False
            self.generate_vertical_short_button.configure(state="normal")
            self.status_var.set(UI_TEXT["error"])
            self.vertical_short_summary_var.set(str(event.get("message", "")))
        elif event_type == "sequence_probe_result":
            package_dir = Path(str(event.get("package_dir") or ""))
            sequence = event.get("sequence", [])
            if isinstance(sequence, list) and self.package_output_dir and package_dir == self.package_output_dir:
                self.sequence_items = [dict(item) for item in sequence if isinstance(item, dict)]
                write_sequence(package_dir, self.sequence_items)
                self._refresh_sequence_ui()
        elif event_type == "horizontal_edit_result":
            result = event.get("result", {})
            if isinstance(result, dict):
                output_path = Path(str(result.get("output_path") or ""))
                selected_dir = Path(str(result.get("selected_dir") or ""))
                package_dir = Path(str(result.get("package_dir") or packages_dir()))
                self.package_output_dir = package_dir
                self.selected_output_dir = selected_dir if selected_dir.exists() else None
                self.horizontal_edit_path = output_path if output_path.exists() else None
                self.sequence_items = read_sequence(package_dir)
                self.sequence_summary_var.set(self._format_sequence_summary(result))
                self.open_selected_button.configure(state="normal" if self.selected_output_dir else "disabled")
                self.open_horizontal_edit_button.configure(state="normal" if self.horizontal_edit_path else "disabled")
                self.progress_var.set(1.0)
                if result.get("status") == "COMPLETED":
                    self.status_var.set(UI_TEXT["complete"])
                    self.eta_var.set(UI_TEXT["horizontal_edit_completed"])
                    self.finish_var.set(UI_TEXT["complete"])
                else:
                    self.status_var.set(UI_TEXT["error"])
            self.sequence_running = False
            self.generate_horizontal_edit_button.configure(state="normal")
        elif event_type == "horizontal_edit_error":
            self.sequence_running = False
            self.generate_horizontal_edit_button.configure(state="normal")
            self.status_var.set(UI_TEXT["error"])
            self.sequence_summary_var.set(str(event.get("message", "")))
        elif event_type == "project_bridge_copy_result":
            result = event.get("result", {})
            if isinstance(result, dict):
                package_dir = Path(str(result.get("package_dir") or packages_dir()))
                selected_dir = Path(str(result.get("selected_dir") or package_dir / "selected"))
                self.package_output_dir = package_dir
                self.selected_output_dir = selected_dir if selected_dir.exists() else None
                self.current_project = ProjectPaths.from_root(package_dir)
                self.output_var.set(str(selected_dir))
                self.open_selected_button.configure(state="normal" if self.selected_output_dir else "disabled")
                self.open_button.configure(state="normal")
                self.project_bridge_summary_var.set(self._format_project_bridge_result_summary(result))
                self.recommendation_path = None
                self.open_recommendation_button.configure(state="disabled")
                self.recommendation_summary_var.set(UI_TEXT["recommend_ready_hint"])
                self.progress_var.set(1.0)
                self.status_var.set(UI_TEXT["complete"] if result.get("status") == "COMPLETED" else UI_TEXT["error"])
                if result.get("status") == "COMPLETED":
                    self._log(LOG_TEXT["project_box_connected"])
                    self._log(LOG_TEXT["project_bridge_ready"])
            self.project_bridge_running = False
            self.add_project_bgm_button.configure(state="normal")
        elif event_type == "project_bridge_metadata_result":
            result = event.get("result", {})
            if isinstance(result, dict):
                package_dir = Path(str(result.get("package_dir") or packages_dir()))
                selected_dir = Path(str(result.get("selected_dir") or package_dir / "selected"))
                metadata_path = Path(str(result.get("metadata_path") or ""))
                self.package_output_dir = package_dir
                self.selected_output_dir = selected_dir if selected_dir.exists() else None
                self.project_bridge_metadata_path = metadata_path if metadata_path.exists() else None
                self.current_project = ProjectPaths.from_root(package_dir)
                self.output_var.set(str(metadata_path))
                self.open_selected_button.configure(state="normal" if self.selected_output_dir else "disabled")
                self.open_button.configure(state="normal")
                self.project_bridge_summary_var.set(self._format_project_bridge_result_summary(result))
                self.progress_var.set(1.0)
                self.status_var.set(UI_TEXT["complete"])
                self.eta_var.set(UI_TEXT["project_bridge_metadata_completed"])
                self.finish_var.set(UI_TEXT["complete"])
                self._log(LOG_TEXT["project_metadata_ready"])
                if not result.get("used_ollama"):
                    self._log(LOG_TEXT["project_ollama_fallback"])
                self._log(LOG_TEXT["project_bridge_ready"])
            self.project_bridge_running = False
            self.generate_bridge_metadata_button.configure(state="normal")
        elif event_type == "project_bridge_error":
            self.project_bridge_running = False
            self.add_project_bgm_button.configure(state="normal")
            self.generate_bridge_metadata_button.configure(state="normal")
            self.status_var.set(UI_TEXT["error"])
            self.project_bridge_summary_var.set(str(event.get("message", "")))
        elif event_type == "memory_save_result":
            result = event.get("result", {})
            if isinstance(result, dict):
                self.memory_output_dir = Path(str(result.get("memory_dir") or ensure_memory_dirs()))
                summary_path = Path(str(result.get("summary_path") or memory_summary_path()))
                self.memory_summary_file_path = summary_path if summary_path.exists() else None
                self.memory_summary_var.set(self._format_memory_summary(result))
                self.output_var.set(str(self.memory_output_dir))
                self.progress_var.set(1.0)
                self.status_var.set(UI_TEXT["complete"])
                self.eta_var.set(UI_TEXT["memory_saved"])
                self.finish_var.set(UI_TEXT["complete"])
                self._log(LOG_TEXT["memory_saved"])
                if not result.get("used_ollama"):
                    self._log(LOG_TEXT["memory_template_fallback"])
                self._log(LOG_TEXT["memory_ready"])
            self.memory_running = False
            self._set_memory_buttons("normal")
        elif event_type == "memory_summary_result":
            result = event.get("result", {})
            if isinstance(result, dict):
                self.memory_output_dir = Path(str(result.get("memory_dir") or ensure_memory_dirs()))
                summary_path = Path(str(result.get("summary_path") or memory_summary_path()))
                self.memory_summary_file_path = summary_path if summary_path.exists() else None
                self.memory_summary_var.set(self._format_memory_summary(result))
                self.output_var.set(str(summary_path))
                self.progress_var.set(1.0)
                self.status_var.set(UI_TEXT["complete"])
                self.eta_var.set(UI_TEXT["memory_updated"])
                self.finish_var.set(UI_TEXT["complete"])
                if not result.get("used_ollama"):
                    self._log(LOG_TEXT["memory_template_fallback"])
                self._log(LOG_TEXT["memory_ready"])
            self.memory_running = False
            self._set_memory_buttons("normal")
        elif event_type == "memory_error":
            self.memory_running = False
            self._set_memory_buttons("normal")
            self.status_var.set(UI_TEXT["error"])
            self.memory_summary_var.set(str(event.get("message", "")))
        elif event_type == "recommendation_result":
            result = event.get("result", {})
            if isinstance(result, dict):
                recommendation_path = Path(str(result.get("recommendation_path") or ""))
                package_dir = Path(str(result.get("package_dir") or packages_dir()))
                self.package_output_dir = package_dir
                self.current_project = ProjectPaths.from_root(package_dir)
                self.recommendation_path = recommendation_path if recommendation_path.exists() else None
                self.recommendation_summary_var.set(self._format_recommendation_summary(result))
                self.output_var.set(str(recommendation_path))
                self.open_recommendation_button.configure(state="normal" if self.recommendation_path else "disabled")
                self.open_package_button.configure(state="normal")
                self.open_button.configure(state="normal")
                self.progress_var.set(1.0)
                self.status_var.set(UI_TEXT["complete"])
                self.eta_var.set(UI_TEXT["complete"])
                self.finish_var.set(UI_TEXT["complete"])
                if "BORINEF" in json.dumps(result, ensure_ascii=False):
                    self._log(LOG_TEXT["recommend_borinef"])
                if not result.get("used_ollama"):
                    self._log(LOG_TEXT["recommend_template_fallback"])
                self._log(LOG_TEXT["recommend_ready"])
            self.recommendation_running = False
            self._set_recommendation_buttons("normal")
        elif event_type == "recommendation_error":
            self.recommendation_running = False
            self._set_recommendation_buttons("normal")
            self.status_var.set(UI_TEXT["error"])
            self.recommendation_summary_var.set(str(event.get("message", "")))
        elif event_type == "youtube":
            self.youtube_status_var.set(str(event.get("message", "")))
        elif event_type == "progress":
            self.progress_var.set(float(event.get("value", 0.0)))
        elif event_type == "eta":
            self._set_eta(float(event.get("seconds", 0.0)))
        elif event_type == "step":
            label = str(event.get("label", ""))
            state = str(event.get("state", ""))
            if label in self.step_vars:
                self.step_vars[label].set(state)
        elif event_type == "output":
            path = Path(str(event.get("path", "")))
            self.current_project = ProjectPaths.from_root(path)
            self.output_var.set(str(path))
            self.open_button.configure(state="normal")
        elif event_type == "complete":
            path = Path(str(event.get("path", "")))
            self.current_project = ProjectPaths.from_root(path)
            self.output_var.set(str(path))
            self.open_button.configure(state="normal")
            self.worker_running = False
            self.start_button.configure(state="normal")
            self.status_var.set(UI_TEXT["complete"])
        elif event_type == "error":
            self.worker_running = False
            self.start_button.configure(state="normal")
            self.status_var.set(UI_TEXT["error"])
            messagebox.showerror(APP_NAME, str(event.get("message", "Unknown error")))
        self._update_dashboard()

    def _apply_cli_status(
        self,
        statuses: object,
        nvenc: object,
        gpu: object,
        install_guide: str = "",
        install_commands: object = None,
    ) -> None:
        if isinstance(statuses, dict):
            self.cli_status = statuses  # type: ignore[assignment]
        if isinstance(nvenc, dict):
            self.nvenc_status = nvenc  # type: ignore[assignment]
        if isinstance(gpu, dict):
            self.gpu_status = gpu  # type: ignore[assignment]
        self.system_check_completed = True
        self.install_guide_path = Path(install_guide) if install_guide else None
        self.install_commands = [str(item) for item in install_commands] if isinstance(install_commands, list) else []

        for key, spec in CLI_TOOLS.items():
            label = spec["label"]
            state = "UNKNOWN"
            detail = ""
            status = self.cli_status.get(key, {})
            if status:
                state = str(status.get("state") or "UNKNOWN")
                detail = str(status.get("detail") or "")
            if label in self.cli_pills:
                self.cli_pills[label].set_state(state, detail)

        self.cli_pills["NVENC"].set_state(self.nvenc_status.get("state", "UNKNOWN"), self.nvenc_status.get("detail", ""))
        self.cli_pills["GPU"].set_state(self.gpu_status.get("state", "UNKNOWN"), self.gpu_status.get("detail", ""))
        self.system_check_status_var.set(UI_TEXT["system_check_done"])
        self._set_system_action_buttons("normal")
        if hasattr(self, "open_install_guide_button"):
            self.open_install_guide_button.configure(state="normal" if self.install_guide_path else "disabled")
        if hasattr(self, "copy_install_commands_button"):
            self.copy_install_commands_button.configure(state="normal" if self.install_commands else "disabled")
        self._log(LOG_TEXT["system_check_complete"])
        if self._has_system_issues():
            self._log(LOG_TEXT["tools_missing"])
        if self.install_commands:
            self._log(LOG_TEXT["cli_disconnected"])
            self._log(LOG_TEXT["install_candidates_ready"])
        if self.cli_status.get("gh", {}).get("state") == "UNAUTHORIZED":
            self._log(LOG_TEXT["github_unauthorized"])
        if self.cli_status.get("wrangler", {}).get("state") == "UNAUTHORIZED":
            self._log(LOG_TEXT["wrangler_unauthorized"])
        if self.cli_status.get("ollama", {}).get("state") == "MISSING":
            self._log(LOG_TEXT["ollama_sleeping"])
        if self.nvenc_status.get("state") == "ONLINE":
            self._log(LOG_TEXT["gpu_encode_ready"])
        if install_guide:
            self._log(f"install_guide.txt: {install_guide}")
        self._log(LOG_TEXT["complete"])
        if self.selected_file is not None:
            self._probe_selected_video()

    def _has_system_issues(self) -> bool:
        issue_states = {"MISSING", "UNAUTHORIZED", "UNAVAILABLE"}
        for status in self.cli_status.values():
            if str(status.get("state")) in issue_states:
                return True
        return self.nvenc_status.get("state") in {"UNAVAILABLE", "CHECK SKIPPED"}

    def _media_info_text(self, info: MediaInfo) -> str:
        lines = [
            f"File: {info.file_name}",
            f"Size: {human_size(info.file_size_bytes)}",
            f"Duration: {format_duration(info.duration)}",
            f"Resolution: {info.width} x {info.height}",
            f"FPS: {info.fps:.3f}" if info.fps else "FPS: unknown",
            f"Video Codec: {info.video_codec or 'unknown'}",
            f"Audio: {'yes' if info.audio_present else 'no'}",
        ]
        if info.audio_codec:
            lines.append(f"Audio Codec: {info.audio_codec}")
        return "\n".join(lines)

    def _log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_lines.append(f"[{timestamp}] {message}")
        self.log_lines = self.log_lines[-20:]
        set_textbox(self.log_box, "\n".join(self.log_lines) + ("\n" if self.log_lines else ""))


def run_launch_check() -> int:
    ensure_app_dirs()
    system = run_system_check()
    result = {
        "app": APP_NAME,
        "version": APP_VERSION,
        "exe_name": EXE_NAME,
        "app_root": str(app_root()),
        "cli": system["cli"],
        "nvenc": system["nvenc"],
        "gpu": system["gpu"],
        "install_guide": system["install_guide"],
        "install_commands": system["install_commands"],
        "faster_whisper": is_faster_whisper_available(),
    }
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


def _write_smoke_wav(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with wave.open(str(path), "wb") as wav:
        wav.setnchannels(1)
        wav.setsampwidth(2)
        wav.setframerate(8000)
        wav.writeframes(b"\x00\x00" * 800)


def _write_smoke_project_box(project_root: Path) -> Path:
    bgm_path = project_root / "bgm" / "smoke_bridge.wav"
    _write_smoke_wav(bgm_path)
    notes_dir = project_root / "notes"
    notes_dir.mkdir(parents=True, exist_ok=True)
    (notes_dir / "project_notes.txt").write_text(
        "\n".join(
            [
                "Project:",
                project_root.name,
                "",
                "Selected Preset:",
                "BORINEF",
                "",
                "BGM:",
                bgm_path.name,
                "",
                "Suggested Title:",
                "深夜、まだ作ってる。",
                "",
                "Mood:",
                "quiet midnight work",
                "",
                "Suggested Use:",
                "静かな余熱",
                "Shorts背景",
                "深夜のコード作業",
                "",
                "Shorts Direction:",
                "静かなタイピング",
                "深夜の机",
                "まだ作ってる。",
            ]
        )
        + "\n",
        encoding="utf-8",
    )
    for name in ["raw", "shorts", "thumbnails", "export", "upload"]:
        (project_root / name).mkdir(parents=True, exist_ok=True)
    return bgm_path


def run_smoke_test() -> int:
    ensure_app_dirs()
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_root = Path(temp_dir)
        projects_root = temp_root / "projects"
        project_root = projects_root / "smoke_project_box"
        bgm_path = _write_smoke_project_box(project_root)
        boxes = list_project_boxes(projects_root)
        if not boxes:
            raise AssertionError("Project Box list was not populated.")
        project = read_project_box(Path(boxes[0]["root"]))
        if project.get("preset") != "BORINEF":
            raise AssertionError("project_notes.txt preset was not read.")
        if not project.get("bgm_files"):
            raise AssertionError("BGM list was not populated.")
        package_dir = temp_root / "package"
        copy_result = add_bgm_to_video_box(package_dir, bgm_path)
        copied_bgm = Path(str(copy_result.get("copied_bgm") or ""))
        if not copied_bgm.exists():
            raise AssertionError("BGM copy failed.")
        metadata_result = generate_bridge_metadata_draft(package_dir, project, copied_bgm, ollama_ready=False)
        metadata_path = Path(str(metadata_result.get("metadata_path") or ""))
        if not metadata_path.exists():
            raise AssertionError("metadata_draft.txt was not created.")
        memory_result = save_package_to_memory(package_dir, ollama_ready=False, base_dir=temp_root / "memory")
        if not Path(str(memory_result.get("index_path") or "")).exists():
            raise AssertionError("memory_index.json was not created.")
        if not Path(str(memory_result.get("summary_path") or "")).exists():
            raise AssertionError("memory_summary.md was not created.")
        if not Path(str(memory_result.get("record_json") or "")).exists():
            raise AssertionError("memory project json was not created.")
        if not Path(str(memory_result.get("record_md") or "")).exists():
            raise AssertionError("memory project md was not created.")
        recommend_result = generate_assistant_recommendation(package_dir, ollama_ready=False, memory_base_dir=temp_root / "memory")
        if not Path(str(recommend_result.get("recommendation_path") or "")).exists():
            raise AssertionError("assistant_recommendation.md was not created.")
        preview = AudioPreviewPlayer(allow_startfile=False)
        preview_result = preview.play(bgm_path)
        stop_result = preview.stop()
        if not stop_result.success:
            raise AssertionError(f"preview stop failed: {stop_result.message}")
        result = {
            "project_boxes": len(boxes),
            "project": project.get("name"),
            "preset": project.get("preset"),
            "bgm": copied_bgm.name,
            "metadata": metadata_path.name,
            "memory_entries": memory_result.get("entries"),
            "memory_used_ollama": memory_result.get("used_ollama"),
            "recommendation": Path(str(recommend_result.get("recommendation_path") or "")).name,
            "recommendation_used_ollama": recommend_result.get("used_ollama"),
            "preview_backend": preview_result.backend,
            "preview_success": preview_result.success,
        }
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


def run_generate_check() -> int:
    ensure_app_dirs()
    boxes = list_project_boxes()
    result: dict[str, object] = {
        "project_boxes": len(boxes),
        "projects_root": str(Path(boxes[0]["root"]).parent if boxes else BRIDGE_TEXT["no_project_boxes"]),
    }
    if boxes:
        project = read_project_box(Path(boxes[0]["root"]))
        bgm_items = project.get("bgm_files", [])
        bgm_path = None
        if isinstance(bgm_items, list) and bgm_items:
            first = bgm_items[0]
            if isinstance(first, dict):
                bgm_path = Path(str(first.get("path") or ""))
        package_dir = packages_dir() / f"project_bridge_generate_check_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        package_dir.mkdir(parents=True, exist_ok=True)
        copied = add_bgm_to_video_box(package_dir, bgm_path) if bgm_path else {}
        copied_value = str(copied.get("copied_bgm") or "") if copied else ""
        copied_bgm = Path(copied_value) if copied_value else None
        system = run_system_check()
        metadata = generate_bridge_metadata_draft(
            package_dir=package_dir,
            project=project,
            bgm_path=copied_bgm if copied_bgm and copied_bgm.exists() else bgm_path,
            ollama_ready=system["cli"].get("ollama", {}).get("state") == "READY",
        )
        memory = save_package_to_memory(
            package_dir=package_dir,
            ollama_ready=system["cli"].get("ollama", {}).get("state") == "READY",
        )
        recommendation = generate_assistant_recommendation(
            package_dir=package_dir,
            ollama_ready=system["cli"].get("ollama", {}).get("state") == "READY",
        )
        result.update(
            {
                "project": project.get("name"),
                "preset": project.get("preset"),
                "bgm_count": len(bgm_items) if isinstance(bgm_items, list) else 0,
                "copied_bgm": copied.get("copied_bgm", "") if copied else "",
                "metadata_path": metadata.get("metadata_path"),
                "used_ollama": metadata.get("used_ollama"),
                "ollama_model": metadata.get("ollama_model", ""),
                "memory_index": memory.get("index_path"),
                "memory_summary": memory.get("summary_path"),
                "memory_record_json": memory.get("record_json"),
                "memory_record_md": memory.get("record_md"),
                "memory_used_ollama": memory.get("used_ollama"),
                "recommendation_path": recommendation.get("recommendation_path"),
                "recommendation_used_ollama": recommendation.get("used_ollama"),
                "recommendation_memory_entries": recommendation.get("memory_entries"),
            }
        )
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=APP_NAME)
    parser.add_argument("--launch-check", action="store_true", help="Run startup checks without opening the GUI.")
    parser.add_argument("--generate-check", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--smoke-test", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--gui-smoke-seconds", type=float, default=0.0, help=argparse.SUPPRESS)
    args = parser.parse_args()

    if args.launch_check:
        return run_launch_check()
    if args.generate_check:
        return run_generate_check()
    if args.smoke_test:
        return run_smoke_test()

    if ctk is None:
        message = "customtkinter is not installed. Run: pip install -r requirements.txt"
        print(message, file=sys.stderr)
        try:
            messagebox.showerror(APP_NAME, message)
        except Exception:
            pass
        return 1

    setup_theme(ctk)
    app = KadouChuApp()
    if args.gui_smoke_seconds > 0:
        app.after(int(args.gui_smoke_seconds * 1000), app.destroy)
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
