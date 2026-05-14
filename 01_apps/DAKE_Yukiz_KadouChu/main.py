from __future__ import annotations

import argparse
import json
import os
import queue
import sys
import threading
import traceback
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
from core.shorts_analyzer import create_shorts_candidates, write_shorts_candidates
from core.transcription import is_faster_whisper_available, transcribe_media
from ui.theme import COLORS, FONT_FAMILY, setup_theme

if ctk is not None:
    from ui.components import StatusPill, append_textbox, make_panel, set_textbox


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
    "process": "PROCESS",
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
        self.preview_source_video_path: Path | None = None
        self.candidate_data: dict[str, object] = {}
        self.short_choice_touched = False
        self.title_choice_touched = False
        self.first_test_running = False
        self.package_running = False
        self.review_running = False
        self.selected_export_running = False
        self.short_preview_running = False
        self.vertical_short_running = False
        self.worker_running = False

        self.file_var = ctk.StringVar(value=UI_TEXT["no_video_selected"])
        self.test_file_var = ctk.StringVar(value=UI_TEXT["test_not_selected"])
        self.first_test_summary_var = ctk.StringVar(value=UI_TEXT["first_test_ready_hint"])
        self.package_summary_var = ctk.StringVar(value=UI_TEXT["package_ready_hint"])
        self.review_summary_var = ctk.StringVar(value=UI_TEXT["review_ready_hint"])
        self.selected_summary_var = ctk.StringVar(value=UI_TEXT["selected_ready_hint"])
        self.short_preview_summary_var = ctk.StringVar(value=UI_TEXT["short_preview_ready_hint"])
        self.vertical_short_summary_var = ctk.StringVar(value=UI_TEXT["vertical_short_ready_hint"])
        self.short_choice_var = ctk.StringVar(value=UI_TEXT["selected_none"])
        self.title_choice_var = ctk.StringVar(value=UI_TEXT["selected_none"])
        self.youtube_var = ctk.StringVar(value="")
        self.youtube_status_var = ctk.StringVar(value=UI_TEXT["metadata_fetch_optional"])
        self.system_check_status_var = ctk.StringVar(value=UI_TEXT["system_check_not_run"])
        self.eta_var = ctk.StringVar(value=UI_TEXT["eta_empty"])
        self.finish_var = ctk.StringVar(value=UI_TEXT["finish_empty"])
        self.output_var = ctk.StringVar(value=UI_TEXT["no_output"])
        self.status_var = ctk.StringVar(value=UI_TEXT["ready"])
        self.progress_var = ctk.DoubleVar(value=0.0)

        self.step_vars: dict[str, ctk.StringVar] = {}
        self.cli_pills: dict[str, StatusPill] = {}

        self._build_ui()
        self._log(LOG_TEXT["startup"])
        self._log(LOG_TEXT["running"])
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
        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=0)

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

        main = ctk.CTkFrame(self, fg_color="transparent")
        main.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 10))
        main.grid_columnconfigure(0, weight=1, uniform="main")
        main.grid_columnconfigure(1, weight=1, uniform="main")
        main.grid_columnconfigure(2, weight=1, uniform="main")
        main.grid_rowconfigure(0, weight=1)

        self._build_input_panel(main).grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self._build_process_panel(main).grid(row=0, column=1, sticky="nsew", padx=8)
        self._build_output_panel(main).grid(row=0, column=2, sticky="nsew", padx=(8, 0))

        bottom = ctk.CTkFrame(self, fg_color="transparent")
        bottom.grid(row=2, column=0, sticky="ew", padx=20, pady=(0, 16))
        bottom.grid_columnconfigure(0, weight=2)
        bottom.grid_columnconfigure(1, weight=3)

        self._build_system_panel(bottom).grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self._build_log_panel(bottom).grid(row=0, column=1, sticky="nsew", padx=(8, 0))

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
        panel, body = make_panel(parent, UI_TEXT["process"])
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
        self.open_test_output_button.grid(row=3, column=0, columnspan=2, sticky="ew", padx=12, pady=(4, 8))
        ctk.CTkLabel(
            test_box,
            textvariable=self.first_test_summary_var,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=COLORS["text"],
            wraplength=330,
            justify="left",
        ).grid(row=4, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 12))
        return panel

    def _build_output_panel(self, parent: ctk.CTkFrame) -> ctk.CTkFrame:
        panel, body = make_panel(parent, UI_TEXT["output"])
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
        self.open_button.grid(row=5, column=0, sticky="ew", pady=(12, 0))
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
        self.selected_output_dir = None
        self.short_preview_path = None
        self.vertical_short_path = None
        self.preview_source_video_path = None
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
        self._log(LOG_TEXT["source_detected"])
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
        self.generate_package_button.configure(state="disabled")
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
        self.system_check_status_var.set(UI_TEXT["checking"])
        self._set_system_action_buttons("disabled")
        for pill in self.cli_pills.values():
            pill.set_state("CHECKING")

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
            self.generate_package_button.configure(state="normal")
        elif event_type == "posting_package_error":
            self.package_running = False
            self.generate_package_button.configure(state="normal")
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
        append_textbox(self.log_box, f"[{timestamp}] {message}\n")


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


def main() -> int:
    parser = argparse.ArgumentParser(description=APP_NAME)
    parser.add_argument("--launch-check", action="store_true", help="Run startup checks without opening the GUI.")
    parser.add_argument("--gui-smoke-seconds", type=float, default=0.0, help=argparse.SUPPRESS)
    args = parser.parse_args()

    if args.launch_check:
        return run_launch_check()

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
