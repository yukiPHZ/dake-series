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
)
from core.cli_checker import (
    CLI_TOOLS,
    check_cli_environment,
    fetch_youtube_metadata,
    run_system_check,
)
from core.ffmpeg_runner import create_preview_clip
from core.media_probe import MediaInfo, probe_media
from core.ollama_client import build_metadata_draft
from core.project_writer import (
    ProjectPaths,
    create_project,
    write_log_files,
    write_media_info,
    write_metadata_files,
    write_preview_note,
    write_source_manifest,
)
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
    "checking": "Checking...",
    "system_check_done": "System check completed.\n整っています。",
    "output": "OUTPUT",
    "open_output": "Open Output Folder",
    "system": "SYSTEM / CLI STATUS",
    "assistant_log": "補助脳 LOG",
    "ready": "READY",
    "running": "RUNNING",
    "complete": "整っています。",
    "error": "ERROR",
    "no_file": "Select a video file first.",
    "file_types": "Video files",
    "transcription": "Transcription",
    "scene_analysis": "Scene Analysis",
    "shorts_candidates": "Shorts Candidates",
    "thumbnail_base": "Thumbnail Base",
    "export_package": "Export Package",
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
        self.worker_running = False

        self.file_var = ctk.StringVar(value="No video selected")
        self.youtube_var = ctk.StringVar(value="")
        self.youtube_status_var = ctk.StringVar(value="Metadata fetch is optional in Phase 1.")
        self.system_check_status_var = ctk.StringVar(value="System check has not run yet.")
        self.eta_var = ctk.StringVar(value="ETA --")
        self.finish_var = ctk.StringVar(value="Expected Finish --")
        self.output_var = ctk.StringVar(value="No output yet")
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
        ctk.CTkButton(
            body,
            text=UI_TEXT["run_system_check"],
            command=self._start_cli_check,
            height=32,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        ).grid(row=9, column=0, sticky="ew")

        ctk.CTkLabel(
            body,
            textvariable=self.status_var,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=COLORS["accent"],
        ).grid(row=10, column=0, sticky="w", pady=(18, 0))
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
            height=230,
            fg_color=COLORS["field"],
            border_width=1,
            border_color=COLORS["line"],
            text_color=COLORS["text"],
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            wrap="word",
        )
        self.media_box.grid(row=1, column=0, sticky="nsew")
        set_textbox(self.media_box, "Media information will appear here.")

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
        self.open_button.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        return panel

    def _build_system_panel(self, parent: ctk.CTkFrame) -> ctk.CTkFrame:
        panel, body = make_panel(parent, UI_TEXT["system"])
        body.grid_columnconfigure(0, weight=1)
        body.grid_columnconfigure(1, weight=1)

        labels = ["FFMPEG", "FFPROBE", "YT-DLP", "GH", "WRANGLER", "OLLAMA", "NVENC", "GPU"]
        for index, label in enumerate(labels):
            pill = StatusPill(body, label)
            pill.grid(row=index // 2, column=index % 2, sticky="ew", padx=4, pady=4)
            self.cli_pills[label] = pill
        self.system_check_button = ctk.CTkButton(
            body,
            text=UI_TEXT["run_system_check"],
            command=self._start_cli_check,
            height=32,
            fg_color=COLORS["button_secondary"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
        )
        self.system_check_button.grid(row=4, column=0, columnspan=2, sticky="ew", padx=4, pady=(10, 4))
        ctk.CTkLabel(
            body,
            textvariable=self.system_check_status_var,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=COLORS["muted"],
            justify="left",
        ).grid(row=5, column=0, columnspan=2, sticky="w", padx=4, pady=(2, 0))
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
            title="Select video file",
            filetypes=[(UI_TEXT["file_types"], "*.mp4 *.mov *.mkv *.webm"), ("All files", "*.*")],
        )
        if not file_path:
            return
        path = Path(file_path)
        if path.suffix.lower() not in VIDEO_EXTENSIONS:
            messagebox.showwarning(APP_NAME, "Supported files: mp4, mov, mkv, webm")
            return
        self.selected_file = path
        self.file_var.set(str(path))
        self.output_var.set("No output yet")
        self.current_project = None
        self.current_media_info = None
        self.open_button.configure(state="disabled")
        self._log(LOG_TEXT["source_detected"])
        self._probe_selected_video()

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
        if hasattr(self, "system_check_button"):
            self.system_check_button.configure(state="disabled")
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
            messagebox.showerror(APP_NAME, f"Could not open output folder.\n{exc}")

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
            )
        elif event_type == "cli_error":
            self.system_check_status_var.set(f"System check failed: {event.get('message', '')}")
            if hasattr(self, "system_check_button"):
                self.system_check_button.configure(state="normal")
            self._log(str(event.get("message", "")))
        elif event_type == "media":
            info = event.get("info")
            if isinstance(info, MediaInfo):
                self.current_media_info = info
                set_textbox(self.media_box, self._media_info_text(info))
                self._set_eta(estimate_processing_seconds(info.duration, is_faster_whisper_available()))
        elif event_type == "media_error":
            set_textbox(self.media_box, str(event.get("message", "Media information unavailable.")))
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

    def _apply_cli_status(self, statuses: object, nvenc: object, gpu: object, install_guide: str = "") -> None:
        if isinstance(statuses, dict):
            self.cli_status = statuses  # type: ignore[assignment]
        if isinstance(nvenc, dict):
            self.nvenc_status = nvenc  # type: ignore[assignment]
        if isinstance(gpu, dict):
            self.gpu_status = gpu  # type: ignore[assignment]

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
        if hasattr(self, "system_check_button"):
            self.system_check_button.configure(state="normal")
        self._log(LOG_TEXT["system_check_complete"])
        if self._has_system_issues():
            self._log(LOG_TEXT["tools_missing"])
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
        "faster_whisper": is_faster_whisper_available(),
    }
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
