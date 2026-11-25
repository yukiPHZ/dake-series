# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import queue
import webbrowser
from datetime import date, datetime
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk
from tkinter import font as tkfont

from app.audio.audio_device_service import AudioDeviceService
from app.audio.capture_service import CaptureEvent, CaptureService
from app.constants import APP_NAME, FONT_CANDIDATES, LINKS, THEME, UI_TEXT, WINDOW_TITLE
from app.core.app_state import CaptureState, SummaryState, TranscriptionState
from app.core.logger import configure_logger
from app.storage.database import VoiceLogDatabase
from app.storage.paths import (
    AppConfig,
    ensure_runtime_paths,
    icon_candidates,
    load_config,
    runtime_paths,
    save_config,
)
from app.storage.recovery import recover_spool_files, recovered_transcript_path
from app.summary.summary_service import SummaryEvent, SummaryService
from app.transcription.worker import TranscriptionEvent, TranscriptionService
from app.tray.tray_service import TrayService
from app.ui.log_viewer import LogViewer
from app.ui.settings_dialog import SettingsDialog
from app.ui.setup_dialog import SetupDialog


CAPTURE_TEXT = {
    CaptureState.INITIALIZING: "status_initializing",
    CaptureState.WAITING: "status_waiting",
    CaptureState.RECORDING: "status_recording",
    CaptureState.PAUSED: "status_paused",
    CaptureState.MIC_ERROR: "status_mic_error",
    CaptureState.STOPPED: "status_stopped",
}


class VoiceLogApp(tk.Tk):
    def __init__(self, start_services: bool = True, smoke_seconds: float | None = None) -> None:
        super().__init__()
        self.title(WINDOW_TITLE)
        self.geometry("940x720")
        self.minsize(860, 700)
        self.configure(bg=THEME["background"])
        self.protocol("WM_DELETE_WINDOW", self.hide_to_tray)
        self.font_family = self._choose_font()
        self._build_styles()
        self._apply_icon()

        initial_paths = runtime_paths()
        self.config = load_config(initial_paths.config_path)
        self.paths = runtime_paths(self.config)
        ensure_runtime_paths(self.paths)
        self.logger = configure_logger(self.paths.logs)
        self.logger.info("app.startup")
        self.db = VoiceLogDatabase(self.paths.state_db)
        self._recover_spool()

        self.capture_events: queue.Queue[CaptureEvent] = queue.Queue()
        self.transcription_events: queue.Queue[TranscriptionEvent] = queue.Queue()
        self.summary_events: queue.Queue[SummaryEvent] = queue.Queue()

        self.capture_state = CaptureState.INITIALIZING
        self.transcription_state = TranscriptionState.IDLE
        self.summary_state = SummaryState.DISABLED if not self.config.summary_enabled else SummaryState.WAITING
        self.active_recording_started_at: datetime | None = None
        self.level_var = tk.DoubleVar(value=0.0)
        self.capture_status_var = tk.StringVar(value=UI_TEXT["status_initializing"])
        self.capture_detail_var = tk.StringVar(value=UI_TEXT["capture_no_mic"])
        self.processing_status_var = tk.StringVar(value=UI_TEXT["processing_idle"])
        self.recorded_var = tk.StringVar(value="0")
        self.transcribed_var = tk.StringVar(value="0")
        self.pending_var = tk.StringVar(value="0")
        self.summary_var = tk.StringVar(value=UI_TEXT["summary_not_created"])
        self.pause_button_text = tk.StringVar(value=UI_TEXT["button_pause"])
        self.recent_frame: tk.Frame | None = None
        self._quitting = False

        self.transcription_service = TranscriptionService(
            self.db,
            self.paths.save_root,
            self.paths.models,
            self.config.transcription_quality,
            self.transcription_events,
        )
        self.summary_service = SummaryService(self.db, self.paths.save_root, self.summary_events)
        self.capture_service: CaptureService | None = None
        self.tray_service = TrayService(
            icon_candidates(),
            open_app=lambda: self.after(0, self.show_window),
            toggle_pause=lambda: self.after(0, self.toggle_pause),
            is_paused=lambda: self.capture_state == CaptureState.PAUSED,
            open_today=lambda: self.after(0, self.open_today_log),
            open_folder=lambda: self.after(0, self.open_save_folder),
            open_settings=lambda: self.after(0, self.open_settings),
            quit_app=lambda: self.after(0, self.quit_app),
        )

        self._build_ui()
        self._refresh_stats()
        self._start_services(start_services)
        self.tray_service.start()
        self.after(120, self._poll_events)
        self.after(1500, self._periodic_refresh)
        if not self.config.first_run_complete and not smoke_seconds:
            self.after(500, self._show_setup_dialog)
        if smoke_seconds is not None:
            self.after(max(1, int(smoke_seconds * 1000)), self.quit_app)

    def _start_services(self, start_capture: bool) -> None:
        self.transcription_service.start()
        self.summary_service.start(self.config.summary_enabled)
        if start_capture:
            self._restart_capture()
        else:
            self.capture_state = CaptureState.WAITING
            self.capture_status_var.set(UI_TEXT["status_waiting"])

    def _restart_capture(self) -> None:
        if self.capture_service is not None:
            self.capture_service.stop()
        device_service = AudioDeviceService()
        mic_name = self.config.mic_label or device_service.default_input_name()
        self.config.mic_label = mic_name
        self.capture_service = CaptureService(
            self.config,
            self.paths.save_root,
            self.paths.spool,
            self.db,
            self.capture_events,
            on_clip_recorded=lambda _clip_id: self.transcription_service.wake(),
            logger=self.logger,
        )
        self.capture_service.start()

    def _build_styles(self) -> None:
        self.option_add("*Font", f"{{{self.font_family}}} 10")
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TButton", padding=(13, 8))
        style.configure("Primary.TButton", background=THEME["accent"], foreground="#FFFFFF", bordercolor=THEME["accent"], padding=(16, 9))
        style.map("Primary.TButton", background=[("active", THEME["accent_hover"])])
        style.configure("Secondary.TButton", background=THEME["card"], foreground=THEME["text"], bordercolor=THEME["border"], padding=(14, 8))
        style.configure("Horizontal.TProgressbar", troughcolor=THEME["soft"], background=THEME["accent"], bordercolor=THEME["border"])

    def _build_ui(self) -> None:
        outer = tk.Frame(self, bg=THEME["background"], padx=24, pady=20)
        outer.pack(fill="both", expand=True)

        header = tk.Frame(outer, bg=THEME["background"])
        header.pack(fill="x")
        tk.Label(header, text=UI_TEXT["header_title"], bg=THEME["background"], fg=THEME["text"], font=(self.font_family, 22, "bold"), anchor="w").pack(fill="x")
        tk.Label(header, text=UI_TEXT["header_subtitle"], bg=THEME["background"], fg=THEME["muted"], anchor="w").pack(fill="x", pady=(4, 0))

        grid = tk.Frame(outer, bg=THEME["background"])
        grid.pack(fill="both", expand=True, pady=(18, 12))
        grid.columnconfigure(0, weight=1)
        grid.columnconfigure(1, weight=1)
        grid.rowconfigure(1, weight=1)

        self._build_capture_card(grid)
        self._build_today_card(grid)
        self._build_recent_card(grid)
        self._build_actions(grid)
        self._build_footer(outer)

    def _build_capture_card(self, parent: tk.Frame) -> None:
        card = self._card(parent)
        card.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=(0, 14))
        tk.Label(card, text=UI_TEXT["capture_title"], bg=THEME["card"], fg=THEME["text"], font=(self.font_family, 14, "bold"), anchor="w").pack(fill="x")
        row = tk.Frame(card, bg=THEME["card"])
        row.pack(fill="x", pady=(16, 8))
        self.status_dot = tk.Label(row, text="●", bg=THEME["card"], fg=THEME["muted"], font=(self.font_family, 14, "bold"))
        self.status_dot.pack(side="left", padx=(0, 8))
        tk.Label(row, textvariable=self.capture_status_var, bg=THEME["card"], fg=THEME["text"], font=(self.font_family, 15, "bold")).pack(side="left")
        tk.Label(card, textvariable=self.capture_detail_var, bg=THEME["card"], fg=THEME["muted"], anchor="w").pack(fill="x")
        ttk.Progressbar(card, variable=self.level_var, maximum=1.0, style="Horizontal.TProgressbar").pack(fill="x", pady=(16, 4))
        ttk.Button(card, textvariable=self.pause_button_text, style="Primary.TButton", command=self.toggle_pause).pack(anchor="w", pady=(12, 0))

    def _build_today_card(self, parent: tk.Frame) -> None:
        card = self._card(parent)
        card.grid(row=0, column=1, sticky="nsew", padx=(10, 0), pady=(0, 14))
        tk.Label(card, text=UI_TEXT["today_title"], bg=THEME["card"], fg=THEME["text"], font=(self.font_family, 14, "bold"), anchor="w").pack(fill="x")
        for label_key, variable in (
            ("today_recorded", self.recorded_var),
            ("today_transcribed", self.transcribed_var),
            ("today_pending", self.pending_var),
            ("today_summary", self.summary_var),
        ):
            row = tk.Frame(card, bg=THEME["card"])
            row.pack(fill="x", pady=(10, 0))
            tk.Label(row, text=UI_TEXT[label_key], bg=THEME["card"], fg=THEME["muted"], width=14, anchor="w").pack(side="left")
            tk.Label(row, textvariable=variable, bg=THEME["card"], fg=THEME["text"], font=(self.font_family, 11, "bold"), anchor="e").pack(side="right")
        tk.Label(card, textvariable=self.processing_status_var, bg=THEME["card"], fg=THEME["muted"], anchor="w").pack(fill="x", pady=(16, 0))

    def _build_recent_card(self, parent: tk.Frame) -> None:
        card = self._card(parent)
        card.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(0, 14))
        tk.Label(card, text=UI_TEXT["recent_title"], bg=THEME["card"], fg=THEME["text"], font=(self.font_family, 14, "bold"), anchor="w").pack(fill="x")
        self.recent_frame = tk.Frame(card, bg=THEME["card"])
        self.recent_frame.pack(fill="both", expand=True, pady=(10, 0))
        self._render_recent()

    def _build_actions(self, parent: tk.Frame) -> None:
        row = tk.Frame(parent, bg=THEME["background"])
        row.grid(row=2, column=0, columnspan=2, sticky="ew")
        ttk.Button(row, text=UI_TEXT["button_today_log"], style="Secondary.TButton", command=self.open_today_log).pack(side="left")
        ttk.Button(row, text=UI_TEXT["button_open_folder"], style="Secondary.TButton", command=self.open_save_folder).pack(side="left", padx=(10, 0))
        ttk.Button(row, text=UI_TEXT["button_settings"], style="Secondary.TButton", command=self.open_settings).pack(side="left", padx=(10, 0))

    def _build_footer(self, parent: tk.Frame) -> None:
        footer = tk.Frame(parent, bg=THEME["background"])
        footer.pack(fill="x")
        tk.Label(footer, text=UI_TEXT["footer_left"], bg=THEME["background"], fg=THEME["muted"], font=(self.font_family, 8)).pack(side="left")
        right = tk.Frame(footer, bg=THEME["background"])
        right.pack(side="right")
        self._footer_link(right, UI_TEXT["footer_link_1"], LINKS["assessment"])
        tk.Label(right, text=UI_TEXT["footer_separator"], bg=THEME["background"], fg=THEME["muted"], font=(self.font_family, 8)).pack(side="left")
        self._footer_link(right, UI_TEXT["footer_link_2"], LINKS["instagram"])
        tk.Label(right, text=UI_TEXT["footer_separator"], bg=THEME["background"], fg=THEME["muted"], font=(self.font_family, 8)).pack(side="left")
        tk.Label(right, text=UI_TEXT["footer_copyright"], bg=THEME["background"], fg=THEME["muted"], font=(self.font_family, 8)).pack(side="left")

    def _card(self, parent: tk.Frame) -> tk.Frame:
        card = tk.Frame(parent, bg=THEME["card"], highlightbackground=THEME["border"], highlightthickness=1, padx=18, pady=16)
        return card

    def _footer_link(self, parent: tk.Frame, label: str, url: str) -> None:
        link = tk.Label(parent, text=label, bg=THEME["background"], fg=THEME["link"], font=(self.font_family, 8, "bold"), cursor="hand2")
        link.pack(side="left")
        link.bind("<Button-1>", lambda _event: webbrowser.open(url, new=2))
        link.bind("<Enter>", lambda _event: link.configure(fg=THEME["accent"]))
        link.bind("<Leave>", lambda _event: link.configure(fg=THEME["link"]))

    def _poll_events(self) -> None:
        self._poll_capture_events()
        self._poll_transcription_events()
        self._poll_summary_events()
        if not self._quitting:
            self.after(120, self._poll_events)

    def _poll_capture_events(self) -> None:
        while True:
            try:
                event = self.capture_events.get_nowait()
            except queue.Empty:
                break
            if event.kind == "level":
                self.level_var.set(event.level)
                continue
            self.capture_state = event.state
            if event.started_at:
                self.active_recording_started_at = event.started_at
            elif event.state != CaptureState.RECORDING:
                self.active_recording_started_at = None
            self._update_capture_ui()
            self.tray_service.update()

    def _poll_transcription_events(self) -> None:
        while True:
            try:
                event = self.transcription_events.get_nowait()
            except queue.Empty:
                break
            self.transcription_state = event.state
            self._update_processing_ui(event.message)

    def _poll_summary_events(self) -> None:
        while True:
            try:
                event = self.summary_events.get_nowait()
            except queue.Empty:
                break
            self.summary_state = event.state
            self._update_processing_ui(event.message)

    def _periodic_refresh(self) -> None:
        self._refresh_stats()
        self._render_recent()
        backlog = self.db.backlog_count()
        self.summary_service.update_activity(self.capture_state, backlog)
        if not self._quitting:
            self.after(1500, self._periodic_refresh)

    def _refresh_stats(self) -> None:
        stats = self.db.today_stats(date.today())
        self.recorded_var.set(str(stats.recorded))
        self.transcribed_var.set(str(stats.transcribed))
        self.pending_var.set(str(stats.pending))
        self.summary_var.set(UI_TEXT["summary_ready"] if stats.summary_status == "COMPLETE" else UI_TEXT["summary_not_created"])
        self._update_processing_ui()

    def _update_capture_ui(self) -> None:
        self.capture_status_var.set(UI_TEXT[CAPTURE_TEXT[self.capture_state]])
        mic_name = self.config.mic_label or (self.capture_service.mic_name if self.capture_service else "")
        if self.capture_state == CaptureState.RECORDING and self.active_recording_started_at:
            self.capture_detail_var.set(UI_TEXT["capture_since"].format(time=f"{self.active_recording_started_at:%H:%M}"))
        else:
            self.capture_detail_var.set(UI_TEXT["capture_mic"].format(name=mic_name) if mic_name else UI_TEXT["capture_no_mic"])
        self.pause_button_text.set(UI_TEXT["button_resume"] if self.capture_state == CaptureState.PAUSED else UI_TEXT["button_pause"])
        color = {
            CaptureState.RECORDING: THEME["error"],
            CaptureState.WAITING: THEME["success"],
            CaptureState.PAUSED: THEME["warning"],
            CaptureState.MIC_ERROR: THEME["error"],
        }.get(self.capture_state, THEME["muted"])
        self.status_dot.configure(fg=color)

    def _update_processing_ui(self, message: str = "") -> None:
        if self.transcription_state == TranscriptionState.PROCESSING:
            self.processing_status_var.set(UI_TEXT["processing_transcribing"])
        elif self.summary_state == SummaryState.PROCESSING:
            self.processing_status_var.set(UI_TEXT["processing_summarizing"])
        elif message:
            self.processing_status_var.set(message)
        elif self.db.backlog_count() > 0:
            self.processing_status_var.set(UI_TEXT["processing_transcribing"])
        else:
            self.processing_status_var.set(UI_TEXT["processing_idle"])

    def _render_recent(self) -> None:
        if self.recent_frame is None:
            return
        for child in self.recent_frame.winfo_children():
            child.destroy()
        clips = self.db.recent_clips(5)
        if not clips:
            tk.Label(self.recent_frame, text=UI_TEXT["recent_empty"], bg=THEME["card"], fg=THEME["muted"], anchor="w").pack(fill="x", pady=6)
            return
        for clip in clips:
            start = datetime.fromisoformat(clip.started_at)
            end = datetime.fromisoformat(clip.ended_at)
            status = UI_TEXT["recent_done"] if clip.transcription_status == "TRANSCRIBED" else UI_TEXT["recent_processing"] if clip.transcription_status in {"RECORDED", "QUEUED", "TRANSCRIBING"} else UI_TEXT["recent_failed"]
            row = tk.Frame(self.recent_frame, bg=THEME["card"])
            row.pack(fill="x", pady=5)
            tk.Label(row, text=f"{start:%H:%M}–{end:%H:%M}", bg=THEME["card"], fg=THEME["text"], width=14, anchor="w").pack(side="left")
            tk.Label(row, text=status, bg=THEME["card"], fg=THEME["muted"], anchor="w").pack(side="left")

    def toggle_pause(self) -> None:
        if self.capture_service is None:
            return
        if self.capture_state == CaptureState.PAUSED:
            self.capture_service.resume()
        else:
            self.capture_service.pause()

    def open_today_log(self) -> None:
        self.show_window()
        LogViewer(self, self.db, date.today())

    def open_save_folder(self) -> None:
        self.paths.save_root.mkdir(parents=True, exist_ok=True)
        if os.name == "nt":
            os.startfile(str(self.paths.save_root))  # type: ignore[attr-defined]
        else:
            webbrowser.open(self.paths.save_root.resolve().as_uri())

    def open_settings(self) -> None:
        self.show_window()
        SettingsDialog(self, self.config, self.save_settings, self.summarize_pending)

    def save_settings(self, config: AppConfig) -> None:
        self.config = config
        self.paths = runtime_paths(self.config)
        ensure_runtime_paths(self.paths)
        save_config(self.paths.config_path, self.config)
        self.transcription_service.set_quality(config.transcription_quality)
        self.summary_service.set_enabled(config.summary_enabled)
        self._restart_capture()
        self._refresh_stats()

    def summarize_pending(self) -> None:
        self.summary_service.set_enabled(True)
        self.summary_service.wake()

    def hide_to_tray(self) -> None:
        if self._quitting:
            return
        if self.tray_service.available:
            self.withdraw()
            if not self.config.tray_notice_shown:
                messagebox.showinfo(UI_TEXT["dialog_info_title"], UI_TEXT["dialog_tray_notice"], parent=self)
                self.config.tray_notice_shown = True
                save_config(self.paths.config_path, self.config)
        else:
            self.iconify()

    def show_window(self) -> None:
        self.deiconify()
        self.lift()
        self.focus_force()

    def quit_app(self) -> None:
        self._quitting = True
        try:
            if self.capture_service:
                self.capture_service.stop()
            self.transcription_service.stop()
            self.summary_service.stop()
            self.tray_service.stop()
        finally:
            self.logger.info("app.shutdown")
            self.destroy()

    def _show_setup_dialog(self) -> None:
        def finish(use_summary: bool) -> None:
            self.config.first_run_complete = True
            self.config.summary_enabled = use_summary
            save_config(self.paths.config_path, self.config)
            self.summary_service.set_enabled(use_summary)

        SetupDialog(self, finish)

    def _recover_spool(self) -> None:
        recovered = recover_spool_files(self.paths.spool, self.paths.save_root)
        for item in recovered:
            transcript_path = recovered_transcript_path(self.paths.save_root, item.wav_path, item.started_at)
            self.db.create_clip(item.session_id, item.started_at, item.ended_at, item.wav_path, transcript_path)
        if recovered:
            self.logger.info("audio.recovered count=%s", len(recovered))
            self.db.queue_recorded_clips()

    def _apply_icon(self) -> None:
        for path in icon_candidates():
            try:
                if path.is_file():
                    self.iconbitmap(str(path))
                    return
            except Exception:
                continue

    def _choose_font(self) -> str:
        available = set(tkfont.families(self))
        for family in FONT_CANDIDATES:
            if family in available:
                return family
        return "TkDefaultFont"
