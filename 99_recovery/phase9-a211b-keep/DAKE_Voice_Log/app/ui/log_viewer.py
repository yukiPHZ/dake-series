# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import webbrowser
from datetime import date, datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk

from app.constants import THEME, UI_TEXT
from app.storage.database import ClipRecord, VoiceLogDatabase
from app.storage.daily_builder import format_time_range


class LogViewer(tk.Toplevel):
    def __init__(self, master: tk.Tk, db: VoiceLogDatabase, target_date: date | None = None) -> None:
        super().__init__(master)
        self.db = db
        self.date_var = tk.StringVar(value=(target_date or date.today()).isoformat())
        self.clips: list[ClipRecord] = []
        self.title(UI_TEXT["log_viewer_title"])
        self.configure(bg=THEME["background"])
        self.geometry("720x560")
        self.minsize(620, 480)
        self._build()
        self.refresh()

    def _build(self) -> None:
        outer = tk.Frame(self, bg=THEME["background"], padx=18, pady=16)
        outer.pack(fill="both", expand=True)
        top = tk.Frame(outer, bg=THEME["background"])
        top.pack(fill="x")
        tk.Label(top, text=UI_TEXT["log_viewer_date"], bg=THEME["background"], fg=THEME["muted"]).pack(side="left")
        ttk.Entry(top, textvariable=self.date_var, width=16).pack(side="left", padx=(8, 8))
        ttk.Button(top, text=UI_TEXT["button_retry"], command=self.refresh).pack(side="left")

        self.listbox = tk.Listbox(outer, activestyle="none", exportselection=False)
        self.listbox.pack(fill="both", expand=True, pady=(14, 10))
        self.listbox.bind("<<ListboxSelect>>", lambda _event: self._update_buttons())

        actions = tk.Frame(outer, bg=THEME["background"])
        actions.pack(fill="x")
        self.text_button = ttk.Button(actions, text=UI_TEXT["button_open_text"], command=self.open_text)
        self.text_button.pack(side="left")
        self.audio_button = ttk.Button(actions, text=UI_TEXT["button_open_audio"], command=self.open_audio)
        self.audio_button.pack(side="left", padx=(8, 0))
        ttk.Button(actions, text=UI_TEXT["button_close"], command=self.destroy).pack(side="right")

    def refresh(self) -> None:
        self.listbox.delete(0, "end")
        try:
            selected_date = self.date_var.get().strip()
            self.clips = self.db.clips_for_date(selected_date, include_no_speech=True)
        except Exception:
            self.clips = []
        if not self.clips:
            self.listbox.insert("end", UI_TEXT["log_viewer_empty"])
        else:
            for clip in self.clips:
                self.listbox.insert("end", f"{format_time_range(clip)}  {clip.transcription_status}")
        self._update_buttons()

    def selected_clip(self) -> ClipRecord | None:
        selection = self.listbox.curselection()
        if not selection or not self.clips:
            return None
        index = int(selection[0])
        if index >= len(self.clips):
            return None
        return self.clips[index]

    def open_text(self) -> None:
        clip = self.selected_clip()
        if clip is None:
            return
        self._open_path(Path(clip.transcript_path))

    def open_audio(self) -> None:
        clip = self.selected_clip()
        if clip is None:
            return
        self._open_path(Path(clip.audio_path))

    def _update_buttons(self) -> None:
        state = tk.NORMAL if self.selected_clip() else tk.DISABLED
        self.text_button.configure(state=state)
        self.audio_button.configure(state=state)

    @staticmethod
    def _open_path(path: Path) -> None:
        if os.name == "nt" and path.exists():
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif path.exists():
            webbrowser.open(path.resolve().as_uri())

