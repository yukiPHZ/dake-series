# -*- coding: utf-8 -*-
from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from app.constants import THEME, UI_TEXT


class SetupDialog(tk.Toplevel):
    def __init__(self, master: tk.Tk, summary_callback) -> None:  # noqa: ANN001
        super().__init__(master)
        self.summary_callback = summary_callback
        self.title(UI_TEXT["privacy_title"])
        self.configure(bg=THEME["background"])
        self.transient(master)
        self.grab_set()
        self.resizable(False, False)
        self.geometry("520x460")
        self._build()

    def _build(self) -> None:
        body = tk.Frame(self, bg=THEME["background"], padx=24, pady=22)
        body.pack(fill="both", expand=True)
        tk.Label(
            body,
            text=UI_TEXT["privacy_title"],
            bg=THEME["background"],
            fg=THEME["text"],
            font=("BIZ UDPGothic", 17, "bold"),
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            body,
            text=UI_TEXT["privacy_message"],
            bg=THEME["background"],
            fg=THEME["muted"],
            justify="left",
            wraplength=460,
        ).pack(fill="x", pady=(12, 18))

        steps = tk.Frame(body, bg=THEME["card"], highlightbackground=THEME["border"], highlightthickness=1, padx=16, pady=14)
        steps.pack(fill="x")
        for key, mark in (
            ("setup_step_mic", "✓"),
            ("setup_step_save", "✓"),
            ("setup_step_transcription", "○"),
            ("setup_step_summary", "○"),
        ):
            tk.Label(
                steps,
                text=f"{mark} {UI_TEXT[key]}",
                bg=THEME["card"],
                fg=THEME["text"] if mark == "✓" else THEME["muted"],
                anchor="w",
            ).pack(fill="x", pady=2)
        tk.Label(steps, text=UI_TEXT["setup_ready"], bg=THEME["card"], fg=THEME["muted"], anchor="w").pack(fill="x", pady=(10, 0))

        tk.Label(
            body,
            text=UI_TEXT["setup_summary_question"],
            bg=THEME["background"],
            fg=THEME["text"],
            font=("BIZ UDPGothic", 10, "bold"),
        ).pack(anchor="w", pady=(18, 8))
        actions = tk.Frame(body, bg=THEME["background"])
        actions.pack(fill="x")
        ttk.Button(actions, text=UI_TEXT["setup_summary_use"], command=lambda: self._finish(True)).pack(side="left")
        ttk.Button(actions, text=UI_TEXT["setup_summary_later"], command=lambda: self._finish(False)).pack(side="left", padx=(10, 0))

    def _finish(self, use_summary: bool) -> None:
        self.summary_callback(use_summary)
        self.grab_release()
        self.destroy()

