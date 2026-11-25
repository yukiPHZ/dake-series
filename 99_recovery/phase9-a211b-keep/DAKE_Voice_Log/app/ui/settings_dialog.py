# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from app.audio.audio_device_service import AudioDeviceService
from app.constants import APP_NAME, OLLAMA_DOWNLOAD_URL, SUMMARY_MODEL, THEME, UI_TEXT
from app.storage.paths import AppConfig
from app.summary.ollama_client import OllamaClient


class SettingsDialog(tk.Toplevel):
    def __init__(
        self,
        master: tk.Tk,
        config: AppConfig,
        save_callback,
        summarize_callback,
    ) -> None:  # noqa: ANN001
        super().__init__(master)
        self.config = config
        self.save_callback = save_callback
        self.summarize_callback = summarize_callback
        self.devices = AudioDeviceService().input_devices()
        self.mic_values = [UI_TEXT["settings_default_mic"], *[device.name for device in self.devices]]
        self.title(UI_TEXT["settings_title"])
        self.configure(bg=THEME["background"])
        self.geometry("620x620")
        self.resizable(False, False)
        self._build()

    def _build(self) -> None:
        outer = tk.Frame(self, bg=THEME["background"], padx=22, pady=18)
        outer.pack(fill="both", expand=True)
        tk.Label(outer, text=UI_TEXT["settings_title"], bg=THEME["background"], fg=THEME["text"], font=("BIZ UDPGothic", 18, "bold")).pack(anchor="w")

        self.mic_var = tk.StringVar(value=self.config.mic_label or UI_TEXT["settings_default_mic"])
        self.sensitivity_var = tk.StringVar(value=UI_TEXT[f"sensitivity_{self.config.sensitivity}"])
        self.silence_var = tk.StringVar(value=UI_TEXT[f"silence_{self.config.silence_timeout_sec}"])
        self.quality_var = tk.StringVar(value=UI_TEXT[f"quality_{self.config.transcription_quality}"])
        self.summary_var = tk.BooleanVar(value=self.config.summary_enabled)
        self.save_root_var = tk.StringVar(value=self.config.save_root)
        self.auto_start_var = tk.BooleanVar(value=self.config.auto_start_windows)
        self.auto_delete_var = tk.StringVar(value=self._auto_delete_label(self.config.auto_delete_days))
        self.ollama_status_var = tk.StringVar(value=UI_TEXT["model_checking"])

        self._section(outer, UI_TEXT["settings_audio"])
        self._combo(outer, UI_TEXT["settings_mic"], self.mic_var, self.mic_values)
        self._combo(outer, UI_TEXT["settings_sensitivity"], self.sensitivity_var, [UI_TEXT["sensitivity_low"], UI_TEXT["sensitivity_standard"], UI_TEXT["sensitivity_high"]])
        self._combo(outer, UI_TEXT["settings_silence"], self.silence_var, [UI_TEXT["silence_10"], UI_TEXT["silence_30"], UI_TEXT["silence_60"]])

        self._section(outer, UI_TEXT["settings_transcription"])
        self._combo(outer, UI_TEXT["settings_quality"], self.quality_var, [UI_TEXT["quality_light"], UI_TEXT["quality_standard"]])
        self._field_label(outer, UI_TEXT["settings_model_status"], UI_TEXT["model_missing"])

        self._section(outer, UI_TEXT["settings_summary"])
        ttk.Checkbutton(outer, text=UI_TEXT["settings_summary_enabled"], variable=self.summary_var).pack(anchor="w", pady=2)
        self._field_value_var(outer, UI_TEXT["settings_ai_status"], self.ollama_status_var)
        ai_buttons = tk.Frame(outer, bg=THEME["background"])
        ai_buttons.pack(fill="x", pady=(2, 8))
        ttk.Button(ai_buttons, text=UI_TEXT["button_retry"], command=self.check_ollama).pack(side="left")
        ttk.Button(ai_buttons, text=UI_TEXT["button_open_ollama"], command=self.open_ollama).pack(side="left", padx=(8, 0))
        ttk.Button(ai_buttons, text=UI_TEXT["button_prepare_summary_ai"], command=self.prepare_summary_ai).pack(side="left", padx=(8, 0))
        ttk.Button(ai_buttons, text=UI_TEXT["button_summarize_pending"], command=self.summarize_callback).pack(side="left", padx=(8, 0))

        self._section(outer, UI_TEXT["settings_save"])
        row = tk.Frame(outer, bg=THEME["background"])
        row.pack(fill="x", pady=3)
        tk.Label(row, text=UI_TEXT["settings_save_root"], bg=THEME["background"], fg=THEME["muted"], width=18, anchor="w").pack(side="left")
        ttk.Entry(row, textvariable=self.save_root_var).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text=UI_TEXT["button_select_folder"], command=self.choose_folder).pack(side="left", padx=(8, 0))
        ttk.Checkbutton(outer, text=UI_TEXT["settings_auto_start"], variable=self.auto_start_var).pack(anchor="w", pady=(8, 2))
        self._combo(outer, UI_TEXT["settings_auto_delete"], self.auto_delete_var, [UI_TEXT["auto_delete_off"], UI_TEXT["auto_delete_30"], UI_TEXT["auto_delete_90"]])

        bottom = tk.Frame(outer, bg=THEME["background"])
        bottom.pack(fill="x", pady=(18, 0))
        ttk.Button(bottom, text=UI_TEXT["button_save"], command=self.save).pack(side="right")
        ttk.Button(bottom, text=UI_TEXT["button_cancel"], command=self.destroy).pack(side="right", padx=(0, 8))
        self.after(250, self.check_ollama)

    def _section(self, parent: tk.Frame, title: str) -> None:
        tk.Label(parent, text=title, bg=THEME["background"], fg=THEME["text"], font=("BIZ UDPGothic", 11, "bold")).pack(anchor="w", pady=(14, 4))

    def _combo(self, parent: tk.Frame, label: str, variable: tk.StringVar, values: list[str]) -> None:
        row = tk.Frame(parent, bg=THEME["background"])
        row.pack(fill="x", pady=3)
        tk.Label(row, text=label, bg=THEME["background"], fg=THEME["muted"], width=18, anchor="w").pack(side="left")
        ttk.Combobox(row, textvariable=variable, values=values, state="readonly").pack(side="left", fill="x", expand=True)

    def _field_label(self, parent: tk.Frame, label: str, value: str) -> None:
        row = tk.Frame(parent, bg=THEME["background"])
        row.pack(fill="x", pady=3)
        tk.Label(row, text=label, bg=THEME["background"], fg=THEME["muted"], width=18, anchor="w").pack(side="left")
        tk.Label(row, text=value, bg=THEME["background"], fg=THEME["text"], anchor="w").pack(side="left", fill="x", expand=True)

    def _field_value_var(self, parent: tk.Frame, label: str, variable: tk.StringVar) -> None:
        row = tk.Frame(parent, bg=THEME["background"])
        row.pack(fill="x", pady=3)
        tk.Label(row, text=label, bg=THEME["background"], fg=THEME["muted"], width=18, anchor="w").pack(side="left")
        tk.Label(row, textvariable=variable, bg=THEME["background"], fg=THEME["text"], anchor="w").pack(side="left", fill="x", expand=True)

    def choose_folder(self) -> None:
        path = filedialog.askdirectory(title=UI_TEXT["settings_save_root"], initialdir=self.save_root_var.get() or str(Path.home()))
        if path:
            self.save_root_var.set(path)

    def check_ollama(self) -> None:
        status = OllamaClient().status()
        if not status.reachable:
            self.ollama_status_var.set(UI_TEXT["status_ollama_missing"])
        elif SUMMARY_MODEL in status.models:
            self.ollama_status_var.set(UI_TEXT["status_ollama_ready"])
        else:
            self.ollama_status_var.set(UI_TEXT["status_ollama_model_missing"])

    @staticmethod
    def open_ollama() -> None:
        import webbrowser

        webbrowser.open(OLLAMA_DOWNLOAD_URL, new=2)

    def prepare_summary_ai(self) -> None:
        if not messagebox.askyesno(UI_TEXT["ollama_setup_title"], UI_TEXT["ollama_model_pull_confirm"], parent=self):
            return
        self.ollama_status_var.set(UI_TEXT["ollama_model_pulling"])

        def worker() -> None:
            try:
                OllamaClient().pull_model(SUMMARY_MODEL, progress=lambda status: self.after(0, self.ollama_status_var.set, status))
                self.after(0, self.ollama_status_var.set, UI_TEXT["status_ollama_ready"])
            except Exception as exc:
                self.after(0, self.ollama_status_var.set, str(exc))

        threading.Thread(target=worker, daemon=True).start()

    def save(self) -> None:
        mic_label = self.mic_var.get()
        mic_device = ""
        if mic_label != UI_TEXT["settings_default_mic"]:
            for device in self.devices:
                if device.name == mic_label:
                    mic_device = device.id
                    break
        config = AppConfig(
            save_root=self.save_root_var.get().strip() or self.config.save_root,
            mic_device=mic_device,
            mic_label="" if mic_label == UI_TEXT["settings_default_mic"] else mic_label,
            sensitivity=self._reverse("sensitivity", self.sensitivity_var.get()),
            silence_timeout_sec=self._silence_value(self.silence_var.get()),
            transcription_quality=self._reverse("quality", self.quality_var.get()),
            summary_enabled=bool(self.summary_var.get()),
            auto_start_windows=bool(self.auto_start_var.get()),
            auto_delete_days=self._auto_delete_value(self.auto_delete_var.get()),
            first_run_complete=True,
            tray_notice_shown=self.config.tray_notice_shown,
        )
        try:
            set_windows_startup(config.auto_start_windows)
        except Exception as exc:
            messagebox.showwarning(UI_TEXT["dialog_error_title"], str(exc), parent=self)
        self.save_callback(config)
        self.destroy()

    @staticmethod
    def _reverse(prefix: str, label: str) -> str:
        for value in ("low", "standard", "high", "light"):
            key = f"{prefix}_{value}"
            if UI_TEXT.get(key) == label:
                return value
        return "standard"

    @staticmethod
    def _silence_value(label: str) -> int:
        if label == UI_TEXT["silence_10"]:
            return 10
        if label == UI_TEXT["silence_60"]:
            return 60
        return 30

    @staticmethod
    def _auto_delete_label(days: int) -> str:
        if days == 30:
            return UI_TEXT["auto_delete_30"]
        if days == 90:
            return UI_TEXT["auto_delete_90"]
        return UI_TEXT["auto_delete_off"]

    @staticmethod
    def _auto_delete_value(label: str) -> int:
        if label == UI_TEXT["auto_delete_30"]:
            return 30
        if label == UI_TEXT["auto_delete_90"]:
            return 90
        return 0


def set_windows_startup(enabled: bool) -> None:
    if os.name != "nt":
        return
    import winreg

    key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
        if enabled:
            exe = sys.executable if getattr(sys, "frozen", False) else str(Path(__file__).resolve().parents[2] / "main.py")
            winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, f'"{exe}"')
        else:
            try:
                winreg.DeleteValue(key, APP_NAME)
            except FileNotFoundError:
                pass
