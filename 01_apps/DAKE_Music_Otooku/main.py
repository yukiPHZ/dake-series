# -*- coding: utf-8 -*-
from __future__ import annotations

import ctypes
import os
import queue
import sys
import tempfile
import threading
import time
import webbrowser
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox

from core.app_config import DEFAULT_DURATION_SECONDS, ensure_data_dirs
from core.audio_probe import probe_audio_info
from core.cli_checker import EnvironmentReport, check_environment
from core.ffmpeg_audio import export_audio_material
from core.musicgen_runner import generate_music
from core.ollama_client import generate_direction
from core.project_writer import (
    create_project,
    make_project_name,
    write_project_files,
    write_setup_needed,
)
from core.prompt_builder import fallback_direction
from ui.components import make_button, make_panel
from ui.theme import APP_USER_MODEL_ID, COLORS, FONT_CANDIDATES, WINDOW_MIN_SIZE, WINDOW_SIZE


APP_NAME = "音を置く"
WINDOW_TITLE = "音を置く"
COPYRIGHT = "© 2026 しまりす不動産 / Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "main_title": "音を置く",
    "main_description": "言葉から、動画や配信用の小さな音素材を作ります。",
    "prompt_label": "PROMPT",
    "prompt_hint": "例：深夜、コード、ミシン、静かな稼働",
    "output_label": "OUTPUT / STATUS",
    "system_label": "SYSTEM",
    "log_label": "補助脳ログ",
    "button_place_sound": "音を置く",
    "button_reference": "参照音源を選ぶ",
    "button_open_output": "出力フォルダを開く",
    "button_check": "環境チェック",
    "button_clear_reference": "参照を解除",
    "reference_none": "参照音源: なし",
    "reference_selected": "参照音源: {name}",
    "reference_probe_waiting": "file name: --\nduration: --\ncodec: --\nsample rate: --\nchannels: --",
    "reference_probe_running": "ffprobe info: checking...",
    "reference_probe_failed": "ffprobe info: unavailable",
    "output_none": "出力先: 未作成",
    "output_ready": "出力先: {path}",
    "status_idle": "稼働中。",
    "status_checking": "環境を確認しています。",
    "status_processing": "音を置いています。",
    "status_complete": "整っています。",
    "status_ffmpeg_processing": "FFmpeg processing...",
    "status_audio_complete": "整いました。 Audio package created.",
    "status_ffmpeg_failed": "FFmpeg failed. Prompt output is still available.",
    "status_error": "処理を止めました。",
    "status_no_prompt": "言葉を入力してください。",
    "env_unknown": "CHECK WAITING",
    "status_ffmpeg_offline_hint": "OFFLINE - FFmpeg is offline. Prompt output is still available.",
    "status_ffprobe_offline_hint": "OFFLINE - FFprobe is offline. Prompt output is still available.",
    "local_brain_label": "LOCAL BRAIN RESPONSE",
    "local_brain_idle": "Response Time: --",
    "local_brain_measuring": "Response Time: measuring...",
    "local_brain_online": "Response Time: {seconds:.1f}s",
    "local_brain_offline": "LOCAL BRAIN OFFLINE / Fallback direction generated.",
    "status_tool_ffmpeg": "FFMPEG",
    "status_tool_ffprobe": "FFPROBE",
    "status_tool_ollama": "OLLAMA",
    "status_tool_musicgen": "MUSICGEN",
    "status_tool_cuda": "CUDA",
    "status_tool_uvr": "UVR",
    "log_ready": "補助脳：稼働中です。",
    "log_check_start": "SYSTEM：環境チェックを開始しました。",
    "log_check_done": "SYSTEM：環境チェックが完了しました。",
    "log_brain_received": "補助脳：言葉を受け取りました。",
    "log_brain_thinking": "補助脳：空気を解析しています。",
    "log_brain_low_temp": "補助脳：静かな低温構成を提案しました。",
    "log_brain_arranged": "補助脳：音の置き方を整えました。",
    "log_brain_start": "補助脳：音の方向性を考えています。",
    "log_brain_ollama": "補助脳：Ollamaで音設計を作成しました。",
    "log_brain_template": "補助脳：固定テンプレートで音設計を作成しました。",
    "log_bpm": "補助脳：BPMを{bpm}に設定しました。",
    "log_project": "SYSTEM：出力フォルダを作成しました。",
    "log_musicgen_start": "MUSICGEN：短い音源生成を試します。",
    "log_musicgen_unavailable": "MUSICGEN：未導入のためプロンプトと設計メモを保存しました。",
    "log_musicgen_done": "MUSICGEN：generated.wavを作成しました。",
    "log_musicgen_failed": "MUSICGEN：生成できなかったためsetup_needed.txtを保存しました。",
    "log_reference": "SYSTEM：参照音源を素材化します。",
    "log_reference_probe": "FFPROBE：{info}",
    "log_ffmpeg_start": "FFMPEG：音量調整と変換を開始しました。",
    "log_ffmpeg_done": "FFMPEG：wav/mp3/loop previewを作成しました。",
    "log_ffmpeg_file": "FFMPEG：{name} を作成しました。",
    "log_ffmpeg_package": "FFMPEG：Audio package created.",
    "log_ffmpeg_failed": "FFMPEG：FFmpeg failed. Prompt output is still available.",
    "log_ffmpeg_missing": "FFMPEG：FFmpeg is required for audio export",
    "log_complete": "整っています。",
    "dialog_error_title": "確認",
    "dialog_open_error": "出力フォルダを開けませんでした。",
    "dialog_reference_title": "参照音源を選ぶ",
    "filetype_audio": "音声ファイル",
    "filetype_all": "すべてのファイル",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_tagline": "止まらない、迷わない、すぐ終わる。",
    "footer_separator": " ・ ",
    "footer_copyright": COPYRIGHT,
}

LINK_URLS = {
    "footer_link_1": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "footer_link_2": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

STATUS_KEYS = ("ffmpeg", "ffprobe", "ollama", "musicgen", "cuda", "uvr")
STATUS_LABEL_KEYS = {
    "ffmpeg": "status_tool_ffmpeg",
    "ffprobe": "status_tool_ffprobe",
    "ollama": "status_tool_ollama",
    "musicgen": "status_tool_musicgen",
    "cuda": "status_tool_cuda",
    "uvr": "status_tool_uvr",
}
AUDIO_FILETYPES = (
    ("Audio", "*.wav *.mp3 *.m4a *.flac *.ogg"),
    ("All", "*.*"),
)


def set_windows_app_id() -> None:
    if os.name != "nt":
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)
    except Exception:
        return


def open_output_folder(path: Path, dry_run: bool = False) -> bool:
    if dry_run:
        return path.exists()
    try:
        if os.name == "nt":
            os.startfile(str(path))
        else:
            webbrowser.open(path.as_uri())
        return True
    except Exception:
        return False


class MusicOtookuApp:
    def __init__(self, root: tk.Tk) -> None:
        ensure_data_dirs()
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(*WINDOW_MIN_SIZE)
        self._apply_icon()

        self.font_family = self._font_family()
        self.fonts = {
            "title": (self.font_family, 24, "bold"),
            "description": (self.font_family, 11),
            "label": (self.font_family, 10, "bold"),
            "body": (self.font_family, 11),
            "button": (self.font_family, 10, "bold"),
            "status": (self.font_family, 10),
            "small": (self.font_family, 9),
            "mono": ("Consolas", 9),
        }

        self.worker_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.processing = False
        self.environment_report: EnvironmentReport | None = None
        self.reference_audio_path: Path | None = None
        self.output_folder: Path | None = None
        self.log_lines: list[str] = []
        self.status_vars: dict[str, tk.StringVar] = {}

        self.status_var = tk.StringVar(value=UI_TEXT["status_idle"])
        self.reference_var = tk.StringVar(value=UI_TEXT["reference_none"])
        self.reference_info_var = tk.StringVar(value=UI_TEXT["reference_probe_waiting"])
        self.output_var = tk.StringVar(value=UI_TEXT["output_none"])
        self.brain_response_var = tk.StringVar(value=UI_TEXT["local_brain_idle"])

        self._build_ui()
        self._append_log(UI_TEXT["log_ready"])
        self._start_environment_check()
        self.root.after(100, self._poll_queue)

    def _apply_icon(self) -> None:
        icon_path = Path(__file__).resolve().parent / ".." / ".." / "02_assets" / "dake_icon.ico"
        try:
            if icon_path.exists():
                self.root.iconbitmap(str(icon_path))
        except Exception:
            return

    def _font_family(self) -> str:
        available = set(tkfont.families())
        for candidate in FONT_CANDIDATES:
            if candidate in available:
                return candidate
        return "TkDefaultFont"

    def _build_ui(self) -> None:
        outer = tk.Frame(self.root, bg=COLORS["background"])
        outer.pack(fill="both", expand=True, padx=24, pady=(22, 16))

        self._build_header(outer)
        self._build_body(outer)
        self._build_system(outer)
        self._build_footer(outer)

    def _build_header(self, parent: tk.Widget) -> None:
        header = tk.Frame(parent, bg=COLORS["background"])
        header.pack(fill="x", pady=(0, 14))
        header.grid_columnconfigure(0, weight=1)

        title_area = tk.Frame(header, bg=COLORS["background"])
        title_area.grid(row=0, column=0, sticky="w")

        tk.Label(
            title_area,
            text=UI_TEXT["main_title"],
            bg=COLORS["background"],
            fg=COLORS["text"],
            font=self.fonts["title"],
        ).pack(anchor="w")
        tk.Label(
            title_area,
            text=UI_TEXT["main_description"],
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=self.fonts["description"],
        ).pack(anchor="w", pady=(4, 0))

        action_area = tk.Frame(header, bg=COLORS["background"])
        action_area.grid(row=0, column=1, sticky="e", padx=(16, 0))
        self.check_button = make_button(
            action_area,
            UI_TEXT["button_check"],
            self._start_environment_check,
            COLORS,
            self.fonts["button"],
            primary=False,
        )
        self.check_button.pack(side="left")
        self.execute_button = make_button(
            action_area,
            UI_TEXT["button_place_sound"],
            self.start_place_sound,
            COLORS,
            self.fonts["button"],
            primary=True,
        )
        self.execute_button.pack(side="left", padx=(10, 0))

    def _build_body(self, parent: tk.Widget) -> None:
        body = tk.Frame(parent, bg=COLORS["background"])
        body.pack(fill="both", expand=True)
        body.grid_columnconfigure(0, weight=3, uniform="body")
        body.grid_columnconfigure(1, weight=2, uniform="body")
        body.grid_rowconfigure(0, weight=1)

        prompt_panel = make_panel(body, COLORS)
        prompt_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 14))
        prompt_panel.grid_rowconfigure(1, weight=1)
        prompt_panel.grid_columnconfigure(0, weight=1)

        tk.Label(
            prompt_panel,
            text=UI_TEXT["prompt_label"],
            bg=COLORS["surface"],
            fg=COLORS["text"],
            font=self.fonts["label"],
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 8))

        self.prompt_text = tk.Text(
            prompt_panel,
            height=10,
            bg=COLORS["surface_soft"],
            fg=COLORS["text"],
            insertbackground=COLORS["text"],
            selectbackground=COLORS["select"],
            font=self.fonts["body"],
            wrap="word",
            relief="flat",
            bd=0,
            padx=14,
            pady=12,
        )
        self.prompt_text.grid(row=1, column=0, sticky="nsew", padx=16, pady=(0, 12))
        self.prompt_text.insert("1.0", UI_TEXT["prompt_hint"])
        self.prompt_text.bind("<FocusIn>", self._clear_prompt_hint)

        control_row = tk.Frame(prompt_panel, bg=COLORS["surface"])
        control_row.grid(row=2, column=0, sticky="ew", padx=16, pady=(0, 14))

        self.reference_button = make_button(
            control_row,
            UI_TEXT["button_reference"],
            self.select_reference_audio,
            COLORS,
            self.fonts["button"],
            primary=False,
        )
        self.reference_button.pack(side="left")
        self.clear_reference_button = make_button(
            control_row,
            UI_TEXT["button_clear_reference"],
            self.clear_reference_audio,
            COLORS,
            self.fonts["button"],
            primary=False,
        )
        self.clear_reference_button.pack(side="left", padx=(10, 0))

        tk.Label(
            prompt_panel,
            textvariable=self.reference_var,
            bg=COLORS["surface"],
            fg=COLORS["muted"],
            font=self.fonts["small"],
            anchor="w",
        ).grid(row=3, column=0, sticky="ew", padx=16, pady=(0, 14))

        tk.Label(
            prompt_panel,
            textvariable=self.reference_info_var,
            bg=COLORS["surface_soft"],
            fg=COLORS["muted"],
            font=self.fonts["mono"],
            anchor="w",
            justify="left",
            padx=10,
            pady=8,
        ).grid(row=4, column=0, sticky="ew", padx=16, pady=(0, 14))

        output_panel = make_panel(body, COLORS)
        output_panel.grid(row=0, column=1, sticky="nsew")
        output_panel.grid_columnconfigure(0, weight=1)

        tk.Label(
            output_panel,
            text=UI_TEXT["output_label"],
            bg=COLORS["surface"],
            fg=COLORS["text"],
            font=self.fonts["label"],
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 8))

        self.status_label = tk.Label(
            output_panel,
            textvariable=self.status_var,
            bg=COLORS["status_bg"],
            fg=COLORS["text"],
            font=self.fonts["status"],
            anchor="w",
            padx=12,
            pady=8,
        )
        self.status_label.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 12))

        status_list = tk.Frame(output_panel, bg=COLORS["surface"])
        status_list.grid(row=2, column=0, sticky="ew", padx=16)
        status_list.grid_columnconfigure(1, weight=1)
        for row, key in enumerate(STATUS_KEYS):
            tk.Label(
                status_list,
                text=UI_TEXT[STATUS_LABEL_KEYS[key]],
                bg=COLORS["surface"],
                fg=COLORS["muted"],
                font=self.fonts["small"],
                anchor="w",
            ).grid(row=row, column=0, sticky="w", pady=3)
            var = tk.StringVar(value=UI_TEXT["env_unknown"])
            self.status_vars[key] = var
            tk.Label(
                status_list,
                textvariable=var,
                bg=COLORS["surface"],
                fg=COLORS["text"],
                font=self.fonts["mono"],
                anchor="e",
                justify="right",
                wraplength=260,
            ).grid(row=row, column=1, sticky="e", pady=3)

        tk.Label(
            output_panel,
            textvariable=self.output_var,
            bg=COLORS["surface"],
            fg=COLORS["muted"],
            font=self.fonts["small"],
            anchor="w",
            wraplength=360,
            justify="left",
        ).grid(row=3, column=0, sticky="ew", padx=16, pady=(18, 10))

        self.open_button = make_button(
            output_panel,
            UI_TEXT["button_open_output"],
            self.open_current_output,
            COLORS,
            self.fonts["button"],
            primary=False,
        )
        self.open_button.grid(row=4, column=0, sticky="ew", padx=16, pady=(0, 14))
        self.open_button.configure(state="disabled")

    def _build_system(self, parent: tk.Widget) -> None:
        system_panel = make_panel(parent, COLORS)
        system_panel.pack(fill="x", pady=(14, 0))
        system_panel.grid_columnconfigure(0, weight=1)

        header = tk.Frame(system_panel, bg=COLORS["surface"])
        header.grid(row=0, column=0, sticky="ew", padx=16, pady=(12, 6))
        tk.Label(
            header,
            text=UI_TEXT["system_label"],
            bg=COLORS["surface"],
            fg=COLORS["text"],
            font=self.fonts["label"],
        ).pack(side="left")
        tk.Label(
            header,
            text=UI_TEXT["log_label"],
            bg=COLORS["surface"],
            fg=COLORS["muted"],
            font=self.fonts["small"],
        ).pack(side="left", padx=(12, 0))
        tk.Label(
            header,
            textvariable=self.brain_response_var,
            bg=COLORS["surface"],
            fg=COLORS["muted"],
            font=self.fonts["small"],
        ).pack(side="right")
        tk.Label(
            header,
            text=UI_TEXT["local_brain_label"],
            bg=COLORS["surface"],
            fg=COLORS["quiet"],
            font=self.fonts["small"],
        ).pack(side="right", padx=(0, 12))

        self.log_text = tk.Text(
            system_panel,
            height=5,
            bg=COLORS["surface_soft"],
            fg=COLORS["text"],
            font=self.fonts["small"],
            wrap="word",
            relief="flat",
            bd=0,
            padx=12,
            pady=8,
            state="disabled",
        )
        self.log_text.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 14))

    def _build_footer(self, parent: tk.Widget) -> None:
        footer = tk.Frame(parent, bg=COLORS["background"])
        footer.pack(fill="x", pady=(12, 0))

        tk.Label(
            footer,
            text=UI_TEXT["footer_left"] + UI_TEXT["footer_separator"] + UI_TEXT["footer_tagline"],
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=self.fonts["small"],
        ).pack(side="left")
        tk.Label(
            footer,
            text=UI_TEXT["footer_copyright"],
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=self.fonts["small"],
        ).pack(side="right")

    def _clear_prompt_hint(self, _event: tk.Event) -> None:
        if self.prompt_text.get("1.0", "end").strip() == UI_TEXT["prompt_hint"]:
            self.prompt_text.delete("1.0", "end")

    def _append_log(self, message: str) -> None:
        self.log_lines.append(message)
        if hasattr(self, "log_text"):
            self.log_text.configure(state="normal")
            self.log_text.insert("end", message + "\n")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")

    def _set_busy(self, busy: bool) -> None:
        self.processing = busy
        state = "disabled" if busy else "normal"
        for button in (
            self.execute_button,
            self.check_button,
            self.reference_button,
            self.clear_reference_button,
        ):
            button.configure(state=state)

    def _start_environment_check(self) -> None:
        if self.processing:
            return
        self.status_var.set(UI_TEXT["status_checking"])
        self._append_log(UI_TEXT["log_check_start"])
        thread = threading.Thread(target=self._check_worker, daemon=True)
        thread.start()

    def _check_worker(self) -> None:
        report = check_environment()
        self.worker_queue.put(("environment", report))
        self.worker_queue.put(("log", UI_TEXT["log_check_done"]))
        self.worker_queue.put(("status", UI_TEXT["status_idle"]))

    def _apply_environment_report(self, report: EnvironmentReport) -> None:
        self.environment_report = report
        for key in STATUS_KEYS:
            status = report.status_for(key)
            if status is None:
                self.status_vars[key].set(UI_TEXT["env_unknown"])
                continue
            if key == "ffmpeg" and status.state == "OFFLINE":
                self.status_vars[key].set(UI_TEXT["status_ffmpeg_offline_hint"])
            elif key == "ffprobe" and status.state == "OFFLINE":
                self.status_vars[key].set(UI_TEXT["status_ffprobe_offline_hint"])
            elif status.detail and key in {"ollama", "cuda", "uvr"}:
                self.status_vars[key].set(f"{status.state} - {status.detail}")
            else:
                self.status_vars[key].set(status.state)

    def select_reference_audio(self) -> None:
        if self.processing:
            return
        selected = filedialog.askopenfilename(
            title=UI_TEXT["dialog_reference_title"],
            filetypes=[
                (UI_TEXT["filetype_audio"], AUDIO_FILETYPES[0][1]),
                (UI_TEXT["filetype_all"], AUDIO_FILETYPES[1][1]),
            ],
        )
        if not selected:
            return
        self.reference_audio_path = Path(selected)
        self.reference_var.set(UI_TEXT["reference_selected"].format(name=self.reference_audio_path.name))
        self.reference_info_var.set(UI_TEXT["reference_probe_running"])
        thread = threading.Thread(target=self._probe_reference_worker, args=(self.reference_audio_path,), daemon=True)
        thread.start()

    def clear_reference_audio(self) -> None:
        self.reference_audio_path = None
        self.reference_var.set(UI_TEXT["reference_none"])
        self.reference_info_var.set(UI_TEXT["reference_probe_waiting"])

    def _probe_reference_worker(self, path: Path) -> None:
        info = probe_audio_info(path)
        if info:
            lines = info.display_lines()
            self.worker_queue.put(("reference_info", "\n".join(lines)))
            self.worker_queue.put(("log", UI_TEXT["log_reference_probe"].format(info=" / ".join(lines))))
        else:
            self.worker_queue.put(("reference_info", UI_TEXT["reference_probe_failed"]))

    def start_place_sound(self) -> None:
        if self.processing:
            return
        prompt = self.prompt_text.get("1.0", "end").strip()
        if not prompt or prompt == UI_TEXT["prompt_hint"]:
            self.status_var.set(UI_TEXT["status_no_prompt"])
            return
        self._set_busy(True)
        self.status_var.set(UI_TEXT["status_processing"])
        thread = threading.Thread(target=self._place_sound_worker, args=(prompt,), daemon=True)
        thread.start()

    def _place_sound_worker(self, prompt: str) -> None:
        process_log: list[str] = []
        final_status = UI_TEXT["status_complete"]

        def log(message: str) -> None:
            process_log.append(message)
            self.worker_queue.put(("log", message))

        try:
            log(UI_TEXT["log_brain_received"])
            log(UI_TEXT["log_brain_thinking"])
            self.worker_queue.put(("brain_response", UI_TEXT["local_brain_measuring"]))
            report = self.environment_report or check_environment()
            if self.environment_report is None:
                self.worker_queue.put(("environment", report))

            try:
                started_at = time.perf_counter()
                direction = generate_direction(prompt, report.ollama_models)
                elapsed = time.perf_counter() - started_at
                self.worker_queue.put(("brain_response", UI_TEXT["local_brain_online"].format(seconds=elapsed)))
                log(UI_TEXT["log_brain_ollama"])
                log(UI_TEXT["log_brain_low_temp"])
            except Exception:
                direction = fallback_direction(prompt)
                self.worker_queue.put(("brain_response", UI_TEXT["local_brain_offline"]))
                log(UI_TEXT["log_brain_template"])

            log(UI_TEXT["log_brain_arranged"])
            log(UI_TEXT["log_bpm"].format(bpm=direction.bpm))
            project_name = make_project_name(prompt)
            project_paths = create_project(project_name)
            log(UI_TEXT["log_project"])

            source_audio: Path | None = None
            if self.reference_audio_path:
                source_audio = self.reference_audio_path
                log(UI_TEXT["log_reference"])
                audio_info = probe_audio_info(source_audio)
                if audio_info:
                    log(UI_TEXT["log_reference_probe"].format(info=" / ".join(audio_info.display_lines())))
            else:
                musicgen_status = report.status_for("musicgen")
                if musicgen_status and musicgen_status.state == "IMPORT READY":
                    log(UI_TEXT["log_musicgen_start"])
                    music_result = generate_music(
                        direction.musicgen_prompt,
                        project_paths.audio / "generated.wav",
                        duration_seconds=DEFAULT_DURATION_SECONDS,
                    )
                    if music_result.success and music_result.output_path:
                        source_audio = music_result.output_path
                        log(UI_TEXT["log_musicgen_done"])
                    else:
                        write_setup_needed(project_paths, music_result.message or "MusicGen failed")
                        log(UI_TEXT["log_musicgen_failed"])
                else:
                    write_setup_needed(project_paths, "AudioCraft / MusicGen is not available")
                    log(UI_TEXT["log_musicgen_unavailable"])

            if source_audio:
                ffmpeg_status = report.status_for("ffmpeg")
                if ffmpeg_status and ffmpeg_status.state == "ONLINE":
                    self.worker_queue.put(("status", UI_TEXT["status_ffmpeg_processing"]))
                    log(UI_TEXT["log_ffmpeg_start"])
                    output_stem = "source_converted" if self.reference_audio_path else "generated"
                    export_result = export_audio_material(source_audio, project_paths.audio, output_stem=output_stem)
                    if export_result.success:
                        for file_path in export_result.files:
                            log(UI_TEXT["log_ffmpeg_file"].format(name=file_path.name))
                        log(UI_TEXT["log_ffmpeg_package"])
                        log(UI_TEXT["log_ffmpeg_done"])
                        final_status = UI_TEXT["status_audio_complete"]
                        self.worker_queue.put(("status", UI_TEXT["status_audio_complete"]))
                    else:
                        log(UI_TEXT["log_ffmpeg_failed"])
                        final_status = UI_TEXT["status_ffmpeg_failed"]
                        self.worker_queue.put(("status", UI_TEXT["status_ffmpeg_failed"]))
                        if export_result.errors:
                            write_setup_needed(project_paths, export_result.errors[0])
                else:
                    log(UI_TEXT["log_ffmpeg_failed"])
                    final_status = UI_TEXT["status_ffmpeg_failed"]
                    self.worker_queue.put(("status", UI_TEXT["status_ffmpeg_failed"]))
                    write_setup_needed(project_paths, "FFmpeg is required for audio export")

            log(UI_TEXT["log_complete"])
            write_project_files(project_paths, prompt, direction, process_log)
            self.worker_queue.put(("done", (project_paths.root, final_status)))
        except Exception as exc:
            process_log.append(str(exc))
            self.worker_queue.put(("error", str(exc)))
        finally:
            self.worker_queue.put(("busy", False))

    def _poll_queue(self) -> None:
        try:
            while True:
                event, payload = self.worker_queue.get_nowait()
                if event == "log":
                    self._append_log(str(payload))
                elif event == "status":
                    self.status_var.set(str(payload))
                elif event == "environment":
                    self._apply_environment_report(payload)
                elif event == "done":
                    if isinstance(payload, tuple):
                        folder_payload, status_payload = payload
                    else:
                        folder_payload, status_payload = payload, UI_TEXT["status_complete"]
                    self.output_folder = Path(folder_payload)
                    self.output_var.set(UI_TEXT["output_ready"].format(path=self.output_folder))
                    self.open_button.configure(state="normal")
                    self.status_var.set(str(status_payload))
                elif event == "error":
                    self.status_var.set(UI_TEXT["status_error"])
                    self._append_log(str(payload))
                elif event == "busy":
                    self._set_busy(bool(payload))
                elif event == "brain_response":
                    self.brain_response_var.set(str(payload))
                elif event == "reference_info":
                    self.reference_info_var.set(str(payload))
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def open_current_output(self) -> None:
        if not self.output_folder:
            return
        if not open_output_folder(self.output_folder):
            messagebox.showinfo(UI_TEXT["dialog_error_title"], UI_TEXT["dialog_open_error"])


def run_smoke_test() -> int:
    with tempfile.TemporaryDirectory() as temp_dir:
        prompt = "深夜、コード、ミシン、静かな稼働"
        direction = fallback_direction(prompt)
        paths = create_project("smoke_otooku", base_output_dir=Path(temp_dir))
        log_lines = [
            UI_TEXT["log_brain_received"],
            UI_TEXT["log_brain_thinking"],
            UI_TEXT["log_brain_template"],
            UI_TEXT["log_brain_arranged"],
            UI_TEXT["log_bpm"].format(bpm=direction.bpm),
            UI_TEXT["log_project"],
            UI_TEXT["log_musicgen_unavailable"],
            UI_TEXT["log_complete"],
        ]
        write_setup_needed(paths, "smoke test")
        write_project_files(paths, prompt, direction, log_lines)
        required = (
            paths.prompts / "music_direction.txt",
            paths.prompts / "musicgen_prompt.txt",
            paths.notes / "usage_note.txt",
            paths.logs / "process_log.txt",
            paths.root / "setup_needed.txt",
        )
        for path in required:
            if not path.exists():
                raise AssertionError(f"missing {path.name}")
        if not open_output_folder(paths.root, dry_run=True):
            raise AssertionError("open output folder check failed")
        missing_ffmpeg = export_audio_material(paths.root / "missing.wav", paths.audio, ffmpeg_command="__missing_ffmpeg__")
        if missing_ffmpeg.success or not missing_ffmpeg.errors:
            raise AssertionError("missing ffmpeg path did not return a safe error")
        check_environment()
    print("smoke ok")
    return 0


def run_generate_check() -> int:
    prompt = "深夜、コード、ミシン、静かな稼働"
    report = check_environment()
    paths = create_project(make_project_name(prompt))
    log_lines = [
        UI_TEXT["log_brain_received"],
        UI_TEXT["log_brain_thinking"],
    ]
    response_time: float | None = None
    try:
        started_at = time.perf_counter()
        direction = generate_direction(prompt, report.ollama_models)
        response_time = time.perf_counter() - started_at
        log_lines.append(UI_TEXT["log_brain_ollama"])
        log_lines.append(UI_TEXT["log_brain_low_temp"])
    except Exception:
        direction = fallback_direction(prompt)
        log_lines.append(UI_TEXT["log_brain_template"])

    log_lines.extend(
        [
            UI_TEXT["log_brain_arranged"],
            UI_TEXT["log_bpm"].format(bpm=direction.bpm),
            UI_TEXT["log_project"],
            UI_TEXT["log_musicgen_unavailable"],
            UI_TEXT["log_complete"],
        ]
    )
    write_setup_needed(paths, "generate check")
    write_project_files(paths, prompt, direction, log_lines)
    required = (
        paths.prompts / "music_direction.txt",
        paths.prompts / "musicgen_prompt.txt",
        paths.notes / "usage_note.txt",
        paths.logs / "process_log.txt",
    )
    for path in required:
        if not path.exists():
            raise AssertionError(f"missing {path}")
    print(paths.root)
    if response_time is None:
        print("local brain: fallback")
    else:
        print(f"local brain: {direction.source} {response_time:.2f}s")
    return 0


def run_launch_check() -> int:
    set_windows_app_id()
    root = tk.Tk()
    app = MusicOtookuApp(root)
    root.after(900, app.root.destroy)
    root.mainloop()
    return 0


def main() -> None:
    if "--smoke-test" in sys.argv:
        raise SystemExit(run_smoke_test())
    if "--generate-check" in sys.argv:
        raise SystemExit(run_generate_check())
    if "--launch-check" in sys.argv:
        raise SystemExit(run_launch_check())

    set_windows_app_id()
    root = tk.Tk()
    MusicOtookuApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
