# -*- coding: utf-8 -*-
from __future__ import annotations

import ctypes
import json
import math
import os
import queue
import struct
import sys
import tempfile
import threading
import time
import webbrowser
import wave
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox

from core.app_config import DEFAULT_DURATION_SECONDS, ensure_data_dirs
from core.audio_preview import AudioPreviewItem, AudioPreviewPlayer, find_audio_preview_items
from core.audio_probe import probe_audio_info
from core.cli_checker import EnvironmentReport, check_environment
from core.favorites import FAVORITES_DIR, ensure_favorites_dirs, save_favorite_audio
from core.ffmpeg_audio import LoopPackOptions, export_audio_material, export_loop_pack
from core.musicgen_runner import generate_music
from core.ollama_client import generate_direction
from core.presets import (
    MusicPreset,
    fallback_direction_with_preset,
    find_preset,
    load_music_presets,
    merge_preset_tags,
)
from core.project_writer import (
    create_project,
    make_project_name,
    write_loop_notes,
    write_project_files,
    write_setup_needed,
)
from core.prompt_builder import MusicDirection, fallback_direction
from core.video_bgm_pack import export_video_bgm_pack
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
    "preset_label": "Preset:",
    "preset_custom": "Custom",
    "output_label": "OUTPUT / STATUS",
    "system_label": "SYSTEM",
    "log_label": "補助脳ログ",
    "button_place_sound": "音を置く",
    "button_reference": "参照音源を選ぶ",
    "button_open_output": "出力フォルダを開く",
    "button_video_bgm_pack": "Video BGM Pack を作る",
    "button_preview_play": "Play",
    "button_preview_stop": "Stop",
    "button_preview_refresh": "Refresh",
    "button_favorite_add": "Add to Favorite",
    "button_favorite_open": "Open Favorites",
    "button_check": "環境チェック",
    "button_clear_reference": "参照を解除",
    "reference_none": "参照音源: なし",
    "reference_selected": "参照音源: {name}",
    "reference_probe_waiting": "file name: --\nduration: --\ncodec: --\nsample rate: --\nchannels: --",
    "reference_probe_running": "ffprobe info: checking...",
    "reference_probe_failed": "ffprobe info: unavailable",
    "loop_options_label": "LOOP PACK",
    "loop_duration_label": "出力尺",
    "duration_30s": "30s",
    "duration_60s": "60s",
    "duration_180s": "180s",
    "fade_label": "fade",
    "fade_in_label": "in",
    "fade_out_label": "out",
    "volume_label": "音量",
    "volume_original": "original",
    "volume_soft": "soft",
    "volume_normalized": "normalized",
    "tag_label": "用途タグ",
    "tag_quiet": "quiet",
    "tag_work": "work",
    "tag_midnight": "midnight",
    "tag_shrine": "shrine",
    "tag_borinef": "borinef",
    "output_none": "出力先: 未作成",
    "output_ready": "出力先: {path}",
    "status_idle": "稼働中。",
    "status_checking": "環境を確認しています。",
    "status_processing": "音を置いています。",
    "status_complete": "整っています。",
    "status_ffmpeg_processing": "FFmpeg processing...",
    "status_audio_complete": "整いました。 Audio package created.",
    "status_video_bgm_processing": "Video BGM Pack processing...",
    "status_video_bgm_complete": "整いました。 Video BGM Pack exported.",
    "status_video_bgm_failed": "Video BGM Pack failed. Loop Pack output is still available.",
    "status_video_bgm_no_output": "出力フォルダがまだありません。",
    "status_video_bgm_no_loop": "Loop Pack がまだありません。先に音を置いてください。",
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
    "preview_label": "Audio Preview",
    "preview_empty": "Preview files: --",
    "preview_ready": "Preview ready.",
    "preview_no_output": "No output folder.",
    "preview_no_audio": "No mp3 / wav found.",
    "preview_playing": "Playing: {name}",
    "preview_stopped": "Preview stopped.",
    "preview_external": "Opened in default player. Stop in player.",
    "preview_failed": "再生できませんでした。",
    "favorite_saved": "Favorite saved.",
    "favorite_no_audio": "Favorite target is not selected.",
    "favorite_failed": "Favorite save failed.",
    "favorite_opened": "Favorite folder opened.",
    "favorite_open_failed": "Favorite folder could not be opened.",
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
    "log_loop_brain": "補助脳：ループ構成を整えています。",
    "log_brain_start": "補助脳：音の方向性を考えています。",
    "log_brain_ollama": "補助脳：Ollamaで音設計を作成しました。",
    "log_brain_template": "補助脳：固定テンプレートで音設計を作成しました。",
    "log_preset_loaded": "補助脳：{name}プリセットを読み込みました。",
    "log_preset_reflected": "補助脳：選択した空気を音の方向へ反映しました。",
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
    "log_loop_file": "FFMPEG：{name} をloop_packへ書き出しました。",
    "log_loop_exported": "FFMPEG：Loop pack exported.",
    "log_loop_failed": "FFMPEG：Loop pack export failed.",
    "log_video_brain": "補助脳：動画向け用途を整理しています。",
    "log_video_classify": "Loop Pack を分類しています。",
    "log_video_file": "Video BGM Pack：{name} を配置しました。",
    "log_video_exported": "Video BGM Pack exported.",
    "log_video_failed": "Video BGM Pack export failed.",
    "log_preview_ready_brain": "補助脳：音を確認できます。",
    "log_preview_ready": "Preview ready.",
    "log_preview_playing": "Playing audio...",
    "log_preview_stopped": "Preview stopped.",
    "log_preview_failed": "再生できませんでした。",
    "log_preview_external": "既定プレイヤーで開きました。停止はプレイヤー側です。",
    "log_favorite_brain": "補助脳：お気に入りへ置きました。",
    "log_favorite_saved": "Favorite saved.",
    "log_favorite_opened": "Favorite folder opened.",
    "log_favorite_failed": "Favorite save failed.",
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
LOOP_DURATION_OPTIONS = (
    (30, "duration_30s"),
    (60, "duration_60s"),
    (180, "duration_180s"),
)
VOLUME_OPTIONS = (
    ("original", "volume_original"),
    ("soft", "volume_soft"),
    ("normalized", "volume_normalized"),
)
TAG_OPTIONS = (
    ("quiet", "tag_quiet"),
    ("work", "tag_work"),
    ("midnight", "tag_midnight"),
    ("shrine", "tag_shrine"),
    ("borinef", "tag_borinef"),
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


def write_generate_check_wav(path: Path, duration_seconds: float = 2.0) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    sample_rate = 44100
    frames = int(sample_rate * duration_seconds)
    with wave.open(str(path), "wb") as audio_file:
        audio_file.setnchannels(1)
        audio_file.setsampwidth(2)
        audio_file.setframerate(sample_rate)
        for index in range(frames):
            sample = int(32767 * 0.12 * math.sin(2.0 * math.pi * 220.0 * index / sample_rate))
            audio_file.writeframes(struct.pack("<h", sample))


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
        self.last_direction: MusicDirection | None = None
        self.last_preset: MusicPreset | None = None
        self.preview_player = AudioPreviewPlayer()
        self.preview_items: list[AudioPreviewItem] = []
        self.music_presets = load_music_presets()
        self.music_presets_by_name = {preset.name: preset for preset in self.music_presets}
        self.log_lines: list[str] = []
        self.status_vars: dict[str, tk.StringVar] = {}

        self.status_var = tk.StringVar(value=UI_TEXT["status_idle"])
        self.reference_var = tk.StringVar(value=UI_TEXT["reference_none"])
        self.reference_info_var = tk.StringVar(value=UI_TEXT["reference_probe_waiting"])
        self.output_var = tk.StringVar(value=UI_TEXT["output_none"])
        self.brain_response_var = tk.StringVar(value=UI_TEXT["local_brain_idle"])
        self.preview_status_var = tk.StringVar(value=UI_TEXT["preview_empty"])
        self.preset_var = tk.StringVar(value=UI_TEXT["preset_custom"])
        self.loop_duration_vars = {
            duration: tk.BooleanVar(value=True)
            for duration, _text_key in LOOP_DURATION_OPTIONS
        }
        self.fade_in_var = tk.StringVar(value="1.5")
        self.fade_out_var = tk.StringVar(value="2.0")
        self.volume_var = tk.StringVar(value="original")
        self.tag_vars = {
            tag: tk.BooleanVar(value=(tag == "quiet"))
            for tag, _text_key in TAG_OPTIONS
        }

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

        prompt_header = tk.Frame(prompt_panel, bg=COLORS["surface"])
        prompt_header.grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 8))
        prompt_header.grid_columnconfigure(0, weight=1)
        tk.Label(
            prompt_header,
            text=UI_TEXT["prompt_label"],
            bg=COLORS["surface"],
            fg=COLORS["text"],
            font=self.fonts["label"],
            anchor="w",
        ).grid(row=0, column=0, sticky="w")

        preset_row = tk.Frame(prompt_header, bg=COLORS["surface"])
        preset_row.grid(row=0, column=1, sticky="e")
        tk.Label(
            preset_row,
            text=UI_TEXT["preset_label"],
            bg=COLORS["surface"],
            fg=COLORS["muted"],
            font=self.fonts["small"],
        ).pack(side="left", padx=(0, 6))
        self.preset_menu = tk.OptionMenu(
            preset_row,
            self.preset_var,
            UI_TEXT["preset_custom"],
            *(preset.name for preset in self.music_presets),
        )
        self.preset_menu.configure(
            bg=COLORS["surface_soft"],
            fg=COLORS["text"],
            activebackground=COLORS["secondary_hover"],
            relief="flat",
            bd=0,
            font=self.fonts["small"],
            highlightthickness=1,
            highlightbackground=COLORS["border"],
        )
        self.preset_menu.pack(side="left")

        self.prompt_text = tk.Text(
            prompt_panel,
            height=9,
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

        self._build_loop_controls(prompt_panel)

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

        self.video_bgm_button = make_button(
            output_panel,
            UI_TEXT["button_video_bgm_pack"],
            self.start_video_bgm_pack,
            COLORS,
            self.fonts["button"],
            primary=False,
        )
        self.video_bgm_button.grid(row=5, column=0, sticky="ew", padx=16, pady=(0, 14))
        self.video_bgm_button.configure(state="disabled")

        preview_panel = tk.Frame(output_panel, bg=COLORS["surface"])
        preview_panel.grid(row=6, column=0, sticky="ew", padx=16, pady=(0, 14))
        preview_panel.grid_columnconfigure(0, weight=1)

        tk.Label(
            preview_panel,
            text=UI_TEXT["preview_label"],
            bg=COLORS["surface"],
            fg=COLORS["text"],
            font=self.fonts["label"],
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", pady=(0, 6))

        self.preview_listbox = tk.Listbox(
            preview_panel,
            height=5,
            bg=COLORS["surface_soft"],
            fg=COLORS["text"],
            selectbackground=COLORS["select"],
            font=self.fonts["small"],
            relief="flat",
            bd=0,
            activestyle="none",
        )
        self.preview_listbox.grid(row=1, column=0, sticky="ew")
        self.preview_listbox.bind("<<ListboxSelect>>", self._on_preview_select)

        preview_button_row = tk.Frame(preview_panel, bg=COLORS["surface"])
        preview_button_row.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        self.preview_play_button = make_button(
            preview_button_row,
            UI_TEXT["button_preview_play"],
            self.play_selected_preview,
            COLORS,
            self.fonts["button"],
            primary=False,
        )
        self.preview_play_button.pack(side="left")
        self.preview_stop_button = make_button(
            preview_button_row,
            UI_TEXT["button_preview_stop"],
            self.stop_preview,
            COLORS,
            self.fonts["button"],
            primary=False,
        )
        self.preview_stop_button.pack(side="left", padx=(8, 0))
        self.preview_refresh_button = make_button(
            preview_button_row,
            UI_TEXT["button_preview_refresh"],
            self.refresh_audio_preview,
            COLORS,
            self.fonts["button"],
            primary=False,
        )
        self.preview_refresh_button.pack(side="left", padx=(8, 0))

        favorite_button_row = tk.Frame(preview_panel, bg=COLORS["surface"])
        favorite_button_row.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        self.favorite_add_button = make_button(
            favorite_button_row,
            UI_TEXT["button_favorite_add"],
            self.add_selected_favorite,
            COLORS,
            self.fonts["button"],
            primary=False,
        )
        self.favorite_add_button.pack(side="left")
        self.favorite_open_button = make_button(
            favorite_button_row,
            UI_TEXT["button_favorite_open"],
            self.open_favorites_folder,
            COLORS,
            self.fonts["button"],
            primary=False,
        )
        self.favorite_open_button.pack(side="left", padx=(8, 0))

        tk.Label(
            preview_panel,
            textvariable=self.preview_status_var,
            bg=COLORS["surface"],
            fg=COLORS["muted"],
            font=self.fonts["small"],
            anchor="w",
            wraplength=320,
            justify="left",
        ).grid(row=4, column=0, sticky="ew", pady=(8, 0))
        self._sync_preview_controls()

    def _build_loop_controls(self, parent: tk.Widget) -> None:
        panel = tk.Frame(parent, bg=COLORS["surface"])
        panel.grid(row=5, column=0, sticky="ew", padx=16, pady=(0, 14))
        panel.grid_columnconfigure(1, weight=1)

        tk.Label(
            panel,
            text=UI_TEXT["loop_options_label"],
            bg=COLORS["surface"],
            fg=COLORS["text"],
            font=self.fonts["label"],
            anchor="w",
        ).grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6))

        duration_row = tk.Frame(panel, bg=COLORS["surface"])
        duration_row.grid(row=1, column=0, columnspan=2, sticky="ew")
        tk.Label(
            duration_row,
            text=UI_TEXT["loop_duration_label"],
            bg=COLORS["surface"],
            fg=COLORS["muted"],
            font=self.fonts["small"],
        ).pack(side="left", padx=(0, 10))
        for duration, text_key in LOOP_DURATION_OPTIONS:
            tk.Checkbutton(
                duration_row,
                text=UI_TEXT[text_key],
                variable=self.loop_duration_vars[duration],
                bg=COLORS["surface"],
                fg=COLORS["text"],
                activebackground=COLORS["surface"],
                font=self.fonts["small"],
                bd=0,
                padx=2,
                pady=0,
            ).pack(side="left", padx=(0, 8))

        setting_row = tk.Frame(panel, bg=COLORS["surface"])
        setting_row.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        tk.Label(
            setting_row,
            text=UI_TEXT["fade_label"],
            bg=COLORS["surface"],
            fg=COLORS["muted"],
            font=self.fonts["small"],
        ).pack(side="left", padx=(0, 8))
        tk.Label(
            setting_row,
            text=UI_TEXT["fade_in_label"],
            bg=COLORS["surface"],
            fg=COLORS["muted"],
            font=self.fonts["small"],
        ).pack(side="left")
        tk.Entry(
            setting_row,
            textvariable=self.fade_in_var,
            width=5,
            bg=COLORS["surface_soft"],
            fg=COLORS["text"],
            font=self.fonts["small"],
            relief="flat",
        ).pack(side="left", padx=(3, 8))
        tk.Label(
            setting_row,
            text=UI_TEXT["fade_out_label"],
            bg=COLORS["surface"],
            fg=COLORS["muted"],
            font=self.fonts["small"],
        ).pack(side="left")
        tk.Entry(
            setting_row,
            textvariable=self.fade_out_var,
            width=5,
            bg=COLORS["surface_soft"],
            fg=COLORS["text"],
            font=self.fonts["small"],
            relief="flat",
        ).pack(side="left", padx=(3, 12))
        tk.Label(
            setting_row,
            text=UI_TEXT["volume_label"],
            bg=COLORS["surface"],
            fg=COLORS["muted"],
            font=self.fonts["small"],
        ).pack(side="left", padx=(0, 6))
        volume_menu = tk.OptionMenu(setting_row, self.volume_var, *(UI_TEXT[text_key] for _value, text_key in VOLUME_OPTIONS))
        volume_menu.configure(
            bg=COLORS["surface_soft"],
            fg=COLORS["text"],
            activebackground=COLORS["secondary_hover"],
            relief="flat",
            bd=0,
            font=self.fonts["small"],
            highlightthickness=1,
            highlightbackground=COLORS["border"],
        )
        volume_menu.pack(side="left")

        tag_row = tk.Frame(panel, bg=COLORS["surface"])
        tag_row.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        tk.Label(
            tag_row,
            text=UI_TEXT["tag_label"],
            bg=COLORS["surface"],
            fg=COLORS["muted"],
            font=self.fonts["small"],
        ).pack(side="left", padx=(0, 10))
        for tag, text_key in TAG_OPTIONS:
            tk.Checkbutton(
                tag_row,
                text=UI_TEXT[text_key],
                variable=self.tag_vars[tag],
                bg=COLORS["surface"],
                fg=COLORS["text"],
                activebackground=COLORS["surface"],
                font=self.fonts["small"],
                bd=0,
                padx=2,
                pady=0,
            ).pack(side="left", padx=(0, 6))

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
        if hasattr(self, "preset_menu"):
            self.preset_menu.configure(state=state)
        if hasattr(self, "video_bgm_button"):
            if busy:
                self.video_bgm_button.configure(state="disabled")
            else:
                self._sync_video_bgm_button()
        if hasattr(self, "preview_play_button"):
            if busy:
                self.preview_play_button.configure(state="disabled")
                self.preview_refresh_button.configure(state="disabled")
                self.favorite_add_button.configure(state="disabled")
                self.favorite_open_button.configure(state="disabled")
            else:
                self._sync_preview_controls()

    def _sync_video_bgm_button(self) -> None:
        if not hasattr(self, "video_bgm_button"):
            return
        loop_pack = self.output_folder / "audio" / "loop_pack" if self.output_folder else None
        enabled = bool(loop_pack and loop_pack.exists() and not self.processing)
        self.video_bgm_button.configure(state="normal" if enabled else "disabled")

    def _sync_preview_controls(self) -> None:
        if not hasattr(self, "preview_play_button"):
            return
        has_output = bool(self.output_folder)
        has_items = bool(self.preview_items)
        play_state = "normal" if has_items and not self.processing else "disabled"
        refresh_state = "normal" if has_output and not self.processing else "disabled"
        self.preview_play_button.configure(state=play_state)
        self.preview_refresh_button.configure(state=refresh_state)
        self.preview_stop_button.configure(state="normal" if has_items else "disabled")
        if hasattr(self, "favorite_add_button"):
            self.favorite_add_button.configure(state=play_state)
            self.favorite_open_button.configure(state="normal" if not self.processing else "disabled")

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

    def _float_from_var(self, value_var: tk.StringVar, default: float) -> float:
        try:
            value = float(value_var.get())
        except ValueError:
            return default
        return max(0.0, min(value, 20.0))

    def _selected_preset(self) -> MusicPreset | None:
        name = self.preset_var.get()
        if name == UI_TEXT["preset_custom"]:
            return None
        return self.music_presets_by_name.get(name)

    def _loop_options_from_ui(self) -> LoopPackOptions:
        durations = tuple(
            duration
            for duration, _text_key in LOOP_DURATION_OPTIONS
            if self.loop_duration_vars[duration].get()
        )
        if not durations:
            durations = (30,)

        volume_mode = self.volume_var.get()
        valid_volume_modes = {value for value, _text_key in VOLUME_OPTIONS}
        if volume_mode not in valid_volume_modes:
            volume_mode = "original"

        tags = tuple(tag for tag, _text_key in TAG_OPTIONS if self.tag_vars[tag].get())
        if not tags:
            tags = ("quiet",)

        return LoopPackOptions(
            durations=durations,
            fade_in=self._float_from_var(self.fade_in_var, 1.5),
            fade_out=self._float_from_var(self.fade_out_var, 2.0),
            volume_mode=volume_mode,
            tags=tags,
        )

    def start_place_sound(self) -> None:
        if self.processing:
            return
        prompt = self.prompt_text.get("1.0", "end").strip()
        if not prompt or prompt == UI_TEXT["prompt_hint"]:
            self.status_var.set(UI_TEXT["status_no_prompt"])
            return
        loop_options = self._loop_options_from_ui()
        selected_preset = self._selected_preset()
        if selected_preset:
            loop_options = LoopPackOptions(
                durations=loop_options.durations,
                fade_in=loop_options.fade_in,
                fade_out=loop_options.fade_out,
                volume_mode=loop_options.volume_mode,
                tags=merge_preset_tags(loop_options.tags, selected_preset),
            )
        reference_audio_path = self.reference_audio_path
        self._set_busy(True)
        self.status_var.set(UI_TEXT["status_processing"])
        thread = threading.Thread(
            target=self._place_sound_worker,
            args=(prompt, loop_options, reference_audio_path, selected_preset),
            daemon=True,
        )
        thread.start()

    def _place_sound_worker(
        self,
        prompt: str,
        loop_options: LoopPackOptions,
        reference_audio_path: Path | None,
        selected_preset: MusicPreset | None,
    ) -> None:
        process_log: list[str] = []
        final_status = UI_TEXT["status_complete"]

        def log(message: str) -> None:
            process_log.append(message)
            self.worker_queue.put(("log", message))

        try:
            log(UI_TEXT["log_brain_received"])
            if selected_preset:
                log(UI_TEXT["log_preset_loaded"].format(name=selected_preset.name))
            log(UI_TEXT["log_brain_thinking"])
            self.worker_queue.put(("brain_response", UI_TEXT["local_brain_measuring"]))
            report = self.environment_report or check_environment()
            if self.environment_report is None:
                self.worker_queue.put(("environment", report))

            try:
                started_at = time.perf_counter()
                direction = generate_direction(prompt, report.ollama_models, selected_preset)
                elapsed = time.perf_counter() - started_at
                self.worker_queue.put(("brain_response", UI_TEXT["local_brain_online"].format(seconds=elapsed)))
                log(UI_TEXT["log_brain_ollama"])
                log(UI_TEXT["log_brain_low_temp"])
            except Exception:
                direction = fallback_direction_with_preset(prompt, selected_preset)
                self.worker_queue.put(("brain_response", UI_TEXT["local_brain_offline"]))
                log(UI_TEXT["log_brain_template"])

            self.worker_queue.put(("direction", direction))
            self.worker_queue.put(("preset", selected_preset))
            if selected_preset:
                log(UI_TEXT["log_preset_reflected"])
            log(UI_TEXT["log_brain_arranged"])
            log(UI_TEXT["log_bpm"].format(bpm=direction.bpm))
            project_name = make_project_name(prompt)
            project_paths = create_project(project_name)
            log(UI_TEXT["log_project"])

            source_audio: Path | None = None
            if reference_audio_path:
                source_audio = reference_audio_path
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
                    output_stem = "source_converted" if reference_audio_path else "generated"
                    export_result = export_audio_material(source_audio, project_paths.audio, output_stem=output_stem)
                    if export_result.success:
                        for file_path in export_result.files:
                            log(UI_TEXT["log_ffmpeg_file"].format(name=file_path.name))
                        log(UI_TEXT["log_ffmpeg_package"])
                        log(UI_TEXT["log_ffmpeg_done"])

                        log(UI_TEXT["log_loop_brain"])
                        loop_result = export_loop_pack(source_audio, project_paths.audio, loop_options)
                        if loop_result.files:
                            for file_path in loop_result.files:
                                log(UI_TEXT["log_loop_file"].format(name=file_path.name))
                            write_loop_notes(
                                project_paths,
                                direction,
                                loop_options.tags,
                                loop_options.durations,
                                loop_options.fade_in,
                                loop_options.fade_out,
                                loop_options.volume_mode,
                                loop_result.files,
                                selected_preset,
                            )
                        if loop_result.success:
                            log(UI_TEXT["log_loop_exported"])
                            final_status = UI_TEXT["status_audio_complete"]
                            self.worker_queue.put(("status", UI_TEXT["status_audio_complete"]))
                        else:
                            log(UI_TEXT["log_loop_failed"])
                            if loop_result.errors:
                                write_setup_needed(project_paths, loop_result.errors[0])
                            final_status = UI_TEXT["status_ffmpeg_failed"]
                            self.worker_queue.put(("status", UI_TEXT["status_ffmpeg_failed"]))
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
            write_project_files(project_paths, prompt, direction, process_log, selected_preset)
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
                    self._sync_video_bgm_button()
                    self.refresh_audio_preview(log_ready=False)
                    self.status_var.set(str(status_payload))
                elif event == "error":
                    self.status_var.set(UI_TEXT["status_error"])
                    self._append_log(str(payload))
                    self._sync_video_bgm_button()
                elif event == "busy":
                    self._set_busy(bool(payload))
                elif event == "brain_response":
                    self.brain_response_var.set(str(payload))
                elif event == "direction":
                    self.last_direction = payload
                elif event == "preset":
                    self.last_preset = payload
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

    def refresh_audio_preview(self, log_ready: bool = True) -> None:
        if not hasattr(self, "preview_listbox"):
            return
        self.preview_items = []
        self.preview_listbox.delete(0, "end")
        if not self.output_folder:
            self.preview_status_var.set(UI_TEXT["preview_no_output"])
            self._sync_preview_controls()
            return

        self.preview_items = find_audio_preview_items(self.output_folder)
        for item in self.preview_items:
            self.preview_listbox.insert("end", item.label)

        if self.preview_items:
            self.preview_listbox.selection_set(0)
            self.preview_status_var.set(UI_TEXT["preview_ready"])
            if log_ready:
                self._append_log(UI_TEXT["log_preview_ready_brain"])
                self._append_log(UI_TEXT["log_preview_ready"])
        else:
            self.preview_status_var.set(UI_TEXT["preview_no_audio"])
        self._sync_preview_controls()

    def _selected_preview_item(self) -> AudioPreviewItem | None:
        if not self.preview_items:
            return None
        selection = self.preview_listbox.curselection() if hasattr(self, "preview_listbox") else ()
        index = selection[0] if selection else 0
        if index < 0 or index >= len(self.preview_items):
            return None
        return self.preview_items[index]

    def _on_preview_select(self, _event: tk.Event) -> None:
        self._sync_preview_controls()

    def play_selected_preview(self) -> None:
        item = self._selected_preview_item()
        if not item:
            self.preview_status_var.set(UI_TEXT["preview_no_audio"])
            return
        result = self.preview_player.play(item.path)
        if result.success:
            self.preview_status_var.set(UI_TEXT["preview_playing"].format(name=item.path.name))
            self._append_log(UI_TEXT["log_preview_playing"])
            if result.mode == "external":
                self.preview_status_var.set(UI_TEXT["preview_external"])
                self._append_log(UI_TEXT["log_preview_external"])
        else:
            self.preview_status_var.set(UI_TEXT["preview_failed"])
            self._append_log(UI_TEXT["log_preview_failed"])
        self._sync_preview_controls()

    def stop_preview(self) -> None:
        result = self.preview_player.stop()
        if result.success:
            if result.mode == "external":
                self.preview_status_var.set(UI_TEXT["preview_external"])
                self._append_log(UI_TEXT["log_preview_external"])
            else:
                self.preview_status_var.set(UI_TEXT["preview_stopped"])
                self._append_log(UI_TEXT["log_preview_stopped"])
        else:
            self.preview_status_var.set(UI_TEXT["preview_failed"])
            self._append_log(UI_TEXT["log_preview_failed"])
        self._sync_preview_controls()

    def add_selected_favorite(self) -> None:
        item = self._selected_preview_item()
        if not item:
            self.preview_status_var.set(UI_TEXT["favorite_no_audio"])
            return
        result = save_favorite_audio(item.path, self.output_folder, self.last_preset)
        if result.success and result.record:
            self.preview_status_var.set(UI_TEXT["favorite_saved"])
            self._append_log(UI_TEXT["log_favorite_brain"])
            self._append_log(UI_TEXT["log_favorite_saved"])
        else:
            self.preview_status_var.set(UI_TEXT["favorite_failed"])
            self._append_log(UI_TEXT["log_favorite_failed"])
            if result.error:
                self._append_log(result.error)
        self._sync_preview_controls()

    def open_favorites_folder(self) -> None:
        ensure_favorites_dirs()
        if open_output_folder(FAVORITES_DIR):
            self.preview_status_var.set(UI_TEXT["favorite_opened"])
            self._append_log(UI_TEXT["log_favorite_opened"])
        else:
            self.preview_status_var.set(UI_TEXT["favorite_open_failed"])
            self._append_log(UI_TEXT["favorite_open_failed"])

    def start_video_bgm_pack(self) -> None:
        if self.processing:
            return
        if not self.output_folder:
            self.status_var.set(UI_TEXT["status_video_bgm_no_output"])
            return
        loop_pack = self.output_folder / "audio" / "loop_pack"
        if not loop_pack.exists():
            self.status_var.set(UI_TEXT["status_video_bgm_no_loop"])
            return

        prompt = self.prompt_text.get("1.0", "end").strip()
        if not prompt or prompt == UI_TEXT["prompt_hint"]:
            prompt = ""
        selected_preset = self.last_preset or self._selected_preset()
        direction = self.last_direction or fallback_direction_with_preset(prompt, selected_preset)
        self._set_busy(True)
        self.status_var.set(UI_TEXT["status_video_bgm_processing"])
        thread = threading.Thread(
            target=self._video_bgm_pack_worker,
            args=(self.output_folder, direction, selected_preset),
            daemon=True,
        )
        thread.start()

    def _video_bgm_pack_worker(
        self,
        output_folder: Path,
        direction: MusicDirection,
        selected_preset: MusicPreset | None,
    ) -> None:
        def log(message: str) -> None:
            self.worker_queue.put(("log", message))

        try:
            report = self.environment_report or check_environment()
            if self.environment_report is None:
                self.worker_queue.put(("environment", report))

            log(UI_TEXT["log_video_brain"])
            if selected_preset:
                log(UI_TEXT["log_preset_loaded"].format(name=selected_preset.name))
            result = export_video_bgm_pack(output_folder, direction, report.ollama_models, selected_preset)
            if result.suggestion and result.suggestion.response_time is not None:
                self.worker_queue.put(
                    ("brain_response", UI_TEXT["local_brain_online"].format(seconds=result.suggestion.response_time))
                )
            elif result.suggestion and result.suggestion.source == "template":
                self.worker_queue.put(("brain_response", UI_TEXT["local_brain_offline"]))

            log(UI_TEXT["log_video_classify"])
            for file_path in result.copied_files:
                log(UI_TEXT["log_video_file"].format(name=file_path.name))

            if result.success:
                log(UI_TEXT["log_video_exported"])
                log(UI_TEXT["log_complete"])
                self.worker_queue.put(("done", (output_folder, UI_TEXT["status_video_bgm_complete"])))
            else:
                log(UI_TEXT["log_video_failed"])
                if result.errors:
                    log(result.errors[0])
                self.worker_queue.put(("status", UI_TEXT["status_video_bgm_failed"]))
        except Exception as exc:
            self.worker_queue.put(("error", str(exc)))
        finally:
            self.worker_queue.put(("busy", False))


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
        if "Preset:\n" in (paths.prompts / "music_direction.txt").read_text(encoding="utf-8"):
            raise AssertionError("custom output unexpectedly contains preset metadata")

        borinef = find_preset("BORINEF")
        if not borinef:
            raise AssertionError("BORINEF preset missing")
        borinef_paths = create_project("smoke_borinef", base_output_dir=Path(temp_dir))
        borinef_direction = fallback_direction_with_preset(prompt, borinef)
        write_project_files(borinef_paths, prompt, borinef_direction, log_lines, borinef)
        if "Preset:\nBORINEF" not in (borinef_paths.prompts / "music_direction.txt").read_text(encoding="utf-8"):
            raise AssertionError("BORINEF preset was not written to music_direction.txt")
        if "Preset Tags:\nborinef, quiet, ember" not in (borinef_paths.prompts / "musicgen_prompt.txt").read_text(encoding="utf-8"):
            raise AssertionError("BORINEF preset tags were not written to musicgen_prompt.txt")

        yukiz = find_preset("YUKIZ稼働中")
        if not yukiz:
            raise AssertionError("YUKIZ稼働中 preset missing")
        yukiz_paths = create_project("smoke_yukiz", base_output_dir=Path(temp_dir))
        yukiz_direction = fallback_direction_with_preset(prompt, yukiz)
        write_project_files(yukiz_paths, prompt, yukiz_direction, log_lines, yukiz)
        if "Preset:\nYUKIZ稼働中" not in (yukiz_paths.notes / "usage_note.txt").read_text(encoding="utf-8"):
            raise AssertionError("YUKIZ稼働中 preset was not written to usage_note.txt")
        preview_wav = paths.audio / "preview_check.wav"
        write_generate_check_wav(preview_wav, duration_seconds=0.2)
        preview_items = find_audio_preview_items(paths.root)
        if not preview_items:
            raise AssertionError("preview audio list was not populated")
        preview_player = AudioPreviewPlayer()
        play_result = preview_player.play(preview_wav)
        if not play_result.success:
            raise AssertionError(f"preview wav did not play safely: {play_result.message}")
        stop_result = preview_player.stop()
        if not stop_result.success:
            raise AssertionError(f"preview stop failed: {stop_result.message}")
        missing_preview = preview_player.play(paths.audio / "missing_preview.mp3")
        if missing_preview.success:
            raise AssertionError("missing preview audio unexpectedly played")
        favorite_dir = Path(temp_dir) / "favorites"
        first_favorite = save_favorite_audio(preview_wav, paths.root, borinef, favorites_dir=favorite_dir)
        second_favorite = save_favorite_audio(preview_wav, paths.root, borinef, favorites_dir=favorite_dir)
        if not first_favorite.success or not second_favorite.success:
            raise AssertionError("favorite save failed")
        if not first_favorite.record or not second_favorite.record:
            raise AssertionError("favorite record was not created")
        if first_favorite.record.favorite_path == second_favorite.record.favorite_path:
            raise AssertionError("duplicate favorite overwrote the existing file")
        favorite_index = favorite_dir / "favorite_index.json"
        favorite_note = favorite_dir / "notes" / "favorite_note.txt"
        if not favorite_index.exists():
            raise AssertionError("favorite_index.json was not created")
        if not favorite_note.exists():
            raise AssertionError("favorite_note.txt was not created")
        index_data = json.loads(favorite_index.read_text(encoding="utf-8"))
        if len(index_data.get("favorites", [])) != 2:
            raise AssertionError("favorite_index.json did not record both favorites")
        if "Preset:\nBORINEF" not in favorite_note.read_text(encoding="utf-8"):
            raise AssertionError("favorite_note.txt did not include preset")
        ensure_favorites_dirs(favorite_dir)
        if not open_output_folder(favorite_dir, dry_run=True):
            raise AssertionError("open favorites folder check failed")
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
    preset = find_preset("BORINEF")
    if not preset:
        raise AssertionError("BORINEF preset missing")
    report = check_environment()
    paths = create_project(make_project_name(prompt))
    log_lines = [
        UI_TEXT["log_brain_received"],
        UI_TEXT["log_preset_loaded"].format(name=preset.name),
        UI_TEXT["log_brain_thinking"],
    ]
    response_time: float | None = None
    try:
        started_at = time.perf_counter()
        direction = generate_direction(prompt, report.ollama_models, preset)
        response_time = time.perf_counter() - started_at
        log_lines.append(UI_TEXT["log_brain_ollama"])
        log_lines.append(UI_TEXT["log_brain_low_temp"])
    except Exception:
        direction = fallback_direction_with_preset(prompt, preset)
        log_lines.append(UI_TEXT["log_brain_template"])

    log_lines.extend(
        [
            UI_TEXT["log_preset_reflected"],
            UI_TEXT["log_brain_arranged"],
            UI_TEXT["log_bpm"].format(bpm=direction.bpm),
            UI_TEXT["log_project"],
            UI_TEXT["log_musicgen_unavailable"],
            UI_TEXT["log_complete"],
        ]
    )
    write_setup_needed(paths, "generate check")
    write_project_files(paths, prompt, direction, log_lines, preset)
    required = (
        paths.prompts / "music_direction.txt",
        paths.prompts / "musicgen_prompt.txt",
        paths.notes / "usage_note.txt",
        paths.logs / "process_log.txt",
    )
    for path in required:
        if not path.exists():
            raise AssertionError(f"missing {path}")
    if "Preset:\nBORINEF" not in (paths.prompts / "music_direction.txt").read_text(encoding="utf-8"):
        raise AssertionError("BORINEF preset was not written to music_direction.txt")
    if "Preset Tags:\nborinef, quiet, ember" not in (paths.notes / "usage_note.txt").read_text(encoding="utf-8"):
        raise AssertionError("BORINEF preset tags were not written to usage_note.txt")
    ffmpeg_status = report.status_for("ffmpeg")
    video_result = None
    if ffmpeg_status and ffmpeg_status.state == "ONLINE":
        source_wav = paths.audio / "_generate_check_source.wav"
        write_generate_check_wav(source_wav)
        loop_options = LoopPackOptions(tags=merge_preset_tags(("quiet", "midnight"), preset))
        loop_result = export_loop_pack(source_wav, paths.audio, loop_options)
        if not loop_result.success:
            raise AssertionError(f"loop pack failed: {loop_result.errors}")
        write_loop_notes(
            paths,
            direction,
            loop_options.tags,
            loop_options.durations,
            loop_options.fade_in,
            loop_options.fade_out,
            loop_options.volume_mode,
            loop_result.files,
            preset,
        )
        video_result = export_video_bgm_pack(paths.root, direction, report.ollama_models, preset)
        if not video_result.success:
            raise AssertionError(f"video bgm pack failed: {video_result.errors}")
        required_video = (
            paths.root / "video_bgm_pack",
            paths.root / "video_bgm_pack" / "bgm" / "shorts",
            paths.root / "video_bgm_pack" / "bgm" / "long",
            paths.root / "video_bgm_pack" / "bgm" / "ambient",
            paths.root / "video_bgm_pack" / "bgm" / "work",
            paths.root / "video_bgm_pack" / "notes" / "usage_note.txt",
            paths.root / "video_bgm_pack" / "notes" / "shorts_ideas.txt",
            paths.root / "video_bgm_pack" / "notes" / "long_video_ideas.txt",
            paths.root / "video_bgm_pack" / "export_log.txt",
        )
        for path in required_video:
            if not path.exists():
                raise AssertionError(f"missing {path}")
        if "Preset:\nBORINEF" not in (paths.root / "video_bgm_pack" / "notes" / "usage_note.txt").read_text(encoding="utf-8"):
            raise AssertionError("BORINEF preset was not written to video bgm usage_note.txt")
        preview_items = find_audio_preview_items(paths.root)
        if not preview_items:
            raise AssertionError("preview audio list was not populated")
    print(paths.root)
    print(f"preset: {preset.name}")
    if response_time is None:
        print("local brain: fallback")
    else:
        print(f"local brain: {direction.source} {response_time:.2f}s")
    if video_result and video_result.suggestion:
        if video_result.suggestion.response_time is None:
            print(f"video bgm brain: {video_result.suggestion.source}")
        else:
            print(f"video bgm brain: {video_result.suggestion.source} {video_result.suggestion.response_time:.2f}s")
        print(f"video bgm pack: {video_result.root}")
    else:
        print("video bgm pack: skipped")
    preview_items = find_audio_preview_items(paths.root)
    print(f"preview files: {len(preview_items)}")
    if preview_items:
        favorite_source = next((item for item in preview_items if "loop_pack" in item.label), preview_items[0])
        favorite_result = save_favorite_audio(favorite_source.path, paths.root, preset)
        if not favorite_result.success or not favorite_result.record:
            raise AssertionError(f"favorite save failed: {favorite_result.error}")
        print(f"favorite saved: {favorite_result.record.favorite_path}")
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
