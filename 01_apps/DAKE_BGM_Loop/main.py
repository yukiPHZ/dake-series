# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import json
import os
import queue
import random
import shutil
import subprocess
import sys
import threading
import time
import webbrowser
import wave
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk

from core.generator_adapter import (
    AceStepGeneratorAdapter,
    GenerateRequest,
    GenerateResult,
    MockGeneratorAdapter,
)


APP_KEY = "DAKE_BGM_Loop"
APP_NAME = "Dake BGM Loop"
WINDOW_TITLE = "Dake BGM Loop"
CONCEPT = "のんきなループを、静かに作る。"
TAGLINE = "Generate, Favorite, Save."
APP_DIR = Path(__file__).resolve().parent
SETTINGS_PATH = APP_DIR / "settings.json"
METADATA_DIR_NAME = "metadata"

MOODS = ("のんき", "静か", "作業用", "神社", "雨", "夜", "ミシン", "コード", "余白")
DURATIONS = (15, 30, 60)
MOOD_SLUGS = {
    "のんき": "nonkina",
    "静か": "shizuka",
    "作業用": "sagyouyou",
    "神社": "jinja",
    "雨": "ame",
    "夜": "yoru",
    "ミシン": "mishin",
    "コード": "code",
    "余白": "yohaku",
}
PROMPTS = {
    "のんき": (
        "gentle carefree ambient loop, soft mallet, warm room tone, "
        "simple rhythm, no vocals, no strong melody, seamless loop"
    ),
    "静か": (
        "quiet ambient background loop, soft pad, minimal movement, "
        "no vocals, no drums, calm workspace atmosphere"
    ),
    "作業用": (
        "calm focus background loop, soft pad, subtle pulse, unobtrusive, "
        "no vocals, no strong melody, seamless loop"
    ),
    "神社": (
        "quiet shrine atmosphere, soft bells far away, airy reverb, "
        "calm ambient loop, no vocals"
    ),
    "雨": (
        "soft rainy ambient loop, warm pad, distant drops, calm and unobtrusive, "
        "no vocals, seamless loop"
    ),
    "夜": (
        "quiet night ambient loop, soft low pad, dim room tone, minimal movement, "
        "no vocals, seamless loop"
    ),
    "ミシン": (
        "gentle sewing machine inspired rhythm, soft clicks, warm ambient pad, "
        "cozy workspace loop, no vocals"
    ),
    "コード": (
        "quiet coding background loop, subtle keyboard-like clicks, soft low pad, "
        "minimal, no vocals"
    ),
    "余白": (
        "minimal spacious ambient loop, lots of silence, soft texture, "
        "no vocals, no strong melody"
    ),
}
LICENSE_NOTICE = (
    "生成音声の商用利用可否は、使用モデル・素材・公開先の規約に依存します。"
    "YouTube等で公開する前に、使用モデルのライセンスを確認してください。 "
    "Commercial usability depends on the model, source material, and platform terms. "
    "Please review the model license before publishing."
)
DEFAULT_SETTINGS: dict[str, Any] = {
    "generator_mode": "mock",
    "output_dir": "outputs",
    "favorite_dir": "favorites",
    "last_mood": "のんき",
    "last_duration_sec": 30,
}

COLORS = {
    "bg": "#F5F6F3",
    "panel": "#FFFFFF",
    "text": "#1F2528",
    "muted": "#687076",
    "border": "#DDE2DC",
    "accent": "#2F7D6D",
    "accent_dark": "#246457",
    "accent_light": "#E4F2ED",
    "warning": "#9A6A22",
    "error": "#B42318",
}
FONT_CANDIDATES = ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo", "Segoe UI")


@dataclass(frozen=True)
class RuntimePaths:
    output_dir: Path
    favorite_dir: Path
    metadata_dir: Path


@dataclass
class CurrentGeneration:
    output_path: Path
    metadata_path: Path
    mood: str
    duration_sec: int
    seed: int
    prompt: str
    adapter_name: str
    favorite: bool = False


class WavPreviewPlayer:
    def __init__(self) -> None:
        self._winsound: Any | None = None
        self._process: subprocess.Popen[bytes] | None = None
        if os.name == "nt":
            try:
                import winsound

                self._winsound = winsound
            except Exception:
                self._winsound = None

    def play(self, path: Path) -> str:
        self.stop()
        if self._winsound is not None:
            flags = self._winsound.SND_FILENAME | self._winsound.SND_ASYNC
            self._winsound.PlaySound(str(path), flags)
            return "Preview playing."

        if sys.platform == "darwin":
            self._process = subprocess.Popen(["afplay", str(path)])
            return "Preview playing via afplay."

        aplay = shutil.which("aplay")
        if aplay:
            self._process = subprocess.Popen([aplay, str(path)])
            return "Preview playing via aplay."

        webbrowser.open(path.resolve().as_uri())
        return "Opened in the default player."

    def stop(self) -> str:
        if self._winsound is not None:
            self._winsound.PlaySound(None, self._winsound.SND_PURGE)
        if self._process is not None and self._process.poll() is None:
            self._process.terminate()
        self._process = None
        return "Preview stopped."


def load_settings() -> dict[str, Any]:
    settings = dict(DEFAULT_SETTINGS)
    if SETTINGS_PATH.is_file():
        try:
            loaded = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
            if isinstance(loaded, dict):
                settings.update(loaded)
        except Exception:
            pass
    return settings


def save_settings(settings: dict[str, Any]) -> None:
    SETTINGS_PATH.write_text(
        json.dumps(settings, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def resolve_runtime_paths(settings: dict[str, Any]) -> RuntimePaths:
    return RuntimePaths(
        output_dir=resolve_app_path(settings.get("output_dir"), "outputs"),
        favorite_dir=resolve_app_path(settings.get("favorite_dir"), "favorites"),
        metadata_dir=APP_DIR / METADATA_DIR_NAME,
    )


def resolve_app_path(value: Any, default_name: str) -> Path:
    raw_value = str(value or default_name)
    path = Path(raw_value).expanduser()
    if not path.is_absolute():
        path = APP_DIR / path
    return path


def ensure_runtime_dirs(paths: RuntimePaths) -> None:
    paths.output_dir.mkdir(parents=True, exist_ok=True)
    paths.favorite_dir.mkdir(parents=True, exist_ok=True)
    paths.metadata_dir.mkdir(parents=True, exist_ok=True)


def build_prompt(mood: str) -> str:
    return PROMPTS.get(mood, PROMPTS["のんき"])


def build_output_path(output_dir: Path, created_at: datetime, mood: str, duration_sec: int, seed: int) -> Path:
    timestamp = created_at.strftime("%Y%m%d_%H%M%S")
    slug = MOOD_SLUGS.get(mood, "mood")
    filename = f"{APP_KEY}_{timestamp}_{slug}_{duration_sec}s_seed{seed}.wav"
    return unique_path(output_dir / filename)


def build_request(mood: str, duration_sec: int, seed: int, output_dir: Path, created_at: datetime) -> GenerateRequest:
    output_path = build_output_path(output_dir, created_at, mood, duration_sec, seed)
    return GenerateRequest(
        prompt=build_prompt(mood),
        mood=mood,
        mood_slug=MOOD_SLUGS.get(mood, "mood"),
        duration_sec=duration_sec,
        seed=seed,
        output_path=output_path,
    )


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    for number in range(2, 1000):
        candidate = path.with_name(f"{path.stem}_{number}{path.suffix}")
        if not candidate.exists():
            return candidate
    raise RuntimeError("Could not create a unique file name.")


def write_metadata_file(
    request: GenerateRequest,
    result: GenerateResult,
    metadata_dir: Path,
    created_at: datetime,
    favorite: bool = False,
) -> Path:
    metadata_dir.mkdir(parents=True, exist_ok=True)
    metadata_path = metadata_dir / f"{result.output_path.stem}.json"
    payload = {
        "created_at": created_at.isoformat(timespec="seconds"),
        "mood": request.mood,
        "duration_sec": request.duration_sec,
        "seed": request.seed,
        "prompt": request.prompt,
        "model_adapter": result.adapter_name,
        "output_path": str(result.output_path.resolve()),
        "favorite": favorite,
        "license_notice": LICENSE_NOTICE,
    }
    metadata_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return metadata_path


def mark_metadata_favorite(metadata_path: Path) -> None:
    payload = json.loads(metadata_path.read_text(encoding="utf-8"))
    payload["favorite"] = True
    metadata_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def open_path(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)
    if os.name == "nt":
        os.startfile(str(path))  # type: ignore[attr-defined]
        return
    webbrowser.open(path.resolve().as_uri())


def resolve_font_family(root: tk.Tk) -> str:
    available = set(tkfont.families(root))
    for family in FONT_CANDIDATES:
        if family in available:
            return family
    return "TkDefaultFont"


class DakeBgmLoopApp(tk.Tk):
    def __init__(self, smoke_seconds: float | None = None) -> None:
        super().__init__()
        self.settings = load_settings()
        self.paths = resolve_runtime_paths(self.settings)
        ensure_runtime_dirs(self.paths)
        self.preview_player = WavPreviewPlayer()
        self.work_queue: queue.Queue[tuple[Any, ...]] = queue.Queue()
        self.current: CurrentGeneration | None = None
        self.busy = False

        initial_mood = str(self.settings.get("last_mood") or "のんき")
        if initial_mood not in MOODS:
            initial_mood = "のんき"
        initial_duration = int(self.settings.get("last_duration_sec") or 30)
        if initial_duration not in DURATIONS:
            initial_duration = 30

        self.mood_var = tk.StringVar(value=initial_mood)
        self.duration_var = tk.IntVar(value=initial_duration)
        self.seed_var = tk.StringVar(value="")
        self.mode_var = tk.StringVar(value=str(self.settings.get("generator_mode") or "mock"))
        if self.mode_var.get() not in ("mock", "ace_step"):
            self.mode_var.set("mock")
        self.status_var = tk.StringVar(value="Ready.")
        self.adapter_status_var = tk.StringVar(value="")
        self.seed_used_var = tk.StringVar(value="Seed: auto")
        self.file_var = tk.StringVar(value="No WAV generated yet.")

        self.title(WINDOW_TITLE)
        self.geometry("880x620")
        self.minsize(780, 560)
        self.configure(bg=COLORS["bg"])
        self.font_family = resolve_font_family(self)
        self._configure_style()
        self._build_ui()
        self._refresh_prompt()
        self._refresh_adapter_status()
        self._update_button_states()
        self.after(100, self._poll_queue)
        if smoke_seconds is not None:
            self.after(max(1, int(smoke_seconds * 1000)), self.destroy)

    def _configure_style(self) -> None:
        self.option_add("*Font", f"{{{self.font_family}}} 10")
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TFrame", background=COLORS["bg"])
        style.configure("Panel.TFrame", background=COLORS["panel"], relief="flat")
        style.configure("TLabel", background=COLORS["bg"], foreground=COLORS["text"])
        style.configure("Panel.TLabel", background=COLORS["panel"], foreground=COLORS["text"])
        style.configure("Muted.TLabel", background=COLORS["bg"], foreground=COLORS["muted"])
        style.configure("PanelMuted.TLabel", background=COLORS["panel"], foreground=COLORS["muted"])
        style.configure("TLabelframe", background=COLORS["panel"], bordercolor=COLORS["border"])
        style.configure("TLabelframe.Label", background=COLORS["panel"], foreground=COLORS["text"])
        style.configure("TButton", padding=(12, 7))
        style.configure(
            "Accent.TButton",
            background=COLORS["accent"],
            foreground="#FFFFFF",
            bordercolor=COLORS["accent"],
            focusthickness=0,
        )
        style.map("Accent.TButton", background=[("active", COLORS["accent_dark"])])

    def _font(self, size: int, weight: str = "normal") -> tkfont.Font:
        return tkfont.Font(family=self.font_family, size=size, weight=weight)

    def _build_ui(self) -> None:
        outer = ttk.Frame(self, padding=(22, 18, 22, 14))
        outer.pack(fill="both", expand=True)

        header = ttk.Frame(outer)
        header.pack(fill="x")
        ttk.Label(header, text=APP_NAME, font=self._font(22, "bold")).pack(anchor="w")
        ttk.Label(header, text=f"{CONCEPT}  {TAGLINE}", style="Muted.TLabel").pack(anchor="w", pady=(2, 0))

        body = ttk.Frame(outer)
        body.pack(fill="both", expand=True, pady=(16, 12))
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        left = ttk.LabelFrame(body, text="空気 / 長さ / Seed", padding=16)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        right = ttk.LabelFrame(body, text="Generate / Preview", padding=16)
        right.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        right.columnconfigure(0, weight=1)

        mood_grid = ttk.Frame(left, style="Panel.TFrame")
        mood_grid.pack(fill="x")
        for index, mood in enumerate(MOODS):
            button = tk.Radiobutton(
                mood_grid,
                text=mood,
                value=mood,
                variable=self.mood_var,
                indicatoron=False,
                width=8,
                padx=8,
                pady=8,
                bg=COLORS["panel"],
                fg=COLORS["text"],
                activebackground=COLORS["accent_light"],
                activeforeground=COLORS["text"],
                selectcolor=COLORS["accent_light"],
                relief="flat",
                command=self._on_mood_changed,
            )
            button.grid(row=index // 3, column=index % 3, sticky="ew", padx=4, pady=4)
            mood_grid.columnconfigure(index % 3, weight=1)

        ttk.Label(left, text="長さ", style="PanelMuted.TLabel").pack(anchor="w", pady=(18, 6))
        duration_frame = ttk.Frame(left, style="Panel.TFrame")
        duration_frame.pack(fill="x")
        for duration in DURATIONS:
            rb = tk.Radiobutton(
                duration_frame,
                text=f"{duration} sec",
                value=duration,
                variable=self.duration_var,
                indicatoron=False,
                padx=10,
                pady=8,
                bg=COLORS["panel"],
                fg=COLORS["text"],
                activebackground=COLORS["accent_light"],
                selectcolor=COLORS["accent_light"],
                relief="flat",
                command=self._remember_light_settings,
            )
            rb.pack(side="left", fill="x", expand=True, padx=4)

        ttk.Label(left, text="Seed", style="PanelMuted.TLabel").pack(anchor="w", pady=(18, 6))
        seed_row = ttk.Frame(left, style="Panel.TFrame")
        seed_row.pack(fill="x")
        self.seed_entry = ttk.Entry(seed_row, textvariable=self.seed_var)
        self.seed_entry.pack(side="left", fill="x", expand=True)
        ttk.Label(seed_row, text="空なら自動生成", style="PanelMuted.TLabel").pack(side="left", padx=(10, 0))
        ttk.Label(left, textvariable=self.seed_used_var, style="PanelMuted.TLabel").pack(anchor="w", pady=(8, 0))

        mode_row = ttk.Frame(right, style="Panel.TFrame")
        mode_row.grid(row=0, column=0, sticky="ew")
        ttk.Label(mode_row, text="Adapter", style="PanelMuted.TLabel").pack(side="left", padx=(0, 10))
        ttk.Radiobutton(mode_row, text="Mock", value="mock", variable=self.mode_var, command=self._on_mode_changed).pack(
            side="left", padx=(0, 8)
        )
        ttk.Radiobutton(
            mode_row,
            text="ACE-Step",
            value="ace_step",
            variable=self.mode_var,
            command=self._on_mode_changed,
        ).pack(side="left")

        ttk.Label(right, textvariable=self.adapter_status_var, style="PanelMuted.TLabel").grid(
            row=1, column=0, sticky="w", pady=(8, 10)
        )

        ttk.Label(right, text="Prompt", style="PanelMuted.TLabel").grid(row=2, column=0, sticky="w")
        self.prompt_text = tk.Text(
            right,
            height=5,
            wrap="word",
            relief="solid",
            borderwidth=1,
            bg="#FAFBF8",
            fg=COLORS["text"],
            insertbackground=COLORS["text"],
        )
        self.prompt_text.grid(row=3, column=0, sticky="ew", pady=(6, 14))

        action_grid = ttk.Frame(right, style="Panel.TFrame")
        action_grid.grid(row=4, column=0, sticky="ew")
        action_grid.columnconfigure(0, weight=1)
        action_grid.columnconfigure(1, weight=1)
        self.generate_button = ttk.Button(
            action_grid,
            text="Generate",
            style="Accent.TButton",
            command=lambda: self._generate_from_ui(fresh_seed=False),
        )
        self.generate_button.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        self.preview_button = ttk.Button(action_grid, text="Preview", command=self._preview_current)
        self.preview_button.grid(row=1, column=0, sticky="ew", padx=(0, 4), pady=4)
        self.stop_button = ttk.Button(action_grid, text="Stop", command=self._stop_preview)
        self.stop_button.grid(row=1, column=1, sticky="ew", padx=(4, 0), pady=4)
        self.save_button = ttk.Button(action_grid, text="WAV保存", command=self._save_wav_as)
        self.save_button.grid(row=2, column=0, sticky="ew", padx=(0, 4), pady=4)
        self.favorite_button = ttk.Button(action_grid, text="♡ お気に入り", command=self._favorite_current)
        self.favorite_button.grid(row=2, column=1, sticky="ew", padx=(4, 0), pady=4)
        self.next_button = ttk.Button(
            action_grid,
            text="Next / 再生成",
            command=lambda: self._generate_from_ui(fresh_seed=True),
        )
        self.next_button.grid(row=3, column=0, sticky="ew", padx=(0, 4), pady=4)
        self.open_output_button = ttk.Button(action_grid, text="出力フォルダを開く", command=self._open_output_folder)
        self.open_output_button.grid(row=3, column=1, sticky="ew", padx=(4, 0), pady=4)

        status_box = ttk.Frame(right, style="Panel.TFrame")
        status_box.grid(row=5, column=0, sticky="ew", pady=(14, 0))
        ttk.Label(status_box, textvariable=self.status_var, style="Panel.TLabel").pack(anchor="w")
        ttk.Label(status_box, textvariable=self.file_var, style="PanelMuted.TLabel", wraplength=370).pack(
            anchor="w", pady=(6, 0)
        )

        license_frame = ttk.Frame(outer)
        license_frame.pack(fill="x", side="bottom")
        jp = "生成音声の商用利用可否は、使用モデル・素材・公開先の規約に依存します。YouTube等で公開する前に、使用モデルのライセンスを確認してください。"
        en = "Commercial usability depends on the model, source material, and platform terms. Please review the model license before publishing."
        ttk.Label(license_frame, text=jp, style="Muted.TLabel", font=self._font(8)).pack(anchor="w")
        ttk.Label(license_frame, text=en, style="Muted.TLabel", font=self._font(8)).pack(anchor="w")

    def _on_mood_changed(self) -> None:
        self._refresh_prompt()
        self._remember_light_settings()

    def _on_mode_changed(self) -> None:
        self._remember_light_settings()
        self._refresh_adapter_status()

    def _remember_light_settings(self) -> None:
        self.settings["generator_mode"] = self.mode_var.get()
        self.settings["last_mood"] = self.mood_var.get()
        self.settings["last_duration_sec"] = int(self.duration_var.get())
        save_settings(self.settings)

    def _refresh_prompt(self) -> None:
        self.prompt_text.configure(state="normal")
        self.prompt_text.delete("1.0", "end")
        self.prompt_text.insert("1.0", build_prompt(self.mood_var.get()))
        self.prompt_text.configure(state="disabled")

    def _refresh_adapter_status(self) -> None:
        if self.mode_var.get() == "ace_step":
            self.adapter_status_var.set(AceStepGeneratorAdapter().status_message())
        else:
            self.adapter_status_var.set("Mock mode: local check WAV generator.")

    def _parse_seed(self) -> int:
        raw_seed = self.seed_var.get().strip()
        if not raw_seed:
            return random.randint(1, 999999)
        try:
            seed = int(raw_seed)
        except ValueError as exc:
            raise ValueError("Seedには数字を入力してください。") from exc
        if seed < 0:
            raise ValueError("Seedには0以上の数字を入力してください。")
        return seed

    def _generate_from_ui(self, fresh_seed: bool) -> None:
        if self.busy:
            self.status_var.set("Generating. Please wait.")
            return
        if fresh_seed:
            self.seed_var.set("")
        try:
            seed = self._parse_seed()
        except ValueError as exc:
            messagebox.showerror("Seed", str(exc))
            return

        mood = self.mood_var.get()
        duration_sec = int(self.duration_var.get())
        created_at = datetime.now()
        request = build_request(mood, duration_sec, seed, self.paths.output_dir, created_at)
        mode = self.mode_var.get()

        self.settings["generator_mode"] = mode
        self.settings["last_mood"] = mood
        self.settings["last_duration_sec"] = duration_sec
        save_settings(self.settings)

        self.busy = True
        self.current = None
        self.status_var.set("Generating loop...")
        self.seed_used_var.set(f"Seed: {seed}")
        self.file_var.set("Writing WAV and metadata...")
        self._update_button_states()

        worker = threading.Thread(target=self._generate_worker, args=(request, mode, created_at), daemon=True)
        worker.start()

    def _generate_worker(self, request: GenerateRequest, mode: str, created_at: datetime) -> None:
        try:
            fallback_message = ""
            if mode == "ace_step":
                ace_adapter = AceStepGeneratorAdapter()
                if ace_adapter.is_available():
                    adapter = ace_adapter
                else:
                    fallback_message = ace_adapter.status_message()
                    adapter = MockGeneratorAdapter()
            else:
                adapter = MockGeneratorAdapter()

            result = adapter.generate(request)
            metadata_path = write_metadata_file(request, result, self.paths.metadata_dir, created_at)
            self.work_queue.put(("generated", request, result, metadata_path, fallback_message))
        except Exception as exc:
            self.work_queue.put(("error", str(exc)))

    def _poll_queue(self) -> None:
        try:
            while True:
                item = self.work_queue.get_nowait()
                if item[0] == "generated":
                    _, request, result, metadata_path, fallback_message = item
                    self._handle_generated(request, result, metadata_path, fallback_message)
                elif item[0] == "error":
                    _, message = item
                    self.busy = False
                    self.status_var.set("Generation failed.")
                    self.file_var.set(str(message))
                    self._update_button_states()
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    def _handle_generated(
        self,
        request: GenerateRequest,
        result: GenerateResult,
        metadata_path: Path,
        fallback_message: str,
    ) -> None:
        self.busy = False
        self.current = CurrentGeneration(
            output_path=result.output_path,
            metadata_path=metadata_path,
            mood=request.mood,
            duration_sec=request.duration_sec,
            seed=request.seed,
            prompt=request.prompt,
            adapter_name=result.adapter_name,
        )
        self.seed_var.set(str(request.seed))
        if fallback_message:
            self.status_var.set(fallback_message)
        else:
            self.status_var.set(f"Generated by {result.adapter_name}.")
        self.file_var.set(str(result.output_path))
        self._update_button_states()

    def _update_button_states(self) -> None:
        has_current = self.current is not None and self.current.output_path.is_file()
        busy_state = "disabled" if self.busy else "normal"
        self.generate_button.configure(state=busy_state)
        self.next_button.configure(state=busy_state)
        for button in (self.preview_button, self.stop_button, self.save_button, self.favorite_button):
            button.configure(state="normal" if has_current and not self.busy else "disabled")
        if self.current and self.current.favorite:
            self.favorite_button.configure(text="♥ お気に入り済み")
        else:
            self.favorite_button.configure(text="♡ お気に入り")

    def _preview_current(self) -> None:
        if not self.current:
            self.status_var.set("No WAV to preview.")
            return
        try:
            message = self.preview_player.play(self.current.output_path)
            self.status_var.set(message)
        except Exception as exc:
            self.status_var.set(f"Preview failed: {exc}")

    def _stop_preview(self) -> None:
        try:
            self.status_var.set(self.preview_player.stop())
        except Exception as exc:
            self.status_var.set(f"Stop failed: {exc}")

    def _save_wav_as(self) -> None:
        if not self.current:
            return
        target = filedialog.asksaveasfilename(
            title="WAVを保存",
            defaultextension=".wav",
            initialfile=self.current.output_path.name,
            filetypes=[("WAV", "*.wav"), ("All files", "*.*")],
        )
        if not target:
            return
        try:
            shutil.copy2(self.current.output_path, target)
            self.status_var.set("WAV saved.")
        except Exception as exc:
            messagebox.showerror("WAV保存", f"保存できませんでした。\n{exc}")

    def _favorite_current(self) -> None:
        if not self.current:
            return
        try:
            self.paths.favorite_dir.mkdir(parents=True, exist_ok=True)
            favorite_path = unique_path(self.paths.favorite_dir / self.current.output_path.name)
            shutil.copy2(self.current.output_path, favorite_path)
            mark_metadata_favorite(self.current.metadata_path)
            self.current.favorite = True
            self.status_var.set("Favorite saved.")
            self.file_var.set(str(favorite_path))
            self._update_button_states()
        except Exception as exc:
            messagebox.showerror("お気に入り", f"お気に入りへ保存できませんでした。\n{exc}")

    def _open_output_folder(self) -> None:
        try:
            open_path(self.paths.output_dir)
            self.status_var.set("Output folder opened.")
        except Exception as exc:
            messagebox.showerror("出力フォルダ", f"出力フォルダを開けませんでした。\n{exc}")

    def destroy(self) -> None:
        try:
            self.preview_player.stop()
        finally:
            super().destroy()


def run_generate_check() -> int:
    settings = load_settings()
    paths = resolve_runtime_paths(settings)
    ensure_runtime_dirs(paths)
    adapter = MockGeneratorAdapter()
    generated: list[tuple[GenerateRequest, GenerateResult, Path]] = []
    for duration_sec in DURATIONS:
        seed = 418220 + duration_sec
        created_at = datetime.now()
        request = build_request("のんき", duration_sec, seed, paths.output_dir, created_at)
        result = adapter.generate(request)
        metadata_path = write_metadata_file(request, result, paths.metadata_dir, created_at)
        generated.append((request, result, metadata_path))
        with wave.open(str(result.output_path), "rb") as wav_file:
            actual_duration = wav_file.getnframes() / wav_file.getframerate()
        print(f"generated {duration_sec}s: {result.output_path.name} ({actual_duration:.2f}s)")

    first_request, first_result, first_metadata = generated[0]
    favorite_path = unique_path(paths.favorite_dir / first_result.output_path.name)
    shutil.copy2(first_result.output_path, favorite_path)
    mark_metadata_favorite(first_metadata)
    print(f"favorite copied: {favorite_path.name}")
    print(f"seed saved: {first_request.seed}")
    return 0


def run_preview_check() -> int:
    settings = load_settings()
    paths = resolve_runtime_paths(settings)
    ensure_runtime_dirs(paths)
    wav_files = sorted(paths.output_dir.glob("*.wav"), key=lambda path: path.stat().st_mtime, reverse=True)
    if not wav_files:
        created_at = datetime.now()
        request = build_request("のんき", 15, 418220, paths.output_dir, created_at)
        result = MockGeneratorAdapter().generate(request)
        write_metadata_file(request, result, paths.metadata_dir, created_at)
        wav_files = [result.output_path]
    player = WavPreviewPlayer()
    message = player.play(wav_files[0])
    time.sleep(0.3)
    stop_message = player.stop()
    print(message)
    print(stop_message)
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=APP_NAME)
    parser.add_argument("--generate-check", action="store_true", help="Generate 15/30/60 sec mock WAVs and metadata.")
    parser.add_argument("--preview-check", action="store_true", help="Exercise the preview/stop code path.")
    parser.add_argument("--gui-smoke-seconds", type=float, default=None, help="Launch the GUI and close after N seconds.")
    args = parser.parse_args(argv)

    if args.generate_check:
        return run_generate_check()
    if args.preview_check:
        return run_preview_check()

    app = DakeBgmLoopApp(smoke_seconds=args.gui_smoke_seconds)
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
