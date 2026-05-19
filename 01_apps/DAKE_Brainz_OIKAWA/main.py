# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import math
import queue
import random
import threading
import time
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk

from core.config import AppConfig, ConfigStore, existing_folder, open_path, resolve_memory_folder
from core.heat_engine import AnalysisResult, analyze_documents
from core.markdown_writer import write_suggestion
from core.scanner import scan_memory


APP_NAME = "DakeBrainzOIKAWA"
WINDOW_TITLE = "OIKAWA"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "app_title": "OIKAWA",
    "app_subtitle": "記憶層を巡回する観測装置",
    "button_scan": "巡回する",
    "button_scanning": "巡回中",
    "button_open_output": "保存先を開く",
    "button_choose_folder": "記憶フォルダを選ぶ",
    "status_idle": "眠っています",
    "status_scanning": "記憶層を巡回中",
    "status_complete": "観測完了",
    "status_no_trace": "強い熱の痕跡は見つかりませんでした",
    "status_error": "巡回に失敗しました",
    "section_traces": "浮上した痕跡",
    "section_related": "関連断片",
    "section_suggestion": "OIKAWA提案",
    "label_memory_folder": "記憶フォルダ",
    "memory_folder_missing": "記憶フォルダ未検出",
    "dialog_choose_memory": "記憶フォルダを選択",
    "dialog_title": "OIKAWA",
    "summary_idle": "呼ばれるまで、記憶層の外側で待機します",
    "summary_scan": "files {files} / skipped {skipped} / traces {traces}",
    "summary_saved": "提案Markdownを保存しました",
    "card_file": "該当ファイル",
    "card_excerpt": "抜粋",
    "card_score": "score {score}",
    "card_empty": "まだ浮上したカードはありません",
    "footer_source": "local scan / no cloud",
    "launch_check_ok": "LAUNCH CHECK OK",
    "gui_smoke_ok": "GUI SMOKE OK",
    "scan_check_ok": "SCAN CHECK OK",
    "ghost_words": ["熾火", "巡り", "側に", "在る", "余白", "記憶", "痕跡"],
}

COLORS = {
    "background": "#06070A",
    "panel": "#101218",
    "panel_light": "#151821",
    "text": "#D7DAE0",
    "muted": "#7B8190",
    "glow": "#6E7FA8",
    "heat": "#C47A3A",
    "line": "#242B3A",
    "line_soft": "#161B26",
}

FONT_JP = ("BIZ UDPGothic", 11)
FONT_JP_SMALL = ("BIZ UDPGothic", 9)
FONT_TITLE = ("Segoe UI", 18, "bold")
FONT_LABEL = ("Segoe UI", 9)
FONT_MONO = ("JetBrains Mono", 9)


@dataclass
class Particle:
    x: float
    y: float
    vx: float
    vy: float
    size: float
    phase: float


class OikawaApp(tk.Tk):
    def __init__(
        self,
        launch_check: bool = False,
        gui_smoke_seconds: float = 0.0,
        memory_folder_override: str = "",
    ) -> None:
        super().__init__()
        self.title(WINDOW_TITLE)
        self.geometry("1120x720")
        self.minsize(980, 620)
        self.configure(bg=COLORS["background"])

        self.config_store = ConfigStore()
        self.config_data = self.config_store.load()
        self.memory_folder = self._resolve_initial_memory_folder(memory_folder_override)
        self.output_path: Path | None = None
        self.events: queue.Queue[tuple[str, object]] = queue.Queue()
        self.scan_thread: threading.Thread | None = None
        self.scanning = False
        self.particles: list[Particle] = []
        self.last_animation = time.monotonic()

        self._build_canvas()
        self._build_overlay()
        self._init_particles()
        self._render_memory_state()
        self._render_empty_cards()

        self.after(80, self._animate)
        self.after(100, self._poll_events)

        if launch_check:
            self.after(500, self._finish_launch_check)
        elif gui_smoke_seconds > 0:
            self.after(max(500, int(gui_smoke_seconds * 1000)), self._finish_gui_smoke)

    def _resolve_initial_memory_folder(self, override: str) -> Path | None:
        if override:
            folder = existing_folder(override)
            if folder:
                self.config_data.memory_folder = str(folder)
                self.config_store.save(self.config_data)
            return folder

        folder = resolve_memory_folder(self.config_data)
        if folder and self.config_data.memory_folder != str(folder):
            self.config_data.memory_folder = str(folder)
            self.config_store.save(self.config_data)
        return folder

    def _build_canvas(self) -> None:
        self.canvas = tk.Canvas(
            self,
            bg=COLORS["background"],
            highlightthickness=0,
            bd=0,
        )
        self.canvas.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.canvas.bind("<Configure>", lambda _event: self._init_particles())

    def _build_overlay(self) -> None:
        self.status_var = tk.StringVar(value=UI_TEXT["status_idle"])
        self.memory_var = tk.StringVar(value="")
        self.summary_var = tk.StringVar(value=UI_TEXT["summary_idle"])

        header = tk.Frame(self, bg=COLORS["background"])
        header.place(x=30, y=24)
        tk.Label(
            header,
            text=UI_TEXT["app_title"],
            fg=COLORS["text"],
            bg=COLORS["background"],
            font=FONT_TITLE,
        ).pack(anchor="w")
        tk.Label(
            header,
            text=UI_TEXT["app_subtitle"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
            font=FONT_JP_SMALL,
        ).pack(anchor="w", pady=(4, 12))
        tk.Label(
            header,
            textvariable=self.status_var,
            fg=COLORS["heat"],
            bg=COLORS["background"],
            font=FONT_JP_SMALL,
        ).pack(anchor="w")
        tk.Label(
            header,
            textvariable=self.memory_var,
            fg=COLORS["muted"],
            bg=COLORS["background"],
            font=FONT_MONO,
            wraplength=520,
            justify="left",
        ).pack(anchor="w", pady=(8, 0))

        self.missing_frame = tk.Frame(
            self,
            bg=COLORS["panel"],
            highlightbackground=COLORS["line"],
            highlightthickness=1,
            padx=18,
            pady=16,
        )
        tk.Label(
            self.missing_frame,
            text=UI_TEXT["memory_folder_missing"],
            fg=COLORS["text"],
            bg=COLORS["panel"],
            font=FONT_JP,
        ).pack(anchor="w")
        tk.Button(
            self.missing_frame,
            text=UI_TEXT["button_choose_folder"],
            command=self._choose_memory_folder,
            **self._button_style(COLORS["panel_light"]),
        ).pack(anchor="w", pady=(12, 0))

        self.results_frame = tk.Frame(self, bg=COLORS["background"])
        self.results_frame.place(relx=0.04, rely=0.50, relwidth=0.58, relheight=0.44)
        tk.Label(
            self.results_frame,
            text=UI_TEXT["section_related"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
            font=FONT_JP_SMALL,
        ).pack(anchor="w")
        self.cards_frame = tk.Frame(self.results_frame, bg=COLORS["background"])
        self.cards_frame.pack(fill="both", expand=True, pady=(10, 0))

        actions = tk.Frame(self, bg=COLORS["background"])
        actions.place(relx=0.97, rely=0.94, anchor="se")
        tk.Label(
            actions,
            text=UI_TEXT["footer_source"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
            font=FONT_LABEL,
        ).pack(anchor="e", pady=(0, 10))
        self.open_output_button = tk.Button(
            actions,
            text=UI_TEXT["button_open_output"],
            command=self._open_output,
            state="disabled",
            **self._button_style(COLORS["panel_light"]),
        )
        self.open_output_button.pack(anchor="e", pady=(0, 10))
        self.scan_button = tk.Button(
            actions,
            text=UI_TEXT["button_scan"],
            command=self._start_scan,
            **self._button_style(COLORS["heat"]),
        )
        self.scan_button.pack(anchor="e")

        footer = tk.Frame(self, bg=COLORS["background"])
        footer.place(relx=0.04, rely=0.96, anchor="sw")
        tk.Label(
            footer,
            text=COPYRIGHT,
            fg=COLORS["muted"],
            bg=COLORS["background"],
            font=FONT_LABEL,
        ).pack(anchor="w")
        tk.Label(
            footer,
            textvariable=self.summary_var,
            fg=COLORS["muted"],
            bg=COLORS["background"],
            font=FONT_MONO,
        ).pack(anchor="w", pady=(4, 0))

    def _button_style(self, background: str) -> dict[str, object]:
        return {
            "bg": background,
            "fg": COLORS["text"],
            "activebackground": COLORS["glow"],
            "activeforeground": COLORS["text"],
            "disabledforeground": COLORS["muted"],
            "font": FONT_JP_SMALL,
            "relief": "flat",
            "bd": 0,
            "padx": 18,
            "pady": 9,
            "highlightthickness": 1,
            "highlightbackground": COLORS["line"],
            "cursor": "hand2",
        }

    def _render_memory_state(self) -> None:
        if self.memory_folder:
            self.memory_var.set(f"{UI_TEXT['label_memory_folder']}: {self.memory_folder}")
            self.missing_frame.place_forget()
            return

        self.memory_var.set(UI_TEXT["memory_folder_missing"])
        self.missing_frame.place(relx=0.04, rely=0.30, relwidth=0.28)

    def _init_particles(self) -> None:
        width = max(1, self.canvas.winfo_width() or 1120)
        height = max(1, self.canvas.winfo_height() or 720)
        random.seed(20260518)
        count = 32
        self.particles = [
            Particle(
                x=random.uniform(0, width),
                y=random.uniform(0, height),
                vx=random.uniform(-0.18, 0.18),
                vy=random.uniform(-0.14, 0.14),
                size=random.uniform(1.0, 2.2),
                phase=random.uniform(0, math.tau),
            )
            for _ in range(count)
        ]

    def _animate(self) -> None:
        now = time.monotonic()
        delta = min(0.2, now - self.last_animation)
        self.last_animation = now
        width = max(1, self.canvas.winfo_width())
        height = max(1, self.canvas.winfo_height())
        speed = 2.2 if self.scanning else 1.0

        self.canvas.delete("field")
        self.canvas.create_rectangle(0, 0, width, height, fill=COLORS["background"], outline="", tags="field")
        self._draw_ghost_words(width, height)

        for particle in self.particles:
            particle.phase += delta * 0.6
            particle.x += (particle.vx + math.sin(particle.phase) * 0.05) * speed * delta * 60
            particle.y += (particle.vy + math.cos(particle.phase) * 0.04) * speed * delta * 60
            if particle.x < 0:
                particle.x += width
            elif particle.x > width:
                particle.x -= width
            if particle.y < 0:
                particle.y += height
            elif particle.y > height:
                particle.y -= height

        for index, current in enumerate(self.particles):
            for other in self.particles[index + 1 :]:
                distance = math.hypot(current.x - other.x, current.y - other.y)
                if distance < 138:
                    color = COLORS["line"] if self.scanning and distance < 96 else COLORS["line_soft"]
                    self.canvas.create_line(current.x, current.y, other.x, other.y, fill=color, width=1, tags="field")

        for particle in self.particles:
            radius = particle.size * (1.25 if self.scanning else 1.0)
            self.canvas.create_oval(
                particle.x - radius,
                particle.y - radius,
                particle.x + radius,
                particle.y + radius,
                fill=COLORS["glow"],
                outline="",
                tags="field",
            )

        self.after(90, self._animate)

    def _draw_ghost_words(self, width: int, height: int) -> None:
        positions = [(0.19, 0.22), (0.70, 0.19), (0.52, 0.37), (0.82, 0.58), (0.28, 0.76), (0.64, 0.82), (0.42, 0.16)]
        for word, (x_ratio, y_ratio) in zip(UI_TEXT["ghost_words"], positions):
            self.canvas.create_text(
                width * x_ratio,
                height * y_ratio,
                text=word,
                fill="#0F1219",
                font=("BIZ UDPGothic", 18),
                tags="field",
            )

    def _start_scan(self) -> None:
        if self.scanning:
            return
        if not self.memory_folder:
            self.status_var.set(UI_TEXT["memory_folder_missing"])
            self._render_memory_state()
            return

        self.scanning = True
        self.status_var.set(UI_TEXT["status_scanning"])
        self.summary_var.set(UI_TEXT["status_scanning"])
        self.scan_button.configure(text=UI_TEXT["button_scanning"], state="disabled")
        self.open_output_button.configure(state="disabled")
        self._render_empty_cards()

        self.scan_thread = threading.Thread(target=self._scan_worker, args=(self.memory_folder,), daemon=True)
        self.scan_thread.start()

    def _scan_worker(self, memory_folder: Path) -> None:
        try:
            documents, skipped = scan_memory(memory_folder)
            result = analyze_documents(documents, memory_folder, skipped_files=skipped)
            output_path = write_suggestion(memory_folder, result)
            self.events.put(("scan_done", (result, output_path)))
        except Exception as exc:  # noqa: BLE001
            self.events.put(("scan_error", exc))

    def _poll_events(self) -> None:
        try:
            while True:
                event_name, payload = self.events.get_nowait()
                if event_name == "scan_done":
                    result, output_path = payload
                    assert isinstance(result, AnalysisResult)
                    assert isinstance(output_path, Path)
                    self._handle_scan_done(result, output_path)
                elif event_name == "scan_error":
                    self._handle_scan_error(payload)
        except queue.Empty:
            pass
        self.after(100, self._poll_events)

    def _handle_scan_done(self, result: AnalysisResult, output_path: Path) -> None:
        self.scanning = False
        self.output_path = output_path
        self.scan_button.configure(text=UI_TEXT["button_scan"], state="normal")
        self.open_output_button.configure(state="normal")
        self.status_var.set(UI_TEXT["status_complete"] if result.traces else UI_TEXT["status_no_trace"])
        self.summary_var.set(
            UI_TEXT["summary_scan"].format(
                files=result.scanned_files,
                skipped=result.skipped_files,
                traces=len(result.traces),
            )
        )
        self._render_cards(result)

    def _handle_scan_error(self, payload: object) -> None:
        self.scanning = False
        self.scan_button.configure(text=UI_TEXT["button_scan"], state="normal")
        self.status_var.set(UI_TEXT["status_error"])
        self.summary_var.set(str(payload))
        messagebox.showerror(UI_TEXT["dialog_title"], f"{UI_TEXT['status_error']}\n{payload}")

    def _render_empty_cards(self) -> None:
        self._clear_cards()
        card = self._card_frame()
        card.pack(fill="x", pady=(0, 10))
        tk.Label(
            card,
            text=UI_TEXT["card_empty"],
            fg=COLORS["muted"],
            bg=COLORS["panel"],
            font=FONT_JP_SMALL,
        ).pack(anchor="w")

    def _render_cards(self, result: AnalysisResult) -> None:
        self._clear_cards()
        if not result.fragments:
            card = self._card_frame()
            card.pack(fill="x", pady=(0, 10))
            tk.Label(
                card,
                text=UI_TEXT["status_no_trace"],
                fg=COLORS["text"],
                bg=COLORS["panel"],
                font=FONT_JP,
            ).pack(anchor="w")
            tk.Label(
                card,
                text=result.suggestion,
                fg=COLORS["muted"],
                bg=COLORS["panel"],
                font=FONT_JP_SMALL,
                wraplength=560,
                justify="left",
            ).pack(anchor="w", pady=(8, 0))
            return

        for fragment in result.fragments[:5]:
            card = self._card_frame()
            card.pack(fill="x", pady=(0, 10))
            top = tk.Frame(card, bg=COLORS["panel"])
            top.pack(fill="x")
            tk.Label(
                top,
                text=fragment.heat_word,
                fg=COLORS["heat"],
                bg=COLORS["panel"],
                font=FONT_JP,
            ).pack(side="left")
            tk.Label(
                top,
                text=UI_TEXT["card_score"].format(score=fragment.score),
                fg=COLORS["muted"],
                bg=COLORS["panel"],
                font=FONT_MONO,
            ).pack(side="right")
            tk.Label(
                card,
                text=f"{UI_TEXT['card_file']}: {Path(fragment.relative_path).name}",
                fg=COLORS["text"],
                bg=COLORS["panel"],
                font=FONT_JP_SMALL,
                wraplength=560,
                justify="left",
            ).pack(anchor="w", pady=(6, 0))
            tk.Label(
                card,
                text=f"{UI_TEXT['card_excerpt']}: {fragment.excerpt}",
                fg=COLORS["muted"],
                bg=COLORS["panel"],
                font=FONT_JP_SMALL,
                wraplength=560,
                justify="left",
            ).pack(anchor="w", pady=(4, 0))
            tk.Label(
                card,
                text=f"{UI_TEXT['section_suggestion']}: {result.suggestion.splitlines()[0]}",
                fg=COLORS["glow"],
                bg=COLORS["panel"],
                font=FONT_JP_SMALL,
                wraplength=560,
                justify="left",
            ).pack(anchor="w", pady=(6, 0))

    def _card_frame(self) -> tk.Frame:
        return tk.Frame(
            self.cards_frame,
            bg=COLORS["panel"],
            highlightbackground=COLORS["line"],
            highlightthickness=1,
            padx=14,
            pady=12,
        )

    def _clear_cards(self) -> None:
        for child in self.cards_frame.winfo_children():
            child.destroy()

    def _choose_memory_folder(self) -> None:
        selected = filedialog.askdirectory(title=UI_TEXT["dialog_choose_memory"])
        if not selected:
            return
        folder = existing_folder(selected)
        if not folder:
            self.status_var.set(UI_TEXT["memory_folder_missing"])
            return
        self.memory_folder = folder
        self.config_data.memory_folder = str(folder)
        self.config_store.save(self.config_data)
        self.status_var.set(UI_TEXT["status_idle"])
        self._render_memory_state()

    def _open_output(self) -> None:
        if self.output_path:
            open_path(self.output_path.parent)
            return
        if self.memory_folder:
            open_path(self.memory_folder / "OIKAWA" / "suggestions")

    def _finish_launch_check(self) -> None:
        print(UI_TEXT["launch_check_ok"])
        self.destroy()

    def _finish_gui_smoke(self) -> None:
        print(UI_TEXT["gui_smoke_ok"])
        self.destroy()


def run_gui(launch_check: bool = False, gui_smoke_seconds: float = 0.0, memory_folder_override: str = "") -> int:
    app = OikawaApp(
        launch_check=launch_check,
        gui_smoke_seconds=gui_smoke_seconds,
        memory_folder_override=memory_folder_override,
    )
    app.mainloop()
    return 0


def run_scan_check(memory_folder: str) -> int:
    root = existing_folder(memory_folder)
    if not root:
        raise RuntimeError(UI_TEXT["memory_folder_missing"])
    documents, skipped = scan_memory(root)
    result = analyze_documents(documents, root, skipped_files=skipped)
    output_path = write_suggestion(root, result)
    print(UI_TEXT["scan_check_ok"])
    print(output_path)
    return 0


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--launch-check", action="store_true")
    parser.add_argument("--gui-smoke-seconds", type=float, default=0.0)
    parser.add_argument("--memory-folder", default="")
    parser.add_argument("--scan-check", default="")
    args = parser.parse_args()

    if args.scan_check:
        return run_scan_check(args.scan_check)

    return run_gui(
        launch_check=args.launch_check,
        gui_smoke_seconds=args.gui_smoke_seconds,
        memory_folder_override=args.memory_folder,
    )


if __name__ == "__main__":
    raise SystemExit(main())
