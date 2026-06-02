# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
import tkinter as tk
from pathlib import Path
from tkinter import font as tkfont


APP_NAME = "DakeAdvancedTimer"
WINDOW_TITLE = "アドバンスドタイマー"
COPYRIGHT = "© 2026 DAKE Series"

UI_TEXT = {
    "main_title": "時間を始める",
    "main_description": "20年前の自分に会いに行くための、静かなタイマーです。",
    "center_helper": "約束を、いま始める。",
    "finished_message": "戻ってきました。",
    "mode_focus": "25分集中",
    "mode_short_break": "5分休憩",
    "mode_long_break": "15分休憩",
    "mode_custom": "カスタム",
    "mode_current_format": "現在モード: {mode}",
    "custom_minutes_label": "カスタム分数",
    "minutes_suffix": "分",
    "button_start": "開始",
    "button_pause": "一時停止",
    "button_resume": "再開",
    "button_reset": "リセット",
    "button_again": "もう一回",
    "status_idle": "準備できています。",
    "status_running": "進行中です。",
    "status_paused": "一時停止中です。",
    "status_finished": "完了しました。",
    "status_error": "入力を確認してください。",
    "error_custom_required": "1〜180分の数字を入力してください。",
    "error_custom_range": "カスタム分数は1〜180分で入力してください。",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_separator": " / ",
    "footer_copyright": COPYRIGHT,
}

STATE_IDLE = "idle"
STATE_RUNNING = "running"
STATE_PAUSED = "paused"
STATE_FINISHED = "finished"
STATE_ERROR = "error"

MIN_CUSTOM_MINUTES = 1
MAX_CUSTOM_MINUTES = 180

MODES = {
    "focus": {"label_key": "mode_focus", "minutes": 25},
    "short_break": {"label_key": "mode_short_break", "minutes": 5},
    "long_break": {"label_key": "mode_long_break", "minutes": 15},
    "custom": {"label_key": "mode_custom", "minutes": None},
}

COLORS = {
    "background": "#F6F7F9",
    "panel": "#FFFFFF",
    "text": "#1E2430",
    "muted": "#667085",
    "border": "#DDE3EA",
    "accent": "#2F6FED",
    "accent_hover": "#255BC4",
    "accent_text": "#FFFFFF",
    "quiet": "#EEF2F7",
    "error": "#B42318",
    "success": "#157347",
}

FONT_FALLBACKS = ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo", "TkDefaultFont")


def get_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def get_project_root(base_dir: Path) -> Path:
    for candidate in (base_dir, *base_dir.parents):
        if (candidate / "02_assets" / "dake_icon.ico").exists():
            return candidate
    return base_dir.parents[1] if len(base_dir.parents) >= 2 else base_dir


def get_icon_path() -> Path:
    return get_project_root(get_base_dir()) / "02_assets" / "dake_icon.ico"


def choose_font_family(root: tk.Misc) -> str:
    available = set(tkfont.families(root))
    for family in FONT_FALLBACKS:
        if family in available:
            return family
    return "TkDefaultFont"


def format_seconds(seconds: int) -> str:
    minutes, remaining_seconds = divmod(max(0, seconds), 60)
    return f"{minutes:02d}:{remaining_seconds:02d}"


class AdvancedTimerApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.font_family = choose_font_family(self)
        self.state = STATE_IDLE
        self.after_id: str | None = None
        self.selected_mode = tk.StringVar(value="focus")
        self.custom_minutes = tk.StringVar(value="25")
        self.remaining_seconds = self._selected_duration_seconds()

        self.time_var = tk.StringVar(value=format_seconds(self.remaining_seconds))
        self.mode_var = tk.StringVar(value=self._current_mode_text())
        self.status_var = tk.StringVar(value=UI_TEXT["status_idle"])
        self.finished_var = tk.StringVar(value="")

        self.mode_buttons: dict[str, tk.Button] = {}

        self.title(WINDOW_TITLE)
        self.geometry("480x560")
        self.minsize(480, 560)
        self.maxsize(480, 560)
        self.resizable(False, False)
        self.configure(bg=COLORS["background"])
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._set_window_icon()
        self._build_ui()
        self.custom_minutes.trace_add("write", self._handle_custom_input_change)
        self._render_state()

    def _set_window_icon(self) -> None:
        icon_path = get_icon_path()
        if not icon_path.exists():
            return
        try:
            self.iconbitmap(str(icon_path))
            self.iconbitmap(default=str(icon_path))
        except tk.TclError:
            return

    def _build_ui(self) -> None:
        shell = tk.Frame(self, bg=COLORS["background"])
        shell.pack(fill="both", expand=True, padx=28, pady=24)

        self._build_header(shell)
        self._build_mode_selector(shell)
        self._build_timer_panel(shell)
        self._build_controls(shell)
        self._build_footer(shell)

    def _build_header(self, parent: tk.Widget) -> None:
        header = tk.Frame(parent, bg=COLORS["background"])
        header.pack(fill="x")

        tk.Label(
            header,
            text=UI_TEXT["main_title"],
            bg=COLORS["background"],
            fg=COLORS["text"],
            font=(self.font_family, 22, "bold"),
            anchor="center",
        ).pack(fill="x")

        tk.Label(
            header,
            text=UI_TEXT["main_description"],
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=(self.font_family, 10),
            wraplength=400,
            justify="center",
        ).pack(fill="x", pady=(8, 0))

    def _build_mode_selector(self, parent: tk.Widget) -> None:
        modes = tk.Frame(parent, bg=COLORS["background"])
        modes.pack(fill="x", pady=(24, 0))
        modes.columnconfigure((0, 1), weight=1, uniform="mode")

        order = ("focus", "short_break", "long_break", "custom")
        for index, mode_key in enumerate(order):
            button = tk.Button(
                modes,
                text=UI_TEXT[MODES[mode_key]["label_key"]],
                command=lambda key=mode_key: self._select_mode(key),
                bd=0,
                relief="flat",
                cursor="hand2",
                font=(self.font_family, 10, "bold"),
                padx=12,
                pady=10,
            )
            button.grid(row=index // 2, column=index % 2, sticky="ew", padx=4, pady=4)
            self.mode_buttons[mode_key] = button

        custom_row = tk.Frame(parent, bg=COLORS["background"])
        custom_row.pack(fill="x", pady=(10, 0))
        custom_row.columnconfigure(1, weight=1)

        tk.Label(
            custom_row,
            text=UI_TEXT["custom_minutes_label"],
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=(self.font_family, 9),
        ).grid(row=0, column=0, sticky="w")

        self.custom_entry = tk.Entry(
            custom_row,
            textvariable=self.custom_minutes,
            justify="right",
            font=(self.font_family, 11),
            bd=0,
            relief="flat",
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            highlightcolor=COLORS["accent"],
            width=8,
        )
        self.custom_entry.grid(row=0, column=1, sticky="e", padx=(10, 6), ipady=6)

        tk.Label(
            custom_row,
            text=UI_TEXT["minutes_suffix"],
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=(self.font_family, 9),
        ).grid(row=0, column=2, sticky="e")

    def _build_timer_panel(self, parent: tk.Widget) -> None:
        panel = tk.Frame(
            parent,
            bg=COLORS["panel"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
            bd=0,
        )
        panel.pack(fill="x", pady=(22, 18))

        tk.Label(
            panel,
            text=UI_TEXT["center_helper"],
            bg=COLORS["panel"],
            fg=COLORS["muted"],
            font=(self.font_family, 11),
        ).pack(fill="x", padx=20, pady=(22, 4))

        tk.Label(
            panel,
            textvariable=self.time_var,
            bg=COLORS["panel"],
            fg=COLORS["text"],
            font=(self.font_family, 56, "bold"),
        ).pack(fill="x", padx=20, pady=(0, 6))

        tk.Label(
            panel,
            textvariable=self.mode_var,
            bg=COLORS["panel"],
            fg=COLORS["muted"],
            font=(self.font_family, 10),
        ).pack(fill="x", padx=20, pady=(0, 12))

        tk.Label(
            panel,
            textvariable=self.finished_var,
            bg=COLORS["panel"],
            fg=COLORS["success"],
            font=(self.font_family, 12, "bold"),
            height=1,
        ).pack(fill="x", padx=20, pady=(0, 20))

    def _build_controls(self, parent: tk.Widget) -> None:
        controls = tk.Frame(parent, bg=COLORS["background"])
        controls.pack(fill="x")

        self.primary_button = tk.Button(
            controls,
            command=self._handle_primary_action,
            bd=0,
            relief="flat",
            cursor="hand2",
            font=(self.font_family, 14, "bold"),
            padx=18,
            pady=14,
            bg=COLORS["accent"],
            activebackground=COLORS["accent_hover"],
            fg=COLORS["accent_text"],
            activeforeground=COLORS["accent_text"],
        )
        self.primary_button.pack(fill="x")

        self.reset_button = tk.Button(
            controls,
            text=UI_TEXT["button_reset"],
            command=self._reset_timer,
            bd=0,
            relief="flat",
            cursor="hand2",
            font=(self.font_family, 10),
            padx=14,
            pady=10,
            bg=COLORS["quiet"],
            activebackground=COLORS["border"],
            fg=COLORS["text"],
            activeforeground=COLORS["text"],
        )
        self.reset_button.pack(fill="x", pady=(10, 0))

        tk.Label(
            controls,
            textvariable=self.status_var,
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=(self.font_family, 9),
        ).pack(fill="x", pady=(12, 0))

    def _build_footer(self, parent: tk.Widget) -> None:
        footer = tk.Label(
            parent,
            text=f"{UI_TEXT['footer_left']}{UI_TEXT['footer_separator']}{UI_TEXT['footer_copyright']}",
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=(self.font_family, 8),
        )
        footer.pack(side="bottom", fill="x", pady=(12, 0))

    def _select_mode(self, mode_key: str) -> None:
        if self.state in (STATE_RUNNING, STATE_PAUSED):
            return
        self.selected_mode.set(mode_key)
        self._clear_after()
        try:
            self.remaining_seconds = self._selected_duration_seconds()
        except ValueError:
            self._set_error(UI_TEXT["error_custom_required"])
            return
        self._set_state(STATE_IDLE)

    def _selected_duration_seconds(self) -> int:
        mode_key = self.selected_mode.get()
        minutes = MODES[mode_key]["minutes"]
        if minutes is None:
            minutes = self._parse_custom_minutes()
        return int(minutes) * 60

    def _parse_custom_minutes(self) -> int:
        raw_value = self.custom_minutes.get().strip()
        if not raw_value:
            raise ValueError(UI_TEXT["error_custom_required"])
        try:
            minutes = int(raw_value)
        except ValueError as exc:
            raise ValueError(UI_TEXT["error_custom_required"]) from exc
        if minutes < MIN_CUSTOM_MINUTES or minutes > MAX_CUSTOM_MINUTES:
            raise ValueError(UI_TEXT["error_custom_range"])
        return minutes

    def _handle_custom_input_change(self, *_args: object) -> None:
        if self.selected_mode.get() != "custom" or self.state in (STATE_RUNNING, STATE_PAUSED):
            return
        try:
            self.remaining_seconds = self._selected_duration_seconds()
        except ValueError:
            if self.state != STATE_ERROR:
                self.time_var.set(format_seconds(0))
            return
        self._set_state(STATE_IDLE)

    def _handle_primary_action(self) -> None:
        if self.state == STATE_RUNNING:
            self._pause_timer()
            return
        if self.state == STATE_PAUSED:
            self._resume_timer()
            return
        self._start_timer()

    def _start_timer(self) -> None:
        try:
            self.remaining_seconds = self._selected_duration_seconds()
        except ValueError as exc:
            self._set_error(str(exc))
            return
        self.finished_var.set("")
        self._set_state(STATE_RUNNING)
        self._schedule_tick()

    def _resume_timer(self) -> None:
        if self.remaining_seconds <= 0:
            self._start_timer()
            return
        self.finished_var.set("")
        self._set_state(STATE_RUNNING)
        self._schedule_tick()

    def _pause_timer(self) -> None:
        self._clear_after()
        self._set_state(STATE_PAUSED)

    def _reset_timer(self) -> None:
        self._clear_after()
        try:
            self.remaining_seconds = self._selected_duration_seconds()
        except ValueError as exc:
            self._set_error(str(exc))
            return
        self.finished_var.set("")
        self._set_state(STATE_IDLE)

    def _schedule_tick(self) -> None:
        self._clear_after()
        self.after_id = self.after(1000, self._tick)

    def _tick(self) -> None:
        self.after_id = None
        if self.state != STATE_RUNNING:
            return
        self.remaining_seconds = max(0, self.remaining_seconds - 1)
        self._update_display()
        if self.remaining_seconds <= 0:
            self._finish_timer()
            return
        self._schedule_tick()

    def _finish_timer(self) -> None:
        self._clear_after()
        self.remaining_seconds = 0
        self.finished_var.set(UI_TEXT["finished_message"])
        self._set_state(STATE_FINISHED)

    def _set_error(self, message: str) -> None:
        self._clear_after()
        self.status_var.set(message)
        self.state = STATE_ERROR
        self._render_state()

    def _set_state(self, state: str) -> None:
        self.state = state
        status_key = {
            STATE_IDLE: "status_idle",
            STATE_RUNNING: "status_running",
            STATE_PAUSED: "status_paused",
            STATE_FINISHED: "status_finished",
            STATE_ERROR: "status_error",
        }[state]
        self.status_var.set(UI_TEXT[status_key])
        self._render_state()

    def _render_state(self) -> None:
        self._update_display()
        primary_text = {
            STATE_IDLE: UI_TEXT["button_start"],
            STATE_RUNNING: UI_TEXT["button_pause"],
            STATE_PAUSED: UI_TEXT["button_resume"],
            STATE_FINISHED: UI_TEXT["button_again"],
            STATE_ERROR: UI_TEXT["button_start"],
        }[self.state]
        self.primary_button.configure(text=primary_text)

        primary_enabled = self.state != STATE_ERROR
        self.primary_button.configure(state="normal" if primary_enabled else "disabled")
        self.reset_button.configure(state="normal")

        selector_enabled = self.state not in (STATE_RUNNING, STATE_PAUSED)
        for mode_key, button in self.mode_buttons.items():
            is_selected = mode_key == self.selected_mode.get()
            button.configure(
                state="normal" if selector_enabled else "disabled",
                bg=COLORS["accent"] if is_selected else COLORS["quiet"],
                activebackground=COLORS["accent_hover"] if is_selected else COLORS["border"],
                fg=COLORS["accent_text"] if is_selected else COLORS["text"],
                activeforeground=COLORS["accent_text"] if is_selected else COLORS["text"],
            )

        custom_enabled = selector_enabled and self.selected_mode.get() == "custom"
        self.custom_entry.configure(state="normal" if custom_enabled else "disabled")

    def _update_display(self) -> None:
        self.time_var.set(format_seconds(self.remaining_seconds))
        self.mode_var.set(self._current_mode_text())

    def _current_mode_text(self) -> str:
        mode_key = self.selected_mode.get()
        mode_label = UI_TEXT[MODES[mode_key]["label_key"]]
        if mode_key == "custom":
            raw_value = self.custom_minutes.get().strip()
            if raw_value:
                mode_label = f"{mode_label} {raw_value}{UI_TEXT['minutes_suffix']}"
        return UI_TEXT["mode_current_format"].format(mode=mode_label)

    def _clear_after(self) -> None:
        if self.after_id is None:
            return
        try:
            self.after_cancel(self.after_id)
        except tk.TclError:
            pass
        self.after_id = None

    def _on_close(self) -> None:
        self._clear_after()
        self.destroy()


def main() -> None:
    app = AdvancedTimerApp()
    app.mainloop()


if __name__ == "__main__":
    main()
