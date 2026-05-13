# -*- coding: utf-8 -*-
from __future__ import annotations

import ctypes
import json
import re
import sys
import webbrowser
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk


APP_NAME = "Dakeずっとメモ"
WINDOW_TITLE = "Dakeずっとメモ"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "main_title": "ずっとメモ",
    "main_description": "1本のメモを、列に流して見渡します。",
    "label_columns": "列数",
    "column_1": "1",
    "column_2": "2",
    "column_3": "3",
    "column_4": "4",
    "button_refresh": "リフレッシュ",
    "button_copy": "コピー",
    "label_input": "メモ入力",
    "label_preview": "段組表示",
    "input_placeholder": "ここにメモを書きます。",
    "preview_empty": "メモ入力に書くと、ここに列で表示します。",
    "status_idle": "自動保存します",
    "status_saved": "保存しました",
    "status_save_error": "保存できませんでした",
    "status_load_error": "保存データを読み込めませんでした",
    "status_refreshed": "表示を更新しました",
    "status_copied": "コピーしました",
    "status_copy_empty": "コピーするメモがありません",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_tagline": "止まらない、迷わない、すぐ終わる。",
    "footer_link_1": "戸建買取査定",
    "footer_link_2": "Instagram",
    "footer_separator": " ｜ ",
    "footer_copyright": COPYRIGHT,
}

LINK_URLS = {
    "footer_link_1": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "footer_link_2": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

COLORS = {
    "background": "#F6F7F9",
    "surface": "#FFFFFF",
    "surface_soft": "#F9FAFB",
    "text": "#1E2430",
    "muted": "#667085",
    "quiet": "#98A2B3",
    "border": "#E5EAF0",
    "accent": "#2F6FED",
    "accent_hover": "#2458BF",
    "button_secondary": "#FFFFFF",
    "button_secondary_hover": "#F2F4F7",
    "selection": "#D7E7FF",
}

FONT_CANDIDATES = ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo")
WINDOW_SIZE = "1080x700"
WINDOW_MIN_SIZE = (820, 560)
FOOTER_NARROW_WIDTH = 930
CONFIG_FILE_NAME = "DAKE_Column_Memo_config.json"
CONFIG_VERSION = 1
DEFAULT_COLUMNS = 2
SAVE_DEBOUNCE_MS = 420
PREVIEW_DEBOUNCE_MS = 120
STATUS_RESET_MS = 1400
APP_USER_MODEL_ID = "Shimarisu.DakeColumnMemo"


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APP_DIR = app_dir()
CONFIG_PATH = APP_DIR / CONFIG_FILE_NAME


def set_windows_app_id() -> None:
    if not sys.platform.startswith("win"):
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)
    except Exception:
        return


def icon_candidates() -> list[Path]:
    base = app_dir()
    candidates = [
        base / ".." / ".." / "02_assets" / "dake_icon.ico",
        base / "dake_icon.ico",
    ]
    bundle_dir = getattr(sys, "_MEIPASS", None)
    if bundle_dir:
        candidates.insert(0, Path(bundle_dir) / "dake_icon.ico")
    return candidates


def choose_font(root: tk.Tk) -> str:
    available = set(tkfont.families(root))
    for family in FONT_CANDIDATES:
        if family in available:
            return family
    return "TkDefaultFont"


def parse_geometry(value: object) -> str:
    if not isinstance(value, str):
        return WINDOW_SIZE
    match = re.fullmatch(r"(\d+)x(\d+)([+-]\d+)?([+-]\d+)?", value)
    if not match:
        return WINDOW_SIZE
    width = max(int(match.group(1)), WINDOW_MIN_SIZE[0])
    height = max(int(match.group(2)), WINDOW_MIN_SIZE[1])
    x = match.group(3) or ""
    y = match.group(4) or ""
    return f"{width}x{height}{x}{y}"


def clamp_columns(value: object) -> int:
    try:
        columns = int(value)
    except (TypeError, ValueError):
        return DEFAULT_COLUMNS
    return min(max(columns, 1), 4)


def load_config() -> tuple[dict[str, object], bool]:
    if not CONFIG_PATH.exists():
        return {}, False
    try:
        with CONFIG_PATH.open("r", encoding="utf-8") as file:
            data = json.load(file)
    except Exception:
        return {}, True
    return data if isinstance(data, dict) else {}, False


class ColumnPreview(tk.Frame):
    def __init__(
        self,
        parent: tk.Widget,
        *,
        preview_font: tkfont.Font,
        muted_font: tkfont.Font,
    ) -> None:
        super().__init__(
            parent,
            bg=COLORS["surface"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
        )
        self.preview_font = preview_font
        self.muted_font = muted_font
        self.column_count = DEFAULT_COLUMNS
        self.memo_text = ""
        self.lines: list[str] = [""]
        self.first_line = 0
        self.lines_per_column = 1
        self.page_capacity = 1

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.canvas = tk.Canvas(
            self,
            bg=COLORS["surface"],
            bd=0,
            highlightthickness=0,
            relief="flat",
        )
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self._on_scrollbar)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")

        self.canvas.bind("<Configure>", self._on_configure)
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Button-4>", self._on_wheel_up)
        self.canvas.bind("<Button-5>", self._on_wheel_down)
        self.canvas.bind("<Enter>", lambda _event: self.canvas.focus_set())

    def set_columns(self, column_count: int) -> None:
        self.column_count = clamp_columns(column_count)
        self._reflow()

    def set_text(self, text: str) -> None:
        self.memo_text = text
        self._reflow()

    def refresh(self) -> None:
        self._reflow()

    def _on_configure(self, _event: tk.Event) -> None:
        self._reflow()

    def _on_mousewheel(self, event: tk.Event) -> str:
        delta = -1 if event.delta > 0 else 1
        self._scroll_by(delta * 3)
        return "break"

    def _on_wheel_up(self, _event: tk.Event) -> str:
        self._scroll_by(-3)
        return "break"

    def _on_wheel_down(self, _event: tk.Event) -> str:
        self._scroll_by(3)
        return "break"

    def _on_scrollbar(self, *args: str) -> None:
        if not args:
            return
        if args[0] == "moveto" and len(args) >= 2:
            total = max(len(self.lines), 1)
            self.first_line = int(float(args[1]) * total)
        elif args[0] == "scroll" and len(args) >= 3:
            amount = int(args[1])
            unit = self.lines_per_column if args[2] == "pages" else 1
            self.first_line += amount * unit
        self._clamp_first_line()
        self._render()

    def _scroll_by(self, amount: int) -> None:
        self.first_line += amount
        self._clamp_first_line()
        self._render()

    def _reflow(self) -> None:
        width = max(self.canvas.winfo_width(), 240)
        height = max(self.canvas.winfo_height(), 180)
        pad_x = 18
        gap = 18 if self.column_count > 1 else 0
        usable_width = width - pad_x * 2 - gap * (self.column_count - 1)
        column_width = max(int(usable_width / self.column_count), 80)
        text_width = max(column_width - 16, 48)
        self.lines = self._wrap_text(self.memo_text, text_width) or [""]

        line_space = self.preview_font.metrics("linespace")
        self.lines_per_column = max((height - 28) // max(line_space + 5, 1), 1)
        self.page_capacity = max(self.lines_per_column * self.column_count, 1)
        self._clamp_first_line()
        self._render()

    def _wrap_text(self, text: str, max_width: int) -> list[str]:
        normalized = text.replace("\r\n", "\n").replace("\r", "\n").replace("\t", "    ")
        if normalized == "":
            return [""]
        wrapped: list[str] = []
        for source_line in normalized.split("\n"):
            wrapped.extend(self._wrap_line(source_line, max_width))
        return wrapped

    def _wrap_line(self, source_line: str, max_width: int) -> list[str]:
        if source_line == "":
            return [""]
        result: list[str] = []
        current = ""
        for char in source_line:
            candidate = current + char
            if current and self.preview_font.measure(candidate) > max_width:
                result.append(current.rstrip())
                current = char.lstrip()
            else:
                current = candidate
        result.append(current)
        return result

    def _clamp_first_line(self) -> None:
        max_first = max(len(self.lines) - self.page_capacity, 0)
        self.first_line = min(max(self.first_line, 0), max_first)

    def _render(self) -> None:
        self.canvas.delete("all")
        width = max(self.canvas.winfo_width(), 240)
        height = max(self.canvas.winfo_height(), 180)
        pad_x = 18
        pad_y = 14
        gap = 18 if self.column_count > 1 else 0
        line_step = self.preview_font.metrics("linespace") + 5
        usable_width = width - pad_x * 2 - gap * (self.column_count - 1)
        column_width = max(int(usable_width / self.column_count), 80)

        if self.memo_text.strip() == "":
            self.canvas.create_text(
                width // 2,
                height // 2,
                text=UI_TEXT["preview_empty"],
                fill=COLORS["quiet"],
                font=self.muted_font,
                anchor="center",
                width=max(width - 64, 120),
            )
            self.scrollbar.set(0, 1)
            return

        for column_index in range(self.column_count):
            x = pad_x + column_index * (column_width + gap)
            if column_index > 0:
                separator_x = x - gap // 2
                self.canvas.create_line(
                    separator_x,
                    pad_y,
                    separator_x,
                    height - pad_y,
                    fill=COLORS["border"],
                )
            start = self.first_line + column_index * self.lines_per_column
            stop = min(start + self.lines_per_column, len(self.lines))
            y = pad_y
            for line in self.lines[start:stop]:
                self.canvas.create_text(
                    x + 2,
                    y,
                    text=line,
                    fill=COLORS["text"],
                    font=self.preview_font,
                    anchor="nw",
                    width=max(column_width - 12, 48),
                )
                y += line_step

        total = max(len(self.lines), 1)
        if total <= self.page_capacity:
            self.scrollbar.set(0, 1)
            return
        first = self.first_line / total
        last = min((self.first_line + self.page_capacity) / total, 1)
        self.scrollbar.set(first, last)


class ColumnMemoApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.config_data, self.load_failed = load_config()
        self.font_family = choose_font(root)
        self.columns = clamp_columns(self.config_data.get("columns", DEFAULT_COLUMNS))
        self.column_var = tk.IntVar(value=self.columns)
        self.save_after_id: str | None = None
        self.preview_after_id: str | None = None
        self.status_after_id: str | None = None
        self.placeholder_active = False
        self.footer: tk.Frame | None = None
        self.footer_mode: str | None = None

        self.fonts = {
            "title": tkfont.Font(root, family=self.font_family, size=20, weight="bold"),
            "description": tkfont.Font(root, family=self.font_family, size=10),
            "label": tkfont.Font(root, family=self.font_family, size=10, weight="bold"),
            "input": tkfont.Font(root, family=self.font_family, size=11),
            "preview": tkfont.Font(root, family=self.font_family, size=10),
            "button": tkfont.Font(root, family=self.font_family, size=10, weight="bold"),
            "status": tkfont.Font(root, family=self.font_family, size=9),
            "footer": tkfont.Font(root, family=self.font_family, size=8),
        }

        self._configure_root()
        self._build_ui()
        self._apply_icon()
        self._load_initial_text()

    def _configure_root(self) -> None:
        self.root.title(WINDOW_TITLE)
        self.root.geometry(parse_geometry(self.config_data.get("window_geometry")))
        self.root.minsize(*WINDOW_MIN_SIZE)
        self.root.configure(bg=COLORS["background"])
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _apply_icon(self) -> None:
        for candidate in icon_candidates():
            try:
                icon_path = candidate.resolve()
            except Exception:
                icon_path = candidate
            if not icon_path.exists():
                continue
            try:
                self.root.iconbitmap(str(icon_path))
                return
            except tk.TclError:
                continue

    def _build_ui(self) -> None:
        outer = tk.Frame(self.root, bg=COLORS["background"])
        outer.pack(fill="both", expand=True, padx=28, pady=(24, 16))

        self._build_header(outer)
        self._build_body(outer)
        self._build_status(outer)
        self._build_footer(outer)
        self.root.bind("<Configure>", self._on_root_configure, add="+")

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
            anchor="w",
        ).pack(anchor="w")
        tk.Label(
            title_area,
            text=UI_TEXT["main_description"],
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=self.fonts["description"],
            anchor="w",
        ).pack(anchor="w", pady=(6, 0))

        actions = tk.Frame(header, bg=COLORS["background"])
        actions.grid(row=0, column=1, sticky="e", padx=(18, 0))

        columns = tk.Frame(actions, bg=COLORS["background"])
        columns.pack(side="left", padx=(0, 12))
        tk.Label(
            columns,
            text=UI_TEXT["label_columns"],
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=self.fonts["description"],
        ).pack(side="left", padx=(0, 8))

        for value in (1, 2, 3, 4):
            self._make_column_button(columns, value).pack(side="left", padx=(0, 4))

        self._make_button(
            actions,
            UI_TEXT["button_refresh"],
            self.refresh_preview,
            primary=False,
        ).pack(side="left", padx=(0, 8))
        self._make_button(
            actions,
            UI_TEXT["button_copy"],
            self.copy_memo,
            primary=True,
        ).pack(side="left")

    def _build_body(self, parent: tk.Widget) -> None:
        body = tk.Frame(parent, bg=COLORS["background"])
        body.pack(fill="both", expand=True)
        body.grid_columnconfigure(0, weight=4, uniform="memo")
        body.grid_columnconfigure(1, weight=6, uniform="memo")
        body.grid_rowconfigure(0, weight=1)

        editor_shell = self._make_panel(body, UI_TEXT["label_input"])
        preview_shell = self._make_panel(body, UI_TEXT["label_preview"])
        editor_shell.grid(row=0, column=0, sticky="nsew", padx=(0, 14))
        preview_shell.grid(row=0, column=1, sticky="nsew")

        editor_shell.grid_columnconfigure(0, weight=1)
        editor_shell.grid_rowconfigure(1, weight=1)
        preview_shell.grid_columnconfigure(0, weight=1)
        preview_shell.grid_rowconfigure(1, weight=1)

        input_frame = tk.Frame(
            editor_shell,
            bg=COLORS["surface"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
        )
        input_frame.grid(row=1, column=0, sticky="nsew")
        input_frame.grid_columnconfigure(0, weight=1)
        input_frame.grid_rowconfigure(0, weight=1)

        self.memo_input = tk.Text(
            input_frame,
            bg=COLORS["surface"],
            fg=COLORS["text"],
            insertbackground=COLORS["text"],
            selectbackground=COLORS["selection"],
            selectforeground=COLORS["text"],
            inactiveselectbackground=COLORS["selection"],
            font=self.fonts["input"],
            relief="flat",
            bd=0,
            padx=14,
            pady=12,
            wrap="word",
            undo=True,
            height=1,
        )
        input_scroll = ttk.Scrollbar(input_frame, orient="vertical", command=self.memo_input.yview)
        self.memo_input.configure(yscrollcommand=input_scroll.set)
        self.memo_input.grid(row=0, column=0, sticky="nsew")
        input_scroll.grid(row=0, column=1, sticky="ns")
        self.memo_input.bind("<<Modified>>", self._on_text_modified)
        self.memo_input.bind("<FocusIn>", self._on_input_focus_in)
        self.memo_input.bind("<Control-a>", self._select_all)

        self.preview = ColumnPreview(
            preview_shell,
            preview_font=self.fonts["preview"],
            muted_font=self.fonts["description"],
        )
        self.preview.grid(row=1, column=0, sticky="nsew")
        self.preview.set_columns(self.columns)

    def _make_panel(self, parent: tk.Widget, label_text: str) -> tk.Frame:
        panel = tk.Frame(parent, bg=COLORS["background"])
        tk.Label(
            panel,
            text=label_text,
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=self.fonts["label"],
            anchor="w",
        ).grid(row=0, column=0, sticky="w", pady=(0, 8))
        return panel

    def _build_status(self, parent: tk.Widget) -> None:
        self.status_label = tk.Label(
            parent,
            text=UI_TEXT["status_load_error"] if self.load_failed else UI_TEXT["status_idle"],
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=self.fonts["status"],
            anchor="w",
        )
        self.status_label.pack(fill="x", pady=(10, 0))

    def _build_footer(self, parent: tk.Widget) -> None:
        self.footer = tk.Frame(parent, bg=COLORS["background"])
        self.footer.pack(fill="x", pady=(12, 0))
        self._render_footer("wide")

    def _on_root_configure(self, event: tk.Event) -> None:
        if event.widget is not self.root:
            return
        next_mode = "narrow" if event.width < FOOTER_NARROW_WIDTH else "wide"
        self._render_footer(next_mode)

    def _render_footer(self, mode: str) -> None:
        if self.footer is None or self.footer_mode == mode:
            return
        self.footer_mode = mode
        for child in self.footer.winfo_children():
            child.destroy()

        if mode == "narrow":
            thought_line = tk.Frame(self.footer, bg=COLORS["background"])
            thought_line.pack(anchor="center")
            self._footer_thought_line(thought_line)
            link_line = tk.Frame(self.footer, bg=COLORS["background"])
            link_line.pack(anchor="center", pady=(4, 0))
            self._footer_link_line(link_line)
            return

        self.footer.grid_columnconfigure(0, weight=0)
        self.footer.grid_columnconfigure(1, weight=1)
        self.footer.grid_columnconfigure(2, weight=0)
        left = tk.Frame(self.footer, bg=COLORS["background"])
        left.grid(row=0, column=0, sticky="w")
        self._footer_thought_line(left)
        right = tk.Frame(self.footer, bg=COLORS["background"])
        right.grid(row=0, column=2, sticky="e")
        self._footer_link_line(right)

    def _footer_thought_line(self, parent: tk.Widget) -> None:
        self._footer_text(parent, UI_TEXT["footer_left"])
        self._footer_text(parent, UI_TEXT["footer_separator"])
        self._footer_text(parent, UI_TEXT["footer_tagline"])

    def _footer_link_line(self, parent: tk.Widget) -> None:
        self._footer_link(parent, "footer_link_1")
        self._footer_text(parent, UI_TEXT["footer_separator"])
        self._footer_link(parent, "footer_link_2")
        self._footer_text(parent, UI_TEXT["footer_separator"])
        self._footer_text(parent, UI_TEXT["footer_copyright"])

    def _footer_text(self, parent: tk.Widget, value: str) -> None:
        tk.Label(
            parent,
            text=value,
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=self.fonts["footer"],
        ).pack(side="left")

    def _footer_link(self, parent: tk.Widget, text_key: str) -> None:
        label = tk.Label(
            parent,
            text=UI_TEXT[text_key],
            bg=COLORS["background"],
            fg=COLORS["muted"],
            cursor="hand2",
            font=self.fonts["footer"],
        )
        label.pack(side="left")
        label.bind("<Button-1>", lambda _event, key=text_key: webbrowser.open(LINK_URLS[key]))
        label.bind("<Enter>", lambda _event, target=label: target.configure(fg=COLORS["accent"]))
        label.bind("<Leave>", lambda _event, target=label: target.configure(fg=COLORS["muted"]))

    def _make_button(self, parent: tk.Widget, label: str, command, primary: bool) -> tk.Button:
        bg = COLORS["accent"] if primary else COLORS["button_secondary"]
        fg = COLORS["surface"] if primary else COLORS["text"]
        hover = COLORS["accent_hover"] if primary else COLORS["button_secondary_hover"]
        border = COLORS["accent"] if primary else COLORS["border"]
        button = tk.Button(
            parent,
            text=label,
            command=command,
            bg=bg,
            fg=fg,
            activebackground=hover,
            activeforeground=fg,
            font=self.fonts["button"],
            relief="flat",
            bd=0,
            padx=15,
            pady=8,
            cursor="hand2",
            highlightthickness=1,
            highlightbackground=border,
        )
        button.bind("<Enter>", lambda _event, target=button, color=hover: target.configure(bg=color))
        button.bind("<Leave>", lambda _event, target=button, color=bg: target.configure(bg=color))
        return button

    def _make_column_button(self, parent: tk.Widget, value: int) -> tk.Radiobutton:
        return tk.Radiobutton(
            parent,
            text=UI_TEXT[f"column_{value}"],
            variable=self.column_var,
            value=value,
            command=self._on_column_changed,
            indicatoron=False,
            bg=COLORS["button_secondary"],
            fg=COLORS["text"],
            selectcolor=COLORS["surface_soft"],
            activebackground=COLORS["button_secondary_hover"],
            activeforeground=COLORS["text"],
            font=self.fonts["button"],
            relief="flat",
            bd=0,
            width=3,
            padx=4,
            pady=7,
            cursor="hand2",
            highlightthickness=1,
            highlightbackground=COLORS["border"],
        )

    def _load_initial_text(self) -> None:
        memo_text = self.config_data.get("memo_text")
        if isinstance(memo_text, str) and memo_text:
            self.memo_input.insert("1.0", memo_text)
            self.memo_input.edit_modified(False)
            self.preview.set_text(memo_text)
            return
        self._show_placeholder()
        self.preview.set_text("")

    def _show_placeholder(self) -> None:
        self.placeholder_active = True
        self.memo_input.configure(fg=COLORS["quiet"])
        self.memo_input.insert("1.0", UI_TEXT["input_placeholder"])
        self.memo_input.edit_modified(False)

    def _hide_placeholder(self) -> None:
        if not self.placeholder_active:
            return
        self.placeholder_active = False
        self.memo_input.delete("1.0", "end")
        self.memo_input.configure(fg=COLORS["text"])
        self.memo_input.edit_modified(False)

    def _on_input_focus_in(self, _event: tk.Event) -> None:
        self._hide_placeholder()

    def _select_all(self, _event: tk.Event) -> str:
        self._hide_placeholder()
        self.memo_input.tag_add("sel", "1.0", "end-1c")
        self.memo_input.mark_set("insert", "1.0")
        self.memo_input.see("insert")
        return "break"

    def _on_text_modified(self, _event: tk.Event) -> None:
        if not self.memo_input.edit_modified():
            return
        self.memo_input.edit_modified(False)
        if self.placeholder_active:
            return
        self._schedule_preview()
        self._schedule_save()

    def _on_column_changed(self) -> None:
        self.columns = clamp_columns(self.column_var.get())
        self.preview.set_columns(self.columns)
        self._schedule_save()

    def _schedule_preview(self) -> None:
        if self.preview_after_id is not None:
            self.root.after_cancel(self.preview_after_id)
        self.preview_after_id = self.root.after(PREVIEW_DEBOUNCE_MS, self._update_preview)

    def _schedule_save(self) -> None:
        if self.save_after_id is not None:
            self.root.after_cancel(self.save_after_id)
        self.save_after_id = self.root.after(SAVE_DEBOUNCE_MS, self._save_now)

    def _update_preview(self) -> None:
        self.preview_after_id = None
        self.preview.set_text(self._current_memo_text())

    def _current_memo_text(self) -> str:
        if self.placeholder_active:
            return ""
        return self.memo_input.get("1.0", "end-1c")

    def _set_status(self, text: str, *, reset: bool = True) -> None:
        self.status_label.configure(text=text)
        if self.status_after_id is not None:
            self.root.after_cancel(self.status_after_id)
            self.status_after_id = None
        if reset:
            self.status_after_id = self.root.after(STATUS_RESET_MS, self._reset_status)

    def _reset_status(self) -> None:
        self.status_after_id = None
        self.status_label.configure(text=UI_TEXT["status_idle"])

    def _save_now(self) -> None:
        if self.save_after_id is not None:
            self.root.after_cancel(self.save_after_id)
            self.save_after_id = None
        payload = {
            "config_version": CONFIG_VERSION,
            "memo_text": self._current_memo_text(),
            "columns": self.columns,
            "window_geometry": self.root.geometry(),
            "last_updated": datetime.now().isoformat(timespec="seconds"),
        }
        try:
            CONFIG_PATH.write_text(
                json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )
        except Exception:
            self._set_status(UI_TEXT["status_save_error"])
            return
        self._set_status(UI_TEXT["status_saved"])

    def refresh_preview(self) -> None:
        self._update_preview()
        self._save_now()
        self._set_status(UI_TEXT["status_refreshed"])

    def copy_memo(self) -> None:
        text = self._current_memo_text()
        if text.strip() == "":
            self._set_status(UI_TEXT["status_copy_empty"])
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self._set_status(UI_TEXT["status_copied"])

    def _on_close(self) -> None:
        self._update_preview()
        self._save_now()
        self.root.destroy()


def main() -> None:
    set_windows_app_id()
    root = tk.Tk()
    ColumnMemoApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
