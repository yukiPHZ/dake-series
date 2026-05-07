# -*- coding: utf-8 -*-
from __future__ import annotations

import ctypes
import re
import sys
import unicodedata
import webbrowser
from dataclasses import dataclass
from datetime import date
from pathlib import Path
import tkinter as tk
from tkinter import font as tkfont


APP_NAME = "Dake築年数"
WINDOW_TITLE = "築年数"
DISPLAY_NAME = "築年数"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "main_title": "築年数を調べる",
    "main_description": "西暦でも和暦でも、築年数をすぐに確認できます。",
    "input_label": "築年",
    "input_placeholder": "1996 / 平成8 / H8",
    "result_prompt": "築年を入力してください",
    "result_age": "築{age}年",
    "result_detail": "{western_year}年築 / {era_year}",
    "status_idle": "築年を入力してください",
    "status_ready": "築年数を表示しています",
    "status_error": "年の形式を確認してください",
    "error_empty": "築年を入力してください",
    "error_format": "年の形式を確認してください",
    "error_future": "未来の年は築年数を計算できません",
    "error_era": "対応している元号は 令和 / 平成 / 昭和 です",
    "era_first_year": "元",
    "footer_line1": "シンプルそれDAKEシリーズ",
    "footer_link_assessment": "戸建買取査定",
    "footer_link_instagram": "Instagram",
    "footer_separator": " ｜ ",
    "footer_copyright": "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta",
}

URLS = {
    "assessment": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "instagram": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

COLORS = {
    "base_bg": "#F6F7F9",
    "panel_bg": "#FFFFFF",
    "text": "#1E2430",
    "muted": "#667085",
    "border": "#E6EAF0",
    "accent": "#2F6FED",
    "accent_soft": "#EAF2FF",
    "error": "#B42318",
    "error_soft": "#FEE4E2",
    "entry_bg": "#FFFFFF",
    "placeholder": "#98A2B3",
}

FONT_CANDIDATES = ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo")
COMMON_ICON_RELATIVE = Path("..") / ".." / "02_assets" / "dake_icon.ico"
COMMON_ICON_FILENAME = "dake_icon.ico"
WINDOW_WIDTH = 600
WINDOW_HEIGHT = 470
WINDOW_SIZE = f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}"
WINDOW_APP_ID = "Shimarisu.DakeYearAge"

ERA_ALIASES = {
    "令和": "令和",
    "R": "令和",
    "平成": "平成",
    "H": "平成",
    "昭和": "昭和",
    "S": "昭和",
}


@dataclass(frozen=True)
class EraDefinition:
    name: str
    start_year: int
    end_year: int | None


ERA_DEFINITIONS = {
    "令和": EraDefinition("令和", 2019, None),
    "平成": EraDefinition("平成", 1989, 2019),
    "昭和": EraDefinition("昭和", 1926, 1989),
}


@dataclass(frozen=True)
class YearResult:
    western_year: int
    era_year: str
    age: int


@dataclass(frozen=True)
class ParsedBuildYear:
    western_year: int
    preferred_era: str | None = None


class YearInputError(Exception):
    def __init__(self, message_key: str) -> None:
        super().__init__(message_key)
        self.message_key = message_key


def get_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def resource_icon_path() -> Path:
    base_dir = get_base_dir()
    candidates = [
        Path(getattr(sys, "_MEIPASS", base_dir)) / COMMON_ICON_FILENAME,
        (base_dir / COMMON_ICON_RELATIVE).resolve(),
        (base_dir / ".." / ".." / ".." / "02_assets" / COMMON_ICON_FILENAME).resolve(),
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[1]


def apply_window_icon(window: tk.Misc) -> None:
    try:
        icon_path = resource_icon_path()
        if icon_path.exists():
            window.iconbitmap(str(icon_path))
            window.iconbitmap(default=str(icon_path))
    except Exception:
        pass


def set_windows_app_id() -> None:
    if not sys.platform.startswith("win"):
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(WINDOW_APP_ID)
    except Exception:
        pass


def choose_font_family(root: tk.Tk) -> str:
    available = set(tkfont.families(root))
    for family in FONT_CANDIDATES:
        if family in available:
            return family
    return "TkDefaultFont"


def open_url(url: str) -> None:
    webbrowser.open(url)


def normalize_input(value: str) -> str:
    normalized = unicodedata.normalize("NFKC", value)
    normalized = normalized.upper()
    normalized = normalized.replace(" ", "").replace("\u3000", "")
    normalized = normalized.replace("年", "")
    return normalized.strip()


def parse_era_year(text: str) -> ParsedBuildYear:
    match = re.fullmatch(r"(令和|平成|昭和|R|H|S)(元|\d{1,2})", text)
    if match is None:
        if text.startswith(("明治", "大正")) or re.fullmatch(r"[MT].*", text):
            raise YearInputError("error_era")
        raise YearInputError("error_format")

    era_key = ERA_ALIASES.get(match.group(1))
    if era_key is None:
        raise YearInputError("error_era")

    era_number_text = match.group(2)
    era_number = 1 if era_number_text == UI_TEXT["era_first_year"] else int(era_number_text)
    if era_number < 1:
        raise YearInputError("error_format")

    era = ERA_DEFINITIONS[era_key]
    western_year = era.start_year + era_number - 1
    if era.end_year is not None and western_year > era.end_year:
        raise YearInputError("error_format")
    return ParsedBuildYear(western_year=western_year, preferred_era=era_key)


def parse_build_year(value: str) -> ParsedBuildYear:
    text = normalize_input(value)
    if not text:
        raise YearInputError("error_empty")

    if re.fullmatch(r"\d{4}", text):
        year = int(text)
        if year < 1926:
            raise YearInputError("error_format")
        return ParsedBuildYear(western_year=year)

    return parse_era_year(text)


def format_japanese_era(western_year: int, preferred_era: str | None = None) -> str:
    if preferred_era is not None:
        era = ERA_DEFINITIONS[preferred_era]
    elif western_year >= 2019:
        era = ERA_DEFINITIONS["令和"]
    elif western_year >= 1989:
        era = ERA_DEFINITIONS["平成"]
    elif western_year >= 1926:
        era = ERA_DEFINITIONS["昭和"]
    else:
        raise YearInputError("error_format")

    era_number = western_year - era.start_year + 1
    number_text = UI_TEXT["era_first_year"] if era_number == 1 else str(era_number)
    return f"{era.name}{number_text}年"


def calculate_year_age(value: str, current_year: int | None = None) -> YearResult:
    parsed = parse_build_year(value)
    year = parsed.western_year
    base_year = current_year if current_year is not None else date.today().year
    if year > base_year:
        raise YearInputError("error_future")
    return YearResult(
        western_year=year,
        era_year=format_japanese_era(year, parsed.preferred_era),
        age=base_year - year,
    )


class DakeYearAgeApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(WINDOW_WIDTH, WINDOW_HEIGHT)
        self.root.resizable(False, False)
        self.root.configure(bg=COLORS["base_bg"])
        apply_window_icon(self.root)

        self.font_family = choose_font_family(root)
        self.fonts = {
            "title": (self.font_family, 18, "bold"),
            "description": (self.font_family, 10),
            "label": (self.font_family, 10, "bold"),
            "entry": (self.font_family, 17),
            "result": (self.font_family, 26, "bold"),
            "detail": (self.font_family, 12),
            "status": (self.font_family, 9, "bold"),
            "footer": (self.font_family, 8),
        }

        self.input_var = tk.StringVar()
        self.result_var = tk.StringVar(value=UI_TEXT["result_prompt"])
        self.detail_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value=UI_TEXT["status_idle"])
        self.placeholder_visible = False
        self.refresh_after_id: str | None = None

        self.build_ui()
        self.show_placeholder()
        self.root.bind("<Escape>", self.clear_input)
        self.root.after(100, self.focus_entry_on_startup)

    def build_ui(self) -> None:
        outer = tk.Frame(self.root, bg=COLORS["base_bg"])
        outer.pack(fill=tk.BOTH, expand=True, padx=22, pady=(18, 10))

        tk.Label(
            outer,
            text=UI_TEXT["main_title"],
            font=self.fonts["title"],
            fg=COLORS["text"],
            bg=COLORS["base_bg"],
        ).pack(anchor=tk.CENTER)
        tk.Label(
            outer,
            text=UI_TEXT["main_description"],
            font=self.fonts["description"],
            fg=COLORS["muted"],
            bg=COLORS["base_bg"],
        ).pack(anchor=tk.CENTER, pady=(5, 14))

        panel = tk.Frame(
            outer,
            bg=COLORS["panel_bg"],
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            highlightcolor=COLORS["border"],
        )
        panel.pack(fill=tk.X)

        tk.Label(
            panel,
            text=UI_TEXT["input_label"],
            font=self.fonts["label"],
            fg=COLORS["muted"],
            bg=COLORS["panel_bg"],
        ).pack(anchor=tk.CENTER, pady=(16, 5))

        self.entry = tk.Entry(
            panel,
            textvariable=self.input_var,
            font=self.fonts["entry"],
            justify=tk.CENTER,
            fg=COLORS["text"],
            bg=COLORS["entry_bg"],
            relief=tk.SOLID,
            bd=1,
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            highlightcolor=COLORS["accent"],
            insertbackground=COLORS["text"],
        )
        self.entry.pack(fill=tk.X, padx=44, ipady=8)
        self.entry.bind("<FocusIn>", self.on_focus_in)
        self.entry.bind("<FocusOut>", self.on_focus_out)
        self.entry.bind("<Button-1>", self.on_entry_click)
        self.entry.bind("<Return>", self.on_enter)
        self.input_var.trace_add("write", self.on_input_changed)

        result_area = tk.Frame(panel, bg=COLORS["panel_bg"])
        result_area.pack(fill=tk.X, padx=20, pady=(18, 14))
        tk.Label(
            result_area,
            textvariable=self.result_var,
            font=self.fonts["result"],
            fg=COLORS["text"],
            bg=COLORS["panel_bg"],
        ).pack(anchor=tk.CENTER)
        tk.Label(
            result_area,
            textvariable=self.detail_var,
            font=self.fonts["detail"],
            fg=COLORS["muted"],
            bg=COLORS["panel_bg"],
        ).pack(anchor=tk.CENTER, pady=(5, 0))

        self.status_label = tk.Label(
            outer,
            textvariable=self.status_var,
            font=self.fonts["status"],
            fg=COLORS["muted"],
            bg=COLORS["base_bg"],
        )
        self.status_label.pack(anchor=tk.CENTER, pady=(10, 0))

        spacer = tk.Frame(outer, bg=COLORS["base_bg"])
        spacer.pack(fill=tk.BOTH, expand=True)

        self.build_footer(outer)

    def build_footer(self, parent: tk.Frame) -> None:
        footer = tk.Frame(parent, bg=COLORS["base_bg"])
        footer.pack(side=tk.BOTTOM, fill=tk.X, pady=(4, 0))

        tk.Label(
            footer,
            text=UI_TEXT["footer_line1"],
            font=self.fonts["footer"],
            fg=COLORS["muted"],
            bg=COLORS["base_bg"],
            anchor="center",
            justify="center",
        ).pack(fill=tk.X, pady=(4, 1))

        line2 = tk.Frame(footer, bg=COLORS["base_bg"])
        line2.pack(anchor=tk.CENTER, pady=(0, 6))

        self.add_footer_link(
            line2,
            UI_TEXT["footer_link_assessment"],
            URLS["assessment"],
        )
        self.add_footer_text(line2, UI_TEXT["footer_separator"])
        self.add_footer_link(
            line2,
            UI_TEXT["footer_link_instagram"],
            URLS["instagram"],
        )
        self.add_footer_text(line2, UI_TEXT["footer_separator"])
        self.add_footer_text(line2, UI_TEXT["footer_copyright"])

    def add_footer_text(self, parent: tk.Frame, text: str) -> None:
        tk.Label(
            parent,
            text=text,
            font=self.fonts["footer"],
            fg=COLORS["muted"],
            bg=COLORS["base_bg"],
            anchor="center",
            justify="center",
        ).pack(side=tk.LEFT)

    def add_footer_link(self, parent: tk.Frame, text: str, url: str) -> None:
        label = tk.Label(
            parent,
            text=text,
            font=self.fonts["footer"],
            fg=COLORS["muted"],
            bg=COLORS["base_bg"],
            cursor="hand2",
            anchor="center",
            justify="center",
        )
        label.pack(side=tk.LEFT)
        label.bind("<Button-1>", lambda _event, target_url=url: open_url(target_url))
        label.bind("<Enter>", lambda _event: label.configure(fg=COLORS["accent"]))
        label.bind("<Leave>", lambda _event: label.configure(fg=COLORS["muted"]))

    def show_placeholder(self) -> None:
        if self.input_var.get():
            return
        self.placeholder_visible = True
        self.entry.configure(fg=COLORS["placeholder"])
        self.input_var.set(UI_TEXT["input_placeholder"])

    def hide_placeholder(self) -> None:
        if not self.placeholder_visible:
            return
        self.placeholder_visible = False
        self.input_var.set("")
        self.entry.configure(fg=COLORS["text"])

    def on_focus_in(self, _event: tk.Event) -> None:
        self.hide_placeholder()
        self.select_entry_text()

    def on_entry_click(self, _event: tk.Event) -> str:
        self.hide_placeholder()
        self.select_entry_text()
        return "break"

    def select_entry_text(self) -> None:
        self.root.after_idle(lambda: self.entry.select_range(0, tk.END))

    def focus_entry_on_startup(self) -> None:
        self.entry.focus_set()
        self.select_entry_text()

    def on_focus_out(self, _event: tk.Event) -> None:
        if not self.input_var.get().strip():
            self.show_placeholder()

    def on_enter(self, _event: tk.Event) -> str:
        self.refresh_result()
        return "break"

    def clear_input(self, _event: tk.Event | None = None) -> str:
        self.placeholder_visible = False
        self.entry.configure(fg=COLORS["text"])
        self.entry.delete(0, tk.END)
        if self.refresh_after_id is not None:
            self.root.after_cancel(self.refresh_after_id)
            self.refresh_after_id = None
        self.result_var.set(UI_TEXT["result_prompt"])
        self.detail_var.set("")
        self.status_var.set(UI_TEXT["status_idle"])
        self.status_label.configure(fg=COLORS["muted"], bg=COLORS["base_bg"])
        self.entry.focus_set()
        return "break"

    def on_input_changed(self, *_args: object) -> None:
        if self.placeholder_visible:
            return
        if self.refresh_after_id is not None:
            self.root.after_cancel(self.refresh_after_id)
        self.refresh_after_id = self.root.after(60, self.refresh_result)

    def refresh_result(self) -> None:
        self.refresh_after_id = None
        value = "" if self.placeholder_visible else self.input_var.get()
        try:
            result = calculate_year_age(value)
        except YearInputError as exc:
            self.show_message(exc.message_key, is_error=exc.message_key not in {"error_empty"})
            return

        self.result_var.set(UI_TEXT["result_age"].format(age=result.age))
        self.detail_var.set(
            UI_TEXT["result_detail"].format(
                western_year=result.western_year,
                era_year=result.era_year,
            )
        )
        self.status_var.set(UI_TEXT["status_ready"])
        self.status_label.configure(fg=COLORS["accent"], bg=COLORS["accent_soft"])

    def show_message(self, key: str, *, is_error: bool) -> None:
        message = UI_TEXT[key]
        self.result_var.set(message)
        self.detail_var.set("")
        if is_error:
            self.status_var.set(UI_TEXT["status_error"])
            self.status_label.configure(fg=COLORS["error"], bg=COLORS["error_soft"])
        else:
            self.status_var.set(UI_TEXT["status_idle"])
            self.status_label.configure(fg=COLORS["muted"], bg=COLORS["base_bg"])


def main() -> None:
    set_windows_app_id()
    root = tk.Tk()
    DakeYearAgeApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
