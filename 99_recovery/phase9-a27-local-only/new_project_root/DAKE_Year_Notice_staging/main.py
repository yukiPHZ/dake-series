# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
import webbrowser
from dataclasses import dataclass
from datetime import date
from pathlib import Path
import tkinter as tk
from tkinter import font as tkfont


APP_NAME = "Dake今年の注意点"
WINDOW_TITLE = "今年の注意点"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "main_title": "今年の注意点",
    "main_description": "閏年と固定資産税の評価替え年を確認します。",
    "label_leap_year": "閏年",
    "label_fixed_tax_revaluation": "固定資産税評価替え",
    "label_next_revaluation": "次回評価替え",
    "status_checked": "確認済み",
    "yes": "はい",
    "no": "いいえ",
    "year_format": "{year}年",
    "year_with_era_format": "{year}年（{era}）",
    "reiwa_first_year": "令和元年",
    "reiwa_year_format": "令和{year}年",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_link_1": "戸建買取査定",
    "footer_link_2": "Instagram",
    "footer_separator": " ｜ ",
    "footer_copyright": COPYRIGHT,
}

LINK_URLS = {
    "footer_link_1": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "footer_link_2": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

FONT_FALLBACKS = ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo", "TkDefaultFont")
REIWA_START_YEAR = 2019
FIXED_TAX_REVALUATION_BASE_YEAR = 2024

COLORS = {
    "background": "#F6F7F9",
    "card": "#FFFFFF",
    "text": "#1E2430",
    "muted": "#667085",
    "border": "#E6EAF0",
    "accent": "#2F6FED",
    "success": "#15803D",
}


@dataclass(frozen=True)
class YearNotice:
    year: int
    era: str
    is_leap: bool
    is_revaluation: bool
    next_revaluation_year: int
    next_revaluation_era: str


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


def is_leap_year(year: int) -> bool:
    return year % 400 == 0 or (year % 4 == 0 and year % 100 != 0)


def is_fixed_tax_revaluation_year(year: int) -> bool:
    return (year - FIXED_TAX_REVALUATION_BASE_YEAR) % 3 == 0


def next_fixed_tax_revaluation_year(year: int) -> int:
    candidate = year + 1
    while not is_fixed_tax_revaluation_year(candidate):
        candidate += 1
    return candidate


def format_reiwa_year(year: int) -> str:
    reiwa_year = year - REIWA_START_YEAR + 1
    if reiwa_year == 1:
        return UI_TEXT["reiwa_first_year"]
    return UI_TEXT["reiwa_year_format"].format(year=reiwa_year)


def format_year_with_era(year: int) -> str:
    return UI_TEXT["year_with_era_format"].format(year=year, era=format_reiwa_year(year))


def yes_no(value: bool) -> str:
    return UI_TEXT["yes"] if value else UI_TEXT["no"]


def build_year_notice(year: int) -> YearNotice:
    next_year = next_fixed_tax_revaluation_year(year)
    return YearNotice(
        year=year,
        era=format_reiwa_year(year),
        is_leap=is_leap_year(year),
        is_revaluation=is_fixed_tax_revaluation_year(year),
        next_revaluation_year=next_year,
        next_revaluation_era=format_reiwa_year(next_year),
    )


class YearNoticeApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.notice = build_year_notice(date.today().year)
        self.font_family = choose_font_family(self)
        self.footer_link_font = tkfont.Font(
            family=self.font_family,
            size=8,
            underline=True,
        )

        self.title(WINDOW_TITLE)
        self.geometry("420x430")
        self.minsize(420, 430)
        self.maxsize(420, 430)
        self.resizable(False, False)
        self.configure(bg=COLORS["background"])
        self._set_window_icon()
        self._build_ui()

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
        shell.pack(fill="both", expand=True, padx=20, pady=18)

        self._build_header(shell)
        self._build_year_panel(shell)
        self._build_result_card(shell)
        self._build_status(shell)
        self._build_footer(shell)

    def _build_header(self, parent: tk.Widget) -> None:
        header = tk.Frame(parent, bg=COLORS["background"])
        header.pack(fill="x")

        tk.Label(
            header,
            text=UI_TEXT["main_title"],
            bg=COLORS["background"],
            fg=COLORS["text"],
            font=(self.font_family, 18, "bold"),
            anchor="w",
        ).pack(fill="x")

        tk.Label(
            header,
            text=UI_TEXT["main_description"],
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=(self.font_family, 10),
            anchor="w",
        ).pack(fill="x", pady=(6, 0))

    def _build_year_panel(self, parent: tk.Widget) -> None:
        panel = tk.Frame(parent, bg=COLORS["background"])
        panel.pack(fill="x", pady=(30, 20))

        tk.Label(
            panel,
            text=format_year_with_era(self.notice.year),
            bg=COLORS["background"],
            fg=COLORS["text"],
            font=(self.font_family, 22, "bold"),
            anchor="center",
        ).pack(fill="x")

    def _build_result_card(self, parent: tk.Widget) -> None:
        card = tk.Frame(
            parent,
            bg=COLORS["card"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
            bd=0,
        )
        card.pack(fill="x", pady=(0, 18))
        card.columnconfigure(1, weight=1)

        rows = [
            (UI_TEXT["label_leap_year"], yes_no(self.notice.is_leap)),
            (UI_TEXT["label_fixed_tax_revaluation"], yes_no(self.notice.is_revaluation)),
            (
                UI_TEXT["label_next_revaluation"],
                format_year_with_era(self.notice.next_revaluation_year),
            ),
        ]

        for row_index, (label_text, value_text) in enumerate(rows):
            tk.Label(
                card,
                text=label_text,
                bg=COLORS["card"],
                fg=COLORS["muted"],
                font=(self.font_family, 10),
                anchor="w",
            ).grid(row=row_index, column=0, sticky="w", padx=(18, 8), pady=(16, 0))

            tk.Label(
                card,
                text=value_text,
                bg=COLORS["card"],
                fg=COLORS["text"],
                font=(self.font_family, 13, "bold"),
                anchor="e",
            ).grid(row=row_index, column=1, sticky="ew", padx=(8, 18), pady=(16, 0))

        tk.Frame(card, bg=COLORS["card"], height=16).grid(row=len(rows), column=0, columnspan=2)

    def _build_status(self, parent: tk.Widget) -> None:
        tk.Label(
            parent,
            text=UI_TEXT["status_checked"],
            bg=COLORS["background"],
            fg=COLORS["success"],
            font=(self.font_family, 10, "bold"),
            anchor="center",
        ).pack(fill="x", pady=(0, 16))

    def _build_footer(self, parent: tk.Widget) -> None:
        footer = tk.Frame(parent, bg=COLORS["background"])
        footer.pack(side="bottom", fill="x")

        footer_row = tk.Frame(footer, bg=COLORS["background"])
        footer_row.pack(anchor="center")

        self._add_footer_label(footer_row, UI_TEXT["footer_left"])
        self._add_footer_label(footer_row, UI_TEXT["footer_separator"])
        self._add_footer_link(footer_row, "footer_link_1")
        self._add_footer_label(footer_row, UI_TEXT["footer_separator"])
        self._add_footer_link(footer_row, "footer_link_2")

        tk.Label(
            footer,
            text=UI_TEXT["footer_copyright"],
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=(self.font_family, 8),
            anchor="center",
        ).pack(fill="x", pady=(4, 0))

    def _add_footer_label(self, parent: tk.Widget, text: str) -> None:
        tk.Label(
            parent,
            text=text,
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=(self.font_family, 8),
        ).pack(side="left")

    def _add_footer_link(self, parent: tk.Widget, key: str) -> None:
        label = tk.Label(
            parent,
            text=UI_TEXT[key],
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=(self.font_family, 8),
            cursor="hand2",
        )
        label.pack(side="left")
        label.bind("<Button-1>", lambda _event, url=LINK_URLS[key]: webbrowser.open(url))
        label.bind("<Enter>", lambda _event: label.configure(fg=COLORS["accent"], font=self.footer_link_font))
        label.bind("<Leave>", lambda _event: label.configure(fg=COLORS["muted"], font=(self.font_family, 8)))


def main() -> None:
    app = YearNoticeApp()
    app.mainloop()


if __name__ == "__main__":
    main()
