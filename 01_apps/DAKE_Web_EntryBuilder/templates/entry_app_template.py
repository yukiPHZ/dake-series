# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import shutil
import subprocess
import sys
import webbrowser
from pathlib import Path
import tkinter as tk
from tkinter import font as tkfont
from tkinter import messagebox


APP_NAME = __DISPLAY_NAME_JSON__
WINDOW_TITLE = __DISPLAY_NAME_JSON__
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "main_heading": __SCREEN_TITLE_JSON__,
    "description": "このサイトはブラウザにより動作が異なる場合があります",
    "button_edge": "Edgeで開く",
    "button_default": "既定ブラウザで開く",
    "button_internal": "このアプリで開く（簡易表示）",
    "unofficial_text": "本アプリは非公式のアクセス補助ツールです。\n表示・動作は対象Webサイトおよびブラウザ環境に依存します。",
    "disclaimer_heading": "免責",
    "disclaimer_text": "本アプリはURLを開くだけです\n対象Webサイトの表示や動作を保証しません\nブラウザや利用環境により表示が異なる場合があります\n操作結果は利用者の責任で確認してください",
    "dialog_error_title": "開けませんでした",
    "error_edge_not_found": "Edgeを起動できませんでした。既定ブラウザをお試しください。",
    "error_default_open": "既定ブラウザを起動できませんでした。",
    "error_internal_open": "簡易表示を起動できませんでした。",
    "error_webview_missing": "pywebviewを読み込めませんでした。requirements.txt を確認してください。",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_estimate": "戸建買取査定",
    "footer_instagram": "Instagram",
    "footer_separator": " ｜ ",
    "footer_copyright": COPYRIGHT,
}

COLORS = {
    "base_bg": "#F6F7F9",
    "card_bg": "#FFFFFF",
    "text": "#1E2430",
    "sub_text": "#667085",
    "border": "#E6EAF0",
    "accent": "#2F6FED",
    "accent_hover": "#2458BF",
    "notice_bg": "#F8FAFC",
}

LINK_URLS = {
    "estimate": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "instagram": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

TARGET_URL = __TARGET_URL_JSON__
RECOMMENDED_MODE = __RECOMMENDED_MODE_JSON__
EXE_NAME = __EXE_NAME_JSON__
ICON_RELATIVE_PATH = Path("..") / ".." / ".." / ".." / "02_assets" / "dake_icon.ico"


def choose_font_family(root: tk.Tk) -> str:
    available = set(tkfont.families(root))
    for candidate in ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo"):
        if candidate in available:
            return candidate
    return "TkDefaultFont"


def show_error_dialog(message: str) -> None:
    try:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(UI_TEXT["dialog_error_title"], message)
        root.destroy()
    except tk.TclError:
        return


def find_edge_path() -> str | None:
    found = shutil.which("msedge")
    if found:
        return found

    candidates = [
        Path(os.environ.get("PROGRAMFILES(X86)", "")) / "Microsoft" / "Edge" / "Application" / "msedge.exe",
        Path(os.environ.get("PROGRAMFILES", "")) / "Microsoft" / "Edge" / "Application" / "msedge.exe",
        Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "Edge" / "Application" / "msedge.exe",
    ]
    for candidate in candidates:
        if candidate.exists():
            return str(candidate)
    return None


def run_webview() -> None:
    try:
        import webview
    except ImportError:
        show_error_dialog(UI_TEXT["error_webview_missing"])
        return

    try:
        webview.create_window(APP_NAME, TARGET_URL, width=1180, height=760)
        webview.start()
    except Exception:
        show_error_dialog(UI_TEXT["error_internal_open"])


class WebEntryApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.configure(bg=COLORS["base_bg"])
        self.root.minsize(720, 620)
        self.font_family = choose_font_family(self.root)
        self.fonts = {
            "title": (self.font_family, 22, "bold"),
            "subtitle": (self.font_family, 11),
            "button_main": (self.font_family, 13, "bold"),
            "button_sub": (self.font_family, 11, "bold"),
            "body": (self.font_family, 10),
            "body_bold": (self.font_family, 10, "bold"),
            "small": (self.font_family, 9),
        }
        self._set_icon()
        self._build_ui()

    def _set_icon(self) -> None:
        icon_path = (Path(__file__).resolve().parent / ICON_RELATIVE_PATH).resolve()
        if not icon_path.exists():
            return
        try:
            self.root.iconbitmap(default=str(icon_path))
        except tk.TclError:
            return

    def _build_ui(self) -> None:
        outer = tk.Frame(self.root, bg=COLORS["base_bg"])
        outer.pack(fill="both", expand=True, padx=28, pady=26)

        header = tk.Frame(outer, bg=COLORS["base_bg"])
        header.pack(fill="x")
        tk.Label(
            header,
            text=UI_TEXT["main_heading"],
            font=self.fonts["title"],
            fg=COLORS["text"],
            bg=COLORS["base_bg"],
            anchor="w",
        ).pack(anchor="w")
        tk.Label(
            header,
            text=UI_TEXT["description"],
            font=self.fonts["subtitle"],
            fg=COLORS["sub_text"],
            bg=COLORS["base_bg"],
            anchor="w",
        ).pack(anchor="w", pady=(8, 0))

        action_card = self._create_card(outer, pady=(28, 18))
        self._build_action_buttons(action_card)

        notice_card = self._create_card(outer, pady=(0, 14))
        tk.Label(
            notice_card,
            text=UI_TEXT["unofficial_text"],
            font=self.fonts["body"],
            fg=COLORS["sub_text"],
            bg=COLORS["card_bg"],
            justify="left",
            anchor="w",
        ).pack(fill="x")

        disclaimer_card = self._create_card(outer)
        tk.Label(
            disclaimer_card,
            text=UI_TEXT["disclaimer_heading"],
            font=self.fonts["body_bold"],
            fg=COLORS["text"],
            bg=COLORS["card_bg"],
            anchor="w",
        ).pack(anchor="w")
        tk.Label(
            disclaimer_card,
            text=UI_TEXT["disclaimer_text"],
            font=self.fonts["body"],
            fg=COLORS["sub_text"],
            bg=COLORS["card_bg"],
            justify="left",
            anchor="w",
        ).pack(fill="x", pady=(8, 0))

        spacer = tk.Frame(outer, bg=COLORS["base_bg"])
        spacer.pack(fill="both", expand=True)
        self._build_footer(outer)

    def _create_card(self, parent: tk.Widget, pady: tuple[int, int] | int = 0) -> tk.Frame:
        card = tk.Frame(parent, bg=COLORS["card_bg"], highlightthickness=1, highlightbackground=COLORS["border"])
        card.pack(fill="x", pady=pady)
        inner = tk.Frame(card, bg=COLORS["card_bg"])
        inner.pack(fill="both", expand=True, padx=22, pady=20)
        return inner

    def _build_action_buttons(self, parent: tk.Widget) -> None:
        actions = self._ordered_actions()
        for index, (mode, label, command) in enumerate(actions):
            is_primary = index == 0
            button = tk.Button(
                parent,
                text=label,
                command=command,
                font=self.fonts["button_main"] if is_primary else self.fonts["button_sub"],
                fg="#FFFFFF" if is_primary else COLORS["text"],
                bg=COLORS["accent"] if is_primary else COLORS["card_bg"],
                activeforeground="#FFFFFF" if is_primary else COLORS["text"],
                activebackground=COLORS["accent_hover"] if is_primary else COLORS["notice_bg"],
                relief="flat",
                bd=0,
                padx=18,
                pady=16 if is_primary else 12,
                cursor="hand2",
                highlightthickness=1,
                highlightbackground=COLORS["accent"] if is_primary else COLORS["border"],
            )
            button.pack(fill="x", pady=(0, 10 if index < len(actions) - 1 else 0))
            self._bind_button_hover(button, is_primary)

    def _ordered_actions(self) -> list[tuple[str, str, object]]:
        action_map = {
            "edge": ("edge", UI_TEXT["button_edge"], self.open_edge),
            "default": ("default", UI_TEXT["button_default"], self.open_default),
            "internal": ("internal", UI_TEXT["button_internal"], self.open_internal),
        }
        order_by_mode = {
            "edge": ["edge", "default", "internal"],
            "default": ["default", "edge", "internal"],
            "internal": ["internal", "edge", "default"],
        }
        return [action_map[mode] for mode in order_by_mode.get(RECOMMENDED_MODE, order_by_mode["edge"])]

    def _bind_button_hover(self, button: tk.Button, is_primary: bool) -> None:
        normal_bg = COLORS["accent"] if is_primary else COLORS["card_bg"]
        hover_bg = COLORS["accent_hover"] if is_primary else COLORS["notice_bg"]
        button.bind("<Enter>", lambda _event: button.configure(bg=hover_bg))
        button.bind("<Leave>", lambda _event: button.configure(bg=normal_bg))

    def _build_footer(self, parent: tk.Widget) -> None:
        footer = tk.Frame(parent, bg=COLORS["base_bg"])
        footer.pack(fill="x", pady=(18, 0))
        tk.Label(
            footer,
            text=UI_TEXT["footer_left"],
            font=self.fonts["small"],
            fg=COLORS["sub_text"],
            bg=COLORS["base_bg"],
        ).pack(side="left")

        right = tk.Frame(footer, bg=COLORS["base_bg"])
        right.pack(side="right")
        self._footer_link(right, UI_TEXT["footer_estimate"], LINK_URLS["estimate"])
        self._footer_label(right, UI_TEXT["footer_separator"])
        self._footer_link(right, UI_TEXT["footer_instagram"], LINK_URLS["instagram"])
        self._footer_label(right, UI_TEXT["footer_separator"])
        self._footer_label(right, UI_TEXT["footer_copyright"])

    def _footer_label(self, parent: tk.Widget, text: str) -> None:
        tk.Label(parent, text=text, font=self.fonts["small"], fg=COLORS["sub_text"], bg=COLORS["base_bg"]).pack(side="left")

    def _footer_link(self, parent: tk.Widget, text: str, url: str) -> None:
        label = tk.Label(parent, text=text, font=self.fonts["small"], fg=COLORS["accent"], bg=COLORS["base_bg"], cursor="hand2")
        label.pack(side="left")
        label.bind("<Button-1>", lambda _event: webbrowser.open(url))

    def open_edge(self) -> None:
        edge_path = find_edge_path()
        if edge_path is None:
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["error_edge_not_found"])
            return
        try:
            subprocess.Popen([edge_path, TARGET_URL])
        except OSError:
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["error_edge_not_found"])

    def open_default(self) -> None:
        try:
            opened = webbrowser.open(TARGET_URL)
        except webbrowser.Error:
            opened = False
        if not opened:
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["error_default_open"])

    def open_internal(self) -> None:
        try:
            if getattr(sys, "frozen", False):
                subprocess.Popen([sys.executable, "--webview"])
            else:
                subprocess.Popen([sys.executable, str(Path(__file__).resolve()), "--webview"])
        except OSError:
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["error_internal_open"])


def main() -> None:
    if "--webview" in sys.argv:
        run_webview()
        return

    root = tk.Tk()
    WebEntryApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
