# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import os
import re
import shutil
import webbrowser
from pathlib import Path
from urllib.parse import urlparse
import tkinter as tk
from tkinter import font as tkfont
from tkinter import messagebox, ttk


APP_NAME = "Dake専用Web入口Builder"
WINDOW_TITLE = "Dake専用Web入口Builder"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "main_heading": "専用Web入口を作る",
    "main_description": "指定したURLだけを開く、迷わない入口を作成します",
    "label_display_name": "表示名（日本語）",
    "label_screen_title": "画面見出し",
    "label_url": "起動URL（必須）",
    "label_exe_name": "exe名（英数字）",
    "label_browser": "推奨ブラウザ",
    "example_display_name": "例：Dake千葉市ハザード入口",
    "example_screen_title": "例：千葉市ハザードマップを開く",
    "example_url": "例：https://www.city.chiba.jp/",
    "example_exe_name": "例：DakeChibaHazard.exe",
    "option_edge": "Edgeで開く",
    "option_default": "既定ブラウザ",
    "option_internal": "このアプリで開く（簡易表示）",
    "button_create": "専用アプリを作成",
    "button_open_output": "出力フォルダを開く",
    "notice_title": "注意",
    "notice_text": "本アプリは非公式のアクセス補助ツールを作成します。\n各Webサイトの運営元とは関係ありません。",
    "http_notice": "httpのURLです。サイトにより表示や動作が異なる場合があります。",
    "status_idle": "入力して作成できます",
    "status_created": "作成しました",
    "dialog_error_title": "確認してください",
    "dialog_complete_title": "作成しました",
    "dialog_complete_message": "専用アプリを作成しました。フォルダを開きます。",
    "dialog_http_title": "確認",
    "dialog_http_message": "httpsではないURLです。このまま作成しますか？",
    "error_display_name_required": "表示名を入力してください",
    "error_screen_title_required": "画面見出しを入力してください",
    "error_url_required": "URLを入力してください",
    "error_url_invalid": "URLを確認してください",
    "error_exe_required": "exe名を入力してください",
    "error_exe_invalid": "exe名は英数字で入力してください",
    "error_template_missing": "テンプレートを読み込めませんでした",
    "error_output_failed": "作成できませんでした。入力内容を確認してください",
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
    "warning": "#B54708",
}

LINK_URLS = {
    "estimate": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "instagram": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

BROWSER_OPTIONS = {
    UI_TEXT["option_edge"]: "edge",
    UI_TEXT["option_default"]: "default",
    UI_TEXT["option_internal"]: "internal",
}

EXE_NAME_PATTERN = re.compile(r"^[A-Za-z0-9_]+(?:\.exe)?$", re.IGNORECASE)
BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "templates" / "entry_app_template.py"
OUTPUT_DIR = BASE_DIR / "output"
ICON_RELATIVE_PATH = Path("..") / ".." / "02_assets" / "dake_icon.ico"


ENTRY_APP_TEMPLATE_FALLBACK = r'''# -*- coding: utf-8 -*-
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
'''


class WebEntryBuilderApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.configure(bg=COLORS["base_bg"])
        self.root.minsize(760, 720)
        self.font_family = self._choose_font_family()
        self.fonts = {
            "title": (self.font_family, 22, "bold"),
            "subtitle": (self.font_family, 11),
            "label": (self.font_family, 10, "bold"),
            "body": (self.font_family, 10),
            "small": (self.font_family, 9),
            "button": (self.font_family, 12, "bold"),
        }

        self.display_name_var = tk.StringVar()
        self.screen_title_var = tk.StringVar()
        self.url_var = tk.StringVar()
        self.exe_name_var = tk.StringVar()
        self.browser_var = tk.StringVar(value=UI_TEXT["option_edge"])
        self.status_var = tk.StringVar(value=UI_TEXT["status_idle"])
        self.http_notice_var = tk.StringVar(value="")

        self._set_icon()
        self._build_ui()
        self.url_var.trace_add("write", self._update_url_notice)

    def _choose_font_family(self) -> str:
        available = set(tkfont.families(self.root))
        for candidate in ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo"):
            if candidate in available:
                return candidate
        return "TkDefaultFont"

    def _set_icon(self) -> None:
        icon_path = (BASE_DIR / ICON_RELATIVE_PATH).resolve()
        if not icon_path.exists():
            return
        try:
            self.root.iconbitmap(default=str(icon_path))
        except tk.TclError:
            return

    def _build_ui(self) -> None:
        outer = tk.Frame(self.root, bg=COLORS["base_bg"])
        outer.pack(fill="both", expand=True, padx=28, pady=24)

        self._build_header(outer)
        form_card = self._create_card(outer, pady=(18, 14))
        self._build_form(form_card)

        notice_card = self._create_card(outer, pady=(0, 14))
        self._build_notice(notice_card)

        action_row = tk.Frame(outer, bg=COLORS["base_bg"])
        action_row.pack(fill="x", pady=(0, 10))
        create_button = tk.Button(
            action_row,
            text=UI_TEXT["button_create"],
            command=self.create_app,
            font=self.fonts["button"],
            fg="#FFFFFF",
            bg=COLORS["accent"],
            activeforeground="#FFFFFF",
            activebackground=COLORS["accent_hover"],
            relief="flat",
            bd=0,
            padx=18,
            pady=14,
            cursor="hand2",
        )
        create_button.pack(fill="x")
        self._bind_button_hover(create_button)

        tk.Label(
            outer,
            textvariable=self.status_var,
            font=self.fonts["small"],
            fg=COLORS["sub_text"],
            bg=COLORS["base_bg"],
            anchor="w",
        ).pack(fill="x", pady=(2, 0))

        spacer = tk.Frame(outer, bg=COLORS["base_bg"])
        spacer.pack(fill="both", expand=True)
        self._build_footer(outer)

    def _build_header(self, parent: tk.Widget) -> None:
        header = tk.Frame(parent, bg=COLORS["base_bg"])
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
            text=UI_TEXT["main_description"],
            font=self.fonts["subtitle"],
            fg=COLORS["sub_text"],
            bg=COLORS["base_bg"],
            anchor="w",
        ).pack(anchor="w", pady=(8, 0))

    def _create_card(self, parent: tk.Widget, pady: tuple[int, int] | int = 0) -> tk.Frame:
        card = tk.Frame(parent, bg=COLORS["card_bg"], highlightbackground=COLORS["border"], highlightthickness=1)
        card.pack(fill="x", pady=pady)
        inner = tk.Frame(card, bg=COLORS["card_bg"])
        inner.pack(fill="both", expand=True, padx=22, pady=20)
        return inner

    def _build_form(self, parent: tk.Widget) -> None:
        self._field(parent, UI_TEXT["label_display_name"], self.display_name_var, UI_TEXT["example_display_name"])
        self._field(parent, UI_TEXT["label_screen_title"], self.screen_title_var, UI_TEXT["example_screen_title"])
        self._field(parent, UI_TEXT["label_url"], self.url_var, UI_TEXT["example_url"])
        tk.Label(
            parent,
            textvariable=self.http_notice_var,
            font=self.fonts["small"],
            fg=COLORS["warning"],
            bg=COLORS["card_bg"],
            anchor="w",
        ).pack(fill="x", pady=(0, 10))
        self._field(parent, UI_TEXT["label_exe_name"], self.exe_name_var, UI_TEXT["example_exe_name"])

        tk.Label(
            parent,
            text=UI_TEXT["label_browser"],
            font=self.fonts["label"],
            fg=COLORS["text"],
            bg=COLORS["card_bg"],
            anchor="w",
        ).pack(fill="x", pady=(2, 6))
        browser_box = ttk.Combobox(
            parent,
            textvariable=self.browser_var,
            values=list(BROWSER_OPTIONS.keys()),
            state="readonly",
            font=self.fonts["body"],
        )
        browser_box.pack(fill="x", ipady=6)

    def _field(self, parent: tk.Widget, label_text: str, variable: tk.StringVar, example_text: str) -> None:
        tk.Label(
            parent,
            text=label_text,
            font=self.fonts["label"],
            fg=COLORS["text"],
            bg=COLORS["card_bg"],
            anchor="w",
        ).pack(fill="x", pady=(0, 6))
        tk.Entry(
            parent,
            textvariable=variable,
            font=self.fonts["body"],
            fg=COLORS["text"],
            bg="#FFFFFF",
            relief="solid",
            bd=1,
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            highlightcolor=COLORS["accent"],
            insertbackground=COLORS["text"],
        ).pack(fill="x", ipady=7)
        tk.Label(
            parent,
            text=example_text,
            font=self.fonts["small"],
            fg=COLORS["sub_text"],
            bg=COLORS["card_bg"],
            anchor="w",
        ).pack(fill="x", pady=(5, 12))

    def _build_notice(self, parent: tk.Widget) -> None:
        tk.Label(
            parent,
            text=UI_TEXT["notice_title"],
            font=self.fonts["label"],
            fg=COLORS["text"],
            bg=COLORS["card_bg"],
            anchor="w",
        ).pack(anchor="w")
        tk.Label(
            parent,
            text=UI_TEXT["notice_text"],
            font=self.fonts["body"],
            fg=COLORS["sub_text"],
            bg=COLORS["card_bg"],
            justify="left",
            anchor="w",
        ).pack(fill="x", pady=(8, 0))

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

    def _bind_button_hover(self, button: tk.Button) -> None:
        button.bind("<Enter>", lambda _event: button.configure(bg=COLORS["accent_hover"]))
        button.bind("<Leave>", lambda _event: button.configure(bg=COLORS["accent"]))

    def _update_url_notice(self, *_args: object) -> None:
        url = self.url_var.get().strip().lower()
        self.http_notice_var.set(UI_TEXT["http_notice"] if url.startswith("http://") else "")

    def create_app(self) -> None:
        try:
            config = self._validate_inputs()
        except ValueError as error:
            messagebox.showerror(UI_TEXT["dialog_error_title"], str(error))
            return

        if config["url"].startswith("http://"):
            proceed = messagebox.askyesno(UI_TEXT["dialog_http_title"], UI_TEXT["dialog_http_message"])
            if not proceed:
                return

        try:
            output_folder = self._write_generated_app(config)
        except OSError:
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["error_output_failed"])
            return

        self.status_var.set(UI_TEXT["status_created"])
        messagebox.showinfo(UI_TEXT["dialog_complete_title"], UI_TEXT["dialog_complete_message"])
        self._open_folder(output_folder)

    def _validate_inputs(self) -> dict[str, str]:
        display_name = self.display_name_var.get().strip()
        screen_title = self.screen_title_var.get().strip()
        url = self.url_var.get().strip()
        exe_name = self.exe_name_var.get().strip()
        browser_label = self.browser_var.get().strip()

        if not display_name:
            raise ValueError(UI_TEXT["error_display_name_required"])
        if not screen_title:
            raise ValueError(UI_TEXT["error_screen_title_required"])
        if not url:
            raise ValueError(UI_TEXT["error_url_required"])

        parsed = urlparse(url)
        if parsed.scheme not in {"http", "https"} or not parsed.netloc:
            raise ValueError(UI_TEXT["error_url_invalid"])

        if not exe_name:
            raise ValueError(UI_TEXT["error_exe_required"])
        if not EXE_NAME_PATTERN.fullmatch(exe_name):
            raise ValueError(UI_TEXT["error_exe_invalid"])

        if not exe_name.lower().endswith(".exe"):
            exe_name = f"{exe_name}.exe"

        return {
            "display_name": display_name,
            "screen_title": screen_title,
            "url": url,
            "exe_name": exe_name,
            "exe_stem": Path(exe_name).stem,
            "browser_mode": BROWSER_OPTIONS.get(browser_label, "edge"),
        }

    def _write_generated_app(self, config: dict[str, str]) -> Path:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        folder = self._unique_output_folder(config["exe_stem"])
        folder.mkdir(parents=True, exist_ok=False)

        template = self._load_template()
        replacements = {
            "__DISPLAY_NAME_JSON__": json.dumps(config["display_name"], ensure_ascii=False),
            "__SCREEN_TITLE_JSON__": json.dumps(config["screen_title"], ensure_ascii=False),
            "__TARGET_URL_JSON__": json.dumps(config["url"], ensure_ascii=False),
            "__RECOMMENDED_MODE_JSON__": json.dumps(config["browser_mode"], ensure_ascii=False),
            "__EXE_NAME_JSON__": json.dumps(config["exe_name"], ensure_ascii=False),
        }
        for marker, value in replacements.items():
            template = template.replace(marker, value)

        self._write_text(folder / "main.py", template)
        self._write_text(folder / "build.bat", self._generated_build_bat(config["exe_stem"]))
        self._write_text(folder / "requirements.txt", "pyinstaller\npywebview\n")
        self._write_text(folder / "README.md", self._generated_readme(config))
        return folder

    def _load_template(self) -> str:
        if TEMPLATE_PATH.exists():
            return TEMPLATE_PATH.read_text(encoding="utf-8")
        if ENTRY_APP_TEMPLATE_FALLBACK:
            return ENTRY_APP_TEMPLATE_FALLBACK
        raise OSError(UI_TEXT["error_template_missing"])

    def _unique_output_folder(self, base_name: str) -> Path:
        candidate = OUTPUT_DIR / base_name
        if not candidate.exists():
            return candidate
        index = 2
        while True:
            numbered = OUTPUT_DIR / f"{base_name}_{index}"
            if not numbered.exists():
                return numbered
            index += 1

    def _generated_build_bat(self, exe_stem: str) -> str:
        return f"""@echo off
chcp 65001 > nul

rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
del *.spec 2>nul

pyinstaller ^
--onefile ^
--noconsole ^
--clean ^
--noconfirm ^
--name {exe_stem} ^
--icon=..\\..\\..\\..\\02_assets\\dake_icon.ico ^
main.py

pause
"""

    def _generated_readme(self, config: dict[str, str]) -> str:
        return f"""# {config["display_name"]}

これは指定URL専用の入口アプリです。
ブラウザではありません。
公式ツールではありません。
対象Webサイトの表示や動作を保証しません。
URLを開く以外の処理は行いません。

## 開くURL

{config["url"]}

## 使い方

1. `main.py` を実行します。
2. 表示されたボタンから、推奨ブラウザまたは既定ブラウザでURLを開きます。
3. 簡易表示が必要な場合のみ、このアプリで開くボタンを使います。

## ビルド

Windows環境で `build.bat` を実行すると、`dist/{config["exe_name"]}` が作成されます。
"""

    def _write_text(self, path: Path, content: str) -> None:
        path.write_text(content, encoding="utf-8", newline="\n")

    def _open_folder(self, folder: Path) -> None:
        try:
            if os.name == "nt":
                os.startfile(str(folder))
            else:
                webbrowser.open(folder.as_uri())
        except OSError:
            return


def main() -> None:
    root = tk.Tk()
    WebEntryBuilderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
