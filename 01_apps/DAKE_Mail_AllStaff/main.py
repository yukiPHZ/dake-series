# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import sys
import webbrowser
from pathlib import Path
from urllib.parse import quote
import tkinter as tk
from tkinter import font as tkfont


APP_NAME = "Dake全社員メール起動"
WINDOW_TITLE = "Dake全社員メール起動"
DISPLAY_NAME = "全社員メール"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "main_title": "全社員宛メールを開く",
    "main_description": "宛先とCCを入力した状態で、既定のメーラーを起動します。送信はされません。",
    "button_open_mail": "メールを開く",
    "status_display": "状態：{status}",
    "status_ready": "準備完了",
    "status_opened": "メーラーを起動しました",
    "status_config_created": "設定ファイルを作成しました。宛先を確認してください。",
    "status_config_invalid": "設定ファイルを読み込めませんでした。内容を確認してください。",
    "status_config_write_failed": "設定ファイルを作成できませんでした。保存場所を確認してください。",
    "status_open_failed": "メーラーを起動できませんでした。既定のメールアプリを確認してください。",
    "error_no_to": "宛先が未設定です。設定ファイルを確認してください。",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_link_1": "戸建買取査定",
    "footer_link_2": "Instagram",
    "footer_separator": " ｜ ",
    "footer_copyright": COPYRIGHT,
}

COLORS = {
    "background": "#F6F7F9",
    "card": "#FFFFFF",
    "text": "#1E2430",
    "muted": "#667085",
    "border": "#E6EAF0",
    "accent": "#2F6FED",
    "accent_hover": "#2458BF",
    "selection_bg": "#EAF2FF",
    "error": "#D92D20",
}

LINK_URLS = {
    "footer_link_1": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "footer_link_2": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

CONFIG_FILE_NAME = "Dake_AllStaff_Mail_config.json"
DEFAULT_CONFIG = {
    "to": "all@example.co.jp",
    "cc": "example@example.co.jp",
}
ICON_RELATIVE_PATH = Path("..") / ".." / "02_assets" / "dake_icon.ico"
MAILTO_SAFE_CHARS = ",@._+-"


def choose_font_family(root: tk.Tk) -> str:
    available = set(tkfont.families(root))
    for family in ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo"):
        if family in available:
            return family
    return "TkDefaultFont"


def get_application_directory() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def get_config_path() -> Path:
    return get_application_directory() / CONFIG_FILE_NAME


def normalize_config_value(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def write_default_config(config_path: Path) -> None:
    payload = json.dumps(DEFAULT_CONFIG, ensure_ascii=False, indent=2)
    config_path.write_text(f"{payload}\n", encoding="utf-8")


def load_or_create_config() -> tuple[dict[str, str], str | None]:
    config_path = get_config_path()

    if not config_path.exists():
        try:
            write_default_config(config_path)
        except OSError:
            return {"to": "", "cc": ""}, "status_config_write_failed"
        return DEFAULT_CONFIG.copy(), "status_config_created"

    try:
        raw_config = json.loads(config_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {"to": "", "cc": ""}, "status_config_invalid"

    if not isinstance(raw_config, dict):
        return {"to": "", "cc": ""}, "status_config_invalid"

    return {
        "to": normalize_config_value(raw_config.get("to")),
        "cc": normalize_config_value(raw_config.get("cc")),
    }, None


def build_mailto_url(to_address: str, cc_address: str) -> str:
    encoded_to = quote(to_address.strip(), safe=MAILTO_SAFE_CHARS)
    query_parts: list[str] = []

    if cc_address.strip():
        encoded_cc = quote(cc_address.strip(), safe=MAILTO_SAFE_CHARS)
        query_parts.append(f"cc={encoded_cc}")

    if not query_parts:
        return f"mailto:{encoded_to}"

    return f"mailto:{encoded_to}?{'&'.join(query_parts)}"


class AllStaffMailApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.configure(bg=COLORS["background"])
        self.root.minsize(560, 390)
        self.root.resizable(False, False)

        self.font_family = choose_font_family(root)
        self.fonts = {
            "title": (self.font_family, 20, "bold"),
            "body": (self.font_family, 10),
            "button": (self.font_family, 14, "bold"),
            "status": (self.font_family, 10),
            "footer": (self.font_family, 9),
        }
        self.status_var = tk.StringVar()

        self.apply_window_icon()
        self.build_ui()

        _config, status_key = load_or_create_config()
        self.set_status(status_key or "status_ready")
        self.center_window()

    def apply_window_icon(self) -> None:
        base_dir = get_application_directory()
        candidate_paths = [
            (base_dir / ICON_RELATIVE_PATH).resolve(),
            (base_dir.parent / ICON_RELATIVE_PATH).resolve(),
        ]

        for icon_path in candidate_paths:
            if not icon_path.exists():
                continue
            try:
                self.root.iconbitmap(str(icon_path))
                return
            except tk.TclError:
                continue

    def build_ui(self) -> None:
        outer = tk.Frame(self.root, bg=COLORS["background"])
        outer.pack(fill="both", expand=True, padx=26, pady=24)

        card = tk.Frame(
            outer,
            bg=COLORS["card"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
            bd=0,
        )
        card.pack(fill="both", expand=True)
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(0, weight=1)

        content = tk.Frame(card, bg=COLORS["card"])
        content.grid(row=0, column=0)

        tk.Label(
            content,
            text=UI_TEXT["main_title"],
            font=self.fonts["title"],
            fg=COLORS["text"],
            bg=COLORS["card"],
        ).pack()

        tk.Label(
            content,
            text=UI_TEXT["main_description"],
            font=self.fonts["body"],
            fg=COLORS["muted"],
            bg=COLORS["card"],
            wraplength=430,
            justify="center",
        ).pack(pady=(12, 0))

        self.open_button = tk.Button(
            content,
            text=UI_TEXT["button_open_mail"],
            command=self.open_mail,
            font=self.fonts["button"],
            fg=COLORS["card"],
            bg=COLORS["accent"],
            activeforeground=COLORS["card"],
            activebackground=COLORS["accent_hover"],
            relief="flat",
            bd=0,
            padx=44,
            pady=16,
            cursor="hand2",
            highlightthickness=0,
        )
        self.open_button.pack(pady=(30, 0))
        self.open_button.bind("<Enter>", lambda _event: self.open_button.configure(bg=COLORS["accent_hover"]))
        self.open_button.bind("<Leave>", lambda _event: self.open_button.configure(bg=COLORS["accent"]))

        self.status_label = tk.Label(
            card,
            textvariable=self.status_var,
            font=self.fonts["status"],
            fg=COLORS["muted"],
            bg=COLORS["card"],
            wraplength=460,
            justify="center",
        )
        self.status_label.grid(row=1, column=0, pady=(0, 22))

        self.build_footer(outer)

    def build_footer(self, parent: tk.Frame) -> None:
        footer = tk.Frame(parent, bg=COLORS["background"])
        footer.pack(fill="x", pady=(14, 0))

        tk.Label(
            footer,
            text=UI_TEXT["footer_left"],
            font=self.fonts["footer"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
        ).pack()

        footer_links = tk.Frame(footer, bg=COLORS["background"])
        footer_links.pack(pady=(5, 0))

        self.create_footer_link(footer_links, "footer_link_1")
        self.create_footer_text(footer_links, "footer_separator")
        self.create_footer_link(footer_links, "footer_link_2")
        self.create_footer_text(footer_links, "footer_separator")
        self.create_footer_text(footer_links, "footer_copyright")

    def create_footer_link(self, parent: tk.Widget, text_key: str) -> None:
        label = tk.Label(
            parent,
            text=UI_TEXT[text_key],
            font=self.fonts["footer"],
            fg=COLORS["accent"],
            bg=COLORS["background"],
            cursor="hand2",
        )
        label.pack(side="left")
        label.bind("<Button-1>", lambda _event, key=text_key: webbrowser.open_new_tab(LINK_URLS[key]))
        label.bind("<Enter>", lambda _event: label.configure(fg=COLORS["accent_hover"]))
        label.bind("<Leave>", lambda _event: label.configure(fg=COLORS["accent"]))

    def create_footer_text(self, parent: tk.Widget, text_key: str) -> None:
        tk.Label(
            parent,
            text=UI_TEXT[text_key],
            font=self.fonts["footer"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
        ).pack(side="left")

    def center_window(self) -> None:
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = max(0, (screen_width - width) // 2)
        y = max(0, (screen_height - height) // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def set_status(self, status_key: str, is_error: bool = False) -> None:
        status_text = UI_TEXT["status_display"].format(status=UI_TEXT[status_key])
        self.status_var.set(status_text)
        color = COLORS["error"] if is_error else COLORS["muted"]
        self.status_label.configure(fg=color)

    def open_mail(self) -> None:
        config, config_status = load_or_create_config()

        if config_status == "status_config_created":
            self.set_status(config_status)
            return

        if config_status in {"status_config_invalid", "status_config_write_failed"}:
            self.set_status(config_status, is_error=True)
            return

        to_address = config["to"].strip()
        cc_address = config["cc"].strip()

        if not to_address:
            self.set_status("error_no_to", is_error=True)
            return

        mailto_url = build_mailto_url(to_address, cc_address)

        try:
            opened = webbrowser.open(mailto_url, new=1)
        except Exception:
            opened = False

        if opened:
            self.set_status("status_opened")
            return

        self.set_status("status_open_failed", is_error=True)

    def run(self) -> None:
        self.root.mainloop()


def main() -> None:
    root = tk.Tk()
    AllStaffMailApp(root).run()


if __name__ == "__main__":
    main()
