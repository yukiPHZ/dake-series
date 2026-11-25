# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
import webbrowser
from pathlib import Path
from tkinter import font as tkfont
from tkinter import messagebox
from urllib.parse import quote
import tkinter as tk


APP_NAME = "Dake菊田メール"
WINDOW_TITLE = "菊田にメール"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

KIKUTA_MAIL_ADDRESS = "kikuta@sakuratoshi.co.jp"

FONT_FAMILY = "BIZ UDPGothic"
FONT_SIZE = 10
TITLE_FONT_SIZE = 18
FOOTER_FONT_SIZE = 8
FONT = (FONT_FAMILY, FONT_SIZE)
FONT_BOLD = (FONT_FAMILY, FONT_SIZE, "bold")
BASE_FONT = FONT
TITLE_FONT = (FONT_FAMILY, TITLE_FONT_SIZE, "bold")
DESCRIPTION_FONT = (FONT_FAMILY, FONT_SIZE)
LABEL_FONT = FONT_BOLD
INPUT_FONT = (FONT_FAMILY, FONT_SIZE)
BUTTON_FONT = FONT_BOLD
FOOTER_FONT = (FONT_FAMILY, FOOTER_FONT_SIZE)
FONT_FALLBACKS = (FONT_FAMILY, "Yu Gothic UI", "Meiryo")

UI_TEXT = {
    "brand_series": "シンプルそれDAKEシリーズ",
    "brand_subtitle": "止まらない、迷わない、すぐ終わる。",
    "main_title": "菊田にメールする",
    "main_description": "件名と本文を入力して、菊田宛のメールを作ります。",
    "label_subject": "件名",
    "label_body": "本文",
    "placeholder_subject": "件名を入力",
    "placeholder_body": "本文を入力",
    "button_create_mail": "メールを作る",
    "status_idle": "入力してください",
    "status_ready": "メールを作成できます",
    "status_opening": "メールソフトを開いています",
    "status_complete": "メール作成画面を開きました",
    "status_error": "メールソフトを開けませんでした",
    "error_no_address": "宛先メールアドレスが設定されていません。",
    "error_mail_open_failed": "既定のメールソフトを開けませんでした。",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_subtitle": "止まらない、迷わない、すぐ終わる。",
    "footer_estimate": "戸建買取査定",
    "footer_instagram": "Instagram",
    "footer_separator": " ｜ ",
    "footer_copyright": COPYRIGHT,
}

LINK_URLS = {
    "footer_estimate": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "footer_instagram": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

COLORS = {
    "background": "#F6F7F9",
    "card": "#FFFFFF",
    "text": "#1E2430",
    "muted": "#667085",
    "border": "#E6EAF0",
    "accent": "#2F6FED",
    "accent_hover": "#2458BF",
    "success": "#12B76A",
    "error": "#D92D20",
}

def get_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def get_project_root(base_dir: Path) -> Path:
    for candidate in (base_dir, *base_dir.parents):
        if candidate.name == "DAKE_series":
            return candidate
        if (candidate / "02_assets" / "dake_icon.ico").exists():
            return candidate
    return base_dir.parents[1] if len(base_dir.parents) >= 2 else base_dir


APP_DIR = get_base_dir()
PROJECT_ROOT = get_project_root(APP_DIR)
ICON_PATH = PROJECT_ROOT / "02_assets" / "dake_icon.ico"


def choose_font_family(root: tk.Tk) -> str:
    available = set(tkfont.families(root))
    for family in FONT_FALLBACKS:
        if family in available:
            return family
    return FONT_FAMILY


def configure_app_font(root: tk.Tk) -> None:
    global FONT_FAMILY, FONT, FONT_BOLD, BASE_FONT
    global TITLE_FONT, DESCRIPTION_FONT, LABEL_FONT, INPUT_FONT, BUTTON_FONT, FOOTER_FONT

    resolved_family = choose_font_family(root)
    FONT_FAMILY = resolved_family
    FONT = (resolved_family, FONT_SIZE)
    FONT_BOLD = (resolved_family, FONT_SIZE, "bold")
    BASE_FONT = FONT
    TITLE_FONT = (resolved_family, TITLE_FONT_SIZE, "bold")
    DESCRIPTION_FONT = (resolved_family, FONT_SIZE)
    LABEL_FONT = FONT_BOLD
    INPUT_FONT = (resolved_family, FONT_SIZE)
    BUTTON_FONT = FONT_BOLD
    FOOTER_FONT = (resolved_family, FOOTER_FONT_SIZE)


def normalize_mail_body(body: str) -> str:
    return body.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")


def is_address_configured(address: str) -> bool:
    return bool(address.strip())


def build_mailto_url(address: str, subject: str, body: str) -> str:
    recipient = quote(address.strip(), safe="@._+-")
    encoded_subject = quote(subject, safe="")
    encoded_body = quote(normalize_mail_body(body), safe="")
    return f"mailto:{recipient}?subject={encoded_subject}&body={encoded_body}"


class KikutaMailApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.configure(bg=COLORS["background"])
        self.root.minsize(900, 640)
        self.root.geometry("900x640")

        configure_app_font(root)
        self.footer_link_hover_font = tkfont.Font(
            family=FONT_FAMILY,
            size=FOOTER_FONT_SIZE,
            underline=True,
        )

        self.status_var = tk.StringVar(value=UI_TEXT["status_idle"])
        self.subject_has_placeholder = True
        self.body_has_placeholder = True

        self._set_window_icon()
        self._build_ui()
        self._set_status("status_idle")

    def _set_window_icon(self) -> None:
        if not ICON_PATH.exists():
            return
        try:
            self.root.iconbitmap(str(ICON_PATH))
            self.root.iconbitmap(default=str(ICON_PATH))
        except Exception:
            return

    def _build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        outer = tk.Frame(self.root, bg=COLORS["background"])
        outer.grid(row=0, column=0, sticky="nsew", padx=32, pady=28)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)

        header = tk.Frame(outer, bg=COLORS["background"])
        header.grid(row=0, column=0, sticky="ew", pady=(0, 18))
        header.columnconfigure(0, weight=1)

        tk.Label(
            header,
            text=UI_TEXT["main_title"],
            bg=COLORS["background"],
            fg=COLORS["text"],
            font=TITLE_FONT,
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", pady=(0, 8))

        tk.Label(
            header,
            text=UI_TEXT["main_description"],
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=DESCRIPTION_FONT,
            anchor="w",
        ).grid(row=1, column=0, sticky="ew")

        card = tk.Frame(
            outer,
            bg=COLORS["card"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
            bd=0,
        )
        card.grid(row=1, column=0, sticky="nsew")
        card.columnconfigure(0, weight=1)
        card.rowconfigure(3, weight=1)

        tk.Label(
            card,
            text=UI_TEXT["label_subject"],
            bg=COLORS["card"],
            fg=COLORS["text"],
            font=LABEL_FONT,
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=28, pady=(26, 8))

        self.subject_entry = tk.Entry(
            card,
            bd=0,
            bg=COLORS["card"],
            fg=COLORS["muted"],
            font=INPUT_FONT,
            highlightbackground=COLORS["border"],
            highlightcolor=COLORS["accent"],
            highlightthickness=1,
            insertbackground=COLORS["text"],
            relief="flat",
        )
        self.subject_entry.grid(row=1, column=0, sticky="ew", padx=28, ipady=10)
        self.subject_entry.insert(0, UI_TEXT["placeholder_subject"])
        self.subject_entry.bind("<FocusIn>", self._handle_subject_focus_in)
        self.subject_entry.bind("<FocusOut>", self._handle_subject_focus_out)
        self.subject_entry.bind("<KeyRelease>", self._handle_input_change)

        tk.Label(
            card,
            text=UI_TEXT["label_body"],
            bg=COLORS["card"],
            fg=COLORS["text"],
            font=LABEL_FONT,
            anchor="w",
        ).grid(row=2, column=0, sticky="ew", padx=28, pady=(22, 8))

        self.body_text = tk.Text(
            card,
            bd=0,
            bg=COLORS["card"],
            fg=COLORS["muted"],
            font=INPUT_FONT,
            height=12,
            highlightbackground=COLORS["border"],
            highlightcolor=COLORS["accent"],
            highlightthickness=1,
            insertbackground=COLORS["text"],
            padx=12,
            pady=10,
            relief="flat",
            undo=False,
            wrap="word",
        )
        self.body_text.grid(row=3, column=0, sticky="nsew", padx=28)
        self.body_text.insert("1.0", UI_TEXT["placeholder_body"])
        self.body_text.bind("<FocusIn>", self._handle_body_focus_in)
        self.body_text.bind("<FocusOut>", self._handle_body_focus_out)
        self.body_text.bind("<KeyRelease>", self._handle_input_change)

        action_row = tk.Frame(card, bg=COLORS["card"])
        action_row.grid(row=4, column=0, sticky="ew", padx=28, pady=(22, 26))
        action_row.columnconfigure(1, weight=1)

        self.create_button = tk.Button(
            action_row,
            text=UI_TEXT["button_create_mail"],
            bg=COLORS["accent"],
            fg="#FFFFFF",
            activebackground=COLORS["accent_hover"],
            activeforeground="#FFFFFF",
            bd=0,
            command=self._create_mail,
            cursor="hand2",
            font=BUTTON_FONT,
            padx=24,
            pady=11,
            relief="flat",
        )
        self.create_button.grid(row=0, column=0, sticky="w")
        self.create_button.bind("<Enter>", self._handle_button_enter)
        self.create_button.bind("<Leave>", self._handle_button_leave)

        self.status_label = tk.Label(
            action_row,
            textvariable=self.status_var,
            bg=COLORS["card"],
            fg=COLORS["muted"],
            font=BASE_FONT,
            anchor="e",
        )
        self.status_label.grid(row=0, column=1, sticky="ew", padx=(18, 0))

        footer = tk.Frame(outer, bg=COLORS["background"])
        footer.grid(row=2, column=0, sticky="ew", pady=(16, 0))
        footer.columnconfigure(0, weight=1)
        footer.columnconfigure(1, weight=1)

        footer_left_text = f"{UI_TEXT['footer_left']} / {UI_TEXT['footer_subtitle']}"
        tk.Label(
            footer,
            text=footer_left_text,
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=FOOTER_FONT,
            anchor="w",
        ).grid(row=0, column=0, sticky="ew")

        footer_right = tk.Frame(footer, bg=COLORS["background"])
        footer_right.grid(row=0, column=1, sticky="e")

        self._build_footer_link(
            footer_right,
            0,
            UI_TEXT["footer_estimate"],
            LINK_URLS["footer_estimate"],
        )
        self._build_footer_text(footer_right, 1, UI_TEXT["footer_separator"])
        self._build_footer_link(
            footer_right,
            2,
            UI_TEXT["footer_instagram"],
            LINK_URLS["footer_instagram"],
        )
        self._build_footer_text(footer_right, 3, UI_TEXT["footer_separator"])
        self._build_footer_text(footer_right, 4, UI_TEXT["footer_copyright"])

    def _build_footer_text(self, parent: tk.Widget, column: int, text: str) -> None:
        tk.Label(
            parent,
            text=text,
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=FOOTER_FONT,
            anchor="e",
        ).grid(row=0, column=column, sticky="e")

    def _build_footer_link(
        self,
        parent: tk.Widget,
        column: int,
        text: str,
        url: str,
    ) -> None:
        label = tk.Label(
            parent,
            text=text,
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=FOOTER_FONT,
            anchor="e",
            cursor="hand2",
        )
        label.grid(row=0, column=column, sticky="e")
        label.bind("<Button-1>", lambda _event: self._open_footer_link(url))
        label.bind("<Enter>", lambda _event: self._handle_footer_link_enter(label))
        label.bind("<Leave>", lambda _event: self._handle_footer_link_leave(label))

    def _handle_subject_focus_in(self, _event: tk.Event) -> None:
        if self.subject_has_placeholder:
            self.subject_entry.delete(0, tk.END)
            self.subject_entry.configure(fg=COLORS["text"])
            self.subject_has_placeholder = False
        self._set_status("status_ready")

    def _handle_subject_focus_out(self, _event: tk.Event) -> None:
        if not self.subject_entry.get():
            self.subject_entry.insert(0, UI_TEXT["placeholder_subject"])
            self.subject_entry.configure(fg=COLORS["muted"])
            self.subject_has_placeholder = True
        self._refresh_ready_status()

    def _handle_body_focus_in(self, _event: tk.Event) -> None:
        if self.body_has_placeholder:
            self.body_text.delete("1.0", tk.END)
            self.body_text.configure(fg=COLORS["text"])
            self.body_has_placeholder = False
        self._set_status("status_ready")

    def _handle_body_focus_out(self, _event: tk.Event) -> None:
        if not self.body_text.get("1.0", "end-1c"):
            self.body_text.insert("1.0", UI_TEXT["placeholder_body"])
            self.body_text.configure(fg=COLORS["muted"])
            self.body_has_placeholder = True
        self._refresh_ready_status()

    def _handle_input_change(self, _event: tk.Event) -> None:
        self._set_status("status_ready")

    def _handle_button_enter(self, _event: tk.Event) -> None:
        self.create_button.configure(bg=COLORS["accent_hover"])

    def _handle_button_leave(self, _event: tk.Event) -> None:
        self.create_button.configure(bg=COLORS["accent"])

    def _handle_footer_link_enter(self, label: tk.Label) -> None:
        label.configure(fg=COLORS["accent"], font=self.footer_link_hover_font)

    def _handle_footer_link_leave(self, label: tk.Label) -> None:
        label.configure(fg=COLORS["muted"], font=FOOTER_FONT)

    def _open_footer_link(self, url: str) -> None:
        try:
            webbrowser.open(url)
        except Exception:
            return

    def _refresh_ready_status(self) -> None:
        if self.subject_has_placeholder and self.body_has_placeholder:
            self._set_status("status_idle")
            return
        self._set_status("status_ready")

    def _get_subject(self) -> str:
        if self.subject_has_placeholder:
            return ""
        return self.subject_entry.get()

    def _get_body(self) -> str:
        if self.body_has_placeholder:
            return ""
        return self.body_text.get("1.0", "end-1c")

    def _set_status(self, text_key: str) -> None:
        self.status_var.set(UI_TEXT[text_key])
        if not hasattr(self, "status_label"):
            return
        if text_key == "status_complete":
            self.status_label.configure(fg=COLORS["success"])
        elif text_key == "status_error":
            self.status_label.configure(fg=COLORS["error"])
        elif text_key == "status_opening":
            self.status_label.configure(fg=COLORS["accent"])
        else:
            self.status_label.configure(fg=COLORS["muted"])

    def _show_error(self, text_key: str) -> None:
        self._set_status("status_error")
        messagebox.showerror(APP_NAME, UI_TEXT[text_key])

    def _create_mail(self) -> None:
        if not is_address_configured(KIKUTA_MAIL_ADDRESS):
            self._show_error("error_no_address")
            return

        self._set_status("status_opening")
        self.root.update_idletasks()

        subject = self._get_subject()
        body = self._get_body()
        mailto_url = build_mailto_url(KIKUTA_MAIL_ADDRESS, subject, body)

        try:
            if not webbrowser.open(mailto_url):
                raise RuntimeError("webbrowser.open returned False")
        except Exception:
            self._show_error("error_mail_open_failed")
            return

        self._set_status("status_complete")


def main() -> None:
    root = tk.Tk()
    KikutaMailApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
