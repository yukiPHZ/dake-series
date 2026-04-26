# -*- coding: utf-8 -*-
from __future__ import annotations

import webbrowser
from pathlib import Path
from tkinter import font as tkfont
from tkinter import messagebox
from urllib.parse import quote
import tkinter as tk


APP_NAME = "Dake菊田メール"
WINDOW_TITLE = "菊田にメール"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

KIKUTA_MAIL_ADDRESS = "ここに菊田のメールアドレスを入れる"
UNSET_MAIL_ADDRESS = "ここに菊田のメールアドレスを入れる"

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
    "success": "#12B76A",
    "error": "#D92D20",
}

BASE_DIR = Path(__file__).resolve().parent
ICON_PATH = (BASE_DIR / ".." / ".." / "02_assets" / "dake_icon.ico").resolve()


def choose_font_family(root: tk.Tk) -> str:
    available = set(tkfont.families(root))
    for family in ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo"):
        if family in available:
            return family
    return "TkDefaultFont"


def normalize_mail_body(body: str) -> str:
    return body.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")


def is_address_configured(address: str) -> bool:
    mail_address = address.strip()
    return bool(mail_address) and mail_address != UNSET_MAIL_ADDRESS


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
        self.root.minsize(680, 560)
        self.root.geometry("760x640")

        self.font_family = choose_font_family(root)
        self.fonts = {
            "title": (self.font_family, 22, "bold"),
            "description": (self.font_family, 11),
            "label": (self.font_family, 10, "bold"),
            "input": (self.font_family, 11),
            "button": (self.font_family, 11, "bold"),
            "status": (self.font_family, 10),
            "footer": (self.font_family, 9),
        }

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
            self.root.iconbitmap(default=str(ICON_PATH))
        except tk.TclError:
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
            font=self.fonts["title"],
            anchor="w",
        ).grid(row=0, column=0, sticky="ew")

        tk.Label(
            header,
            text=UI_TEXT["main_description"],
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=self.fonts["description"],
            anchor="w",
        ).grid(row=1, column=0, sticky="ew", pady=(7, 0))

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
            font=self.fonts["label"],
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=28, pady=(26, 8))

        self.subject_entry = tk.Entry(
            card,
            bd=0,
            bg=COLORS["card"],
            fg=COLORS["muted"],
            font=self.fonts["input"],
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
            font=self.fonts["label"],
            anchor="w",
        ).grid(row=2, column=0, sticky="ew", padx=28, pady=(22, 8))

        self.body_text = tk.Text(
            card,
            bd=0,
            bg=COLORS["card"],
            fg=COLORS["muted"],
            font=self.fonts["input"],
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
            font=self.fonts["button"],
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
            font=self.fonts["status"],
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
            font=self.fonts["footer"],
            anchor="w",
        ).grid(row=0, column=0, sticky="ew")

        tk.Label(
            footer,
            text=UI_TEXT["footer_copyright"],
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=self.fonts["footer"],
            anchor="e",
        ).grid(row=0, column=1, sticky="ew")

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
