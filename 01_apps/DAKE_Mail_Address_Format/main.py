# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import ctypes
import re
import sys
import tkinter as tk
from pathlib import Path
from tkinter import font as tkfont, messagebox


APP_NAME = "DAKE_Mail_Address_Format"
WINDOW_TITLE = "メール整形"
COPYRIGHT = "© 2026 しまりす不動産 / Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "main_title": "メール整形",
    "main_description_1": "メールアドレスをまとめて貼り付けるだけでOKです。",
    "main_description_2": "カンマ区切り・改行・名前付き表記から自動で整えます。",
    "input_label": "貼り付け入力",
    "result_label": "整形結果",
    "button_format": "整形する",
    "button_copy": "コピー",
    "button_clear": "クリア",
    "status_idle": "貼り付けて、整形して、コピーできます。",
    "status_formatted": "{count}件のメールアドレスを整形しました。",
    "status_copied": "カンマ区切りの宛先をコピーしました。",
    "status_cleared": "入力と結果をクリアしました。",
    "dialog_warning_title": "確認してください",
    "dialog_no_address": "メールアドレスが見つかりませんでした。貼り付け内容を確認してください。",
    "dialog_copy_error": "コピーできませんでした。もう一度お試しください。",
    "launch_check_ok": "DAKE_Mail_Address_Format launch-check OK",
}

COLORS = {
    "background": "#F6F7F9",
    "panel": "#FFFFFF",
    "field": "#FBFCFE",
    "text": "#1E2430",
    "muted": "#667085",
    "border": "#D9E2EC",
    "accent": "#2F6FED",
    "accent_hover": "#2458BF",
    "copy": "#118A4E",
    "copy_hover": "#0B6F3B",
    "neutral": "#EEF2F7",
    "neutral_hover": "#E2E8F0",
    "warning": "#B42318",
}

EMAIL_RE = re.compile(
    r"(?<![A-Z0-9._%+-])([A-Z0-9._%+-]+@(?:[A-Z0-9-]+\.)+[A-Z]{2,})(?![A-Z0-9_-])",
    re.IGNORECASE,
)
APP_USER_MODEL_ID = "Shimarisu.DakeMailAddressFormat"
WINDOW_SIZE = "780x620"
WINDOW_MIN_SIZE = (660, 540)
FONT_CANDIDATES = ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo")


def extract_email_addresses(raw_text: str) -> list[str]:
    addresses: list[str] = []
    seen: set[str] = set()

    for match in EMAIL_RE.finditer(raw_text or ""):
        address = match.group(1).strip()
        key = address.lower()
        if key in seen:
            continue
        seen.add(key)
        addresses.append(address)

    return addresses


def format_email_addresses(raw_text: str) -> str:
    return ",".join(extract_email_addresses(raw_text))


def app_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def find_icon_path() -> Path | None:
    base = app_base_dir()
    candidates = [base, *base.parents]
    for parent in candidates[:5]:
        icon_path = parent / "02_assets" / "dake_icon.ico"
        if icon_path.exists():
            return icon_path
    return None


def apply_window_icon(root: tk.Tk) -> None:
    icon_path = find_icon_path()
    if icon_path is None:
        return
    try:
        root.iconbitmap(str(icon_path))
    except tk.TclError:
        return


def choose_font_family(root: tk.Tk) -> str:
    try:
        available = set(tkfont.families(root))
    except tk.TclError:
        available = set()
    for family in FONT_CANDIDATES:
        if family in available:
            return family
    return "TkDefaultFont"


def set_windows_app_id() -> None:
    if not sys.platform.startswith("win"):
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)
    except Exception:
        return


class MailAddressFormatApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(*WINDOW_MIN_SIZE)
        self.root.configure(bg=COLORS["background"])

        self.font_family = choose_font_family(root)
        self.fonts = {
            "title": (self.font_family, 22, "bold"),
            "description": (self.font_family, 10),
            "label": (self.font_family, 10, "bold"),
            "text": (self.font_family, 11),
            "button": (self.font_family, 10, "bold"),
            "status": (self.font_family, 9),
        }
        self.status_var = tk.StringVar(value=UI_TEXT["status_idle"])
        self.auto_format_job: str | None = None
        self.suspend_auto_format = False

        apply_window_icon(self.root)
        self._build_ui()
        self.input_text.bind("<<Modified>>", self._handle_input_modified)
        self.input_text.focus_set()

    def _build_ui(self) -> None:
        outer = tk.Frame(self.root, bg=COLORS["background"])
        outer.pack(fill="both", expand=True, padx=24, pady=22)
        outer.grid_columnconfigure(0, weight=1)
        outer.grid_rowconfigure(1, weight=1)
        outer.grid_rowconfigure(2, weight=1)

        header = tk.Frame(outer, bg=COLORS["background"])
        header.grid(row=0, column=0, sticky="ew", pady=(0, 16))
        header.grid_columnconfigure(0, weight=1)

        tk.Label(
            header,
            text=UI_TEXT["main_title"],
            font=self.fonts["title"],
            fg=COLORS["text"],
            bg=COLORS["background"],
            anchor="w",
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            header,
            text=UI_TEXT["main_description_1"],
            font=self.fonts["description"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
            anchor="w",
        ).grid(row=1, column=0, sticky="ew", pady=(8, 0))
        tk.Label(
            header,
            text=UI_TEXT["main_description_2"],
            font=self.fonts["description"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
            anchor="w",
        ).grid(row=2, column=0, sticky="ew", pady=(2, 0))

        self.input_text = self._create_text_panel(outer, 1, UI_TEXT["input_label"], readonly=False)
        self.result_text = self._create_text_panel(outer, 2, UI_TEXT["result_label"], readonly=True)

        actions = tk.Frame(outer, bg=COLORS["background"])
        actions.grid(row=3, column=0, sticky="ew", pady=(16, 0))
        actions.grid_columnconfigure(3, weight=1)

        self._button(actions, UI_TEXT["button_format"], self.format_input, COLORS["accent"], COLORS["accent_hover"]).grid(
            row=0, column=0, sticky="w"
        )
        self._button(actions, UI_TEXT["button_copy"], self.copy_result, COLORS["copy"], COLORS["copy_hover"]).grid(
            row=0, column=1, sticky="w", padx=(10, 0)
        )
        self._button(
            actions,
            UI_TEXT["button_clear"],
            self.clear_all,
            COLORS["neutral"],
            COLORS["neutral_hover"],
            fg=COLORS["text"],
        ).grid(row=0, column=2, sticky="w", padx=(10, 0))

        tk.Label(
            actions,
            textvariable=self.status_var,
            font=self.fonts["status"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
            anchor="e",
        ).grid(row=0, column=3, sticky="e", padx=(14, 0))

    def _create_text_panel(self, parent: tk.Widget, row: int, title: str, readonly: bool) -> tk.Text:
        panel = tk.Frame(
            parent,
            bg=COLORS["panel"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
            bd=0,
        )
        panel.grid(row=row, column=0, sticky="nsew", pady=(0, 12 if row == 1 else 0))
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(1, weight=1)

        tk.Label(
            panel,
            text=title,
            font=self.fonts["label"],
            fg=COLORS["text"],
            bg=COLORS["panel"],
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=14, pady=(12, 6))

        text_frame = tk.Frame(panel, bg=COLORS["field"])
        text_frame.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 14))
        text_frame.grid_columnconfigure(0, weight=1)
        text_frame.grid_rowconfigure(0, weight=1)

        text = tk.Text(
            text_frame,
            height=7,
            wrap="word",
            undo=not readonly,
            bg=COLORS["field"],
            fg=COLORS["text"],
            insertbackground=COLORS["text"],
            relief="flat",
            bd=0,
            padx=10,
            pady=9,
            font=self.fonts["text"],
        )
        scrollbar = tk.Scrollbar(text_frame, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=scrollbar.set)
        text.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        if readonly:
            text.configure(state="disabled", cursor="arrow")
        return text

    def _button(
        self,
        parent: tk.Widget,
        label: str,
        command,
        bg: str,
        hover_bg: str,
        fg: str = "#FFFFFF",
    ) -> tk.Button:
        button = tk.Button(
            parent,
            text=label,
            command=command,
            bg=bg,
            fg=fg,
            activebackground=hover_bg,
            activeforeground=fg,
            relief="flat",
            bd=0,
            padx=20,
            pady=9,
            cursor="hand2",
            font=self.fonts["button"],
        )
        button._normal_bg = bg  # type: ignore[attr-defined]
        button._hover_bg = hover_bg  # type: ignore[attr-defined]
        button.bind("<Enter>", self._button_enter)
        button.bind("<Leave>", self._button_leave)
        return button

    def _button_enter(self, event: tk.Event) -> None:
        button = event.widget
        if str(button.cget("state")) == "normal":
            button.configure(bg=button._hover_bg)

    def _button_leave(self, event: tk.Event) -> None:
        button = event.widget
        if str(button.cget("state")) == "normal":
            button.configure(bg=button._normal_bg)

    def _input_value(self) -> str:
        return self.input_text.get("1.0", "end-1c")

    def _set_result(self, value: str) -> None:
        self.result_text.configure(state="normal")
        self.result_text.delete("1.0", "end")
        self.result_text.insert("1.0", value)
        self.result_text.configure(state="disabled")

    def _show_no_address_warning(self) -> None:
        self.status_var.set(UI_TEXT["dialog_no_address"])
        messagebox.showwarning(
            UI_TEXT["dialog_warning_title"],
            UI_TEXT["dialog_no_address"],
            parent=self.root,
        )

    def _handle_input_modified(self, _event: tk.Event) -> None:
        if not self.input_text.edit_modified():
            return
        self.input_text.edit_modified(False)
        if self.suspend_auto_format:
            return
        if self.auto_format_job is not None:
            self.root.after_cancel(self.auto_format_job)
        self.auto_format_job = self.root.after(180, self.preview_format)

    def preview_format(self) -> None:
        self.auto_format_job = None
        addresses = extract_email_addresses(self._input_value())
        if not addresses:
            self._set_result("")
            self.status_var.set(UI_TEXT["status_idle"])
            return
        self._set_result(",".join(addresses))
        self.status_var.set(UI_TEXT["status_formatted"].format(count=len(addresses)))

    def format_input(self) -> str:
        addresses = extract_email_addresses(self._input_value())
        if not addresses:
            self._set_result("")
            self._show_no_address_warning()
            return ""

        result = ",".join(addresses)
        self._set_result(result)
        self.status_var.set(UI_TEXT["status_formatted"].format(count=len(addresses)))
        return result

    def copy_result(self) -> None:
        result = self.format_input()
        if not result:
            return
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(result)
            self.root.update()
        except tk.TclError:
            messagebox.showwarning(
                UI_TEXT["dialog_warning_title"],
                UI_TEXT["dialog_copy_error"],
                parent=self.root,
            )
            return
        self.status_var.set(UI_TEXT["status_copied"])

    def clear_all(self) -> None:
        self.suspend_auto_format = True
        try:
            self.input_text.delete("1.0", "end")
            self.input_text.edit_modified(False)
            self._set_result("")
            self.status_var.set(UI_TEXT["status_cleared"])
            self.input_text.focus_set()
        finally:
            self.suspend_auto_format = False

    def run(self) -> None:
        self.root.mainloop()


def run_launch_check() -> int:
    cases = {
        "comma": (
            "tanaka@example.co.jp, suzuki@example.co.jp",
            "tanaka@example.co.jp,suzuki@example.co.jp",
        ),
        "newline": (
            "sato@example.co.jp\nyamada@example.co.jp",
            "sato@example.co.jp,yamada@example.co.jp",
        ),
        "semicolon": (
            "sato@example.co.jp; yamada@example.co.jp",
            "sato@example.co.jp,yamada@example.co.jp",
        ),
        "named": (
            "Tanaka <tanaka@example.co.jp> Suzuki <suzuki@example.co.jp>",
            "tanaka@example.co.jp,suzuki@example.co.jp",
        ),
        "mailto_markdown": (
            "[tanaka@example.co.jp](mailto:tanaka@example.co.jp), [suzuki@example.co.jp](mailto:suzuki@example.co.jp)",
            "tanaka@example.co.jp,suzuki@example.co.jp",
        ),
        "duplicate_case": (
            "TANAKA@example.co.jp, tanaka@example.co.jp, Tanaka@example.co.jp",
            "TANAKA@example.co.jp",
        ),
    }
    for label, (source, expected) in cases.items():
        actual = format_email_addresses(source)
        if actual != expected:
            raise RuntimeError(f"{label} failed: {actual!r}")

    if extract_email_addresses("no address here"):
        raise RuntimeError("empty input check failed")

    print(UI_TEXT["launch_check_ok"])
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=APP_NAME)
    parser.add_argument("--launch-check", action="store_true")
    args = parser.parse_args(argv)

    if args.launch_check:
        return run_launch_check()

    set_windows_app_id()
    root = tk.Tk()
    MailAddressFormatApp(root).run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
