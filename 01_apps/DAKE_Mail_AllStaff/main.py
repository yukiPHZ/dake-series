# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import queue
import re
import sys
import tempfile
import threading
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
    "main_description": "宛先・CC・件名入りで既定メーラーを開きます。送信はされません。",
    "button_open_mail": "メールを開く",
    "button_settings": "宛先設定",
    "button_save": "保存",
    "button_back": "戻る",
    "settings_title": "宛先設定",
    "settings_description": "メールアドレスをまとめて貼り付けるだけでOKです。",
    "label_to": "TO",
    "label_cc": "CC",
    "label_subject": "件名",
    "field_hint_addresses": "カンマ区切り・改行・名前付き表記から自動で整えます。",
    "status_display": "状態：{status}",
    "status_ready": "準備完了",
    "status_opening": "メーラー起動中",
    "status_opened": "メーラーを起動しました",
    "status_config_created": "設定ファイルを作成しました。宛先設定を確認してください。",
    "status_config_invalid": "設定ファイルを読み込めませんでした。内容を確認してください。",
    "status_config_write_failed": "設定ファイルを作成できませんでした。保存場所を確認してください。",
    "status_settings_loaded": "宛先設定を表示しました",
    "status_settings_saved": "宛先設定を保存しました",
    "status_settings_save_failed": "宛先設定を保存できませんでした。保存場所を確認してください。",
    "status_open_failed": "メーラーを起動できませんでした。既定のメールアプリを確認してください。",
    "error_no_to": "宛先が未設定です。宛先設定を確認してください。",
    "error_save_no_to": "TOを1件以上入力してから保存してください。",
    "launch_check_ok": "DAKE_Mail_AllStaff launch-check OK",
    "footer_left": "シンプルそれDAKEシリーズ ｜ 止まらない、迷わない、すぐ終わる。",
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
    "input_bg": "#FFFFFF",
}

LINK_URLS = {
    "footer_link_1": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "footer_link_2": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

CONFIG_FILE_NAME = "Dake_AllStaff_Mail_config.json"
DEFAULT_CONFIG = {
    "to": "all@example.co.jp",
    "cc": "example@example.co.jp",
    "subject": "【全社員連絡】",
}
ICON_RELATIVE_PATH = Path("..") / ".." / "02_assets" / "dake_icon.ico"
MAILTO_SAFE_CHARS = ",@._+-"
QUEUE_POLL_INTERVAL_MS = 100
STATUS_DOT_INTERVAL_MS = 350
EMAIL_PATTERN = re.compile(
    r"[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@"
    r"[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?"
    r"(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?)+"
)


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


def empty_config() -> dict[str, str]:
    return {"to": "", "cc": "", "subject": ""}


def write_config(config_path: Path, config: dict[str, str]) -> None:
    payload = json.dumps(config, ensure_ascii=False, indent=2)
    config_path.write_text(f"{payload}\n", encoding="utf-8")


def write_default_config(config_path: Path) -> None:
    write_config(config_path, DEFAULT_CONFIG)


def save_config(config: dict[str, str]) -> None:
    write_config(get_config_path(), config)


def load_or_create_config() -> tuple[dict[str, str], str | None]:
    config_path = get_config_path()

    if not config_path.exists():
        try:
            write_default_config(config_path)
        except OSError:
            return empty_config(), "status_config_write_failed"
        return DEFAULT_CONFIG.copy(), "status_config_created"

    try:
        raw_config = json.loads(config_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return empty_config(), "status_config_invalid"

    if not isinstance(raw_config, dict):
        return empty_config(), "status_config_invalid"

    return {
        "to": normalize_config_value(raw_config.get("to")),
        "cc": normalize_config_value(raw_config.get("cc")),
        "subject": normalize_config_value(raw_config.get("subject")),
    }, None


def config_addresses_to_lines(addresses: str) -> str:
    return "\n".join(part.strip() for part in addresses.split(",") if part.strip())


def extract_email_addresses(addresses: str) -> list[str]:
    extracted_addresses: list[str] = []
    seen_addresses: set[str] = set()

    for match in EMAIL_PATTERN.finditer(addresses):
        email_address = match.group(0)
        normalized_email_address = email_address.lower()

        if normalized_email_address in seen_addresses:
            continue

        seen_addresses.add(normalized_email_address)
        extracted_addresses.append(email_address)

    return extracted_addresses


def pasted_addresses_to_config(addresses: str) -> str:
    return ",".join(extract_email_addresses(addresses))


def build_mailto_url(to_address: str, cc_address: str, subject: str) -> str:
    encoded_to = quote(to_address.strip(), safe=MAILTO_SAFE_CHARS)
    query_parts: list[str] = []

    if cc_address.strip():
        encoded_cc = quote(cc_address.strip(), safe=MAILTO_SAFE_CHARS)
        query_parts.append(f"cc={encoded_cc}")

    if subject.strip():
        encoded_subject = quote(subject.strip(), safe="")
        query_parts.append(f"subject={encoded_subject}")

    if not query_parts:
        return f"mailto:{encoded_to}"

    return f"mailto:{encoded_to}?{'&'.join(query_parts)}"


class AllStaffMailApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.configure(bg=COLORS["background"])
        self.root.minsize(660, 560)
        self.root.resizable(False, False)

        self.font_family = choose_font_family(root)
        self.fonts = {
            "title": (self.font_family, 20, "bold"),
            "body": (self.font_family, 10),
            "button": (self.font_family, 14, "bold"),
            "sub_button": (self.font_family, 11, "bold"),
            "status": (self.font_family, 10),
            "footer": (self.font_family, 9),
            "field": (self.font_family, 10),
            "field_label": (self.font_family, 10, "bold"),
        }
        self.status_var = tk.StringVar()
        self.subject_var = tk.StringVar()
        self.event_queue: queue.Queue[dict[str, object]] = queue.Queue()
        self.opening_mail = False
        self.status_animation_job: str | None = None
        self.status_animation_phase = 0

        self.apply_window_icon()
        self.build_ui()

        _config, status_key = load_or_create_config()
        self.set_status(status_key or "status_ready")
        self.center_window()
        self.root.after(QUEUE_POLL_INTERVAL_MS, self.poll_queue)

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

        view_container = tk.Frame(card, bg=COLORS["card"])
        view_container.grid(row=0, column=0, sticky="nsew", padx=34, pady=(28, 12))
        view_container.grid_columnconfigure(0, weight=1)
        view_container.grid_rowconfigure(0, weight=1)

        self.main_view = tk.Frame(view_container, bg=COLORS["card"])
        self.settings_view = tk.Frame(view_container, bg=COLORS["card"])
        self.main_view.grid(row=0, column=0, sticky="nsew")
        self.settings_view.grid(row=0, column=0, sticky="nsew")

        self.build_main_view(self.main_view)
        self.build_settings_view(self.settings_view)

        self.status_label = tk.Label(
            card,
            textvariable=self.status_var,
            font=self.fonts["status"],
            fg=COLORS["muted"],
            bg=COLORS["card"],
            wraplength=520,
            justify="center",
        )
        self.status_label.grid(row=1, column=0, pady=(0, 22), padx=20)

        self.build_footer(outer)
        self.main_view.tkraise()

    def build_main_view(self, parent: tk.Frame) -> None:
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)

        content = tk.Frame(parent, bg=COLORS["card"])
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

        self.open_button = self.create_primary_button(
            content,
            "button_open_mail",
            command=self.open_mail,
            padx=44,
            pady=16,
        )
        self.open_button.pack(pady=(30, 0))

        self.settings_button = self.create_secondary_button(
            content,
            "button_settings",
            command=self.show_settings,
            padx=30,
            pady=9,
        )
        self.settings_button.pack(pady=(14, 0))

    def build_settings_view(self, parent: tk.Frame) -> None:
        parent.grid_columnconfigure(0, weight=1)

        tk.Label(
            parent,
            text=UI_TEXT["settings_title"],
            font=self.fonts["title"],
            fg=COLORS["text"],
            bg=COLORS["card"],
        ).pack(anchor="w")

        tk.Label(
            parent,
            text=UI_TEXT["settings_description"],
            font=self.fonts["body"],
            fg=COLORS["muted"],
            bg=COLORS["card"],
        ).pack(anchor="w", pady=(8, 18))

        self.create_field_label(parent, "label_to", "field_hint_addresses")
        self.to_text = self.create_address_text(parent)
        self.to_text.pack(fill="x", pady=(6, 14))

        self.create_field_label(parent, "label_cc", "field_hint_addresses")
        self.cc_text = self.create_address_text(parent)
        self.cc_text.pack(fill="x", pady=(6, 14))

        self.create_field_label(parent, "label_subject")
        self.subject_entry = tk.Entry(
            parent,
            textvariable=self.subject_var,
            font=self.fonts["field"],
            fg=COLORS["text"],
            bg=COLORS["input_bg"],
            insertbackground=COLORS["text"],
            relief="solid",
            bd=1,
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            highlightcolor=COLORS["accent"],
        )
        self.subject_entry.pack(fill="x", ipady=7, pady=(6, 18))

        button_row = tk.Frame(parent, bg=COLORS["card"])
        button_row.pack(fill="x")

        self.back_button = self.create_secondary_button(
            button_row,
            "button_back",
            command=self.show_main,
            padx=24,
            pady=8,
        )
        self.back_button.pack(side="left")

        self.save_button = self.create_primary_button(
            button_row,
            "button_save",
            command=self.save_settings,
            font_key="sub_button",
            padx=30,
            pady=9,
        )
        self.save_button.pack(side="right")

    def create_primary_button(
        self,
        parent: tk.Widget,
        text_key: str,
        command: object,
        font_key: str = "button",
        padx: int = 24,
        pady: int = 10,
    ) -> tk.Button:
        button = tk.Button(
            parent,
            text=UI_TEXT[text_key],
            command=command,
            font=self.fonts[font_key],
            fg=COLORS["card"],
            bg=COLORS["accent"],
            activeforeground=COLORS["card"],
            activebackground=COLORS["accent_hover"],
            relief="flat",
            bd=0,
            padx=padx,
            pady=pady,
            cursor="hand2",
            highlightthickness=0,
        )
        self.bind_button_hover(button, COLORS["accent"], COLORS["accent_hover"])
        return button

    def create_secondary_button(
        self,
        parent: tk.Widget,
        text_key: str,
        command: object,
        padx: int = 24,
        pady: int = 10,
    ) -> tk.Button:
        button = tk.Button(
            parent,
            text=UI_TEXT[text_key],
            command=command,
            font=self.fonts["sub_button"],
            fg=COLORS["accent"],
            bg=COLORS["card"],
            activeforeground=COLORS["accent_hover"],
            activebackground=COLORS["selection_bg"],
            relief="solid",
            bd=1,
            padx=padx,
            pady=pady,
            cursor="hand2",
            highlightthickness=0,
        )
        self.bind_button_hover(button, COLORS["card"], COLORS["selection_bg"])
        return button

    def bind_button_hover(self, button: tk.Button, normal_bg: str, hover_bg: str) -> None:
        button.bind(
            "<Enter>",
            lambda _event: button.configure(bg=hover_bg) if button.cget("state") == tk.NORMAL else None,
        )
        button.bind("<Leave>", lambda _event: button.configure(bg=normal_bg))

    def create_field_label(self, parent: tk.Widget, label_key: str, hint_key: str | None = None) -> None:
        label_row = tk.Frame(parent, bg=COLORS["card"])
        label_row.pack(fill="x")

        tk.Label(
            label_row,
            text=UI_TEXT[label_key],
            font=self.fonts["field_label"],
            fg=COLORS["text"],
            bg=COLORS["card"],
        ).pack(side="left")

        if hint_key is None:
            return

        tk.Label(
            label_row,
            text=UI_TEXT[hint_key],
            font=self.fonts["footer"],
            fg=COLORS["muted"],
            bg=COLORS["card"],
        ).pack(side="left", padx=(10, 0))

    def create_address_text(self, parent: tk.Widget) -> tk.Text:
        return tk.Text(
            parent,
            height=4,
            font=self.fonts["field"],
            fg=COLORS["text"],
            bg=COLORS["input_bg"],
            insertbackground=COLORS["text"],
            relief="solid",
            bd=1,
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            highlightcolor=COLORS["accent"],
            wrap="none",
            padx=8,
            pady=6,
        )

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
            fg=COLORS["muted"],
            bg=COLORS["background"],
            cursor="hand2",
        )
        label.pack(side="left")
        label.bind("<Button-1>", lambda _event, key=text_key: webbrowser.open_new_tab(LINK_URLS[key]))
        label.bind("<Enter>", lambda _event: label.configure(fg=COLORS["accent_hover"]))
        label.bind("<Leave>", lambda _event: label.configure(fg=COLORS["muted"]))

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

    def set_status_text(self, status_text: str, is_error: bool = False) -> None:
        self.status_var.set(UI_TEXT["status_display"].format(status=status_text))
        color = COLORS["error"] if is_error else COLORS["muted"]
        self.status_label.configure(fg=color)

    def show_main(self) -> None:
        self.main_view.tkraise()
        self.set_status("status_ready")

    def show_settings(self) -> None:
        if self.opening_mail:
            return

        config, config_status = load_or_create_config()

        if config_status in {"status_config_invalid", "status_config_write_failed"}:
            self.set_status(config_status, is_error=True)
            return

        self.populate_settings(config)
        self.settings_view.tkraise()
        self.set_status(config_status or "status_settings_loaded")

    def populate_settings(self, config: dict[str, str]) -> None:
        self.to_text.delete("1.0", tk.END)
        self.to_text.insert("1.0", config_addresses_to_lines(config["to"]))
        self.cc_text.delete("1.0", tk.END)
        self.cc_text.insert("1.0", config_addresses_to_lines(config["cc"]))
        self.subject_var.set(config["subject"])

    def save_settings(self) -> None:
        to_value = pasted_addresses_to_config(self.to_text.get("1.0", tk.END))

        if not to_value:
            self.set_status("error_save_no_to", is_error=True)
            self.to_text.focus_set()
            return

        config = {
            "to": to_value,
            "cc": pasted_addresses_to_config(self.cc_text.get("1.0", tk.END)),
            "subject": self.subject_var.get().strip(),
        }

        try:
            save_config(config)
        except OSError:
            self.set_status("status_settings_save_failed", is_error=True)
            return

        self.populate_settings(config)
        self.set_status("status_settings_saved")

    def start_status_animation(self) -> None:
        self.stop_status_animation()
        self.status_animation_phase = 0
        self.animate_status()

    def animate_status(self) -> None:
        dot_count = (self.status_animation_phase % 3) + 1
        self.set_status_text(f"{UI_TEXT['status_opening']}{'.' * dot_count}")
        self.status_animation_phase += 1
        self.status_animation_job = self.root.after(STATUS_DOT_INTERVAL_MS, self.animate_status)

    def stop_status_animation(self) -> None:
        if self.status_animation_job is None:
            return
        try:
            self.root.after_cancel(self.status_animation_job)
        except tk.TclError:
            pass
        self.status_animation_job = None

    def poll_queue(self) -> None:
        try:
            while True:
                event = self.event_queue.get_nowait()
                self.handle_queue_event(event)
        except queue.Empty:
            pass

        if self.root.winfo_exists():
            self.root.after(QUEUE_POLL_INTERVAL_MS, self.poll_queue)

    def handle_queue_event(self, event: dict[str, object]) -> None:
        if event.get("type") == "mail_open_result":
            self.finish_open_mail(bool(event.get("opened")))

    def open_mail(self) -> None:
        if self.opening_mail:
            return

        config, config_status = load_or_create_config()

        if config_status == "status_config_created":
            self.set_status(config_status)
            return

        if config_status in {"status_config_invalid", "status_config_write_failed"}:
            self.set_status(config_status, is_error=True)
            return

        to_address = config["to"].strip()
        cc_address = config["cc"].strip()
        subject = config["subject"].strip()

        if not to_address:
            self.set_status("error_no_to", is_error=True)
            return

        mailto_url = build_mailto_url(to_address, cc_address, subject)

        self.opening_mail = True
        self.open_button.configure(state=tk.DISABLED)
        self.start_status_animation()

        worker = threading.Thread(
            target=self.open_mail_worker,
            args=(mailto_url,),
            daemon=True,
        )
        worker.start()

    def open_mail_worker(self, mailto_url: str) -> None:
        try:
            opened = webbrowser.open(mailto_url, new=1)
        except Exception:
            opened = False

        self.event_queue.put({"type": "mail_open_result", "opened": opened})

    def finish_open_mail(self, opened: bool) -> None:
        self.stop_status_animation()
        self.opening_mail = False
        self.open_button.configure(state=tk.NORMAL)

        if opened:
            self.set_status("status_opened")
            return

        self.set_status("status_open_failed", is_error=True)

    def run(self) -> None:
        try:
            self.root.mainloop()
        finally:
            self.stop_status_animation()


def run_launch_check() -> int:
    raw_to = "\n".join(
        [
            "Tanaka <tanaka@example.co.jp>, suzuki@example.co.jp",
            "sato@example.co.jp; yamada@example.co.jp",
            "TANAKA@example.co.jp",
        ]
    )
    raw_cc = "Soumu <soumu@example.co.jp>; cc2@example.co.jp, CC2@example.co.jp"
    expected_config = {
        "to": "tanaka@example.co.jp,suzuki@example.co.jp,sato@example.co.jp,yamada@example.co.jp",
        "cc": "soumu@example.co.jp,cc2@example.co.jp",
        "subject": DEFAULT_CONFIG["subject"],
    }

    original_get_application_directory = get_application_directory
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        globals()["get_application_directory"] = lambda: temp_path
        try:
            root = tk.Tk()
            root.withdraw()
            app = AllStaffMailApp(root)
            app.show_settings()
            app.to_text.delete("1.0", tk.END)
            app.to_text.insert("1.0", raw_to)
            app.cc_text.delete("1.0", tk.END)
            app.cc_text.insert("1.0", raw_cc)
            app.subject_var.set(DEFAULT_CONFIG["subject"])
            app.save_settings()
            root.destroy()

            config_path = temp_path / CONFIG_FILE_NAME
            saved_config = json.loads(config_path.read_text(encoding="utf-8"))
            if saved_config != expected_config:
                raise RuntimeError(f"config mismatch: {saved_config!r}")

            restored_config, status_key = load_or_create_config()
            if status_key is not None or restored_config != expected_config:
                raise RuntimeError(f"restore mismatch: {restored_config!r}")

            mailto_url = build_mailto_url(
                restored_config["to"],
                restored_config["cc"],
                restored_config["subject"],
            )
            if not mailto_url.startswith(f"mailto:{expected_config['to']}?"):
                raise RuntimeError(f"mailto to mismatch: {mailto_url}")
            if f"cc={expected_config['cc']}" not in mailto_url:
                raise RuntimeError(f"mailto cc mismatch: {mailto_url}")
            if "subject=" not in mailto_url:
                raise RuntimeError(f"mailto subject missing: {mailto_url}")
            if "body" + "=" in mailto_url:
                raise RuntimeError(f"mailto body must be absent: {mailto_url}")

            root_empty = tk.Tk()
            root_empty.withdraw()
            empty_app = AllStaffMailApp(root_empty)
            empty_app.show_settings()
            before_config = json.loads(config_path.read_text(encoding="utf-8"))
            empty_app.to_text.delete("1.0", tk.END)
            empty_app.to_text.insert("1.0", "no address")
            empty_app.save_settings()
            after_config = json.loads(config_path.read_text(encoding="utf-8"))
            root_empty.destroy()
            if before_config != after_config:
                raise RuntimeError("empty TO save must be blocked")
        finally:
            globals()["get_application_directory"] = original_get_application_directory

    print(UI_TEXT["launch_check_ok"])
    return 0


def main(argv: list[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    if "--launch-check" in args:
        return run_launch_check()

    root = tk.Tk()
    AllStaffMailApp(root).run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
