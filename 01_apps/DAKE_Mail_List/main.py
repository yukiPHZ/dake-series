# -*- coding: utf-8 -*-
from __future__ import annotations

import csv
import ctypes
import html
import json
import os
import queue
import re
import sys
import threading
from dataclasses import dataclass
from datetime import datetime
from email.header import decode_header, make_header
from email.utils import parseaddr
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    ROOT_CLASS = TkinterDnD.Tk
    DND_READY = True
except ImportError:
    DND_FILES = None
    ROOT_CLASS = tk.Tk
    DND_READY = False


APP_NAME = "Dakeメールリスト"
WINDOW_TITLE = APP_NAME
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "main_title": "メールをCSVにする",
    "main_description": "Outlookから保存した .msg ファイルをここへドロップします。",
    "drop_area": "メールをここへドロップ",
    "button_select_folder": "CSV保存先を選ぶ",
    "button_open_folder": "保存フォルダを開く",
    "status_idle": "待機中",
    "status_loading": "読み込み中",
    "status_saved": "CSV保存完了",
    "status_error": "読み込みできませんでした",
    "dialog_folder_title": "CSV保存先を選ぶ",
    "dialog_error_title": "確認してください",
    "dialog_drop_error": ".msg ファイルをドロップしてください。",
    "dialog_dependency_error": "ドラッグ＆ドロップ用ライブラリを読み込めませんでした。requirements.txt を確認してください。",
    "dialog_open_error": "保存フォルダを開けませんでした。",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_caption": "止まらない、迷わない、すぐ終わる。",
    "footer_link_1": "戸建買取査定",
    "footer_link_2": "Instagram",
    "footer_separator": "｜",
    "footer_copyright": COPYRIGHT,
}

COLORS = {
    "base_bg": "#F6F7F9",
    "card_bg": "#FFFFFF",
    "text": "#1E2430",
    "muted": "#667085",
    "border": "#E6EAF0",
    "accent": "#2F6FED",
    "accent_hover": "#2458BF",
    "button_hover": "#F2F4F7",
    "drop_bg": "#FBFCFE",
    "drop_border": "#CDD5DF",
    "error": "#B42318",
}

CSV_HEADER = ("会社名", "お名前", "メールアドレス")
CONFIG_FILE_NAME = "Dake_Mail_List_config.json"
ICON_RELATIVE_PATH = Path("..") / ".." / "02_assets" / "dake_icon.ico"
QUEUE_POLL_INTERVAL_MS = 80
EMAIL_RE = re.compile(
    r"(?<![\w.+-])([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})(?![\w.-])",
    re.IGNORECASE,
)
HTML_TAG_RE = re.compile(r"<[^>]+>")
COMPANY_KEYWORDS = ("株式会社", "有限会社", "合同会社", "㈱", "（株）", "(株)")
HONORIFIC_SUFFIXES = ("様", "さん", "先生")
FONT_CANDIDATES = ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo")
URL_PREFIXES = ("h" "ttp://", "h" "ttps://")
SENDER_EMAIL_KEYS = (
    "senderEmail",
    "sender_email",
    "sender" "Sm" "tpAddress",
    "sender_" "sm" "tp_address",
)


@dataclass(frozen=True)
class ContactRow:
    company: str
    name: str
    email: str


@dataclass(frozen=True)
class ProcessResult:
    output_path: Path
    row_count: int
    skipped_count: int
    failed_count: int


def set_windows_app_id() -> None:
    if not sys.platform.startswith("win"):
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Shimarisu.DakeMailList")
    except Exception:
        return


def app_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def get_config_path() -> Path:
    return app_base_dir() / CONFIG_FILE_NAME


def get_default_desktop() -> Path:
    desktop = Path.home() / "Desktop"
    if desktop.exists():
        return desktop
    return Path.home()


def load_save_folder() -> Path:
    config_path = get_config_path()
    if not config_path.exists():
        return get_default_desktop()

    try:
        payload = json.loads(config_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return get_default_desktop()

    if not isinstance(payload, dict):
        return get_default_desktop()

    raw_folder = payload.get("save_folder")
    if not raw_folder:
        return get_default_desktop()

    folder = Path(str(raw_folder)).expanduser()
    if folder.exists() and folder.is_dir():
        return folder
    return get_default_desktop()


def save_config(save_folder: Path) -> None:
    payload = {"save_folder": str(save_folder)}
    get_config_path().write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def find_icon_path() -> Path | None:
    base = app_base_dir()
    candidates = [
        base / ICON_RELATIVE_PATH,
        base.parent.parent / "02_assets" / "dake_icon.ico",
        base.parent.parent.parent / "02_assets" / "dake_icon.ico",
    ]
    for candidate in candidates:
        try:
            resolved = candidate.resolve()
        except OSError:
            continue
        if resolved.exists():
            return resolved
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
    for candidate in FONT_CANDIDATES:
        if candidate in available:
            return candidate
    return "TkDefaultFont"


def open_folder(folder: Path) -> bool:
    try:
        if os.name == "nt":
            os.startfile(str(folder))
            return True
        import webbrowser

        webbrowser.open(folder.resolve().as_uri())
        return True
    except Exception:
        return False


def decode_mime_header(value: str) -> str:
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        return value


def safe_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, bytes):
        for encoding in ("utf-8", "cp932", "latin-1"):
            try:
                return value.decode(encoding)
            except UnicodeDecodeError:
                continue
        return value.decode("utf-8", errors="replace")
    return str(value)


def body_to_text(value: object) -> str:
    text = safe_text(value)
    if "<" in text and ">" in text:
        text = HTML_TAG_RE.sub(" ", text)
    text = html.unescape(text)
    return text.replace("\r\n", "\n").replace("\r", "\n")


def normalize_space(value: str) -> str:
    return re.sub(r"[ \t\u3000]+", " ", value).strip()


def first_email_from_text(text: str) -> str:
    match = EMAIL_RE.search(text or "")
    if match is None:
        return ""
    return match.group(1).strip(".,;:<>")


def get_message_attr(message: object, names: tuple[str, ...]) -> str:
    for name in names:
        try:
            value = getattr(message, name, None)
        except Exception:
            continue
        if value:
            return safe_text(value)
    return ""


def get_header_value(message: object, header_name: str) -> str:
    wanted = header_name.lower()
    for attr_name in ("header", "headerDict", "header_dict"):
        try:
            header = getattr(message, attr_name, None)
        except Exception:
            continue
        if not header:
            continue
        if isinstance(header, str):
            for line in header.splitlines():
                if line.lower().startswith(f"{wanted}:"):
                    return line.split(":", 1)[1].strip()
            continue
        if hasattr(header, "items"):
            try:
                for key, value in header.items():
                    if str(key).lower() == wanted:
                        return safe_text(value)
            except Exception:
                continue
        if hasattr(header, "get"):
            for key in (header_name, header_name.lower(), header_name.title()):
                try:
                    value = header.get(key)
                except Exception:
                    value = None
                if value:
                    return safe_text(value)
    return ""


def extract_email(sender_email: str, sender_text: str, from_header: str, body: str) -> str:
    for candidate in (sender_email, sender_text, from_header):
        email_address = first_email_from_text(candidate)
        if email_address:
            return email_address
        parsed = parseaddr(candidate)[1]
        if first_email_from_text(parsed):
            return parsed
    return first_email_from_text(body)


def clean_name(sender_text: str, from_header: str, email_address: str) -> str:
    for candidate in (sender_text, from_header):
        decoded = decode_mime_header(candidate)
        display_name = parseaddr(decoded)[0].strip().strip("\"'")
        display_name = normalize_space(display_name)
        if not display_name:
            continue
        if email_address and display_name.lower() == email_address.lower():
            continue
        if first_email_from_text(display_name):
            continue
        for suffix in HONORIFIC_SUFFIXES:
            if display_name.endswith(suffix):
                display_name = display_name[: -len(suffix)].strip()
        return display_name
    return ""


def clean_company_line(line: str) -> str:
    line = normalize_space(line)
    line = re.sub(r"^[\-\*_・|｜\s]+", "", line)
    line = re.sub(r"^(会社名|社名|所属)\s*[:：]\s*", "", line)
    line = line.strip(" -_：:｜|")
    if not line:
        return ""
    if len(line) > 80:
        return ""
    if first_email_from_text(line):
        return ""
    if line.lower().startswith(URL_PREFIXES):
        return ""
    return line


def extract_company(body: str) -> str:
    lines = [normalize_space(line) for line in body.splitlines()]
    lines = [line for line in lines if line]
    for line in reversed(lines[-24:]):
        if any(keyword in line for keyword in COMPANY_KEYWORDS):
            return clean_company_line(line)
    return ""


def open_msg_file(path: Path) -> object:
    import extract_msg

    opener = getattr(extract_msg, "openMsg", None)
    if opener is not None:
        return opener(str(path))
    return extract_msg.Message(str(path))


def extract_contact_from_msg(path: Path) -> ContactRow:
    message = open_msg_file(path)
    try:
        sender_text = get_message_attr(message, ("sender", "senderName", "sender_name"))
        sender_email = get_message_attr(
            message,
            SENDER_EMAIL_KEYS,
        )
        from_header = get_header_value(message, "From")
        body = body_to_text(get_message_attr(message, ("body", "htmlBody", "html_body")))
    finally:
        close = getattr(message, "close", None)
        if callable(close):
            close()

    email_address = extract_email(sender_email, sender_text, from_header, body)
    name = clean_name(sender_text, from_header, email_address)
    company = extract_company(body)
    return ContactRow(company=company, name=name, email=email_address)


def build_output_path(save_folder: Path) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return save_folder / f"mail_list_{timestamp}.csv"


def write_csv(output_path: Path, rows: list[ContactRow]) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(CSV_HEADER)
        for row in rows:
            writer.writerow((row.company, row.name, row.email))


def process_msg_files(raw_paths: list[str], save_folder: Path) -> ProcessResult:
    paths = [Path(path) for path in raw_paths]
    msg_paths = [path for path in paths if path.suffix.lower() == ".msg" and path.exists() and path.is_file()]
    skipped_count = len(paths) - len(msg_paths)
    rows: list[ContactRow] = []
    failed_count = 0

    for path in msg_paths:
        try:
            rows.append(extract_contact_from_msg(path))
        except Exception:
            failed_count += 1

    if not rows:
        raise RuntimeError("no readable msg files")

    output_path = build_output_path(save_folder)
    write_csv(output_path, rows)
    return ProcessResult(
        output_path=output_path,
        row_count=len(rows),
        skipped_count=skipped_count,
        failed_count=failed_count,
    )


class MailListApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.geometry("720x540")
        self.root.minsize(640, 500)
        self.root.configure(bg=COLORS["base_bg"])
        self.root.resizable(True, True)

        self.font_family = choose_font_family(root)
        self.fonts = {
            "title": (self.font_family, 22, "bold"),
            "body": (self.font_family, 10),
            "drop": (self.font_family, 17, "bold"),
            "button": (self.font_family, 10),
            "status": (self.font_family, 10),
            "footer": (self.font_family, 8),
        }
        self.save_folder = load_save_folder()
        self.event_queue: queue.Queue[dict[str, object]] = queue.Queue()
        self.processing = False
        self.last_output_folder = self.save_folder
        self.status_var = tk.StringVar(value=UI_TEXT["status_idle"])

        apply_window_icon(self.root)
        self.build_ui()
        self.register_drop_targets()
        self.root.after(QUEUE_POLL_INTERVAL_MS, self.poll_queue)

        if not DND_READY:
            self.set_status("status_error", is_error=True)

    def build_ui(self) -> None:
        outer = tk.Frame(self.root, bg=COLORS["base_bg"])
        outer.pack(fill="both", expand=True, padx=26, pady=24)
        outer.grid_columnconfigure(0, weight=1)
        outer.grid_rowconfigure(1, weight=1)

        header = tk.Frame(outer, bg=COLORS["base_bg"])
        header.grid(row=0, column=0, sticky="ew", pady=(0, 18))
        header.grid_columnconfigure(0, weight=1)

        tk.Label(
            header,
            text=UI_TEXT["main_title"],
            font=self.fonts["title"],
            fg=COLORS["text"],
            bg=COLORS["base_bg"],
            anchor="w",
        ).grid(row=0, column=0, sticky="w")

        tk.Label(
            header,
            text=UI_TEXT["main_description"],
            font=self.fonts["body"],
            fg=COLORS["muted"],
            bg=COLORS["base_bg"],
            anchor="w",
            justify="left",
        ).grid(row=1, column=0, sticky="ew", pady=(8, 0))

        self.drop_area = tk.Frame(
            outer,
            bg=COLORS["drop_bg"],
            highlightbackground=COLORS["drop_border"],
            highlightthickness=2,
            bd=0,
        )
        self.drop_area.grid(row=1, column=0, sticky="nsew")
        self.drop_area.grid_columnconfigure(0, weight=1)
        self.drop_area.grid_rowconfigure(0, weight=1)

        self.drop_label = tk.Label(
            self.drop_area,
            text=UI_TEXT["drop_area"],
            font=self.fonts["drop"],
            fg=COLORS["muted"],
            bg=COLORS["drop_bg"],
        )
        self.drop_label.grid(row=0, column=0)

        button_row = tk.Frame(outer, bg=COLORS["base_bg"])
        button_row.grid(row=2, column=0, sticky="ew", pady=(18, 0))
        button_row.grid_columnconfigure(2, weight=1)

        self.select_folder_button = self.create_button(
            button_row,
            UI_TEXT["button_select_folder"],
            self.select_save_folder,
            primary=True,
        )
        self.select_folder_button.grid(row=0, column=0, sticky="w")

        self.open_folder_button = self.create_button(
            button_row,
            UI_TEXT["button_open_folder"],
            self.open_save_folder,
            primary=False,
        )
        self.open_folder_button.grid(row=0, column=1, sticky="w", padx=(10, 0))

        self.status_label = tk.Label(
            outer,
            textvariable=self.status_var,
            font=self.fonts["status"],
            fg=COLORS["muted"],
            bg=COLORS["base_bg"],
            anchor="w",
        )
        self.status_label.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        self.build_footer(outer)

    def build_footer(self, parent: tk.Frame) -> None:
        footer = tk.Frame(parent, bg=COLORS["base_bg"])
        footer.grid(row=4, column=0, sticky="ew", pady=(16, 0))
        footer.grid_columnconfigure(0, weight=1)

        first_line = (
            f"{UI_TEXT['footer_left']} {UI_TEXT['footer_separator']} "
            f"{UI_TEXT['footer_caption']}"
        )
        second_line = (
            f"{UI_TEXT['footer_link_1']} {UI_TEXT['footer_separator']} "
            f"{UI_TEXT['footer_link_2']} {UI_TEXT['footer_separator']} "
            f"{UI_TEXT['footer_copyright']}"
        )
        tk.Label(
            footer,
            text=first_line,
            font=self.fonts["footer"],
            fg=COLORS["muted"],
            bg=COLORS["base_bg"],
            anchor="center",
        ).grid(row=0, column=0, sticky="ew")
        tk.Label(
            footer,
            text=second_line,
            font=self.fonts["footer"],
            fg=COLORS["muted"],
            bg=COLORS["base_bg"],
            anchor="center",
        ).grid(row=1, column=0, sticky="ew", pady=(2, 0))

    def create_button(self, parent: tk.Widget, label: str, command, primary: bool) -> tk.Button:
        normal_bg = COLORS["accent"] if primary else COLORS["card_bg"]
        normal_fg = "#FFFFFF" if primary else COLORS["text"]
        hover_bg = COLORS["accent_hover"] if primary else COLORS["button_hover"]
        button = tk.Button(
            parent,
            text=label,
            command=command,
            font=self.fonts["button"],
            fg=normal_fg,
            bg=normal_bg,
            activeforeground=normal_fg,
            activebackground=hover_bg,
            relief="flat",
            bd=0,
            padx=16,
            pady=9,
            cursor="hand2",
            highlightthickness=1,
            highlightbackground=COLORS["accent"] if primary else COLORS["border"],
        )
        button._normal_bg = normal_bg  # type: ignore[attr-defined]
        button._hover_bg = hover_bg  # type: ignore[attr-defined]
        button.bind("<Enter>", self.button_enter)
        button.bind("<Leave>", self.button_leave)
        return button

    def button_enter(self, event) -> None:
        button = event.widget
        if str(button.cget("state")) == "normal":
            button.configure(bg=button._hover_bg)

    def button_leave(self, event) -> None:
        button = event.widget
        if str(button.cget("state")) == "normal":
            button.configure(bg=button._normal_bg)

    def register_drop_targets(self) -> None:
        if not DND_READY or DND_FILES is None:
            return
        for widget in (self.root, self.drop_area, self.drop_label):
            try:
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", self.handle_drop)
            except tk.TclError:
                continue

    def set_status(self, status_key: str, is_error: bool = False) -> None:
        self.status_var.set(UI_TEXT[status_key])
        self.status_label.configure(fg=COLORS["error"] if is_error else COLORS["muted"])

    def set_processing(self, processing: bool) -> None:
        self.processing = processing
        state = "disabled" if processing else "normal"
        cursor = "arrow" if processing else "hand2"
        self.select_folder_button.configure(state=state, cursor=cursor)
        self.open_folder_button.configure(state=state, cursor=cursor)

    def select_save_folder(self) -> None:
        if self.processing:
            return
        selected = filedialog.askdirectory(
            title=UI_TEXT["dialog_folder_title"],
            initialdir=str(self.save_folder),
            parent=self.root,
        )
        if not selected:
            return
        self.save_folder = Path(selected)
        self.last_output_folder = self.save_folder
        try:
            save_config(self.save_folder)
        except OSError:
            pass
        self.set_status("status_idle")

    def open_save_folder(self) -> None:
        if self.processing:
            return
        folder = self.last_output_folder if self.last_output_folder.exists() else self.save_folder
        if not open_folder(folder):
            messagebox.showerror(
                UI_TEXT["dialog_error_title"],
                UI_TEXT["dialog_open_error"],
                parent=self.root,
            )

    def handle_drop(self, event) -> None:
        if self.processing:
            return
        raw_paths = list(self.root.tk.splitlist(event.data))
        if not any(Path(path).suffix.lower() == ".msg" for path in raw_paths):
            self.set_status("status_error", is_error=True)
            messagebox.showwarning(
                UI_TEXT["dialog_error_title"],
                UI_TEXT["dialog_drop_error"],
                parent=self.root,
            )
            return
        self.start_processing(raw_paths)

    def start_processing(self, raw_paths: list[str]) -> None:
        self.set_processing(True)
        self.set_status("status_loading")
        worker = threading.Thread(
            target=self.process_worker,
            args=(raw_paths, self.save_folder),
            daemon=True,
        )
        worker.start()

    def process_worker(self, raw_paths: list[str], save_folder: Path) -> None:
        try:
            result = process_msg_files(raw_paths, save_folder)
        except Exception:
            self.event_queue.put({"type": "error"})
            return
        self.event_queue.put({"type": "done", "result": result})

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
        if event["type"] == "error":
            self.set_processing(False)
            self.set_status("status_error", is_error=True)
            return
        if event["type"] == "done":
            result = event["result"]
            if isinstance(result, ProcessResult):
                self.last_output_folder = result.output_path.parent
            self.set_processing(False)
            self.set_status("status_saved")
            open_folder(self.last_output_folder)

    def run(self) -> None:
        self.root.mainloop()


def main() -> None:
    set_windows_app_id()
    root = ROOT_CLASS()
    MailListApp(root).run()


if __name__ == "__main__":
    main()
