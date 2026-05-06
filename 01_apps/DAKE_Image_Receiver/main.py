# -*- coding: utf-8 -*-
from __future__ import annotations

import ctypes
import html
import os
import queue
import re
import socket
import subprocess
import sys
import threading
import time
import unicodedata
import webbrowser
from datetime import datetime
from pathlib import Path
from typing import Any

import qrcode
import tkinter as tk
from flask import Flask, Response, request
from PIL import Image, ImageTk
from werkzeug.exceptions import RequestEntityTooLarge
from werkzeug.serving import BaseWSGIServer, make_server


APP_NAME = "DakeImage_Receiver"
WINDOW_TITLE = "スマホから画像受け取るDAKE"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "brand_series": "シンプルそれDAKEシリーズ",
    "main_title": "スマホから画像を受け取る",
    "main_description": "QRコードを読み取って、スマホの画像をこのPCへ送ります。",
    "qr_help": "スマホのカメラで読み取ってください",
    "fallback_url_label": "読み取れない場合：",
    "button_open_folder": "保存フォルダを開く",
    "button_stop": "受信を停止",
    "button_restart": "再起動",
    "status_idle": "待機中",
    "status_receiving": "受信中",
    "status_receiving_dots": ["受信中.", "受信中..", "受信中..."],
    "status_complete": "受信完了",
    "status_complete_template": "受信完了：{count}枚",
    "status_stopped": "受信を停止しました",
    "status_error": "エラー",
    "status_error_template": "エラー：{message}",
    "status_server_starting": "起動中",
    "status_folder_open_error": "保存フォルダを開けませんでした。",
    "status_server_error": "受信サーバーを起動できませんでした。",
    "status_restart_error": "再起動できませんでした。",
    "mobile_title": "画像を送る",
    "mobile_description": "写真を選んで、このPCへ送信します。",
    "mobile_select": "画像を選ぶ",
    "mobile_send": "送信する",
    "mobile_sending": "送信中です。画面を閉じずにお待ちください。",
    "mobile_complete": "送信しました。PC側を確認してください。",
    "mobile_stopped": "現在、PC側の受信が停止しています。PC側で再起動してください。",
    "mobile_error_too_large": "枚数または容量が大きすぎます。枚数を減らしてもう一度送ってください。",
    "mobile_error_no_file": "画像を選択してください。",
    "mobile_error_unsupported": "対応していない画像形式です。",
    "mobile_error_save": "保存できませんでした。PC側を確認してください。",
    "mobile_error_button": "戻る",
    "mobile_selected_none": "画像が選択されていません",
    "mobile_selected_count_suffix": "枚選択中",
    "mobile_limit_notice": "一度に送る目安は30枚までです。",
    "mobile_limit_detail": "最大100枚まで送れますが、枚数が多い場合は分けて送ってください。",
    "mobile_size_notice": "動画や大きすぎる画像は送れません。",
    "mobile_file_input_label": "送信する画像",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_subcopy": "止まらない、迷わない、すぐ終わる。",
    "footer_link_1": "戸建買取査定",
    "footer_link_2": "Instagram",
    "footer_separator": " ｜ ",
    "footer_copyright": "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta",
}

LINK_URLS = {
    "footer_link_1": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "footer_link_2": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

THEME = {
    "background": "#F6F7F9",
    "card": "#FFFFFF",
    "text": "#1E2430",
    "muted": "#667085",
    "border": "#E6EAF0",
    "accent": "#2F6FED",
    "accent_hover": "#2458BF",
    "complete": "#12B76A",
    "error": "#D92D20",
}

FONT_CANDIDATES = ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo")
WINDOW_SIZE = "620x640"
WINDOW_WIDTH = 620
WINDOW_HEIGHT = 640
START_PORT = 8765
MAX_PORT_SCAN = 80
MAX_FILES = 100
MAX_FILE_BYTES = 50 * 1024 * 1024
MAX_TOTAL_BYTES = 500 * 1024 * 1024
READ_CHUNK_SIZE = 1024 * 1024
ALLOWED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".heic", ".webp"}
SAVE_ROOT = Path.home() / "Downloads" / "DAKE_Image_Receiver"
COMMON_ICON_RELATIVE = Path("..") / ".." / "02_assets" / "dake_icon.ico"
COMMON_ICON_FILENAME = "dake_icon.ico"
QUEUE_POLL_MS = 120
DOT_INTERVAL_MS = 420


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def find_icon_path() -> Path | None:
    base = app_dir()
    candidates = [
        base / COMMON_ICON_RELATIVE,
        base.parent.parent / "02_assets" / COMMON_ICON_FILENAME,
        base.parent.parent.parent / "02_assets" / COMMON_ICON_FILENAME,
        Path(__file__).resolve().parent / COMMON_ICON_RELATIVE,
    ]
    if getattr(sys, "frozen", False):
        candidates.append(Path(getattr(sys, "_MEIPASS", base)) / COMMON_ICON_FILENAME)
    for candidate in candidates:
        try:
            resolved = candidate.resolve()
        except OSError:
            resolved = candidate
        if resolved.exists():
            return resolved
    return None


def set_windows_app_id() -> None:
    if not sys.platform.startswith("win"):
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Shimarisu.DakeImageReceiver")
    except Exception:
        pass


def get_local_ip() -> str:
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        sock.connect(("8.8.8.8", 80))
        return sock.getsockname()[0]
    except OSError:
        try:
            return socket.gethostbyname(socket.gethostname())
        except OSError:
            return "127.0.0.1"
    finally:
        sock.close()


def find_available_port(start_port: int = START_PORT) -> int:
    for offset in range(MAX_PORT_SCAN + 1):
        port = start_port + offset
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            try:
                sock.bind(("0.0.0.0", port))
            except OSError:
                continue
            return port
    raise OSError(UI_TEXT["status_server_error"])


def make_qr_image(url: str, size: int = 260) -> ImageTk.PhotoImage:
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10,
        border=2,
    )
    qr.add_data(url)
    qr.make(fit=True)
    image = qr.make_image(fill_color=THEME["text"], back_color=THEME["card"]).convert("RGB")
    image = image.resize((size, size), Image.Resampling.NEAREST)
    return ImageTk.PhotoImage(image)


def ensure_save_root() -> Path:
    SAVE_ROOT.mkdir(parents=True, exist_ok=True)
    return SAVE_ROOT


def safe_filename(raw_name: str) -> str:
    try:
        raw_name = raw_name.encode("latin-1").decode("utf-8")
    except UnicodeError:
        pass
    name = raw_name.replace("\\", "/").split("/")[-1]
    name = unicodedata.normalize("NFC", name)
    name = re.sub(r"[\x00-\x1f\x7f]", "", name)
    name = re.sub(r'[<>:"/\\|?*]+', "_", name)
    name = name.strip().strip(".")
    if not name:
        name = "image"
    reserved = {
        "CON",
        "PRN",
        "AUX",
        "NUL",
        "COM1",
        "COM2",
        "COM3",
        "COM4",
        "COM5",
        "COM6",
        "COM7",
        "COM8",
        "COM9",
        "LPT1",
        "LPT2",
        "LPT3",
        "LPT4",
        "LPT5",
        "LPT6",
        "LPT7",
        "LPT8",
        "LPT9",
    }
    stem = Path(name).stem
    if stem.upper() in reserved:
        name = f"_{name}"
    return name


def allowed_extension(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


def unique_path(folder: Path, filename: str) -> Path:
    safe_name = safe_filename(filename)
    candidate = folder / safe_name
    suffix = candidate.suffix
    stem = candidate.stem or "image"
    index = 2
    while candidate.exists():
        candidate = folder / f"{stem}_{index}{suffix}"
        index += 1
    return candidate


def create_received_folder() -> Path:
    ensure_save_root()
    while True:
        folder = SAVE_ROOT / f"received_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        try:
            folder.mkdir(parents=False, exist_ok=False)
            return folder
        except FileExistsError:
            time.sleep(0.12)


def open_folder(path: Path) -> None:
    if sys.platform.startswith("win"):
        os.startfile(str(path))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


def html_page(body: str, title: str | None = None) -> Response:
    page_title = html.escape(title or UI_TEXT["mobile_title"])
    content = f"""<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{page_title}</title>
<style>
:root {{
  color-scheme: light;
  --bg: {THEME["background"]};
  --card: {THEME["card"]};
  --text: {THEME["text"]};
  --muted: {THEME["muted"]};
  --border: {THEME["border"]};
  --accent: {THEME["accent"]};
  --accent-hover: {THEME["accent_hover"]};
  --complete: {THEME["complete"]};
}}
* {{ box-sizing: border-box; }}
body {{
  margin: 0;
  min-height: 100vh;
  background: var(--bg);
  color: var(--text);
  font-family: "BIZ UDPGothic", "Yu Gothic UI", Meiryo, sans-serif;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 24px;
}}
main {{
  width: min(100%, 420px);
  background: var(--card);
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 26px 22px;
}}
h1 {{
  margin: 0 0 10px;
  font-size: 24px;
  line-height: 1.35;
  letter-spacing: 0;
}}
p {{
  margin: 0 0 20px;
  color: var(--muted);
  font-size: 14px;
  line-height: 1.7;
}}
input[type=file] {{
  position: absolute;
  inset: 0;
  width: 100%;
  height: 100%;
  opacity: 0;
  cursor: pointer;
}}
.select-button,
button,
.back-link {{
  width: 100%;
  min-height: 48px;
  border-radius: 8px;
  border: 1px solid var(--accent);
  font-size: 16px;
  font-weight: 700;
  font-family: inherit;
  display: flex;
  align-items: center;
  justify-content: center;
  text-decoration: none;
}}
.select-button {{
  position: relative;
  overflow: hidden;
  background: var(--card);
  color: var(--accent);
  margin-bottom: 12px;
  cursor: pointer;
}}
button {{
  background: var(--accent);
  color: #fff;
  cursor: pointer;
}}
button:active {{
  background: var(--accent-hover);
}}
.status {{
  min-height: 22px;
  margin: 12px 0 0;
  color: var(--muted);
  font-size: 13px;
  line-height: 1.6;
}}
.selected-count {{
  margin: 0 0 14px;
  padding: 12px 14px;
  border: 1px solid var(--border);
  border-radius: 8px;
  background: #F8FAFC;
  color: var(--text);
  font-size: 15px;
  font-weight: 700;
  line-height: 1.5;
}}
.notice {{
  margin: 14px 0 18px;
  color: var(--muted);
  font-size: 12px;
  line-height: 1.65;
}}
.notice p {{
  margin: 0 0 4px;
  color: inherit;
  font-size: inherit;
  line-height: inherit;
}}
.complete {{
  color: var(--complete);
  font-weight: 700;
}}
.error {{
  color: #D92D20;
  font-weight: 700;
}}
.back-link {{
  color: var(--accent);
  background: var(--card);
}}
</style>
</head>
<body>
<main>{body}</main>
</body>
</html>"""
    return Response(content, mimetype="text/html; charset=utf-8")


def build_upload_page() -> Response:
    title = html.escape(UI_TEXT["mobile_title"])
    description = html.escape(UI_TEXT["mobile_description"])
    select_label = html.escape(UI_TEXT["mobile_select"])
    send_label = html.escape(UI_TEXT["mobile_send"])
    file_label = html.escape(UI_TEXT["mobile_file_input_label"])
    selected_none = html.escape(UI_TEXT["mobile_selected_none"])
    selected_count_suffix = html.escape(UI_TEXT["mobile_selected_count_suffix"])
    limit_notice = html.escape(UI_TEXT["mobile_limit_notice"])
    limit_detail = html.escape(UI_TEXT["mobile_limit_detail"])
    size_notice = html.escape(UI_TEXT["mobile_size_notice"])
    body = f"""
<h1>{title}</h1>
<p>{description}</p>
<form method="POST" action="/upload" enctype="multipart/form-data">
  <label class="select-button" for="fileInput">
    <span>{select_label}</span>
    <input id="fileInput" type="file" name="files" accept="image/*" multiple aria-label="{file_label}">
  </label>
  <div id="selectedCount" class="selected-count">{selected_none}</div>
  <div class="notice">
    <p>{limit_notice}</p>
    <p>{limit_detail}</p>
    <p>{size_notice}</p>
  </div>
  <button type="submit">{send_label}</button>
</form>
<script>
const fileInput = document.getElementById("fileInput");
const selectedCount = document.getElementById("selectedCount");
const selectedNoneText = "{selected_none}";
const selectedCountSuffix = "{selected_count_suffix}";
fileInput.addEventListener("change", function () {{
  const count = fileInput.files.length;
  if (count === 0) {{
    selectedCount.textContent = selectedNoneText;
  }} else {{
    selectedCount.textContent = count + selectedCountSuffix;
  }}
}});
</script>
"""
    return html_page(body)


def build_message_page(message_key: str, class_name: str = "complete", back: bool = False) -> Response:
    title = html.escape(UI_TEXT["mobile_title"])
    description = html.escape(UI_TEXT[message_key])
    class_attr = html.escape(class_name)
    body = f"""
<h1>{title}</h1>
<p class="{class_attr}">{description}</p>
"""
    if back:
        body += f"""<a class="back-link" href="/">{html.escape(UI_TEXT["mobile_error_button"])}</a>"""
    return html_page(body)


class ReceiverState:
    def __init__(self) -> None:
        self.accepting = True
        self.lock = threading.Lock()


class ServerController:
    def __init__(self, notify_queue: queue.Queue[dict[str, Any]], state: ReceiverState | None = None) -> None:
        self.notify_queue = notify_queue
        self.state = state or ReceiverState()
        self.server: BaseWSGIServer | None = None
        self.thread: threading.Thread | None = None
        self.save_lock = threading.Lock()
        self.app = self._create_app()

    def _create_app(self) -> Flask:
        app = Flask(APP_NAME)
        app.config["MAX_CONTENT_LENGTH"] = MAX_TOTAL_BYTES + (8 * 1024 * 1024)

        @app.errorhandler(RequestEntityTooLarge)
        def request_too_large(_error: RequestEntityTooLarge) -> Response:
            self._notify_error(UI_TEXT["mobile_error_too_large"])
            return build_message_page("mobile_error_too_large", "error", True)

        @app.route("/", methods=["GET"])
        def index() -> Response:
            if not self.state.accepting:
                return build_message_page("mobile_stopped", "error", False)
            return build_upload_page()

        @app.route("/", methods=["POST"])
        @app.route("/upload", methods=["POST"])
        def upload() -> Response:
            if not self.state.accepting:
                return build_message_page("mobile_stopped", "error", False)
            return self._handle_upload()

        return app

    def start(self, port: int) -> None:
        if self.server is not None:
            return
        self.server = make_server("0.0.0.0", port, self.app, threaded=True)
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()

    def shutdown(self) -> None:
        if self.server is None:
            return
        self.server.shutdown()
        if self.thread is not None:
            self.thread.join(timeout=2.0)
        self.server = None
        self.thread = None

    def stop_receiving(self) -> None:
        self.state.accepting = False

    def restart_receiving(self) -> None:
        self.state.accepting = True

    def _notify_error(self, message: str) -> None:
        self.notify_queue.put({"type": "error", "message": message})

    def _handle_upload(self) -> Response:
        content_length = request.content_length
        if content_length is not None and content_length > MAX_TOTAL_BYTES + (8 * 1024 * 1024):
            self._notify_error(UI_TEXT["mobile_error_too_large"])
            return build_message_page("mobile_error_too_large", "error", True)

        files = request.files.getlist("files")
        if not files:
            files = request.files.getlist("images")
        files = [item for item in files if item and item.filename]
        if not files:
            self._notify_error(UI_TEXT["mobile_error_no_file"])
            return build_message_page("mobile_error_no_file", "error", True)
        if len(files) > MAX_FILES:
            self._notify_error(UI_TEXT["mobile_error_too_large"])
            return build_message_page("mobile_error_too_large", "error", True)
        if not all(allowed_extension(file.filename or "") for file in files):
            self._notify_error(UI_TEXT["mobile_error_unsupported"])
            return build_message_page("mobile_error_unsupported", "error", True)

        self.notify_queue.put({"type": "receiving"})
        folder: Path | None = None
        total_bytes = 0
        saved_count = 0
        try:
            with self.save_lock:
                folder = create_received_folder()
                for file_storage in files:
                    destination = unique_path(folder, file_storage.filename or "")
                    current_size = 0
                    with destination.open("wb") as output:
                        while True:
                            chunk = file_storage.stream.read(READ_CHUNK_SIZE)
                            if not chunk:
                                break
                            current_size += len(chunk)
                            total_bytes += len(chunk)
                            if current_size > MAX_FILE_BYTES or total_bytes > MAX_TOTAL_BYTES:
                                raise ValueError(UI_TEXT["mobile_error_too_large"])
                            output.write(chunk)
                    saved_count += 1
        except ValueError as exc:
            if folder is not None:
                self._remove_failed_folder(folder)
            message = str(exc) or UI_TEXT["mobile_error_too_large"]
            self._notify_error(message)
            return build_message_page("mobile_error_too_large", "error", True)
        except Exception:
            if folder is not None:
                self._remove_failed_folder(folder)
            self._notify_error(UI_TEXT["mobile_error_save"])
            return build_message_page("mobile_error_save", "error", True)

        self.notify_queue.put({"type": "complete", "count": saved_count})
        return build_message_page("mobile_complete", "complete", False)

    def _remove_failed_folder(self, folder: Path) -> None:
        try:
            for child in folder.iterdir():
                if child.is_file():
                    child.unlink()
            folder.rmdir()
        except Exception:
            pass


class DakeImageReceiverApp:
    def __init__(self) -> None:
        set_windows_app_id()
        ensure_save_root()
        self.root = tk.Tk()
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.resizable(False, False)
        self.root.configure(bg=THEME["background"])

        self.font_family = self._choose_font()
        self.root.option_add("*Font", (self.font_family, 10))

        self.notify_queue: queue.Queue[dict[str, Any]] = queue.Queue()
        self.state = ReceiverState()
        self.server = ServerController(self.notify_queue, self.state)
        self.local_ip = get_local_ip()
        self.port = find_available_port()
        self.url = f"http://{self.local_ip}:{self.port}/"
        self.qr_photo: ImageTk.PhotoImage | None = None
        self.status_var = tk.StringVar(value=UI_TEXT["status_idle"])
        self.receiving = False
        self.dot_index = 0

        self._apply_window_icon()
        self._build_ui()
        self._start_server()
        self.root.protocol("WM_DELETE_WINDOW", self._close)
        self.root.after(QUEUE_POLL_MS, self._poll_queue)
        self.root.after(DOT_INTERVAL_MS, self._tick_receiving_status)

    def _choose_font(self) -> str:
        import tkinter.font as tkfont

        available = set(tkfont.families(self.root))
        for family in FONT_CANDIDATES:
            if family in available:
                return family
        return FONT_CANDIDATES[-1]

    def _apply_window_icon(self) -> None:
        icon = find_icon_path()
        if icon is None:
            return
        try:
            self.root.iconbitmap(str(icon))
        except tk.TclError:
            pass

    def _build_ui(self) -> None:
        outer = tk.Frame(self.root, bg=THEME["background"])
        outer.pack(fill="both", expand=True, padx=28, pady=(20, 12))

        header = tk.Frame(outer, bg=THEME["background"])
        header.pack(fill="x")
        title = tk.Label(
            header,
            text=UI_TEXT["main_title"],
            bg=THEME["background"],
            fg=THEME["text"],
            font=(self.font_family, 20, "bold"),
            anchor="w",
        )
        title.pack(fill="x", anchor="w")
        description = tk.Label(
            header,
            text=UI_TEXT["main_description"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 10),
            anchor="w",
            justify="left",
            wraplength=560,
        )
        description.pack(fill="x", pady=(5, 0))

        self.qr_card = tk.Frame(
            outer,
            bg=THEME["card"],
            highlightthickness=1,
            highlightbackground=THEME["border"],
        )
        self.qr_card.pack(fill="x", pady=(18, 12))

        self.qr_area = tk.Frame(self.qr_card, bg=THEME["card"], width=320, height=306)
        self.qr_area.pack(pady=(18, 10))
        self.qr_area.pack_propagate(False)

        self.qr_label = tk.Label(self.qr_area, bg=THEME["card"])
        self.qr_label.pack(expand=True)

        self.qr_help_label = tk.Label(
            self.qr_card,
            text=UI_TEXT["qr_help"],
            bg=THEME["card"],
            fg=THEME["text"],
            font=(self.font_family, 11, "bold"),
        )
        self.qr_help_label.pack()

        fallback_row = tk.Frame(self.qr_card, bg=THEME["card"])
        fallback_row.pack(fill="x", padx=24, pady=(8, 16))
        tk.Label(
            fallback_row,
            text=UI_TEXT["fallback_url_label"],
            bg=THEME["card"],
            fg=THEME["muted"],
            font=(self.font_family, 8),
        ).pack(side="left")
        self.url_label = tk.Label(
            fallback_row,
            text=self.url,
            bg=THEME["card"],
            fg=THEME["muted"],
            font=(self.font_family, 8),
            anchor="w",
        )
        self.url_label.pack(side="left", fill="x", expand=True)

        button_row = tk.Frame(outer, bg=THEME["background"])
        button_row.pack(fill="x")

        self.open_button = self._button(
            button_row,
            UI_TEXT["button_open_folder"],
            self._open_save_folder,
            kind="secondary",
        )
        self.open_button.pack(side="left")

        right_buttons = tk.Frame(button_row, bg=THEME["background"])
        right_buttons.pack(side="right")
        self.stop_button = self._button(right_buttons, UI_TEXT["button_stop"], self._stop_receiving, kind="secondary")
        self.stop_button.pack(side="left", padx=(0, 8))
        self.restart_button = self._button(
            right_buttons,
            UI_TEXT["button_restart"],
            self._restart_receiving,
            kind="primary",
        )
        self.restart_button.pack(side="left")
        self.restart_button.configure(state=tk.DISABLED)

        spacer = tk.Frame(outer, bg=THEME["background"])
        spacer.pack(fill="both", expand=True)

        status_row = tk.Frame(outer, bg=THEME["background"])
        status_row.pack(fill="x", pady=(4, 8))
        self.status_label = tk.Label(
            status_row,
            textvariable=self.status_var,
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 9, "bold"),
            anchor="w",
        )
        self.status_label.pack(side="left")

        self._build_footer(outer)
        self._show_qr()

    def _button(self, parent: tk.Widget, label: str, command: Any, kind: str) -> tk.Button:
        if kind == "primary":
            bg = THEME["accent"]
            fg = "#FFFFFF"
            active_bg = THEME["accent_hover"]
            relief = "flat"
            border = 0
        else:
            bg = THEME["card"]
            fg = THEME["text"]
            active_bg = "#EEF4FF"
            relief = "solid"
            border = 1
        return tk.Button(
            parent,
            text=label,
            command=command,
            bg=bg,
            fg=fg,
            activebackground=active_bg,
            activeforeground=fg,
            disabledforeground="#A8B0BF",
            font=(self.font_family, 9, "bold"),
            bd=border,
            relief=relief,
            highlightthickness=0,
            width=14,
            padx=12,
            pady=9,
            cursor="hand2",
        )

    def _build_footer(self, parent: tk.Frame) -> None:
        footer = tk.Frame(parent, bg=THEME["background"], height=66)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)
        tk.Label(
            footer,
            text=UI_TEXT["footer_left"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 8, "bold"),
            anchor="center",
        ).pack(fill="x")
        tk.Label(
            footer,
            text=UI_TEXT["footer_subcopy"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 7),
            anchor="center",
        ).pack(fill="x", pady=(1, 0))

        link_row = tk.Frame(footer, bg=THEME["background"])
        link_row.pack(anchor="center", pady=(2, 0))
        self._footer_link(link_row, UI_TEXT["footer_link_1"], LINK_URLS["footer_link_1"])
        self._footer_text(link_row, UI_TEXT["footer_separator"])
        self._footer_link(link_row, UI_TEXT["footer_link_2"], LINK_URLS["footer_link_2"])
        self._footer_text(link_row, UI_TEXT["footer_separator"])
        self._footer_text(link_row, UI_TEXT["footer_copyright"])

    def _footer_text(self, parent: tk.Frame, label: str) -> None:
        tk.Label(
            parent,
            text=label,
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 7),
        ).pack(side="left")

    def _footer_link(self, parent: tk.Frame, label: str, url: str) -> None:
        item = tk.Label(
            parent,
            text=label,
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 7),
            cursor="hand2",
        )
        item.pack(side="left")
        item.bind("<Enter>", lambda _event: item.configure(fg=THEME["accent"]))
        item.bind("<Leave>", lambda _event: item.configure(fg=THEME["muted"]))
        item.bind("<Button-1>", lambda _event: webbrowser.open(url, new=2))

    def _start_server(self) -> None:
        self._set_status(UI_TEXT["status_server_starting"])
        try:
            self.server.start(self.port)
        except Exception:
            self._set_error(UI_TEXT["status_server_error"])
            self.stop_button.configure(state=tk.DISABLED)
            self.restart_button.configure(state=tk.NORMAL)
            self._show_stopped()
            return
        self._set_status(UI_TEXT["status_idle"])

    def _show_qr(self) -> None:
        self.qr_photo = make_qr_image(self.url)
        self.qr_label.configure(image=self.qr_photo, text="")
        self.qr_help_label.configure(text=UI_TEXT["qr_help"], fg=THEME["text"])
        self.url_label.configure(text=self.url)

    def _show_stopped(self) -> None:
        self.qr_label.configure(
            image="",
            text=UI_TEXT["status_stopped"],
            bg=THEME["card"],
            fg=THEME["muted"],
            font=(self.font_family, 16, "bold"),
        )
        self.qr_help_label.configure(text=UI_TEXT["status_stopped"], fg=THEME["muted"])

    def _stop_receiving(self) -> None:
        self.server.stop_receiving()
        self.receiving = False
        self._set_status(UI_TEXT["status_stopped"])
        self._show_stopped()
        self.stop_button.configure(state=tk.DISABLED)
        self.restart_button.configure(state=tk.NORMAL)

    def _restart_receiving(self) -> None:
        try:
            if self.server.server is None:
                self.port = find_available_port()
                self.url = f"http://{self.local_ip}:{self.port}/"
                self.server.start(self.port)
            self.server.restart_receiving()
        except Exception:
            self._set_error(UI_TEXT["status_restart_error"])
            return
        self._show_qr()
        self._set_status(UI_TEXT["status_idle"])
        self.stop_button.configure(state=tk.NORMAL)
        self.restart_button.configure(state=tk.DISABLED)

    def _open_save_folder(self) -> None:
        try:
            open_folder(ensure_save_root())
        except Exception:
            self._set_error(UI_TEXT["status_folder_open_error"])

    def _set_status(self, value: str) -> None:
        self.status_var.set(value)
        color = THEME["muted"]
        if value.startswith(UI_TEXT["status_error"]):
            color = THEME["error"]
        elif value.startswith(UI_TEXT["status_complete"]):
            color = THEME["complete"]
        self.status_label.configure(fg=color)

    def _set_error(self, message: str) -> None:
        self.receiving = False
        self._set_status(UI_TEXT["status_error_template"].format(message=message))

    def _poll_queue(self) -> None:
        while True:
            try:
                event = self.notify_queue.get_nowait()
            except queue.Empty:
                break
            event_type = event.get("type")
            if event_type == "receiving":
                self.receiving = True
                self.dot_index = 0
                self._set_status(UI_TEXT["status_receiving_dots"][self.dot_index])
            elif event_type == "complete":
                self.receiving = False
                self._set_status(UI_TEXT["status_complete_template"].format(count=event.get("count", 0)))
            elif event_type == "error":
                self._set_error(str(event.get("message", "")))
        self.root.after(QUEUE_POLL_MS, self._poll_queue)

    def _tick_receiving_status(self) -> None:
        if self.receiving:
            dots = UI_TEXT["status_receiving_dots"]
            self.dot_index = (self.dot_index + 1) % len(dots)
            self._set_status(dots[self.dot_index])
        self.root.after(DOT_INTERVAL_MS, self._tick_receiving_status)

    def _close(self) -> None:
        try:
            self.server.shutdown()
        finally:
            self.root.destroy()

    def run(self) -> None:
        self.root.mainloop()


def main() -> None:
    app = DakeImageReceiverApp()
    app.run()


if __name__ == "__main__":
    main()
