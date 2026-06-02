from __future__ import annotations

import ctypes
import ctypes.wintypes
import datetime as dt
import json
import math
import os
import random
import re
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import asdict, dataclass
from pathlib import Path
from tkinter import BOTH, DISABLED, NORMAL, Canvas, Entry, Frame, Label, StringVar, Tk, messagebox, ttk


UI_TEXT = {
    "app_name": "note素材受信箱",
    "window_title": "note素材受信箱",
    "section_status": "同期状態",
    "section_settings": "設定",
    "slack_status": "Slack接続状態",
    "last_synced_at": "最終同期日時",
    "sync_count": "同期件数",
    "save_to": "保存先",
    "today_count": "今日の同期件数",
    "not_connected": "未接続",
    "ready": "待機中",
    "syncing": "同期中",
    "connected": "接続済み",
    "failed": "失敗",
    "never": "未実行",
    "button_sync": "今すぐ同期",
    "button_open_obsidian": "Obsidianを開く",
    "button_open_inbox": "INBOXを開く",
    "button_open_notes": "NOTESを開く",
    "button_open_articles": "ARTICLESを開く",
    "button_save_settings": "設定保存",
    "label_token": "Slack Bot Token",
    "label_channel": "Slack Channel ID",
    "label_root": "PEAKHEADZ_ROOT",
    "label_obsidian": "Obsidian実行ファイル",
    "label_interval": "同期間隔（秒）",
    "settings_saved": "設定を保存しました。",
    "missing_slack": "Slack Bot Token と Slack Channel ID を設定してください。",
    "missing_root": "PEAKHEADZ_ROOT を設定してください。",
    "sync_done": "{count}件を保存しました。",
    "sync_none": "新しいSlack素材はありません。",
    "open_failed": "開けませんでした: {path}",
    "obsidian_failed": "Obsidianを開けませんでした。設定を確認してください。",
    "soft_error": "処理に失敗しました。",
    "markdown_heading": "Slack原文",
    "tray_open": "開く",
    "tray_sync": "今すぐ同期",
    "tray_obsidian": "Obsidianを開く",
    "tray_exit": "終了",
    "self_test_ok": "SELF TEST OK",
    "launch_check_ok": "LAUNCH CHECK OK",
}


APP_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
CONFIG_PATH = APP_DIR / "data" / "note_inbox_config.json"
DEFAULT_PEAKHEADZ_ROOT = Path.home() / "Documents" / "PEAKHEADZ_ROOT"
SLACK_HISTORY_URL = "https://slack.com/api/conversations.history"
COMMON_ICON_PATH = APP_DIR.parent.parent / "02_assets" / "dake_icon.ico"
APP_NAME = "DAKE_Note_Inbox"


COLORS = {
    "bg": "#09101a",
    "panel": "#111b28",
    "panel_2": "#142033",
    "line": "#27364a",
    "text": "#eaf1f8",
    "muted": "#99a7b7",
    "accent": "#8fb8ff",
    "accent_2": "#9dd7c5",
    "danger": "#ff9a9a",
    "entry": "#0d1624",
}


@dataclass
class AppConfig:
    slack_bot_token: str = ""
    slack_channel_id: str = ""
    peakheadz_root: str = str(DEFAULT_PEAKHEADZ_ROOT)
    obsidian_path: str = ""
    sync_interval_seconds: int = 300
    slack_last_ts: str = ""
    last_synced_at: str = ""
    last_sync_count: int = 0
    today_sync_date: str = ""
    today_sync_count: int = 0


@dataclass
class SlackMessage:
    ts: str
    text: str
    user: str = ""


class ConfigStore:
    def __init__(self, path: Path = CONFIG_PATH) -> None:
        self.path = path

    def load(self) -> AppConfig:
        if not self.path.exists():
            return AppConfig()
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return AppConfig()
        base = AppConfig()
        for key in asdict(base):
            if key in data:
                setattr(base, key, data[key])
        base.sync_interval_seconds = normalize_interval(base.sync_interval_seconds)
        base.last_sync_count = safe_int(base.last_sync_count, 0)
        base.today_sync_count = safe_int(base.today_sync_count, 0)
        return base

    def save(self, config: AppConfig) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        payload = json.dumps(asdict(config), ensure_ascii=False, indent=2)
        tmp_path = self.path.with_suffix(".tmp")
        tmp_path.write_text(payload + "\n", encoding="utf-8")
        tmp_path.replace(self.path)


def safe_int(value: object, default: int) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def normalize_interval(value: object) -> int:
    seconds = safe_int(value, 300)
    if seconds < 0:
        return 0
    return min(seconds, 86_400)


def now_iso() -> str:
    return dt.datetime.now().replace(microsecond=0).isoformat()


def today_key() -> str:
    return dt.date.today().isoformat()


def app_icon_path() -> Path | None:
    if COMMON_ICON_PATH.exists():
        return COMMON_ICON_PATH
    return None


def safe_filename(value: str) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|\s]+', "_", value).strip("._")
    return cleaned[:80] or "slack"


def slack_ts_to_local(ts: str) -> dt.datetime:
    try:
        return dt.datetime.fromtimestamp(float(ts))
    except (TypeError, ValueError, OSError):
        return dt.datetime.now()


def yaml_quote(value: str) -> str:
    return json.dumps(value or "", ensure_ascii=False)


def slack_error_message(error: str) -> str:
    known = {
        "not_authed": "Slack Bot Token が設定されていません。",
        "invalid_auth": "Slack Bot Token が無効です。",
        "channel_not_found": "Slack Channel ID が見つかりません。",
        "not_in_channel": "Bot が対象チャンネルに参加していません。",
        "missing_scope": "Slack Bot Token の権限が不足しています。",
    }
    return known.get(error, f"Slack API error: {error}")


def slack_api_get(token: str, params: dict[str, str], timeout: int = 20) -> dict:
    query = urllib.parse.urlencode(params)
    request = urllib.request.Request(
        f"{SLACK_HISTORY_URL}?{query}",
        headers={
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
        },
        method="GET",
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            body = response.read().decode("utf-8")
    except urllib.error.HTTPError as exc:
        raise RuntimeError(f"HTTP {exc.code}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(str(exc.reason)) from exc
    data = json.loads(body)
    if not data.get("ok"):
        raise RuntimeError(slack_error_message(str(data.get("error", "unknown"))))
    return data


def fetch_slack_messages(config: AppConfig) -> list[SlackMessage]:
    params: dict[str, str] = {
        "channel": config.slack_channel_id,
        "limit": "100",
    }
    if config.slack_last_ts:
        params["oldest"] = config.slack_last_ts
        params["inclusive"] = "false"
    data = slack_api_get(config.slack_bot_token, params)
    messages: list[SlackMessage] = []
    for item in data.get("messages", []):
        text = str(item.get("text", ""))
        ts = str(item.get("ts", ""))
        if not ts or not text:
            continue
        messages.append(SlackMessage(ts=ts, text=text, user=str(item.get("user", ""))))
    messages.sort(key=lambda message: float(message.ts))
    return messages


def target_inbox(root_path: str) -> Path:
    return Path(root_path).expanduser() / "INBOX"


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    for index in range(2, 10_000):
        candidate = parent / f"{stem}_{index}{suffix}"
        if not candidate.exists():
            return candidate
    raise RuntimeError("filename collision limit reached")


def markdown_for_slack_message(message: SlackMessage, channel_id: str) -> str:
    local_time = slack_ts_to_local(message.ts).strftime("%Y-%m-%d %H:%M:%S")
    return "\n".join(
        [
            "---",
            "source: slack",
            f"channel_id: {yaml_quote(channel_id)}",
            f"timestamp: {yaml_quote(local_time)}",
            f"slack_ts: {yaml_quote(message.ts)}",
            "status: raw",
            "---",
            "",
            f"# {UI_TEXT['markdown_heading']}",
            "",
            message.text,
            "",
        ]
    )


def save_slack_message(root_path: str, channel_id: str, message: SlackMessage) -> Path:
    inbox = target_inbox(root_path)
    inbox.mkdir(parents=True, exist_ok=True)
    stamp = slack_ts_to_local(message.ts).strftime("%Y%m%d_%H%M%S")
    filename = f"{stamp}_{safe_filename(message.ts)}.md"
    path = unique_path(inbox / filename)
    path.write_text(markdown_for_slack_message(message, channel_id), encoding="utf-8", newline="\n")
    return path


def open_path(path: Path) -> None:
    resolved = path.expanduser()
    if not resolved.exists():
        raise FileNotFoundError(str(resolved))
    if os.name == "nt":
        os.startfile(str(resolved))  # type: ignore[attr-defined]
    else:
        subprocess.Popen(["xdg-open", str(resolved)])


class WindowsTrayIcon:
    def __init__(self, app: "NoteInboxApp") -> None:
        self.app = app
        self.active = False
        self.thread: threading.Thread | None = None
        self.hwnd = None
        self._class_atom = None
        self._callback = None

    def start(self) -> None:
        if os.name != "nt" or self.active:
            return
        self.thread = threading.Thread(target=self._run, name="note-inbox-tray", daemon=True)
        self.thread.start()

    def stop(self) -> None:
        if os.name != "nt" or not self.hwnd:
            return
        try:
            self._delete_icon()
            ctypes.windll.user32.PostMessageW(self.hwnd, 0x0010, 0, 0)
        except Exception:
            pass

    def _run(self) -> None:
        try:
            self._message_loop()
        except Exception:
            self.active = False

    def _message_loop(self) -> None:
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        kernel32.GetModuleHandleW.restype = ctypes.c_void_p
        user32.CreateWindowExW.restype = ctypes.c_void_p
        user32.CreateWindowExW.argtypes = [
            ctypes.wintypes.DWORD,
            ctypes.wintypes.LPCWSTR,
            ctypes.wintypes.LPCWSTR,
            ctypes.wintypes.DWORD,
            ctypes.c_int,
            ctypes.c_int,
            ctypes.c_int,
            ctypes.c_int,
            ctypes.c_void_p,
            ctypes.c_void_p,
            ctypes.c_void_p,
            ctypes.c_void_p,
        ]
        user32.DefWindowProcW.restype = ctypes.c_ssize_t
        user32.DefWindowProcW.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_size_t, ctypes.c_size_t]
        user32.RegisterClassW.restype = ctypes.c_ushort
        user32.TrackPopupMenu.restype = ctypes.c_uint

        WNDPROC = ctypes.WINFUNCTYPE(ctypes.c_ssize_t, ctypes.c_void_p, ctypes.c_uint, ctypes.c_size_t, ctypes.c_size_t)

        class WNDCLASS(ctypes.Structure):
            _fields_ = [
                ("style", ctypes.c_uint),
                ("lpfnWndProc", WNDPROC),
                ("cbClsExtra", ctypes.c_int),
                ("cbWndExtra", ctypes.c_int),
                ("hInstance", ctypes.c_void_p),
                ("hIcon", ctypes.c_void_p),
                ("hCursor", ctypes.c_void_p),
                ("hbrBackground", ctypes.c_void_p),
                ("lpszMenuName", ctypes.c_wchar_p),
                ("lpszClassName", ctypes.c_wchar_p),
            ]

        def wnd_proc(hwnd, msg, wparam, lparam):
            if msg == 0x0400 + 20:
                if lparam in (0x0205, 0x0206):
                    self._show_menu(hwnd)
                elif lparam == 0x0203:
                    self.app.root.after(0, self.app.show_window)
                return 0
            if msg == 0x0010:
                self._delete_icon()
                user32.DestroyWindow(hwnd)
                return 0
            if msg == 0x0002:
                user32.PostQuitMessage(0)
                return 0
            return user32.DefWindowProcW(hwnd, msg, wparam, lparam)

        self._callback = WNDPROC(wnd_proc)
        instance = kernel32.GetModuleHandleW(None)
        class_name = f"{APP_NAME}_TrayWindow"
        wndclass = WNDCLASS()
        wndclass.lpfnWndProc = self._callback
        wndclass.hInstance = instance
        wndclass.lpszClassName = class_name
        self._class_atom = user32.RegisterClassW(ctypes.byref(wndclass))
        hwnd = user32.CreateWindowExW(0, class_name, class_name, 0, 0, 0, 0, 0, None, None, instance, None)
        if not hwnd:
            return
        self.hwnd = hwnd
        self._add_icon(hwnd)
        self.active = True
        msg = ctypes.wintypes.MSG()
        while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))
        self.active = False

    def _notify_data(self, hwnd):
        class GUID(ctypes.Structure):
            _fields_ = [
                ("Data1", ctypes.c_ulong),
                ("Data2", ctypes.c_ushort),
                ("Data3", ctypes.c_ushort),
                ("Data4", ctypes.c_ubyte * 8),
            ]

        class NOTIFYICONDATA(ctypes.Structure):
            _fields_ = [
                ("cbSize", ctypes.c_ulong),
                ("hWnd", ctypes.c_void_p),
                ("uID", ctypes.c_uint),
                ("uFlags", ctypes.c_uint),
                ("uCallbackMessage", ctypes.c_uint),
                ("hIcon", ctypes.c_void_p),
                ("szTip", ctypes.c_wchar * 128),
                ("dwState", ctypes.c_ulong),
                ("dwStateMask", ctypes.c_ulong),
                ("szInfo", ctypes.c_wchar * 256),
                ("uTimeoutOrVersion", ctypes.c_uint),
                ("szInfoTitle", ctypes.c_wchar * 64),
                ("dwInfoFlags", ctypes.c_ulong),
                ("guidItem", GUID),
                ("hBalloonIcon", ctypes.c_void_p),
            ]

        shell32 = ctypes.windll.shell32
        user32 = ctypes.windll.user32
        user32.LoadImageW.restype = ctypes.c_void_p
        user32.LoadIconW.restype = ctypes.c_void_p
        icon_handle = None
        icon_path = app_icon_path()
        if icon_path:
            icon_handle = user32.LoadImageW(None, str(icon_path), 1, 0, 0, 0x00000010 | 0x00000040)
        if not icon_handle:
            icon_handle = user32.LoadIconW(None, 32512)
        data = NOTIFYICONDATA()
        data.cbSize = ctypes.sizeof(NOTIFYICONDATA)
        data.hWnd = hwnd
        data.uID = 1
        data.uFlags = 0x00000001 | 0x00000002 | 0x00000004
        data.uCallbackMessage = 0x0400 + 20
        data.hIcon = icon_handle
        data.szTip = UI_TEXT["app_name"]
        return shell32, data

    def _add_icon(self, hwnd) -> None:
        shell32, data = self._notify_data(hwnd)
        shell32.Shell_NotifyIconW(0x00000000, ctypes.byref(data))

    def _delete_icon(self) -> None:
        if not self.hwnd:
            return
        shell32, data = self._notify_data(self.hwnd)
        shell32.Shell_NotifyIconW(0x00000002, ctypes.byref(data))

    def _show_menu(self, hwnd) -> None:
        user32 = ctypes.windll.user32
        menu = user32.CreatePopupMenu()
        commands = [
            (1001, UI_TEXT["tray_open"], self.app.show_window),
            (1002, UI_TEXT["tray_sync"], self.app.sync_now),
            (1003, UI_TEXT["tray_obsidian"], self.app.open_obsidian),
            (1004, UI_TEXT["tray_exit"], self.app.exit_app),
        ]
        for command_id, label, _callback in commands:
            user32.AppendMenuW(menu, 0x00000000, command_id, label)
        point = ctypes.wintypes.POINT()
        user32.GetCursorPos(ctypes.byref(point))
        user32.SetForegroundWindow(hwnd)
        selected = user32.TrackPopupMenu(menu, 0x0100, point.x, point.y, 0, hwnd, None)
        user32.DestroyMenu(menu)
        for command_id, _label, callback in commands:
            if selected == command_id:
                self.app.root.after(0, callback)
                break


class StarField:
    def __init__(self, canvas: Canvas, count: int = 42) -> None:
        self.canvas = canvas
        self.count = count
        self.items: list[tuple[int, float, float]] = []
        self.running = True
        self.resize_after_id: str | None = None
        self.canvas.bind("<Configure>", self._on_resize, add="+")
        self._create_stars()
        self._tick()

    def _create_stars(self) -> None:
        self.resize_after_id = None
        self.canvas.update_idletasks()
        width = max(self.canvas.winfo_width(), 920)
        height = max(self.canvas.winfo_height(), 620)
        for item, _phase, _speed in self.items:
            self.canvas.delete(item)
        self.items.clear()
        for _index in range(self.count):
            x = random.randint(12, width - 12)
            y = random.randint(12, height - 12)
            size = random.choice([1, 1, 2])
            phase = random.random() * 6.28
            speed = random.uniform(0.25, 0.55)
            item = self.canvas.create_oval(x, y, x + size, y + size, fill="#53667f", outline="")
            self.items.append((item, phase, speed))

    def _on_resize(self, _event) -> None:
        if self.resize_after_id:
            self.canvas.after_cancel(self.resize_after_id)
        self.resize_after_id = self.canvas.after(350, self._create_stars)

    def _tick(self) -> None:
        if not self.running:
            return
        current = time.time()
        palette = ["#405066", "#53667f", "#7288a8", "#9fb5d6"]
        for item, phase, speed in self.items:
            value = int((1 + math.sin(current * speed + phase)) * 1.5)
            self.canvas.itemconfigure(item, fill=palette[max(0, min(value, len(palette) - 1))])
        self.canvas.after(1200, self._tick)


class NoteInboxApp:
    def __init__(self, root: Tk, store: ConfigStore | None = None) -> None:
        self.root = root
        self.store = store or ConfigStore()
        self.config = self.store.load()
        self.sync_lock = threading.Lock()
        self.status_vars: dict[str, StringVar] = {}
        self.entry_vars: dict[str, StringVar] = {}
        self.buttons: list[ttk.Button] = []
        self.tray = WindowsTrayIcon(self)
        self.auto_sync_after_id: str | None = None
        self._setup_window()
        self._build_ui()
        self._refresh_status()
        self.tray.start()
        self._schedule_auto_sync()

    def _setup_window(self) -> None:
        self.root.title(UI_TEXT["window_title"])
        self.root.geometry("980x720")
        self.root.minsize(680, 500)
        icon_path = app_icon_path()
        if icon_path:
            try:
                self.root.iconbitmap(str(icon_path))
            except Exception:
                pass
        self.root.configure(bg=COLORS["bg"])
        self.root.protocol("WM_DELETE_WINDOW", self.hide_to_tray)
        self.root.bind("<Unmap>", self._on_unmap)
        self.root.after(80, self._maximize_window)

    def _maximize_window(self) -> None:
        try:
            self.root.state("zoomed")
        except Exception:
            self.root.attributes("-zoomed", True)

    def _build_ui(self) -> None:
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TButton", padding=(12, 7), font=("Yu Gothic UI", 10))
        style.configure("Accent.TButton", padding=(14, 8), font=("Yu Gothic UI", 10, "bold"))

        canvas = Canvas(self.root, bg=COLORS["bg"], highlightthickness=0)
        canvas.pack(fill=BOTH, expand=True)
        StarField(canvas)

        frame = Frame(canvas, bg=COLORS["bg"])
        window = canvas.create_window(0, 18, anchor="n", window=frame)

        def resize_content(event) -> None:
            content_width = min(max(event.width - 72, 680), 980)
            canvas.itemconfigure(window, width=content_width)
            canvas.coords(window, event.width / 2, 18)

        canvas.bind("<Configure>", resize_content, add="+")

        title = Label(frame, text=UI_TEXT["app_name"], bg=COLORS["bg"], fg=COLORS["text"], font=("Yu Gothic UI", 20, "bold"))
        title.pack(anchor="center", pady=(24, 8))

        status_panel = Frame(frame, bg=COLORS["panel"], highlightbackground=COLORS["line"], highlightthickness=1)
        status_panel.pack(fill="x", pady=(8, 14))
        Label(status_panel, text=UI_TEXT["section_status"], bg=COLORS["panel"], fg=COLORS["accent"], font=("Yu Gothic UI", 12, "bold")).pack(anchor="w", padx=18, pady=(14, 8))

        grid = Frame(status_panel, bg=COLORS["panel"])
        grid.pack(fill="x", padx=18, pady=(0, 16))
        status_items = [
            ("slack_status", UI_TEXT["slack_status"]),
            ("last_synced_at", UI_TEXT["last_synced_at"]),
            ("sync_count", UI_TEXT["sync_count"]),
            ("today_count", UI_TEXT["today_count"]),
            ("save_to", UI_TEXT["save_to"]),
        ]
        for row, (key, label_text) in enumerate(status_items):
            Label(grid, text=label_text, bg=COLORS["panel"], fg=COLORS["muted"], font=("Yu Gothic UI", 9)).grid(row=row, column=0, sticky="w", pady=3)
            var = StringVar(value="")
            self.status_vars[key] = var
            Label(grid, textvariable=var, bg=COLORS["panel"], fg=COLORS["text"], font=("Yu Gothic UI", 10), wraplength=540, justify="left").grid(row=row, column=1, sticky="w", padx=(18, 0), pady=3)
        grid.columnconfigure(1, weight=1)

        button_panel = Frame(frame, bg=COLORS["bg"])
        button_panel.pack(anchor="center", pady=(0, 14))
        self._add_button(button_panel, UI_TEXT["button_sync"], self.sync_now, "Accent.TButton")
        self._add_button(button_panel, UI_TEXT["button_open_obsidian"], self.open_obsidian)
        self._add_button(button_panel, UI_TEXT["button_open_inbox"], self.open_inbox)
        self._add_button(button_panel, UI_TEXT["button_open_notes"], self.open_notes)
        self._add_button(button_panel, UI_TEXT["button_open_articles"], self.open_articles)

        settings_panel = Frame(frame, bg=COLORS["panel"], highlightbackground=COLORS["line"], highlightthickness=1)
        settings_panel.pack(fill="x", pady=(0, 24))
        Label(settings_panel, text=UI_TEXT["section_settings"], bg=COLORS["panel"], fg=COLORS["accent_2"], font=("Yu Gothic UI", 12, "bold")).pack(anchor="w", padx=18, pady=(14, 8))

        form = Frame(settings_panel, bg=COLORS["panel"])
        form.pack(fill="x", padx=18, pady=(0, 12))
        self._add_entry(form, "slack_bot_token", UI_TEXT["label_token"], self.config.slack_bot_token, show="*")
        self._add_entry(form, "slack_channel_id", UI_TEXT["label_channel"], self.config.slack_channel_id)
        self._add_entry(form, "peakheadz_root", UI_TEXT["label_root"], self.config.peakheadz_root)
        self._add_entry(form, "obsidian_path", UI_TEXT["label_obsidian"], self.config.obsidian_path)
        self._add_entry(form, "sync_interval_seconds", UI_TEXT["label_interval"], str(self.config.sync_interval_seconds))
        form.columnconfigure(1, weight=1)

        save_row = Frame(settings_panel, bg=COLORS["panel"])
        save_row.pack(anchor="center", pady=(0, 16))
        self._add_button(save_row, UI_TEXT["button_save_settings"], self.save_settings, "Accent.TButton")

    def _add_button(self, parent: Frame, text: str, command, style_name: str = "TButton") -> None:
        button = ttk.Button(parent, text=text, command=command, style=style_name)
        button.pack(side="left", padx=(0, 8), pady=4)
        self.buttons.append(button)

    def _add_entry(self, parent: Frame, key: str, label_text: str, value: str, show: str | None = None) -> None:
        row = len(self.entry_vars)
        Label(parent, text=label_text, bg=COLORS["panel"], fg=COLORS["muted"], font=("Yu Gothic UI", 9)).grid(row=row, column=0, sticky="w", pady=5)
        var = StringVar(value=value)
        self.entry_vars[key] = var
        entry = Entry(
            parent,
            textvariable=var,
            show=show,
            bg=COLORS["entry"],
            fg=COLORS["text"],
            insertbackground=COLORS["text"],
            relief="flat",
            highlightbackground=COLORS["line"],
            highlightcolor=COLORS["accent"],
            highlightthickness=1,
            font=("Yu Gothic UI", 10),
        )
        entry.grid(row=row, column=1, sticky="ew", padx=(18, 0), pady=5, ipady=5)

    def _refresh_status(self) -> None:
        connected = UI_TEXT["connected"] if self.config.slack_bot_token and self.config.slack_channel_id else UI_TEXT["not_connected"]
        self.status_vars["slack_status"].set(connected)
        self.status_vars["last_synced_at"].set(self.config.last_synced_at or UI_TEXT["never"])
        self.status_vars["sync_count"].set(str(self.config.last_sync_count))
        self.status_vars["today_count"].set(str(self.config.today_sync_count))
        self.status_vars["save_to"].set(str(target_inbox(self.config.peakheadz_root)))

    def _set_buttons(self, state: str) -> None:
        for button in self.buttons:
            button.configure(state=state)

    def _on_unmap(self, _event) -> None:
        if self.root.state() == "iconic":
            self.root.after(120, self.hide_to_tray)

    def hide_to_tray(self) -> None:
        self.root.withdraw()

    def show_window(self) -> None:
        self.root.deiconify()
        self._maximize_window()
        self.root.lift()
        self.root.focus_force()

    def exit_app(self) -> None:
        if self.auto_sync_after_id:
            self.root.after_cancel(self.auto_sync_after_id)
            self.auto_sync_after_id = None
        self.tray.stop()
        self.root.after(150, self.root.destroy)

    def save_settings(self) -> None:
        self.config.slack_bot_token = self.entry_vars["slack_bot_token"].get().strip()
        self.config.slack_channel_id = self.entry_vars["slack_channel_id"].get().strip()
        self.config.peakheadz_root = self.entry_vars["peakheadz_root"].get().strip()
        self.config.obsidian_path = self.entry_vars["obsidian_path"].get().strip()
        self.config.sync_interval_seconds = normalize_interval(self.entry_vars["sync_interval_seconds"].get().strip())
        self.entry_vars["sync_interval_seconds"].set(str(self.config.sync_interval_seconds))
        self.store.save(self.config)
        self._refresh_status()
        self._schedule_auto_sync()
        messagebox.showinfo(UI_TEXT["app_name"], UI_TEXT["settings_saved"])

    def sync_now(self, show_dialog: bool = True) -> None:
        self.save_settings_without_dialog()
        if not self.config.slack_bot_token or not self.config.slack_channel_id:
            if show_dialog:
                messagebox.showwarning(UI_TEXT["app_name"], UI_TEXT["missing_slack"])
            return
        if not self.config.peakheadz_root:
            if show_dialog:
                messagebox.showwarning(UI_TEXT["app_name"], UI_TEXT["missing_root"])
            return
        if not self.sync_lock.acquire(blocking=False):
            return
        self.status_vars["slack_status"].set(UI_TEXT["syncing"])
        self._set_buttons(DISABLED)
        threading.Thread(target=self._sync_worker, args=(show_dialog,), name="note-inbox-sync", daemon=True).start()

    def save_settings_without_dialog(self) -> None:
        self.config.slack_bot_token = self.entry_vars["slack_bot_token"].get().strip()
        self.config.slack_channel_id = self.entry_vars["slack_channel_id"].get().strip()
        self.config.peakheadz_root = self.entry_vars["peakheadz_root"].get().strip()
        self.config.obsidian_path = self.entry_vars["obsidian_path"].get().strip()
        self.config.sync_interval_seconds = normalize_interval(self.entry_vars["sync_interval_seconds"].get().strip())
        self.store.save(self.config)

    def _sync_worker(self, show_dialog: bool) -> None:
        try:
            messages = fetch_slack_messages(self.config)
            saved_paths: list[Path] = []
            for message in messages:
                saved_paths.append(save_slack_message(self.config.peakheadz_root, self.config.slack_channel_id, message))
            if messages:
                self.config.slack_last_ts = messages[-1].ts
            now = now_iso()
            if self.config.today_sync_date != today_key():
                self.config.today_sync_date = today_key()
                self.config.today_sync_count = 0
            self.config.last_sync_count = len(saved_paths)
            self.config.today_sync_count += len(saved_paths)
            self.config.last_synced_at = now
            self.store.save(self.config)
            self.root.after(0, lambda: self._sync_complete(len(saved_paths), None, show_dialog))
        except Exception as exc:
            self.root.after(0, lambda: self._sync_complete(0, str(exc), show_dialog))
        finally:
            self.sync_lock.release()

    def _sync_complete(self, count: int, error: str | None, show_dialog: bool) -> None:
        self._set_buttons(NORMAL)
        if error:
            self.status_vars["slack_status"].set(UI_TEXT["failed"])
            if show_dialog:
                messagebox.showerror(UI_TEXT["app_name"], f"{UI_TEXT['soft_error']}\n{error}")
        else:
            self.status_vars["slack_status"].set(UI_TEXT["connected"])
            message = UI_TEXT["sync_done"].format(count=count) if count else UI_TEXT["sync_none"]
            if show_dialog:
                messagebox.showinfo(UI_TEXT["app_name"], message)
        self._refresh_status()

    def _schedule_auto_sync(self) -> None:
        if self.auto_sync_after_id:
            self.root.after_cancel(self.auto_sync_after_id)
            self.auto_sync_after_id = None
        seconds = normalize_interval(self.config.sync_interval_seconds)
        if seconds <= 0:
            return
        self.auto_sync_after_id = self.root.after(seconds * 1000, self._auto_sync)

    def _auto_sync(self) -> None:
        self.auto_sync_after_id = None
        if self.config.slack_bot_token and self.config.slack_channel_id:
            self.sync_now(show_dialog=False)
        self._schedule_auto_sync()

    def open_obsidian(self) -> None:
        self.save_settings_without_dialog()
        root_path = Path(self.config.peakheadz_root).expanduser()
        obsidian = Path(self.config.obsidian_path).expanduser() if self.config.obsidian_path else None
        try:
            if obsidian and obsidian.exists() and obsidian.is_file():
                subprocess.Popen([str(obsidian), str(root_path)])
            else:
                raise FileNotFoundError(str(obsidian or ""))
        except Exception:
            messagebox.showerror(UI_TEXT["app_name"], UI_TEXT["obsidian_failed"])

    def open_inbox(self) -> None:
        self._open_root_child("INBOX")

    def open_notes(self) -> None:
        self._open_root_child("NOTES")

    def open_articles(self) -> None:
        self._open_root_child("ARTICLES")

    def _open_root_child(self, child: str) -> None:
        self.save_settings_without_dialog()
        path = Path(self.config.peakheadz_root).expanduser() / child
        try:
            open_path(path)
        except Exception:
            messagebox.showerror(UI_TEXT["app_name"], UI_TEXT["open_failed"].format(path=path))


def run_self_test() -> int:
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        store = ConfigStore(tmp_path / "data" / "note_inbox_config.json")
        config = AppConfig(
            slack_bot_token="test-token-placeholder",
            slack_channel_id="C123456",
            peakheadz_root=str(tmp_path / "PEAKHEADZ_ROOT"),
            obsidian_path="",
            sync_interval_seconds=60,
        )
        store.save(config)
        loaded = store.load()
        if loaded.slack_channel_id != "C123456":
            raise AssertionError("config restore failed")
        message = SlackMessage(ts="1700000000.000100", text="hello from slack")
        saved = save_slack_message(loaded.peakheadz_root, loaded.slack_channel_id, message)
        body = saved.read_text(encoding="utf-8")
        if "source: slack" not in body or "hello from slack" not in body:
            raise AssertionError("markdown save failed")
    print(UI_TEXT["self_test_ok"])
    return 0


def run_launch_check() -> int:
    ConfigStore().load()
    if not UI_TEXT["app_name"]:
        raise AssertionError("UI_TEXT is empty")
    print(UI_TEXT["launch_check_ok"])
    return 0


def main() -> int:
    if "--self-test" in sys.argv:
        return run_self_test()
    if "--launch-check" in sys.argv:
        return run_launch_check()
    root = Tk()
    NoteInboxApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
