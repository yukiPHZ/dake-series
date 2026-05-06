# -*- coding: utf-8 -*-
from __future__ import annotations

import ctypes
import os
import queue
import re
import sys
import threading
import uuid
import webbrowser
from dataclasses import dataclass
from pathlib import Path
from tkinter import font as tkfont
from tkinter import messagebox, ttk
import tkinter as tk

from PIL import Image, ImageGrab

try:
    import mss
except ImportError:
    mss = None


APP_NAME = "DakeScreen_WebP"
WINDOW_TITLE = "スクショ→WebP"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "main_title": "スクショしてWebP保存",
    "main_description": "ボタンを押して、撮りたいウインドウを選ぶだけでWebP保存します。",
    "center_title": "待機中",
    "center_shortcut": "ボタン：選んで保存 / Ctrl + Shift + 1：即保存",
    "button_execute": "選んでWebP保存",
    "status_label": "ステータス",
    "status_idle": "待機中",
    "status_select_window": "撮影するウインドウを選んでください",
    "status_select_countdown": "撮影するウインドウを選んでください：{count}",
    "status_saving": "保存中.",
    "status_saved": "保存しました：{filename}",
    "status_error": "保存できませんでした",
    "status_hotkey_error": "ショートカットを登録できませんでした",
    "save_location_label": "保存先：デスクトップ\\DAKE_screenshots",
    "dialog_error_title": "保存できませんでした",
    "dialog_capture_error_message": "{message}",
    "dialog_hotkey_error_title": "ショートカットを使えません",
    "dialog_hotkey_error_message": "Ctrl + Shift + 1 が他のアプリで使われている可能性があります。ボタンから保存してください。",
    "error_windows_only": "Windowsで実行してください。",
    "error_no_window": "撮影するウインドウを前面に出してから、もう一度実行してください。",
    "error_no_target": "撮影するウインドウを前面に出してから、もう一度実行してください。",
    "error_minimized": "撮影するウインドウが最小化されています。前面に出してから、もう一度実行してください。",
    "error_invalid_size": "撮影できる大きさのウインドウが見つかりませんでした。前面に出してから、もう一度実行してください。",
    "error_capture_failed": "ウインドウを撮影できませんでした。前面に出してから、もう一度実行してください。",
    "error_save_failed": "画像を保存できませんでした。保存先フォルダを確認してください。",
    "footer_series": "シンプルそれDAKEシリーズ",
    "footer_phrase": "止まらない、迷わない、すぐ終わる。",
    "footer_link_assessment": "戸建買取査定",
    "footer_link_instagram": "Instagram",
    "footer_separator": " ｜ ",
    "footer_copyright": COPYRIGHT,
}

THEME = {
    "background": "#F6F7F9",
    "card": "#FFFFFF",
    "text": "#1E2430",
    "muted": "#667085",
    "border": "#E6EAF0",
    "accent": "#2F6FED",
    "accent_hover": "#2458BF",
    "accent_disabled": "#A9C0F7",
    "success": "#12B76A",
    "success_bg": "#EAFBF3",
    "error": "#D92D20",
    "error_bg": "#FDECEC",
    "soft": "#EEF2F7",
}

LINK_URLS = {
    "assessment": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "instagram": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

STATUS_THEME = {
    "idle": (THEME["soft"], THEME["muted"]),
    "saving": ("#EAF2FF", THEME["accent"]),
    "success": (THEME["success_bg"], THEME["success"]),
    "error": (THEME["error_bg"], THEME["error"]),
}

WINDOW_SIZE = "760x520"
WINDOW_MIN_SIZE = (700, 500)
TARGET_WIDTH = 1200
WEBP_QUALITY = 88
OUTPUT_FOLDER_NAME = "DAKE_screenshots"
FILE_PATTERN = re.compile(r"^screenshot-(\d+)\.webp$", re.IGNORECASE)
BUTTON_SELECT_COUNTDOWN_SECONDS = 3
COUNTDOWN_INTERVAL_MS = 1000
IMMEDIATE_SELF_CAPTURE_DELAY_MS = 220
QUEUE_POLL_INTERVAL_MS = 80
HOTKEY_POLL_INTERVAL_MS = 80
FOREGROUND_POLL_INTERVAL_MS = 250
RESAMPLE = Image.Resampling.LANCZOS if hasattr(Image, "Resampling") else Image.LANCZOS

if os.name == "nt":
    from ctypes import wintypes
else:
    wintypes = None


class UserFacingError(RuntimeError):
    pass


@dataclass(frozen=True)
class WindowRect:
    left: int
    top: int
    right: int
    bottom: int

    @property
    def width(self) -> int:
        return self.right - self.left

    @property
    def height(self) -> int:
        return self.bottom - self.top

    def as_bbox(self) -> tuple[int, int, int, int]:
        return self.left, self.top, self.right, self.bottom


class RECT(ctypes.Structure):
    _fields_ = [
        ("left", ctypes.c_long),
        ("top", ctypes.c_long),
        ("right", ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]


class GUID(ctypes.Structure):
    _fields_ = [
        ("Data1", ctypes.c_ulong),
        ("Data2", ctypes.c_ushort),
        ("Data3", ctypes.c_ushort),
        ("Data4", ctypes.c_ubyte * 8),
    ]


def guid_from_string(value: str) -> GUID:
    parsed = uuid.UUID(value)
    data4 = (ctypes.c_ubyte * 8).from_buffer_copy(parsed.bytes[8:])
    return GUID(parsed.time_low, parsed.time_mid, parsed.time_hi_version, data4)


def enable_dpi_awareness() -> None:
    if os.name != "nt":
        return

    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        return
    except Exception:
        pass

    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def get_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def get_common_icon_candidates() -> list[Path]:
    source_dir = Path(__file__).resolve().parent
    base_dir = get_base_dir()
    return [
        source_dir / ".." / ".." / "02_assets" / "dake_icon.ico",
        base_dir / ".." / ".." / "02_assets" / "dake_icon.ico",
        base_dir / ".." / ".." / ".." / "02_assets" / "dake_icon.ico",
        base_dir / "dake_icon.ico",
    ]


def choose_font_family(root: tk.Tk) -> str:
    preferred = ["BIZ UDPGothic", "Yu Gothic UI", "Meiryo"]
    available = set(tkfont.families(root))
    for family in preferred:
        if family in available:
            return family
    return "TkDefaultFont"


def get_desktop_dir() -> Path:
    if os.name == "nt":
        desktop_guid = guid_from_string("B4BFCC3A-DB2C-424C-B029-7FE99A87C641")
        path_pointer = ctypes.c_wchar_p()
        try:
            result = ctypes.windll.shell32.SHGetKnownFolderPath(
                ctypes.byref(desktop_guid),
                0,
                None,
                ctypes.byref(path_pointer),
            )
            if result == 0 and path_pointer.value:
                return Path(path_pointer.value)
        except Exception:
            pass
        finally:
            if path_pointer:
                try:
                    ctypes.windll.ole32.CoTaskMemFree(path_pointer)
                except Exception:
                    pass

    return Path.home() / "Desktop"


def normalize_to_white_rgb(image: Image.Image) -> Image.Image:
    if image.mode == "RGB":
        return image.copy()

    if "A" in image.getbands():
        rgba_image = image.convert("RGBA")
        background = Image.new("RGB", rgba_image.size, (255, 255, 255))
        background.paste(rgba_image, mask=rgba_image.getchannel("A"))
        return background

    return image.convert("RGB")


def resize_for_output(image: Image.Image) -> Image.Image:
    if image.width <= TARGET_WIDTH:
        return image

    new_height = max(1, round(image.height * (TARGET_WIDTH / image.width)))
    return image.resize((TARGET_WIDTH, new_height), RESAMPLE)


class ActiveWindowCaptureService:
    DWMWA_EXTENDED_FRAME_BOUNDS = 9
    GA_ROOT = 2
    EXCLUDED_WINDOW_CLASSES = {"Shell_TrayWnd", "Shell_SecondaryTrayWnd", "Progman", "WorkerW"}

    def __init__(self) -> None:
        if os.name != "nt":
            raise UserFacingError(UI_TEXT["error_windows_only"])
        self.user32 = ctypes.windll.user32
        self.dwmapi = ctypes.windll.dwmapi

    def get_foreground_window(self) -> int | None:
        hwnd = self.user32.GetForegroundWindow()
        return int(hwnd) if hwnd else None

    def get_root_window(self, hwnd: int | None) -> int | None:
        if not hwnd:
            return None
        root = self.user32.GetAncestor(wintypes.HWND(hwnd), self.GA_ROOT)
        return int(root) if root else int(hwnd)

    def is_same_root(self, first_hwnd: int | None, second_hwnd: int | None) -> bool:
        if not first_hwnd or not second_hwnd:
            return False
        return self.get_root_window(first_hwnd) == self.get_root_window(second_hwnd)

    def is_window_usable(self, hwnd: int | None) -> bool:
        if not hwnd:
            return False
        if not self.user32.IsWindow(wintypes.HWND(hwnd)):
            return False
        if not self.user32.IsWindowVisible(wintypes.HWND(hwnd)):
            return False
        return True

    def is_capture_candidate(self, hwnd: int | None) -> bool:
        if not self.is_window_usable(hwnd):
            return False
        return self.get_window_class_name(hwnd) not in self.EXCLUDED_WINDOW_CLASSES

    def get_window_class_name(self, hwnd: int | None) -> str:
        if not hwnd:
            return ""
        buffer = ctypes.create_unicode_buffer(256)
        self.user32.GetClassNameW(wintypes.HWND(hwnd), buffer, len(buffer))
        return buffer.value

    def ensure_window_ready(self, hwnd: int | None) -> int:
        if not self.is_capture_candidate(hwnd):
            raise UserFacingError(UI_TEXT["error_no_window"])
        if self.user32.IsIconic(wintypes.HWND(hwnd)):
            raise UserFacingError(UI_TEXT["error_minimized"])
        return int(hwnd)

    def get_window_rect(self, hwnd: int) -> WindowRect:
        rect = RECT()
        result = self.dwmapi.DwmGetWindowAttribute(
            wintypes.HWND(hwnd),
            self.DWMWA_EXTENDED_FRAME_BOUNDS,
            ctypes.byref(rect),
            ctypes.sizeof(rect),
        )
        if result != 0:
            if not self.user32.GetWindowRect(wintypes.HWND(hwnd), ctypes.byref(rect)):
                raise UserFacingError(UI_TEXT["error_capture_failed"])

        window_rect = WindowRect(
            int(rect.left),
            int(rect.top),
            int(rect.right),
            int(rect.bottom),
        )
        if window_rect.width <= 0 or window_rect.height <= 0:
            raise UserFacingError(UI_TEXT["error_invalid_size"])
        return window_rect

    def capture_window(self, hwnd: int | None) -> Image.Image:
        ready_hwnd = self.ensure_window_ready(hwnd)
        rect = self.get_window_rect(ready_hwnd)

        try:
            image = self._grab_with_mss(rect)
        except Exception:
            image = self._grab_with_pillow(rect)

        return resize_for_output(normalize_to_white_rgb(image))

    def _grab_with_mss(self, rect: WindowRect) -> Image.Image:
        if mss is None:
            raise RuntimeError("mss is not available")

        monitor = {
            "left": rect.left,
            "top": rect.top,
            "width": rect.width,
            "height": rect.height,
        }
        with mss.mss() as screen_capture:
            shot = screen_capture.grab(monitor)
            return Image.frombytes("RGB", shot.size, shot.rgb)

    def _grab_with_pillow(self, rect: WindowRect) -> Image.Image:
        try:
            return ImageGrab.grab(bbox=rect.as_bbox(), all_screens=True)
        except TypeError:
            return ImageGrab.grab(bbox=rect.as_bbox())


class WebPFileService:
    def __init__(self) -> None:
        self.output_dir = get_desktop_dir() / OUTPUT_FOLDER_NAME

    def save(self, image: Image.Image) -> Path:
        try:
            self.output_dir.mkdir(parents=True, exist_ok=True)
            destination = self._build_next_path()
            image.save(destination, format="WEBP", quality=WEBP_QUALITY, method=6)
            return destination
        except Exception as exc:
            raise UserFacingError(UI_TEXT["error_save_failed"]) from exc

    def _build_next_path(self) -> Path:
        max_number = 0
        for path in self.output_dir.glob("screenshot-*.webp"):
            match = FILE_PATTERN.match(path.name)
            if match:
                max_number = max(max_number, int(match.group(1)))

        next_number = max_number + 1
        while True:
            candidate = self.output_dir / f"screenshot-{next_number:02d}.webp"
            if not candidate.exists():
                return candidate
            next_number += 1


class GlobalHotkeyListener:
    MOD_CONTROL = 0x0002
    MOD_SHIFT = 0x0004
    MOD_NOREPEAT = 0x4000
    VK_1 = 0x31
    WM_HOTKEY = 0x0312
    WM_QUIT = 0x0012

    def __init__(self, on_hotkey) -> None:
        self.on_hotkey = on_hotkey
        self.hotkey_id = 1201
        self.thread: threading.Thread | None = None
        self.thread_id: int | None = None
        self.error: str | None = None
        self._ready = threading.Event()

    def start(self) -> None:
        if os.name != "nt":
            self.error = UI_TEXT["error_windows_only"]
            return

        self.thread = threading.Thread(target=self._run, daemon=True)
        self.thread.start()

    def stop(self) -> None:
        if os.name != "nt" or self.thread_id is None:
            return
        try:
            ctypes.windll.user32.PostThreadMessageW(self.thread_id, self.WM_QUIT, 0, 0)
        except Exception:
            pass

    def _run(self) -> None:
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        self.thread_id = int(kernel32.GetCurrentThreadId())

        modifiers = self.MOD_CONTROL | self.MOD_SHIFT | self.MOD_NOREPEAT
        if not user32.RegisterHotKey(None, self.hotkey_id, modifiers, self.VK_1):
            self.error = UI_TEXT["status_hotkey_error"]
            self._ready.set()
            return

        self._ready.set()
        message = wintypes.MSG()
        try:
            while user32.GetMessageW(ctypes.byref(message), None, 0, 0) != 0:
                if message.message == self.WM_HOTKEY and message.wParam == self.hotkey_id:
                    self.on_hotkey()
        finally:
            user32.UnregisterHotKey(None, self.hotkey_id)


class ScreenWebPApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(*WINDOW_MIN_SIZE)
        self.root.configure(bg=THEME["background"])

        self.font_family = choose_font_family(self.root)
        self.root.option_add("*Font", (self.font_family, 10))

        self.capture_service = ActiveWindowCaptureService()
        self.file_service = WebPFileService()
        self.capture_queue: queue.Queue[tuple[str, str, Path | None]] = queue.Queue()
        self.hotkey_queue: queue.Queue[str] = queue.Queue()
        self.status_var = tk.StringVar(value=UI_TEXT["status_idle"])
        self.busy = False
        self.root_hidden_for_capture = False
        self.last_target_hwnd: int | None = None
        self.hotkey_error_reported = False
        self.footer_layout_stacked: bool | None = None

        self._apply_window_icon()
        self._build_styles()
        self._build_ui()
        self._set_status(UI_TEXT["status_idle"], "idle")

        self.hotkey_listener = GlobalHotkeyListener(lambda: self.hotkey_queue.put("capture"))
        self.hotkey_listener.start()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(FOREGROUND_POLL_INTERVAL_MS, self._poll_foreground_window)
        self.root.after(HOTKEY_POLL_INTERVAL_MS, self._poll_hotkey_queue)
        self.root.after(300, self._check_hotkey_registration)

    def run(self) -> None:
        self.root.mainloop()

    def _build_styles(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(
            "TSeparator",
            background=THEME["border"],
        )

    def _build_ui(self) -> None:
        container = tk.Frame(self.root, bg=THEME["background"])
        container.pack(fill="both", expand=True, padx=26, pady=22)

        header = tk.Frame(container, bg=THEME["background"])
        header.pack(fill="x")

        title = tk.Label(
            header,
            text=UI_TEXT["main_title"],
            bg=THEME["background"],
            fg=THEME["text"],
            font=(self.font_family, 20, "bold"),
            anchor="w",
        )
        title.pack(fill="x")

        description = tk.Label(
            header,
            text=UI_TEXT["main_description"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 10),
            anchor="w",
        )
        description.pack(fill="x", pady=(6, 0))

        main_card = tk.Frame(
            container,
            bg=THEME["card"],
            highlightbackground=THEME["border"],
            highlightthickness=1,
        )
        main_card.pack(fill="both", expand=True, pady=(18, 14))

        center = tk.Frame(main_card, bg=THEME["card"])
        center.pack(expand=True)

        waiting_label = tk.Label(
            center,
            text=UI_TEXT["center_title"],
            bg=THEME["card"],
            fg=THEME["text"],
            font=(self.font_family, 22, "bold"),
        )
        waiting_label.pack(pady=(0, 8))

        shortcut_label = tk.Label(
            center,
            text=UI_TEXT["center_shortcut"],
            bg=THEME["card"],
            fg=THEME["muted"],
            font=(self.font_family, 13, "bold"),
        )
        shortcut_label.pack(pady=(0, 22))

        self.capture_button = tk.Button(
            center,
            text=UI_TEXT["button_execute"],
            command=self.request_button_capture,
            bg=THEME["accent"],
            fg="#FFFFFF",
            activebackground=THEME["accent_hover"],
            activeforeground="#FFFFFF",
            disabledforeground="#FFFFFF",
            relief="flat",
            bd=0,
            cursor="hand2",
            padx=34,
            pady=14,
            font=(self.font_family, 12, "bold"),
        )
        self.capture_button.pack()

        status_row = tk.Frame(main_card, bg=THEME["card"])
        status_row.pack(fill="x", padx=18, pady=(0, 8))

        status_label = tk.Label(
            status_row,
            text=UI_TEXT["status_label"],
            bg=THEME["card"],
            fg=THEME["muted"],
            font=(self.font_family, 9),
        )
        status_label.pack(side="left", padx=(0, 8))

        self.status_badge = tk.Label(
            status_row,
            textvariable=self.status_var,
            bg=THEME["soft"],
            fg=THEME["muted"],
            font=(self.font_family, 9, "bold"),
            padx=10,
            pady=4,
        )
        self.status_badge.pack(side="left")

        save_location = tk.Label(
            main_card,
            text=UI_TEXT["save_location_label"],
            bg=THEME["card"],
            fg=THEME["muted"],
            font=(self.font_family, 9),
            anchor="w",
        )
        save_location.pack(fill="x", padx=18, pady=(0, 18))

        self._build_footer(container)

    def _build_footer(self, container: tk.Frame) -> None:
        self.footer = tk.Frame(container, bg=THEME["background"])
        self.footer.pack(fill="x", pady=(4, 0))
        self.footer.grid_columnconfigure(0, weight=1)

        self.footer_left = tk.Frame(self.footer, bg=THEME["background"])

        footer_series = tk.Label(
            self.footer_left,
            text=UI_TEXT["footer_series"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 8, "bold"),
            anchor="w",
        )
        footer_series.pack(anchor="w")

        footer_phrase = tk.Label(
            self.footer_left,
            text=UI_TEXT["footer_phrase"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 8),
            anchor="w",
        )
        footer_phrase.pack(anchor="w", pady=(2, 0))

        self.footer_right = tk.Frame(self.footer, bg=THEME["background"])

        self._make_footer_link(self.footer_right, UI_TEXT["footer_link_assessment"], LINK_URLS["assessment"])
        self._make_footer_text(self.footer_right, UI_TEXT["footer_separator"])
        self._make_footer_link(self.footer_right, UI_TEXT["footer_link_instagram"], LINK_URLS["instagram"])
        self._make_footer_text(self.footer_right, UI_TEXT["footer_separator"])
        self._make_footer_text(self.footer_right, UI_TEXT["footer_copyright"])
        self.footer.bind("<Configure>", lambda event: self._layout_footer(event.width))
        self._layout_footer(WINDOW_MIN_SIZE[0])

    def _layout_footer(self, width: int) -> None:
        should_stack = width < 720
        if self.footer_layout_stacked == should_stack:
            return

        self.footer_layout_stacked = should_stack
        self.footer_left.grid_forget()
        self.footer_right.grid_forget()

        if should_stack:
            self.footer.grid_columnconfigure(0, weight=1)
            self.footer.grid_columnconfigure(1, weight=0)
            self.footer_left.grid(row=0, column=0, sticky="", pady=(2, 4))
            self.footer_right.grid(row=1, column=0, sticky="", pady=(0, 2))
            return

        self.footer.grid_columnconfigure(0, weight=1)
        self.footer.grid_columnconfigure(1, weight=0)
        self.footer_left.grid(row=0, column=0, sticky="w", pady=(2, 2))
        self.footer_right.grid(row=0, column=1, sticky="e", pady=(2, 2))

    def _make_footer_link(self, parent: tk.Frame, label: str, url: str) -> None:
        link = tk.Label(
            parent,
            text=label,
            bg=THEME["background"],
            fg=THEME["accent"],
            font=(self.font_family, 8, "bold"),
            cursor="hand2",
        )
        link.pack(side="left")
        link.bind("<Button-1>", lambda _event: webbrowser.open(url, new=2))

    def _make_footer_text(self, parent: tk.Frame, label: str) -> None:
        text = tk.Label(
            parent,
            text=label,
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 8),
        )
        text.pack(side="left")

    def _apply_window_icon(self) -> None:
        for icon_path in get_common_icon_candidates():
            try:
                resolved_path = icon_path.resolve()
            except OSError:
                resolved_path = icon_path
            if resolved_path.exists():
                try:
                    self.root.iconbitmap(str(resolved_path))
                except tk.TclError:
                    pass
                return

    def _poll_foreground_window(self) -> None:
        try:
            hwnd = self.capture_service.get_foreground_window()
            if hwnd and not self._is_own_window(hwnd) and self.capture_service.is_capture_candidate(hwnd):
                self.last_target_hwnd = hwnd
        finally:
            self.root.after(FOREGROUND_POLL_INTERVAL_MS, self._poll_foreground_window)

    def _poll_hotkey_queue(self) -> None:
        while True:
            try:
                self.hotkey_queue.get_nowait()
            except queue.Empty:
                break
            self.request_immediate_capture()

        self.root.after(HOTKEY_POLL_INTERVAL_MS, self._poll_hotkey_queue)

    def _check_hotkey_registration(self) -> None:
        if self.hotkey_error_reported:
            return
        if self.hotkey_listener.error:
            self.hotkey_error_reported = True
            self._set_status(UI_TEXT["status_hotkey_error"], "error")
            messagebox.showwarning(
                UI_TEXT["dialog_hotkey_error_title"],
                UI_TEXT["dialog_hotkey_error_message"],
            )

    def request_button_capture(self) -> None:
        if self.busy:
            return

        self._set_busy(True)
        self._clear_capture_queue()
        self._set_status(UI_TEXT["status_select_window"], "saving")
        self.root.update_idletasks()
        self.root_hidden_for_capture = True
        self.root.iconify()
        self.root.after(100, lambda: self._run_select_countdown(BUTTON_SELECT_COUNTDOWN_SECONDS))

    def _run_select_countdown(self, count: int) -> None:
        if not self.busy:
            return

        if count > 0:
            self._set_status(UI_TEXT["status_select_countdown"].format(count=count), "saving")
            self.root.after(COUNTDOWN_INTERVAL_MS, lambda: self._run_select_countdown(count - 1))
            return

        self._set_status(UI_TEXT["status_saving"], "saving")
        target_hwnd = self.capture_service.get_foreground_window()
        if target_hwnd and self._is_own_window(target_hwnd):
            target_hwnd = None
        self._start_capture_worker(target_hwnd)

    def request_immediate_capture(self) -> None:
        if self.busy:
            return

        self._set_busy(True)
        self._set_status(UI_TEXT["status_saving"], "saving")
        self._clear_capture_queue()

        foreground_hwnd = self.capture_service.get_foreground_window()
        if foreground_hwnd and not self._is_own_window(foreground_hwnd):
            self._start_capture_worker(foreground_hwnd)
            return

        target_hwnd = self.last_target_hwnd if self.capture_service.is_capture_candidate(self.last_target_hwnd) else None
        self.root_hidden_for_capture = True
        self.root.withdraw()
        self.root.after(IMMEDIATE_SELF_CAPTURE_DELAY_MS, lambda: self._capture_after_hiding(target_hwnd))

    def _capture_after_hiding(self, target_hwnd: int | None) -> None:
        if target_hwnd is None:
            foreground_hwnd = self.capture_service.get_foreground_window()
            if foreground_hwnd and not self._is_own_window(foreground_hwnd):
                target_hwnd = foreground_hwnd
        self._start_capture_worker(target_hwnd)

    def _start_capture_worker(self, hwnd: int | None) -> None:
        worker = threading.Thread(target=self._capture_worker, args=(hwnd,), daemon=True)
        worker.start()
        self.root.after(QUEUE_POLL_INTERVAL_MS, self._poll_capture_queue)

    def _capture_worker(self, hwnd: int | None) -> None:
        try:
            image = self.capture_service.capture_window(hwnd)
            destination = self.file_service.save(image)
            self.capture_queue.put(("complete", destination.name, destination))
        except UserFacingError as exc:
            self.capture_queue.put(("error", str(exc), None))
        except Exception:
            self.capture_queue.put(("error", UI_TEXT["error_capture_failed"], None))

    def _poll_capture_queue(self) -> None:
        try:
            result, message, destination = self.capture_queue.get_nowait()
        except queue.Empty:
            if self.busy:
                self.root.after(QUEUE_POLL_INTERVAL_MS, self._poll_capture_queue)
            return

        self._restore_root_after_capture()
        self._set_busy(False)

        if result == "complete" and destination is not None:
            self._set_status(UI_TEXT["status_saved"].format(filename=destination.name), "success")
            self.last_target_hwnd = None
            return

        self._set_status(UI_TEXT["status_error"], "error")
        messagebox.showerror(
            UI_TEXT["dialog_error_title"],
            UI_TEXT["dialog_capture_error_message"].format(message=message),
        )

    def _restore_root_after_capture(self) -> None:
        if not self.root_hidden_for_capture:
            return
        self.root_hidden_for_capture = False
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

    def _clear_capture_queue(self) -> None:
        while True:
            try:
                self.capture_queue.get_nowait()
            except queue.Empty:
                return

    def _set_busy(self, busy: bool) -> None:
        self.busy = busy
        state = tk.DISABLED if busy else tk.NORMAL
        self.capture_button.configure(state=state)
        if busy:
            self.capture_button.configure(bg=THEME["accent_disabled"])
        else:
            self.capture_button.configure(bg=THEME["accent"])

    def _set_status(self, text: str, status_type: str) -> None:
        self.status_var.set(text)
        background, foreground = STATUS_THEME.get(status_type, STATUS_THEME["idle"])
        self.status_badge.configure(bg=background, fg=foreground)

    def _is_own_window(self, hwnd: int | None) -> bool:
        if not hwnd:
            return False
        try:
            own_hwnd = int(self.root.winfo_id())
        except tk.TclError:
            return False
        return self.capture_service.is_same_root(hwnd, own_hwnd)

    def _on_close(self) -> None:
        self.hotkey_listener.stop()
        self.root.destroy()


def main() -> None:
    enable_dpi_awareness()
    app = ScreenWebPApp()
    app.run()


if __name__ == "__main__":
    main()
