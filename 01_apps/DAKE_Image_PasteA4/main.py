# -*- coding: utf-8 -*-
from __future__ import annotations

import ctypes
import os
import queue
import subprocess
import sys
import threading
import time
import webbrowser
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, font as tkfont, messagebox, ttk
import tkinter as tk
from ctypes import wintypes

from PIL import Image, ImageGrab, ImageTk

try:
    import mss
except Exception:
    mss = None


APP_NAME = "Dake貼る"
WINDOW_TITLE = "Dake貼る"
COPYRIGHT = "© 2026 しまりす不動産 / Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "brand_series": "シンプルそれDAKEシリーズ",
    "main_title": "ウインドウを貼る",
    "main_description": "見えているウインドウを貼って、A4画像にします。",
    "button_refresh_windows": "更新",
    "button_paste_window": "貼る",
    "button_portrait": "A4縦",
    "button_landscape": "A4横",
    "button_smaller": "小さく",
    "button_larger": "大きく",
    "button_delete": "削除",
    "button_clear": "クリア",
    "button_export": "画像出力",
    "empty_title": "まだ貼られていません",
    "empty_subtitle": "左の一覧からウインドウを選んで貼ってください。",
    "status_idle": "待機中",
    "status_loading_windows": "一覧を更新中",
    "status_ready": "準備完了",
    "status_capturing": "貼り付け中",
    "status_exporting": "画像を出力中",
    "status_complete": "完了しました",
    "status_error": "処理できませんでした",
    "footer_left": "シンプルそれDAKEシリーズ / 止まらない、迷わない、すぐ終わる。",
    "footer_link_1": "戸建買取査定",
    "footer_link_2": "Instagram",
    "footer_separator": " ｜ ",
    "footer_copyright": COPYRIGHT,
    "window_list_title": "起動中のウインドウ",
    "window_list_hint": "タイトルがある、表示中のウインドウだけを表示します。",
    "selection_label": "選択画像",
    "status_label": "状態",
    "dialog_error_title": "エラー",
    "dialog_no_window_title": "ウインドウを選んでください",
    "dialog_no_window_message": "一覧から貼りたいウインドウを選んでください。",
    "dialog_capture_error_message": "このウインドウを取得できませんでした。\n前面に表示してから、もう一度お試しください。\n\n{error}",
    "dialog_export_error_message": "画像を保存できませんでした。\n保存先を確認してから、もう一度お試しください。\n\n{error}",
    "dialog_export_title": "PNG画像として保存",
    "filetype_png": "PNG画像",
    "export_file_template": "Dake貼る_{timestamp}.png",
    "status_windows_count": "{count}件のウインドウを表示しています",
    "status_pasted": "貼りました",
    "status_deleted": "削除しました",
    "status_cleared": "クリアしました",
    "status_export_complete": "保存しました: {name}",
    "error_windows_only": "Windowsで実行してください。",
    "error_no_windows": "貼れるウインドウが見つかりませんでした。",
    "error_minimized": "最小化されているため取得できません。",
    "error_invalid_size": "取得できる大きさではありません。",
    "error_capture_failed": "ウインドウを取得できませんでした。",
}

THEME = {
    "background": "#F6F7F9",
    "card": "#FFFFFF",
    "paper": "#FFFFFF",
    "text": "#1E2430",
    "muted": "#667085",
    "border": "#E6EAF0",
    "accent": "#2F6FED",
    "accent_hover": "#2458BF",
    "accent_soft": "#EAF2FF",
    "soft": "#EEF2F7",
    "success": "#12B76A",
    "success_bg": "#EAFBF3",
    "error": "#D92D20",
    "error_bg": "#FDECEC",
    "shadow": "#D8DEE8",
    "canvas_bg": "#EEF1F5",
    "link": "#58677D",
    "link_hover": "#2F6FED",
}

LINK_URLS = {
    "footer_link_1": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "footer_link_2": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

STATUS_THEME = {
    "idle": (THEME["soft"], THEME["muted"]),
    "working": (THEME["accent_soft"], THEME["accent"]),
    "success": (THEME["success_bg"], THEME["success"]),
    "error": (THEME["error_bg"], THEME["error"]),
}

WINDOW_SIZE = "1180x780"
WINDOW_MIN_SIZE = (980, 660)
A4_PORTRAIT = (2480, 3508)
A4_LANDSCAPE = (3508, 2480)
PREVIEW_PADDING = 26
QUEUE_POLL_INTERVAL_MS = 80
CAPTURE_DELAY_MS = 220
RESIZE_STEP = 1.12
MIN_ITEM_SIZE = 80
MAX_ITEM_PAGE_RATIO = 0.98
RESAMPLE = Image.Resampling.LANCZOS if hasattr(Image, "Resampling") else Image.LANCZOS


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


@dataclass(frozen=True)
class WindowRecord:
    hwnd: int
    title: str
    rect: WindowRect

    @property
    def display_name(self) -> str:
        return f"{self.title}  ({self.rect.width}x{self.rect.height})"


@dataclass
class PlacedImage:
    item_id: int
    image: Image.Image
    x: float
    y: float
    width: float
    height: float
    photo: ImageTk.PhotoImage | None = field(default=None, repr=False)


class RECT(ctypes.Structure):
    _fields_ = [
        ("left", ctypes.c_long),
        ("top", ctypes.c_long),
        ("right", ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]


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


def normalize_image(image: Image.Image) -> Image.Image:
    if image.mode == "RGB":
        return image.copy()

    if "A" in image.getbands() or "transparency" in image.info:
        rgba_image = image.convert("RGBA")
        background = Image.new("RGB", rgba_image.size, "white")
        background.paste(rgba_image, mask=rgba_image.getchannel("A"))
        return background

    return image.convert("RGB")


def fit_size(source_size: tuple[int, int], box_size: tuple[int, int]) -> tuple[int, int]:
    source_width, source_height = source_size
    box_width, box_height = box_size
    if source_width <= 0 or source_height <= 0:
        return 1, 1

    scale = min(box_width / source_width, box_height / source_height, 1.0)
    return max(1, round(source_width * scale)), max(1, round(source_height * scale))


def default_output_dir() -> Path:
    downloads = Path.home() / "Downloads"
    if downloads.exists():
        return downloads
    return Path.home()


def open_folder(path: Path) -> None:
    try:
        if os.name == "nt":
            os.startfile(str(path))  # type: ignore[attr-defined]
            return
        if sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
            return
        subprocess.run(["xdg-open", str(path)], check=False)
    except Exception:
        pass


class WindowCaptureService:
    DWMWA_EXTENDED_FRAME_BOUNDS = 9
    DWMWA_CLOAKED = 14
    GA_ROOT = 2
    GWL_EXSTYLE = -20
    WS_EX_TOOLWINDOW = 0x00000080
    SW_RESTORE = 9
    EXCLUDED_CLASSES = {
        "Shell_TrayWnd",
        "Shell_SecondaryTrayWnd",
        "Progman",
        "WorkerW",
        "Windows.UI.Core.CoreWindow",
    }

    def __init__(self) -> None:
        if os.name != "nt":
            raise UserFacingError(UI_TEXT["error_windows_only"])

        self.user32 = ctypes.windll.user32
        self.dwmapi = ctypes.windll.dwmapi
        self.enum_windows_proc_type = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

        self.user32.EnumWindows.argtypes = [self.enum_windows_proc_type, wintypes.LPARAM]
        self.user32.EnumWindows.restype = wintypes.BOOL
        self.user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
        self.user32.GetWindowTextLengthW.restype = ctypes.c_int
        self.user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
        self.user32.GetWindowTextW.restype = ctypes.c_int
        self.user32.IsWindow.argtypes = [wintypes.HWND]
        self.user32.IsWindow.restype = wintypes.BOOL
        self.user32.IsWindowVisible.argtypes = [wintypes.HWND]
        self.user32.IsWindowVisible.restype = wintypes.BOOL
        self.user32.IsIconic.argtypes = [wintypes.HWND]
        self.user32.IsIconic.restype = wintypes.BOOL
        self.user32.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
        self.user32.GetAncestor.restype = wintypes.HWND
        self.user32.GetWindowRect.argtypes = [wintypes.HWND, ctypes.POINTER(RECT)]
        self.user32.GetWindowRect.restype = wintypes.BOOL
        self.user32.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
        self.user32.GetClassNameW.restype = ctypes.c_int
        self.user32.GetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int]
        self.user32.GetWindowLongPtrW.restype = ctypes.c_ssize_t
        self.user32.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
        self.user32.SetForegroundWindow.argtypes = [wintypes.HWND]

    def enumerate_windows(self, own_hwnd: int | None) -> list[WindowRecord]:
        records: list[WindowRecord] = []
        own_root = self.get_root_window(own_hwnd) if own_hwnd else None

        def callback(hwnd, _lparam) -> bool:
            hwnd_int = int(hwnd)
            try:
                record = self._build_window_record(hwnd_int, own_root)
            except Exception:
                record = None
            if record is not None:
                records.append(record)
            return True

        enum_callback = self.enum_windows_proc_type(callback)
        self.user32.EnumWindows(enum_callback, 0)
        records.sort(key=lambda item: item.title.casefold())
        return records

    def _build_window_record(self, hwnd: int, own_root: int | None) -> WindowRecord | None:
        if not self.is_window_usable(hwnd):
            return None
        if own_root is not None and self.get_root_window(hwnd) == own_root:
            return None
        if self.user32.IsIconic(wintypes.HWND(hwnd)):
            return None
        if self.is_cloaked(hwnd):
            return None
        if self.get_window_class_name(hwnd) in self.EXCLUDED_CLASSES:
            return None
        if self.is_tool_window(hwnd):
            return None

        title = self.get_window_title(hwnd).strip()
        if not title:
            return None
        if title == WINDOW_TITLE:
            return None

        rect = self.get_window_rect(hwnd)
        if rect.width < 80 or rect.height < 60:
            return None
        return WindowRecord(hwnd=hwnd, title=title, rect=rect)

    def is_window_usable(self, hwnd: int | None) -> bool:
        if not hwnd:
            return False
        handle = wintypes.HWND(hwnd)
        return bool(self.user32.IsWindow(handle)) and bool(self.user32.IsWindowVisible(handle))

    def is_same_root(self, first_hwnd: int | None, second_hwnd: int | None) -> bool:
        if not first_hwnd or not second_hwnd:
            return False
        return self.get_root_window(first_hwnd) == self.get_root_window(second_hwnd)

    def get_root_window(self, hwnd: int | None) -> int | None:
        if not hwnd:
            return None
        root = self.user32.GetAncestor(wintypes.HWND(hwnd), self.GA_ROOT)
        return int(root) if root else int(hwnd)

    def get_window_title(self, hwnd: int) -> str:
        length = self.user32.GetWindowTextLengthW(wintypes.HWND(hwnd))
        if length <= 0:
            return ""
        buffer = ctypes.create_unicode_buffer(length + 1)
        self.user32.GetWindowTextW(wintypes.HWND(hwnd), buffer, length + 1)
        return buffer.value

    def get_window_class_name(self, hwnd: int) -> str:
        buffer = ctypes.create_unicode_buffer(256)
        self.user32.GetClassNameW(wintypes.HWND(hwnd), buffer, len(buffer))
        return buffer.value

    def is_tool_window(self, hwnd: int) -> bool:
        try:
            ex_style = int(self.user32.GetWindowLongPtrW(wintypes.HWND(hwnd), self.GWL_EXSTYLE))
            return bool(ex_style & self.WS_EX_TOOLWINDOW)
        except Exception:
            return False

    def is_cloaked(self, hwnd: int) -> bool:
        cloaked = ctypes.c_int(0)
        try:
            result = self.dwmapi.DwmGetWindowAttribute(
                wintypes.HWND(hwnd),
                self.DWMWA_CLOAKED,
                ctypes.byref(cloaked),
                ctypes.sizeof(cloaked),
            )
            return result == 0 and cloaked.value != 0
        except Exception:
            return False

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
            left=int(rect.left),
            top=int(rect.top),
            right=int(rect.right),
            bottom=int(rect.bottom),
        )
        if window_rect.width <= 0 or window_rect.height <= 0:
            raise UserFacingError(UI_TEXT["error_invalid_size"])
        return window_rect

    def capture_window(self, hwnd: int) -> Image.Image:
        if not self.is_window_usable(hwnd):
            raise UserFacingError(UI_TEXT["error_capture_failed"])
        if self.user32.IsIconic(wintypes.HWND(hwnd)):
            raise UserFacingError(UI_TEXT["error_minimized"])

        try:
            self.user32.ShowWindow(wintypes.HWND(hwnd), self.SW_RESTORE)
            self.user32.SetForegroundWindow(wintypes.HWND(hwnd))
            time.sleep(0.12)
        except Exception:
            pass

        rect = self.get_window_rect(hwnd)
        if rect.width <= 0 or rect.height <= 0:
            raise UserFacingError(UI_TEXT["error_invalid_size"])

        try:
            image = self._grab_with_mss(rect)
        except Exception:
            image = self._grab_with_pillow(rect)
        return normalize_image(image)

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


class DakePasteA4App:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(*WINDOW_MIN_SIZE)
        self.root.configure(bg=THEME["background"])

        self.font_family = choose_font_family(self.root)
        self.root.option_add("*Font", (self.font_family, 10))

        self.capture_service: WindowCaptureService | None = None
        if os.name == "nt":
            self.capture_service = WindowCaptureService()

        self.window_records: list[WindowRecord] = []
        self.placed_images: list[PlacedImage] = []
        self.next_item_id = 1
        self.selected_item_id: int | None = None
        self.orientation = "portrait"
        self.busy = False
        self.capture_queue: queue.Queue[tuple[str, Image.Image | None, str | None]] = queue.Queue()
        self.export_queue: queue.Queue[tuple[str, Path | None, str | None]] = queue.Queue()
        self.root_hidden_for_capture = False
        self.window_queue: queue.Queue[tuple[str, list[WindowRecord] | None, str | None]] = queue.Queue()
        self.footer_stacked: bool | None = None
        self.paper_rect: tuple[int, int, int, int] | None = None
        self.paper_scale: tuple[float, float] | None = None
        self.drag_state: tuple[int, int, float, float] | None = None

        self.status_var = tk.StringVar(value=UI_TEXT["status_idle"])

        self._apply_window_icon()
        self._build_styles()
        self._build_ui()
        self._set_status(UI_TEXT["status_idle"], "idle")
        self._update_buttons()

        self.root.bind("<Delete>", self.delete_selected, add="+")
        self.root.after(200, self.refresh_windows)

    def run(self) -> None:
        self.root.mainloop()

    def _build_styles(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(
            "Primary.TButton",
            font=(self.font_family, 10, "bold"),
            padding=(18, 10),
            background=THEME["accent"],
            foreground="#FFFFFF",
            bordercolor=THEME["accent"],
            lightcolor=THEME["accent"],
            darkcolor=THEME["accent"],
        )
        style.map(
            "Primary.TButton",
            background=[("active", THEME["accent_hover"]), ("disabled", "#A9C0F7")],
            foreground=[("disabled", "#FFFFFF")],
        )

        style.configure(
            "Secondary.TButton",
            font=(self.font_family, 10, "bold"),
            padding=(14, 9),
            background=THEME["card"],
            foreground=THEME["text"],
            bordercolor=THEME["border"],
            lightcolor=THEME["card"],
            darkcolor=THEME["card"],
        )
        style.map("Secondary.TButton", background=[("active", THEME["soft"])])

    def _build_ui(self) -> None:
        container = tk.Frame(self.root, bg=THEME["background"])
        container.pack(fill="both", expand=True, padx=24, pady=(20, 16))

        self._build_header(container)
        self._build_body(container)
        self._build_control_bar(container)
        self._build_footer(container)

    def _build_header(self, parent: tk.Frame) -> None:
        header = tk.Frame(parent, bg=THEME["background"])
        header.pack(fill="x")
        header.grid_columnconfigure(0, weight=1)

        title_area = tk.Frame(header, bg=THEME["background"])
        title_area.grid(row=0, column=0, sticky="ew")

        title = tk.Label(
            title_area,
            text=UI_TEXT["main_title"],
            bg=THEME["background"],
            fg=THEME["text"],
            font=(self.font_family, 22, "bold"),
            anchor="w",
        )
        title.pack(fill="x")

        description = tk.Label(
            title_area,
            text=UI_TEXT["main_description"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 10),
            anchor="w",
        )
        description.pack(fill="x", pady=(6, 0))

        orientation_area = tk.Frame(header, bg=THEME["background"])
        orientation_area.grid(row=0, column=1, sticky="e", padx=(18, 0))

        self.portrait_button = tk.Button(
            orientation_area,
            text=UI_TEXT["button_portrait"],
            command=lambda: self.set_orientation("portrait"),
            relief="flat",
            bd=0,
            padx=16,
            pady=9,
            cursor="hand2",
            font=(self.font_family, 10, "bold"),
        )
        self.portrait_button.pack(side="left")

        self.landscape_button = tk.Button(
            orientation_area,
            text=UI_TEXT["button_landscape"],
            command=lambda: self.set_orientation("landscape"),
            relief="flat",
            bd=0,
            padx=16,
            pady=9,
            cursor="hand2",
            font=(self.font_family, 10, "bold"),
        )
        self.landscape_button.pack(side="left", padx=(8, 0))

    def _build_body(self, parent: tk.Frame) -> None:
        body = tk.Frame(parent, bg=THEME["background"])
        body.pack(fill="both", expand=True, pady=(18, 14))
        body.grid_columnconfigure(0, weight=0)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)

        self._build_window_panel(body)
        self._build_canvas_panel(body)

    def _build_window_panel(self, parent: tk.Frame) -> None:
        panel = tk.Frame(
            parent,
            bg=THEME["card"],
            highlightbackground=THEME["border"],
            highlightthickness=1,
        )
        panel.grid(row=0, column=0, sticky="nsw", padx=(0, 16))
        panel.configure(width=300)
        panel.grid_propagate(False)
        panel.grid_rowconfigure(2, weight=1)

        title = tk.Label(
            panel,
            text=UI_TEXT["window_list_title"],
            bg=THEME["card"],
            fg=THEME["text"],
            font=(self.font_family, 12, "bold"),
            anchor="w",
        )
        title.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 4))

        hint = tk.Label(
            panel,
            text=UI_TEXT["window_list_hint"],
            bg=THEME["card"],
            fg=THEME["muted"],
            font=(self.font_family, 9),
            anchor="w",
            justify="left",
            wraplength=250,
        )
        hint.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 12))

        list_frame = tk.Frame(panel, bg=THEME["card"])
        list_frame.grid(row=2, column=0, sticky="nsew", padx=16)
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_rowconfigure(0, weight=1)

        self.window_listbox = tk.Listbox(
            list_frame,
            activestyle="none",
            bg="#FFFFFF",
            fg=THEME["text"],
            selectbackground=THEME["accent_soft"],
            selectforeground=THEME["text"],
            highlightthickness=1,
            highlightbackground=THEME["border"],
            highlightcolor=THEME["accent"],
            relief="flat",
            exportselection=False,
            font=(self.font_family, 10),
        )
        self.window_listbox.grid(row=0, column=0, sticky="nsew")
        self.window_listbox.bind("<<ListboxSelect>>", lambda _event: self._update_buttons())

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.window_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.window_listbox.configure(yscrollcommand=scrollbar.set)

        actions = tk.Frame(panel, bg=THEME["card"])
        actions.grid(row=3, column=0, sticky="ew", padx=16, pady=16)

        self.refresh_button = ttk.Button(
            actions,
            text=UI_TEXT["button_refresh_windows"],
            style="Secondary.TButton",
            command=self.refresh_windows,
        )
        self.refresh_button.pack(side="left")

        self.paste_button = ttk.Button(
            actions,
            text=UI_TEXT["button_paste_window"],
            style="Primary.TButton",
            command=self.paste_selected_window,
        )
        self.paste_button.pack(side="right")

    def _build_canvas_panel(self, parent: tk.Frame) -> None:
        panel = tk.Frame(
            parent,
            bg=THEME["card"],
            highlightbackground=THEME["border"],
            highlightthickness=1,
        )
        panel.grid(row=0, column=1, sticky="nsew")
        panel.grid_rowconfigure(0, weight=1)
        panel.grid_columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(
            panel,
            bg=THEME["canvas_bg"],
            bd=0,
            highlightthickness=0,
        )
        self.canvas.grid(row=0, column=0, sticky="nsew", padx=16, pady=16)
        self.canvas.bind("<Configure>", lambda _event: self.draw_canvas())
        self.canvas.bind("<Button-1>", self._on_canvas_press)
        self.canvas.bind("<B1-Motion>", self._on_canvas_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_canvas_release)
        self.canvas.bind("<Delete>", self.delete_selected)

    def _build_control_bar(self, parent: tk.Frame) -> None:
        bar = tk.Frame(parent, bg=THEME["background"])
        bar.pack(fill="x")
        bar.grid_columnconfigure(1, weight=1)

        selection = tk.Frame(bar, bg=THEME["background"])
        selection.grid(row=0, column=0, sticky="w")

        label = tk.Label(
            selection,
            text=UI_TEXT["selection_label"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 9, "bold"),
        )
        label.pack(side="left", padx=(0, 10))

        self.smaller_button = ttk.Button(
            selection,
            text=UI_TEXT["button_smaller"],
            style="Secondary.TButton",
            command=lambda: self.resize_selected(1 / RESIZE_STEP),
        )
        self.smaller_button.pack(side="left", padx=(0, 8))

        self.larger_button = ttk.Button(
            selection,
            text=UI_TEXT["button_larger"],
            style="Secondary.TButton",
            command=lambda: self.resize_selected(RESIZE_STEP),
        )
        self.larger_button.pack(side="left", padx=(0, 8))

        self.delete_button = ttk.Button(
            selection,
            text=UI_TEXT["button_delete"],
            style="Secondary.TButton",
            command=self.delete_selected,
        )
        self.delete_button.pack(side="left")

        status = tk.Frame(bar, bg=THEME["background"])
        status.grid(row=0, column=1, sticky="ew", padx=18)

        status_label = tk.Label(
            status,
            text=UI_TEXT["status_label"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 9),
        )
        status_label.pack(side="left", padx=(0, 8))

        self.status_badge = tk.Label(
            status,
            textvariable=self.status_var,
            bg=THEME["soft"],
            fg=THEME["muted"],
            font=(self.font_family, 9, "bold"),
            padx=10,
            pady=5,
        )
        self.status_badge.pack(side="left")

        actions = tk.Frame(bar, bg=THEME["background"])
        actions.grid(row=0, column=2, sticky="e")

        self.clear_button = ttk.Button(
            actions,
            text=UI_TEXT["button_clear"],
            style="Secondary.TButton",
            command=self.clear_all,
        )
        self.clear_button.pack(side="left", padx=(0, 8))

        self.export_button = ttk.Button(
            actions,
            text=UI_TEXT["button_export"],
            style="Primary.TButton",
            command=self.export_png,
        )
        self.export_button.pack(side="left")

    def _build_footer(self, parent: tk.Frame) -> None:
        self.footer = tk.Frame(parent, bg=THEME["background"])
        self.footer.pack(fill="x", pady=(12, 0))
        self.footer.grid_columnconfigure(0, weight=1)

        self.footer_left = tk.Frame(self.footer, bg=THEME["background"])
        footer_left_label = tk.Label(
            self.footer_left,
            text=UI_TEXT["footer_left"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 8),
            anchor="w",
        )
        footer_left_label.pack(anchor="w")

        self.footer_right = tk.Frame(self.footer, bg=THEME["background"])
        self._make_footer_link(self.footer_right, UI_TEXT["footer_link_1"], LINK_URLS["footer_link_1"])
        self._make_footer_text(self.footer_right, UI_TEXT["footer_separator"])
        self._make_footer_link(self.footer_right, UI_TEXT["footer_link_2"], LINK_URLS["footer_link_2"])
        self._make_footer_text(self.footer_right, UI_TEXT["footer_separator"])
        self._make_footer_text(self.footer_right, UI_TEXT["footer_copyright"])

        self.footer.bind("<Configure>", lambda event: self._layout_footer(event.width))
        self._layout_footer(WINDOW_MIN_SIZE[0])

    def _make_footer_link(self, parent: tk.Frame, label: str, url: str) -> None:
        link = tk.Label(
            parent,
            text=label,
            bg=THEME["background"],
            fg=THEME["link"],
            font=(self.font_family, 8, "bold"),
            cursor="hand2",
        )
        link.pack(side="left")
        link.bind("<Button-1>", lambda _event: webbrowser.open(url, new=2))
        link.bind("<Enter>", lambda _event: link.configure(fg=THEME["link_hover"]))
        link.bind("<Leave>", lambda _event: link.configure(fg=THEME["link"]))

    def _make_footer_text(self, parent: tk.Frame, label: str) -> None:
        text = tk.Label(
            parent,
            text=label,
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 8),
        )
        text.pack(side="left")

    def _layout_footer(self, width: int) -> None:
        stacked = width < 860
        if stacked == self.footer_stacked:
            return

        self.footer_stacked = stacked
        self.footer_left.grid_forget()
        self.footer_right.grid_forget()

        if stacked:
            self.footer_left.grid(row=0, column=0, sticky="", pady=(0, 3))
            self.footer_right.grid(row=1, column=0, sticky="")
            return

        self.footer_left.grid(row=0, column=0, sticky="w")
        self.footer_right.grid(row=0, column=1, sticky="e")

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

    def refresh_windows(self) -> None:
        if self.busy:
            return
        if self.capture_service is None:
            self._set_status(UI_TEXT["error_windows_only"], "error")
            return

        self.busy = True
        self._set_status(UI_TEXT["status_loading_windows"], "working")
        self._update_buttons()

        own_hwnd = int(self.root.winfo_id())
        worker = threading.Thread(target=self._refresh_windows_worker, args=(own_hwnd,), daemon=True)
        worker.start()
        self.root.after(QUEUE_POLL_INTERVAL_MS, self._poll_refresh_queue)

    def _refresh_windows_worker(self, own_hwnd: int) -> None:
        try:
            records = self.capture_service.enumerate_windows(own_hwnd)
            self.window_queue.put(("windows", records, None))
        except Exception as exc:
            self.window_queue.put(("refresh_error", None, str(exc)))

    def _poll_refresh_queue(self) -> None:
        try:
            kind, records, message = self.window_queue.get_nowait()
        except queue.Empty:
            if self.busy:
                self.root.after(QUEUE_POLL_INTERVAL_MS, self._poll_refresh_queue)
            return

        self.busy = False
        if kind == "windows" and records is not None:
            self.window_records = records
            self.window_listbox.delete(0, tk.END)
            for record in self.window_records:
                self.window_listbox.insert(tk.END, record.display_name)
            if self.window_records:
                self.window_listbox.selection_set(0)
                self._set_status(UI_TEXT["status_windows_count"].format(count=len(self.window_records)), "success")
            else:
                self._set_status(UI_TEXT["error_no_windows"], "error")
        else:
            self._set_status(UI_TEXT["status_error"], "error")
            messagebox.showerror(
                UI_TEXT["dialog_error_title"],
                UI_TEXT["dialog_capture_error_message"].format(error=message or ""),
            )
        self._update_buttons()

    def paste_selected_window(self) -> None:
        if self.busy:
            return

        record = self._get_selected_window_record()
        if record is None:
            self._set_status(UI_TEXT["status_error"], "error")
            messagebox.showinfo(UI_TEXT["dialog_no_window_title"], UI_TEXT["dialog_no_window_message"])
            return

        self.busy = True
        self._set_status(UI_TEXT["status_capturing"], "working")
        self._update_buttons()
        self._clear_capture_queue()
        self.root_hidden_for_capture = True
        self.root.withdraw()
        self.root.after(CAPTURE_DELAY_MS, lambda: self._start_capture_worker(record.hwnd))

    def _start_capture_worker(self, hwnd: int) -> None:
        worker = threading.Thread(target=self._capture_worker, args=(hwnd,), daemon=True)
        worker.start()
        self.root.after(QUEUE_POLL_INTERVAL_MS, self._poll_capture_queue)

    def _capture_worker(self, hwnd: int) -> None:
        try:
            if self.capture_service is None:
                raise UserFacingError(UI_TEXT["error_windows_only"])
            image = self.capture_service.capture_window(hwnd)
            self.capture_queue.put(("capture_complete", image, None))
        except UserFacingError as exc:
            self.capture_queue.put(("capture_error", None, str(exc)))
        except Exception as exc:
            self.capture_queue.put(("capture_error", None, str(exc) or UI_TEXT["error_capture_failed"]))

    def _poll_capture_queue(self) -> None:
        try:
            kind, image, message = self.capture_queue.get_nowait()
        except queue.Empty:
            if self.busy:
                self.root.after(QUEUE_POLL_INTERVAL_MS, self._poll_capture_queue)
            return

        self._restore_root_after_capture()
        self.busy = False

        if kind == "capture_complete" and image is not None:
            self.add_image_to_canvas(image)
            self._set_status(UI_TEXT["status_pasted"], "success")
        else:
            self._set_status(UI_TEXT["status_error"], "error")
            messagebox.showerror(
                UI_TEXT["dialog_error_title"],
                UI_TEXT["dialog_capture_error_message"].format(error=message or UI_TEXT["error_capture_failed"]),
            )
        self._update_buttons()

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

    def _get_selected_window_record(self) -> WindowRecord | None:
        selection = self.window_listbox.curselection()
        if not selection:
            return None
        index = int(selection[0])
        if 0 <= index < len(self.window_records):
            return self.window_records[index]
        return None

    def set_orientation(self, orientation: str) -> None:
        if self.busy or orientation == self.orientation:
            return
        self.orientation = orientation
        self._clamp_all_items()
        self.draw_canvas()
        self._update_buttons()

    def add_image_to_canvas(self, image: Image.Image) -> None:
        page_width, page_height = self._page_size()
        fit_width, fit_height = fit_size(
            image.size,
            (round(page_width * 0.78), round(page_height * 0.78)),
        )
        offset = min(160, 28 * len(self.placed_images))
        x = max(0, (page_width - fit_width) / 2 + offset)
        y = max(0, (page_height - fit_height) / 2 + offset)

        item = PlacedImage(
            item_id=self.next_item_id,
            image=image,
            x=x,
            y=y,
            width=fit_width,
            height=fit_height,
        )
        self.next_item_id += 1
        self.placed_images.append(item)
        self.selected_item_id = item.item_id
        self._clamp_item(item)
        self.draw_canvas()
        self._update_buttons()

    def draw_canvas(self) -> None:
        canvas = self.canvas
        canvas.delete("all")

        canvas_width = max(1, canvas.winfo_width())
        canvas_height = max(1, canvas.winfo_height())
        page_width, page_height = self._page_size()
        ratio = page_width / page_height
        available_width = max(160, canvas_width - PREVIEW_PADDING * 2)
        available_height = max(160, canvas_height - PREVIEW_PADDING * 2)

        paper_width = min(available_width, round(available_height * ratio))
        paper_height = round(paper_width / ratio)
        if paper_height > available_height:
            paper_height = available_height
            paper_width = round(paper_height * ratio)

        paper_x = (canvas_width - paper_width) // 2
        paper_y = (canvas_height - paper_height) // 2
        self.paper_rect = (paper_x, paper_y, paper_width, paper_height)
        self.paper_scale = (paper_width / page_width, paper_height / page_height)

        canvas.create_rectangle(
            paper_x + 4,
            paper_y + 4,
            paper_x + paper_width + 4,
            paper_y + paper_height + 4,
            fill=THEME["shadow"],
            outline="",
        )
        canvas.create_rectangle(
            paper_x,
            paper_y,
            paper_x + paper_width,
            paper_y + paper_height,
            fill=THEME["paper"],
            outline=THEME["border"],
            width=1,
        )

        if not self.placed_images:
            center_x = paper_x + paper_width // 2
            center_y = paper_y + paper_height // 2
            canvas.create_text(
                center_x,
                center_y - 14,
                text=UI_TEXT["empty_title"],
                fill=THEME["muted"],
                font=(self.font_family, 14, "bold"),
            )
            canvas.create_text(
                center_x,
                center_y + 14,
                text=UI_TEXT["empty_subtitle"],
                fill=THEME["muted"],
                font=(self.font_family, 9),
                width=max(160, paper_width - 40),
            )
            return

        scale_x, scale_y = self.paper_scale
        for item in self.placed_images:
            x = paper_x + round(item.x * scale_x)
            y = paper_y + round(item.y * scale_y)
            width = max(1, round(item.width * scale_x))
            height = max(1, round(item.height * scale_y))
            preview = item.image.resize((width, height), RESAMPLE)
            item.photo = ImageTk.PhotoImage(preview)
            canvas.create_image(x, y, image=item.photo, anchor="nw", tags=(f"item_{item.item_id}",))
            outline = THEME["accent"] if item.item_id == self.selected_item_id else THEME["border"]
            line_width = 2 if item.item_id == self.selected_item_id else 1
            canvas.create_rectangle(
                x,
                y,
                x + width,
                y + height,
                outline=outline,
                width=line_width,
                dash=(5, 3) if item.item_id == self.selected_item_id else None,
            )

    def _on_canvas_press(self, event) -> None:
        self.canvas.focus_set()
        item = self._hit_test(event.x, event.y)
        if item is None:
            self.selected_item_id = None
            self.drag_state = None
            self.draw_canvas()
            self._update_buttons()
            return

        self.selected_item_id = item.item_id
        self.drag_state = (event.x, event.y, item.x, item.y)
        self.draw_canvas()
        self._update_buttons()

    def _on_canvas_drag(self, event) -> None:
        if self.drag_state is None or self.selected_item_id is None or self.paper_scale is None:
            return
        item = self._get_selected_item()
        if item is None:
            return

        start_x, start_y, original_x, original_y = self.drag_state
        scale_x, scale_y = self.paper_scale
        item.x = original_x + (event.x - start_x) / scale_x
        item.y = original_y + (event.y - start_y) / scale_y
        self._clamp_item(item)
        self.draw_canvas()

    def _on_canvas_release(self, _event) -> None:
        self.drag_state = None

    def _hit_test(self, event_x: int, event_y: int) -> PlacedImage | None:
        if self.paper_rect is None or self.paper_scale is None:
            return None

        paper_x, paper_y, _paper_width, _paper_height = self.paper_rect
        scale_x, scale_y = self.paper_scale
        logical_x = (event_x - paper_x) / scale_x
        logical_y = (event_y - paper_y) / scale_y

        for item in reversed(self.placed_images):
            if item.x <= logical_x <= item.x + item.width and item.y <= logical_y <= item.y + item.height:
                return item
        return None

    def resize_selected(self, factor: float) -> None:
        if self.busy:
            return
        item = self._get_selected_item()
        if item is None:
            return

        page_width, page_height = self._page_size()
        center_x = item.x + item.width / 2
        center_y = item.y + item.height / 2
        max_width = page_width * MAX_ITEM_PAGE_RATIO
        max_height = page_height * MAX_ITEM_PAGE_RATIO

        new_width = item.width * factor
        new_height = item.height * factor
        ratio = item.width / item.height if item.height else 1

        if new_width > max_width:
            new_width = max_width
            new_height = new_width / ratio
        if new_height > max_height:
            new_height = max_height
            new_width = new_height * ratio

        if new_width < MIN_ITEM_SIZE or new_height < MIN_ITEM_SIZE:
            return

        item.width = new_width
        item.height = new_height
        item.x = center_x - item.width / 2
        item.y = center_y - item.height / 2
        self._clamp_item(item)
        self.draw_canvas()

    def delete_selected(self, _event=None) -> None:
        if self.busy:
            return
        if self.selected_item_id is None:
            return

        self.placed_images = [item for item in self.placed_images if item.item_id != self.selected_item_id]
        self.selected_item_id = self.placed_images[-1].item_id if self.placed_images else None
        self._set_status(UI_TEXT["status_deleted"], "success")
        self.draw_canvas()
        self._update_buttons()

    def clear_all(self) -> None:
        if self.busy:
            return
        self.placed_images.clear()
        self.selected_item_id = None
        self._set_status(UI_TEXT["status_cleared"], "success")
        self.draw_canvas()
        self._update_buttons()

    def export_png(self) -> None:
        if self.busy or not self.placed_images:
            return

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        initial_file = UI_TEXT["export_file_template"].format(timestamp=timestamp)
        selected = filedialog.asksaveasfilename(
            parent=self.root,
            title=UI_TEXT["dialog_export_title"],
            initialdir=str(default_output_dir()),
            initialfile=initial_file,
            defaultextension=".png",
            filetypes=[(UI_TEXT["filetype_png"], "*.png")],
        )
        if not selected:
            return

        output_path = Path(selected)
        page_size = self._page_size()
        snapshot = [
            (
                item.image.copy(),
                int(round(item.x)),
                int(round(item.y)),
                max(1, int(round(item.width))),
                max(1, int(round(item.height))),
            )
            for item in self.placed_images
        ]

        self.busy = True
        self._set_status(UI_TEXT["status_exporting"], "working")
        self._update_buttons()
        self._clear_export_queue()
        worker = threading.Thread(
            target=self._export_worker,
            args=(output_path, page_size, snapshot),
            daemon=True,
        )
        worker.start()
        self.root.after(QUEUE_POLL_INTERVAL_MS, self._poll_export_queue)

    def _export_worker(
        self,
        output_path: Path,
        page_size: tuple[int, int],
        snapshot: list[tuple[Image.Image, int, int, int, int]],
    ) -> None:
        try:
            page = Image.new("RGB", page_size, "white")
            for image, x, y, width, height in snapshot:
                try:
                    resized = image.resize((width, height), RESAMPLE)
                    page.paste(normalize_image(resized), (x, y))
                finally:
                    image.close()
            output_path.parent.mkdir(parents=True, exist_ok=True)
            page.save(output_path, format="PNG")
            self.export_queue.put(("export_complete", output_path, None))
        except Exception as exc:
            self.export_queue.put(("export_error", None, str(exc)))

    def _poll_export_queue(self) -> None:
        try:
            kind, path, message = self.export_queue.get_nowait()
        except queue.Empty:
            if self.busy:
                self.root.after(QUEUE_POLL_INTERVAL_MS, self._poll_export_queue)
            return

        self.busy = False
        if kind == "export_complete" and path is not None:
            self._set_status(UI_TEXT["status_export_complete"].format(name=path.name), "success")
            open_folder(path.parent)
        else:
            self._set_status(UI_TEXT["status_error"], "error")
            messagebox.showerror(
                UI_TEXT["dialog_error_title"],
                UI_TEXT["dialog_export_error_message"].format(error=message or ""),
            )
        self._update_buttons()

    def _clear_export_queue(self) -> None:
        while True:
            try:
                self.export_queue.get_nowait()
            except queue.Empty:
                return

    def _get_selected_item(self) -> PlacedImage | None:
        if self.selected_item_id is None:
            return None
        for item in self.placed_images:
            if item.item_id == self.selected_item_id:
                return item
        return None

    def _page_size(self) -> tuple[int, int]:
        return A4_LANDSCAPE if self.orientation == "landscape" else A4_PORTRAIT

    def _clamp_all_items(self) -> None:
        for item in self.placed_images:
            self._clamp_item(item)

    def _clamp_item(self, item: PlacedImage) -> None:
        page_width, page_height = self._page_size()
        if item.width > page_width * MAX_ITEM_PAGE_RATIO or item.height > page_height * MAX_ITEM_PAGE_RATIO:
            new_width, new_height = fit_size(
                (int(item.width), int(item.height)),
                (round(page_width * MAX_ITEM_PAGE_RATIO), round(page_height * MAX_ITEM_PAGE_RATIO)),
            )
            item.width = new_width
            item.height = new_height

        item.x = min(max(0, item.x), max(0, page_width - item.width))
        item.y = min(max(0, item.y), max(0, page_height - item.height))

    def _set_status(self, text: str, status_type: str) -> None:
        self.status_var.set(text)
        background, foreground = STATUS_THEME.get(status_type, STATUS_THEME["idle"])
        self.status_badge.configure(bg=background, fg=foreground)

    def _update_buttons(self) -> None:
        has_selection = self.selected_item_id is not None
        has_images = bool(self.placed_images)
        has_window_selection = bool(self.window_listbox.curselection())
        base_state = tk.DISABLED if self.busy else tk.NORMAL
        selected_state = tk.NORMAL if has_selection and not self.busy else tk.DISABLED
        image_state = tk.NORMAL if has_images and not self.busy else tk.DISABLED
        paste_state = tk.NORMAL if has_window_selection and not self.busy else tk.DISABLED

        self.refresh_button.configure(state=base_state)
        self.paste_button.configure(state=paste_state)
        self.smaller_button.configure(state=selected_state)
        self.larger_button.configure(state=selected_state)
        self.delete_button.configure(state=selected_state)
        self.clear_button.configure(state=image_state)
        self.export_button.configure(state=image_state)
        self.portrait_button.configure(state=base_state)
        self.landscape_button.configure(state=base_state)

        self._update_orientation_buttons()

    def _update_orientation_buttons(self) -> None:
        for key, button in (("portrait", self.portrait_button), ("landscape", self.landscape_button)):
            selected = key == self.orientation
            button.configure(
                bg=THEME["accent"] if selected else THEME["card"],
                fg="#FFFFFF" if selected else THEME["text"],
                activebackground=THEME["accent_hover"] if selected else THEME["soft"],
                activeforeground="#FFFFFF" if selected else THEME["text"],
                highlightthickness=1,
                highlightbackground=THEME["accent"] if selected else THEME["border"],
                disabledforeground="#FFFFFF" if selected else THEME["muted"],
            )


def main() -> None:
    enable_dpi_awareness()
    app = DakePasteA4App()
    app.run()


if __name__ == "__main__":
    main()
