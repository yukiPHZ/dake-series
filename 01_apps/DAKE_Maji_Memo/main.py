# -*- coding: utf-8 -*-
from __future__ import annotations

import ctypes
import json
import os
import re
import sys
import webbrowser
from datetime import datetime
from pathlib import Path
from ctypes import wintypes


APP_NAME = "Dakeマジでメモ"
WINDOW_TITLE = "Dakeマジでメモ"
COPYRIGHT = "© 2026 しまりす不動産 / Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "main_title": "メモを置く",
    "main_description": "保存を気にせず、今の考えをそのまま置いておきます。",
    "label_main": "メイン",
    "label_sub": "補助",
    "button_refresh": "リフレッシュ",
    "status_idle": "そのまま残ります",
    "status_refreshed": "リフレッシュしました",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_tagline": "止まらない、迷わない、すぐ終わる。",
    "footer_link_1": "戸建買取査定",
    "footer_link_2": "Instagram",
    "footer_separator": " ｜ ",
    "footer_copyright": COPYRIGHT,
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
}

FONT_FACE = "BIZ UDPGothic"
WINDOW_DEFAULT_SIZE = (960, 620)
WINDOW_MIN_SIZE = (760, 480)
CONFIG_FILE_NAME = "DAKE_Maji_Memo_config.json"
SAVE_DEBOUNCE_MS = 380
STATUS_RESET_MS = 1800
APP_USER_MODEL_ID = "Shimarisu.DakeMajiMemo"

CLASS_NAME = "DakeMajiMemoWindow"
SAVE_TIMER_ID = 1
STATUS_TIMER_ID = 2

ID_TITLE = 101
ID_DESCRIPTION = 102
ID_REFRESH = 103
ID_LABEL_MAIN = 104
ID_LABEL_SUB = 105
ID_EDIT_MAIN = 201
ID_EDIT_SUB = 202
ID_STATUS = 301
ID_FOOTER_LEFT = 302
ID_FOOTER_LINK_1 = 303
ID_FOOTER_SEPARATOR_1 = 304
ID_FOOTER_LINK_2 = 305
ID_FOOTER_SEPARATOR_2 = 306
ID_FOOTER_COPYRIGHT = 307

WS_OVERLAPPEDWINDOW = 0x00CF0000
WS_VISIBLE = 0x10000000
WS_CHILD = 0x40000000
WS_CLIPSIBLINGS = 0x04000000
WS_TABSTOP = 0x00010000
WS_VSCROLL = 0x00200000
WS_EX_CLIENTEDGE = 0x00000200

ES_MULTILINE = 0x0004
ES_AUTOVSCROLL = 0x0040
ES_WANTRETURN = 0x1000
ES_NOHIDESEL = 0x0100
SS_NOTIFY = 0x0100
BS_PUSHBUTTON = 0x00000000

SW_SHOW = 5
COLOR_WINDOW = 5
IDC_ARROW = 32512
IDC_HAND = 32649

WM_CREATE = 0x0001
WM_DESTROY = 0x0002
WM_SIZE = 0x0005
WM_PAINT = 0x000F
WM_CLOSE = 0x0010
WM_GETMINMAXINFO = 0x0024
WM_SETFONT = 0x0030
WM_SETTEXT = 0x000C
WM_GETTEXT = 0x000D
WM_GETTEXTLENGTH = 0x000E
WM_COMMAND = 0x0111
WM_TIMER = 0x0113
WM_CTLCOLORSTATIC = 0x0138
WM_CTLCOLOREDIT = 0x0133
WM_SETCURSOR = 0x0020
EM_SETMARGINS = 0x00D3
EC_LEFTMARGIN = 0x0001
EC_RIGHTMARGIN = 0x0002

EN_CHANGE = 0x0300
STN_CLICKED = 0
BN_CLICKED = 0
TRANSPARENT = 1
OPAQUE = 2
FW_NORMAL = 400
FW_BOLD = 700
DEFAULT_CHARSET = 1
CLEARTYPE_QUALITY = 5
DEFAULT_PITCH = 0
LOGPIXELSY = 90

IMAGE_ICON = 1
LR_LOADFROMFILE = 0x0010
LR_DEFAULTSIZE = 0x0040
ICON_SMALL = 0
ICON_BIG = 1

GWL_ID = -12

HICON = wintypes.HANDLE
HCURSOR = wintypes.HANDLE
HBRUSH = wintypes.HANDLE
HFONT = wintypes.HANDLE
HMENU = wintypes.HANDLE
UINT_PTR = ctypes.c_size_t

WNDPROC = ctypes.WINFUNCTYPE(ctypes.c_ssize_t, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)

user32 = ctypes.windll.user32
gdi32 = ctypes.windll.gdi32
kernel32 = ctypes.windll.kernel32
shell32 = ctypes.windll.shell32


class WNDCLASSEXW(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.UINT),
        ("style", wintypes.UINT),
        ("lpfnWndProc", WNDPROC),
        ("cbClsExtra", ctypes.c_int),
        ("cbWndExtra", ctypes.c_int),
        ("hInstance", wintypes.HINSTANCE),
        ("hIcon", HICON),
        ("hCursor", HCURSOR),
        ("hbrBackground", HBRUSH),
        ("lpszMenuName", wintypes.LPCWSTR),
        ("lpszClassName", wintypes.LPCWSTR),
        ("hIconSm", HICON),
    ]


class PAINTSTRUCT(ctypes.Structure):
    _fields_ = [
        ("hdc", wintypes.HDC),
        ("fErase", wintypes.BOOL),
        ("rcPaint", wintypes.RECT),
        ("fRestore", wintypes.BOOL),
        ("fIncUpdate", wintypes.BOOL),
        ("rgbReserved", wintypes.BYTE * 32),
    ]


class POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]


class MINMAXINFO(ctypes.Structure):
    _fields_ = [
        ("ptReserved", POINT),
        ("ptMaxSize", POINT),
        ("ptMaxPosition", POINT),
        ("ptMinTrackSize", POINT),
        ("ptMaxTrackSize", POINT),
    ]


class MSG(ctypes.Structure):
    _fields_ = [
        ("hwnd", wintypes.HWND),
        ("message", wintypes.UINT),
        ("wParam", wintypes.WPARAM),
        ("lParam", wintypes.LPARAM),
        ("time", wintypes.DWORD),
        ("pt", POINT),
    ]


class SIZE(ctypes.Structure):
    _fields_ = [("cx", ctypes.c_long), ("cy", ctypes.c_long)]


user32.CreateWindowExW.restype = wintypes.HWND
user32.CreateWindowExW.argtypes = [
    wintypes.DWORD,
    wintypes.LPCWSTR,
    wintypes.LPCWSTR,
    wintypes.DWORD,
    ctypes.c_int,
    ctypes.c_int,
    ctypes.c_int,
    ctypes.c_int,
    wintypes.HWND,
    HMENU,
    wintypes.HINSTANCE,
    wintypes.LPVOID,
]
user32.DefWindowProcW.restype = ctypes.c_ssize_t
user32.DefWindowProcW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
user32.GetWindowLongPtrW.restype = ctypes.c_ssize_t
user32.GetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int]
user32.SendMessageW.restype = ctypes.c_ssize_t
user32.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
user32.LoadCursorW.restype = HCURSOR
user32.LoadImageW.restype = wintypes.HANDLE
user32.BeginPaint.restype = wintypes.HDC
user32.BeginPaint.argtypes = [wintypes.HWND, ctypes.POINTER(PAINTSTRUCT)]
user32.EndPaint.restype = wintypes.BOOL
user32.EndPaint.argtypes = [wintypes.HWND, ctypes.POINTER(PAINTSTRUCT)]
user32.FillRect.restype = ctypes.c_int
user32.FillRect.argtypes = [wintypes.HDC, ctypes.POINTER(wintypes.RECT), HBRUSH]
user32.FrameRect.restype = ctypes.c_int
user32.FrameRect.argtypes = [wintypes.HDC, ctypes.POINTER(wintypes.RECT), HBRUSH]
user32.SetTimer.restype = UINT_PTR
user32.SetTimer.argtypes = [wintypes.HWND, UINT_PTR, wintypes.UINT, wintypes.LPVOID]
user32.KillTimer.restype = wintypes.BOOL
user32.KillTimer.argtypes = [wintypes.HWND, UINT_PTR]
user32.DestroyIcon.restype = wintypes.BOOL
user32.DestroyIcon.argtypes = [HICON]
kernel32.GetModuleHandleW.restype = wintypes.HINSTANCE
kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]
gdi32.CreateSolidBrush.restype = HBRUSH
gdi32.CreateSolidBrush.argtypes = [wintypes.COLORREF]
gdi32.CreateFontW.restype = HFONT
gdi32.DeleteObject.restype = wintypes.BOOL
gdi32.DeleteObject.argtypes = [wintypes.HANDLE]
gdi32.SetBkMode.restype = ctypes.c_int
gdi32.SetBkMode.argtypes = [wintypes.HDC, ctypes.c_int]
gdi32.SetBkColor.restype = wintypes.COLORREF
gdi32.SetBkColor.argtypes = [wintypes.HDC, wintypes.COLORREF]
gdi32.SetTextColor.restype = wintypes.COLORREF
gdi32.SetTextColor.argtypes = [wintypes.HDC, wintypes.COLORREF]


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APP_DIR = app_dir()
CONFIG_PATH = APP_DIR / CONFIG_FILE_NAME


def rgb(value: str) -> int:
    color = value.lstrip("#")
    r = int(color[0:2], 16)
    g = int(color[2:4], 16)
    b = int(color[4:6], 16)
    return r | (g << 8) | (b << 16)


def loword(value: int) -> int:
    return value & 0xFFFF


def hiword(value: int) -> int:
    return (value >> 16) & 0xFFFF


def makelong(low: int, high: int) -> int:
    return (high << 16) | (low & 0xFFFF)


def to_control_text(value: str) -> str:
    return value.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")


def from_control_text(value: str) -> str:
    return value.replace("\r\n", "\n").replace("\r", "\n")


def set_windows_app_id() -> None:
    if not sys.platform.startswith("win"):
        return
    try:
        shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)
    except Exception:
        return


def icon_candidates() -> list[Path]:
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        return [
            exe_dir / ".." / ".." / ".." / "02_assets" / "dake_icon.ico",
            exe_dir / ".." / ".." / "02_assets" / "dake_icon.ico",
        ]
    base = Path(__file__).resolve().parent
    return [
        base / ".." / ".." / "02_assets" / "dake_icon.ico",
        Path("..") / ".." / "02_assets" / "dake_icon.ico",
    ]


def load_icon() -> int:
    for candidate in icon_candidates():
        try:
            icon_path = candidate.resolve()
        except Exception:
            icon_path = candidate
        if not icon_path.exists():
            continue
        handle = user32.LoadImageW(None, str(icon_path), IMAGE_ICON, 0, 0, LR_LOADFROMFILE | LR_DEFAULTSIZE)
        if handle:
            return int(handle)
    return 0


def load_config() -> dict[str, object]:
    if not CONFIG_PATH.exists():
        return {}
    try:
        with CONFIG_PATH.open("r", encoding="utf-8") as file:
            data = json.load(file)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def text_value(data: dict[str, object], key: str) -> str:
    value = data.get(key)
    return value if isinstance(value, str) else ""


def parse_geometry(value: object) -> tuple[int, int, int | None, int | None]:
    if not isinstance(value, str):
        return (*WINDOW_DEFAULT_SIZE, None, None)
    match = re.fullmatch(r"(\d+)x(\d+)([+-]\d+)?([+-]\d+)?", value)
    if not match:
        return (*WINDOW_DEFAULT_SIZE, None, None)
    width = max(int(match.group(1)), WINDOW_MIN_SIZE[0])
    height = max(int(match.group(2)), WINDOW_MIN_SIZE[1])
    x = int(match.group(3)) if match.group(3) else None
    y = int(match.group(4)) if match.group(4) else None
    return width, height, x, y


def center_position(width: int, height: int) -> tuple[int, int]:
    screen_w = user32.GetSystemMetrics(0)
    screen_h = user32.GetSystemMetrics(1)
    return max((screen_w - width) // 2, 0), max((screen_h - height) // 2, 0)


class MajiMemoApp:
    def __init__(self) -> None:
        self.instance = kernel32.GetModuleHandleW(None)
        self.config = load_config()
        self.window_proc = WNDPROC(self._wnd_proc)
        self.hwnd = wintypes.HWND()
        self.icon = load_icon()
        self.suppress_change = False
        self.card_rects: list[wintypes.RECT] = []
        self.controls: dict[int, int] = {}
        self.card_controls: set[int] = set()
        self.link_controls: set[int] = set()
        self.brushes = {
            "background": gdi32.CreateSolidBrush(rgb(THEME["background"])),
            "card": gdi32.CreateSolidBrush(rgb(THEME["card"])),
            "border": gdi32.CreateSolidBrush(rgb(THEME["border"])),
        }
        self.fonts: dict[str, int] = {}

    def run(self) -> int:
        self._register_class()
        width, height, x, y = parse_geometry(self.config.get("window_geometry"))
        if x is None or y is None:
            x, y = center_position(width, height)

        self.hwnd = user32.CreateWindowExW(
            0,
            CLASS_NAME,
            WINDOW_TITLE,
            WS_OVERLAPPEDWINDOW | WS_VISIBLE,
            x,
            y,
            width,
            height,
            None,
            None,
            self.instance,
            None,
        )
        if not self.hwnd:
            return 1

        if self.icon:
            user32.SendMessageW(self.hwnd, 0x0080, ICON_BIG, self.icon)
            user32.SendMessageW(self.hwnd, 0x0080, ICON_SMALL, self.icon)

        user32.ShowWindow(self.hwnd, SW_SHOW)
        user32.UpdateWindow(self.hwnd)
        user32.SetFocus(self.controls[ID_EDIT_MAIN])

        message = MSG()
        while user32.GetMessageW(ctypes.byref(message), None, 0, 0) != 0:
            user32.TranslateMessage(ctypes.byref(message))
            user32.DispatchMessageW(ctypes.byref(message))
        self._cleanup_gdi()
        return int(message.wParam)

    def _register_class(self) -> None:
        h_icon = self.icon or None
        wnd_class = WNDCLASSEXW(
            ctypes.sizeof(WNDCLASSEXW),
            0,
            self.window_proc,
            0,
            0,
            self.instance,
            h_icon,
            user32.LoadCursorW(None, IDC_ARROW),
            self.brushes["background"],
            None,
            CLASS_NAME,
            h_icon,
        )
        user32.RegisterClassExW(ctypes.byref(wnd_class))

    def _create_font(self, hwnd: int, point_size: int, weight: int = FW_NORMAL, underline: bool = False) -> int:
        hdc = user32.GetDC(hwnd)
        dpi = gdi32.GetDeviceCaps(hdc, LOGPIXELSY) if hdc else 96
        if hdc:
            user32.ReleaseDC(hwnd, hdc)
        height = -int(point_size * dpi / 72)
        return int(
            gdi32.CreateFontW(
                height,
                0,
                0,
                0,
                weight,
                0,
                int(underline),
                0,
                DEFAULT_CHARSET,
                0,
                0,
                CLEARTYPE_QUALITY,
                DEFAULT_PITCH,
                FONT_FACE,
            )
        )

    def _init_fonts(self, hwnd: int) -> None:
        self.fonts = {
            "title": self._create_font(hwnd, 18, FW_BOLD),
            "description": self._create_font(hwnd, 10),
            "label": self._create_font(hwnd, 10, FW_BOLD),
            "input": self._create_font(hwnd, 12),
            "button": self._create_font(hwnd, 10),
            "status": self._create_font(hwnd, 9),
            "footer": self._create_font(hwnd, 8),
            "footer_link": self._create_font(hwnd, 8, FW_NORMAL, True),
        }

    def _create_controls(self, hwnd: int) -> None:
        self._init_fonts(hwnd)
        self.controls[ID_TITLE] = self._static(hwnd, ID_TITLE, UI_TEXT["main_title"], self.fonts["title"])
        self.controls[ID_DESCRIPTION] = self._static(
            hwnd,
            ID_DESCRIPTION,
            UI_TEXT["main_description"],
            self.fonts["description"],
        )
        self.controls[ID_REFRESH] = self._button(hwnd, ID_REFRESH, UI_TEXT["button_refresh"], self.fonts["button"])
        self.controls[ID_LABEL_MAIN] = self._static(hwnd, ID_LABEL_MAIN, UI_TEXT["label_main"], self.fonts["label"])
        self.controls[ID_LABEL_SUB] = self._static(hwnd, ID_LABEL_SUB, UI_TEXT["label_sub"], self.fonts["label"])
        self.controls[ID_EDIT_MAIN] = self._edit(
            hwnd,
            ID_EDIT_MAIN,
            text_value(self.config, "left_text"),
            self.fonts["input"],
        )
        self.controls[ID_EDIT_SUB] = self._edit(
            hwnd,
            ID_EDIT_SUB,
            text_value(self.config, "right_text"),
            self.fonts["input"],
        )
        self.controls[ID_STATUS] = self._static(hwnd, ID_STATUS, UI_TEXT["status_idle"], self.fonts["status"])
        footer_left = UI_TEXT["footer_left"] + UI_TEXT["footer_separator"] + UI_TEXT["footer_tagline"]
        self.controls[ID_FOOTER_LEFT] = self._static(hwnd, ID_FOOTER_LEFT, footer_left, self.fonts["footer"])
        self.controls[ID_FOOTER_LINK_1] = self._static(
            hwnd,
            ID_FOOTER_LINK_1,
            UI_TEXT["footer_link_1"],
            self.fonts["footer_link"],
            notify=True,
        )
        self.controls[ID_FOOTER_SEPARATOR_1] = self._static(
            hwnd,
            ID_FOOTER_SEPARATOR_1,
            UI_TEXT["footer_separator"],
            self.fonts["footer"],
        )
        self.controls[ID_FOOTER_LINK_2] = self._static(
            hwnd,
            ID_FOOTER_LINK_2,
            UI_TEXT["footer_link_2"],
            self.fonts["footer_link"],
            notify=True,
        )
        self.controls[ID_FOOTER_SEPARATOR_2] = self._static(
            hwnd,
            ID_FOOTER_SEPARATOR_2,
            UI_TEXT["footer_separator"],
            self.fonts["footer"],
        )
        self.controls[ID_FOOTER_COPYRIGHT] = self._static(
            hwnd,
            ID_FOOTER_COPYRIGHT,
            UI_TEXT["footer_copyright"],
            self.fonts["footer"],
        )

        self.card_controls.update(
            {
                self.controls[ID_LABEL_MAIN],
                self.controls[ID_LABEL_SUB],
            }
        )
        self.link_controls.update(
            {
                self.controls[ID_FOOTER_LINK_1],
                self.controls[ID_FOOTER_LINK_2],
            }
        )

    def _static(self, parent: int, control_id: int, text: str, font: int, notify: bool = False) -> int:
        style = WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS
        if notify:
            style |= SS_NOTIFY
        hwnd = user32.CreateWindowExW(
            0,
            "STATIC",
            text,
            style,
            0,
            0,
            10,
            10,
            parent,
            control_id,
            self.instance,
            None,
        )
        user32.SendMessageW(hwnd, WM_SETFONT, font, True)
        return int(hwnd)

    def _button(self, parent: int, control_id: int, text: str, font: int) -> int:
        hwnd = user32.CreateWindowExW(
            0,
            "BUTTON",
            text,
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
            0,
            0,
            10,
            10,
            parent,
            control_id,
            self.instance,
            None,
        )
        user32.SendMessageW(hwnd, WM_SETFONT, font, True)
        return int(hwnd)

    def _edit(self, parent: int, control_id: int, text: str, font: int) -> int:
        hwnd = user32.CreateWindowExW(
            WS_EX_CLIENTEDGE,
            "EDIT",
            to_control_text(text),
            WS_CHILD
            | WS_VISIBLE
            | WS_TABSTOP
            | WS_VSCROLL
            | ES_MULTILINE
            | ES_AUTOVSCROLL
            | ES_WANTRETURN
            | ES_NOHIDESEL,
            0,
            0,
            10,
            10,
            parent,
            control_id,
            self.instance,
            None,
        )
        user32.SendMessageW(hwnd, WM_SETFONT, font, True)
        user32.SendMessageW(hwnd, EM_SETMARGINS, EC_LEFTMARGIN | EC_RIGHTMARGIN, makelong(12, 12))
        return int(hwnd)

    def _layout(self, width: int, height: int) -> None:
        margin_x = 30
        margin_top = 24
        margin_bottom = 24
        footer_h = 40 if width < 900 else 20
        status_h = 22
        header_h = 72
        body_top = margin_top + header_h + 8
        footer_y = height - margin_bottom - footer_h
        status_y = footer_y - 8 - status_h
        body_h = max(status_y - 12 - body_top, 220)
        content_w = max(width - margin_x * 2, WINDOW_MIN_SIZE[0] - margin_x * 2)

        button_w = 116
        button_h = 34
        header_text_w = max(content_w - button_w - 18, 320)

        self._move(ID_TITLE, margin_x, margin_top, header_text_w, 30)
        self._move(ID_DESCRIPTION, margin_x, margin_top + 35, header_text_w, 24)
        self._move(ID_REFRESH, width - margin_x - button_w, margin_top + 8, button_w, button_h)

        gap = 24
        left_w = (content_w - gap) * 58 // 100
        right_w = content_w - gap - left_w
        left_x = margin_x
        right_x = margin_x + left_w + gap
        self.card_rects = [
            wintypes.RECT(left_x, body_top, left_x + left_w, body_top + body_h),
            wintypes.RECT(right_x, body_top, right_x + right_w, body_top + body_h),
        ]

        self._move(ID_LABEL_MAIN, left_x + 18, body_top + 14, left_w - 36, 22)
        self._move(ID_LABEL_SUB, right_x + 18, body_top + 14, right_w - 36, 22)
        self._move(ID_EDIT_MAIN, left_x + 18, body_top + 44, left_w - 36, body_h - 62)
        self._move(ID_EDIT_SUB, right_x + 18, body_top + 44, right_w - 36, body_h - 62)
        self._move(ID_STATUS, margin_x, status_y, content_w, status_h)

        self._move(ID_FOOTER_LEFT, margin_x, footer_y, min(380, content_w), 18)
        if width < 900:
            right_y = footer_y + 18
            right_x = margin_x
        else:
            right_y = footer_y
            right_x = width - margin_x - 482
        self._move(ID_FOOTER_LINK_1, right_x, right_y, 92, 18)
        self._move(ID_FOOTER_SEPARATOR_1, right_x + 92, right_y, 26, 18)
        self._move(ID_FOOTER_LINK_2, right_x + 118, right_y, 78, 18)
        self._move(ID_FOOTER_SEPARATOR_2, right_x + 196, right_y, 26, 18)
        self._move(ID_FOOTER_COPYRIGHT, right_x + 222, right_y, 260, 18)
        user32.InvalidateRect(self.hwnd, None, True)

    def _move(self, control_id: int, x: int, y: int, width: int, height: int) -> None:
        handle = self.controls.get(control_id)
        if handle:
            user32.MoveWindow(handle, x, y, width, height, True)

    def _paint(self, hwnd: int) -> None:
        ps = PAINTSTRUCT()
        hdc = user32.BeginPaint(hwnd, ctypes.byref(ps))
        client = wintypes.RECT()
        user32.GetClientRect(hwnd, ctypes.byref(client))
        user32.FillRect(hdc, ctypes.byref(client), self.brushes["background"])
        for rect in self.card_rects:
            user32.FillRect(hdc, ctypes.byref(rect), self.brushes["card"])
            user32.FrameRect(hdc, ctypes.byref(rect), self.brushes["border"])
        user32.EndPaint(hwnd, ctypes.byref(ps))

    def _get_edit_text(self, control_id: int) -> str:
        hwnd = self.controls[control_id]
        length = user32.GetWindowTextLengthW(hwnd)
        buffer = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buffer, length + 1)
        return from_control_text(buffer.value)

    def _set_status(self, text: str) -> None:
        user32.SetWindowTextW(self.controls[ID_STATUS], text)

    def _schedule_save(self) -> None:
        user32.KillTimer(self.hwnd, SAVE_TIMER_ID)
        user32.SetTimer(self.hwnd, SAVE_TIMER_ID, SAVE_DEBOUNCE_MS, None)

    def _save_now(self) -> None:
        user32.KillTimer(self.hwnd, SAVE_TIMER_ID)
        payload = {
            "left_text": self._get_edit_text(ID_EDIT_MAIN),
            "right_text": self._get_edit_text(ID_EDIT_SUB),
            "window_geometry": self._window_geometry(),
            "last_updated": datetime.now().isoformat(timespec="seconds"),
        }
        try:
            CONFIG_PATH.write_text(
                json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )
        except Exception:
            return

    def _window_geometry(self) -> str:
        rect = wintypes.RECT()
        user32.GetWindowRect(self.hwnd, ctypes.byref(rect))
        width = max(rect.right - rect.left, WINDOW_MIN_SIZE[0])
        height = max(rect.bottom - rect.top, WINDOW_MIN_SIZE[1])
        return f"{width}x{height}+{rect.left}+{rect.top}"

    def _refresh(self) -> None:
        self.suppress_change = True
        try:
            user32.SetWindowTextW(self.controls[ID_EDIT_MAIN], "")
            user32.SetWindowTextW(self.controls[ID_EDIT_SUB], "")
        finally:
            self.suppress_change = False
        self._save_now()
        self._set_status(UI_TEXT["status_refreshed"])
        user32.KillTimer(self.hwnd, STATUS_TIMER_ID)
        user32.SetTimer(self.hwnd, STATUS_TIMER_ID, STATUS_RESET_MS, None)
        user32.SetFocus(self.controls[ID_EDIT_MAIN])

    def _handle_command(self, wparam: int) -> int:
        control_id = loword(wparam)
        notification = hiword(wparam)
        if control_id in (ID_EDIT_MAIN, ID_EDIT_SUB) and notification == EN_CHANGE:
            if not self.suppress_change:
                self._schedule_save()
            return 0
        if control_id == ID_REFRESH and notification == BN_CLICKED:
            self._refresh()
            return 0
        if control_id == ID_FOOTER_LINK_1 and notification == STN_CLICKED:
            webbrowser.open(LINK_URLS["footer_link_1"])
            return 0
        if control_id == ID_FOOTER_LINK_2 and notification == STN_CLICKED:
            webbrowser.open(LINK_URLS["footer_link_2"])
            return 0
        return 0

    def _color_static(self, hdc: int, hwnd: int) -> int:
        handle = int(hwnd)
        bg_key = "card" if handle in self.card_controls else "background"
        gdi32.SetBkMode(hdc, OPAQUE)
        gdi32.SetBkColor(hdc, rgb(THEME[bg_key]))
        color = THEME["text"] if handle == self.controls.get(ID_TITLE) else THEME["muted"]
        gdi32.SetTextColor(hdc, rgb(color))
        return int(self.brushes[bg_key])

    def _color_edit(self, hdc: int) -> int:
        gdi32.SetBkMode(hdc, OPAQUE)
        gdi32.SetBkColor(hdc, rgb(THEME["card"]))
        gdi32.SetTextColor(hdc, rgb(THEME["text"]))
        return int(self.brushes["card"])

    def _cleanup_gdi(self) -> None:
        for font in self.fonts.values():
            if font:
                gdi32.DeleteObject(font)
        for brush in self.brushes.values():
            if brush:
                gdi32.DeleteObject(brush)
        if self.icon:
            user32.DestroyIcon(self.icon)

    def _wnd_proc(self, hwnd: int, message: int, wparam: int, lparam: int) -> int:
        if message == WM_CREATE:
            self.hwnd = hwnd
            self._create_controls(hwnd)
            rect = wintypes.RECT()
            user32.GetClientRect(hwnd, ctypes.byref(rect))
            self._layout(rect.right - rect.left, rect.bottom - rect.top)
            return 0
        if message == WM_SIZE:
            width = loword(lparam)
            height = hiword(lparam)
            if self.controls:
                self._layout(width, height)
            return 0
        if message == WM_PAINT:
            self._paint(hwnd)
            return 0
        if message == WM_COMMAND:
            return self._handle_command(wparam)
        if message == WM_TIMER:
            if wparam == SAVE_TIMER_ID:
                self._save_now()
                return 0
            if wparam == STATUS_TIMER_ID:
                user32.KillTimer(hwnd, STATUS_TIMER_ID)
                self._set_status(UI_TEXT["status_idle"])
                return 0
        if message == WM_CTLCOLORSTATIC:
            return self._color_static(wparam, lparam)
        if message == WM_CTLCOLOREDIT:
            return self._color_edit(wparam)
        if message == WM_SETCURSOR:
            if int(wparam) in self.link_controls:
                user32.SetCursor(user32.LoadCursorW(None, IDC_HAND))
                return 1
        if message == WM_GETMINMAXINFO:
            info = ctypes.cast(lparam, ctypes.POINTER(MINMAXINFO)).contents
            info.ptMinTrackSize.x = WINDOW_MIN_SIZE[0]
            info.ptMinTrackSize.y = WINDOW_MIN_SIZE[1]
            return 0
        if message == WM_CLOSE:
            self._save_now()
            user32.DestroyWindow(hwnd)
            return 0
        if message == WM_DESTROY:
            user32.PostQuitMessage(0)
            return 0
        return user32.DefWindowProcW(hwnd, message, wparam, lparam)


def main() -> None:
    set_windows_app_id()
    app = MajiMemoApp()
    raise SystemExit(app.run())


if __name__ == "__main__":
    main()
