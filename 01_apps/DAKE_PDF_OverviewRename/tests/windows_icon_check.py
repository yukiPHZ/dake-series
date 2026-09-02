# -*- coding: utf-8 -*-
"""Check the built exe icon resource and the running Tk window icon handle."""

from __future__ import annotations

import ctypes
import json
import sys
from ctypes import wintypes
from pathlib import Path


WINDOW_TITLE = "DakePDF俯瞰名前変更"
WM_GETICON = 0x007F
ICON_SMALL = 0
ICON_BIG = 1
ICON_SMALL2 = 2
GCLP_HICON = -14
GCLP_HICONSM = -34


def main() -> None:
    if not sys.platform.startswith("win"):
        raise SystemExit("Windows only")
    app_dir = Path(__file__).resolve().parents[1]
    exe = app_dir / "dist" / "DakePDF_OverviewRename.exe"
    common_icon = app_dir / ".." / ".." / "02_assets" / "dake_icon.ico"
    user32 = ctypes.windll.user32
    shell32 = ctypes.windll.shell32
    user32.FindWindowW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR]
    user32.FindWindowW.restype = wintypes.HWND
    user32.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
    user32.SendMessageW.restype = wintypes.LPARAM
    user32.GetClassLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int]
    user32.GetClassLongPtrW.restype = ctypes.c_void_p
    shell32.ExtractIconExW.argtypes = [wintypes.LPCWSTR, ctypes.c_int, ctypes.POINTER(wintypes.HICON), ctypes.POINTER(wintypes.HICON), wintypes.UINT]
    shell32.ExtractIconExW.restype = wintypes.UINT

    large = wintypes.HICON()
    small = wintypes.HICON()
    icon_count = shell32.ExtractIconExW(str(exe), 0, ctypes.byref(large), ctypes.byref(small), 1)
    hwnd = user32.FindWindowW(None, WINDOW_TITLE)
    window_icon = 0
    if hwnd:
        for icon_type in (ICON_SMALL2, ICON_SMALL, ICON_BIG):
            window_icon = int(user32.SendMessageW(hwnd, WM_GETICON, icon_type, 0))
            if window_icon:
                break
        if not window_icon:
            window_icon = int(user32.GetClassLongPtrW(hwnd, GCLP_HICONSM) or user32.GetClassLongPtrW(hwnd, GCLP_HICON) or 0)

    result = {
        "exe_exists": exe.is_file(),
        "exe_icon_resources": int(icon_count),
        "common_icon_exists": common_icon.resolve().is_file(),
        "window_found": bool(hwnd),
        "window_icon_handle_nonzero": bool(window_icon),
    }
    if large:
        user32.DestroyIcon(large)
    if small:
        user32.DestroyIcon(small)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    assert all(result.values()), result


if __name__ == "__main__":
    main()
