# -*- coding: utf-8 -*-
from __future__ import annotations

import ctypes
import os
import queue
import sys
import threading
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import font as tkfont
from tkinter import messagebox, ttk

from PIL import Image, ImageGrab, ImageOps, ImageTk


APP_NAME = "Dakeスクショ印刷"
WINDOW_TITLE = "スクショ印刷"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "main_title": "スクショをA4縦で印刷する",
    "main_description": "Win + Shift + S で切り取った画像を、そのままA4縦に印刷します。",
    "button_clipboard": "クリップボードから取得",
    "button_print": "印刷",
    "button_clear": "クリア",
    "status_label": "状態",
    "status_idle": "未取得",
    "status_acquired": "取得しました",
    "status_printing": "印刷中",
    "status_print_complete": "印刷完了",
    "status_no_image": "クリップボードに画像がありません",
    "status_print_failed": "印刷できませんでした",
    "dialog_print_timeout_title": "印刷を確認してください",
    "dialog_print_timeout_message": "プリンターの応答に時間がかかっています。プリンター側の状態を確認してください。\n\nアプリの操作は続けられます。",
    "empty_title": "画像がありません",
    "empty_description": "Win + Shift + S のあと、クリップボードから取得してください。",
    "dialog_no_image_title": "画像がありません",
    "dialog_no_image_message": "Win + Shift + S で切り取ったあとに、もう一度取得してください。",
    "dialog_clipboard_error_title": "取得できませんでした",
    "dialog_clipboard_error_message": "クリップボードを確認してから、もう一度お試しください。\n\n{error}",
    "dialog_print_error_title": "印刷できませんでした",
    "dialog_print_error_message": "プリンター設定を確認してから、もう一度お試しください。\n\n{error}",
    "error_windows_only": "Windowsで実行してください。",
    "error_no_default_printer": "通常使うプリンターが見つかりません。",
    "error_printer_dc": "プリンターへ接続できません。",
    "error_start_doc": "印刷ジョブを開始できません。",
    "error_start_page": "印刷ページを開始できません。",
    "error_end_page": "印刷ページを完了できません。",
    "error_end_doc": "印刷ジョブを完了できません。",
    "footer_left": "シンプルそれDAKEシリーズ / 止まらない、迷わない、すぐ終わる。",
    "footer_link_1": "戸建買取査定",
    "footer_link_2": "Instagram",
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
    "soft": "#EEF2F7",
    "success": "#12B76A",
    "success_bg": "#EAFBF3",
    "error": "#D92D20",
    "error_bg": "#FDECEC",
}

LINKS = {
    "assessment": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "instagram": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

STATUS_THEME = {
    "status_idle": (THEME["soft"], THEME["muted"]),
    "status_acquired": ("#EAF2FF", THEME["accent"]),
    "status_printing": ("#EAF2FF", THEME["accent"]),
    "status_print_complete": (THEME["success_bg"], THEME["success"]),
    "status_no_image": (THEME["error_bg"], THEME["error"]),
    "status_print_failed": (THEME["error_bg"], THEME["error"]),
}

WINDOW_SIZE = "900x760"
WINDOW_MIN_SIZE = (760, 660)
A4_LOGICAL_WIDTH = 2480
A4_LOGICAL_HEIGHT = 3508
A4_MARGIN = 180
MAX_LAYOUT_UPSCALE = 2.0
PREVIEW_PADDING = 24
QUEUE_POLL_INTERVAL_MS = 80
PRINT_TIMEOUT_MS = 45000
RESAMPLE = Image.Resampling.LANCZOS if hasattr(Image, "Resampling") else Image.LANCZOS


def get_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def get_common_icon_candidates() -> list[Path]:
    base_dir = get_base_dir()
    source_dir = Path(__file__).resolve().parent
    return [
        source_dir / ".." / ".." / "02_assets" / "dake_icon.ico",
        base_dir / ".." / ".." / "02_assets" / "dake_icon.ico",
        base_dir / ".." / ".." / ".." / "02_assets" / "dake_icon.ico",
    ]


def choose_font_family(root: tk.Tk) -> str:
    preferred = ["BIZ UDPGothic", "Yu Gothic UI", "Meiryo"]
    available = set(tkfont.families(root))
    for family in preferred:
        if family in available:
            return family
    return "TkDefaultFont"


def normalize_clipboard_image(image: Image.Image) -> Image.Image:
    fixed_image = ImageOps.exif_transpose(image)
    if fixed_image.mode == "RGB":
        return fixed_image.copy()

    if "A" in fixed_image.getbands() or "transparency" in fixed_image.info:
        rgba_image = fixed_image.convert("RGBA")
        background = Image.new("RGB", rgba_image.size, "white")
        background.paste(rgba_image, mask=rgba_image.getchannel("A"))
        return background

    return fixed_image.convert("RGB")


def calculate_fit_size(
    source_size: tuple[int, int],
    box_size: tuple[int, int],
    max_upscale: float = MAX_LAYOUT_UPSCALE,
) -> tuple[int, int]:
    source_width, source_height = source_size
    box_width, box_height = box_size
    if source_width <= 0 or source_height <= 0 or box_width <= 0 or box_height <= 0:
        return 1, 1

    scale = min(box_width / source_width, box_height / source_height)
    scale = min(scale, max_upscale)
    width = max(1, round(source_width * scale))
    height = max(1, round(source_height * scale))
    return min(width, box_width), min(height, box_height)


def calculate_a4_image_rect(image_size: tuple[int, int]) -> tuple[int, int, int, int]:
    printable_width = A4_LOGICAL_WIDTH - (A4_MARGIN * 2)
    printable_height = A4_LOGICAL_HEIGHT - (A4_MARGIN * 2)
    image_width, image_height = calculate_fit_size(image_size, (printable_width, printable_height))
    image_x = (A4_LOGICAL_WIDTH - image_width) // 2
    image_y = (A4_LOGICAL_HEIGHT - image_height) // 2
    return image_x, image_y, image_width, image_height


class ClipboardImageService:
    def load_image(self) -> Image.Image | None:
        data = ImageGrab.grabclipboard()
        if isinstance(data, Image.Image):
            return normalize_clipboard_image(data)
        return None


class DOCINFOW(ctypes.Structure):
    _fields_ = [
        ("cbSize", ctypes.c_int),
        ("lpszDocName", ctypes.c_wchar_p),
        ("lpszOutput", ctypes.c_wchar_p),
        ("lpszDatatype", ctypes.c_wchar_p),
        ("fwType", ctypes.c_uint32),
    ]


class DEVMODEW(ctypes.Structure):
    _fields_ = [
        ("dmDeviceName", ctypes.c_wchar * 32),
        ("dmSpecVersion", ctypes.c_uint16),
        ("dmDriverVersion", ctypes.c_uint16),
        ("dmSize", ctypes.c_uint16),
        ("dmDriverExtra", ctypes.c_uint16),
        ("dmFields", ctypes.c_uint32),
        ("dmOrientation", ctypes.c_int16),
        ("dmPaperSize", ctypes.c_int16),
        ("dmPaperLength", ctypes.c_int16),
        ("dmPaperWidth", ctypes.c_int16),
        ("dmScale", ctypes.c_int16),
        ("dmCopies", ctypes.c_int16),
        ("dmDefaultSource", ctypes.c_int16),
        ("dmPrintQuality", ctypes.c_int16),
        ("dmColor", ctypes.c_int16),
        ("dmDuplex", ctypes.c_int16),
        ("dmYResolution", ctypes.c_int16),
        ("dmTTOption", ctypes.c_int16),
        ("dmCollate", ctypes.c_int16),
        ("dmFormName", ctypes.c_wchar * 32),
        ("dmLogPixels", ctypes.c_uint16),
        ("dmBitsPerPel", ctypes.c_uint32),
        ("dmPelsWidth", ctypes.c_uint32),
        ("dmPelsHeight", ctypes.c_uint32),
        ("dmDisplayFlags", ctypes.c_uint32),
        ("dmDisplayFrequency", ctypes.c_uint32),
        ("dmICMMethod", ctypes.c_uint32),
        ("dmICMIntent", ctypes.c_uint32),
        ("dmMediaType", ctypes.c_uint32),
        ("dmDitherType", ctypes.c_uint32),
        ("dmReserved1", ctypes.c_uint32),
        ("dmReserved2", ctypes.c_uint32),
        ("dmPanningWidth", ctypes.c_uint32),
        ("dmPanningHeight", ctypes.c_uint32),
    ]


class WindowsPrintService:
    HORZRES = 8
    VERTRES = 10
    LOGPIXELSX = 88
    LOGPIXELSY = 90
    DM_ORIENTATION = 0x00000001
    DM_PAPERSIZE = 0x00000002
    DMORIENT_PORTRAIT = 1
    DMPAPER_A4 = 9
    DM_OUT_BUFFER = 0x00000002
    DM_IN_BUFFER = 0x00000008

    def print_image(self, image: Image.Image, document_name: str) -> None:
        if os.name != "nt":
            raise RuntimeError(UI_TEXT["error_windows_only"])

        printer_name = self._get_default_printer_name()
        gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)
        hdc = self._create_printer_dc(printer_name, gdi32)
        if not hdc:
            raise RuntimeError(UI_TEXT["error_printer_dc"])

        doc_started = False
        try:
            doc_info = DOCINFOW()
            doc_info.cbSize = ctypes.sizeof(DOCINFOW)
            doc_info.lpszDocName = document_name

            gdi32.StartDocW.argtypes = [ctypes.c_void_p, ctypes.POINTER(DOCINFOW)]
            gdi32.StartDocW.restype = ctypes.c_int
            gdi32.StartPage.argtypes = [ctypes.c_void_p]
            gdi32.StartPage.restype = ctypes.c_int
            gdi32.EndPage.argtypes = [ctypes.c_void_p]
            gdi32.EndPage.restype = ctypes.c_int
            gdi32.EndDoc.argtypes = [ctypes.c_void_p]
            gdi32.EndDoc.restype = ctypes.c_int
            gdi32.AbortDoc.argtypes = [ctypes.c_void_p]
            gdi32.DeleteDC.argtypes = [ctypes.c_void_p]

            if gdi32.StartDocW(hdc, ctypes.byref(doc_info)) <= 0:
                raise RuntimeError(UI_TEXT["error_start_doc"])
            doc_started = True

            if gdi32.StartPage(hdc) <= 0:
                raise RuntimeError(UI_TEXT["error_start_page"])

            self._draw_image(gdi32, hdc, image)

            if gdi32.EndPage(hdc) <= 0:
                raise RuntimeError(UI_TEXT["error_end_page"])
            if gdi32.EndDoc(hdc) <= 0:
                raise RuntimeError(UI_TEXT["error_end_doc"])
            doc_started = False
        except Exception:
            if doc_started:
                gdi32.AbortDoc(hdc)
            raise
        finally:
            gdi32.DeleteDC(hdc)

    def _get_default_printer_name(self) -> str:
        winspool = ctypes.WinDLL("winspool.drv", use_last_error=True)
        winspool.GetDefaultPrinterW.argtypes = [ctypes.c_wchar_p, ctypes.POINTER(ctypes.c_uint32)]
        winspool.GetDefaultPrinterW.restype = ctypes.c_int

        needed = ctypes.c_uint32(0)
        winspool.GetDefaultPrinterW(None, ctypes.byref(needed))
        if needed.value <= 0:
            raise RuntimeError(UI_TEXT["error_no_default_printer"])

        buffer = ctypes.create_unicode_buffer(needed.value)
        if not winspool.GetDefaultPrinterW(buffer, ctypes.byref(needed)):
            raise RuntimeError(UI_TEXT["error_no_default_printer"])
        return buffer.value

    def _create_printer_dc(self, printer_name: str, gdi32) -> ctypes.c_void_p:
        devmode_pointer = self._build_a4_portrait_devmode(printer_name)

        gdi32.CreateDCW.argtypes = [ctypes.c_wchar_p, ctypes.c_wchar_p, ctypes.c_wchar_p, ctypes.c_void_p]
        gdi32.CreateDCW.restype = ctypes.c_void_p

        if devmode_pointer is not None:
            hdc = gdi32.CreateDCW("WINSPOOL", printer_name, None, ctypes.cast(devmode_pointer, ctypes.c_void_p))
            if hdc:
                return hdc

        return gdi32.CreateDCW("WINSPOOL", printer_name, None, None)

    def _build_a4_portrait_devmode(self, printer_name: str):
        winspool = ctypes.WinDLL("winspool.drv", use_last_error=True)
        printer_handle = ctypes.c_void_p()

        winspool.OpenPrinterW.argtypes = [ctypes.c_wchar_p, ctypes.POINTER(ctypes.c_void_p), ctypes.c_void_p]
        winspool.OpenPrinterW.restype = ctypes.c_int
        winspool.ClosePrinter.argtypes = [ctypes.c_void_p]
        winspool.DocumentPropertiesW.argtypes = [
            ctypes.c_void_p,
            ctypes.c_void_p,
            ctypes.c_wchar_p,
            ctypes.c_void_p,
            ctypes.c_void_p,
            ctypes.c_uint32,
        ]
        winspool.DocumentPropertiesW.restype = ctypes.c_long

        if not winspool.OpenPrinterW(printer_name, ctypes.byref(printer_handle), None):
            return None

        try:
            needed = winspool.DocumentPropertiesW(None, printer_handle, printer_name, None, None, 0)
            if needed <= 0:
                return None

            buffer = ctypes.create_string_buffer(needed)
            devmode_pointer = ctypes.cast(buffer, ctypes.POINTER(DEVMODEW))
            result = winspool.DocumentPropertiesW(
                None,
                printer_handle,
                printer_name,
                ctypes.cast(devmode_pointer, ctypes.c_void_p),
                None,
                self.DM_OUT_BUFFER,
            )
            if result < 0:
                return None

            devmode = devmode_pointer.contents
            devmode.dmFields |= self.DM_ORIENTATION | self.DM_PAPERSIZE
            devmode.dmOrientation = self.DMORIENT_PORTRAIT
            devmode.dmPaperSize = self.DMPAPER_A4

            result = winspool.DocumentPropertiesW(
                None,
                printer_handle,
                printer_name,
                ctypes.cast(devmode_pointer, ctypes.c_void_p),
                ctypes.cast(devmode_pointer, ctypes.c_void_p),
                self.DM_IN_BUFFER | self.DM_OUT_BUFFER,
            )
            if result < 0:
                return None

            self._devmode_buffer = buffer
            return devmode_pointer
        finally:
            winspool.ClosePrinter(printer_handle)

    def _draw_image(self, gdi32, hdc, image: Image.Image) -> None:
        from PIL import ImageWin

        gdi32.GetDeviceCaps.argtypes = [ctypes.c_void_p, ctypes.c_int]
        gdi32.GetDeviceCaps.restype = ctypes.c_int

        page_width = gdi32.GetDeviceCaps(hdc, self.HORZRES)
        page_height = gdi32.GetDeviceCaps(hdc, self.VERTRES)
        dpi_x = max(72, gdi32.GetDeviceCaps(hdc, self.LOGPIXELSX))
        dpi_y = max(72, gdi32.GetDeviceCaps(hdc, self.LOGPIXELSY))
        margin_x = max(1, round(dpi_x * 15 / 25.4))
        margin_y = max(1, round(dpi_y * 15 / 25.4))
        printable_width = max(1, page_width - (margin_x * 2))
        printable_height = max(1, page_height - (margin_y * 2))

        image_width, image_height = calculate_fit_size(image.size, (printable_width, printable_height))
        left = max(0, (page_width - image_width) // 2)
        top = max(0, (page_height - image_height) // 2)
        right = left + image_width
        bottom = top + image_height

        dib = ImageWin.Dib(image.convert("RGB"))
        dib.draw(hdc, (left, top, right, bottom))


class ScreenshotPrintApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(*WINDOW_MIN_SIZE)
        self.root.configure(bg=THEME["background"])

        self.font_family = choose_font_family(self.root)
        self.root.option_add("*Font", (self.font_family, 10))

        self.clipboard_service = ClipboardImageService()
        self.print_service = WindowsPrintService()
        self.print_queue: queue.Queue[tuple[int, str, str | None]] = queue.Queue()
        self.image: Image.Image | None = None
        self.preview_photo: ImageTk.PhotoImage | None = None
        self.busy = False
        self.print_job_id = 0
        self.current_print_job_id: int | None = None

        self.status_var = tk.StringVar(value=UI_TEXT["status_idle"])

        self._apply_window_icon()
        self._build_styles()
        self._build_ui()
        self._set_status("status_idle")
        self.root.after(150, self.load_from_clipboard)

    def run(self) -> None:
        self.root.mainloop()

    def _build_styles(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(
            "Secondary.TButton",
            font=(self.font_family, 10, "bold"),
            padding=(16, 10),
            background=THEME["card"],
            foreground=THEME["text"],
            bordercolor=THEME["border"],
            lightcolor=THEME["card"],
            darkcolor=THEME["card"],
        )
        style.map("Secondary.TButton", background=[("active", THEME["soft"])])

        style.configure(
            "Primary.TButton",
            font=(self.font_family, 10, "bold"),
            padding=(22, 10),
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

    def _build_ui(self) -> None:
        container = tk.Frame(self.root, bg=THEME["background"])
        container.pack(fill="both", expand=True, padx=24, pady=20)

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

        actions = tk.Frame(container, bg=THEME["background"])
        actions.pack(fill="x", pady=(16, 14))

        self.clipboard_button = ttk.Button(
            actions,
            text=UI_TEXT["button_clipboard"],
            style="Secondary.TButton",
            command=self.load_from_clipboard,
        )
        self.clipboard_button.pack(side="left", padx=(0, 10))

        self.clear_button = ttk.Button(
            actions,
            text=UI_TEXT["button_clear"],
            style="Secondary.TButton",
            command=self.clear_image,
        )
        self.clear_button.pack(side="left")

        self.print_button = ttk.Button(
            actions,
            text=UI_TEXT["button_print"],
            style="Primary.TButton",
            command=self.print_current_image,
        )
        self.print_button.pack(side="right")

        preview_card = tk.Frame(
            container,
            bg=THEME["card"],
            highlightbackground=THEME["border"],
            highlightthickness=1,
        )
        preview_card.pack(fill="both", expand=True)

        status_row = tk.Frame(preview_card, bg=THEME["card"])
        status_row.pack(fill="x", padx=18, pady=(14, 0))

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

        self.preview_canvas = tk.Canvas(
            preview_card,
            bg=THEME["card"],
            bd=0,
            highlightthickness=0,
        )
        self.preview_canvas.pack(fill="both", expand=True, padx=18, pady=14)
        self.preview_canvas.bind("<Configure>", lambda _event: self._draw_preview())

        self._build_footer(container)
        self._update_buttons()

    def _build_footer(self, container: tk.Frame) -> None:
        footer = tk.Frame(container, bg=THEME["background"])
        footer.pack(fill="x", pady=(12, 0))

        footer_left = tk.Label(
            footer,
            text=UI_TEXT["footer_left"],
            bg=THEME["background"],
            fg=THEME["muted"],
            font=(self.font_family, 8),
            anchor="w",
        )
        footer_left.pack(side="left", fill="x", expand=True)

        footer_right = tk.Frame(footer, bg=THEME["background"])
        footer_right.pack(side="right")

        self._make_footer_link(footer_right, UI_TEXT["footer_link_1"], LINKS["assessment"])
        self._make_footer_text(footer_right, UI_TEXT["footer_separator"])
        self._make_footer_link(footer_right, UI_TEXT["footer_link_2"], LINKS["instagram"])
        self._make_footer_text(footer_right, UI_TEXT["footer_separator"])
        self._make_footer_text(footer_right, UI_TEXT["footer_copyright"])

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

    def load_from_clipboard(self) -> None:
        if self.busy:
            return

        try:
            image = self.clipboard_service.load_image()
        except Exception as exc:
            self.image = None
            self._set_status("status_no_image")
            self._update_buttons()
            self._draw_preview()
            messagebox.showerror(
                UI_TEXT["dialog_clipboard_error_title"],
                UI_TEXT["dialog_clipboard_error_message"].format(error=exc),
            )
            return

        if image is None:
            self.image = None
            self._set_status("status_no_image")
            self._update_buttons()
            self._draw_preview()
            return

        self.image = image
        self._set_status("status_acquired")
        self._update_buttons()
        self._draw_preview()

    def clear_image(self) -> None:
        if self.busy:
            return
        self.image = None
        self.preview_photo = None
        self._set_status("status_idle")
        self._update_buttons()
        self._draw_preview()

    def print_current_image(self) -> None:
        if self.busy:
            return

        if self.image is None:
            self._set_status("status_no_image")
            messagebox.showinfo(UI_TEXT["dialog_no_image_title"], UI_TEXT["dialog_no_image_message"])
            return

        image_for_print = self.image.copy()
        self._clear_print_queue()
        self.print_job_id += 1
        job_id = self.print_job_id
        self.current_print_job_id = job_id
        self.busy = True
        self._set_status("status_printing")
        self._update_buttons()

        worker = threading.Thread(target=self._print_worker, args=(job_id, image_for_print), daemon=True)
        worker.start()
        self.root.after(QUEUE_POLL_INTERVAL_MS, self._poll_print_queue)
        self.root.after(PRINT_TIMEOUT_MS, lambda: self._handle_print_timeout(job_id))

    def _print_worker(self, job_id: int, image: Image.Image) -> None:
        try:
            self.print_service.print_image(image, APP_NAME)
        except Exception as exc:
            self.print_queue.put((job_id, "error", str(exc)))
            return
        self.print_queue.put((job_id, "complete", None))

    def _poll_print_queue(self) -> None:
        try:
            job_id, result, message = self.print_queue.get_nowait()
        except queue.Empty:
            if self.busy:
                self.root.after(QUEUE_POLL_INTERVAL_MS, self._poll_print_queue)
            return

        if job_id != self.current_print_job_id:
            if self.busy:
                self.root.after(QUEUE_POLL_INTERVAL_MS, self._poll_print_queue)
            return

        self.busy = False
        self.current_print_job_id = None
        if result == "complete":
            self._set_status("status_print_complete")
        else:
            self._set_status("status_print_failed")
            messagebox.showerror(
                UI_TEXT["dialog_print_error_title"],
                UI_TEXT["dialog_print_error_message"].format(error=message or ""),
            )
        self._update_buttons()

    def _handle_print_timeout(self, job_id: int) -> None:
        if not self.busy or self.current_print_job_id != job_id:
            return

        self.busy = False
        self.current_print_job_id = None
        self._set_status("status_print_failed")
        self._update_buttons()
        messagebox.showwarning(UI_TEXT["dialog_print_timeout_title"], UI_TEXT["dialog_print_timeout_message"])

    def _clear_print_queue(self) -> None:
        while True:
            try:
                self.print_queue.get_nowait()
            except queue.Empty:
                return

    def _set_status(self, status_key: str) -> None:
        self.status_var.set(UI_TEXT[status_key])
        background, foreground = STATUS_THEME.get(status_key, STATUS_THEME["status_idle"])
        self.status_badge.configure(bg=background, fg=foreground)

    def _update_buttons(self) -> None:
        normal = tk.NORMAL if not self.busy else tk.DISABLED
        image_state = tk.NORMAL if self.image is not None and not self.busy else tk.DISABLED
        self.clipboard_button.configure(state=normal)
        self.clear_button.configure(state=image_state)
        self.print_button.configure(state=image_state)

    def _draw_preview(self) -> None:
        canvas = self.preview_canvas
        canvas.delete("all")

        canvas_width = max(1, canvas.winfo_width())
        canvas_height = max(1, canvas.winfo_height())
        available_width = max(120, canvas_width - PREVIEW_PADDING * 2)
        available_height = max(160, canvas_height - PREVIEW_PADDING * 2)
        a4_ratio = A4_LOGICAL_WIDTH / A4_LOGICAL_HEIGHT

        paper_width = min(available_width, round(available_height * a4_ratio))
        paper_height = round(paper_width / a4_ratio)
        if paper_height > available_height:
            paper_height = available_height
            paper_width = round(paper_height * a4_ratio)

        paper_x = (canvas_width - paper_width) // 2
        paper_y = (canvas_height - paper_height) // 2
        paper_right = paper_x + paper_width
        paper_bottom = paper_y + paper_height

        canvas.create_rectangle(
            paper_x,
            paper_y,
            paper_right,
            paper_bottom,
            fill="#FFFFFF",
            outline=THEME["border"],
            width=1,
        )

        margin_x = max(1, round(paper_width * A4_MARGIN / A4_LOGICAL_WIDTH))
        margin_y = max(1, round(paper_height * A4_MARGIN / A4_LOGICAL_HEIGHT))
        canvas.create_rectangle(
            paper_x + margin_x,
            paper_y + margin_y,
            paper_right - margin_x,
            paper_bottom - margin_y,
            outline="#F1F3F6",
            width=1,
        )

        if self.image is None:
            self.preview_photo = None
            center_x = paper_x + paper_width // 2
            center_y = paper_y + paper_height // 2
            canvas.create_text(
                center_x,
                center_y - 12,
                text=UI_TEXT["empty_title"],
                fill=THEME["muted"],
                font=(self.font_family, 13, "bold"),
            )
            canvas.create_text(
                center_x,
                center_y + 14,
                text=UI_TEXT["empty_description"],
                fill=THEME["muted"],
                font=(self.font_family, 9),
                width=max(160, paper_width - margin_x * 2),
            )
            return

        logical_x, logical_y, logical_width, logical_height = calculate_a4_image_rect(self.image.size)
        scale_x = paper_width / A4_LOGICAL_WIDTH
        scale_y = paper_height / A4_LOGICAL_HEIGHT
        preview_x = paper_x + round(logical_x * scale_x)
        preview_y = paper_y + round(logical_y * scale_y)
        preview_width = max(1, round(logical_width * scale_x))
        preview_height = max(1, round(logical_height * scale_y))

        preview_image = self.image.resize((preview_width, preview_height), RESAMPLE)
        self.preview_photo = ImageTk.PhotoImage(preview_image)
        canvas.create_image(preview_x, preview_y, anchor="nw", image=self.preview_photo)
        canvas.create_rectangle(
            preview_x,
            preview_y,
            preview_x + preview_width,
            preview_y + preview_height,
            outline=THEME["border"],
            width=1,
        )


def main() -> None:
    app = ScreenshotPrintApp()
    app.run()


if __name__ == "__main__":
    main()
