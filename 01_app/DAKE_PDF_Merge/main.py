import json
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import ctypes
except Exception:
    ctypes = None

from pypdf import PdfReader, PdfWriter

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except Exception:
    DND_FILES = None
    TkinterDnD = None
    HAS_DND = False

APP_NAME = "DakePDF結合"
WINDOW_TITLE = "DakePDF結合"
CONFIG_NAME = "dake_pdf_merge_config.json"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

BG = "#F6F7F9"
CARD = "#FFFFFF"
TEXT = "#1E2430"
SUBTEXT = "#667085"
ACCENT = "#2F6FED"
ACCENT_HOVER = "#2458BF"
BORDER = "#E6EAF0"
PREVIEW_BG = "#F6F8FC"
PREVIEW_BORDER = "#C9D3E3"
SUCCESS = ACCENT
DISABLED_BG = "#E8ECF3"
DISABLED_FG = "#98A2B3"
ERROR = "#D92D20"
FOOTER_TEXT = "#AAB2BD"
FONT_CANDIDATES = ["BIZ UDPGothic", "Yu Gothic UI", "Meiryo"]


def app_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def make_root():
    if HAS_DND:
        return TkinterDnD.Tk()
    return tk.Tk()


def detect_font_name():
    root = tk.Tk()
    root.withdraw()
    try:
        families = set(root.tk.call("font", "families"))
    finally:
        root.destroy()
    for name in FONT_CANDIDATES:
        if name in families:
            return name
    return "TkDefaultFont"


FONT_NAME = detect_font_name()


def icon_ico_path() -> str:
    return os.path.join(app_dir(), "app.ico")


def icon_png_path() -> str:
    return os.path.join(app_dir(), "icon.png")


def apply_window_icon(window: tk.Misc):
    try:
        ico = icon_ico_path()
        if os.path.exists(ico):
            window.iconbitmap(default=ico)
    except Exception:
        pass
    try:
        png = icon_png_path()
        if os.path.exists(png):
            image = tk.PhotoImage(file=png)
            window._app_icon_ref = image
            window.iconphoto(True, image)
            try:
                window.wm_iconphoto(True, image)
            except Exception:
                pass
    except Exception:
        pass


def set_windows_app_id():
    if not sys.platform.startswith("win") or ctypes is None:
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Shimarisu.DakePDFMerge")
    except Exception:
        pass


def config_path() -> str:
    return os.path.join(app_dir(), CONFIG_NAME)


def default_downloads() -> str:
    return os.path.join(os.path.expanduser("~"), "Downloads")


def load_config() -> dict:
    try:
        with open(config_path(), "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_config(data: dict) -> None:
    try:
        with open(config_path(), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def shorten_path(path: str, max_len: int = 56) -> str:
    if len(path) <= max_len:
        return path
    drive, rest = os.path.splitdrive(path)
    parts = rest.strip("\\/").split(os.sep)
    if len(parts) <= 2:
        return path[: max_len - 1] + "…"
    return f"{drive}\\…\\{parts[-2]}\\{parts[-1]}"


def open_folder(path: str) -> None:
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


def extract_drop_paths(raw: str) -> list[str]:
    try:
        parts = list(tk.Tcl().splitlist(raw))
    except Exception:
        parts = raw.split()
    cleaned = []
    for part in parts:
        part = part.strip().strip("{}")
        if part:
            cleaned.append(part)
    return cleaned


class ModernButton(tk.Label):
    def __init__(self, parent, text, command, *, primary=False, width=12):
        self.command = command
        self.primary = primary
        self.enabled = True
        self.normal_bg = ACCENT if primary else "#FFFFFF"
        self.normal_fg = "#FFFFFF" if primary else TEXT
        self.hover_bg = ACCENT_HOVER if primary else "#F2F4F7"
        super().__init__(
            parent,
            text=text,
            bg=self.normal_bg,
            fg=self.normal_fg,
            padx=16,
            pady=9,
            cursor="hand2",
            bd=1,
            relief="solid",
            width=width,
            font=(FONT_NAME, 10, "bold"),
        )
        self.bind("<Button-1>", self._on_click)
        self.bind("<Enter>", self._hover_in)
        self.bind("<Leave>", self._hover_out)

    def _on_click(self, _event):
        if self.enabled:
            self.command()

    def _hover_in(self, _event):
        if self.enabled:
            self.configure(bg=self.hover_bg)

    def _hover_out(self, _event):
        self._apply_visual()

    def set_enabled(self, enabled: bool):
        self.enabled = enabled
        self._apply_visual()

    def _apply_visual(self):
        if self.enabled:
            self.configure(bg=self.normal_bg, fg=self.normal_fg, cursor="hand2")
        else:
            self.configure(bg=DISABLED_BG, fg=DISABLED_FG, cursor="arrow")


class MergeFileCard(tk.Frame):
    def __init__(self, parent, app, path: str, index: int):
        super().__init__(parent, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        self.app = app
        self.path = path
        self.index = index
        self.thumb_image = None
        self.top = None
        self.number_label = None
        self.thumb_frame = None
        self.thumb_label = None
        self.name_label = None
        self.meta = None
        self.buttons = []
        self.build_ui()

    def build_ui(self):
        self.top = tk.Frame(self, bg=CARD)
        self.top.pack(fill="x", padx=12, pady=(10, 6))

        self.number_label = tk.Label(
            self.top,
            text=f"{self.index + 1:02d}",
            font=(FONT_NAME, 11, "bold"),
            bg=CARD,
            fg=ACCENT,
        )
        self.number_label.pack(side="left", padx=(0, 8))

        btns = tk.Frame(self.top, bg=CARD)
        btns.pack(side="right")

        actions = [
            ("↑", lambda: self.app.move_file(self.index, -1)),
            ("↓", lambda: self.app.move_file(self.index, 1)),
            ("削除", lambda: self.app.remove_file(self.index)),
        ]
        for label, cmd in actions:
            button = tk.Label(
                btns,
                text=label,
                bg="#FFFFFF",
                fg=TEXT,
                padx=8,
                pady=5,
                cursor="hand2",
                bd=1,
                relief="solid",
                font=(FONT_NAME, 9, "bold"),
            )
            button.pack(side="left", padx=3)
            button.bind("<Button-1>", lambda _e, c=cmd: c())
            button.bind("<Enter>", lambda _e, w=button: w.configure(bg="#F2F4F7"))
            button.bind("<Leave>", lambda _e, w=button: w.configure(bg="#FFFFFF"))
            self.buttons.append(button)

        self.thumb_frame = tk.Frame(self, bg=PREVIEW_BG, highlightthickness=1, highlightbackground=PREVIEW_BORDER)
        self.thumb_frame.pack(padx=12, pady=(0, 8))
        self.thumb_label = tk.Label(
            self.thumb_frame,
            text="読み込み中",
            bg="#FFFFFF",
            fg=SUBTEXT,
            width=18,
            height=14,
            font=(FONT_NAME, 9),
        )
        self.thumb_label.pack(padx=1, pady=1)

        self.name_label = tk.Label(
            self,
            text=os.path.basename(self.path),
            bg=CARD,
            fg=TEXT,
            font=(FONT_NAME, 10, "bold"),
            wraplength=150,
            justify="left",
        )
        self.name_label.pack(fill="x", padx=12)

        self.meta = tk.Label(
            self,
            text="ページ数取得中",
            bg=CARD,
            fg=SUBTEXT,
            font=(FONT_NAME, 9),
            anchor="w",
        )
        self.meta.pack(fill="x", padx=12, pady=(3, 10))

    def set_page_count(self, count: int | None):
        self.meta.configure(text=f"{count}ページ" if count is not None else "ページ数不明")

    def set_thumbnail(self, img):
        self.thumb_image = img
        self.thumb_label.configure(image=img, text="", width=160, height=220, bg="#FFFFFF")

    def update_visual(self):
        self.configure(bg=CARD, highlightbackground=BORDER)
        self.top.configure(bg=CARD)
        self.number_label.configure(bg=CARD, fg=ACCENT)
        self.thumb_frame.configure(highlightbackground=PREVIEW_BORDER)
        self.name_label.configure(bg=CARD)
        self.meta.configure(bg=CARD)
        for button in self.buttons:
            button.configure(bg="#FFFFFF")


class DAKEPDFMergeApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.geometry("1480x900")
        self.root.minsize(1180, 760)
        self.root.configure(bg=BG)
        apply_window_icon(self.root)

        self.cfg = load_config()
        self.save_folder = self.cfg.get("last_folder", default_downloads())
        self.files: list[str] = []
        self.page_count_cache: dict[str, int] = {}
        self.merge_card_by_path: dict[str, MergeFileCard] = {}
        self.thumbnail_queue: queue.Queue = queue.Queue()
        self.worker_running = False
        self.cancel_requested = False
        self._status_anim_job = None
        self._status_anim_base = "待機中"
        self._status_anim_dots = 0

        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")
        except Exception:
            pass
        self.style.configure(
            "Horizontal.TProgressbar",
            troughcolor="#E9EEF7",
            bordercolor="#E9EEF7",
            lightcolor=ACCENT,
            darkcolor=ACCENT,
            background=ACCENT,
            thickness=16,
        )

        self.action_buttons: list[ModernButton] = []

        self.build_ui()
        self.root.after(120, self.process_thumbnail_queue)

    def build_ui(self):
        shell = tk.Frame(self.root, bg=BG)
        shell.pack(fill="both", expand=True)

        header = tk.Frame(shell, bg=BG)
        header.pack(fill="x", padx=26, pady=(18, 10))
        tk.Label(header, text="DAKE PDF Merge", font=(FONT_NAME, 20, "bold"), bg=BG, fg=TEXT).pack(side="left")
        tk.Label(header, text="結合専用 / 迷わず、止まらず、すぐ終わる", font=(FONT_NAME, 10), bg=BG, fg=SUBTEXT).pack(side="left", padx=(12, 0), pady=(6, 0))

        main = tk.Frame(shell, bg=BG)
        main.pack(fill="both", expand=True, padx=20, pady=(0, 8))

        title = tk.Frame(main, bg=BG)
        title.pack(fill="x", pady=(2, 10))
        tk.Label(title, text="PDFを結合する", font=(FONT_NAME, 22, "bold"), bg=BG, fg=TEXT).pack(anchor="w")
        tk.Label(title, text="順番を整えて、そのまま保存します", font=(FONT_NAME, 10), bg=BG, fg=SUBTEXT).pack(anchor="w", pady=(4, 0))

        top_card = tk.Frame(main, bg="#FFFFFF", highlightthickness=1, highlightbackground=BORDER)
        top_card.pack(fill="x", pady=(0, 12))

        row1 = tk.Frame(top_card, bg="#FFFFFF")
        row1.pack(fill="x", padx=14, pady=(12, 8))
        self.add_button = ModernButton(row1, "PDFを追加", self.add_files, primary=True, width=10)
        self.add_button.pack(side="left")
        self.folder_button = ModernButton(row1, "保存先を選ぶ", self.choose_folder, width=12)
        self.folder_button.pack(side="left", padx=8)
        self.refresh_button = ModernButton(row1, "リフレッシュ", self.reset_merge, width=10)
        self.refresh_button.pack(side="left", padx=8)
        self.action_buttons.extend([self.add_button, self.folder_button, self.refresh_button])

        self.status_label = tk.Label(row1, text="まだPDFがありません", font=(FONT_NAME, 10, "bold"), bg="#FFFFFF", fg=TEXT)
        self.status_label.pack(side="left", padx=(16, 8))
        self.folder_short_label = tk.Label(row1, text=f"保存先: {shorten_path(self.save_folder)}", font=(FONT_NAME, 9), bg="#FFFFFF", fg=SUBTEXT)
        self.folder_short_label.pack(side="right")

        row2 = tk.Frame(top_card, bg="#FFFFFF")
        row2.pack(fill="x", padx=14, pady=(0, 12))
        dnd_text = "画面へのドラッグ＆ドロップ対応" if HAS_DND else "ドラッグ＆ドロップは未導入環境です（ボタン追加は使用可能）"
        tk.Label(row2, text=f"横並び表示 / ↑↓で順番変更 / {dnd_text}", font=(FONT_NAME, 9), bg="#FFFFFF", fg=SUBTEXT).pack(anchor="w")

        body = tk.Frame(main, bg=BG)
        body.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(body, bg=BG, highlightthickness=0)
        self.scrollbar = tk.Scrollbar(body, orient="vertical", command=self.canvas.yview)
        self.cards_wrap = tk.Frame(self.canvas, bg=BG)
        self.cards_wrap.bind("<Configure>", lambda _e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas_window = self.canvas.create_window((0, 0), window=self.cards_wrap, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.bind("<Configure>", self.on_canvas_resize)
        self.bind_mousewheel(self.canvas)
        self.bind_mousewheel(self.cards_wrap)

        bottom = tk.Frame(main, bg=BG)
        bottom.pack(fill="x", pady=(12, 0))

        left = tk.Frame(bottom, bg=BG)
        left.pack(side="left", fill="x", expand=True)
        self.progress = ttk.Progressbar(left, orient="horizontal", mode="determinate")
        self.progress.pack(anchor="w", fill="x", expand=True)
        self.progress_label = tk.Label(left, text="待機中", font=(FONT_NAME, 9), bg=BG, fg=ACCENT)
        self.progress_label.pack(anchor="w", pady=(6, 0))

        right = tk.Frame(bottom, bg=BG)
        right.pack(side="right", padx=(14, 0))
        self.cancel_button = ModernButton(right, "キャンセル", self.cancel_task, width=10)
        self.cancel_button.pack(side="left", padx=(0, 10))
        self.merge_button = ModernButton(right, "結合して保存", self.merge_files, primary=True, width=14)
        self.merge_button.pack(side="left")
        self.action_buttons.append(self.merge_button)

        footer = tk.Frame(shell, bg=BG)
        footer.pack(fill="x", padx=24, pady=(0, 10))
        tk.Label(footer, text=COPYRIGHT, font=(FONT_NAME, 9), bg=BG, fg=FOOTER_TEXT, anchor="e").pack(fill="x")

        if HAS_DND:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind("<<Drop>>", self.on_drop)

        self.refresh_merge_cards()
        self.set_processing_state(False)

    def bind_mousewheel(self, widget):
        def _on_mousewheel(event):
            delta = -1 * int(event.delta / 120) if event.delta else 0
            if delta:
                self.canvas.yview_scroll(delta, "units")
        widget.bind("<Enter>", lambda _e: self.root.bind_all("<MouseWheel>", _on_mousewheel))
        widget.bind("<Leave>", lambda _e: self.root.unbind_all("<MouseWheel>"))

    def start_status_animation(self, base_text: str):
        self.stop_status_animation()
        self._status_anim_base = base_text
        self._status_anim_dots = 0
        self._run_status_animation()

    def _run_status_animation(self):
        if not self.worker_running:
            return
        dots = "." * ((self._status_anim_dots % 3) + 1)
        self.progress_label.configure(text=f"{self._status_anim_base}{dots}", fg=ACCENT)
        self._status_anim_dots += 1
        self._status_anim_job = self.root.after(800, self._run_status_animation)

    def stop_status_animation(self):
        if self._status_anim_job is not None:
            try:
                self.root.after_cancel(self._status_anim_job)
            except Exception:
                pass
            self._status_anim_job = None

    def set_processing_state(self, processing: bool):
        for button in self.action_buttons:
            button.set_enabled(not processing)
        self.cancel_button.set_enabled(processing)
        if not processing:
            self.stop_status_animation()
            self.progress_label.configure(fg=ACCENT)

    def choose_folder(self):
        if self.worker_running:
            return
        path = filedialog.askdirectory(initialdir=self.save_folder or default_downloads())
        if path:
            self.save_folder = path
            self.cfg["last_folder"] = path
            save_config(self.cfg)
            self.refresh_status()

    def add_files(self):
        if self.worker_running:
            return
        paths = filedialog.askopenfilenames(filetypes=[("PDFファイル", "*.pdf")])
        self.add_pdf_paths(paths)

    def add_pdf_paths(self, paths):
        added = 0
        for raw in paths:
            path = os.path.abspath(str(raw))
            if not path.lower().endswith(".pdf"):
                continue
            if not os.path.isfile(path):
                continue
            if path in self.files:
                continue
            self.files.append(path)
            added += 1
        if added:
            self.refresh_merge_cards()

    def on_drop(self, event):
        if self.worker_running:
            return
        paths = extract_drop_paths(event.data)
        self.add_pdf_paths(paths)

    def reset_merge(self):
        if self.worker_running:
            messagebox.showinfo(APP_NAME, "処理中はリフレッシュできません。")
            return
        self.files.clear()
        self.page_count_cache.clear()
        self.progress["value"] = 0
        self.progress_label.configure(text="待機中", fg=ACCENT)
        self.refresh_merge_cards()

    def remove_file(self, index: int):
        if self.worker_running:
            return
        if 0 <= index < len(self.files):
            self.files.pop(index)
            self.refresh_merge_cards()

    def move_file(self, index: int, delta: int):
        if self.worker_running:
            return
        new_index = index + delta
        if 0 <= index < len(self.files) and 0 <= new_index < len(self.files):
            self.files[index], self.files[new_index] = self.files[new_index], self.files[index]
            self.refresh_merge_cards()

    def get_page_count(self, path: str) -> int | None:
        if path in self.page_count_cache:
            return self.page_count_cache[path]
        try:
            count = len(PdfReader(path).pages)
            self.page_count_cache[path] = count
            return count
        except Exception:
            return None

    def queue_thumbnail_job(self, path: str):
        def worker():
            payload = None
            page_count = self.get_page_count(path)
            if fitz is not None:
                try:
                    doc = fitz.open(path)
                    page_count = len(doc)
                    page = doc.load_page(0)
                    rect = page.rect
                    target_w, target_h = 160, 220
                    scale = min(target_w / rect.width, target_h / rect.height)
                    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
                    payload = pix.tobytes("ppm")
                    doc.close()
                except Exception:
                    payload = None
            self.thumbnail_queue.put((path, payload, page_count))
        threading.Thread(target=worker, daemon=True).start()

    def process_thumbnail_queue(self):
        try:
            while True:
                path, payload, page_count = self.thumbnail_queue.get_nowait()
                card = self.merge_card_by_path.get(path)
                if card is None:
                    continue
                if page_count is not None:
                    card.set_page_count(page_count)
                if payload:
                    image = tk.PhotoImage(data=payload)
                    card.set_thumbnail(image)
        except queue.Empty:
            pass
        self.root.after(120, self.process_thumbnail_queue)

    def refresh_status(self):
        count = len(self.files)
        self.status_label.configure(text=f"{count}件選択中" if count else "まだPDFがありません")
        self.folder_short_label.configure(text=f"保存先: {shorten_path(self.save_folder)}")

    def refresh_merge_cards(self):
        for child in self.cards_wrap.winfo_children():
            child.destroy()
        self.merge_card_by_path = {}
        self.refresh_status()

        if not self.files:
            empty = tk.Frame(self.cards_wrap, bg="#FFFFFF", highlightthickness=1, highlightbackground=BORDER)
            empty.pack(fill="x", padx=4, pady=4)
            tk.Label(empty, text="まだPDFがありません", font=(FONT_NAME, 12, "bold"), bg="#FFFFFF", fg=TEXT, pady=18).pack()
            tk.Label(empty, text="「PDFを追加」または画面へのドラッグ＆ドロップで追加できます", font=(FONT_NAME, 10), bg="#FFFFFF", fg=SUBTEXT, pady=4).pack()
            return

        width = max(self.canvas.winfo_width() - 24, 320)
        card_outer = 210
        cols = max(1, width // card_outer)
        while cols > 1 and ((cols * card_outer) + 12) > width:
            cols -= 1
        used = cols * card_outer
        extra = max(0, width - used)
        pad_x = max(4, extra // (cols * 2 + 2) if cols else 4)

        for i, path in enumerate(self.files):
            card = MergeFileCard(self.cards_wrap, self, path, i)
            self.merge_card_by_path[path] = card
            card.grid(row=i // cols, column=i % cols, padx=pad_x, pady=8, sticky="n")
            card.update_visual()
            self.queue_thumbnail_job(path)

    def on_canvas_resize(self, _event=None):
        try:
            width = max(self.canvas.winfo_width() - 2, 200)
            self.canvas.itemconfigure(self.canvas_window, width=width)
        except Exception:
            pass
        self.reflow_cards()

    def reflow_cards(self):
        children = list(self.cards_wrap.winfo_children())
        if not children:
            return
        width = max(self.canvas.winfo_width() - 24, 320)
        card_outer = 210
        cols = max(1, width // card_outer)
        while cols > 1 and ((cols * card_outer) + 12) > width:
            cols -= 1
        used = cols * card_outer
        extra = max(0, width - used)
        pad_x = max(4, extra // (cols * 2 + 2) if cols else 4)
        for i, child in enumerate(children):
            child.grid_forget()
            child.grid(row=i // cols, column=i % cols, padx=pad_x, pady=8, sticky="n")

    def make_output_name(self) -> str:
        from datetime import datetime
        base = os.path.splitext(os.path.basename(self.files[0]))[0]
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        return f"{base}_{ts}_{len(self.files)}files_dake.pdf"

    def merge_files(self):
        if self.worker_running:
            return
        if not self.files:
            messagebox.showwarning(APP_NAME, "PDFを追加してください。")
            return

        folder = self.save_folder or default_downloads()
        try:
            os.makedirs(folder, exist_ok=True)
        except Exception as e:
            messagebox.showerror("保存エラー", f"保存先フォルダを作成できませんでした。\n{e}")
            return

        self.cfg["last_folder"] = folder
        save_config(self.cfg)

        base_name = self.make_output_name()
        output = os.path.join(folder, base_name)
        if os.path.exists(output):
            stem, ext = os.path.splitext(base_name)
            n = 1
            while True:
                candidate = os.path.join(folder, f"{stem}_{n:02d}{ext}")
                if not os.path.exists(candidate):
                    output = candidate
                    break
                n += 1

        self.cancel_requested = False
        self.worker_running = True
        self.set_processing_state(True)
        self.progress["value"] = 0
        self.start_status_animation("結合しています")
        threading.Thread(target=self._merge_worker, args=(output,), daemon=True).start()

    def _merge_worker(self, output: str):
        try:
            writer = PdfWriter()
            total = len(self.files)
            for i, path in enumerate(self.files, start=1):
                if self.cancel_requested:
                    self.root.after(0, self.finish_cancel)
                    return
                try:
                    reader = PdfReader(path)
                    for page in reader.pages:
                        writer.add_page(page)
                except Exception as e:
                    self.root.after(0, lambda p=path, err=e: self.handle_file_error(p, err))
                    return
                progress = int(i / total * 100)
                self.root.after(0, lambda p=progress, i=i, t=total: self.update_progress(p, f"{i} / {t} 件を処理中"))
            with open(output, "wb") as f:
                writer.write(f)
            self.root.after(0, lambda: self.finish_success(output))
        except Exception as e:
            self.root.after(0, lambda err=e: self.finish_error(err))

    def update_progress(self, value: int, text: str):
        self.stop_status_animation()
        self.progress["value"] = value
        self.progress_label.configure(text=text, fg=ACCENT)

    def cancel_task(self):
        if self.worker_running:
            self.cancel_requested = True
            self.start_status_animation("キャンセル中")

    def finish_cancel(self):
        self.worker_running = False
        self.set_processing_state(False)
        self.progress["value"] = 0
        self.progress_label.configure(text="キャンセルしました", fg=ACCENT)

    def finish_success(self, output: str):
        self.worker_running = False
        self.set_processing_state(False)
        self.progress["value"] = 100
        self.progress_label.configure(text="完了しました", fg=ACCENT)
        self.root.lift()
        self.root.focus_force()
        messagebox.showinfo(APP_NAME, "結合が完了しました。", parent=self.root)
        open_folder(os.path.dirname(output))

    def finish_error(self, error: Exception):
        self.worker_running = False
        self.set_processing_state(False)
        self.progress["value"] = 0
        self.progress_label.configure(text="エラーが発生しました", fg=ERROR)
        messagebox.showerror("エラー", f"処理中にエラーが発生しました。\n{error}", parent=self.root)

    def handle_file_error(self, path: str, error: Exception):
        self.worker_running = False
        self.set_processing_state(False)
        self.progress["value"] = 0
        self.progress_label.configure(text="読み込みエラー", fg=ERROR)
        messagebox.showerror("PDF読み込みエラー", f"次のPDFを読み込めませんでした。\n\n{path}\n\n{error}", parent=self.root)


def main():
    set_windows_app_id()
    root = make_root()
    apply_window_icon(root)
    app = DAKEPDFMergeApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
