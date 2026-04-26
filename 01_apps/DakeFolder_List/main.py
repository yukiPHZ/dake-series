# -*- coding: utf-8 -*-
from __future__ import annotations

import ctypes
import os
import queue
import sys
import threading
import time
import webbrowser
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox


APP_NAME = "Dakeフォルダ一覧"
WINDOW_TITLE = "フォルダ一覧"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "brand_series": "シンプルそれDAKEシリーズ",
    "main_title": "フォルダ構成を見る",
    "main_description": "フォルダを選ぶだけで、構成とファイル一覧を表示します。",
    "button_select_folder": "フォルダを選ぶ",
    "button_refresh": "再読み込み",
    "button_copy": "一覧をコピー",
    "button_save_txt": "txt保存",
    "empty_title": "フォルダを選んでください",
    "empty_subtitle": "選択したフォルダを基準に、構成とファイル一覧を表示します。",
    "status_idle": "未選択",
    "status_loading": "読み込み中",
    "status_ready": "表示しました",
    "status_copied": "コピーしました",
    "status_saved": "保存しました",
    "status_error": "エラーが発生しました",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_link_1": "戸建買取査定",
    "footer_link_2": "Instagram",
    "footer_separator": " ｜ ",
    "footer_copyright": COPYRIGHT,
    "dialog_select_folder": "一覧化するフォルダを選んでください",
    "dialog_save_title": "txt保存",
    "dialog_error_title": "確認してください",
    "dialog_copy_error": "クリップボードへコピーできませんでした。",
    "dialog_save_error": "txtファイルを保存できませんでした。",
    "filetype_txt": "テキストファイル",
    "filetype_all": "すべてのファイル",
    "status_loading_count": "読み込み中... {count}件",
    "status_ready_count": "表示しました（フォルダ {dirs} / ファイル {files} / スキップ {skipped}）",
    "status_no_data": "一覧がありません",
    "status_refresh_empty": "先にフォルダを選んでください",
    "path_label_empty": "対象フォルダ：未選択",
    "path_label_selected": "対象フォルダ：{path}",
    "output_target": "対象フォルダ",
    "output_scanned_at": "作成日時",
    "output_summary": "集計",
    "output_folders": "フォルダ",
    "output_files": "ファイル",
    "output_others": "その他",
    "output_skipped": "スキップ",
    "meta_folder": "フォルダ",
    "meta_file": "ファイル",
    "meta_other": "その他",
    "meta_extension": "拡張子",
    "meta_size": "サイズ",
    "meta_updated": "更新",
    "meta_no_extension": "なし",
    "meta_unknown": "不明",
    "tree_access_denied": "アクセスできないためスキップ",
    "tree_stat_error": "情報を取得できないためスキップ",
    "error_scan_failed": "フォルダの読み込み中に問題が発生しました。別のフォルダでもお試しください。",
}

COLORS = {
    "base_bg": "#F6F7F9",
    "card_bg": "#FFFFFF",
    "text": "#1E2430",
    "sub_text": "#667085",
    "border": "#E6EAF0",
    "accent": "#2F6FED",
    "accent_hover": "#2458BF",
    "button_sub_hover": "#F2F4F7",
    "text_bg": "#FBFCFE",
    "disabled_bg": "#E8ECF3",
    "disabled_fg": "#98A2B3",
    "error": "#D92D20",
}

LINK_URLS = {
    "footer_link_1": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "footer_link_2": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

FONT_CANDIDATES = ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo")
TEXT_FONT_CANDIDATES = ("BIZ UDGothic", "BIZ UDPGothic", "Consolas", "Yu Gothic UI", "Meiryo")
ICON_RELATIVE_PATH = Path("..") / ".." / "02_assets" / "dake_icon.ico"


@dataclass
class ScanStats:
    dirs: int = 0
    files: int = 0
    others: int = 0
    skipped: int = 0

    @property
    def scanned_count(self) -> int:
        return self.dirs + self.files + self.others + self.skipped


@dataclass
class EntryInfo:
    name: str
    path: Path
    kind: str
    size: int | None
    mtime: float | None
    stat_ok: bool


def set_windows_app_id() -> None:
    if not sys.platform.startswith("win"):
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Shimarisu.DakeFolderList")
    except Exception:
        return


def app_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


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
        pass
    try:
        root.iconbitmap(default=str(icon_path))
    except tk.TclError:
        pass


def choose_font_family(root: tk.Tk, candidates: tuple[str, ...]) -> str:
    try:
        available = set(tkfont.families(root))
    except tk.TclError:
        available = set()
    for candidate in candidates:
        if candidate in available:
            return candidate
    return "TkDefaultFont"


def format_mtime(timestamp: float | None) -> str:
    if timestamp is None:
        return UI_TEXT["meta_unknown"]
    try:
        return datetime.fromtimestamp(timestamp).strftime("%Y-%m-%d %H:%M:%S")
    except (OSError, OverflowError, ValueError):
        return UI_TEXT["meta_unknown"]


def format_size(size: int | None) -> str:
    if size is None:
        return UI_TEXT["meta_unknown"]
    value = float(size)
    units = ("B", "KB", "MB", "GB", "TB")
    unit = units[0]
    for unit in units:
        if value < 1024 or unit == units[-1]:
            break
        value /= 1024
    if unit == "B":
        return f"{int(value)} {unit}"
    return f"{value:.1f} {unit}"


def safe_suffix(name: str) -> str:
    suffix = Path(name).suffix.lower()
    return suffix if suffix else UI_TEXT["meta_no_extension"]


def read_entry_info(entry: os.DirEntry[str]) -> EntryInfo:
    path = Path(entry.path)
    try:
        is_dir = entry.is_dir(follow_symlinks=False)
        is_file = entry.is_file(follow_symlinks=False)
        stat_result = entry.stat(follow_symlinks=False)
    except OSError:
        return EntryInfo(entry.name, path, "error", None, None, False)

    if is_dir:
        kind = "dir"
        size = None
    elif is_file:
        kind = "file"
        size = stat_result.st_size
    else:
        kind = "other"
        size = stat_result.st_size
    return EntryInfo(entry.name, path, kind, size, stat_result.st_mtime, True)


def entry_sort_key(entry: EntryInfo) -> tuple[int, str]:
    order = {"dir": 0, "file": 1, "other": 2, "error": 3}
    return order.get(entry.kind, 3), entry.name.lower()


def make_directory_line(name: str, mtime: float | None) -> str:
    return (
        f"{name}/  "
        f"[{UI_TEXT['meta_folder']} | "
        f"{UI_TEXT['meta_updated']}: {format_mtime(mtime)}]"
    )


def make_file_line(entry: EntryInfo) -> str:
    return (
        f"{entry.name}  "
        f"[{UI_TEXT['meta_extension']}: {safe_suffix(entry.name)} | "
        f"{UI_TEXT['meta_size']}: {format_size(entry.size)} | "
        f"{UI_TEXT['meta_updated']}: {format_mtime(entry.mtime)}]"
    )


def make_other_line(entry: EntryInfo) -> str:
    return (
        f"{entry.name}  "
        f"[{UI_TEXT['meta_other']} | "
        f"{UI_TEXT['meta_size']}: {format_size(entry.size)} | "
        f"{UI_TEXT['meta_updated']}: {format_mtime(entry.mtime)}]"
    )


def build_folder_listing(folder: Path, progress_callback) -> tuple[str, ScanStats]:
    stats = ScanStats()
    root = folder.resolve()
    root_mtime = None
    try:
        root_mtime = root.stat().st_mtime
    except OSError:
        pass

    stats.dirs = 1
    tree_lines = [make_directory_line(root.name or str(root), root_mtime)]
    last_progress = 0.0

    def push_progress(force: bool = False) -> None:
        nonlocal last_progress
        now = time.monotonic()
        if force or now - last_progress >= 0.2:
            last_progress = now
            progress_callback(stats.scanned_count)

    def scan_dir(current: Path, prefix: str) -> None:
        try:
            with os.scandir(current) as iterator:
                entries = [read_entry_info(entry) for entry in iterator]
        except OSError:
            stats.skipped += 1
            tree_lines.append(f"{prefix}└─ {UI_TEXT['tree_access_denied']}")
            push_progress(True)
            return

        entries.sort(key=entry_sort_key)
        for index, entry in enumerate(entries):
            connector = "└─ " if index == len(entries) - 1 else "├─ "
            child_prefix = "   " if index == len(entries) - 1 else "│  "
            line_prefix = f"{prefix}{connector}"

            if not entry.stat_ok:
                stats.skipped += 1
                tree_lines.append(f"{line_prefix}{entry.name}  [{UI_TEXT['tree_stat_error']}]")
                push_progress()
                continue

            if entry.kind == "dir":
                stats.dirs += 1
                tree_lines.append(f"{line_prefix}{make_directory_line(entry.name, entry.mtime)}")
                push_progress()
                scan_dir(entry.path, f"{prefix}{child_prefix}")
            elif entry.kind == "file":
                stats.files += 1
                tree_lines.append(f"{line_prefix}{make_file_line(entry)}")
                push_progress()
            else:
                stats.others += 1
                tree_lines.append(f"{line_prefix}{make_other_line(entry)}")
                push_progress()

    scan_dir(root, "")
    push_progress(True)

    scanned_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    summary = (
        f"{UI_TEXT['output_summary']}: "
        f"{UI_TEXT['output_folders']} {stats.dirs} / "
        f"{UI_TEXT['output_files']} {stats.files} / "
        f"{UI_TEXT['output_others']} {stats.others} / "
        f"{UI_TEXT['output_skipped']} {stats.skipped}"
    )
    header_lines = [
        f"{UI_TEXT['output_target']}: {root}",
        f"{UI_TEXT['output_scanned_at']}: {scanned_at}",
        summary,
        "",
    ]
    return "\n".join(header_lines + tree_lines), stats


class FolderListApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.geometry("1040x760")
        self.root.minsize(860, 620)
        self.root.configure(bg=COLORS["base_bg"])
        apply_window_icon(self.root)

        self.font_family = choose_font_family(self.root, FONT_CANDIDATES)
        self.text_font_family = choose_font_family(self.root, TEXT_FONT_CANDIDATES)
        self.fonts = {
            "brand": (self.font_family, 9),
            "title": (self.font_family, 22, "bold"),
            "subtitle": (self.font_family, 10),
            "button": (self.font_family, 10, "bold"),
            "body": (self.font_family, 10),
            "small": (self.font_family, 9),
            "tree": (self.text_font_family, 10),
        }

        self.current_folder: Path | None = None
        self.current_listing = ""
        self.scan_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.scanning = False
        self.status_var = tk.StringVar(value=UI_TEXT["status_idle"])
        self.path_var = tk.StringVar(value=UI_TEXT["path_label_empty"])
        self.buttons: list[tk.Button] = []

        self._build_ui()
        self._set_empty_text()
        self._sync_buttons()
        self.root.after(80, self._process_scan_queue)

    def _build_ui(self) -> None:
        outer = tk.Frame(self.root, bg=COLORS["base_bg"])
        outer.pack(fill="both", expand=True, padx=24, pady=(22, 12))

        self._build_header(outer)
        self._build_action_card(outer)
        self._build_text_card(outer)
        self._build_status(outer)
        self._build_footer(outer)

    def _build_header(self, parent: tk.Widget) -> None:
        header = tk.Frame(parent, bg=COLORS["base_bg"])
        header.pack(fill="x")

        tk.Label(
            header,
            text=UI_TEXT["brand_series"],
            font=self.fonts["brand"],
            fg=COLORS["accent"],
            bg=COLORS["base_bg"],
            anchor="w",
        ).pack(anchor="w")
        tk.Label(
            header,
            text=UI_TEXT["main_title"],
            font=self.fonts["title"],
            fg=COLORS["text"],
            bg=COLORS["base_bg"],
            anchor="w",
        ).pack(anchor="w", pady=(4, 0))
        tk.Label(
            header,
            text=UI_TEXT["main_description"],
            font=self.fonts["subtitle"],
            fg=COLORS["sub_text"],
            bg=COLORS["base_bg"],
            anchor="w",
        ).pack(anchor="w", pady=(6, 0))

    def _create_card(self, parent: tk.Widget, pady: tuple[int, int] | int) -> tk.Frame:
        card = tk.Frame(parent, bg=COLORS["card_bg"], highlightbackground=COLORS["border"], highlightthickness=1)
        card.pack(fill="x", pady=pady)
        inner = tk.Frame(card, bg=COLORS["card_bg"])
        inner.pack(fill="both", expand=True, padx=18, pady=16)
        return inner

    def _build_action_card(self, parent: tk.Widget) -> None:
        card = self._create_card(parent, pady=(18, 12))

        button_row = tk.Frame(card, bg=COLORS["card_bg"])
        button_row.pack(fill="x")
        self.select_button = self._create_button(button_row, UI_TEXT["button_select_folder"], self.select_folder, True)
        self.refresh_button = self._create_button(button_row, UI_TEXT["button_refresh"], self.refresh_folder, False)
        self.copy_button = self._create_button(button_row, UI_TEXT["button_copy"], self.copy_listing, False)
        self.save_button = self._create_button(button_row, UI_TEXT["button_save_txt"], self.save_listing, False)

        self.select_button.pack(side="left")
        self.refresh_button.pack(side="left", padx=(8, 0))
        self.copy_button.pack(side="left", padx=(8, 0))
        self.save_button.pack(side="left", padx=(8, 0))

        tk.Label(
            card,
            textvariable=self.path_var,
            font=self.fonts["small"],
            fg=COLORS["sub_text"],
            bg=COLORS["card_bg"],
            anchor="w",
        ).pack(fill="x", pady=(12, 0))

    def _create_button(self, parent: tk.Widget, label: str, command, primary: bool) -> tk.Button:
        normal_bg = COLORS["accent"] if primary else COLORS["card_bg"]
        normal_fg = "#FFFFFF" if primary else COLORS["text"]
        hover_bg = COLORS["accent_hover"] if primary else COLORS["button_sub_hover"]
        active_bg = hover_bg
        button = tk.Button(
            parent,
            text=label,
            command=command,
            font=self.fonts["button"],
            fg=normal_fg,
            bg=normal_bg,
            activeforeground=normal_fg,
            activebackground=active_bg,
            disabledforeground=COLORS["disabled_fg"],
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
        button.bind("<Enter>", self._button_enter)
        button.bind("<Leave>", self._button_leave)
        self.buttons.append(button)
        return button

    def _button_enter(self, event) -> None:
        button = event.widget
        if str(button.cget("state")) == "normal":
            button.configure(bg=button._hover_bg)

    def _button_leave(self, event) -> None:
        button = event.widget
        if str(button.cget("state")) == "normal":
            button.configure(bg=button._normal_bg)

    def _build_text_card(self, parent: tk.Widget) -> None:
        card = tk.Frame(parent, bg=COLORS["card_bg"], highlightbackground=COLORS["border"], highlightthickness=1)
        card.pack(fill="both", expand=True)
        card.grid_rowconfigure(0, weight=1)
        card.grid_columnconfigure(0, weight=1)

        self.text = tk.Text(
            card,
            wrap="none",
            undo=False,
            font=self.fonts["tree"],
            fg=COLORS["text"],
            bg=COLORS["text_bg"],
            insertbackground=COLORS["text"],
            selectbackground="#D8E6FF",
            relief="flat",
            padx=14,
            pady=12,
            state="disabled",
        )
        y_scroll = tk.Scrollbar(card, orient="vertical", command=self.text.yview)
        x_scroll = tk.Scrollbar(card, orient="horizontal", command=self.text.xview)
        self.text.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        self.text.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

    def _build_status(self, parent: tk.Widget) -> None:
        status = tk.Label(
            parent,
            textvariable=self.status_var,
            font=self.fonts["small"],
            fg=COLORS["sub_text"],
            bg=COLORS["base_bg"],
            anchor="w",
        )
        status.pack(fill="x", pady=(10, 0))

    def _build_footer(self, parent: tk.Widget) -> None:
        footer = tk.Frame(parent, bg=COLORS["base_bg"])
        footer.pack(fill="x", pady=(14, 0))
        tk.Label(
            footer,
            text=UI_TEXT["footer_left"],
            font=self.fonts["small"],
            fg=COLORS["sub_text"],
            bg=COLORS["base_bg"],
        ).pack(side="left")

        right = tk.Frame(footer, bg=COLORS["base_bg"])
        right.pack(side="right")
        self._footer_link(right, UI_TEXT["footer_link_1"], LINK_URLS["footer_link_1"])
        self._footer_label(right, UI_TEXT["footer_separator"])
        self._footer_link(right, UI_TEXT["footer_link_2"], LINK_URLS["footer_link_2"])
        self._footer_label(right, UI_TEXT["footer_separator"])
        self._footer_label(right, UI_TEXT["footer_copyright"])

    def _footer_label(self, parent: tk.Widget, text: str) -> None:
        tk.Label(parent, text=text, font=self.fonts["small"], fg=COLORS["sub_text"], bg=COLORS["base_bg"]).pack(side="left")

    def _footer_link(self, parent: tk.Widget, text: str, url: str) -> None:
        label = tk.Label(parent, text=text, font=self.fonts["small"], fg=COLORS["accent"], bg=COLORS["base_bg"], cursor="hand2")
        label.pack(side="left")
        label.bind("<Button-1>", lambda _event: webbrowser.open(url))

    def _set_text_content(self, content: str) -> None:
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.insert("1.0", content)
        self.text.configure(state="disabled")

    def _set_empty_text(self) -> None:
        self._set_text_content(f"{UI_TEXT['empty_title']}\n\n{UI_TEXT['empty_subtitle']}")

    def _sync_buttons(self) -> None:
        if self.scanning:
            for button in self.buttons:
                button.configure(state="disabled", bg=COLORS["disabled_bg"], cursor="arrow")
            return

        self.select_button.configure(state="normal", bg=self.select_button._normal_bg, cursor="hand2")
        refresh_state = "normal" if self.current_folder is not None else "disabled"
        data_state = "normal" if self.current_listing else "disabled"
        self.refresh_button.configure(state=refresh_state)
        self.copy_button.configure(state=data_state)
        self.save_button.configure(state=data_state)
        for button in (self.refresh_button, self.copy_button, self.save_button):
            if str(button.cget("state")) == "normal":
                button.configure(bg=button._normal_bg, cursor="hand2")
            else:
                button.configure(bg=COLORS["disabled_bg"], cursor="arrow")

    def select_folder(self) -> None:
        if self.scanning:
            return
        initial = str(self.current_folder or Path.home())
        selected = filedialog.askdirectory(title=UI_TEXT["dialog_select_folder"], initialdir=initial)
        if selected:
            self.start_scan(Path(selected))

    def refresh_folder(self) -> None:
        if self.scanning:
            return
        if self.current_folder is None:
            self.status_var.set(UI_TEXT["status_refresh_empty"])
            return
        self.start_scan(self.current_folder)

    def start_scan(self, folder: Path) -> None:
        self.current_folder = folder
        self.current_listing = ""
        self.scanning = True
        self.path_var.set(UI_TEXT["path_label_selected"].format(path=str(folder)))
        self.status_var.set(f"{UI_TEXT['status_loading']}...")
        self._set_text_content(f"{UI_TEXT['status_loading']}...")
        self._sync_buttons()
        worker = threading.Thread(target=self._scan_worker, args=(folder,), daemon=True)
        worker.start()

    def _scan_worker(self, folder: Path) -> None:
        def progress(count: int) -> None:
            self.scan_queue.put(("progress", count))

        try:
            listing, stats = build_folder_listing(folder, progress)
            self.scan_queue.put(("done", (listing, stats)))
        except Exception as error:
            self.scan_queue.put(("error", str(error)))

    def _process_scan_queue(self) -> None:
        try:
            while True:
                event, payload = self.scan_queue.get_nowait()
                if event == "progress":
                    self.status_var.set(UI_TEXT["status_loading_count"].format(count=payload))
                elif event == "done":
                    listing, stats = payload
                    self._finish_scan(str(listing), stats)
                elif event == "error":
                    self._finish_scan_error(str(payload))
        except queue.Empty:
            pass
        self.root.after(80, self._process_scan_queue)

    def _finish_scan(self, listing: str, stats: ScanStats) -> None:
        self.current_listing = listing
        self.scanning = False
        self._set_text_content(listing)
        self.status_var.set(
            UI_TEXT["status_ready_count"].format(
                dirs=stats.dirs,
                files=stats.files,
                skipped=stats.skipped,
            )
        )
        self._sync_buttons()

    def _finish_scan_error(self, detail: str) -> None:
        self.current_listing = ""
        self.scanning = False
        self.status_var.set(UI_TEXT["status_error"])
        self._set_text_content(f"{UI_TEXT['error_scan_failed']}\n\n{detail}")
        self._sync_buttons()

    def copy_listing(self) -> None:
        if not self.current_listing:
            self.status_var.set(UI_TEXT["status_no_data"])
            return
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(self.current_listing)
            self.root.update_idletasks()
        except tk.TclError:
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["dialog_copy_error"], parent=self.root)
            self.status_var.set(UI_TEXT["status_error"])
            return
        self.status_var.set(UI_TEXT["status_copied"])

    def save_listing(self) -> None:
        if not self.current_listing:
            self.status_var.set(UI_TEXT["status_no_data"])
            return
        initial_name = f"folder_list_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        output_path = filedialog.asksaveasfilename(
            title=UI_TEXT["dialog_save_title"],
            initialfile=initial_name,
            defaultextension=".txt",
            filetypes=((UI_TEXT["filetype_txt"], "*.txt"), (UI_TEXT["filetype_all"], "*.*")),
        )
        if not output_path:
            return
        try:
            with open(output_path, "w", encoding="utf-8", newline="\n") as file:
                file.write(self.current_listing)
        except OSError:
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["dialog_save_error"], parent=self.root)
            self.status_var.set(UI_TEXT["status_error"])
            return
        self.status_var.set(UI_TEXT["status_saved"])


def main() -> None:
    set_windows_app_id()
    root = tk.Tk()
    FolderListApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
