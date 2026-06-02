# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import json
import os
import queue
import re
import subprocess
import sys
import threading
import webbrowser
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tkinter import font as tkfont
from tkinter import messagebox, ttk
from urllib.parse import urlparse
import tkinter as tk

try:
    import ctypes
except Exception:
    ctypes = None


APP_NAME = "DAKE Web Index"
APP_ID = "dake.web.index"
APP_FOLDER_NAME = "DAKE_Web_Index"
DIST_DIR_NAME = "dist"
README_NAME = "README.md"
DAKE_WEB_META_NAME = "DAKE_WEB_META"
DEFAULT_DEV_ROOT = Path(os.environ.get("DAKE_WEB_INDEX_ROOT", os.environ.get("DAKE_WEB_DASHBOARD_ROOT", r"C:\Users\yukiz\devlop")))
DEFAULT_SERIES_ROOT = Path(os.environ.get("DAKE_SERIES_ROOT", r"C:\Users\yukiz\devlop\DAKE_series"))
MAX_SCAN_DEPTH = 2
WORKER_POLL_MS = 80
AUTO_RELOAD_MS = 5 * 60 * 1000

UI_TEXT = {
    "window_title": "DAKE Web Index",
    "header_title": "Webサイト索引",
    "header_subtitle": "README正本から自動生成します",
    "button_reload": "再読込",
    "column_folder_name": "フォルダ名",
    "column_site_name": "サイト名",
    "column_url": "URL",
    "column_github": "GitHub",
    "column_readme": "README",
    "column_updated": "最終更新日",
    "status_waiting": "未読込",
    "status_loading": "読込中",
    "status_loaded": "{count}件 / {time}",
    "status_error": "読込できませんでした",
    "status_no_selection": "行を選択してください",
    "auto_reload_interval": "自動更新: 5分ごと",
    "status_folder_missing": "フォルダを開けません: {path}",
    "status_folder_opened": "フォルダを開きました: {folder}",
    "value_unset": "未設定",
    "value_readme": "README.md",
    "dialog_error_title": "エラー",
    "dialog_notice_title": "確認",
    "dialog_url_missing": "URL が設定されていません。",
    "dialog_url_invalid": "http/https URL ではありません。",
    "dialog_path_missing": "README が見つかりません。\n\n{path}",
    "dialog_open_error": "開けませんでした。\n\n{target}\n\n{error}",
    "launch_check_ok": "LAUNCH CHECK OK",
    "open_check_ok": "OPEN CHECK OK",
}

THEME = {
    "bg": "#F7F8FA",
    "surface": "#FFFFFF",
    "border": "#DADDE3",
    "text": "#1D2433",
    "muted": "#667085",
    "accent": "#2657B8",
    "accent_hover": "#1F4CA6",
    "selection": "#E8EEF9",
    "scrollbar": "#C8CDD6",
    "scrollbar_hover": "#B8BFCA",
}

FONT_CANDIDATES = ["BIZ UDPGothic", "Yu Gothic UI", "Meiryo", "MS Gothic"]
EXCLUDED_DIR_NAMES = {
    ".git",
    ".next",
    ".venv",
    ".wrangler",
    "__pycache__",
    "build",
    "dist",
    "node_modules",
    "venv",
}

DAKE_WEB_META_PATTERN = re.compile(
    r"##\s*DAKE_WEB_META\b.*?```(?:json)?\s*(\{.*?\})\s*```",
    re.IGNORECASE | re.DOTALL,
)
URL_PATTERN = re.compile(r"https?://[^\s)>\]\"']+")
GITHUB_PATTERN = re.compile(r"https?://github\.com/[^\s)>\]\"']+", re.IGNORECASE)


@dataclass(frozen=True)
class SiteRecord:
    folder_name: str
    site_name: str
    url: str
    github_url: str
    readme_path: Path
    folder_path: Path
    last_modified: float

    def updated_text(self) -> str:
        return format_timestamp(self.last_modified)

    def to_json(self) -> dict[str, str]:
        return {
            "folder_name": self.folder_name,
            "site_name": self.site_name,
            "url": self.url,
            "github": self.github_url,
            "readme": str(self.readme_path),
            "last_updated": self.updated_text(),
        }


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        if exe_dir.name.lower() == DIST_DIR_NAME:
            return exe_dir.parent
        return exe_dir
    return Path(__file__).resolve().parent


def icon_path() -> Path:
    if getattr(sys, "frozen", False):
        bundled = Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent)) / "dake_icon.ico"
        if bundled.exists():
            return bundled
    return DEFAULT_SERIES_ROOT / "02_assets" / "dake_icon.ico"


def set_windows_app_id() -> None:
    if not sys.platform.startswith("win") or ctypes is None:
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
    except Exception:
        pass


def apply_window_icon(window: tk.Misc) -> None:
    try:
        icon = icon_path()
        if icon.exists():
            window.iconbitmap(str(icon))
            window.iconbitmap(default=str(icon))
    except Exception:
        pass


def choose_font_family(root: tk.Tk) -> str:
    try:
        available = set(tkfont.families(root))
    except Exception:
        return "TkDefaultFont"
    for family in FONT_CANDIDATES:
        if family in available:
            return family
    return "TkDefaultFont"


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError:
        return path.read_text(encoding="utf-8", errors="replace")


def safe_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def first_text(*values: object) -> str:
    for value in values:
        text = safe_text(value)
        if text:
            return text
    return ""


def load_json_object(text: str) -> dict[str, object]:
    try:
        loaded = json.loads(text)
    except json.JSONDecodeError:
        return {}
    return loaded if isinstance(loaded, dict) else {}


def extract_meta_block(text: str) -> dict[str, object]:
    match = DAKE_WEB_META_PATTERN.search(text)
    if not match:
        return {}
    return load_json_object(match.group(1))


def load_site_meta(folder: Path, readme_text: str) -> tuple[dict[str, object], Path | None]:
    readme_meta = extract_meta_block(readme_text)
    meta_path = folder / DAKE_WEB_META_NAME
    file_meta: dict[str, object] = {}
    if meta_path.exists():
        meta_text = read_text(meta_path)
        file_meta = load_json_object(meta_text) or extract_meta_block(meta_text)
    if readme_meta and file_meta:
        merged = dict(file_meta)
        merged.update(readme_meta)
        return merged, meta_path
    if readme_meta:
        return readme_meta, None
    if file_meta:
        return file_meta, meta_path
    return {}, meta_path if meta_path.exists() else None


def normalize_http_url(value: str) -> str:
    text = safe_text(value).rstrip(".,;")
    if not text:
        return ""
    if re.search(r"\s", text):
        return ""
    if "://" not in text:
        text = f"https://{text}"
    parsed = urlparse(text)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        return ""
    return text


def first_http_url(text: str) -> str:
    for match in URL_PATTERN.finditer(text):
        value = normalize_http_url(match.group(0))
        if value and "github.com" not in value.lower():
            return value
    return ""


def first_github_url(text: str) -> str:
    match = GITHUB_PATTERN.search(text)
    return normalize_http_url(match.group(0)) if match else ""


def url_from_meta(meta: dict[str, object], readme_text: str) -> str:
    value = first_text(
        meta.get("production_url"),
        meta.get("url"),
        meta.get("site_url"),
        meta.get("canonical_url"),
        meta.get("domain"),
    )
    return normalize_http_url(value) or first_http_url(readme_text)


def run_hidden_command(args: list[str], cwd: Path | None = None, timeout: float | None = None) -> subprocess.CompletedProcess[str]:
    kwargs: dict[str, object] = {
        "cwd": str(cwd) if cwd is not None else None,
        "capture_output": True,
        "text": True,
        "encoding": "utf-8",
        "errors": "replace",
        "timeout": timeout,
        "check": False,
    }
    if sys.platform.startswith("win"):
        kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
    return subprocess.run(args, **kwargs)


def github_from_git_remote(folder: Path) -> str:
    if not (folder / ".git").exists():
        return ""
    try:
        result = run_hidden_command(
            ["git", "-C", str(folder), "config", "--get", "remote.origin.url"],
            timeout=2.0,
        )
    except Exception:
        return ""
    return normalize_github_remote(result.stdout.strip())


def normalize_github_remote(value: str) -> str:
    text = safe_text(value)
    if not text:
        return ""
    if text.startswith("git@github.com:"):
        path = text.removeprefix("git@github.com:").removesuffix(".git")
        return normalize_http_url(f"https://github.com/{path}")
    if text.startswith("https://github.com/") or text.startswith("http://github.com/"):
        return normalize_http_url(text.removesuffix(".git"))
    return ""


def github_from_meta(meta: dict[str, object], readme_text: str, folder: Path) -> str:
    value = first_text(
        meta.get("github_url"),
        meta.get("repo_url"),
        meta.get("repository_url"),
        meta.get("git_url"),
    )
    return normalize_http_url(value) or first_github_url(readme_text) or github_from_git_remote(folder)


def site_name_from_meta(meta: dict[str, object], folder: Path) -> str:
    return first_text(
        meta.get("display_name"),
        meta.get("site_title"),
        meta.get("name"),
        meta.get("project_name"),
        meta.get("folder_name"),
        folder.name,
    )


def latest_mtime(*paths: Path | None) -> float:
    values: list[float] = []
    for path in paths:
        if path is None:
            continue
        try:
            values.append(path.stat().st_mtime)
        except OSError:
            pass
    return max(values) if values else 0.0


def format_timestamp(timestamp: float) -> str:
    if not timestamp:
        return UI_TEXT["value_unset"]
    return datetime.fromtimestamp(timestamp).strftime("%Y-%m-%d %H:%M")


def has_site_signal(folder: Path, meta: dict[str, object], readme_text: str) -> bool:
    if meta:
        return True
    lower_name = folder.name.lower()
    if "site" in lower_name and (folder / README_NAME).exists():
        return True
    if (folder / "wrangler.toml").exists():
        return True
    if (folder / "package.json").exists() and ((folder / "public").exists() or (folder / "functions").exists()):
        return True
    return False


def iter_scan_dirs(root: Path, max_depth: int = MAX_SCAN_DEPTH):
    if not root.exists():
        return
    stack: list[tuple[Path, int]] = [(root, 0)]
    while stack:
        folder, depth = stack.pop()
        yield folder
        if depth >= max_depth:
            continue
        try:
            children = sorted((child for child in folder.iterdir() if child.is_dir()), key=lambda item: item.name.lower())
        except OSError:
            continue
        for child in reversed(children):
            if child.name.lower() in EXCLUDED_DIR_NAMES:
                continue
            stack.append((child, depth + 1))


def build_record(folder: Path) -> SiteRecord | None:
    readme_path = folder / README_NAME
    if not readme_path.exists():
        return None
    try:
        readme_text = read_text(readme_path)
    except OSError:
        return None
    meta, meta_path = load_site_meta(folder, readme_text)
    if not has_site_signal(folder, meta, readme_text):
        return None
    return SiteRecord(
        folder_name=folder.name,
        site_name=site_name_from_meta(meta, folder),
        url=url_from_meta(meta, readme_text),
        github_url=github_from_meta(meta, readme_text, folder),
        readme_path=readme_path,
        folder_path=folder,
        last_modified=latest_mtime(readme_path, meta_path),
    )


def scan_sites(root: Path) -> list[SiteRecord]:
    records: list[SiteRecord] = []
    seen: set[Path] = set()
    for folder in iter_scan_dirs(root):
        resolved = folder.resolve()
        if resolved in seen:
            continue
        seen.add(resolved)
        record = build_record(folder)
        if record is not None:
            records.append(record)
    records.sort(key=lambda item: (item.folder_name.lower(), item.site_name.lower()))
    return records


def open_url_external(url: str, opener=webbrowser.open) -> bool:
    normalized = normalize_http_url(url)
    if not normalized:
        return False
    opener(normalized)
    return True


def open_path_external(path: Path, starter=None) -> bool:
    if not path.exists():
        return False
    if starter is not None:
        starter(path)
        return True
    if sys.platform.startswith("win"):
        os.startfile(str(path))
    else:
        webbrowser.open(path.as_uri())
    return True


class AutoHideScrollbar(ttk.Scrollbar):
    def set(self, first: str, last: str) -> None:
        if float(first) <= 0.0 and float(last) >= 1.0:
            self.grid_remove()
        else:
            self.grid()
        super().set(first, last)


class WebIndexApp:
    def __init__(self, root: tk.Tk, root_path: Path) -> None:
        self.root = root
        self.root_path = root_path
        self.font_family = choose_font_family(root)
        self.records: list[SiteRecord] = []
        self.record_by_iid: dict[str, SiteRecord] = {}
        self.worker_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.worker_thread: threading.Thread | None = None
        self.auto_reload_job: str | None = None
        self.sort_column = "folder_name"
        self.sort_reverse = False
        self.status_var = tk.StringVar(value=UI_TEXT["status_waiting"])
        self.auto_status_var = tk.StringVar(value=UI_TEXT["auto_reload_interval"])

        self.root.title(UI_TEXT["window_title"])
        self.root.geometry("1120x680")
        self.root.minsize(880, 520)
        self.root.configure(bg=THEME["bg"])
        apply_window_icon(root)
        self.configure_style()
        self.build_ui()
        self.refresh(reset_auto_timer=True)

    def configure_style(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(
            "Index.Treeview",
            background=THEME["surface"],
            fieldbackground=THEME["surface"],
            foreground=THEME["text"],
            rowheight=29,
            borderwidth=0,
            font=(self.font_family, 10),
        )
        style.configure(
            "Index.Treeview.Heading",
            background=THEME["bg"],
            foreground=THEME["muted"],
            relief="flat",
            font=(self.font_family, 9, "bold"),
        )
        style.map("Index.Treeview", background=[("selected", THEME["selection"])], foreground=[("selected", THEME["text"])])
        style.configure(
            "Index.Vertical.TScrollbar",
            gripcount=0,
            width=10,
            background=THEME["scrollbar"],
            darkcolor=THEME["scrollbar"],
            lightcolor=THEME["scrollbar"],
            bordercolor=THEME["bg"],
            troughcolor=THEME["bg"],
            arrowcolor=THEME["muted"],
            relief="flat",
        )
        style.map(
            "Index.Vertical.TScrollbar",
            background=[("active", THEME["scrollbar_hover"])],
            arrowcolor=[("active", THEME["muted"])],
        )

    def build_ui(self) -> None:
        outer = tk.Frame(self.root, bg=THEME["bg"])
        outer.pack(fill="both", expand=True, padx=22, pady=18)

        header = tk.Frame(outer, bg=THEME["bg"])
        header.pack(fill="x", pady=(0, 14))
        tk.Label(header, text=UI_TEXT["header_title"], bg=THEME["bg"], fg=THEME["text"], font=(self.font_family, 24, "bold")).pack(anchor="w")
        tk.Label(header, text=UI_TEXT["header_subtitle"], bg=THEME["bg"], fg=THEME["muted"], font=(self.font_family, 11)).pack(anchor="w", pady=(4, 0))

        table_wrap = tk.Frame(outer, bg=THEME["surface"], highlightthickness=1, highlightbackground=THEME["border"])
        table_wrap.pack(fill="both", expand=True)
        table_wrap.grid_columnconfigure(0, weight=1)
        table_wrap.grid_rowconfigure(0, weight=1)

        columns = ("folder_name", "site_name", "url", "github", "readme", "updated")
        self.tree = ttk.Treeview(table_wrap, columns=columns, show="headings", selectmode="browse", style="Index.Treeview")
        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scroll = AutoHideScrollbar(table_wrap, orient="vertical", command=self.tree.yview, style="Index.Vertical.TScrollbar")
        y_scroll.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=y_scroll.set)

        headings = {
            "folder_name": UI_TEXT["column_folder_name"],
            "site_name": UI_TEXT["column_site_name"],
            "url": UI_TEXT["column_url"],
            "github": UI_TEXT["column_github"],
            "readme": UI_TEXT["column_readme"],
            "updated": UI_TEXT["column_updated"],
        }
        widths = {
            "folder_name": 220,
            "site_name": 220,
            "url": 340,
            "github": 340,
            "readme": 100,
            "updated": 150,
        }
        self.column_headings = headings
        for column in columns:
            if column in {"folder_name", "site_name", "updated"}:
                self.tree.heading(column, text=headings[column], command=lambda key=column: self.sort_by_column(key))
            else:
                self.tree.heading(column, text=headings[column])
            self.tree.column(column, width=widths[column], minwidth=80, stretch=False, anchor="w")
        self.tree.bind("<Double-Button-1>", self.on_tree_double_click)
        self.tree.bind("<Return>", self.open_selected_url)
        self.tree.bind("<Configure>", self.on_tree_configure)
        self.root.after_idle(self.adjust_column_widths)

        footer = tk.Frame(outer, bg=THEME["bg"])
        footer.pack(fill="x", pady=(12, 0))
        tk.Label(footer, textvariable=self.status_var, bg=THEME["bg"], fg=THEME["muted"], font=(self.font_family, 9)).pack(side="left")
        self.reload_button = tk.Button(
            footer,
            text=UI_TEXT["button_reload"],
            command=lambda: self.refresh(reset_auto_timer=True),
            bg=THEME["accent"],
            fg="#FFFFFF",
            activebackground=THEME["accent_hover"],
            activeforeground="#FFFFFF",
            relief="flat",
            bd=0,
            padx=18,
            pady=8,
            cursor="hand2",
            font=(self.font_family, 9, "bold"),
        )
        self.reload_button.pack(side="right")
        tk.Label(footer, textvariable=self.auto_status_var, bg=THEME["bg"], fg=THEME["muted"], font=(self.font_family, 9)).pack(side="right", padx=(0, 14))

    def on_tree_configure(self, _event=None) -> None:
        self.adjust_column_widths()

    def adjust_column_widths(self) -> None:
        available = max(self.tree.winfo_width() - 4, 820)
        updated_width = 150
        readme_width = 92
        folder_width = 200
        site_width = 200
        flex_width = available - updated_width - readme_width - folder_width - site_width
        if flex_width < 340:
            deficit = 340 - flex_width
            folder_reduce = min(50, deficit // 2)
            site_reduce = min(50, deficit - folder_reduce)
            folder_width -= folder_reduce
            site_width -= site_reduce
            flex_width = available - updated_width - readme_width - folder_width - site_width
        url_width = max(170, flex_width // 2)
        github_width = max(170, flex_width - url_width)
        widths = {
            "folder_name": max(150, folder_width),
            "site_name": max(150, site_width),
            "url": url_width,
            "github": github_width,
            "readme": readme_width,
            "updated": updated_width,
        }
        for column, width in widths.items():
            self.tree.column(column, width=int(width))

    def sort_by_column(self, column: str) -> None:
        if column == self.sort_column:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = column
            self.sort_reverse = column == "updated"
        self.render_records()

    def sorted_records(self, records: list[SiteRecord]) -> list[SiteRecord]:
        if self.sort_column == "updated":
            return sorted(records, key=lambda record: record.last_modified, reverse=self.sort_reverse)
        if self.sort_column == "site_name":
            return sorted(records, key=lambda record: (record.site_name.lower(), record.folder_name.lower()), reverse=self.sort_reverse)
        return sorted(records, key=lambda record: (record.folder_name.lower(), record.site_name.lower()), reverse=self.sort_reverse)

    def schedule_auto_reload(self) -> None:
        if self.auto_reload_job is not None:
            self.root.after_cancel(self.auto_reload_job)
        self.auto_reload_job = self.root.after(AUTO_RELOAD_MS, self.on_auto_reload)

    def on_auto_reload(self) -> None:
        self.auto_reload_job = None
        self.refresh()
        self.schedule_auto_reload()

    def refresh(self, reset_auto_timer: bool = False) -> None:
        if reset_auto_timer:
            self.schedule_auto_reload()
        if self.worker_thread and self.worker_thread.is_alive():
            return
        self.status_var.set(UI_TEXT["status_loading"])
        self.reload_button.configure(state="disabled")
        self.worker_thread = threading.Thread(target=self.scan_worker, daemon=True)
        self.worker_thread.start()
        self.root.after(WORKER_POLL_MS, self.poll_worker)

    def scan_worker(self) -> None:
        try:
            records = scan_sites(self.root_path)
            self.worker_queue.put(("records", records))
        except Exception as exc:
            self.worker_queue.put(("error", exc))

    def poll_worker(self) -> None:
        try:
            while True:
                event, payload = self.worker_queue.get_nowait()
                if event == "records" and isinstance(payload, list):
                    self.apply_records(payload)
                elif event == "error":
                    self.status_var.set(UI_TEXT["status_error"])
        except queue.Empty:
            pass
        if self.worker_thread and self.worker_thread.is_alive():
            self.root.after(WORKER_POLL_MS, self.poll_worker)
        else:
            self.reload_button.configure(state="normal")

    def apply_records(self, records: list[SiteRecord]) -> None:
        self.records = records
        self.render_records()
        self.status_var.set(UI_TEXT["status_loaded"].format(count=len(records), time=datetime.now().strftime("%Y-%m-%d %H:%M")))

    def render_records(self) -> None:
        self.record_by_iid.clear()
        self.tree.delete(*self.tree.get_children())
        for index, record in enumerate(self.sorted_records(self.records)):
            iid = str(index)
            self.record_by_iid[iid] = record
            self.tree.insert(
                "",
                "end",
                iid=iid,
                values=(
                    record.folder_name,
                    record.site_name,
                    record.url or UI_TEXT["value_unset"],
                    record.github_url or UI_TEXT["value_unset"],
                    UI_TEXT["value_readme"],
                    record.updated_text(),
                ),
            )

    def selected_record(self) -> SiteRecord | None:
        selection = self.tree.selection()
        if not selection:
            self.status_var.set(UI_TEXT["status_no_selection"])
            return None
        return self.record_by_iid.get(selection[0])

    def on_tree_double_click(self, event) -> None:
        if self.tree.identify("region", event.x, event.y) != "cell":
            return
        row_id = self.tree.identify_row(event.y)
        column_id = self.tree.identify_column(event.x)
        record = self.record_by_iid.get(row_id)
        if record is None:
            return
        if column_id == "#1":
            self.open_folder(record.folder_path)
        elif column_id == "#3":
            self.open_url(record.url)
        elif column_id == "#4":
            self.open_url(record.github_url)
        elif column_id == "#5":
            self.open_path(record.readme_path)

    def open_selected_url(self, _event=None) -> None:
        record = self.selected_record()
        if record is not None:
            self.open_url(record.url)

    def open_url(self, url: str) -> None:
        if not safe_text(url):
            messagebox.showinfo(UI_TEXT["dialog_notice_title"], UI_TEXT["dialog_url_missing"], parent=self.root)
            return
        try:
            if not open_url_external(url):
                messagebox.showinfo(UI_TEXT["dialog_notice_title"], UI_TEXT["dialog_url_invalid"], parent=self.root)
        except Exception as exc:
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["dialog_open_error"].format(target=url, error=exc), parent=self.root)

    def open_path(self, path: Path) -> None:
        if not path.exists():
            messagebox.showinfo(UI_TEXT["dialog_notice_title"], UI_TEXT["dialog_path_missing"].format(path=path), parent=self.root)
            return
        try:
            open_path_external(path)
        except Exception as exc:
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["dialog_open_error"].format(target=path, error=exc), parent=self.root)

    def open_folder(self, path: Path) -> None:
        if not path.exists():
            self.status_var.set(UI_TEXT["status_folder_missing"].format(path=path))
            return
        try:
            if open_path_external(path):
                self.status_var.set(UI_TEXT["status_folder_opened"].format(folder=path.name))
            else:
                self.status_var.set(UI_TEXT["status_folder_missing"].format(path=path))
        except Exception:
            self.status_var.set(UI_TEXT["status_folder_missing"].format(path=path))


def records_to_payload(records: list[SiteRecord], root_path: Path) -> dict[str, object]:
    return {
        "status": "ok",
        "root": str(root_path),
        "loaded_at": datetime.now().isoformat(timespec="seconds"),
        "count": len(records),
        "records": [record.to_json() for record in records],
    }


def run_reload_api(root_path: Path) -> int:
    records = scan_sites(root_path)
    print(json.dumps(records_to_payload(records, root_path), ensure_ascii=False, indent=2))
    return 0


def run_launch_check(root_path: Path) -> int:
    records = scan_sites(root_path)
    print(f"{UI_TEXT['launch_check_ok']}: sites={len(records)} root={root_path}")
    return 0


def run_open_check(root_path: Path) -> int:
    records = scan_sites(root_path)
    url_record = next((record for record in records if record.url), None)
    github_record = next((record for record in records if record.github_url), None)
    readme_record = next((record for record in records if record.readme_path.exists()), None)
    url_ok = bool(url_record and open_url_external(url_record.url, opener=lambda _url: True))
    github_ok = bool(github_record and open_url_external(github_record.github_url, opener=lambda _url: True))
    readme_ok = bool(readme_record and open_path_external(readme_record.readme_path, starter=lambda _path: True))
    print(f"{UI_TEXT['open_check_ok']}: url={url_ok} github={github_ok} readme={readme_ok}")
    return 0 if url_ok and github_ok and readme_ok else 1


def run_gui(root_path: Path) -> int:
    set_windows_app_id()
    root = tk.Tk()
    WebIndexApp(root, root_path)
    root.mainloop()
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=APP_NAME)
    parser.add_argument("--root", default=str(DEFAULT_DEV_ROOT))
    parser.add_argument("--launch-check", action="store_true")
    parser.add_argument("--reload-api", action="store_true")
    parser.add_argument("--qpsc-reload", action="store_true")
    parser.add_argument("--open-check", action="store_true")
    args = parser.parse_args()
    root_path = Path(args.root)
    if args.launch_check:
        return run_launch_check(root_path)
    if args.reload_api or args.qpsc_reload:
        return run_reload_api(root_path)
    if args.open_check:
        return run_open_check(root_path)
    return run_gui(root_path)


if __name__ == "__main__":
    raise SystemExit(main())
