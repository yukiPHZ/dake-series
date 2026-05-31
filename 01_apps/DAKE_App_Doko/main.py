# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import os
import queue
import re
import subprocess
import sys
import threading
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import font as tkfont, messagebox, ttk

def get_app_base_dir() -> Path:
    try:
        if getattr(sys, "frozen", False):
            return Path(sys.executable).resolve().parent
        return Path(__file__).resolve().parent
    except Exception:
        return Path.cwd()


def get_series_root(base: Path | None = None) -> Path | None:
    try:
        start = base or get_app_base_dir()
        for candidate in (start, *start.parents):
            if candidate.name == "DAKE_series":
                return candidate
            if (candidate / "01_apps").exists() and (candidate / "02_assets").exists():
                return candidate
    except Exception:
        return None
    return None


def get_core_dir() -> Path | None:
    series_root = get_series_root()
    if series_root is None:
        return None
    core_dir = series_root / "00_core"
    return core_dir if core_dir.exists() else None


CORE_DIR = get_core_dir()
if CORE_DIR is not None and str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

from dake_quality_engine import run_launch_check, safe_load_json_config, safe_run, safe_save_json_config
from dake_quality_engine.logging import write_debug_log


APP_KEY = "DAKE_App_Doko"
APP_NAME = "Dakeアプリどこ"
WINDOW_TITLE = APP_NAME
COPYRIGHT = "© 2026 しまリス不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "display_title": "アプリどこ",
    "display_subtitle": "PC内に散らばったDAKE系アプリを探して、忘れていたアプリへ再接続します。",
    "primary_button": "アプリを探す",
    "secondary_button": "最近使っていないアプリ",
    "cancel_button": "キャンセル",
    "search_label": "検索",
    "search_placeholder": "圧縮 / 工程 / PDF / メール",
    "status_not_searched": "未検索",
    "status_searching": "探索中",
    "status_readme": "README確認中",
    "status_completed": "検索完了",
    "status_launched": "起動しました",
    "status_error": "エラー",
    "series_root_missing": "DAKE_series ルート未検出",
    "status_cancel_requested": "キャンセルを受け付けました。途中結果を表示します。",
    "status_cancelled": "キャンセルしました",
    "status_folder_opened": "フォルダを開きました",
    "status_path_copied": "パスをコピーしました",
    "summary_initial": "青いボタンでPC内のDAKE系exeを探索します。",
    "summary_count": "{count}件のアプリが見つかりました",
    "summary_filter_count": "{shown}件を表示中 / {total}件",
    "empty_title": "まだ探索していません",
    "empty_description": "中央のボタンから探索を始めてください。",
    "no_result_title": "該当するアプリがありません",
    "no_result_description": "検索語を変えるか、もう一度探索してください。",
    "readme_yes": "あり",
    "readme_no": "なし",
    "not_launched": "未起動",
    "unknown_description": "README正本が見つからないDAKE系アプリです。",
    "label_description": "説明",
    "label_exe_path": "exeパス",
    "label_folder_path": "フォルダ",
    "label_readme": "README",
    "label_updated": "最終更新",
    "label_launched": "最終起動",
    "button_launch": "起動",
    "button_open_folder": "フォルダを開く",
    "button_copy_path": "パスをコピー",
    "dialog_error_title": "アプリどこ",
    "dialog_launch_error": "起動できませんでした。",
    "dialog_folder_error": "フォルダを開けませんでした。",
    "scan_root_line": "探索対象: {roots}",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_caption": "探して、思い出して、また使う。",
    "footer_link_1": "戸建買取横浜",
    "footer_link_2": "Instagram",
    "footer_separator": "・",
    "footer_copyright": COPYRIGHT,
}

COLORS = {
    "background": "#F6F7F9",
    "card": "#FFFFFF",
    "text": "#1E2430",
    "muted": "#667085",
    "border": "#E6EAF0",
    "accent": "#2F6FED",
    "accent_hover": "#2458BF",
    "surface_alt": "#F9FAFB",
    "success": "#027A48",
    "error": "#B42318",
}

EXCLUDED_DIRS = {"build", "__pycache__", ".git", "node_modules", "venv", ".venv"}
CONFIG_FILE_NAME = "DAKE_App_Doko_config.json"
QUEUE_POLL_INTERVAL_MS = 100
CARD_WRAP_LENGTH = 720


@dataclass(frozen=True)
class AppCandidate:
    identity: str
    display_name: str
    description: str
    exe_name: str
    exe_path: Path
    folder_path: Path
    readme_path: Path | None
    meta: dict[str, object]
    modified_at: datetime
    launch_history_at: str
    discovered_order: int
    search_text: str


APP_DIR = get_app_base_dir()
LOG_DIR = APP_DIR / "logs"
CONFIG_PATH = APP_DIR / CONFIG_FILE_NAME


def get_readme_path() -> Path:
    candidates = [APP_DIR / "README.md", APP_DIR.parent / "README.md"]
    series_root = get_series_root(APP_DIR)
    if series_root is not None:
        candidates.append(series_root / "01_apps" / APP_KEY / "README.md")
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[0]


def get_icon_path() -> Path | None:
    series_root = get_series_root(APP_DIR)
    if series_root is None:
        return None
    icon_path = series_root / "02_assets" / "dake_icon.ico"
    return icon_path if icon_path.exists() else None


def read_text_fallback(path: Path) -> str:
    for encoding in ("utf-8", "utf-8-sig", "cp932"):
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding="utf-8", errors="replace")


def extract_dake_meta(readme_text: str) -> dict[str, object]:
    match = re.search(r"##\s+DAKE_META\s*```json\s*(\{.*?\})\s*```", readme_text, re.DOTALL)
    if not match:
        return {}
    try:
        loaded = json.loads(match.group(1))
    except json.JSONDecodeError:
        return {}
    return loaded if isinstance(loaded, dict) else {}


def read_own_meta() -> dict[str, object]:
    readme_path = get_readme_path()
    if not readme_path.exists():
        write_debug_log("README.md not found", log_dir=LOG_DIR, context={"path": readme_path})
        return {}
    text = read_text_fallback(readme_path)
    meta = extract_dake_meta(text)
    if not meta:
        write_debug_log("DAKE_META not found", log_dir=LOG_DIR, context={"path": readme_path})
    return meta


def find_readme_for_exe(exe_path: Path) -> Path | None:
    for candidate in (exe_path.parent / "README.md", exe_path.parent.parent / "README.md"):
        if candidate.exists() and candidate.is_file():
            return candidate
    return None


def prettify_exe_name(exe_name: str) -> str:
    stem = Path(exe_name).stem
    stem = re.sub(r"^(Dake|DAKE)_?", "", stem)
    return stem.replace("_", " ").strip() or Path(exe_name).stem


def format_datetime(value: datetime | None) -> str:
    if value is None:
        return UI_TEXT["not_launched"]
    return value.strftime("%Y-%m-%d %H:%M")


def parse_history_datetime(value: str) -> datetime | None:
    if not value:
        return None
    try:
        return datetime.fromisoformat(value)
    except ValueError:
        return None


def normalize_identity(path: Path) -> str:
    try:
        return str(path.resolve())
    except OSError:
        return str(path.absolute())


def build_default_search_roots(series_root: Path | None = None) -> tuple[Path, ...]:
    roots: list[Path] = []
    try:
        home = Path.home()
    except Exception:
        home = Path(r"C:\Users\yukiz")

    for candidate in (home / "Downloads", home / "Desktop", home / "Documents", Path("D:/")):
        try:
            if candidate.exists():
                roots.append(candidate)
        except OSError:
            continue

    resolved_series_root = series_root if series_root is not None else get_series_root(APP_DIR)
    if resolved_series_root is not None:
        apps_dir = resolved_series_root / "01_apps"
        try:
            if apps_dir.exists():
                roots.append(apps_dir)
        except OSError:
            pass

    unique_roots: list[Path] = []
    seen: set[str] = set()
    for root in roots:
        identity = normalize_identity(root).lower()
        if identity not in seen:
            seen.add(identity)
            unique_roots.append(root)
    return tuple(unique_roots)


def load_config() -> dict[str, object]:
    config = safe_load_json_config(CONFIG_PATH, {"launch_history": {}, "last_scan_roots": []})
    if not isinstance(config.get("launch_history"), dict):
        config["launch_history"] = {}
    if not isinstance(config.get("last_scan_roots"), list):
        config["last_scan_roots"] = []
    return config


def save_config(config: dict[str, object]) -> None:
    safe_save_json_config(CONFIG_PATH, config)


def build_candidate(exe_path: Path, history: dict[str, str], discovered_order: int) -> AppCandidate:
    readme_path = find_readme_for_exe(exe_path)
    meta: dict[str, object] = {}
    readme_text = ""
    if readme_path is not None:
        readme_text = read_text_fallback(readme_path)
        meta = extract_dake_meta(readme_text)

    exe_name = exe_path.name
    display_name = str(
        meta.get("launcher_title")
        or meta.get("display_name")
        or meta.get("site_title")
        or prettify_exe_name(exe_name)
    )
    description = str(
        meta.get("launcher_description")
        or meta.get("site_description")
        or meta.get("update_summary")
        or UI_TEXT["unknown_description"]
    )
    identity = normalize_identity(exe_path)
    modified_at = datetime.fromtimestamp(exe_path.stat().st_mtime)
    meta_text = json.dumps(meta, ensure_ascii=False, sort_keys=True) if meta else ""
    search_text = "\n".join(
        [
            display_name,
            description,
            exe_name,
            str(exe_path.parent.name),
            str(meta_text),
            readme_text,
        ]
    ).lower()
    return AppCandidate(
        identity=identity,
        display_name=display_name,
        description=description,
        exe_name=exe_name,
        exe_path=exe_path,
        folder_path=exe_path.parent,
        readme_path=readme_path,
        meta=meta,
        modified_at=modified_at,
        launch_history_at=str(history.get(identity, "")),
        discovered_order=discovered_order,
        search_text=search_text,
    )


def is_dake_exe(file_name: str) -> bool:
    lower_name = file_name.lower()
    return lower_name.startswith("dake") and lower_name.endswith(".exe")


def scan_worker(
    output_queue: queue.Queue[dict[str, object]],
    cancel_event: threading.Event,
    roots: tuple[Path, ...],
    history: dict[str, str],
) -> None:
    discovered: set[str] = set()
    order = 0
    existing_roots = [root for root in roots if root.exists()]
    try:
        output_queue.put({"type": "status", "value": UI_TEXT["status_searching"]})
        for root in existing_roots:
            if cancel_event.is_set():
                break
            for current_root, dir_names, file_names in os.walk(root):
                if cancel_event.is_set():
                    break
                dir_names[:] = [name for name in dir_names if name.lower() not in EXCLUDED_DIRS]
                for file_name in file_names:
                    if cancel_event.is_set():
                        break
                    if not is_dake_exe(file_name):
                        continue
                    exe_path = Path(current_root) / file_name
                    identity = normalize_identity(exe_path)
                    if identity.lower() in discovered:
                        continue
                    discovered.add(identity.lower())
                    output_queue.put({"type": "status", "value": UI_TEXT["status_readme"]})
                    try:
                        candidate = build_candidate(exe_path, history, order)
                    except OSError as exc:
                        write_debug_log("candidate read skipped", log_dir=LOG_DIR, exc=exc, context={"path": exe_path})
                        continue
                    order += 1
                    output_queue.put({"type": "found", "candidate": candidate})
        output_queue.put(
            {
                "type": "done",
                "cancelled": cancel_event.is_set(),
                "roots": [str(root) for root in existing_roots],
            }
        )
    except Exception as exc:
        output_queue.put({"type": "error", "error": exc})


class AppDoko:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.geometry("980x680")
        self.root.minsize(860, 560)
        self.root.configure(bg=COLORS["background"])
        self.series_root = get_series_root(APP_DIR)
        if self.series_root is None:
            write_debug_log(UI_TEXT["series_root_missing"], log_dir=LOG_DIR, context={"base": APP_DIR})
        self.search_roots = build_default_search_roots(self.series_root)

        icon_path = get_icon_path()
        if icon_path is not None:
            try:
                self.root.iconbitmap(icon_path)
            except tk.TclError as exc:
                write_debug_log("icon setup skipped", log_dir=LOG_DIR, exc=exc, context={"path": icon_path})

        self.config = load_config()
        self.candidates: list[AppCandidate] = []
        self.filtered_candidates: list[AppCandidate] = []
        self.event_queue: queue.Queue[dict[str, object]] = queue.Queue()
        self.cancel_event = threading.Event()
        self.worker_thread: threading.Thread | None = None
        self.recent_mode = False
        self.search_var = tk.StringVar()
        self.status_var = tk.StringVar(
            value=UI_TEXT["series_root_missing"] if self.series_root is None else UI_TEXT["status_not_searched"]
        )
        self.summary_var = tk.StringVar(value=UI_TEXT["summary_initial"])

        self._configure_fonts()
        self._build_styles()
        self._build_ui()
        self.root.after(QUEUE_POLL_INTERVAL_MS, self._poll_queue)

    def _configure_fonts(self) -> None:
        available = set(tkfont.families(self.root))
        for family in ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo"):
            if family in available:
                self.base_font = family
                break
        else:
            self.base_font = "TkDefaultFont"

    def _build_styles(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(
            "Doko.TEntry",
            fieldbackground=COLORS["card"],
            foreground=COLORS["text"],
            bordercolor=COLORS["border"],
            lightcolor=COLORS["border"],
            darkcolor=COLORS["border"],
            padding=8,
        )
        style.configure("Vertical.TScrollbar", troughcolor=COLORS["background"], background=COLORS["border"])

    def _build_ui(self) -> None:
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(0, weight=1)

        shell = tk.Frame(self.root, bg=COLORS["background"])
        shell.grid(row=0, column=0, sticky="nsew", padx=22, pady=18)
        shell.grid_columnconfigure(0, weight=1)
        shell.grid_rowconfigure(3, weight=1)

        header = tk.Frame(shell, bg=COLORS["background"])
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(0, weight=1)

        tk.Label(
            header,
            text=UI_TEXT["display_title"],
            bg=COLORS["background"],
            fg=COLORS["text"],
            font=(self.base_font, 22, "bold"),
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            header,
            text=UI_TEXT["display_subtitle"],
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=(self.base_font, 10),
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

        status_frame = tk.Frame(header, bg=COLORS["background"])
        status_frame.grid(row=0, column=1, rowspan=2, sticky="e")
        tk.Label(
            status_frame,
            textvariable=self.status_var,
            bg=COLORS["background"],
            fg=COLORS["accent"],
            font=(self.base_font, 10, "bold"),
        ).pack(anchor="e")
        tk.Label(
            status_frame,
            textvariable=self.summary_var,
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=(self.base_font, 9),
        ).pack(anchor="e", pady=(5, 0))

        action_panel = tk.Frame(shell, bg=COLORS["background"])
        action_panel.grid(row=1, column=0, sticky="ew", pady=(22, 14))
        action_panel.grid_columnconfigure(0, weight=1)
        action_panel.grid_columnconfigure(4, weight=1)

        self.scan_button = self._make_button(
            action_panel,
            UI_TEXT["primary_button"],
            self._start_scan,
            variant="primary",
            font_size=13,
            padx=26,
            pady=10,
        )
        self.scan_button.grid(row=0, column=1, padx=(0, 10))

        self.recent_button = self._make_button(
            action_panel,
            UI_TEXT["secondary_button"],
            self._sort_by_recently_unused,
            variant="secondary",
            font_size=10,
            padx=16,
            pady=8,
        )
        self.recent_button.grid(row=0, column=2, padx=(0, 10))

        self.cancel_button = self._make_button(
            action_panel,
            UI_TEXT["cancel_button"],
            self._cancel_scan,
            variant="secondary",
            font_size=10,
            padx=16,
            pady=8,
        )
        self.cancel_button.grid(row=0, column=3)
        self.cancel_button.configure(state=tk.DISABLED)

        search_frame = tk.Frame(shell, bg=COLORS["background"])
        search_frame.grid(row=2, column=0, sticky="ew", pady=(0, 12))
        search_frame.grid_columnconfigure(1, weight=1)

        tk.Label(
            search_frame,
            text=UI_TEXT["search_label"],
            bg=COLORS["background"],
            fg=COLORS["text"],
            font=(self.base_font, 10, "bold"),
        ).grid(row=0, column=0, sticky="w", padx=(0, 10))
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var, style="Doko.TEntry")
        self.search_entry.grid(row=0, column=1, sticky="ew")
        self.search_entry.insert(0, UI_TEXT["search_placeholder"])
        self.search_entry.configure(foreground=COLORS["muted"])
        self.search_entry.bind("<FocusIn>", self._clear_placeholder)
        self.search_entry.bind("<FocusOut>", self._restore_placeholder)
        self.search_var.trace_add("write", self._on_search_changed)

        root_text = UI_TEXT["scan_root_line"].format(roots=" / ".join(str(root) for root in self.search_roots))
        tk.Label(
            shell,
            text=root_text,
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=(self.base_font, 8),
            anchor="w",
        ).grid(row=4, column=0, sticky="ew", pady=(10, 0))

        self.results_canvas = tk.Canvas(
            shell,
            bg=COLORS["background"],
            highlightthickness=0,
            borderwidth=0,
        )
        self.results_canvas.grid(row=3, column=0, sticky="nsew")
        self.results_scrollbar = ttk.Scrollbar(shell, orient=tk.VERTICAL, command=self.results_canvas.yview)
        self.results_scrollbar.grid(row=3, column=1, sticky="ns")
        self.results_canvas.configure(yscrollcommand=self.results_scrollbar.set)

        self.cards_frame = tk.Frame(self.results_canvas, bg=COLORS["background"])
        self.cards_window = self.results_canvas.create_window((0, 0), window=self.cards_frame, anchor="nw")
        self.cards_frame.bind("<Configure>", self._refresh_scroll_region)
        self.results_canvas.bind("<Configure>", self._resize_cards_window)
        self.results_canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        footer = tk.Frame(shell, bg=COLORS["background"])
        footer.grid(row=5, column=0, sticky="ew", pady=(12, 0))
        footer.grid_columnconfigure(0, weight=1)
        footer_text = (
            f"{UI_TEXT['footer_left']} {UI_TEXT['footer_separator']} "
            f"{UI_TEXT['footer_caption']} {UI_TEXT['footer_separator']} {UI_TEXT['footer_copyright']}"
        )
        tk.Label(
            footer,
            text=footer_text,
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=(self.base_font, 8),
        ).grid(row=0, column=0, sticky="w")

        self._render_results()

    def _make_button(
        self,
        parent: tk.Misc,
        label: str,
        command: object,
        *,
        variant: str,
        font_size: int,
        padx: int,
        pady: int,
    ) -> tk.Button:
        is_primary = variant == "primary"
        normal_bg = COLORS["accent"] if is_primary else COLORS["card"]
        hover_bg = COLORS["accent_hover"] if is_primary else COLORS["surface_alt"]
        fg = "#FFFFFF" if is_primary else COLORS["text"]
        button = tk.Button(
            parent,
            text=label,
            command=command,
            bg=normal_bg,
            fg=fg,
            activebackground=hover_bg,
            activeforeground=fg,
            relief=tk.FLAT,
            bd=0,
            highlightthickness=1,
            highlightbackground=COLORS["accent"] if is_primary else COLORS["border"],
            font=(self.base_font, font_size, "bold" if is_primary else "normal"),
            padx=padx,
            pady=pady,
            cursor="hand2",
        )
        button.bind("<Enter>", lambda _event: self._set_button_bg(button, hover_bg))
        button.bind("<Leave>", lambda _event: self._set_button_bg(button, normal_bg))
        return button

    def _set_button_bg(self, button: tk.Button, color: str) -> None:
        if str(button.cget("state")) == tk.NORMAL:
            button.configure(bg=color)

    def _clear_placeholder(self, _event: tk.Event[tk.Entry]) -> None:
        if self.search_var.get() == UI_TEXT["search_placeholder"]:
            self.search_entry.delete(0, tk.END)
            self.search_entry.configure(foreground=COLORS["text"])

    def _restore_placeholder(self, _event: tk.Event[tk.Entry]) -> None:
        if not self.search_var.get():
            self.search_entry.insert(0, UI_TEXT["search_placeholder"])
            self.search_entry.configure(foreground=COLORS["muted"])

    def _query_text(self) -> str:
        value = self.search_var.get().strip()
        return "" if value == UI_TEXT["search_placeholder"] else value.lower()

    def _on_search_changed(self, *_args: object) -> None:
        if self.search_var.get() == UI_TEXT["search_placeholder"]:
            return
        self._apply_filter()

    def _start_scan(self) -> None:
        if self.worker_thread is not None and self.worker_thread.is_alive():
            return
        self.recent_mode = False
        self.candidates = []
        self.filtered_candidates = []
        self.cancel_event.clear()
        self.scan_button.configure(state=tk.DISABLED)
        self.cancel_button.configure(state=tk.NORMAL)
        self.status_var.set(UI_TEXT["status_searching"])
        self.summary_var.set(UI_TEXT["summary_initial"])
        self._render_results()
        history = self._launch_history()
        self.worker_thread = threading.Thread(
            target=scan_worker,
            args=(self.event_queue, self.cancel_event, self.search_roots, history),
            daemon=True,
        )
        self.worker_thread.start()

    def _cancel_scan(self) -> None:
        self.cancel_event.set()
        self.cancel_button.configure(state=tk.DISABLED)
        self.status_var.set(UI_TEXT["status_cancel_requested"])

    def _sort_by_recently_unused(self) -> None:
        self.recent_mode = True
        self._apply_filter()

    def _launch_history(self) -> dict[str, str]:
        history = self.config.get("launch_history", {})
        return {str(key): str(value) for key, value in history.items()} if isinstance(history, dict) else {}

    def _save_launch_history(self, identity: str, launched_at: str) -> None:
        history = self._launch_history()
        history[identity] = launched_at
        self.config["launch_history"] = history
        save_config(self.config)

    def _poll_queue(self) -> None:
        try:
            while True:
                event = self.event_queue.get_nowait()
                self._handle_worker_event(event)
        except queue.Empty:
            pass
        if self.root.winfo_exists():
            self.root.after(QUEUE_POLL_INTERVAL_MS, self._poll_queue)

    def _handle_worker_event(self, event: dict[str, object]) -> None:
        event_type = event.get("type")
        if event_type == "status":
            self.status_var.set(str(event.get("value", "")))
            return
        if event_type == "found":
            candidate = event.get("candidate")
            if isinstance(candidate, AppCandidate):
                self.candidates.append(candidate)
                self._apply_filter()
            return
        if event_type == "done":
            self.scan_button.configure(state=tk.NORMAL)
            self.cancel_button.configure(state=tk.DISABLED)
            cancelled = bool(event.get("cancelled", False))
            roots = event.get("roots", [])
            self.config["last_scan_roots"] = roots if isinstance(roots, list) else []
            save_config(self.config)
            self.status_var.set(UI_TEXT["status_cancelled"] if cancelled else UI_TEXT["status_completed"])
            self._apply_filter()
            return
        if event_type == "error":
            self.scan_button.configure(state=tk.NORMAL)
            self.cancel_button.configure(state=tk.DISABLED)
            error = event.get("error")
            self.status_var.set(UI_TEXT["status_error"])
            write_debug_log("scan failed", log_dir=LOG_DIR, exc=error if isinstance(error, Exception) else None)
            messagebox.showerror(UI_TEXT["dialog_error_title"], str(error))

    def _apply_filter(self) -> None:
        query = self._query_text()
        if query:
            candidates = [candidate for candidate in self.candidates if query in candidate.search_text]
        else:
            candidates = list(self.candidates)
        if self.recent_mode:
            candidates.sort(key=self._recent_sort_key)
        else:
            candidates.sort(key=lambda candidate: candidate.discovered_order)
        self.filtered_candidates = candidates
        total = len(self.candidates)
        shown = len(candidates)
        if total and shown != total:
            self.summary_var.set(UI_TEXT["summary_filter_count"].format(shown=shown, total=total))
        elif total:
            self.summary_var.set(UI_TEXT["summary_count"].format(count=total))
        else:
            self.summary_var.set(UI_TEXT["summary_initial"])
        self._render_results()

    def _recent_sort_key(self, candidate: AppCandidate) -> tuple[int, datetime, str]:
        launched_at = parse_history_datetime(candidate.launch_history_at)
        if launched_at is None:
            return (0, datetime.min, candidate.display_name)
        return (1, launched_at, candidate.display_name)

    def _render_results(self) -> None:
        for child in self.cards_frame.winfo_children():
            child.destroy()
        if not self.candidates:
            self.results_scrollbar.grid_remove()
            self._render_empty(UI_TEXT["empty_title"], UI_TEXT["empty_description"])
            return
        if not self.filtered_candidates:
            self.results_scrollbar.grid_remove()
            self._render_empty(UI_TEXT["no_result_title"], UI_TEXT["no_result_description"])
            return
        self.results_scrollbar.grid(row=3, column=1, sticky="ns")
        for row, candidate in enumerate(self.filtered_candidates):
            self._render_card(candidate, row)

    def _render_empty(self, title: str, description: str) -> None:
        empty = tk.Frame(self.cards_frame, bg=COLORS["background"])
        empty.pack(fill=tk.BOTH, expand=True, pady=90)
        tk.Label(
            empty,
            text=title,
            bg=COLORS["background"],
            fg=COLORS["text"],
            font=(self.base_font, 15, "bold"),
        ).pack()
        tk.Label(
            empty,
            text=description,
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=(self.base_font, 10),
        ).pack(pady=(8, 0))

    def _render_card(self, candidate: AppCandidate, row: int) -> None:
        card = tk.Frame(
            self.cards_frame,
            bg=COLORS["card"],
            highlightthickness=1,
            highlightbackground=COLORS["border"],
        )
        card.pack(fill=tk.X, padx=(0, 8), pady=(0 if row == 0 else 8, 8))
        card.grid_columnconfigure(0, weight=1)

        header = tk.Frame(card, bg=COLORS["card"])
        header.grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 6))
        header.grid_columnconfigure(0, weight=1)
        tk.Label(
            header,
            text=candidate.display_name,
            bg=COLORS["card"],
            fg=COLORS["text"],
            font=(self.base_font, 13, "bold"),
            anchor="w",
        ).grid(row=0, column=0, sticky="ew")

        actions = tk.Frame(header, bg=COLORS["card"])
        actions.grid(row=0, column=1, sticky="e", padx=(12, 0))
        self._make_button(
            actions,
            UI_TEXT["button_launch"],
            lambda target=candidate: self._launch_app(target),
            variant="primary",
            font_size=9,
            padx=12,
            pady=5,
        ).pack(side=tk.LEFT, padx=(0, 6))
        self._make_button(
            actions,
            UI_TEXT["button_open_folder"],
            lambda target=candidate: self._open_folder(target),
            variant="secondary",
            font_size=9,
            padx=10,
            pady=5,
        ).pack(side=tk.LEFT, padx=(0, 6))
        self._make_button(
            actions,
            UI_TEXT["button_copy_path"],
            lambda target=candidate: self._copy_path(target),
            variant="secondary",
            font_size=9,
            padx=10,
            pady=5,
        ).pack(side=tk.LEFT)

        description = tk.Label(
            card,
            text=candidate.description,
            bg=COLORS["card"],
            fg=COLORS["muted"],
            font=(self.base_font, 10),
            anchor="w",
            justify=tk.LEFT,
            wraplength=CARD_WRAP_LENGTH,
        )
        description.grid(row=1, column=0, sticky="ew", padx=16)

        details = tk.Frame(card, bg=COLORS["card"])
        details.grid(row=2, column=0, sticky="ew", padx=16, pady=(10, 14))
        details.grid_columnconfigure(1, weight=1)
        rows = (
            (UI_TEXT["label_exe_path"], str(candidate.exe_path)),
            (UI_TEXT["label_folder_path"], str(candidate.folder_path)),
            (UI_TEXT["label_readme"], UI_TEXT["readme_yes"] if candidate.readme_path else UI_TEXT["readme_no"]),
            (UI_TEXT["label_updated"], format_datetime(candidate.modified_at)),
            (UI_TEXT["label_launched"], format_datetime(parse_history_datetime(candidate.launch_history_at))),
        )
        for index, (label, value) in enumerate(rows):
            tk.Label(
                details,
                text=label,
                bg=COLORS["card"],
                fg=COLORS["muted"],
                font=(self.base_font, 8, "bold"),
                anchor="w",
            ).grid(row=index, column=0, sticky="nw", padx=(0, 12), pady=2)
            tk.Label(
                details,
                text=value,
                bg=COLORS["card"],
                fg=COLORS["text"],
                font=(self.base_font, 8),
                anchor="w",
                justify=tk.LEFT,
                wraplength=CARD_WRAP_LENGTH,
            ).grid(row=index, column=1, sticky="ew", pady=2)

    def _launch_app(self, candidate: AppCandidate) -> None:
        try:
            subprocess.Popen([str(candidate.exe_path)], cwd=str(candidate.folder_path))
            launched_at = datetime.now().isoformat(timespec="seconds")
            self._save_launch_history(candidate.identity, launched_at)
            self.status_var.set(UI_TEXT["status_launched"])
            self._refresh_candidate_history(candidate.identity, launched_at)
        except Exception as exc:
            self.status_var.set(UI_TEXT["status_error"])
            write_debug_log("launch failed", log_dir=LOG_DIR, exc=exc, context={"path": candidate.exe_path})
            messagebox.showerror(UI_TEXT["dialog_error_title"], f"{UI_TEXT['dialog_launch_error']}\n{exc}")

    def _refresh_candidate_history(self, identity: str, launched_at: str) -> None:
        refreshed: list[AppCandidate] = []
        for candidate in self.candidates:
            if candidate.identity == identity:
                refreshed.append(
                    AppCandidate(
                        identity=candidate.identity,
                        display_name=candidate.display_name,
                        description=candidate.description,
                        exe_name=candidate.exe_name,
                        exe_path=candidate.exe_path,
                        folder_path=candidate.folder_path,
                        readme_path=candidate.readme_path,
                        meta=candidate.meta,
                        modified_at=candidate.modified_at,
                        launch_history_at=launched_at,
                        discovered_order=candidate.discovered_order,
                        search_text=candidate.search_text,
                    )
                )
            else:
                refreshed.append(candidate)
        self.candidates = refreshed
        self._apply_filter()

    def _open_folder(self, candidate: AppCandidate) -> None:
        try:
            os.startfile(str(candidate.folder_path))
            self.status_var.set(UI_TEXT["status_folder_opened"])
        except Exception as exc:
            self.status_var.set(UI_TEXT["status_error"])
            write_debug_log("folder open failed", log_dir=LOG_DIR, exc=exc, context={"path": candidate.folder_path})
            messagebox.showerror(UI_TEXT["dialog_error_title"], f"{UI_TEXT['dialog_folder_error']}\n{exc}")

    def _copy_path(self, candidate: AppCandidate) -> None:
        self.root.clipboard_clear()
        self.root.clipboard_append(str(candidate.exe_path))
        self.status_var.set(UI_TEXT["status_path_copied"])

    def _refresh_scroll_region(self, _event: tk.Event[tk.Frame]) -> None:
        self.results_canvas.configure(scrollregion=self.results_canvas.bbox("all"))

    def _resize_cards_window(self, event: tk.Event[tk.Canvas]) -> None:
        self.results_canvas.itemconfigure(self.cards_window, width=event.width)

    def _on_mousewheel(self, event: tk.Event[tk.Canvas]) -> None:
        if self.results_canvas.winfo_viewable():
            self.results_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def run(self) -> None:
        self.root.mainloop()


def create_launch_check_window() -> tk.Tk:
    root = tk.Tk()
    root.withdraw()
    AppDoko(root)
    return root


def run_app() -> None:
    root = tk.Tk()
    AppDoko(root).run()


def main(argv: list[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    if "--launch-check" in args:
        return run_launch_check(
            checks=(read_own_meta, load_config),
            create_window=create_launch_check_window,
            log_dir=str(LOG_DIR),
        )
    result = safe_run(run_app, title=APP_NAME, log_dir=str(LOG_DIR), show_error=True)
    return 0 if result.ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
