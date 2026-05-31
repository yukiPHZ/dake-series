# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import importlib.util
import json
import os
import queue
import re
import subprocess
import sys
import threading
import time
import webbrowser
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tkinter import font as tkfont
import tkinter as tk
from tkinter import messagebox, ttk

try:
    import ctypes
except Exception:
    ctypes = None


APP_NAME = "QPSC Dashboard"
WINDOW_TITLE = "QPSC Dashboard"
APP_ID = "dake.qpsc.dashboard"
APP_FOLDER_NAME = "DAKE_QPSC_Dashboard"
APP_DASHBOARD_FOLDER = "DAKE_App_Dashboard"
WEB_DASHBOARD_FOLDER = "DAKE_Web_Dashboard"
APP_DASHBOARD_EXE = "DakeApp_Dashboard.exe"
WEB_DASHBOARD_EXE = "DakeWeb_Dashboard.exe"
DIST_DIR_NAME = "dist"
DEFAULT_SERIES_ROOT = Path(os.environ.get("DAKE_SERIES_ROOT", r"C:\Users\yukiz\devlop\DAKE_series"))
DEFAULT_APPS_ROOT = Path(os.environ.get("QPSC_SERIES_APPS_ROOT", str(DEFAULT_SERIES_ROOT / "01_apps")))
WORKER_POLL_MS = 80

UI_TEXT = {
    "app_title": "QPSC Dashboard",
    "header_title": "QPSC",
    "header_subtitle": "今日の状態",
    "button_reload": "再確認",
    "button_open_app_dashboard": "App Dashboardを開く",
    "button_open_web_dashboard": "Web Dashboardを開く",
    "button_open_failed_title": "起動できません",
    "button_open_failed": "対象が見つかりません。\n\n{path}",
    "button_open_error": "起動に失敗しました。\n\n{path}\n\n{error}",
    "metric_apps": "アプリ",
    "metric_sites": "サイト",
    "metric_brain": "補助脳",
    "metric_git": "Git",
    "metric_notice": "通知",
    "metric_app_suffix": "未完了 {count}",
    "metric_site_suffix": "要確認 {count}",
    "metric_brain_value": "未接続",
    "metric_git_suffix": "未commit {count}",
    "metric_notice_suffix": "{count}件",
    "section_nodes": "監視ノード",
    "section_next": "次にやる候補",
    "node_app_title": "アプリ",
    "node_app_source": "取得元: DAKE_App_Dashboard",
    "node_web_title": "サイト",
    "node_web_source": "取得元: DAKE_Web_Dashboard",
    "node_brain_title": "補助脳",
    "node_brain_source": "現時点: ダミー",
    "node_git_title": "Git",
    "node_git_source": "取得元: DAKE_App_Dashboard",
    "node_notice_title": "通知",
    "node_notice_source": "現時点: ダミー",
    "label_total_apps": "総アプリ数",
    "label_available": "available数",
    "label_internal": "internal数",
    "label_ship_rate": "正式出荷率",
    "label_app_review": "要確認数",
    "label_total_sites": "サイト数",
    "label_api_warning": "API警告数",
    "label_health_issue": "health異常数",
    "label_git_unapplied": "Git未反映数",
    "label_brain_state": "状態",
    "label_git_uncommitted": "未commit件数",
    "label_git_untracked": "未追跡件数",
    "label_notice_count": "通知件数",
    "value_waiting": "確認待ち",
    "value_percent": "{value}%",
    "value_none": "なし",
    "value_ready": "準備中",
    "status_checking": "確認中…",
    "status_ready": "確認完了",
    "status_attention": "要確認あり",
    "status_launch_check_ok": "LAUNCH CHECK OK",
    "last_loaded_waiting": "未確認",
    "last_loaded_value": "確認: {time}",
    "next_none": "現時点で明確な候補はありません。",
    "next_app_header": "アプリ: 要確認 {count}件",
    "next_site_header": "サイト: 要確認 {count}件",
    "next_git_header": "Git: 未commit {count}件",
    "next_source_error": "{source}: 状態取得に失敗しました",
    "error_missing_main": "main.py が見つかりません",
    "error_missing_function": "既存Dashboardの取得関数が見つかりません",
    "error_import_failed": "既存Dashboardを読み込めません: {error}",
    "error_scan_failed": "既存Dashboardの状態取得に失敗しました: {error}",
    "source_app": "App Dashboard",
    "source_web": "Web Dashboard",
    "summary_template": "アプリ {apps}件 / サイト {sites}件を確認。要確認 {attention}件。",
    "launch_check_template": "LAUNCH CHECK OK: apps={apps} app_needs_review={app_needs} sites={sites} site_attention={site_attention} git_uncommitted={git_uncommitted}",
}

THEME = {
    "bg": "#F5F7FB",
    "panel": "#FFFFFF",
    "panel_alt": "#F9FAFD",
    "border": "#D7DCE7",
    "text": "#182230",
    "muted": "#667085",
    "quiet": "#98A2B3",
    "accent": "#2457C5",
    "accent_hover": "#1D4CB4",
    "green": "#1F7A4D",
    "green_bg": "#EAF7F0",
    "amber": "#9A5B00",
    "amber_bg": "#FFF4D8",
    "blue_bg": "#EAF1FF",
    "red": "#B42318",
    "red_bg": "#FFF0EE",
    "shadow": "#EEF2F7",
}

FONT_CANDIDATES = ["BIZ UDPGothic", "Yu Gothic UI", "Meiryo", "MS Gothic"]


@dataclass(frozen=True)
class AppSummary:
    total: int = 0
    available: int = 0
    internal: int = 0
    ship_rate: int = 0
    needs_review: int = 0
    candidates: tuple[str, ...] = ()
    error: str = ""


@dataclass(frozen=True)
class SiteSummary:
    total: int = 0
    api_warnings: int = 0
    health_issues: int = 0
    git_unapplied: int = 0
    attention: int = 0
    candidates: tuple[str, ...] = ()
    error: str = ""


@dataclass(frozen=True)
class GitSummary:
    uncommitted: int = 0
    untracked: int = 0
    error: str = ""


@dataclass(frozen=True)
class DashboardSummary:
    app: AppSummary
    site: SiteSummary
    git: GitSummary
    loaded_at: datetime
    warnings: tuple[str, ...] = ()

    @property
    def notification_count(self) -> int:
        return min(99, self.app.needs_review + self.site.attention + len(self.warnings))

    @property
    def attention_count(self) -> int:
        return self.app.needs_review + self.site.attention + len(self.warnings)


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


def dashboard_dir(env_name: str, folder_name: str) -> Path:
    env_value = os.environ.get(env_name, "").strip()
    candidates: list[Path] = []
    if env_value:
        candidates.append(Path(env_value))
    candidates.extend(
        [
            app_dir().parent / folder_name,
            DEFAULT_APPS_ROOT / folder_name,
            app_dir().parent / "codex_staging_dake_dashboard_phase4" if folder_name == APP_DASHBOARD_FOLDER else app_dir().parent / "codex_staging_web_dashboard" / WEB_DASHBOARD_FOLDER,
        ]
    )
    for candidate in candidates:
        if (candidate / "main.py").exists():
            return candidate
    return candidates[0] if candidates else DEFAULT_APPS_ROOT / folder_name


def app_dashboard_dir() -> Path:
    return dashboard_dir("QPSC_APP_DASHBOARD_DIR", APP_DASHBOARD_FOLDER)


def web_dashboard_dir() -> Path:
    return dashboard_dir("QPSC_WEB_DASHBOARD_DIR", WEB_DASHBOARD_FOLDER)


def load_dashboard_module(name: str, folder: Path):
    module_path = folder / "main.py"
    if not module_path.exists():
        raise FileNotFoundError(UI_TEXT["error_missing_main"])
    module_name = f"qpsc_external_{name}_{abs(hash(str(module_path.resolve())))}"
    spec = importlib.util.spec_from_file_location(module_name, module_path)
    if spec is None or spec.loader is None:
        raise ImportError(UI_TEXT["error_import_failed"].format(error=module_path))
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    try:
        spec.loader.exec_module(module)
    except Exception as exc:
        raise ImportError(UI_TEXT["error_import_failed"].format(error=exc)) from exc
    return module


def bool_from_meta(value: object) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.strip().lower() in {"true", "1", "yes", "on"}
    if isinstance(value, (int, float)):
        return bool(value)
    return False


def collect_app_summary() -> tuple[AppSummary, GitSummary, tuple[str, ...]]:
    folder = app_dashboard_dir()
    try:
        module = load_dashboard_module("app", folder)
        if not hasattr(module, "scan_apps") or not hasattr(module, "apps_root"):
            raise AttributeError(UI_TEXT["error_missing_function"])
        records = list(module.scan_apps(folder.parent))
        is_internal = getattr(module, "is_internal_app", None)
        formal_ship_line_reached = getattr(module, "formal_ship_line_reached", None)
        next_candidates = getattr(module, "next_candidates", None)

        def record_is_internal(record: object) -> bool:
            return bool(is_internal(record)) if callable(is_internal) else False

        internal_count = sum(1 for record in records if record_is_internal(record))
        public_records = [record for record in records if not record_is_internal(record)]
        available_count = sum(1 for record in records if bool_from_meta(getattr(record, "meta_fields", {}).get("show_in_launcher", "")))
        shipped_count = (
            sum(1 for record in public_records if bool(formal_ship_line_reached(record)))
            if callable(formal_ship_line_reached)
            else 0
        )
        ship_rate = round(shipped_count / len(public_records) * 100) if public_records else 0
        needs_review = sum(
            1
            for record in records
            if getattr(record, "status_key", "") == "needs_review" or bool(getattr(record, "issue_messages", ()))
        )
        candidate_lines: list[str] = []
        if callable(next_candidates):
            for record, reason in next_candidates(records, limit=3):
                candidate_lines.append(f"{getattr(record, 'folder_name', '')}: {reason}")
        git_summary = collect_git_summary(module, folder.parent.parent)
        return (
            AppSummary(
                total=len(records),
                available=available_count,
                internal=internal_count,
                ship_rate=ship_rate,
                needs_review=needs_review,
                candidates=tuple(candidate_lines),
            ),
            git_summary,
            (),
        )
    except Exception as exc:
        warning = UI_TEXT["next_source_error"].format(source=UI_TEXT["source_app"])
        return AppSummary(error=UI_TEXT["error_scan_failed"].format(error=exc)), GitSummary(error=str(exc)), (warning,)


def collect_git_summary(app_module, series_root_path: Path) -> GitSummary:
    try:
        read_git_status = getattr(app_module, "read_git_status")
        status = read_git_status(series_root_path)
        return GitSummary(
            uncommitted=int(getattr(status, "uncommitted_count", 0)),
            untracked=int(getattr(status, "untracked_count", 0)),
            error=str(getattr(status, "error", "")),
        )
    except Exception as exc:
        return GitSummary(error=str(exc))


def collect_site_summary() -> tuple[SiteSummary, tuple[str, ...]]:
    folder = web_dashboard_dir()
    try:
        module = load_dashboard_module("web", folder)
        if not hasattr(module, "scan_sites"):
            raise AttributeError(UI_TEXT["error_missing_function"])
        dev_root = getattr(module, "DEV_ROOT", DEFAULT_SERIES_ROOT.parent)
        records = list(module.scan_sites(dev_root))
        api_warnings = 0
        health_issues = 0
        git_unapplied = 0
        attention = 0
        candidates: list[str] = []
        for record in records:
            class_key = getattr(record, "class_key", "")
            api = getattr(record, "api", None)
            cloudflare = getattr(record, "cloudflare", None)
            files = getattr(record, "files", None)
            git = getattr(record, "git", None)
            if class_key == "api_review" or bool(getattr(api, "has_hardcoded_key_suspect", False)) or bool(getattr(api, "has_frontend_openai_direct", False)):
                api_warnings += 1
            if bool(getattr(files, "has_functions", False)) and not bool(getattr(cloudflare, "has_health_file", False)):
                health_issues += 1
            if bool(getattr(git, "has_dirty", False)) or int(getattr(git, "ahead", 0) or 0) > 0:
                git_unapplied += 1
            if class_key in {"needs_review", "api_review", "deploy_review"}:
                attention += 1
                if len(candidates) < 3:
                    title = getattr(record, "display_name", "") or getattr(record, "folder_name", "")
                    class_text = getattr(record, "class_text", class_key)
                    candidates.append(f"{title}: {class_text}")
        return (
            SiteSummary(
                total=len(records),
                api_warnings=api_warnings,
                health_issues=health_issues,
                git_unapplied=git_unapplied,
                attention=attention,
                candidates=tuple(candidates),
            ),
            (),
        )
    except Exception as exc:
        warning = UI_TEXT["next_source_error"].format(source=UI_TEXT["source_web"])
        return SiteSummary(error=UI_TEXT["error_scan_failed"].format(error=exc)), (warning,)


def collect_dashboard_summary() -> DashboardSummary:
    app_summary, git_summary, app_warnings = collect_app_summary()
    site_summary, site_warnings = collect_site_summary()
    return DashboardSummary(
        app=app_summary,
        site=site_summary,
        git=git_summary,
        loaded_at=datetime.now(),
        warnings=tuple(app_warnings + site_warnings),
    )


def launch_target(folder: Path, exe_name: str) -> Path:
    exe_path = folder / DIST_DIR_NAME / exe_name
    if exe_path.exists():
        return exe_path
    main_path = folder / "main.py"
    if main_path.exists():
        return main_path
    return exe_path


def launch_dashboard(folder: Path, exe_name: str) -> None:
    target = launch_target(folder, exe_name)
    if not target.exists():
        raise FileNotFoundError(str(target))
    if target.suffix.lower() == ".exe":
        subprocess.Popen([str(target)], cwd=str(folder))
        return
    subprocess.Popen([sys.executable, str(target)], cwd=str(folder))


class QpscDashboardApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.font_family = choose_font_family(root)
        self.worker_thread: threading.Thread | None = None
        self.worker_queue: queue.Queue[tuple[str, object]] = queue.Queue()

        self.metric_vars = {
            "apps": tk.StringVar(value=UI_TEXT["value_waiting"]),
            "sites": tk.StringVar(value=UI_TEXT["value_waiting"]),
            "brain": tk.StringVar(value=UI_TEXT["metric_brain_value"]),
            "git": tk.StringVar(value=UI_TEXT["value_waiting"]),
            "notice": tk.StringVar(value=UI_TEXT["value_waiting"]),
        }
        self.app_vars = {
            "total": tk.StringVar(value="0"),
            "available": tk.StringVar(value="0"),
            "internal": tk.StringVar(value="0"),
            "ship_rate": tk.StringVar(value=UI_TEXT["value_percent"].format(value=0)),
            "needs_review": tk.StringVar(value="0"),
        }
        self.site_vars = {
            "total": tk.StringVar(value="0"),
            "api_warnings": tk.StringVar(value="0"),
            "health_issues": tk.StringVar(value="0"),
            "git_unapplied": tk.StringVar(value="0"),
        }
        self.brain_vars = {"state": tk.StringVar(value=UI_TEXT["metric_brain_value"])}
        self.git_vars = {
            "uncommitted": tk.StringVar(value="0"),
            "untracked": tk.StringVar(value="0"),
        }
        self.notice_vars = {"count": tk.StringVar(value="0")}
        self.summary_var = tk.StringVar(value=UI_TEXT["last_loaded_waiting"])
        self.next_var = tk.StringVar(value=UI_TEXT["next_none"])
        self.status_var = tk.StringVar(value=UI_TEXT["status_checking"])

        self.configure_root()
        self.build_ui()
        self.root.after(120, self.refresh)

    def configure_root(self) -> None:
        self.root.title(WINDOW_TITLE)
        self.root.geometry("1200x800")
        self.root.minsize(1100, 720)
        self.root.configure(bg=THEME["bg"])
        apply_window_icon(self.root)

    def build_ui(self) -> None:
        outer = tk.Frame(self.root, bg=THEME["bg"])
        outer.pack(fill="both", expand=True, padx=24, pady=22)

        header = tk.Frame(outer, bg=THEME["bg"])
        header.pack(fill="x")
        title_box = tk.Frame(header, bg=THEME["bg"])
        title_box.pack(side="left", anchor="w")
        tk.Label(
            title_box,
            text=UI_TEXT["header_title"],
            bg=THEME["bg"],
            fg=THEME["text"],
            font=(self.font_family, 28, "bold"),
        ).pack(anchor="w")
        tk.Label(
            title_box,
            text=UI_TEXT["header_subtitle"],
            bg=THEME["bg"],
            fg=THEME["muted"],
            font=(self.font_family, 13),
        ).pack(anchor="w", pady=(2, 0))

        action_box = tk.Frame(header, bg=THEME["bg"])
        action_box.pack(side="right", anchor="e")
        self.reload_button = self.make_button(action_box, UI_TEXT["button_reload"], self.refresh, primary=True)
        self.reload_button.pack(side="left", padx=(0, 8))
        self.make_button(action_box, UI_TEXT["button_open_app_dashboard"], self.open_app_dashboard).pack(side="left", padx=(0, 8))
        self.make_button(action_box, UI_TEXT["button_open_web_dashboard"], self.open_web_dashboard).pack(side="left")

        metrics = tk.Frame(outer, bg=THEME["bg"])
        metrics.pack(fill="x", pady=(24, 18))
        metric_specs = [
            ("apps", UI_TEXT["metric_apps"]),
            ("sites", UI_TEXT["metric_sites"]),
            ("brain", UI_TEXT["metric_brain"]),
            ("git", UI_TEXT["metric_git"]),
            ("notice", UI_TEXT["metric_notice"]),
        ]
        for index, (key, label) in enumerate(metric_specs):
            metrics.grid_columnconfigure(index, weight=1, uniform="metric")
            self.metric_card(metrics, label, self.metric_vars[key]).grid(row=0, column=index, sticky="ew", padx=(0 if index == 0 else 8, 0))

        tk.Label(
            outer,
            text=UI_TEXT["section_nodes"],
            bg=THEME["bg"],
            fg=THEME["text"],
            font=(self.font_family, 12, "bold"),
        ).pack(anchor="w", pady=(0, 8))

        nodes = tk.Frame(outer, bg=THEME["bg"])
        nodes.pack(fill="x")
        nodes.grid_columnconfigure(0, weight=1, uniform="node")
        nodes.grid_columnconfigure(1, weight=1, uniform="node")
        app_card = self.node_card(
            nodes,
            UI_TEXT["node_app_title"],
            UI_TEXT["node_app_source"],
            [
                (UI_TEXT["label_total_apps"], self.app_vars["total"]),
                (UI_TEXT["label_available"], self.app_vars["available"]),
                (UI_TEXT["label_internal"], self.app_vars["internal"]),
                (UI_TEXT["label_ship_rate"], self.app_vars["ship_rate"]),
                (UI_TEXT["label_app_review"], self.app_vars["needs_review"]),
            ],
            self.open_app_dashboard,
            UI_TEXT["button_open_app_dashboard"],
        )
        web_card = self.node_card(
            nodes,
            UI_TEXT["node_web_title"],
            UI_TEXT["node_web_source"],
            [
                (UI_TEXT["label_total_sites"], self.site_vars["total"]),
                (UI_TEXT["label_api_warning"], self.site_vars["api_warnings"]),
                (UI_TEXT["label_health_issue"], self.site_vars["health_issues"]),
                (UI_TEXT["label_git_unapplied"], self.site_vars["git_unapplied"]),
            ],
            self.open_web_dashboard,
            UI_TEXT["button_open_web_dashboard"],
        )
        app_card.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        web_card.grid(row=0, column=1, sticky="nsew", padx=(10, 0))

        lower = tk.Frame(outer, bg=THEME["bg"])
        lower.pack(fill="x", pady=(18, 0))
        for index in range(3):
            lower.grid_columnconfigure(index, weight=1, uniform="lower")
        self.compact_card(
            lower,
            UI_TEXT["node_brain_title"],
            UI_TEXT["node_brain_source"],
            [(UI_TEXT["label_brain_state"], self.brain_vars["state"])],
        ).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.compact_card(
            lower,
            UI_TEXT["node_git_title"],
            UI_TEXT["node_git_source"],
            [
                (UI_TEXT["label_git_uncommitted"], self.git_vars["uncommitted"]),
                (UI_TEXT["label_git_untracked"], self.git_vars["untracked"]),
            ],
        ).grid(row=0, column=1, sticky="ew", padx=8)
        self.compact_card(
            lower,
            UI_TEXT["node_notice_title"],
            UI_TEXT["node_notice_source"],
            [(UI_TEXT["label_notice_count"], self.notice_vars["count"])],
        ).grid(row=0, column=2, sticky="ew", padx=(8, 0))

        next_card = self.make_card(outer)
        next_card.pack(fill="both", expand=True, pady=(18, 0))
        tk.Label(
            next_card,
            text=UI_TEXT["section_next"],
            bg=THEME["panel"],
            fg=THEME["text"],
            font=(self.font_family, 13, "bold"),
        ).pack(anchor="w", padx=16, pady=(14, 4))
        tk.Label(
            next_card,
            textvariable=self.summary_var,
            bg=THEME["panel"],
            fg=THEME["muted"],
            font=(self.font_family, 9),
        ).pack(anchor="w", padx=16)
        tk.Label(
            next_card,
            textvariable=self.next_var,
            bg=THEME["panel"],
            fg=THEME["text"],
            justify="left",
            anchor="nw",
            font=(self.font_family, 10),
        ).pack(fill="both", expand=True, padx=16, pady=(10, 14))

        footer = tk.Frame(self.root, bg=THEME["bg"])
        footer.pack(fill="x", padx=24, pady=(0, 12))
        tk.Label(
            footer,
            textvariable=self.status_var,
            bg=THEME["bg"],
            fg=THEME["muted"],
            font=(self.font_family, 10),
        ).pack(side="left")

    def make_card(self, parent: tk.Misc) -> tk.Frame:
        frame = tk.Frame(parent, bg=THEME["panel"], highlightthickness=1, highlightbackground=THEME["border"])
        return frame

    def make_button(self, parent: tk.Misc, text: str, command, primary: bool = False) -> tk.Button:
        bg = THEME["accent"] if primary else THEME["panel"]
        fg = "#FFFFFF" if primary else THEME["text"]
        active_bg = THEME["accent_hover"] if primary else THEME["panel_alt"]
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg=fg,
            activebackground=active_bg,
            activeforeground=fg,
            relief="flat",
            bd=0,
            padx=14,
            pady=8,
            cursor="hand2",
            font=(self.font_family, 9, "bold" if primary else "normal"),
        )

    def metric_card(self, parent: tk.Misc, label: str, value_var: tk.StringVar) -> tk.Frame:
        frame = self.make_card(parent)
        tk.Label(frame, text=label, bg=THEME["panel"], fg=THEME["muted"], font=(self.font_family, 9)).pack(anchor="w", padx=14, pady=(12, 4))
        tk.Label(frame, textvariable=value_var, bg=THEME["panel"], fg=THEME["text"], font=(self.font_family, 18, "bold")).pack(anchor="w", padx=14, pady=(0, 12))
        return frame

    def node_card(
        self,
        parent: tk.Misc,
        title: str,
        source: str,
        rows: list[tuple[str, tk.StringVar]],
        open_command,
        button_text: str,
    ) -> tk.Frame:
        frame = self.make_card(parent)
        frame.bind("<Button-1>", lambda _event: open_command())
        header = tk.Frame(frame, bg=THEME["panel"])
        header.pack(fill="x", padx=16, pady=(14, 6))
        tk.Label(header, text=title, bg=THEME["panel"], fg=THEME["text"], font=(self.font_family, 14, "bold")).pack(side="left")
        tk.Label(header, text=source, bg=THEME["panel"], fg=THEME["muted"], font=(self.font_family, 9)).pack(side="right")
        body = tk.Frame(frame, bg=THEME["panel"])
        body.pack(fill="x", padx=16, pady=(4, 12))
        for index, (label, var) in enumerate(rows):
            body.grid_columnconfigure(index, weight=1, uniform="metric")
            self.value_pair(body, label, var).grid(row=0, column=index, sticky="ew", padx=(0 if index == 0 else 8, 0))
        self.make_button(frame, button_text, open_command, primary=True).pack(anchor="e", padx=16, pady=(0, 14))
        return frame

    def compact_card(self, parent: tk.Misc, title: str, source: str, rows: list[tuple[str, tk.StringVar]]) -> tk.Frame:
        frame = self.make_card(parent)
        tk.Label(frame, text=title, bg=THEME["panel"], fg=THEME["text"], font=(self.font_family, 12, "bold")).pack(anchor="w", padx=14, pady=(12, 2))
        tk.Label(frame, text=source, bg=THEME["panel"], fg=THEME["muted"], font=(self.font_family, 8)).pack(anchor="w", padx=14)
        body = tk.Frame(frame, bg=THEME["panel"])
        body.pack(fill="x", padx=14, pady=(10, 12))
        for index, (label, var) in enumerate(rows):
            body.grid_columnconfigure(index, weight=1, uniform="compact")
            self.value_pair(body, label, var, compact=True).grid(row=0, column=index, sticky="ew", padx=(0 if index == 0 else 8, 0))
        return frame

    def value_pair(self, parent: tk.Misc, label: str, var: tk.StringVar, compact: bool = False) -> tk.Frame:
        frame = tk.Frame(parent, bg=THEME["panel_alt"], highlightthickness=1, highlightbackground=THEME["shadow"])
        tk.Label(frame, text=label, bg=THEME["panel_alt"], fg=THEME["muted"], font=(self.font_family, 8)).pack(anchor="w", padx=10, pady=(8, 2))
        tk.Label(frame, textvariable=var, bg=THEME["panel_alt"], fg=THEME["text"], font=(self.font_family, 14 if compact else 16, "bold")).pack(anchor="w", padx=10, pady=(0, 8))
        return frame

    def refresh(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            return
        self.status_var.set(UI_TEXT["status_checking"])
        self.reload_button.configure(state="disabled")
        self.worker_thread = threading.Thread(target=self.scan_worker, daemon=True)
        self.worker_thread.start()
        self.root.after(WORKER_POLL_MS, self.poll_worker)

    def scan_worker(self) -> None:
        try:
            summary = collect_dashboard_summary()
            self.worker_queue.put(("summary", summary))
        except Exception as exc:
            self.worker_queue.put(("error", exc))

    def poll_worker(self) -> None:
        try:
            while True:
                event, payload = self.worker_queue.get_nowait()
                if event == "summary" and isinstance(payload, DashboardSummary):
                    self.apply_summary(payload)
                elif event == "error":
                    self.status_var.set(UI_TEXT["status_attention"])
        except queue.Empty:
            pass
        if self.worker_thread and self.worker_thread.is_alive():
            self.root.after(WORKER_POLL_MS, self.poll_worker)
        else:
            self.reload_button.configure(state="normal")

    def apply_summary(self, summary: DashboardSummary) -> None:
        self.metric_vars["apps"].set(UI_TEXT["metric_app_suffix"].format(count=summary.app.needs_review))
        self.metric_vars["sites"].set(UI_TEXT["metric_site_suffix"].format(count=summary.site.attention))
        self.metric_vars["brain"].set(UI_TEXT["metric_brain_value"])
        self.metric_vars["git"].set(UI_TEXT["metric_git_suffix"].format(count=summary.git.uncommitted))
        self.metric_vars["notice"].set(UI_TEXT["metric_notice_suffix"].format(count=summary.notification_count))

        self.app_vars["total"].set(str(summary.app.total))
        self.app_vars["available"].set(str(summary.app.available))
        self.app_vars["internal"].set(str(summary.app.internal))
        self.app_vars["ship_rate"].set(UI_TEXT["value_percent"].format(value=summary.app.ship_rate))
        self.app_vars["needs_review"].set(str(summary.app.needs_review))

        self.site_vars["total"].set(str(summary.site.total))
        self.site_vars["api_warnings"].set(str(summary.site.api_warnings))
        self.site_vars["health_issues"].set(str(summary.site.health_issues))
        self.site_vars["git_unapplied"].set(str(summary.site.git_unapplied))

        self.git_vars["uncommitted"].set(str(summary.git.uncommitted))
        self.git_vars["untracked"].set(str(summary.git.untracked))
        self.notice_vars["count"].set(str(summary.notification_count))
        self.summary_var.set(
            UI_TEXT["summary_template"].format(
                apps=summary.app.total,
                sites=summary.site.total,
                attention=summary.attention_count,
            )
        )
        self.next_var.set(self.build_next_text(summary))
        self.status_var.set(UI_TEXT["status_attention"] if summary.attention_count else UI_TEXT["status_ready"])

    def build_next_text(self, summary: DashboardSummary) -> str:
        lines: list[str] = []
        if summary.app.needs_review:
            lines.append(UI_TEXT["next_app_header"].format(count=summary.app.needs_review))
            lines.extend(summary.app.candidates[:3])
        if summary.site.attention:
            if lines:
                lines.append("")
            lines.append(UI_TEXT["next_site_header"].format(count=summary.site.attention))
            lines.extend(summary.site.candidates[:3])
        if summary.git.uncommitted:
            if lines:
                lines.append("")
            lines.append(UI_TEXT["next_git_header"].format(count=summary.git.uncommitted))
        if summary.warnings:
            if lines:
                lines.append("")
            lines.extend(summary.warnings)
        return "\n".join(lines) if lines else UI_TEXT["next_none"]

    def open_app_dashboard(self) -> None:
        self.open_dashboard(app_dashboard_dir(), APP_DASHBOARD_EXE)

    def open_web_dashboard(self) -> None:
        self.open_dashboard(web_dashboard_dir(), WEB_DASHBOARD_EXE)

    def open_dashboard(self, folder: Path, exe_name: str) -> None:
        try:
            launch_dashboard(folder, exe_name)
        except FileNotFoundError:
            messagebox.showinfo(
                UI_TEXT["button_open_failed_title"],
                UI_TEXT["button_open_failed"].format(path=launch_target(folder, exe_name)),
                parent=self.root,
            )
        except Exception as exc:
            messagebox.showerror(
                UI_TEXT["button_open_failed_title"],
                UI_TEXT["button_open_error"].format(path=folder, error=exc),
                parent=self.root,
            )


def run_gui() -> int:
    set_windows_app_id()
    root = tk.Tk()
    QpscDashboardApp(root)
    root.mainloop()
    return 0


def run_launch_check() -> int:
    summary = collect_dashboard_summary()
    print(
        UI_TEXT["launch_check_template"].format(
            apps=summary.app.total,
            app_needs=summary.app.needs_review,
            sites=summary.site.total,
            site_attention=summary.site.attention,
            git_uncommitted=summary.git.uncommitted,
        )
    )
    for warning in summary.warnings:
        print(warning)
    if summary.app.error:
        print(summary.app.error)
    if summary.site.error:
        print(summary.site.error)
    if summary.git.error:
        print(summary.git.error)
    return 0 if not summary.warnings and not summary.app.error and not summary.site.error else 1


def main() -> int:
    parser = argparse.ArgumentParser(description=APP_NAME)
    parser.add_argument("--launch-check", action="store_true")
    args = parser.parse_args()
    if args.launch_check:
        return run_launch_check()
    return run_gui()


if __name__ == "__main__":
    raise SystemExit(main())
