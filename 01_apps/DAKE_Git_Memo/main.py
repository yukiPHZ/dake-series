# -*- coding: utf-8 -*-
from __future__ import annotations

import ctypes
import difflib
import json
import os
import sys
import tempfile
import webbrowser
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import font as tkfont
from tkinter import messagebox


APP_NAME = "DakeGitメモ"
WINDOW_TITLE = "DakeGitメモ"
COPYRIGHT = "© 2026 しまりす不動産 / Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "app_title": APP_NAME,
    "subtitle": "書くたびに、変更が残るメモです。",
    "memo_label": "メモ",
    "history": "履歴",
    "save": "保存",
    "show_changes": "変更を見る",
    "restore": "この時点に戻す",
    "close": "閉じる",
    "autosaved": "保存しました",
    "saved": "保存しました",
    "save_waiting": "保存待ち",
    "no_changes": "変更はありません",
    "no_history": "まだ履歴はありません",
    "select_history": "履歴を選んでください",
    "history_load_error": "履歴を読み込めませんでした",
    "save_error": "保存できませんでした",
    "restore_confirm_title": "戻す前の確認",
    "restore_confirm_message": "この時点の内容に戻します。現在の内容も履歴に残してから戻します。",
    "restore_complete": "戻しました",
    "restore_same": "すでにこの内容です",
    "changes_title": "変更を見る",
    "changes_description": "選んだ時点から現在までの変更です。",
    "changes_empty": "変更はありません",
    "empty_line": "（空行）",
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

COLORS = {
    "background": "#F6F7F9",
    "surface": "#FFFFFF",
    "surface_soft": "#FAFBFC",
    "text": "#1E2430",
    "muted": "#667085",
    "quiet": "#98A2B3",
    "border": "#E6EAF0",
    "button": "#2F6FED",
    "button_hover": "#2458BF",
    "secondary_button": "#FFFFFF",
    "secondary_hover": "#F2F4F7",
    "delete_bg": "#FFF2F2",
    "delete_fg": "#A64040",
    "add_bg": "#EEF8F2",
    "add_fg": "#236846",
    "select": "#EAF2FF",
}

FONT_CANDIDATES = ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo")
WINDOW_SIZE = "1040x660"
WINDOW_MIN_SIZE = (820, 520)
APP_USER_MODEL_ID = "Shimarisu.DakeGitMemo"
DATA_DIR_ENV = "DAKE_GIT_MEMO_DATA_DIR"
DATA_DIR_NAME = "DAKE_Git_Memo"
MEMO_FILE_NAME = "memo.txt"
CONFIG_FILE_NAME = "config.json"
HISTORY_DIR_NAME = "history"
AUTOSAVE_DELAY_MS = 1400
STATUS_RESET_MS = 2200
FOOTER_NARROW_WIDTH = 940


@dataclass(frozen=True)
class HistoryEntry:
    path: Path
    label: str


@dataclass(frozen=True)
class SaveResult:
    changed: bool
    snapshot_path: Path | None


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def default_data_dir() -> Path:
    override = os.environ.get(DATA_DIR_ENV)
    if override:
        return Path(override).expanduser()

    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        return Path(local_app_data) / DATA_DIR_NAME

    return Path.home() / "AppData" / "Local" / DATA_DIR_NAME


def write_text_atomic(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_name(path.name + ".tmp")
    temp_path.write_text(text, encoding="utf-8")
    temp_path.replace(path)


def read_text_safe(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except FileNotFoundError:
        return ""


def format_history_label(path: Path) -> str:
    try:
        timestamp = datetime.strptime(path.stem[:15], "%Y%m%d_%H%M%S")
        return timestamp.strftime("%Y-%m-%d %H:%M")
    except ValueError:
        return path.stem


def build_change_lines(old_text: str, new_text: str) -> list[tuple[str, str]]:
    old_lines = old_text.splitlines()
    new_lines = new_text.splitlines()
    changes: list[tuple[str, str]] = []

    for line in difflib.ndiff(old_lines, new_lines):
        prefix = line[:2]
        value = line[2:] or UI_TEXT["empty_line"]
        if prefix == "- ":
            changes.append(("delete", f"- {value}"))
        elif prefix == "+ ":
            changes.append(("add", f"+ {value}"))

    return changes


class MemoStore:
    def __init__(self, data_dir: Path | None = None) -> None:
        self.data_dir = data_dir or default_data_dir()
        self.memo_path = self.data_dir / MEMO_FILE_NAME
        self.history_dir = self.data_dir / HISTORY_DIR_NAME
        self.config_path = self.data_dir / CONFIG_FILE_NAME
        self.ensure_ready()

    def ensure_ready(self) -> None:
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.history_dir.mkdir(parents=True, exist_ok=True)
        if not self.memo_path.exists():
            write_text_atomic(self.memo_path, "")
        if not self.config_path.exists():
            self.write_config(
                {
                    "app_key": DATA_DIR_NAME,
                    "data_version": 1,
                    "created_at": now_iso(),
                    "last_saved_at": "",
                    "last_history_file": "",
                }
            )

    def read_memo(self) -> str:
        return read_text_safe(self.memo_path)

    def write_memo(self, text: str) -> None:
        write_text_atomic(self.memo_path, text)

    def write_config(self, payload: dict[str, object]) -> None:
        write_text_atomic(
            self.config_path,
            json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        )

    def update_config(self, snapshot_path: Path | None = None) -> None:
        payload = {
            "app_key": DATA_DIR_NAME,
            "data_version": 1,
            "last_saved_at": now_iso(),
            "last_history_file": snapshot_path.name if snapshot_path else "",
        }
        self.write_config(payload)

    def history_entries(self) -> list[HistoryEntry]:
        files = sorted(self.history_dir.glob("*.txt"), key=lambda item: item.name, reverse=True)
        return [HistoryEntry(path=item, label=format_history_label(item)) for item in files]

    def read_history(self, entry: HistoryEntry) -> str:
        return entry.path.read_text(encoding="utf-8")

    def save_if_changed(self, text: str, force_history: bool = False) -> SaveResult:
        if not force_history and text == self.read_memo():
            return SaveResult(changed=False, snapshot_path=None)

        snapshot_path = self.create_snapshot(text)
        self.write_memo(text)
        self.update_config(snapshot_path)
        return SaveResult(changed=True, snapshot_path=snapshot_path)

    def create_snapshot(self, text: str) -> Path:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        snapshot_path = self.history_dir / f"{timestamp}.txt"
        if snapshot_path.exists():
            for index in range(2, 100):
                candidate = self.history_dir / f"{timestamp}_{index:02d}.txt"
                if not candidate.exists():
                    snapshot_path = candidate
                    break
        write_text_atomic(snapshot_path, text)
        return snapshot_path


def set_windows_app_id() -> None:
    if not sys.platform.startswith("win"):
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)
    except Exception:
        return


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def icon_candidates() -> list[Path]:
    base = app_dir()
    candidates = [
        base / ".." / ".." / "02_assets" / "dake_icon.ico",
        base / "dake_icon.ico",
    ]
    bundle_dir = getattr(sys, "_MEIPASS", None)
    if bundle_dir:
        candidates.insert(0, Path(bundle_dir) / "dake_icon.ico")
    return candidates


def choose_font(root: tk.Tk) -> str:
    available = set(tkfont.families(root))
    for family in FONT_CANDIDATES:
        if family in available:
            return family
    return "TkDefaultFont"


class DakeGitMemoApp:
    def __init__(self, root: tk.Tk, store: MemoStore | None = None) -> None:
        self.root = root
        self.store = store or MemoStore()
        self.font_family = choose_font(root)
        self.history_items: list[HistoryEntry] = []
        self.autosave_id: str | None = None
        self.status_reset_id: str | None = None
        self.footer: tk.Frame | None = None
        self.footer_mode: str | None = None
        self.suppress_text_event = False

        self.fonts = {
            "title": tkfont.Font(root, family=self.font_family, size=20, weight="bold"),
            "description": tkfont.Font(root, family=self.font_family, size=10),
            "label": tkfont.Font(root, family=self.font_family, size=10, weight="bold"),
            "body": tkfont.Font(root, family=self.font_family, size=11),
            "button": tkfont.Font(root, family=self.font_family, size=10, weight="bold"),
            "history": tkfont.Font(root, family=self.font_family, size=10),
            "status": tkfont.Font(root, family=self.font_family, size=9),
            "footer": tkfont.Font(root, family=self.font_family, size=8),
        }

        self._configure_root()
        self._build_ui()
        self._apply_icon()
        self._load_initial_text()
        self.refresh_history()

    def _configure_root(self) -> None:
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(*WINDOW_MIN_SIZE)
        self.root.configure(bg=COLORS["background"])
        self.root.protocol("WM_DELETE_WINDOW", self.request_close)
        self.root.bind_all("<Control-s>", self._save_shortcut)
        self.root.bind_all("<Control-S>", self._save_shortcut)

    def _apply_icon(self) -> None:
        for candidate in icon_candidates():
            try:
                icon_path = candidate.resolve()
            except Exception:
                icon_path = candidate
            if not icon_path.exists():
                continue
            try:
                self.root.iconbitmap(str(icon_path))
                return
            except tk.TclError:
                continue

    def _build_ui(self) -> None:
        outer = tk.Frame(self.root, bg=COLORS["background"])
        outer.pack(fill="both", expand=True, padx=24, pady=(22, 16))

        self._build_header(outer)
        self._build_body(outer)
        self._build_status(outer)
        self._build_footer(outer)
        self.root.bind("<Configure>", self._on_root_configure, add="+")

    def _build_header(self, parent: tk.Widget) -> None:
        header = tk.Frame(parent, bg=COLORS["background"])
        header.pack(fill="x", pady=(0, 14))
        header.grid_columnconfigure(0, weight=1)

        title_area = tk.Frame(header, bg=COLORS["background"])
        title_area.grid(row=0, column=0, sticky="w")

        tk.Label(
            title_area,
            text=UI_TEXT["app_title"],
            bg=COLORS["background"],
            fg=COLORS["text"],
            font=self.fonts["title"],
            anchor="w",
        ).pack(anchor="w")

        tk.Label(
            title_area,
            text=UI_TEXT["subtitle"],
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=self.fonts["description"],
            anchor="w",
        ).pack(anchor="w", pady=(5, 0))

        action_area = tk.Frame(header, bg=COLORS["background"])
        action_area.grid(row=0, column=1, sticky="e", padx=(18, 0))

        self.save_button = self._make_button(
            action_area,
            UI_TEXT["save"],
            self.save_manual,
            primary=True,
        )
        self.save_button.pack(side="left")

    def _build_body(self, parent: tk.Widget) -> None:
        body = tk.Frame(parent, bg=COLORS["background"])
        body.pack(fill="both", expand=True)
        body.grid_columnconfigure(0, weight=1)
        body.grid_columnconfigure(1, minsize=270)
        body.grid_rowconfigure(0, weight=1)

        memo_shell = self._make_panel(body)
        memo_shell.grid(row=0, column=0, sticky="nsew", padx=(0, 14))
        memo_shell.grid_rowconfigure(1, weight=1)
        memo_shell.grid_columnconfigure(0, weight=1)

        tk.Label(
            memo_shell,
            text=UI_TEXT["memo_label"],
            bg=COLORS["surface"],
            fg=COLORS["text"],
            font=self.fonts["label"],
            anchor="w",
        ).grid(row=0, column=0, columnspan=2, sticky="ew", padx=14, pady=(12, 8))

        self.memo_text = tk.Text(
            memo_shell,
            bg=COLORS["surface"],
            fg=COLORS["text"],
            insertbackground=COLORS["text"],
            selectbackground=COLORS["select"],
            selectforeground=COLORS["text"],
            inactiveselectbackground=COLORS["select"],
            font=self.fonts["body"],
            wrap="word",
            undo=True,
            relief="flat",
            bd=0,
            padx=14,
            pady=12,
            spacing1=2,
            spacing3=2,
        )
        self.memo_text.grid(row=1, column=0, sticky="nsew", padx=(1, 0), pady=(0, 1))
        self.memo_text.bind("<<Modified>>", self._on_text_modified)

        memo_scroll = tk.Scrollbar(memo_shell, command=self.memo_text.yview)
        memo_scroll.grid(row=1, column=1, sticky="ns", pady=(0, 1))
        self.memo_text.configure(yscrollcommand=memo_scroll.set)

        history_shell = self._make_panel(body)
        history_shell.grid(row=0, column=1, sticky="nsew")
        history_shell.grid_rowconfigure(1, weight=1)
        history_shell.grid_columnconfigure(0, weight=1)

        tk.Label(
            history_shell,
            text=UI_TEXT["history"],
            bg=COLORS["surface"],
            fg=COLORS["text"],
            font=self.fonts["label"],
            anchor="w",
        ).grid(row=0, column=0, columnspan=2, sticky="ew", padx=14, pady=(12, 8))

        self.history_list = tk.Listbox(
            history_shell,
            bg=COLORS["surface_soft"],
            fg=COLORS["text"],
            selectbackground=COLORS["select"],
            selectforeground=COLORS["text"],
            activestyle="none",
            font=self.fonts["history"],
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            exportselection=False,
        )
        self.history_list.grid(row=1, column=0, sticky="nsew", padx=(14, 0), pady=(0, 12))
        self.history_list.bind("<<ListboxSelect>>", lambda _event: self._sync_history_buttons())
        self.history_list.bind("<Double-Button-1>", lambda _event: self.show_changes())

        history_scroll = tk.Scrollbar(history_shell, command=self.history_list.yview)
        history_scroll.grid(row=1, column=1, sticky="ns", padx=(0, 14), pady=(0, 12))
        self.history_list.configure(yscrollcommand=history_scroll.set)

        action_stack = tk.Frame(history_shell, bg=COLORS["surface"])
        action_stack.grid(row=2, column=0, columnspan=2, sticky="ew", padx=14, pady=(0, 14))
        action_stack.grid_columnconfigure(0, weight=1)

        self.changes_button = self._make_button(
            action_stack,
            UI_TEXT["show_changes"],
            self.show_changes,
            primary=False,
        )
        self.changes_button.grid(row=0, column=0, sticky="ew", pady=(0, 8))

        self.restore_button = self._make_button(
            action_stack,
            UI_TEXT["restore"],
            self.restore_selected,
            primary=False,
        )
        self.restore_button.grid(row=1, column=0, sticky="ew")

    def _build_status(self, parent: tk.Widget) -> None:
        self.status_label = tk.Label(
            parent,
            text=UI_TEXT["no_changes"],
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=self.fonts["status"],
            anchor="w",
        )
        self.status_label.pack(fill="x", pady=(10, 0))

    def _build_footer(self, parent: tk.Widget) -> None:
        self.footer = tk.Frame(parent, bg=COLORS["background"])
        self.footer.pack(fill="x", pady=(10, 0))
        self._render_footer("wide")

    def _make_panel(self, parent: tk.Widget) -> tk.Frame:
        return tk.Frame(
            parent,
            bg=COLORS["surface"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
        )

    def _make_button(self, parent: tk.Widget, label: str, command, primary: bool) -> tk.Button:
        bg = COLORS["button"] if primary else COLORS["secondary_button"]
        fg = COLORS["surface"] if primary else COLORS["text"]
        hover = COLORS["button_hover"] if primary else COLORS["secondary_hover"]
        border = COLORS["button"] if primary else COLORS["border"]
        button = tk.Button(
            parent,
            text=label,
            command=command,
            bg=bg,
            fg=fg,
            activebackground=hover,
            activeforeground=fg,
            disabledforeground=COLORS["quiet"],
            font=self.fonts["button"],
            relief="flat",
            bd=0,
            padx=18,
            pady=8,
            cursor="hand2",
            highlightthickness=1,
            highlightbackground=border,
        )
        button.bind("<Enter>", lambda _event, target=button, color=hover: self._hover_button(target, color))
        button.bind("<Leave>", lambda _event, target=button, color=bg: self._hover_button(target, color))
        return button

    def _hover_button(self, button: tk.Button, color: str) -> None:
        if str(button.cget("state")) != "disabled":
            button.configure(bg=color)

    def _footer_text(self, parent: tk.Widget, value: str) -> None:
        tk.Label(
            parent,
            text=value,
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=self.fonts["footer"],
        ).pack(side="left")

    def _footer_link(self, parent: tk.Widget, text_key: str) -> None:
        label = tk.Label(
            parent,
            text=UI_TEXT[text_key],
            bg=COLORS["background"],
            fg=COLORS["muted"],
            cursor="hand2",
            font=self.fonts["footer"],
        )
        label.pack(side="left")
        label.bind("<Button-1>", lambda _event, key=text_key: webbrowser.open(LINK_URLS[key]))
        label.bind("<Enter>", lambda _event, target=label: target.configure(fg=COLORS["button"]))
        label.bind("<Leave>", lambda _event, target=label: target.configure(fg=COLORS["muted"]))

    def _on_root_configure(self, event: tk.Event) -> None:
        if event.widget is not self.root:
            return
        next_mode = "narrow" if event.width < FOOTER_NARROW_WIDTH else "wide"
        self._render_footer(next_mode)

    def _render_footer(self, mode: str) -> None:
        if self.footer is None or self.footer_mode == mode:
            return

        self.footer_mode = mode
        for child in self.footer.winfo_children():
            child.destroy()

        if mode == "narrow":
            thought_line = tk.Frame(self.footer, bg=COLORS["background"])
            thought_line.pack(anchor="center")
            self._footer_thought_line(thought_line)

            link_line = tk.Frame(self.footer, bg=COLORS["background"])
            link_line.pack(anchor="center", pady=(4, 0))
            self._footer_link_line(link_line)
            return

        self.footer.grid_columnconfigure(0, weight=0)
        self.footer.grid_columnconfigure(1, weight=1)
        self.footer.grid_columnconfigure(2, weight=0)

        left = tk.Frame(self.footer, bg=COLORS["background"])
        left.grid(row=0, column=0, sticky="w")
        self._footer_thought_line(left)

        right = tk.Frame(self.footer, bg=COLORS["background"])
        right.grid(row=0, column=2, sticky="e")
        self._footer_link_line(right)

    def _footer_thought_line(self, parent: tk.Widget) -> None:
        self._footer_text(parent, UI_TEXT["footer_left"])
        self._footer_text(parent, UI_TEXT["footer_separator"])
        self._footer_text(parent, UI_TEXT["footer_tagline"])

    def _footer_link_line(self, parent: tk.Widget) -> None:
        self._footer_link(parent, "footer_link_1")
        self._footer_text(parent, UI_TEXT["footer_separator"])
        self._footer_link(parent, "footer_link_2")
        self._footer_text(parent, UI_TEXT["footer_separator"])
        self._footer_text(parent, UI_TEXT["footer_copyright"])

    def _load_initial_text(self) -> None:
        self.suppress_text_event = True
        try:
            self.memo_text.delete("1.0", "end")
            self.memo_text.insert("1.0", self.store.read_memo())
            self.memo_text.edit_modified(False)
        finally:
            self.suppress_text_event = False
        self.memo_text.focus_set()

    def _on_text_modified(self, _event: tk.Event) -> None:
        if not self.memo_text.edit_modified():
            return
        self.memo_text.edit_modified(False)
        if self.suppress_text_event:
            return
        self._schedule_autosave()

    def _schedule_autosave(self) -> None:
        if self.autosave_id is not None:
            self.root.after_cancel(self.autosave_id)
        self.autosave_id = self.root.after(AUTOSAVE_DELAY_MS, self.save_auto)
        self._set_status(UI_TEXT["save_waiting"], reset=False)

    def _current_text(self) -> str:
        return self.memo_text.get("1.0", "end-1c")

    def _save_shortcut(self, _event: tk.Event) -> str:
        self.save_manual()
        return "break"

    def save_auto(self) -> None:
        self.autosave_id = None
        self._save_now(UI_TEXT["autosaved"])

    def save_manual(self) -> None:
        if self.autosave_id is not None:
            self.root.after_cancel(self.autosave_id)
            self.autosave_id = None
        self._save_now(UI_TEXT["saved"])

    def _save_now(self, success_text: str, force_history: bool = False, quiet: bool = False) -> SaveResult:
        try:
            result = self.store.save_if_changed(self._current_text(), force_history=force_history)
        except Exception:
            if not quiet:
                self._set_status(UI_TEXT["save_error"], reset=True)
            return SaveResult(changed=False, snapshot_path=None)

        if result.changed:
            self.refresh_history()
            if not quiet:
                self._set_status(success_text, reset=True)
        elif not quiet:
            self._set_status(UI_TEXT["no_changes"], reset=True)
        return result

    def refresh_history(self) -> None:
        self.history_items = self.store.history_entries()
        self.history_list.delete(0, "end")
        if not self.history_items:
            self.history_list.insert("end", UI_TEXT["no_history"])
            self.history_list.itemconfigure(0, fg=COLORS["quiet"])
        else:
            for item in self.history_items:
                self.history_list.insert("end", item.label)
        self._sync_history_buttons()

    def _selected_history(self) -> HistoryEntry | None:
        if not self.history_items:
            return None
        selection = self.history_list.curselection()
        if not selection:
            return None
        index = selection[0]
        if index >= len(self.history_items):
            return None
        return self.history_items[index]

    def _sync_history_buttons(self) -> None:
        state = "normal" if self._selected_history() is not None else "disabled"
        cursor = "hand2" if state == "normal" else "arrow"
        self.changes_button.configure(state=state, cursor=cursor)
        self.restore_button.configure(state=state, cursor=cursor)

    def show_changes(self) -> None:
        entry = self._selected_history()
        if entry is None:
            self._set_status(UI_TEXT["select_history"], reset=True)
            return

        try:
            old_text = self.store.read_history(entry)
        except Exception:
            self._set_status(UI_TEXT["history_load_error"], reset=True)
            return

        changes = build_change_lines(old_text, self._current_text())
        self._open_changes_window(entry.label, changes)

    def _open_changes_window(self, label: str, changes: list[tuple[str, str]]) -> None:
        window = tk.Toplevel(self.root)
        window.title(UI_TEXT["changes_title"])
        window.geometry("680x460")
        window.minsize(520, 320)
        window.configure(bg=COLORS["background"])
        window.transient(self.root)

        shell = tk.Frame(window, bg=COLORS["background"])
        shell.pack(fill="both", expand=True, padx=20, pady=18)
        shell.grid_columnconfigure(0, weight=1)
        shell.grid_rowconfigure(2, weight=1)

        tk.Label(
            shell,
            text=UI_TEXT["changes_title"],
            bg=COLORS["background"],
            fg=COLORS["text"],
            font=self.fonts["label"],
            anchor="w",
        ).grid(row=0, column=0, sticky="ew")

        tk.Label(
            shell,
            text=f"{label}  {UI_TEXT['changes_description']}",
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=self.fonts["description"],
            anchor="w",
        ).grid(row=1, column=0, sticky="ew", pady=(6, 12))

        text_shell = tk.Frame(
            shell,
            bg=COLORS["surface"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
        )
        text_shell.grid(row=2, column=0, sticky="nsew")
        text_shell.grid_columnconfigure(0, weight=1)
        text_shell.grid_rowconfigure(0, weight=1)

        changes_text = tk.Text(
            text_shell,
            bg=COLORS["surface"],
            fg=COLORS["text"],
            font=self.fonts["body"],
            wrap="word",
            relief="flat",
            bd=0,
            padx=14,
            pady=12,
            state="normal",
        )
        changes_text.grid(row=0, column=0, sticky="nsew")
        changes_text.tag_configure("delete", background=COLORS["delete_bg"], foreground=COLORS["delete_fg"])
        changes_text.tag_configure("add", background=COLORS["add_bg"], foreground=COLORS["add_fg"])
        changes_text.tag_configure("quiet", foreground=COLORS["quiet"])

        if not changes:
            changes_text.insert("end", UI_TEXT["changes_empty"], "quiet")
        else:
            for tag, line in changes:
                changes_text.insert("end", line + "\n", tag)
        changes_text.configure(state="disabled")

        changes_scroll = tk.Scrollbar(text_shell, command=changes_text.yview)
        changes_scroll.grid(row=0, column=1, sticky="ns")
        changes_text.configure(yscrollcommand=changes_scroll.set)

        self._make_button(shell, UI_TEXT["close"], window.destroy, primary=True).grid(
            row=3,
            column=0,
            sticky="e",
            pady=(14, 0),
        )

    def restore_selected(self) -> None:
        entry = self._selected_history()
        if entry is None:
            self._set_status(UI_TEXT["select_history"], reset=True)
            return

        try:
            selected_text = self.store.read_history(entry)
        except Exception:
            self._set_status(UI_TEXT["history_load_error"], reset=True)
            return

        current_text = self._current_text()
        if current_text == selected_text:
            self._set_status(UI_TEXT["restore_same"], reset=True)
            return

        confirmed = messagebox.askokcancel(
            UI_TEXT["restore_confirm_title"],
            UI_TEXT["restore_confirm_message"],
            parent=self.root,
        )
        if not confirmed:
            return

        checkpoint = self._save_now(UI_TEXT["saved"], force_history=True, quiet=True)
        if not checkpoint.changed:
            self._set_status(UI_TEXT["save_error"], reset=True)
            return

        try:
            self.store.write_memo(selected_text)
            self.store.update_config(None)
        except Exception:
            self._set_status(UI_TEXT["save_error"], reset=True)
            return

        self.suppress_text_event = True
        try:
            self.memo_text.delete("1.0", "end")
            self.memo_text.insert("1.0", selected_text)
            self.memo_text.edit_modified(False)
        finally:
            self.suppress_text_event = False

        self.refresh_history()
        self._set_status(UI_TEXT["restore_complete"], reset=True)

    def _set_status(self, text: str, reset: bool) -> None:
        self.status_label.configure(text=text)
        if self.status_reset_id is not None:
            self.root.after_cancel(self.status_reset_id)
            self.status_reset_id = None
        if reset:
            self.status_reset_id = self.root.after(STATUS_RESET_MS, self._reset_status)

    def _reset_status(self) -> None:
        self.status_reset_id = None
        self.status_label.configure(text=UI_TEXT["no_changes"])

    def request_close(self) -> None:
        if self.autosave_id is not None:
            self.root.after_cancel(self.autosave_id)
            self.autosave_id = None
        self._save_now(UI_TEXT["saved"], quiet=True)
        self.root.destroy()


def run_smoke_test() -> int:
    with tempfile.TemporaryDirectory() as temp_dir:
        store = MemoStore(Path(temp_dir))
        if not store.memo_path.exists() or not store.config_path.exists() or not store.history_dir.exists():
            raise AssertionError("data files were not created")

        first = store.save_if_changed("ひとつめ")
        if not first.changed:
            raise AssertionError("first save did not create history")

        repeat = store.save_if_changed("ひとつめ")
        if repeat.changed:
            raise AssertionError("unchanged save created history")

        second = store.save_if_changed("ひとつめ\nふたつめ")
        if not second.changed:
            raise AssertionError("changed save did not create history")

        entries = store.history_entries()
        if len(entries) != 2:
            raise AssertionError("history count mismatch")

        checkpoint = store.save_if_changed("復元前の本文", force_history=True)
        if not checkpoint.changed:
            raise AssertionError("restore checkpoint was not saved")
        store.write_memo(store.read_history(entries[-1]))

        if len(store.history_entries()) != 3:
            raise AssertionError("restore checkpoint count mismatch")

        changes = build_change_lines("今日は最悪だった", "風呂に入ったら少し戻った")
        if ("delete", "- 今日は最悪だった") not in changes:
            raise AssertionError("delete line was not detected")
        if ("add", "+ 風呂に入ったら少し戻った") not in changes:
            raise AssertionError("add line was not detected")

    print("smoke ok")
    return 0


def run_launch_check() -> int:
    with tempfile.TemporaryDirectory() as temp_dir:
        set_windows_app_id()
        root = tk.Tk()
        app = DakeGitMemoApp(root, MemoStore(Path(temp_dir)))
        root.after(700, app.request_close)
        root.mainloop()
    return 0


def main() -> None:
    if "--smoke-test" in sys.argv:
        raise SystemExit(run_smoke_test())
    if "--launch-check" in sys.argv:
        raise SystemExit(run_launch_check())

    set_windows_app_id()
    root = tk.Tk()
    DakeGitMemoApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
