# -*- coding: utf-8 -*-
from __future__ import annotations

import copy
import ctypes
import json
import queue
import socket
import sys
import threading
import uuid
import webbrowser
from datetime import datetime, timezone
from pathlib import Path
import tkinter as tk
from tkinter import font as tkfont


APP_NAME = "Dake二人メモ"
WINDOW_TITLE = APP_NAME
COPYRIGHT = "© 2026 しまりす不動産 / Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "main_title": APP_NAME,
    "main_description": "右の人、左の人。互いのメモを奪わず、確認だけ返せます。",
    "label_role": "役割",
    "label_status": "状態",
    "label_host": "ホスト",
    "label_join": "参加先",
    "label_port": "ポート",
    "host_not_started": "未開始",
    "role_left": "左の人",
    "role_right": "右の人",
    "status_disconnected": "未接続",
    "status_hosting": "ホスト中",
    "status_connecting": "接続中",
    "status_syncing": "同期中",
    "status_cut": "切断",
    "status_saved": "保存済",
    "status_save_failed": "保存できません",
    "status_load_failed": "保存を読めません",
    "status_host_failed": "ホスト開始できません",
    "status_join_failed": "参加できません",
    "status_port_failed": "ポートを確認してください",
    "button_start_host": "ホスト開始",
    "button_join": "参加",
    "button_disconnect": "切断",
    "button_add_left": "左のメモを追加",
    "button_add_right": "右のメモを追加",
    "button_delete": "削除",
    "column_left": "左の人",
    "column_right": "右の人",
    "empty_left": "左のメモを追加できます",
    "empty_right": "右のメモを追加できます",
    "owner_label": "本文",
    "action_label": "相手の確認",
    "action_seen": "見た",
    "action_check": "確認",
    "action_done": "完了",
    "action_hold": "保留",
    "action_clear": "取消",
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
    "surface_soft": "#F9FAFB",
    "line": "#E6EAF0",
    "line_strong": "#D5DAE3",
    "text": "#1E2430",
    "muted": "#667085",
    "quiet": "#98A2B3",
    "button": "#2F6FED",
    "button_hover": "#2458BF",
    "button_text": "#FFFFFF",
    "secondary_button": "#FFFFFF",
    "secondary_hover": "#F2F4F7",
    "disabled": "#EEF1F5",
    "success": "#2E7D32",
    "danger": "#B42318",
    "action": "#EAF2FF",
}

FONT_CANDIDATES = ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo")
WINDOW_SIZE = "1120x720"
WINDOW_MIN_SIZE = (860, 600)
APP_USER_MODEL_ID = "Shimarisu.DakeTwoPersonMemo"
DATA_FILE_NAME = "two_person_memo_data.json"
DEFAULT_PORT = "8765"
ROLES = ("left", "right")
ACTION_OPTIONS = (
    UI_TEXT["action_seen"],
    UI_TEXT["action_check"],
    UI_TEXT["action_done"],
    UI_TEXT["action_hold"],
)
ALLOWED_ACTIONS = set(ACTION_OPTIONS) | {""}
FOOTER_NARROW_WIDTH = 900


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="milliseconds")


def is_newer(candidate: str | None, current: str | None) -> bool:
    return str(candidate or "") > str(current or "")


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def data_path() -> Path:
    return app_dir() / DATA_FILE_NAME


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


def set_windows_app_id() -> None:
    if not sys.platform.startswith("win"):
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)
    except Exception:
        return


def choose_font(root: tk.Tk) -> str:
    available = set(tkfont.families(root))
    for family in FONT_CANDIDATES:
        if family in available:
            return family
    return "TkDefaultFont"


def role_label(role: str) -> str:
    return UI_TEXT["role_left"] if role == "left" else UI_TEXT["role_right"]


def other_role(role: str) -> str:
    return "right" if role == "left" else "left"


def safe_action(value: object) -> str:
    if not isinstance(value, str):
        return ""
    return value if value in ALLOWED_ACTIONS else ""


def normalize_block(raw: object) -> dict[str, object] | None:
    if not isinstance(raw, dict):
        return None
    owner = raw.get("owner")
    if owner not in ROLES:
        return None
    created_at = str(raw.get("created_at") or now_iso())
    updated_at = str(raw.get("updated_at") or created_at)
    actions_raw = raw.get("actions")
    actions = actions_raw if isinstance(actions_raw, dict) else {}
    return {
        "id": str(raw.get("id") or uuid.uuid4().hex),
        "owner": owner,
        "text": str(raw.get("text") or ""),
        "created_at": created_at,
        "updated_at": updated_at,
        "actions": {
            "left": safe_action(actions.get("left")),
            "right": safe_action(actions.get("right")),
        },
    }


def new_block(owner: str) -> dict[str, object]:
    timestamp = now_iso()
    return {
        "id": uuid.uuid4().hex,
        "owner": owner,
        "text": "",
        "created_at": timestamp,
        "updated_at": timestamp,
        "actions": {"left": "", "right": ""},
    }


def guess_local_ip() -> str:
    try:
        for ip_address in socket.gethostbyname_ex(socket.gethostname())[2]:
            if ip_address and not ip_address.startswith("127."):
                return ip_address
    except OSError:
        pass

    probe = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        probe.connect(("192.0.2.1", 80))
        ip_address = probe.getsockname()[0]
        if ip_address:
            return ip_address
    except OSError:
        pass
    finally:
        probe.close()
    return "127.0.0.1"


class TwoPersonMemoApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.font_family = choose_font(root)
        self.blocks: list[dict[str, object]] = []
        self.deleted_blocks: dict[str, str] = {}
        self.incoming: queue.Queue[tuple[str, object]] = queue.Queue()
        self.socket_lock = threading.Lock()
        self.stop_event = threading.Event()
        self.server_socket: socket.socket | None = None
        self.client_socket: socket.socket | None = None
        self.peer_sockets: list[socket.socket] = []
        self.pending_save_after: str | None = None
        self.is_rendering = False
        self.footer: tk.Frame | None = None
        self.footer_mode: str | None = None

        self.role_var = tk.StringVar(value="left")
        self.status_var = tk.StringVar(value=UI_TEXT["status_disconnected"])
        self.remote_host_var = tk.StringVar(value="127.0.0.1")
        self.port_var = tk.StringVar(value=DEFAULT_PORT)
        self.host_info_var = tk.StringVar(value=UI_TEXT["host_not_started"])

        self.fonts = {
            "title": tkfont.Font(root, family=self.font_family, size=20, weight="bold"),
            "description": tkfont.Font(root, family=self.font_family, size=10),
            "label": tkfont.Font(root, family=self.font_family, size=10, weight="bold"),
            "body": tkfont.Font(root, family=self.font_family, size=10),
            "button": tkfont.Font(root, family=self.font_family, size=10, weight="bold"),
            "small_button": tkfont.Font(root, family=self.font_family, size=9),
            "small": tkfont.Font(root, family=self.font_family, size=9),
            "footer": tkfont.Font(root, family=self.font_family, size=8),
        }

        self._load_data()
        self.current_role = self.role_var.get()
        self._configure_root()
        self._build_ui()
        self._apply_icon()
        self._render_blocks()
        self.role_var.trace_add("write", self._on_role_change)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(120, self._process_incoming)

    def _configure_root(self) -> None:
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(*WINDOW_MIN_SIZE)
        self.root.configure(bg=COLORS["background"])

    def _apply_icon(self) -> None:
        for candidate in icon_candidates():
            try:
                icon_path = candidate.resolve()
            except OSError:
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
        outer.pack(fill="both", expand=True, padx=24, pady=(20, 16))

        self._build_header(outer)
        self._build_connection_bar(outer)
        self._build_memo_area(outer)
        self._build_footer(outer)
        self.root.bind("<Configure>", self._on_root_configure, add="+")

    def _build_header(self, parent: tk.Widget) -> None:
        header = tk.Frame(parent, bg=COLORS["background"])
        header.pack(fill="x", pady=(0, 12))
        header.grid_columnconfigure(0, weight=1)

        title_area = tk.Frame(header, bg=COLORS["background"])
        title_area.grid(row=0, column=0, sticky="w")

        tk.Label(
            title_area,
            text=UI_TEXT["main_title"],
            bg=COLORS["background"],
            fg=COLORS["text"],
            font=self.fonts["title"],
            anchor="w",
        ).pack(anchor="w")

        tk.Label(
            title_area,
            text=UI_TEXT["main_description"],
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=self.fonts["description"],
            anchor="w",
        ).pack(anchor="w", pady=(5, 0))

        role_box = tk.Frame(header, bg=COLORS["background"])
        role_box.grid(row=0, column=1, sticky="e", padx=(20, 0))

        tk.Label(
            role_box,
            text=UI_TEXT["label_role"],
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=self.fonts["small"],
        ).pack(side="left", padx=(0, 8))

        for role in ROLES:
            tk.Radiobutton(
                role_box,
                text=role_label(role),
                value=role,
                variable=self.role_var,
                bg=COLORS["background"],
                fg=COLORS["text"],
                activebackground=COLORS["background"],
                activeforeground=COLORS["text"],
                selectcolor=COLORS["surface"],
                font=self.fonts["body"],
                padx=4,
            ).pack(side="left", padx=(0, 6))

    def _build_connection_bar(self, parent: tk.Widget) -> None:
        bar = tk.Frame(
            parent,
            bg=COLORS["surface"],
            highlightbackground=COLORS["line"],
            highlightthickness=1,
        )
        bar.pack(fill="x", pady=(0, 14))
        bar.grid_columnconfigure(8, weight=1)

        status_box = tk.Frame(bar, bg=COLORS["surface"], padx=12, pady=10)
        status_box.grid(row=0, column=0, sticky="w")

        tk.Label(
            status_box,
            text=UI_TEXT["label_status"],
            bg=COLORS["surface"],
            fg=COLORS["muted"],
            font=self.fonts["small"],
        ).pack(side="left", padx=(0, 6))

        tk.Label(
            status_box,
            textvariable=self.status_var,
            bg=COLORS["surface"],
            fg=COLORS["button"],
            font=self.fonts["label"],
        ).pack(side="left")

        self._make_button(
            bar,
            UI_TEXT["button_start_host"],
            self.start_host,
            primary=True,
        ).grid(row=0, column=1, sticky="w", padx=(2, 8), pady=8)

        tk.Label(
            bar,
            text=UI_TEXT["label_host"],
            bg=COLORS["surface"],
            fg=COLORS["muted"],
            font=self.fonts["small"],
        ).grid(row=0, column=2, sticky="e", padx=(4, 4))

        tk.Label(
            bar,
            textvariable=self.host_info_var,
            bg=COLORS["surface"],
            fg=COLORS["text"],
            font=self.fonts["small"],
            width=18,
            anchor="w",
        ).grid(row=0, column=3, sticky="w", padx=(0, 12))

        tk.Label(
            bar,
            text=UI_TEXT["label_join"],
            bg=COLORS["surface"],
            fg=COLORS["muted"],
            font=self.fonts["small"],
        ).grid(row=0, column=4, sticky="e", padx=(0, 4))

        tk.Entry(
            bar,
            textvariable=self.remote_host_var,
            bg=COLORS["surface_soft"],
            fg=COLORS["text"],
            insertbackground=COLORS["text"],
            relief="flat",
            highlightthickness=1,
            highlightbackground=COLORS["line"],
            highlightcolor=COLORS["button"],
            font=self.fonts["body"],
            width=16,
        ).grid(row=0, column=5, sticky="w", padx=(0, 8), ipady=5)

        tk.Label(
            bar,
            text=UI_TEXT["label_port"],
            bg=COLORS["surface"],
            fg=COLORS["muted"],
            font=self.fonts["small"],
        ).grid(row=0, column=6, sticky="e", padx=(0, 4))

        tk.Entry(
            bar,
            textvariable=self.port_var,
            bg=COLORS["surface_soft"],
            fg=COLORS["text"],
            insertbackground=COLORS["text"],
            relief="flat",
            highlightthickness=1,
            highlightbackground=COLORS["line"],
            highlightcolor=COLORS["button"],
            font=self.fonts["body"],
            width=6,
        ).grid(row=0, column=7, sticky="w", padx=(0, 8), ipady=5)

        actions = tk.Frame(bar, bg=COLORS["surface"])
        actions.grid(row=0, column=9, sticky="e", padx=(8, 10), pady=8)

        self._make_button(
            actions,
            UI_TEXT["button_join"],
            self.join_host,
            primary=False,
        ).pack(side="left", padx=(0, 8))

        self._make_button(
            actions,
            UI_TEXT["button_disconnect"],
            self.disconnect_network,
            primary=False,
        ).pack(side="left")

    def _build_memo_area(self, parent: tk.Widget) -> None:
        shell = tk.Frame(
            parent,
            bg=COLORS["surface"],
            highlightbackground=COLORS["line"],
            highlightthickness=1,
        )
        shell.pack(fill="both", expand=True)
        shell.grid_columnconfigure(0, weight=1)
        shell.grid_columnconfigure(1, weight=1)
        shell.grid_rowconfigure(1, weight=1)

        tk.Label(
            shell,
            text=UI_TEXT["column_left"],
            bg=COLORS["surface_soft"],
            fg=COLORS["text"],
            font=self.fonts["label"],
            pady=10,
        ).grid(row=0, column=0, sticky="ew")

        tk.Label(
            shell,
            text=UI_TEXT["column_right"],
            bg=COLORS["surface_soft"],
            fg=COLORS["text"],
            font=self.fonts["label"],
            pady=10,
        ).grid(row=0, column=1, sticky="ew")

        separator = tk.Frame(shell, bg=COLORS["line"], width=1)
        separator.grid(row=0, column=1, rowspan=2, sticky="nsw")

        body = tk.Frame(shell, bg=COLORS["surface"])
        body.grid(row=1, column=0, columnspan=2, sticky="nsew")
        body.grid_rowconfigure(0, weight=1)
        body.grid_columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(
            body,
            bg=COLORS["surface"],
            bd=0,
            highlightthickness=0,
            relief="flat",
        )
        self.canvas.grid(row=0, column=0, sticky="nsew")

        scrollbar = tk.Scrollbar(body, orient="vertical", command=self.canvas.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.table_frame = tk.Frame(self.canvas, bg=COLORS["surface"])
        self.table_window = self.canvas.create_window(
            0,
            0,
            window=self.table_frame,
            anchor="nw",
        )
        self.table_frame.bind("<Configure>", self._on_table_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind("<Enter>", self._bind_mousewheel)
        self.canvas.bind("<Leave>", self._unbind_mousewheel)

    def _build_footer(self, parent: tk.Widget) -> None:
        self.footer = tk.Frame(parent, bg=COLORS["background"])
        self.footer.pack(fill="x", pady=(12, 0))
        self._render_footer("wide")

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

    def _make_button(self, parent: tk.Widget, label: str, command, primary: bool) -> tk.Button:
        bg = COLORS["button"] if primary else COLORS["secondary_button"]
        fg = COLORS["button_text"] if primary else COLORS["text"]
        hover = COLORS["button_hover"] if primary else COLORS["secondary_hover"]
        border = COLORS["button"] if primary else COLORS["line_strong"]
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
            padx=14,
            pady=7,
            cursor="hand2",
            highlightthickness=1,
            highlightbackground=border,
        )
        button.bind("<Enter>", lambda _event, target=button, color=hover: self._hover_button(target, color))
        button.bind("<Leave>", lambda _event, target=button, color=bg: self._hover_button(target, color))
        return button

    def _make_action_button(
        self,
        parent: tk.Widget,
        label: str,
        command,
        selected: bool = False,
    ) -> tk.Button:
        bg = COLORS["button"] if selected else COLORS["secondary_button"]
        fg = COLORS["button_text"] if selected else COLORS["text"]
        hover = COLORS["button_hover"] if selected else COLORS["secondary_hover"]
        border = COLORS["button"] if selected else COLORS["line_strong"]
        button = tk.Button(
            parent,
            text=label,
            command=command,
            bg=bg,
            fg=fg,
            activebackground=hover,
            activeforeground=fg,
            font=self.fonts["small_button"],
            relief="flat",
            bd=0,
            padx=10,
            pady=4,
            cursor="hand2",
            highlightthickness=1,
            highlightbackground=border,
        )
        button.bind("<Enter>", lambda _event, target=button, color=hover: self._hover_button(target, color))
        button.bind("<Leave>", lambda _event, target=button, color=bg: self._hover_button(target, color))
        return button

    def _hover_button(self, button: tk.Button, color: str) -> None:
        if str(button.cget("state")) == "disabled":
            return
        button.configure(bg=color)

    def _on_table_configure(self, _event: tk.Event) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event: tk.Event) -> None:
        self.canvas.itemconfigure(self.table_window, width=event.width)

    def _bind_mousewheel(self, _event: tk.Event) -> None:
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _unbind_mousewheel(self, _event: tk.Event) -> None:
        self.canvas.unbind_all("<MouseWheel>")

    def _on_mousewheel(self, event: tk.Event) -> None:
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _render_blocks(self) -> None:
        self.is_rendering = True
        for child in self.table_frame.winfo_children():
            child.destroy()

        self.table_frame.grid_columnconfigure(0, weight=1, uniform="memo_column")
        self.table_frame.grid_columnconfigure(1, weight=1, uniform="memo_column")

        sorted_blocks = sorted(self.blocks, key=lambda item: str(item.get("created_at", "")))

        if not sorted_blocks:
            self._render_empty_row(0)
            next_row = 1
        else:
            next_row = 0
            for block in sorted_blocks:
                self._render_block_row(next_row, block)
                next_row += 1

        self._render_add_row(next_row)
        self.is_rendering = False

    def _render_empty_row(self, row: int) -> None:
        left_cell = self._make_cell(row, 0)
        right_cell = self._make_cell(row, 1)
        self._empty_label(left_cell, UI_TEXT["empty_left"])
        self._empty_label(right_cell, UI_TEXT["empty_right"])

    def _render_block_row(self, row: int, block: dict[str, object]) -> None:
        left_cell = self._make_cell(row, 0)
        right_cell = self._make_cell(row, 1)
        if block.get("owner") == "left":
            self._build_block_card(left_cell, block)
            self._blank_cell(right_cell)
        else:
            self._blank_cell(left_cell)
            self._build_block_card(right_cell, block)

    def _render_add_row(self, row: int) -> None:
        left_cell = self._make_cell(row, 0)
        right_cell = self._make_cell(row, 1)
        role = self.role_var.get()

        left_button = self._make_button(
            left_cell,
            UI_TEXT["button_add_left"],
            lambda: self.add_block("left"),
            primary=role == "left",
        )
        left_button.pack(anchor="w")
        if role != "left":
            left_button.configure(state="disabled", bg=COLORS["disabled"], cursor="arrow")

        right_button = self._make_button(
            right_cell,
            UI_TEXT["button_add_right"],
            lambda: self.add_block("right"),
            primary=role == "right",
        )
        right_button.pack(anchor="e")
        if role != "right":
            right_button.configure(state="disabled", bg=COLORS["disabled"], cursor="arrow")

    def _make_cell(self, row: int, column: int) -> tk.Frame:
        cell = tk.Frame(
            self.table_frame,
            bg=COLORS["surface"],
            padx=14,
            pady=10,
        )
        cell.grid(row=row, column=column, sticky="nsew")
        return cell

    def _empty_label(self, parent: tk.Widget, value: str) -> None:
        tk.Label(
            parent,
            text=value,
            bg=COLORS["surface"],
            fg=COLORS["quiet"],
            font=self.fonts["body"],
            pady=50,
        ).pack(expand=True)

    def _blank_cell(self, parent: tk.Widget) -> None:
        tk.Frame(parent, bg=COLORS["surface"], height=1).pack(fill="both", expand=True)

    def _build_block_card(self, parent: tk.Widget, block: dict[str, object]) -> None:
        block_id = str(block["id"])
        owner = str(block["owner"])
        role = self.role_var.get()
        is_owner = owner == role

        card = tk.Frame(
            parent,
            bg=COLORS["surface_soft"] if is_owner else COLORS["surface"],
            highlightbackground=COLORS["line_strong"],
            highlightthickness=1,
            padx=12,
            pady=10,
        )
        card.pack(fill="both", expand=True)
        card.grid_columnconfigure(0, weight=1)

        top = tk.Frame(card, bg=card.cget("bg"))
        top.grid(row=0, column=0, sticky="ew")
        top.grid_columnconfigure(0, weight=1)

        tk.Label(
            top,
            text=role_label(owner),
            bg=card.cget("bg"),
            fg=COLORS["muted"],
            font=self.fonts["small"],
            anchor="w",
        ).grid(row=0, column=0, sticky="w")

        if is_owner:
            self._make_action_button(
                top,
                UI_TEXT["button_delete"],
                lambda target_id=block_id: self.delete_block(target_id),
            ).grid(row=0, column=1, sticky="e")

        text_widget = tk.Text(
            card,
            bg=COLORS["surface"] if is_owner else COLORS["surface_soft"],
            fg=COLORS["text"],
            insertbackground=COLORS["text"],
            selectbackground="#D7E7FF",
            selectforeground=COLORS["text"],
            inactiveselectbackground="#EAF2FF",
            font=self.fonts["body"],
            relief="flat",
            bd=0,
            padx=10,
            pady=9,
            wrap="word",
            undo=is_owner,
            height=self._text_height(str(block.get("text") or "")),
        )
        text_widget.grid(row=1, column=0, sticky="ew", pady=(8, 8))
        text_widget.insert("1.0", str(block.get("text") or ""))
        if is_owner:
            text_widget.edit_modified(False)
            text_widget.bind(
                "<<Modified>>",
                lambda _event, target_id=block_id, widget=text_widget: self._on_text_modified(target_id, widget),
            )
        else:
            text_widget.configure(state="disabled", cursor="arrow")

        if is_owner:
            self._build_action_summary(card, block, other_role(owner))
        else:
            self._build_action_buttons(card, block)

    def _text_height(self, text: str) -> int:
        line_count = max(3, text.count("\n") + 3)
        return min(8, line_count)

    def _build_action_summary(self, parent: tk.Widget, block: dict[str, object], actor: str) -> None:
        actions = block.get("actions")
        action = ""
        if isinstance(actions, dict):
            action = safe_action(actions.get(actor))
        value = f"{role_label(actor)}: {action}" if action else UI_TEXT["action_label"]
        tk.Label(
            parent,
            text=value,
            bg=parent.cget("bg"),
            fg=COLORS["muted"] if action else COLORS["quiet"],
            font=self.fonts["small"],
            anchor="w",
        ).grid(row=2, column=0, sticky="w")

    def _build_action_buttons(self, parent: tk.Widget, block: dict[str, object]) -> None:
        block_id = str(block["id"])
        role = self.role_var.get()
        actions = block.get("actions")
        current_action = ""
        if isinstance(actions, dict):
            current_action = safe_action(actions.get(role))

        action_row = tk.Frame(parent, bg=parent.cget("bg"))
        action_row.grid(row=2, column=0, sticky="w")

        for action in ACTION_OPTIONS:
            self._make_action_button(
                action_row,
                action,
                lambda value=action, target_id=block_id: self.set_action(target_id, value),
                selected=current_action == action,
            ).pack(side="left", padx=(0, 6), pady=(0, 2))

        self._make_action_button(
            action_row,
            UI_TEXT["action_clear"],
            lambda target_id=block_id: self.set_action(target_id, ""),
        ).pack(side="left", pady=(0, 2))

    def _on_role_change(self, *_args: object) -> None:
        self.current_role = self.role_var.get()
        self._render_blocks()
        self._save_data()

    def add_block(self, owner: str) -> None:
        if owner != self.role_var.get():
            return
        self.blocks.append(new_block(owner))
        self._render_blocks()
        self._save_and_sync()

    def delete_block(self, block_id: str) -> None:
        block = self._find_block(block_id)
        if block is None or block.get("owner") != self.role_var.get():
            return
        self.blocks = [item for item in self.blocks if str(item.get("id")) != block_id]
        self.deleted_blocks[block_id] = now_iso()
        self._render_blocks()
        self._save_and_sync()

    def set_action(self, block_id: str, action: str) -> None:
        block = self._find_block(block_id)
        role = self.role_var.get()
        if block is None or block.get("owner") == role:
            return
        safe_value = safe_action(action)
        actions = block.get("actions")
        if not isinstance(actions, dict):
            actions = {"left": "", "right": ""}
            block["actions"] = actions
        actions[role] = safe_value
        block["updated_at"] = now_iso()
        self._render_blocks()
        self._save_and_sync()

    def _on_text_modified(self, block_id: str, widget: tk.Text) -> None:
        if self.is_rendering or not widget.edit_modified():
            return
        widget.edit_modified(False)
        block = self._find_block(block_id)
        if block is None or block.get("owner") != self.role_var.get():
            return
        block["text"] = widget.get("1.0", "end-1c")
        block["updated_at"] = now_iso()
        self._schedule_save_and_sync()

    def _find_block(self, block_id: str) -> dict[str, object] | None:
        for block in self.blocks:
            if str(block.get("id")) == block_id:
                return block
        return None

    def _load_data(self) -> None:
        path = data_path()
        if not path.exists():
            return
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            self.status_var.set(UI_TEXT["status_load_failed"])
            return

        blocks_raw = payload.get("blocks") if isinstance(payload, dict) else []
        if isinstance(blocks_raw, list):
            self.blocks = [
                normalized
                for normalized in (normalize_block(item) for item in blocks_raw)
                if normalized is not None
            ]

        deleted_raw = payload.get("deleted_blocks") if isinstance(payload, dict) else {}
        if isinstance(deleted_raw, dict):
            self.deleted_blocks = {
                str(block_id): str(timestamp)
                for block_id, timestamp in deleted_raw.items()
                if isinstance(block_id, str)
            }

        role = payload.get("last_role") if isinstance(payload, dict) else None
        if role in ROLES:
            self.role_var.set(str(role))

        last_host = payload.get("last_host") if isinstance(payload, dict) else None
        if isinstance(last_host, str) and last_host:
            self.remote_host_var.set(last_host)

        last_port = payload.get("last_port") if isinstance(payload, dict) else None
        if isinstance(last_port, (str, int)):
            self.port_var.set(str(last_port))

    def _save_data(self) -> None:
        payload = {
            "blocks": self.blocks,
            "deleted_blocks": self.deleted_blocks,
            "last_role": self.role_var.get(),
            "last_host": self.remote_host_var.get().strip(),
            "last_port": self.port_var.get().strip() or DEFAULT_PORT,
        }
        path = data_path()
        tmp_path = path.with_suffix(".json.tmp")
        try:
            tmp_path.write_text(
                json.dumps(payload, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            tmp_path.replace(path)
        except OSError:
            self._set_status_key("status_save_failed")

    def _schedule_save_and_sync(self) -> None:
        if self.pending_save_after is not None:
            self.root.after_cancel(self.pending_save_after)
        self.pending_save_after = self.root.after(450, self._flush_save_and_sync)

    def _flush_save_and_sync(self) -> None:
        self.pending_save_after = None
        self._save_and_sync()

    def _save_and_sync(self) -> None:
        self._save_data()
        self._send_state()

    def _read_port(self) -> int | None:
        try:
            port = int(self.port_var.get().strip())
        except ValueError:
            self._set_status_key("status_port_failed")
            return None
        if port < 1 or port > 65535:
            self._set_status_key("status_port_failed")
            return None
        return port

    def start_host(self) -> None:
        port = self._read_port()
        if port is None:
            return

        self.disconnect_network(update_status=False)
        self.stop_event = threading.Event()
        self.host_info_var.set(f"{guess_local_ip()}:{port}")
        self._set_status_key("status_hosting")
        self._save_data()

        thread = threading.Thread(
            target=self._host_worker,
            args=(port, self.stop_event),
            daemon=True,
        )
        thread.start()

    def _host_worker(self, port: int, stop_event: threading.Event) -> None:
        server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            server.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            server.bind(("", port))
            server.listen(4)
            server.settimeout(0.5)
            with self.socket_lock:
                self.server_socket = server
            self.incoming.put(("host_ready", port))

            while not stop_event.is_set():
                try:
                    peer, _address = server.accept()
                except socket.timeout:
                    continue
                except OSError:
                    break
                peer.settimeout(0.5)
                with self.socket_lock:
                    self.peer_sockets.append(peer)
                self._send_packet(peer, self._state_packet())
                self.incoming.put(("status_key", "status_syncing"))
                threading.Thread(
                    target=self._reader_worker,
                    args=(peer, stop_event, True),
                    daemon=True,
                ).start()
        except OSError:
            self.incoming.put(("status_key", "status_host_failed"))
        finally:
            try:
                server.close()
            except OSError:
                pass

    def join_host(self) -> None:
        port = self._read_port()
        host = self.remote_host_var.get().strip()
        if port is None or not host:
            self._set_status_key("status_join_failed")
            return

        self.disconnect_network(update_status=False)
        self.stop_event = threading.Event()
        self._set_status_key("status_connecting")
        self._save_data()

        thread = threading.Thread(
            target=self._client_worker,
            args=(host, port, self.stop_event),
            daemon=True,
        )
        thread.start()

    def _client_worker(self, host: str, port: int, stop_event: threading.Event) -> None:
        client = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            client.settimeout(5)
            client.connect((host, port))
            client.settimeout(0.5)
            with self.socket_lock:
                self.client_socket = client
            self.incoming.put(("status_key", "status_syncing"))
            self._send_packet(client, self._state_packet())
            self._reader_worker(client, stop_event, False)
        except OSError:
            try:
                client.close()
            except OSError:
                pass
            self.incoming.put(("status_key", "status_join_failed"))

    def disconnect_network(self, update_status: bool = True) -> None:
        self.stop_event.set()
        sockets: list[socket.socket] = []
        with self.socket_lock:
            if self.server_socket is not None:
                sockets.append(self.server_socket)
            if self.client_socket is not None:
                sockets.append(self.client_socket)
            sockets.extend(self.peer_sockets)
            self.server_socket = None
            self.client_socket = None
            self.peer_sockets = []

        for sock in sockets:
            self._close_socket(sock)

        if update_status:
            self.host_info_var.set(UI_TEXT["host_not_started"])
            self._set_status_key("status_cut")

    def _reader_worker(self, sock: socket.socket, stop_event: threading.Event, from_peer: bool) -> None:
        buffer = b""
        try:
            while not stop_event.is_set():
                try:
                    chunk = sock.recv(65536)
                except socket.timeout:
                    continue
                except OSError:
                    break
                if not chunk:
                    break
                buffer += chunk
                while b"\n" in buffer:
                    line, buffer = buffer.split(b"\n", 1)
                    if not line.strip():
                        continue
                    try:
                        packet = json.loads(line.decode("utf-8"))
                    except (UnicodeDecodeError, json.JSONDecodeError):
                        continue
                    kind = "remote_state_from_peer" if from_peer else "remote_state"
                    self.incoming.put((kind, packet))
        finally:
            self._remove_socket(sock)
            self._close_socket(sock)
            if not stop_event.is_set():
                next_status = "status_hosting" if from_peer and self.server_socket is not None else "status_cut"
                self.incoming.put(("status_key", next_status))

    def _close_socket(self, sock: socket.socket) -> None:
        try:
            sock.shutdown(socket.SHUT_RDWR)
        except OSError:
            pass
        try:
            sock.close()
        except OSError:
            pass

    def _remove_socket(self, sock: socket.socket) -> None:
        with self.socket_lock:
            if self.client_socket is sock:
                self.client_socket = None
            self.peer_sockets = [item for item in self.peer_sockets if item is not sock]

    def _state_packet(self) -> dict[str, object]:
        return {
            "type": "state",
            "role": self.current_role,
            "sent_at": now_iso(),
            "blocks": copy.deepcopy(self.blocks),
            "deleted_blocks": copy.deepcopy(self.deleted_blocks),
        }

    def _send_state(self) -> None:
        packet = self._state_packet()
        targets: list[socket.socket] = []
        with self.socket_lock:
            if self.server_socket is not None:
                targets = list(self.peer_sockets)
            elif self.client_socket is not None:
                targets = [self.client_socket]

        for sock in targets:
            self._send_packet(sock, packet)

    def _send_packet(self, sock: socket.socket, packet: dict[str, object]) -> None:
        try:
            payload = json.dumps(packet, ensure_ascii=False, separators=(",", ":")).encode("utf-8") + b"\n"
            sock.sendall(payload)
        except OSError:
            self._remove_socket(sock)
            self._close_socket(sock)

    def _process_incoming(self) -> None:
        try:
            while True:
                kind, payload = self.incoming.get_nowait()
                if kind == "status_key" and isinstance(payload, str):
                    self._set_status_key(payload)
                elif kind == "host_ready" and isinstance(payload, int):
                    self.host_info_var.set(f"{guess_local_ip()}:{payload}")
                    self._set_status_key("status_hosting")
                elif kind == "remote_state":
                    self._merge_remote_packet(payload)
                elif kind == "remote_state_from_peer":
                    changed = self._merge_remote_packet(payload)
                    if changed:
                        self._send_state()
        except queue.Empty:
            pass
        self.root.after(120, self._process_incoming)

    def _merge_remote_packet(self, packet: object) -> bool:
        if not isinstance(packet, dict) or packet.get("type") != "state":
            return False

        changed = False
        remote_deleted = packet.get("deleted_blocks")
        if isinstance(remote_deleted, dict):
            for block_id, timestamp in remote_deleted.items():
                if not isinstance(block_id, str):
                    continue
                timestamp_str = str(timestamp)
                if is_newer(timestamp_str, self.deleted_blocks.get(block_id)):
                    self.deleted_blocks[block_id] = timestamp_str
                    changed = True
                local_block = self._find_block(block_id)
                if local_block is not None and is_newer(timestamp_str, str(local_block.get("updated_at"))):
                    self.blocks = [item for item in self.blocks if str(item.get("id")) != block_id]
                    changed = True

        remote_blocks_raw = packet.get("blocks")
        if not isinstance(remote_blocks_raw, list):
            return changed

        local_role = self.current_role
        for raw_block in remote_blocks_raw:
            remote_block = normalize_block(raw_block)
            if remote_block is None:
                continue

            block_id = str(remote_block["id"])
            deleted_at = self.deleted_blocks.get(block_id)
            if deleted_at and not is_newer(str(remote_block["updated_at"]), deleted_at):
                continue

            local_block = self._find_block(block_id)
            if local_block is None:
                self.blocks.append(remote_block)
                changed = True
                continue

            if not is_newer(str(remote_block["updated_at"]), str(local_block.get("updated_at"))):
                continue

            merged_block = self._merge_block_fields(local_block, remote_block, local_role)
            local_block.update(merged_block)
            changed = True

        if changed:
            self._set_status_key("status_syncing")
            self._save_data()
            self._render_blocks()
        return changed

    def _merge_block_fields(
        self,
        local_block: dict[str, object],
        remote_block: dict[str, object],
        local_role: str,
    ) -> dict[str, object]:
        owner = str(remote_block["owner"])
        merged_block = copy.deepcopy(remote_block)
        local_actions_raw = local_block.get("actions")
        remote_actions_raw = remote_block.get("actions")
        local_actions = local_actions_raw if isinstance(local_actions_raw, dict) else {}
        remote_actions = remote_actions_raw if isinstance(remote_actions_raw, dict) else {}

        if owner == local_role:
            merged_block["text"] = str(local_block.get("text") or "")
            merged_block["created_at"] = str(local_block.get("created_at") or remote_block.get("created_at"))

        merged_actions = {
            "left": safe_action(remote_actions.get("left")),
            "right": safe_action(remote_actions.get("right")),
        }
        merged_actions[local_role] = safe_action(local_actions.get(local_role))
        merged_block["actions"] = merged_actions
        return merged_block

    def _set_status_key(self, key: str) -> None:
        value = UI_TEXT.get(key)
        if value:
            self.status_var.set(value)

    def _on_close(self) -> None:
        if self.pending_save_after is not None:
            self.root.after_cancel(self.pending_save_after)
            self.pending_save_after = None
            self._save_and_sync()
        else:
            self._save_data()
        self.disconnect_network(update_status=False)
        self.root.destroy()


def main() -> None:
    set_windows_app_id()
    root = tk.Tk()
    TwoPersonMemoApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
