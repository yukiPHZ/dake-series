# -*- coding: utf-8 -*-
from __future__ import annotations

import ctypes
import json
import sys
import webbrowser
from datetime import date, datetime, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import font as tkfont, messagebox


APP_NAME = "Dake昨日タスクメモ"
WINDOW_TITLE = "昨日タスクメモ"
COPYRIGHT = "© 2026 しまりす不動産 / Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "main_title": "昨日を見る。今日を書く。",
    "main_description": "前営業日のメモを見ながら、今日やることだけを整理します。",
    "previous_title": "前営業日のメモ",
    "previous_description": "昨日の自分が残したメモです。",
    "today_title": "今日のメモ",
    "today_description": "整理しながら、今日やることを書き出します。",
    "date_previous": "対象日：{date}",
    "date_today": "今日：{date}",
    "previous_empty": "前営業日のメモはまだありません。",
    "button_save": "今日のメモを保存",
    "button_clear": "今日のメモをクリア",
    "status_ready": "入力すると自動保存します",
    "status_saved": "保存しました",
    "status_autosaved": "自動保存しました",
    "status_save_failed": "保存できませんでした",
    "status_cleared": "今日のメモをクリアしました",
    "status_no_previous": "前営業日のメモはまだありません",
    "dialog_clear_title": "確認",
    "dialog_clear_body": "今日のメモを空にします。よろしいですか？",
    "dialog_error_title": "確認してください",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_catchcopy": "止まらない、迷わない、すぐ終わる。",
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
    "base_bg": "#F6F7F9",
    "panel_bg": "#FFFFFF",
    "quiet_panel_bg": "#F1F3F6",
    "text_bg": "#FBFCFE",
    "quiet_text_bg": "#F7F8FA",
    "text": "#1E2430",
    "muted": "#667085",
    "quiet_text": "#475467",
    "border": "#E6EAF0",
    "quiet_border": "#DDE2EA",
    "accent": "#2F6FED",
    "accent_hover": "#2458BF",
    "button_sub_hover": "#EEF2F7",
    "error": "#D92D20",
}

FONT_CANDIDATES = ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo")
ICON_RELATIVE_PATH = Path("..") / ".." / "02_assets" / "dake_icon.ico"
SAVE_DEBOUNCE_MS = 650
STATUS_RESET_MS = 1800
WINDOW_APP_ID = "Shimarisu.DakeYesterdayTaskMemo"


def set_windows_app_id() -> None:
    if not sys.platform.startswith("win"):
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(WINDOW_APP_ID)
    except Exception:
        return


def app_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def memo_dir(base_dir: Path | None = None) -> Path:
    return (base_dir or app_base_dir()) / "data" / "memos"


def memo_path(target_date: date, base_dir: Path | None = None) -> Path:
    return memo_dir(base_dir) / f"{target_date.isoformat()}.json"


def is_business_day(target_date: date) -> bool:
    return target_date.weekday() < 5


def previous_business_day(today: date) -> date:
    candidate = today - timedelta(days=1)
    while not is_business_day(candidate):
        candidate -= timedelta(days=1)
    return candidate


def load_memo(target_date: date, base_dir: Path | None = None) -> str:
    path = memo_path(target_date, base_dir)
    if not path.exists():
        return ""
    try:
        with path.open("r", encoding="utf-8") as file:
            payload = json.load(file)
    except Exception:
        return ""
    if not isinstance(payload, dict):
        return ""
    memo = payload.get("memo")
    return memo if isinstance(memo, str) else ""


def save_memo(target_date: date, memo: str, base_dir: Path | None = None) -> None:
    target_dir = memo_dir(base_dir)
    target_dir.mkdir(parents=True, exist_ok=True)
    payload = {
        "date": target_date.isoformat(),
        "memo": memo,
        "updated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
    }
    with memo_path(target_date, base_dir).open("w", encoding="utf-8") as file:
        json.dump(payload, file, ensure_ascii=False, indent=2)
        file.write("\n")


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


def choose_font_family(root: tk.Tk) -> str:
    try:
        available = set(tkfont.families(root))
    except tk.TclError:
        available = set()
    for candidate in FONT_CANDIDATES:
        if candidate in available:
            return candidate
    return "TkDefaultFont"


def format_date_label(value: date) -> str:
    return value.isoformat()


class YesterdayTaskMemoApp:
    def __init__(self, root: tk.Tk, today: date | None = None) -> None:
        self.root = root
        self.today = today or date.today()
        self.previous_day = previous_business_day(self.today)
        self.font_family = choose_font_family(root)
        self.save_after_id: str | None = None
        self.status_after_id: str | None = None
        self.suppress_modified = False
        self.footer_mode = ""

        self.status_var = tk.StringVar(value=UI_TEXT["status_ready"])
        self.previous_date_var = tk.StringVar(
            value=UI_TEXT["date_previous"].format(date=format_date_label(self.previous_day))
        )
        self.today_date_var = tk.StringVar(
            value=UI_TEXT["date_today"].format(date=format_date_label(self.today))
        )

        self.root.title(WINDOW_TITLE)
        self.root.geometry("1040x640")
        self.root.minsize(860, 520)
        self.root.configure(bg=COLORS["base_bg"])
        apply_window_icon(self.root)

        self._build_fonts()
        self._build_ui()
        self._load_initial_memos()

        self.root.protocol("WM_DELETE_WINDOW", self._close)
        self.root.bind("<Configure>", self._handle_resize)
        self.root.after(80, lambda: self.today_text.focus_set())

    def _build_fonts(self) -> None:
        self.fonts = {
            "title": (self.font_family, 18, "bold"),
            "description": (self.font_family, 10),
            "section": (self.font_family, 12, "bold"),
            "small": (self.font_family, 9),
            "memo": (self.font_family, 11),
            "button": (self.font_family, 10, "bold"),
            "button_sub": (self.font_family, 10),
        }

    def _build_ui(self) -> None:
        outer = tk.Frame(self.root, bg=COLORS["base_bg"])
        outer.pack(fill="both", expand=True, padx=22, pady=(16, 10))

        self._build_header(outer)
        self._build_memo_area(outer)
        self._build_status(outer)
        self._build_footer(outer)

    def _build_header(self, parent: tk.Widget) -> None:
        header = tk.Frame(parent, bg=COLORS["base_bg"])
        header.pack(fill="x", pady=(0, 12))

        tk.Label(
            header,
            text=UI_TEXT["main_title"],
            bg=COLORS["base_bg"],
            fg=COLORS["text"],
            font=self.fonts["title"],
            anchor="w",
        ).pack(fill="x")
        description = tk.Label(
            header,
            text=UI_TEXT["main_description"],
            bg=COLORS["base_bg"],
            fg=COLORS["muted"],
            font=self.fonts["description"],
            anchor="w",
            justify="left",
        )
        description.pack(fill="x", pady=(3, 0))

    def _build_memo_area(self, parent: tk.Widget) -> None:
        area = tk.Frame(parent, bg=COLORS["base_bg"])
        area.pack(fill="both", expand=True)
        area.grid_columnconfigure(0, weight=42, uniform="memo")
        area.grid_columnconfigure(1, weight=58, uniform="memo")
        area.grid_rowconfigure(0, weight=1)

        self.previous_text = self._build_memo_panel(
            area,
            column=0,
            title=UI_TEXT["previous_title"],
            description=UI_TEXT["previous_description"],
            date_var=self.previous_date_var,
            panel_bg=COLORS["quiet_panel_bg"],
            text_bg=COLORS["quiet_text_bg"],
            fg=COLORS["quiet_text"],
            border=COLORS["quiet_border"],
            editable=False,
        )
        self.today_text = self._build_memo_panel(
            area,
            column=1,
            title=UI_TEXT["today_title"],
            description=UI_TEXT["today_description"],
            date_var=self.today_date_var,
            panel_bg=COLORS["panel_bg"],
            text_bg=COLORS["text_bg"],
            fg=COLORS["text"],
            border=COLORS["accent"],
            editable=True,
        )
        self.today_text.bind("<<Modified>>", self._on_text_modified)

    def _build_memo_panel(
        self,
        parent: tk.Widget,
        column: int,
        title: str,
        description: str,
        date_var: tk.StringVar,
        panel_bg: str,
        text_bg: str,
        fg: str,
        border: str,
        editable: bool,
    ) -> tk.Text:
        padx = (0, 12) if column == 0 else (12, 0)
        panel = tk.Frame(parent, bg=panel_bg, highlightbackground=border, highlightthickness=1)
        panel.grid(row=0, column=column, sticky="nsew", padx=padx)
        panel.grid_rowconfigure(1, weight=1)
        panel.grid_columnconfigure(0, weight=1)

        header = tk.Frame(panel, bg=panel_bg)
        header.grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 9))
        header.grid_columnconfigure(0, weight=1)

        title_area = tk.Frame(header, bg=panel_bg)
        title_area.grid(row=0, column=0, sticky="ew")
        tk.Label(
            title_area,
            text=title,
            bg=panel_bg,
            fg=COLORS["text"],
            font=self.fonts["section"],
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            title_area,
            text=description,
            bg=panel_bg,
            fg=COLORS["muted"],
            font=self.fonts["small"],
            anchor="w",
            justify="left",
        ).pack(fill="x", pady=(3, 0))
        tk.Label(
            title_area,
            textvariable=date_var,
            bg=panel_bg,
            fg=COLORS["muted"],
            font=self.fonts["small"],
            anchor="w",
        ).pack(fill="x", pady=(5, 0))

        if editable:
            buttons = tk.Frame(header, bg=panel_bg)
            buttons.grid(row=0, column=1, sticky="ne", padx=(16, 0))
            self._button(
                buttons,
                UI_TEXT["button_save"],
                self._manual_save,
                primary=True,
            ).pack(side="left")
            self._button(
                buttons,
                UI_TEXT["button_clear"],
                self._clear_today,
                primary=False,
            ).pack(side="left", padx=(8, 0))

        body = tk.Frame(panel, bg=panel_bg)
        body.grid(row=1, column=0, sticky="nsew", padx=16, pady=(0, 16))
        body.grid_rowconfigure(0, weight=1)
        body.grid_columnconfigure(0, weight=1)

        text = tk.Text(
            body,
            wrap="word",
            undo=editable,
            maxundo=80,
            bg=text_bg,
            fg=fg,
            insertbackground=COLORS["text"],
            selectbackground="#DCEBFF",
            selectforeground=COLORS["text"],
            relief="flat",
            borderwidth=0,
            highlightbackground=COLORS["border"],
            highlightcolor=COLORS["border"],
            highlightthickness=1,
            padx=13,
            pady=12,
            font=self.fonts["memo"],
            spacing1=3,
            spacing3=3,
        )
        text.grid(row=0, column=0, sticky="nsew")
        scrollbar = tk.Scrollbar(body, orient="vertical", command=text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        text.configure(yscrollcommand=scrollbar.set)
        if not editable:
            text.configure(cursor="arrow")
        return text

    def _button(self, parent: tk.Widget, text: str, command, primary: bool) -> tk.Button:
        if primary:
            normal_bg = COLORS["accent"]
            active_bg = COLORS["accent_hover"]
            fg = "#FFFFFF"
            font = self.fonts["button"]
        else:
            normal_bg = COLORS["panel_bg"]
            active_bg = COLORS["button_sub_hover"]
            fg = COLORS["text"]
            font = self.fonts["button_sub"]

        button = tk.Button(
            parent,
            text=text,
            command=command,
            font=font,
            bg=normal_bg,
            fg=fg,
            activebackground=active_bg,
            activeforeground=fg,
            relief="flat",
            borderwidth=0,
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            padx=13,
            pady=7,
            cursor="hand2",
        )
        button.bind("<Enter>", lambda _event: button.configure(bg=active_bg))
        button.bind("<Leave>", lambda _event: button.configure(bg=normal_bg))
        return button

    def _build_status(self, parent: tk.Widget) -> None:
        tk.Label(
            parent,
            textvariable=self.status_var,
            bg=COLORS["base_bg"],
            fg=COLORS["muted"],
            font=self.fonts["small"],
            anchor="w",
        ).pack(fill="x", pady=(8, 0))

    def _build_footer(self, parent: tk.Widget) -> None:
        self.footer = tk.Frame(parent, bg=COLORS["base_bg"])
        self.footer.pack(fill="x", pady=(8, 0))
        self._render_footer("wide")

    def _handle_resize(self, event) -> None:
        if event.widget is not self.root:
            return
        mode = "wide" if event.width >= 980 else "narrow"
        if mode != self.footer_mode:
            self._render_footer(mode)

    def _render_footer(self, mode: str) -> None:
        self.footer_mode = mode
        for child in self.footer.winfo_children():
            child.destroy()

        if mode == "wide":
            left = tk.Frame(self.footer, bg=COLORS["base_bg"])
            left.pack(side="left")
            self._footer_thought_line(left)

            right = tk.Frame(self.footer, bg=COLORS["base_bg"])
            right.pack(side="right")
            self._footer_link_line(right)
            return

        thought = tk.Frame(self.footer, bg=COLORS["base_bg"])
        thought.pack(anchor="center")
        self._footer_thought_line(thought)

        links = tk.Frame(self.footer, bg=COLORS["base_bg"])
        links.pack(anchor="center", pady=(4, 0))
        self._footer_link_line(links)

    def _footer_thought_line(self, parent: tk.Widget) -> None:
        self._footer_label(parent, UI_TEXT["footer_left"])
        self._footer_label(parent, UI_TEXT["footer_separator"])
        self._footer_label(parent, UI_TEXT["footer_catchcopy"])

    def _footer_link_line(self, parent: tk.Widget) -> None:
        self._footer_link(parent, UI_TEXT["footer_link_1"], LINK_URLS["footer_link_1"])
        self._footer_label(parent, UI_TEXT["footer_separator"])
        self._footer_link(parent, UI_TEXT["footer_link_2"], LINK_URLS["footer_link_2"])
        self._footer_label(parent, UI_TEXT["footer_separator"])
        self._footer_label(parent, UI_TEXT["footer_copyright"])

    def _footer_label(self, parent: tk.Widget, text: str) -> None:
        tk.Label(
            parent,
            text=text,
            font=self.fonts["small"],
            fg=COLORS["muted"],
            bg=COLORS["base_bg"],
        ).pack(side="left")

    def _footer_link(self, parent: tk.Widget, text: str, url: str) -> None:
        label = tk.Label(
            parent,
            text=text,
            font=self.fonts["small"],
            fg=COLORS["muted"],
            bg=COLORS["base_bg"],
            cursor="hand2",
        )
        label.pack(side="left")
        label.bind("<Button-1>", lambda _event: webbrowser.open(url))
        label.bind("<Enter>", lambda _event: label.configure(fg=COLORS["accent"]))
        label.bind("<Leave>", lambda _event: label.configure(fg=COLORS["muted"]))

    def _load_initial_memos(self) -> None:
        previous_memo = load_memo(self.previous_day)
        self.previous_text.configure(state="normal")
        self.previous_text.delete("1.0", "end")
        if previous_memo.strip():
            self.previous_text.insert("1.0", previous_memo)
        else:
            self.previous_text.insert("1.0", UI_TEXT["previous_empty"])
            self.status_var.set(UI_TEXT["status_no_previous"])
        self.previous_text.configure(state="disabled")

        today_memo = load_memo(self.today)
        self.suppress_modified = True
        try:
            self.today_text.delete("1.0", "end")
            if today_memo:
                self.today_text.insert("1.0", today_memo)
            self.today_text.edit_modified(False)
        finally:
            self.suppress_modified = False

    def _on_text_modified(self, event) -> None:
        if self.suppress_modified:
            event.widget.edit_modified(False)
            return
        if event.widget.edit_modified():
            event.widget.edit_modified(False)
            self._schedule_save()

    def _schedule_save(self) -> None:
        if self.save_after_id is not None:
            self.root.after_cancel(self.save_after_id)
        self.status_var.set(UI_TEXT["status_ready"])
        self.save_after_id = self.root.after(SAVE_DEBOUNCE_MS, self._autosave)

    def _memo_text(self) -> str:
        return self.today_text.get("1.0", "end-1c")

    def _save_current(self, status_key: str) -> bool:
        if self.save_after_id is not None:
            self.root.after_cancel(self.save_after_id)
            self.save_after_id = None
        try:
            save_memo(self.today, self._memo_text())
        except Exception:
            self.status_var.set(UI_TEXT["status_save_failed"])
            return False
        self.status_var.set(UI_TEXT[status_key])
        self._schedule_status_reset()
        return True

    def _autosave(self) -> None:
        self.save_after_id = None
        self._save_current("status_autosaved")

    def _manual_save(self) -> None:
        self._save_current("status_saved")
        self.today_text.focus_set()

    def _clear_today(self) -> None:
        if not self._memo_text():
            self.today_text.focus_set()
            return
        confirmed = messagebox.askyesno(
            UI_TEXT["dialog_clear_title"],
            UI_TEXT["dialog_clear_body"],
            parent=self.root,
        )
        if not confirmed:
            self.today_text.focus_set()
            return
        self.suppress_modified = True
        try:
            self.today_text.delete("1.0", "end")
            self.today_text.edit_modified(False)
        finally:
            self.suppress_modified = False
        self._save_current("status_cleared")
        self.today_text.focus_set()

    def _schedule_status_reset(self) -> None:
        if self.status_after_id is not None:
            self.root.after_cancel(self.status_after_id)
        self.status_after_id = self.root.after(
            STATUS_RESET_MS,
            lambda: self.status_var.set(UI_TEXT["status_ready"]),
        )

    def _close(self) -> None:
        self._save_current("status_saved")
        self.root.destroy()


def main() -> None:
    set_windows_app_id()
    root = tk.Tk()
    YesterdayTaskMemoApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
