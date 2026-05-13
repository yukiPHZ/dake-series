# -*- coding: utf-8 -*-
from __future__ import annotations

import ctypes
import sys
import uuid
import webbrowser
from dataclasses import dataclass
from pathlib import Path
import tkinter as tk
from tkinter import font as tkfont


APP_NAME = "付箋メモ"
WINDOW_TITLE = APP_NAME
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "app_title": APP_NAME,
    "main_title": "付箋を並べる",
    "main_description": "思いついたことを付箋に書いて、動かして、捨てられます。",
    "button_add_note": "付箋を追加",
    "button_clear_all": "全部消す",
    "button_delete_note": "×",
    "empty_state": "付箋を追加してください",
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
    "text": "#1E2430",
    "muted": "#667085",
    "quiet": "#A7B0BF",
    "border": "#E6EAF0",
    "button": "#2F6FED",
    "button_hover": "#2458BF",
    "secondary_button": "#FFFFFF",
    "secondary_hover": "#F2F4F7",
    "note": "#FFF7C2",
    "note_header": "#FFF1A8",
    "note_border": "#E7D98E",
    "note_text": "#2D2A1F",
    "danger": "#B42318",
    "danger_hover": "#FDECEC",
}

FONT_CANDIDATES = ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo")
WINDOW_SIZE = "940x640"
WINDOW_MIN_SIZE = (720, 520)
APP_USER_MODEL_ID = "Shimarisu.DakeStickyMemo"
FOOTER_NARROW_WIDTH = 860

NOTE_WIDTH = 220
NOTE_HEIGHT = 170
NOTE_START_X = 28
NOTE_START_Y = 24
NOTE_CASCADE = 26


@dataclass
class StickyNote:
    id: str
    text: str
    x: int
    y: int


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


class StickyMemoApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.font_family = choose_font(root)
        self.notes: dict[str, StickyNote] = {}
        self.note_items: dict[str, int] = {}
        self.note_frames: dict[str, tk.Frame] = {}
        self.drag_note_id: str | None = None
        self.drag_offset_x = 0
        self.drag_offset_y = 0
        self.empty_text_item: int | None = None
        self.footer: tk.Frame | None = None
        self.footer_mode: str | None = None

        self.fonts = {
            "title": tkfont.Font(root, family=self.font_family, size=20, weight="bold"),
            "description": tkfont.Font(root, family=self.font_family, size=10),
            "button": tkfont.Font(root, family=self.font_family, size=10, weight="bold"),
            "note": tkfont.Font(root, family=self.font_family, size=10),
            "note_button": tkfont.Font(root, family=self.font_family, size=9, weight="bold"),
            "empty": tkfont.Font(root, family=self.font_family, size=13),
            "footer": tkfont.Font(root, family=self.font_family, size=8),
        }

        self._configure_root()
        self._build_ui()
        self._apply_icon()

    def _configure_root(self) -> None:
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(*WINDOW_MIN_SIZE)
        self.root.configure(bg=COLORS["background"])

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
        outer.pack(fill="both", expand=True, padx=28, pady=(24, 18))

        self._build_header(outer)
        self._build_canvas(outer)
        self._build_footer(outer)
        self.root.bind("<Configure>", self._on_root_configure, add="+")

    def _build_header(self, parent: tk.Widget) -> None:
        header = tk.Frame(parent, bg=COLORS["background"])
        header.pack(fill="x", pady=(0, 16))
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
        ).pack(anchor="w", pady=(6, 0))

        actions = tk.Frame(header, bg=COLORS["background"])
        actions.grid(row=0, column=1, sticky="e", padx=(18, 0))

        self._make_button(
            actions,
            UI_TEXT["button_add_note"],
            self.add_note,
            primary=True,
        ).pack(side="left", padx=(0, 10))

        self._make_button(
            actions,
            UI_TEXT["button_clear_all"],
            self.clear_all_notes,
            primary=False,
        ).pack(side="left")

    def _build_canvas(self, parent: tk.Widget) -> None:
        shell = tk.Frame(
            parent,
            bg=COLORS["surface"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
        )
        shell.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(
            shell,
            bg=COLORS["surface"],
            bd=0,
            highlightthickness=0,
            relief="flat",
        )
        self.canvas.pack(fill="both", expand=True)
        self.empty_text_item = self.canvas.create_text(
            0,
            0,
            text=UI_TEXT["empty_state"],
            fill=COLORS["quiet"],
            font=self.fonts["empty"],
            anchor="center",
        )
        self.canvas.bind("<Configure>", self._on_canvas_configure)

    def _build_footer(self, parent: tk.Widget) -> None:
        self.footer = tk.Frame(parent, bg=COLORS["background"])
        self.footer.pack(fill="x", pady=(14, 0))
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
            font=self.fonts["button"],
            relief="flat",
            bd=0,
            padx=18,
            pady=8,
            cursor="hand2",
            highlightthickness=1,
            highlightbackground=border,
        )
        button.bind("<Enter>", lambda _event, target=button, color=hover: target.configure(bg=color))
        button.bind("<Leave>", lambda _event, target=button, color=bg: target.configure(bg=color))
        return button

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

    def add_note(self) -> None:
        self.root.update_idletasks()
        note_id = uuid.uuid4().hex
        index = len(self.notes)
        x = NOTE_START_X + (index % 5) * NOTE_CASCADE
        y = NOTE_START_Y + (index % 5) * NOTE_CASCADE
        x, y = self._clamp_position(x, y)
        note = StickyNote(id=note_id, text="", x=x, y=y)
        self.notes[note_id] = note
        self._create_note_widget(note)
        self._sync_empty_state()

    def _create_note_widget(self, note: StickyNote) -> None:
        frame = tk.Frame(
            self.canvas,
            width=NOTE_WIDTH,
            height=NOTE_HEIGHT,
            bg=COLORS["note"],
            highlightbackground=COLORS["note_border"],
            highlightthickness=1,
        )
        frame.pack_propagate(False)

        header = tk.Frame(frame, height=30, bg=COLORS["note_header"], cursor="fleur")
        header.pack(fill="x")
        header.pack_propagate(False)

        spacer = tk.Frame(header, bg=COLORS["note_header"], cursor="fleur")
        spacer.pack(side="left", fill="both", expand=True)

        close_button = tk.Button(
            header,
            text=UI_TEXT["button_delete_note"],
            command=lambda note_id=note.id: self.delete_note(note_id),
            bg=COLORS["note_header"],
            fg=COLORS["muted"],
            activebackground=COLORS["danger_hover"],
            activeforeground=COLORS["danger"],
            font=self.fonts["note_button"],
            relief="flat",
            bd=0,
            width=3,
            cursor="hand2",
            padx=2,
            pady=1,
        )
        close_button.pack(side="right", padx=(0, 8), pady=3)

        text = tk.Text(
            frame,
            bg=COLORS["note"],
            fg=COLORS["note_text"],
            insertbackground=COLORS["note_text"],
            selectbackground="#D7E7FF",
            selectforeground=COLORS["note_text"],
            inactiveselectbackground="#EAF2FF",
            font=self.fonts["note"],
            relief="flat",
            bd=0,
            padx=13,
            pady=10,
            wrap="word",
            undo=True,
            height=1,
        )
        text.pack(fill="both", expand=True)
        text.insert("1.0", note.text)
        text.edit_modified(False)
        text.bind("<<Modified>>", lambda _event, note_id=note.id, widget=text: self._on_note_text_changed(note_id, widget))

        for drag_target in (frame, header, spacer):
            drag_target.bind("<ButtonPress-1>", lambda event, note_id=note.id: self._begin_drag(note_id, event))
            drag_target.bind("<B1-Motion>", self._drag_note)
            drag_target.bind("<ButtonRelease-1>", self._end_drag)

        item_id = self.canvas.create_window(note.x, note.y, window=frame, width=NOTE_WIDTH, height=NOTE_HEIGHT, anchor="nw")
        self.note_items[note.id] = item_id
        self.note_frames[note.id] = frame
        self.canvas.tag_raise(item_id)
        text.focus_set()

    def _on_note_text_changed(self, note_id: str, widget: tk.Text) -> None:
        if not widget.edit_modified():
            return
        note = self.notes.get(note_id)
        if note is not None:
            note.text = widget.get("1.0", "end-1c")
        widget.edit_modified(False)

    def delete_note(self, note_id: str) -> None:
        item_id = self.note_items.pop(note_id, None)
        frame = self.note_frames.pop(note_id, None)
        self.notes.pop(note_id, None)
        if item_id is not None:
            self.canvas.delete(item_id)
        if frame is not None:
            frame.destroy()
        if self.drag_note_id == note_id:
            self.drag_note_id = None
        self._sync_empty_state()

    def clear_all_notes(self) -> None:
        for note_id in list(self.notes):
            self.delete_note(note_id)
        self._sync_empty_state()

    def _begin_drag(self, note_id: str, event: tk.Event) -> None:
        note = self.notes.get(note_id)
        item_id = self.note_items.get(note_id)
        if note is None or item_id is None:
            return
        self.drag_note_id = note_id
        self.drag_offset_x = event.x_root - self.canvas.winfo_rootx() - note.x
        self.drag_offset_y = event.y_root - self.canvas.winfo_rooty() - note.y
        self.canvas.tag_raise(item_id)

    def _drag_note(self, event: tk.Event) -> None:
        if self.drag_note_id is None:
            return
        note = self.notes.get(self.drag_note_id)
        item_id = self.note_items.get(self.drag_note_id)
        if note is None or item_id is None:
            return
        next_x = event.x_root - self.canvas.winfo_rootx() - self.drag_offset_x
        next_y = event.y_root - self.canvas.winfo_rooty() - self.drag_offset_y
        note.x, note.y = self._clamp_position(next_x, next_y)
        self.canvas.coords(item_id, note.x, note.y)

    def _end_drag(self, _event: tk.Event) -> None:
        self.drag_note_id = None

    def _on_canvas_configure(self, event: tk.Event) -> None:
        if self.empty_text_item is not None:
            self.canvas.coords(self.empty_text_item, event.width // 2, event.height // 2)
        self._clamp_all_notes()

    def _clamp_all_notes(self) -> None:
        for note_id, note in self.notes.items():
            x, y = self._clamp_position(note.x, note.y)
            if (x, y) == (note.x, note.y):
                continue
            note.x, note.y = x, y
            item_id = self.note_items.get(note_id)
            if item_id is not None:
                self.canvas.coords(item_id, x, y)

    def _clamp_position(self, x: int | float, y: int | float) -> tuple[int, int]:
        canvas_w = max(self.canvas.winfo_width(), NOTE_WIDTH)
        canvas_h = max(self.canvas.winfo_height(), NOTE_HEIGHT)
        max_x = max(canvas_w - NOTE_WIDTH, 0)
        max_y = max(canvas_h - NOTE_HEIGHT, 0)
        clamped_x = min(max(int(x), 0), max_x)
        clamped_y = min(max(int(y), 0), max_y)
        return clamped_x, clamped_y

    def _sync_empty_state(self) -> None:
        if self.empty_text_item is None:
            return
        state = "hidden" if self.notes else "normal"
        self.canvas.itemconfigure(self.empty_text_item, state=state)


def main() -> None:
    set_windows_app_id()
    root = tk.Tk()
    StickyMemoApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
