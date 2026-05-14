from __future__ import annotations

from tkinter import font as tkfont


def choose_font_family(root, candidates: tuple[str, ...]) -> str:
    try:
        available = set(tkfont.families(root))
    except Exception:
        return "TkDefaultFont"
    for family in candidates:
        if family in available:
            return family
    return "TkDefaultFont"


def set_textbox_text(textbox, value: str) -> None:
    textbox.configure(state="normal")
    textbox.delete("1.0", "end")
    textbox.insert("1.0", value)
    textbox.configure(state="disabled")
