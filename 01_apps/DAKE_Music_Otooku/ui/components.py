# -*- coding: utf-8 -*-
from __future__ import annotations

import tkinter as tk


def make_panel(parent: tk.Widget, colors: dict[str, str]) -> tk.Frame:
    return tk.Frame(
        parent,
        bg=colors["surface"],
        highlightbackground=colors["border"],
        highlightthickness=1,
        bd=0,
    )


def make_button(
    parent: tk.Widget,
    label: str,
    command,
    colors: dict[str, str],
    font,
    primary: bool = False,
) -> tk.Button:
    bg = colors["accent"] if primary else colors["surface_soft"]
    fg = colors["text"]
    hover = colors["accent_hover"] if primary else colors["secondary_hover"]
    border = colors["accent"] if primary else colors["border"]
    button = tk.Button(
        parent,
        text=label,
        command=command,
        bg=bg,
        fg=fg,
        activebackground=hover,
        activeforeground=fg,
        disabledforeground=colors["quiet"],
        font=font,
        relief="flat",
        bd=0,
        padx=16,
        pady=8,
        cursor="hand2",
        highlightthickness=1,
        highlightbackground=border,
    )
    button.bind("<Enter>", lambda _event, widget=button, color=hover: _hover(widget, color))
    button.bind("<Leave>", lambda _event, widget=button, color=bg: _hover(widget, color))
    return button


def _hover(button: tk.Button, color: str) -> None:
    if str(button.cget("state")) != "disabled":
        button.configure(bg=color)
