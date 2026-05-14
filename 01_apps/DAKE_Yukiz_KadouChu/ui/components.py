from __future__ import annotations

import customtkinter as ctk

from ui.theme import COLORS, FONT_FAMILY


def make_panel(parent: ctk.CTkFrame, title: str) -> tuple[ctk.CTkFrame, ctk.CTkFrame]:
    panel = ctk.CTkFrame(parent, fg_color=COLORS["panel"], border_width=1, border_color=COLORS["line"], corner_radius=8)
    panel.grid_columnconfigure(0, weight=1)
    panel.grid_rowconfigure(1, weight=1)
    ctk.CTkLabel(
        panel,
        text=title,
        font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
        text_color=COLORS["accent_soft"],
    ).grid(row=0, column=0, sticky="w", padx=14, pady=(12, 4))
    body = ctk.CTkFrame(panel, fg_color="transparent")
    body.grid(row=1, column=0, sticky="nsew", padx=14, pady=(4, 14))
    return panel, body


class StatusPill(ctk.CTkFrame):
    def __init__(self, parent: ctk.CTkFrame, label: str) -> None:
        super().__init__(parent, fg_color=COLORS["panel_alt"], corner_radius=6)
        self.grid_columnconfigure(0, weight=1)
        self.label = label
        self.name_label = ctk.CTkLabel(
            self,
            text=label,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"),
            text_color=COLORS["text"],
        )
        self.name_label.grid(row=0, column=0, sticky="w", padx=10, pady=(7, 0))
        self.state_label = ctk.CTkLabel(
            self,
            text="CHECKING",
            font=ctk.CTkFont(family=FONT_FAMILY, size=10),
            text_color=COLORS["muted"],
        )
        self.state_label.grid(row=1, column=0, sticky="w", padx=10, pady=(0, 7))

    def set_state(self, state: str, detail: str = "") -> None:
        color = COLORS["muted"]
        state_upper = state.upper()
        if any(token in state_upper for token in ["ONLINE", "READY"]):
            color = COLORS["success"]
        elif any(token in state_upper for token in ["MISSING", "UNAVAILABLE", "ERROR"]):
            color = COLORS["danger"]
        elif any(token in state_upper for token in ["SKIPPED", "FOUND", "CHECK"]):
            color = COLORS["warning"]
        suffix = f" / {detail}" if detail and state_upper not in detail.upper() else ""
        self.state_label.configure(text=f"{state}{suffix}"[:64], text_color=color)


def set_textbox(textbox: ctk.CTkTextbox, text: str) -> None:
    textbox.configure(state="normal")
    textbox.delete("1.0", "end")
    textbox.insert("end", text)
    textbox.configure(state="disabled")


def append_textbox(textbox: ctk.CTkTextbox, text: str) -> None:
    textbox.configure(state="normal")
    textbox.insert("end", text)
    textbox.see("end")
    textbox.configure(state="disabled")
