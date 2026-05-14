from __future__ import annotations

FONT_FAMILY = "Yu Gothic UI"

COLORS = {
    "bg": "#070A0E",
    "panel": "#0D131A",
    "panel_alt": "#101922",
    "field": "#080D12",
    "line": "#1B2B35",
    "text": "#E7EDF2",
    "muted": "#80909C",
    "accent": "#69D2FF",
    "accent_soft": "#8AB7CC",
    "button": "#173142",
    "button_secondary": "#121E28",
    "button_hover": "#21485F",
    "success": "#7CE0A3",
    "warning": "#D8B86C",
    "danger": "#E06C75",
}


def setup_theme(ctk_module: object) -> None:
    ctk_module.set_appearance_mode("dark")
    ctk_module.set_default_color_theme("dark-blue")
