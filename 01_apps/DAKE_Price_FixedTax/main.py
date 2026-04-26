from __future__ import annotations

import os
import tkinter as tk
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from tkinter import font as tkfont
from tkinter import ttk
from typing import Any, Dict, Optional, Tuple

APP_NAME = "Dake固都税計算"
WINDOW_TITLE = "固都税計算"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "main_title": "固定資産税・都市計画税を概算する",
    "main_description": "評価額を入力すると、概算年税額をすぐ表示します。",
    "section_input": "入力",
    "section_result": "計算結果",
    "label_land_value": "土地評価額",
    "label_building_value": "建物評価額",
    "label_residential_type": "住宅用地区分",
    "radio_none": "特例なし",
    "radio_small": "小規模住宅用地",
    "radio_general": "一般住宅用地",
    "label_city_tax_enabled": "都市計画税を計算する",
    "label_fixed_rate": "固定資産税率",
    "label_city_rate": "都市計画税率",
    "label_status": "状態",
    "toggle_on": "ON",
    "toggle_off": "OFF",
    "result_fixed_tax": "固定資産税",
    "result_city_tax": "都市計画税",
    "result_total_tax": "合計年税額",
    "result_monthly_tax": "月額目安",
    "status_idle": "入力してください",
    "status_ready": "概算を表示しています",
    "status_error": "入力内容を確認してください",
    "error_required_land": "土地評価額を入力してください",
    "error_required_building": "建物評価額を入力してください",
    "error_required_fixed_rate": "固定資産税率を入力してください",
    "error_required_city_rate": "都市計画税率を入力してください",
    "error_invalid_number": "{field}は数字で入力してください",
    "error_negative": "{field}は0以上の数値を入力してください",
    "disclaimer": "入力値に基づく参考計算です。税務上の適正性を保証するものではありません。最終判断は税理士等の専門家に確認してください。",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_copyright": COPYRIGHT,
    "unit_yen": "円",
    "unit_percent": "%",
    "value_placeholder": "—",
}

BG_COLOR = "#F6F7F9"
CARD_COLOR = "#FFFFFF"
TEXT_COLOR = "#1E2430"
MUTED_COLOR = "#667085"
BORDER_COLOR = "#E6EAF0"
ACCENT_COLOR = "#2F6FED"
ERROR_COLOR = "#D92D20"

DECIMAL_ONE = Decimal("1")
DECIMAL_TWELVE = Decimal("12")
DECIMAL_HUNDRED = Decimal("100")
ICON_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "..",
    "..",
    "02_assets",
    "dake_icon.ico",
)


@dataclass
class TaxBreakdown:
    fixed_tax: int
    city_tax: int
    total_tax: int
    monthly_tax: int


class InputValidationError(Exception):
    def __init__(self, message: str) -> None:
        super().__init__(message)
        self.message = message


def choose_font_family(root: tk.Misc) -> str:
    preferred = ["BIZ UDPGothic", "Yu Gothic UI", "Meiryo"]
    available = set(tkfont.families(root))
    for family in preferred:
        if family in available:
            return family
    return "TkDefaultFont"


def normalize_numeric_text(value: str) -> str:
    return value.replace(",", "").replace(" ", "").replace("\u3000", "").strip()


def parse_non_negative_decimal(
    raw_value: str,
    *,
    field_label: str,
    required_message: str,
) -> Decimal:
    cleaned = normalize_numeric_text(raw_value)
    if not cleaned:
        raise InputValidationError(required_message)

    try:
        number = Decimal(cleaned)
    except InvalidOperation as exc:
        raise InputValidationError(
            UI_TEXT["error_invalid_number"].format(field=field_label)
        ) from exc

    if not number.is_finite():
        raise InputValidationError(UI_TEXT["error_invalid_number"].format(field=field_label))

    if number < 0:
        raise InputValidationError(UI_TEXT["error_negative"].format(field=field_label))

    return number


def calculate_land_bases(land_value: Decimal, residential_type: str) -> Tuple[Decimal, Decimal]:
    if residential_type == "small":
        return land_value / Decimal("6"), land_value / Decimal("3")
    if residential_type == "general":
        return land_value / Decimal("3"), land_value * Decimal("2") / Decimal("3")
    return land_value, land_value


def round_yen(value: Decimal) -> int:
    return int(value.quantize(DECIMAL_ONE, rounding=ROUND_HALF_UP))


def calculate_tax_breakdown(
    *,
    land_value: Decimal,
    building_value: Decimal,
    residential_type: str,
    city_tax_enabled: bool,
    fixed_rate_percent: Decimal,
    city_rate_percent: Decimal,
) -> TaxBreakdown:
    land_fixed_base, land_city_base = calculate_land_bases(land_value, residential_type)
    fixed_rate = fixed_rate_percent / DECIMAL_HUNDRED
    city_rate = city_rate_percent / DECIMAL_HUNDRED

    fixed_tax_raw = (land_fixed_base + building_value) * fixed_rate
    city_tax_raw = (
        (land_city_base + building_value) * city_rate if city_tax_enabled else Decimal("0")
    )
    total_tax_raw = fixed_tax_raw + city_tax_raw
    monthly_tax_raw = total_tax_raw / DECIMAL_TWELVE

    return TaxBreakdown(
        fixed_tax=round_yen(fixed_tax_raw),
        city_tax=round_yen(city_tax_raw),
        total_tax=round_yen(total_tax_raw),
        monthly_tax=round_yen(monthly_tax_raw),
    )


def format_yen(value: int) -> str:
    return f"{value:,}{UI_TEXT['unit_yen']}"


class FixedTaxApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.pending_refresh: Optional[str] = None
        self.font_family = choose_font_family(self)

        self.title(WINDOW_TITLE)
        self.geometry("470x680")
        self.resizable(False, False)
        self.configure(bg=BG_COLOR)

        try:
            self.iconbitmap(ICON_PATH)
        except Exception:
            pass

        self.land_value_var = tk.StringVar()
        self.building_value_var = tk.StringVar()
        self.residential_type_var = tk.StringVar(value="none")
        self.city_tax_enabled_var = tk.BooleanVar(value=True)
        self.fixed_rate_var = tk.StringVar(value="1.4")
        self.city_rate_var = tk.StringVar(value="0.3")
        self.status_var = tk.StringVar(value=UI_TEXT["status_idle"])
        self.city_toggle_var = tk.StringVar()

        self.result_vars = {
            "fixed_tax": tk.StringVar(value=UI_TEXT["value_placeholder"]),
            "city_tax": tk.StringVar(value=UI_TEXT["value_placeholder"]),
            "total_tax": tk.StringVar(value=UI_TEXT["value_placeholder"]),
            "monthly_tax": tk.StringVar(value=UI_TEXT["value_placeholder"]),
        }

        self.style = ttk.Style(self)
        self.configure_styles()
        self.build_ui()
        self.bind_live_updates()
        self.update_city_toggle_text()

    def configure_styles(self) -> None:
        self.style.theme_use("clam")
        self.style.configure("App.TFrame", background=BG_COLOR)
        self.style.configure(
            "Card.TFrame",
            background=CARD_COLOR,
            bordercolor=BORDER_COLOR,
            borderwidth=1,
            relief="solid",
        )
        self.style.configure(
            "Title.TLabel",
            background=CARD_COLOR,
            foreground=TEXT_COLOR,
            font=(self.font_family, 17, "bold"),
        )
        self.style.configure(
            "Section.TLabel",
            background=CARD_COLOR,
            foreground=TEXT_COLOR,
            font=(self.font_family, 12, "bold"),
        )
        self.style.configure(
            "Body.TLabel",
            background=CARD_COLOR,
            foreground=MUTED_COLOR,
            font=(self.font_family, 10),
        )
        self.style.configure(
            "Field.TLabel",
            background=CARD_COLOR,
            foreground=TEXT_COLOR,
            font=(self.font_family, 10),
        )
        self.style.configure(
            "Unit.TLabel",
            background=CARD_COLOR,
            foreground=MUTED_COLOR,
            font=(self.font_family, 10),
        )
        self.style.configure(
            "ResultName.TLabel",
            background=CARD_COLOR,
            foreground=TEXT_COLOR,
            font=(self.font_family, 11),
        )
        self.style.configure(
            "ResultValue.TLabel",
            background=CARD_COLOR,
            foreground=TEXT_COLOR,
            font=(self.font_family, 15, "bold"),
        )
        self.style.configure(
            "TotalName.TLabel",
            background=CARD_COLOR,
            foreground=TEXT_COLOR,
            font=(self.font_family, 11, "bold"),
        )
        self.style.configure(
            "TotalValue.TLabel",
            background=CARD_COLOR,
            foreground=ACCENT_COLOR,
            font=(self.font_family, 16, "bold"),
        )
        self.style.configure(
            "App.TEntry",
            fieldbackground=CARD_COLOR,
            foreground=TEXT_COLOR,
            bordercolor=BORDER_COLOR,
            lightcolor=BORDER_COLOR,
            darkcolor=BORDER_COLOR,
            padding=(10, 7),
            relief="solid",
        )
        self.style.map(
            "App.TEntry",
            bordercolor=[("focus", ACCENT_COLOR)],
            lightcolor=[("focus", ACCENT_COLOR)],
            darkcolor=[("focus", ACCENT_COLOR)],
        )
        self.style.configure(
            "App.TRadiobutton",
            background=CARD_COLOR,
            foreground=TEXT_COLOR,
            font=(self.font_family, 10),
        )
        self.style.configure(
            "App.TCheckbutton",
            background=CARD_COLOR,
            foreground=TEXT_COLOR,
            font=(self.font_family, 10, "bold"),
        )
        self.style.configure(
            "StatusIdle.TLabel",
            background=CARD_COLOR,
            foreground=MUTED_COLOR,
            font=(self.font_family, 10),
        )
        self.style.configure(
            "StatusReady.TLabel",
            background=CARD_COLOR,
            foreground=ACCENT_COLOR,
            font=(self.font_family, 10),
        )
        self.style.configure(
            "StatusError.TLabel",
            background=CARD_COLOR,
            foreground=ERROR_COLOR,
            font=(self.font_family, 10),
        )
        self.style.configure(
            "Disclaimer.TLabel",
            background=CARD_COLOR,
            foreground=MUTED_COLOR,
            font=(self.font_family, 9),
        )
        self.style.configure(
            "Footer.TLabel",
            background=BG_COLOR,
            foreground=MUTED_COLOR,
            font=(self.font_family, 9),
        )
        self.style.configure("Card.TSeparator", background=BORDER_COLOR)

    def build_ui(self) -> None:
        container = ttk.Frame(self, style="App.TFrame", padding=18)
        container.pack(fill="both", expand=True)

        self.build_header_card(container)
        self.build_input_card(container)
        self.build_result_card(container)
        self.build_footer(container)

    def build_header_card(self, parent: ttk.Frame) -> None:
        card = ttk.Frame(parent, style="Card.TFrame", padding=18)
        card.pack(fill="x", pady=(0, 14))

        ttk.Label(card, text=UI_TEXT["main_title"], style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            card,
            text=UI_TEXT["main_description"],
            style="Body.TLabel",
            wraplength=390,
            justify="left",
        ).pack(anchor="w", pady=(8, 0))

    def build_input_card(self, parent: ttk.Frame) -> None:
        card = ttk.Frame(parent, style="Card.TFrame", padding=18)
        card.pack(fill="x", pady=(0, 14))
        card.columnconfigure(1, weight=1)

        ttk.Label(card, text=UI_TEXT["section_input"], style="Section.TLabel").grid(
            row=0, column=0, columnspan=3, sticky="w"
        )

        self.add_entry_row(
            card,
            row=1,
            label_text=UI_TEXT["label_land_value"],
            text_variable=self.land_value_var,
            unit_text=UI_TEXT["unit_yen"],
        )
        self.add_entry_row(
            card,
            row=2,
            label_text=UI_TEXT["label_building_value"],
            text_variable=self.building_value_var,
            unit_text=UI_TEXT["unit_yen"],
        )

        ttk.Label(card, text=UI_TEXT["label_residential_type"], style="Field.TLabel").grid(
            row=3, column=0, sticky="nw", pady=(16, 0)
        )
        radio_group = ttk.Frame(card, style="Card.TFrame")
        radio_group.grid(row=3, column=1, columnspan=2, sticky="w", pady=(16, 0))
        radio_items = [
            ("none", UI_TEXT["radio_none"]),
            ("small", UI_TEXT["radio_small"]),
            ("general", UI_TEXT["radio_general"]),
        ]
        for index, (value, label) in enumerate(radio_items):
            ttk.Radiobutton(
                radio_group,
                text=label,
                value=value,
                variable=self.residential_type_var,
                style="App.TRadiobutton",
            ).grid(row=index, column=0, sticky="w", pady=(0 if index == 0 else 6, 0))

        ttk.Label(card, text=UI_TEXT["label_city_tax_enabled"], style="Field.TLabel").grid(
            row=4, column=0, sticky="w", pady=(16, 0)
        )
        ttk.Checkbutton(
            card,
            textvariable=self.city_toggle_var,
            variable=self.city_tax_enabled_var,
            style="App.TCheckbutton",
        ).grid(row=4, column=1, columnspan=2, sticky="w", pady=(16, 0))

        self.add_entry_row(
            card,
            row=5,
            label_text=UI_TEXT["label_fixed_rate"],
            text_variable=self.fixed_rate_var,
            unit_text=UI_TEXT["unit_percent"],
        )
        self.add_entry_row(
            card,
            row=6,
            label_text=UI_TEXT["label_city_rate"],
            text_variable=self.city_rate_var,
            unit_text=UI_TEXT["unit_percent"],
        )

        ttk.Label(card, text=UI_TEXT["label_status"], style="Field.TLabel").grid(
            row=7, column=0, sticky="w", pady=(18, 0)
        )
        self.status_label = ttk.Label(
            card,
            textvariable=self.status_var,
            style="StatusIdle.TLabel",
            wraplength=300,
            justify="left",
        )
        self.status_label.grid(row=7, column=1, columnspan=2, sticky="w", pady=(18, 0))

    def build_result_card(self, parent: ttk.Frame) -> None:
        card = ttk.Frame(parent, style="Card.TFrame", padding=18)
        card.pack(fill="x")
        card.columnconfigure(1, weight=1)

        ttk.Label(card, text=UI_TEXT["section_result"], style="Section.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w"
        )

        rows = [
            ("result_fixed_tax", "fixed_tax", "ResultName.TLabel", "ResultValue.TLabel"),
            ("result_city_tax", "city_tax", "ResultName.TLabel", "ResultValue.TLabel"),
            ("result_total_tax", "total_tax", "TotalName.TLabel", "TotalValue.TLabel"),
            ("result_monthly_tax", "monthly_tax", "ResultName.TLabel", "ResultValue.TLabel"),
        ]

        for row_index, (label_key, value_key, name_style, value_style) in enumerate(rows, start=1):
            if row_index == 3:
                ttk.Separator(card, style="Card.TSeparator").grid(
                    row=row_index, column=0, columnspan=2, sticky="ew", pady=(6, 10)
                )
                content_row = row_index + 1
            else:
                content_row = row_index if row_index < 3 else row_index + 1

            ttk.Label(card, text=UI_TEXT[label_key], style=name_style).grid(
                row=content_row, column=0, sticky="w", pady=(0, 10)
            )
            ttk.Label(
                card,
                textvariable=self.result_vars[value_key],
                style=value_style,
                anchor="e",
                width=15,
            ).grid(row=content_row, column=1, sticky="e", pady=(0, 10))

        ttk.Separator(card, style="Card.TSeparator").grid(
            row=6, column=0, columnspan=2, sticky="ew", pady=(4, 12)
        )
        ttk.Label(
            card,
            text=UI_TEXT["disclaimer"],
            style="Disclaimer.TLabel",
            wraplength=390,
            justify="left",
        ).grid(row=7, column=0, columnspan=2, sticky="w")

    def build_footer(self, parent: ttk.Frame) -> None:
        footer = ttk.Frame(parent, style="App.TFrame", padding=(0, 14, 0, 0))
        footer.pack(fill="x")

        ttk.Label(footer, text=UI_TEXT["footer_left"], style="Footer.TLabel").pack(anchor="w")
        ttk.Label(
            footer,
            text=UI_TEXT["footer_copyright"],
            style="Footer.TLabel",
            wraplength=410,
            justify="left",
        ).pack(anchor="w", pady=(2, 0))

    def add_entry_row(
        self,
        parent: ttk.Frame,
        *,
        row: int,
        label_text: str,
        text_variable: tk.StringVar,
        unit_text: str,
    ) -> None:
        pady = (16, 0) if row > 1 else (16, 0)
        ttk.Label(parent, text=label_text, style="Field.TLabel").grid(
            row=row, column=0, sticky="w", pady=pady
        )
        ttk.Entry(
            parent,
            textvariable=text_variable,
            justify="right",
            width=24,
            style="App.TEntry",
        ).grid(row=row, column=1, sticky="ew", pady=pady, padx=(0, 8))
        ttk.Label(parent, text=unit_text, style="Unit.TLabel").grid(
            row=row, column=2, sticky="w", pady=pady
        )

    def bind_live_updates(self) -> None:
        variables = [
            self.land_value_var,
            self.building_value_var,
            self.residential_type_var,
            self.city_tax_enabled_var,
            self.fixed_rate_var,
            self.city_rate_var,
        ]
        for variable in variables:
            variable.trace_add("write", self.schedule_refresh)

    def schedule_refresh(self, *_args: object) -> None:
        self.update_city_toggle_text()
        if self.pending_refresh is not None:
            self.after_cancel(self.pending_refresh)
        self.pending_refresh = self.after(0, self.refresh_outputs)

    def update_city_toggle_text(self) -> None:
        self.city_toggle_var.set(
            UI_TEXT["toggle_on"] if self.city_tax_enabled_var.get() else UI_TEXT["toggle_off"]
        )

    def refresh_outputs(self) -> None:
        self.pending_refresh = None
        try:
            values = self.collect_inputs()
            result = calculate_tax_breakdown(**values)
        except InputValidationError as exc:
            self.clear_results()
            self.set_status(exc.message, "error")
            return
        except Exception:
            self.clear_results()
            self.set_status(UI_TEXT["status_error"], "error")
            return

        self.result_vars["fixed_tax"].set(format_yen(result.fixed_tax))
        self.result_vars["city_tax"].set(format_yen(result.city_tax))
        self.result_vars["total_tax"].set(format_yen(result.total_tax))
        self.result_vars["monthly_tax"].set(format_yen(result.monthly_tax))
        self.set_status(UI_TEXT["status_ready"], "ready")

    def collect_inputs(self) -> Dict[str, Any]:
        land_value = parse_non_negative_decimal(
            self.land_value_var.get(),
            field_label=UI_TEXT["label_land_value"],
            required_message=UI_TEXT["error_required_land"],
        )
        building_value = parse_non_negative_decimal(
            self.building_value_var.get(),
            field_label=UI_TEXT["label_building_value"],
            required_message=UI_TEXT["error_required_building"],
        )
        fixed_rate_percent = parse_non_negative_decimal(
            self.fixed_rate_var.get(),
            field_label=UI_TEXT["label_fixed_rate"],
            required_message=UI_TEXT["error_required_fixed_rate"],
        )
        city_rate_percent = parse_non_negative_decimal(
            self.city_rate_var.get(),
            field_label=UI_TEXT["label_city_rate"],
            required_message=UI_TEXT["error_required_city_rate"],
        )

        return {
            "land_value": land_value,
            "building_value": building_value,
            "residential_type": self.residential_type_var.get(),
            "city_tax_enabled": self.city_tax_enabled_var.get(),
            "fixed_rate_percent": fixed_rate_percent,
            "city_rate_percent": city_rate_percent,
        }

    def clear_results(self) -> None:
        for variable in self.result_vars.values():
            variable.set(UI_TEXT["value_placeholder"])

    def set_status(self, message: str, state: str) -> None:
        self.status_var.set(message)
        style_map = {
            "idle": "StatusIdle.TLabel",
            "ready": "StatusReady.TLabel",
            "error": "StatusError.TLabel",
        }
        self.status_label.configure(style=style_map.get(state, "StatusIdle.TLabel"))


def main() -> None:
    app = FixedTaxApp()
    app.mainloop()


if __name__ == "__main__":
    main()
