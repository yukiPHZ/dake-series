from __future__ import annotations

import html
import re
import tempfile
import webbrowser
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
import tkinter as tk
from tkinter import font as tkfont


APP_NAME = "Dake価格按分"
WINDOW_TITLE = "価格按分"
DISPLAY_NAME = "価格按分"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "main_heading": "売買価格を按分する",
    "description": "売買価格を、土地・建物の評価額比率で参考按分します。",
    "label_sale_price": "売買価格（税込）",
    "label_land_evaluation": "土地評価額",
    "label_building_evaluation": "建物評価額",
    "label_tax_rate": "消費税率",
    "label_tax_toggle": "消費税を計算する",
    "unit_yen": "円",
    "unit_percent": "%",
    "result_heading": "【参考計算結果】",
    "result_land_ratio": "土地比率",
    "result_building_ratio": "建物比率",
    "result_land_price": "土地価格",
    "result_building_price_gross": "建物価格（税込）",
    "result_building_price_net": "建物価格（税抜）",
    "result_building_tax": "うち消費税額",
    "result_placeholder": "-",
    "button_calculate": "計算する",
    "button_reset": "リセット",
    "button_copy_result": "結果をコピー",
    "button_open_print": "印刷用を開く",
    "status_display": "状態: {value}",
    "status_idle": "未計算",
    "status_calculating": "計算中",
    "status_completed": "計算完了",
    "status_copied": "コピーしました",
    "status_input_error": "入力エラー",
    "error_required": "{field}を入力してください。",
    "error_integer": "{field}は円単位の数値で入力してください。",
    "error_tax_rate": "消費税率は0以上100未満の数値で入力してください。",
    "error_sale_positive": "売買価格（税込）は0より大きい数値で入力してください。",
    "error_land_non_negative": "土地評価額は0円以上で入力してください。",
    "error_building_non_negative": "建物評価額は0円以上で入力してください。",
    "error_total_positive": "土地評価額と建物評価額の合計は0より大きい数値で入力してください。",
    "error_copy_unavailable": "先に計算してください。",
    "error_print_unavailable": "先に計算してください。",
    "error_print_output": "印刷用HTMLを開けませんでした。",
    "warning_land_zero": "土地評価額が0円です。土地価格の参考値は0円になります。",
    "warning_building_zero": "建物評価額が0円です。建物価格の参考値は0円になります。",
    "warning_tax_disabled": "消費税計算がオフのため、税抜額と消費税額は未計算で表示します。",
    "not_calculated": "未計算",
    "disclaimer_heading": "免責事項",
    "disclaimer_text": (
        "本ツールは、入力された評価額に基づく機械的な按分計算を行うものです。\n"
        "税務上の適正な区分を保証するものではありません。\n"
        "最終的な判断は税理士等の専門家にご確認ください。"
    ),
    "copy_result_template": (
        "{heading}\n"
        "{land_ratio_label}\t{land_ratio}\n"
        "{building_ratio_label}\t{building_ratio}\n"
        "{land_price_label}\t{land_price}\n"
        "{building_price_gross_label}\t{building_price_gross}\n"
        "{building_price_net_label}\t{building_price_net}\n"
        "{building_tax_label}\t{building_tax}"
    ),
    "footer_series": "シンプルそれDAKEシリーズ / 止まらない、迷わない、すぐ終わる。",
    "footer_estimate": "戸建買取査定",
    "footer_instagram": "Instagram",
    "footer_separator": " ｜ ",
    "footer_copyright": COPYRIGHT,
    "print_title": "参考計算結果",
}

LINK_TARGETS = {
    "estimate": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "instagram": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

COLORS = {
    "background": "#F6F7F9",
    "card": "#FFFFFF",
    "text": "#1E2430",
    "muted": "#667085",
    "border": "#E6EAF0",
    "accent": "#2F6FED",
    "accent_hover": "#2458BF",
    "selection_bg": "#EAF2FF",
    "selection_border": "#7AA7FF",
    "success": "#12B76A",
    "warning": "#B54708",
    "error": "#D92D20",
}

ICON_RELATIVE_PATH = Path("..") / ".." / "02_assets" / "dake_icon.ico"
OUTPUT_DIRECTORY_NAME = "DakePriceApportionment_Output"
INTEGER_PATTERN = re.compile(r"^[+-]?\d+$")
DECIMAL_PATTERN = re.compile(r"^[+-]?\d+(?:\.\d+)?$")


@dataclass(frozen=True)
class CalculationResult:
    sale_price: int
    land_evaluation: int
    building_evaluation: int
    total_evaluation: int
    tax_rate: Decimal
    tax_enabled: bool
    land_ratio: Decimal
    building_ratio: Decimal
    land_price: int
    building_price_gross: int
    building_price_net: int | None
    building_tax: int | None
    warnings: tuple[str, ...]


class PriceApportionmentApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.configure(bg=COLORS["background"])
        self.root.minsize(920, 760)

        self.font_family = self._choose_font_family()
        self.fonts = {
            "title": (self.font_family, 22, "bold"),
            "subtitle": (self.font_family, 11),
            "section": (self.font_family, 13, "bold"),
            "body": (self.font_family, 11),
            "body_bold": (self.font_family, 11, "bold"),
            "small": (self.font_family, 10),
            "button": (self.font_family, 10, "bold"),
            "result": (self.font_family, 12, "bold"),
            "status": (self.font_family, 10),
        }

        self.status_var = tk.StringVar()
        self.error_var = tk.StringVar(value="")
        self.warning_var = tk.StringVar(value="")

        self.input_vars = {
            "sale_price": tk.StringVar(),
            "land_evaluation": tk.StringVar(),
            "building_evaluation": tk.StringVar(),
            "tax_rate": tk.StringVar(value="10"),
        }
        self.tax_enabled_var = tk.BooleanVar(value=True)
        self.result_vars = {
            "land_ratio": tk.StringVar(value=UI_TEXT["result_placeholder"]),
            "building_ratio": tk.StringVar(value=UI_TEXT["result_placeholder"]),
            "land_price": tk.StringVar(value=UI_TEXT["result_placeholder"]),
            "building_price_gross": tk.StringVar(value=UI_TEXT["result_placeholder"]),
            "building_price_net": tk.StringVar(value=UI_TEXT["result_placeholder"]),
            "building_tax": tk.StringVar(value=UI_TEXT["result_placeholder"]),
        }

        self.entries: dict[str, tk.Entry] = {}
        self.action_buttons: list[tk.Button] = []
        self.last_result: CalculationResult | None = None

        self._set_window_icon()
        self._build_ui()
        self._bind_shortcuts()
        self._set_status("status_idle")

    def _choose_font_family(self) -> str:
        available = set(tkfont.families(self.root))
        for candidate in ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo"):
            if candidate in available:
                return candidate
        return "TkDefaultFont"

    def _set_window_icon(self) -> None:
        icon_path = (Path(__file__).resolve().parent / ICON_RELATIVE_PATH).resolve()
        if not icon_path.exists():
            return
        try:
            self.root.iconbitmap(default=str(icon_path))
        except tk.TclError:
            return

    def _build_ui(self) -> None:
        outer = tk.Frame(self.root, bg=COLORS["background"])
        outer.pack(fill="both", expand=True, padx=24, pady=24)

        header = tk.Frame(outer, bg=COLORS["background"])
        header.pack(fill="x", pady=(0, 16))
        tk.Label(
            header,
            text=UI_TEXT["main_heading"],
            font=self.fonts["title"],
            fg=COLORS["text"],
            bg=COLORS["background"],
            anchor="w",
        ).pack(anchor="w")
        tk.Label(
            header,
            text=UI_TEXT["description"],
            font=self.fonts["subtitle"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
            anchor="w",
        ).pack(anchor="w", pady=(6, 0))

        input_card = self._create_card(outer)
        input_card.pack(fill="x", pady=(0, 16))
        self._build_input_section(input_card)

        result_card = self._create_card(outer)
        result_card.pack(fill="x", pady=(0, 16))
        self._build_result_section(result_card)

        disclaimer_card = self._create_card(outer)
        disclaimer_card.pack(fill="x")
        self._build_disclaimer_section(disclaimer_card)

        spacer = tk.Frame(outer, bg=COLORS["background"])
        spacer.pack(fill="both", expand=True)

        status_row = tk.Frame(outer, bg=COLORS["background"])
        status_row.pack(fill="x", pady=(16, 10))
        tk.Label(
            status_row,
            textvariable=self.status_var,
            font=self.fonts["status"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
            anchor="w",
        ).pack(side="left")

        footer = tk.Frame(outer, bg=COLORS["background"])
        footer.pack(fill="x")
        tk.Label(
            footer,
            text=UI_TEXT["footer_series"],
            font=self.fonts["small"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
            anchor="w",
        ).pack(side="left")

        footer_right = tk.Frame(footer, bg=COLORS["background"])
        footer_right.pack(side="right")
        self._create_link_label(
            footer_right,
            UI_TEXT["footer_estimate"],
            LINK_TARGETS["estimate"],
        ).pack(side="left")
        self._create_footer_label(footer_right, UI_TEXT["footer_separator"]).pack(side="left")
        self._create_link_label(
            footer_right,
            UI_TEXT["footer_instagram"],
            LINK_TARGETS["instagram"],
        ).pack(side="left")
        self._create_footer_label(footer_right, UI_TEXT["footer_separator"]).pack(side="left")
        self._create_footer_label(footer_right, UI_TEXT["footer_copyright"]).pack(side="left")

    def _build_input_section(self, parent: tk.Frame) -> None:
        content = tk.Frame(parent, bg=COLORS["card"])
        content.pack(fill="x", padx=20, pady=20)
        content.grid_columnconfigure(1, weight=1)

        row = 0
        self._add_input_row(
            content,
            row,
            "sale_price",
            UI_TEXT["label_sale_price"],
            UI_TEXT["unit_yen"],
        )
        row += 1
        self._add_input_row(
            content,
            row,
            "land_evaluation",
            UI_TEXT["label_land_evaluation"],
            UI_TEXT["unit_yen"],
        )
        row += 1
        self._add_input_row(
            content,
            row,
            "building_evaluation",
            UI_TEXT["label_building_evaluation"],
            UI_TEXT["unit_yen"],
        )
        row += 1
        self._add_input_row(
            content,
            row,
            "tax_rate",
            UI_TEXT["label_tax_rate"],
            UI_TEXT["unit_percent"],
        )
        tax_toggle = tk.Checkbutton(
            content,
            text=UI_TEXT["label_tax_toggle"],
            variable=self.tax_enabled_var,
            font=self.fonts["body"],
            fg=COLORS["text"],
            bg=COLORS["card"],
            activebackground=COLORS["card"],
            activeforeground=COLORS["text"],
            selectcolor=COLORS["card"],
            highlightthickness=0,
            bd=0,
        )
        tax_toggle.grid(row=row, column=3, sticky="w", padx=(12, 0))

        row += 1
        error_label = tk.Label(
            content,
            textvariable=self.error_var,
            font=self.fonts["small"],
            fg=COLORS["error"],
            bg=COLORS["card"],
            justify="left",
            wraplength=720,
            anchor="w",
        )
        error_label.grid(row=row, column=0, columnspan=4, sticky="w", pady=(14, 0))

        row += 1
        warning_label = tk.Label(
            content,
            textvariable=self.warning_var,
            font=self.fonts["small"],
            fg=COLORS["warning"],
            bg=COLORS["card"],
            justify="left",
            wraplength=720,
            anchor="w",
        )
        warning_label.grid(row=row, column=0, columnspan=4, sticky="w", pady=(6, 0))

        row += 1
        button_row = tk.Frame(content, bg=COLORS["card"])
        button_row.grid(row=row, column=0, columnspan=4, sticky="ew", pady=(18, 0))
        for index in range(4):
            button_row.grid_columnconfigure(index, weight=1)

        calculate_button = self._create_button(
            button_row,
            UI_TEXT["button_calculate"],
            self._start_calculation,
            primary=True,
        )
        calculate_button.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        reset_button = self._create_button(
            button_row,
            UI_TEXT["button_reset"],
            self._reset_form,
        )
        reset_button.grid(row=0, column=1, sticky="ew", padx=8)

        copy_result_button = self._create_button(
            button_row,
            UI_TEXT["button_copy_result"],
            self._copy_result_text,
        )
        copy_result_button.grid(row=0, column=2, sticky="ew", padx=8)

        print_button = self._create_button(
            button_row,
            UI_TEXT["button_open_print"],
            self._open_print_html,
        )
        print_button.grid(row=0, column=3, sticky="ew", padx=(8, 0))

        self.action_buttons = [copy_result_button, print_button]
        self._toggle_action_buttons(False)

    def _build_result_section(self, parent: tk.Frame) -> None:
        content = tk.Frame(parent, bg=COLORS["card"])
        content.pack(fill="x", padx=20, pady=20)
        content.grid_columnconfigure(1, weight=1)

        tk.Label(
            content,
            text=UI_TEXT["result_heading"],
            font=self.fonts["section"],
            fg=COLORS["text"],
            bg=COLORS["card"],
            anchor="w",
        ).grid(row=0, column=0, columnspan=2, sticky="w")

        result_keys = [
            ("land_ratio", UI_TEXT["result_land_ratio"]),
            ("building_ratio", UI_TEXT["result_building_ratio"]),
            ("land_price", UI_TEXT["result_land_price"]),
            ("building_price_gross", UI_TEXT["result_building_price_gross"]),
            ("building_price_net", UI_TEXT["result_building_price_net"]),
            ("building_tax", UI_TEXT["result_building_tax"]),
        ]

        for index, (key, label_text) in enumerate(result_keys, start=1):
            tk.Label(
                content,
                text=label_text,
                font=self.fonts["body"],
                fg=COLORS["muted"],
                bg=COLORS["card"],
                anchor="w",
            ).grid(row=index, column=0, sticky="w", pady=(12, 0))
            tk.Label(
                content,
                textvariable=self.result_vars[key],
                font=self.fonts["result"],
                fg=COLORS["text"],
                bg=COLORS["card"],
                anchor="e",
            ).grid(row=index, column=1, sticky="e", pady=(12, 0))

    def _build_disclaimer_section(self, parent: tk.Frame) -> None:
        content = tk.Frame(parent, bg=COLORS["card"])
        content.pack(fill="x", padx=20, pady=20)
        tk.Label(
            content,
            text=UI_TEXT["disclaimer_heading"],
            font=self.fonts["section"],
            fg=COLORS["text"],
            bg=COLORS["card"],
            anchor="w",
        ).pack(anchor="w")
        tk.Label(
            content,
            text=UI_TEXT["disclaimer_text"],
            font=self.fonts["small"],
            fg=COLORS["muted"],
            bg=COLORS["card"],
            justify="left",
            anchor="w",
        ).pack(anchor="w", pady=(10, 0))

    def _create_card(self, parent: tk.Widget) -> tk.Frame:
        return tk.Frame(
            parent,
            bg=COLORS["card"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
            bd=0,
        )

    def _add_input_row(
        self,
        parent: tk.Frame,
        row: int,
        key: str,
        label_text: str,
        unit_text: str,
    ) -> None:
        tk.Label(
            parent,
            text=label_text,
            font=self.fonts["body"],
            fg=COLORS["text"],
            bg=COLORS["card"],
            anchor="w",
        ).grid(row=row, column=0, sticky="w", pady=(0 if row == 0 else 12, 0))

        entry = tk.Entry(
            parent,
            textvariable=self.input_vars[key],
            font=self.fonts["body"],
            fg=COLORS["text"],
            bg=COLORS["card"],
            relief="flat",
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            highlightcolor=COLORS["selection_border"],
            insertbackground=COLORS["text"],
            selectbackground=COLORS["selection_bg"],
            selectforeground=COLORS["text"],
            bd=0,
            width=24,
        )
        entry.grid(row=row, column=1, sticky="ew", padx=(16, 8), pady=(0 if row == 0 else 12, 0), ipady=8)
        entry.bind("<FocusOut>", lambda event, input_key=key: self._format_entry_value(input_key))
        self.entries[key] = entry

        tk.Label(
            parent,
            text=unit_text,
            font=self.fonts["body"],
            fg=COLORS["muted"],
            bg=COLORS["card"],
            anchor="w",
        ).grid(row=row, column=2, sticky="w", pady=(0 if row == 0 else 12, 0))

    def _create_button(
        self,
        parent: tk.Widget,
        text: str,
        command: callable,
        primary: bool = False,
    ) -> tk.Button:
        button = tk.Button(
            parent,
            text=text,
            command=command,
            font=self.fonts["button"],
            relief="flat",
            bd=0,
            cursor="hand2",
            padx=12,
            pady=10,
            activeforeground=COLORS["card"],
            highlightthickness=0,
        )
        if primary:
            button.configure(
                bg=COLORS["accent"],
                fg=COLORS["card"],
                activebackground=COLORS["accent_hover"],
            )
            button.bind("<Enter>", lambda event, widget=button: widget.configure(bg=COLORS["accent_hover"]))
            button.bind("<Leave>", lambda event, widget=button: widget.configure(bg=COLORS["accent"]))
        else:
            button.configure(
                bg=COLORS["card"],
                fg=COLORS["text"],
                activebackground=COLORS["selection_bg"],
                highlightthickness=1,
                highlightbackground=COLORS["border"],
                highlightcolor=COLORS["selection_border"],
            )
        return button

    def _create_link_label(self, parent: tk.Widget, text: str, url: str) -> tk.Label:
        label = tk.Label(
            parent,
            text=text,
            font=self.fonts["small"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
            cursor="hand2",
        )
        label.bind("<Enter>", lambda event, widget=label: widget.configure(fg=COLORS["accent"]))
        label.bind("<Leave>", lambda event, widget=label: widget.configure(fg=COLORS["muted"]))
        label.bind("<Button-1>", lambda event: webbrowser.open_new_tab(url))
        return label

    def _create_footer_label(self, parent: tk.Widget, text: str) -> tk.Label:
        return tk.Label(
            parent,
            text=text,
            font=self.fonts["small"],
            fg=COLORS["muted"],
            bg=COLORS["background"],
        )

    def _bind_shortcuts(self) -> None:
        for entry in self.entries.values():
            entry.bind("<Return>", lambda event: self._start_calculation())

    def _set_status(self, status_key: str) -> None:
        self.status_var.set(UI_TEXT["status_display"].format(value=UI_TEXT[status_key]))

    def _clear_feedback(self) -> None:
        self.error_var.set("")
        self.warning_var.set("")

    def _toggle_action_buttons(self, enabled: bool) -> None:
        state = tk.NORMAL if enabled else tk.DISABLED
        disabled_text = COLORS["muted"]
        for button in self.action_buttons:
            button.configure(state=state)
            if enabled:
                button.configure(fg=COLORS["text"])
            else:
                button.configure(fg=disabled_text)

    def _reset_form(self) -> None:
        for key in ("sale_price", "land_evaluation", "building_evaluation"):
            self.input_vars[key].set("")
        self.input_vars["tax_rate"].set("10")
        self.tax_enabled_var.set(True)
        self.last_result = None
        self._clear_feedback()
        self._set_result_placeholders()
        self._toggle_action_buttons(False)
        self._set_status("status_idle")
        self.entries["sale_price"].focus_set()

    def _set_result_placeholders(self) -> None:
        for variable in self.result_vars.values():
            variable.set(UI_TEXT["result_placeholder"])

    def _start_calculation(self) -> None:
        self._clear_feedback()
        self._set_status("status_calculating")
        self.root.after(1, self._perform_calculation)

    def _perform_calculation(self) -> None:
        try:
            result = self._calculate()
        except ValueError as error:
            self.last_result = None
            self._set_result_placeholders()
            self._toggle_action_buttons(False)
            self.error_var.set(str(error))
            self._set_status("status_input_error")
            return

        self.last_result = result
        self._apply_result(result)
        self.warning_var.set("\n".join(result.warnings))
        self._toggle_action_buttons(True)
        self._set_status("status_completed")

    def _calculate(self) -> CalculationResult:
        sale_price = self._parse_integer(self.input_vars["sale_price"].get(), UI_TEXT["label_sale_price"])
        land_evaluation = self._parse_integer(
            self.input_vars["land_evaluation"].get(),
            UI_TEXT["label_land_evaluation"],
        )
        building_evaluation = self._parse_integer(
            self.input_vars["building_evaluation"].get(),
            UI_TEXT["label_building_evaluation"],
        )
        tax_rate = self._parse_tax_rate(self.input_vars["tax_rate"].get())
        tax_enabled = self.tax_enabled_var.get()

        if sale_price <= 0:
            raise ValueError(UI_TEXT["error_sale_positive"])
        if land_evaluation < 0:
            raise ValueError(UI_TEXT["error_land_non_negative"])
        if building_evaluation < 0:
            raise ValueError(UI_TEXT["error_building_non_negative"])

        total_evaluation = land_evaluation + building_evaluation
        if total_evaluation <= 0:
            raise ValueError(UI_TEXT["error_total_positive"])

        warnings: list[str] = []
        if land_evaluation == 0:
            warnings.append(UI_TEXT["warning_land_zero"])
        if building_evaluation == 0:
            warnings.append(UI_TEXT["warning_building_zero"])
        if not tax_enabled:
            warnings.append(UI_TEXT["warning_tax_disabled"])

        total_decimal = Decimal(total_evaluation)
        land_ratio = Decimal(land_evaluation) / total_decimal
        building_ratio = Decimal(building_evaluation) / total_decimal
        land_price = self._round_yen(Decimal(sale_price) * land_ratio)
        building_price_gross = sale_price - land_price

        if tax_enabled:
            divisor = Decimal("1") + (tax_rate / Decimal("100"))
            building_price_net = self._round_yen(Decimal(building_price_gross) / divisor)
            building_tax = building_price_gross - building_price_net
        else:
            building_price_net = None
            building_tax = None

        return CalculationResult(
            sale_price=sale_price,
            land_evaluation=land_evaluation,
            building_evaluation=building_evaluation,
            total_evaluation=total_evaluation,
            tax_rate=tax_rate,
            tax_enabled=tax_enabled,
            land_ratio=land_ratio,
            building_ratio=building_ratio,
            land_price=land_price,
            building_price_gross=building_price_gross,
            building_price_net=building_price_net,
            building_tax=building_tax,
            warnings=tuple(warnings),
        )

    def _apply_result(self, result: CalculationResult) -> None:
        self.result_vars["land_ratio"].set(self._format_ratio(result.land_ratio))
        self.result_vars["building_ratio"].set(self._format_ratio(result.building_ratio))
        self.result_vars["land_price"].set(self._format_currency(result.land_price))
        self.result_vars["building_price_gross"].set(self._format_currency(result.building_price_gross))
        self.result_vars["building_price_net"].set(self._format_optional_currency(result.building_price_net))
        self.result_vars["building_tax"].set(self._format_optional_currency(result.building_tax))
        self._format_entry_value("sale_price")
        self._format_entry_value("land_evaluation")
        self._format_entry_value("building_evaluation")
        self._format_entry_value("tax_rate")

    def _copy_result_text(self) -> None:
        if self.last_result is None:
            self.error_var.set(UI_TEXT["error_copy_unavailable"])
            self._set_status("status_input_error")
            return
        payload = self._build_result_copy_text(self.last_result)
        self._copy_to_clipboard(payload)

    def _copy_to_clipboard(self, payload: str) -> None:
        self.root.clipboard_clear()
        self.root.clipboard_append(payload)
        self.root.update_idletasks()
        self._set_status("status_copied")

    def _open_print_html(self) -> None:
        if self.last_result is None:
            self.error_var.set(UI_TEXT["error_print_unavailable"])
            self._set_status("status_input_error")
            return

        try:
            output_directory = self._get_output_directory()
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = output_directory / f"price_apportionment_print_{timestamp}.html"
            output_path.write_text(self._build_print_html(self.last_result), encoding="utf-8")
            webbrowser.open_new_tab(output_path.as_uri())
            self._set_status("status_completed")
        except OSError:
            self.error_var.set(UI_TEXT["error_print_output"])
            self._set_status("status_input_error")

    def _get_output_directory(self) -> Path:
        documents_directory = Path.home() / "Documents"
        base_directory = documents_directory if documents_directory.exists() else Path(tempfile.gettempdir())
        output_directory = base_directory / OUTPUT_DIRECTORY_NAME
        output_directory.mkdir(parents=True, exist_ok=True)
        return output_directory

    def _build_print_html(self, result: CalculationResult) -> str:
        result_table = self._build_print_table_rows(self._build_result_rows(result))
        return f"""<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="utf-8">
  <title>{html.escape(UI_TEXT["print_title"])}</title>
  <style>
    @page {{
      size: A4 portrait;
      margin: 12mm;
    }}

    * {{
      box-sizing: border-box;
      color: #000000 !important;
      background: #ffffff !important;
      box-shadow: none !important;
    }}

    body {{
      margin: 0;
      font-family: "BIZ UDPGothic", "Yu Gothic UI", "Meiryo", sans-serif;
      font-size: 12px;
      line-height: 1.6;
    }}

    h1,
    p {{
      margin: 0;
      font-weight: 400;
    }}

    h1 {{
      font-size: 20px;
      margin-bottom: 10px;
    }}

    .page {{
      width: 100%;
    }}

    table {{
      width: 100%;
      border-collapse: collapse;
      table-layout: fixed;
    }}

    th,
    td {{
      border: 1px solid #000000;
      padding: 8px 10px;
      text-align: left;
      vertical-align: top;
      font-weight: 400;
    }}

    @media print {{
      html,
      body {{
        width: 210mm;
        height: 297mm;
      }}

      a {{
        color: #000000 !important;
        text-decoration: none !important;
      }}
    }}
  </style>
</head>
<body>
  <main class="page">
    <h1>{html.escape(UI_TEXT["print_title"])}</h1>
    <table>
      <tbody>
        {result_table}
      </tbody>
    </table>
  </main>
</body>
</html>
"""

    def _build_print_table_rows(self, rows: list[tuple[str, str]]) -> str:
        html_rows = []
        for label, value in rows:
            html_rows.append(
                "<tr>"
                f"<th>{html.escape(label)}</th>"
                f"<td>{html.escape(value)}</td>"
                "</tr>"
            )
        return "".join(html_rows)

    def _build_result_copy_text(self, result: CalculationResult) -> str:
        rows = self._build_result_rows(result)
        return UI_TEXT["copy_result_template"].format(
            heading=UI_TEXT["print_title"],
            land_ratio_label=rows[0][0],
            land_ratio=rows[0][1],
            building_ratio_label=rows[1][0],
            building_ratio=rows[1][1],
            land_price_label=rows[2][0],
            land_price=rows[2][1],
            building_price_gross_label=rows[3][0],
            building_price_gross=rows[3][1],
            building_price_net_label=rows[4][0],
            building_price_net=rows[4][1],
            building_tax_label=rows[5][0],
            building_tax=rows[5][1],
        )

    def _build_result_rows(self, result: CalculationResult) -> list[tuple[str, str]]:
        return [
            (UI_TEXT["result_land_ratio"], self._format_ratio(result.land_ratio)),
            (UI_TEXT["result_building_ratio"], self._format_ratio(result.building_ratio)),
            (UI_TEXT["result_land_price"], self._format_currency(result.land_price)),
            (UI_TEXT["result_building_price_gross"], self._format_currency(result.building_price_gross)),
            (UI_TEXT["result_building_price_net"], self._format_optional_currency(result.building_price_net)),
            (UI_TEXT["result_building_tax"], self._format_optional_currency(result.building_tax)),
        ]

    def _parse_integer(self, raw_value: str, field_name: str) -> int:
        normalized = raw_value.replace(",", "").strip()
        if not normalized:
            raise ValueError(UI_TEXT["error_required"].format(field=field_name))
        if not INTEGER_PATTERN.fullmatch(normalized):
            raise ValueError(UI_TEXT["error_integer"].format(field=field_name))
        return int(normalized)

    def _parse_tax_rate(self, raw_value: str) -> Decimal:
        normalized = raw_value.replace(",", "").strip()
        if not normalized:
            raise ValueError(UI_TEXT["error_required"].format(field=UI_TEXT["label_tax_rate"]))
        if not DECIMAL_PATTERN.fullmatch(normalized):
            raise ValueError(UI_TEXT["error_tax_rate"])
        tax_rate = Decimal(normalized)
        if tax_rate < 0 or tax_rate >= 100:
            raise ValueError(UI_TEXT["error_tax_rate"])
        return tax_rate

    def _format_entry_value(self, key: str) -> None:
        raw_value = self.input_vars[key].get().strip()
        if not raw_value:
            return

        normalized = raw_value.replace(",", "")
        if key == "tax_rate":
            if not DECIMAL_PATTERN.fullmatch(normalized):
                return
            self.input_vars[key].set(self._format_decimal_text(Decimal(normalized)))
            return

        if not INTEGER_PATTERN.fullmatch(normalized):
            return
        self.input_vars[key].set(f"{int(normalized):,}")

    def _format_ratio(self, ratio: Decimal) -> str:
        percentage = (ratio * Decimal("100")).quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP)
        return f"{self._format_decimal_text(percentage)}{UI_TEXT['unit_percent']}"

    def _format_tax_rate(self, tax_rate: Decimal) -> str:
        return f"{self._format_decimal_text(tax_rate)}{UI_TEXT['unit_percent']}"

    def _format_currency(self, amount: int) -> str:
        return f"{self._format_number(amount)}{UI_TEXT['unit_yen']}"

    def _format_optional_currency(self, amount: int | None) -> str:
        if amount is None:
            return UI_TEXT["not_calculated"]
        return self._format_currency(amount)

    def _format_number(self, amount: int) -> str:
        return f"{amount:,}"

    def _format_decimal_text(self, value: Decimal) -> str:
        if value == value.to_integral():
            return str(value.quantize(Decimal("1")))
        return format(value.normalize(), "f").rstrip("0").rstrip(".")

    def _round_yen(self, value: Decimal) -> int:
        return int(value.quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def main() -> None:
    root = tk.Tk()
    app = PriceApportionmentApp(root)
    app._reset_form()
    root.mainloop()


if __name__ == "__main__":
    main()
