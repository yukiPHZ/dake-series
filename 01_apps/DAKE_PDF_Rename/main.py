# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import re
import subprocess
import sys
import webbrowser
from dataclasses import dataclass
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, ttk

try:
    import ctypes
except Exception:
    ctypes = None

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    ROOT_CLASS = TkinterDnD.Tk
    DND_ENABLED = True
except Exception:
    DND_FILES = None
    ROOT_CLASS = tk.Tk
    DND_ENABLED = False


APP_NAME = "DakePDFファイル名整理"
WINDOW_TITLE = "DakePDFファイル名整理"
COPYRIGHT = "© 2026 しまリス不動産 - Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "main_title": "PDFのファイル名を整える",
    "main_description": "PDFをドロップして、書類種別と相手名を指定します。",
    "drop_title": "PDFをここにドロップ",
    "drop_subtitle": "複数PDFもまとめて追加できます。",
    "drop_subtitle_no_dnd": "ドラッグ操作が使えない環境では、PDFを選択してください。",
    "label_pdf_list": "選択中PDF",
    "label_doc_type": "書類種別",
    "label_person_name": "相手名",
    "person_name_placeholder": "例：山田",
    "button_select_pdf": "PDFを選択",
    "button_execute": "ファイル名を整える",
    "button_clear": "クリア",
    "status_idle": "PDFを追加してください。",
    "status_ready": "{count}件のPDFを整理できます。",
    "status_processing": "処理中...",
    "status_complete": "ファイル名の整理が完了しました。",
    "status_error": "確認が必要です。",
    "error_no_pdf": "PDFがまだ選択されていません。PDFを追加してから実行してください。",
    "error_no_name": "相手名を入力してください。",
    "error_invalid_file": "PDF以外のファイルが含まれています。PDFファイルだけを追加してください。",
    "error_missing_file": "選択したPDFが見つかりません。ファイルの場所を確認してください。",
    "error_file_busy": "PDFが開かれている可能性があります。閉じてからもう一度お試しください。",
    "error_name_generation": "同名ファイルを避けた名前を作成できませんでした。",
    "error_rename_failed": "ファイル名の変更ができませんでした。ファイルの状態を確認してください。",
    "complete_message": "{count}件のPDFファイル名を整えました。",
    "dialog_complete_title": "完了",
    "dialog_error_title": "確認",
    "file_dialog_title": "PDFを選択",
    "filetype_pdf": "PDFファイル",
    "filetype_all": "すべてのファイル",
    "list_empty": "PDFが追加されていません。",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_link_1": "戸建買取相談所",
    "footer_link_2": "Instagram",
    "footer_separator": " ・ ",
    "footer_copyright": COPYRIGHT,
    "doc_types": [
        "売買契約書",
        "重要事項説明書",
        "付帯設備表",
        "告知書",
        "領収書手付金",
        "領収書残代金",
        "媒介契約書",
        "仲介手数料約定書",
        "覚書",
    ],
}

COLORS = {
    "base_bg": "#F6F7F9",
    "card_bg": "#FFFFFF",
    "text": "#1E2430",
    "muted": "#667085",
    "border": "#E6EAF0",
    "accent": "#2F6FED",
    "accent_hover": "#2458BF",
    "selection_bg": "#EAF2FF",
    "success": "#12B76A",
    "error": "#D92D20",
    "idle_bg": "#F2F4F7",
    "disabled_bg": "#E8ECF3",
    "disabled_fg": "#98A2B3",
}

LINK_URLS = {
    "footer_link_1": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "footer_link_2": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

COMMON_ICON_RELATIVE = Path("..") / ".." / "02_assets" / "dake_icon.ico"
WINDOW_WIDTH = 820
WINDOW_HEIGHT = 620
MAX_SEQUENCE_NUMBER = 9999
FONT_CANDIDATES = ["BIZ UDPGothic", "Yu Gothic UI", "Meiryo"]
INVALID_CHAR_TRANSLATION = str.maketrans(
    {
        "\\": "＼",
        "/": "／",
        ":": "：",
        "*": "＊",
        "?": "？",
        '"': "”",
        "<": "＜",
        ">": "＞",
        "|": "｜",
    }
)


@dataclass(frozen=True)
class RenameResult:
    old_path: Path
    new_path: Path


class RenameError(Exception):
    def __init__(self, code: str, original: Exception | None = None):
        super().__init__(code)
        self.code = code
        self.original = original


def app_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def choose_font_family(root: tk.Tk) -> str:
    available = set(tkfont.families(root))
    for family in FONT_CANDIDATES:
        if family in available:
            return family
    return "TkDefaultFont"


def set_windows_app_id() -> None:
    if not sys.platform.startswith("win") or ctypes is None:
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            "Shimarisu.DakePDFRename"
        )
    except Exception:
        pass


def icon_candidates() -> list[Path]:
    base = app_base_dir()
    candidates = [
        Path(__file__).resolve().parent / COMMON_ICON_RELATIVE,
        base / COMMON_ICON_RELATIVE,
        base / ".." / ".." / "02_assets" / "dake_icon.ico",
        base / ".." / ".." / ".." / "02_assets" / "dake_icon.ico",
        Path.cwd() / COMMON_ICON_RELATIVE,
    ]
    return [candidate.resolve() for candidate in candidates]


def apply_window_icon(root: tk.Tk) -> None:
    for icon_path in icon_candidates():
        if not icon_path.exists():
            continue
        try:
            root.iconbitmap(str(icon_path))
        except Exception:
            pass
        try:
            root.iconbitmap(default=str(icon_path))
        except Exception:
            pass
        return


def open_folder(path: Path) -> None:
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception:
        pass


def open_web_link(key: str) -> None:
    url = LINK_URLS.get(key)
    if url:
        webbrowser.open_new(url)


def normalize_path_key(path: Path) -> str:
    return os.path.normcase(str(path.resolve()))


def needs_rename(source_path: Path, target_path: Path) -> bool:
    same_folder = normalize_path_key(source_path.parent) == normalize_path_key(
        target_path.parent
    )
    return not same_folder or source_path.name != target_path.name


def sanitize_filename_part(value: str) -> str:
    translated = value.translate(INVALID_CHAR_TRANSLATION)
    without_controls = "".join(ch for ch in translated if ord(ch) >= 32)
    compacted = re.sub(r"\s+", " ", without_controls)
    return compacted.strip().strip(". ")


def normalize_person_name(value: str) -> str:
    cleaned = sanitize_filename_part(value)
    if not cleaned:
        return ""
    without_suffix = cleaned.rstrip("様")
    return f"{without_suffix}様" if without_suffix else "様"


def build_target_stem(doc_type: str, person_name: str) -> str:
    safe_doc_type = sanitize_filename_part(doc_type)
    safe_person_name = sanitize_filename_part(person_name)
    return f"{safe_doc_type}_{safe_person_name}"


def find_unique_target(source_path: Path, stem: str, reserved: set[str]) -> Path:
    folder = source_path.parent
    source_key = normalize_path_key(source_path)

    for number in range(1, MAX_SEQUENCE_NUMBER + 1):
        suffix = "" if number == 1 else f"_{number}"
        candidate = folder / f"{stem}{suffix}.pdf"
        candidate_key = normalize_path_key(candidate)
        if candidate_key in reserved and candidate_key != source_key:
            continue
        if candidate.exists() and candidate_key != source_key:
            continue
        return candidate

    raise RenameError("name_generation")


def rename_pdf_files(paths: list[Path], doc_type: str, person_name: str) -> list[RenameResult]:
    if not paths:
        raise RenameError("no_pdf")

    stem = build_target_stem(doc_type, person_name)
    if not stem:
        raise RenameError("name_generation")

    reserved: set[str] = set()
    plan: list[tuple[Path, Path]] = []

    for source_path in paths:
        if not source_path.exists():
            raise RenameError("missing_file")
        if source_path.suffix.lower() != ".pdf":
            raise RenameError("invalid_file")
        target_path = find_unique_target(source_path, stem, reserved)
        reserved.add(normalize_path_key(target_path))
        plan.append((source_path, target_path))

    results: list[RenameResult] = []
    for source_path, target_path in plan:
        try:
            if needs_rename(source_path, target_path):
                source_path.rename(target_path)
            results.append(RenameResult(old_path=source_path, new_path=target_path))
        except PermissionError as exc:
            raise RenameError("file_busy", exc) from exc
        except FileExistsError as exc:
            raise RenameError("name_generation", exc) from exc
        except OSError as exc:
            raise RenameError("rename_failed", exc) from exc

    return results


def shorten_path(path: Path, max_chars: int = 76) -> str:
    text = str(path)
    if len(text) <= max_chars:
        return text
    keep = max_chars - 5
    front = keep // 2
    back = keep - front
    return f"{text[:front]} ... {text[-back:]}"


class PdfRenameApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.font_family = choose_font_family(root)
        self.pdf_paths: list[Path] = []
        self.drop_widgets: list[tk.Widget] = []
        self.status_var = tk.StringVar(value=UI_TEXT["status_idle"])
        self.doc_type_var = tk.StringVar(value=UI_TEXT["doc_types"][0])
        self.person_name_var = tk.StringVar()
        self.empty_list_var = tk.StringVar(value=UI_TEXT["list_empty"])
        self.execute_button: tk.Button | None = None
        self.drop_area: tk.Frame | None = None
        self.drop_title_label: tk.Label | None = None
        self.drop_subtitle_label: tk.Label | None = None
        self.listbox: tk.Listbox | None = None

        self._setup_root()
        self._setup_styles()
        self._build_ui()
        self._refresh_list()
        self._update_action_state()

    def _setup_root(self) -> None:
        set_windows_app_id()
        self.root.title(WINDOW_TITLE)
        self.root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
        self.root.minsize(720, 560)
        self.root.configure(bg=COLORS["base_bg"])
        self.root.option_add("*Font", (self.font_family, 10))
        apply_window_icon(self.root)

    def _setup_styles(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(
            "Dake.TCombobox",
            fieldbackground=COLORS["card_bg"],
            background=COLORS["card_bg"],
            foreground=COLORS["text"],
            bordercolor=COLORS["border"],
            lightcolor=COLORS["border"],
            darkcolor=COLORS["border"],
            arrowcolor=COLORS["accent"],
            padding=(8, 6),
        )

    def _build_ui(self) -> None:
        outer = tk.Frame(self.root, bg=COLORS["base_bg"])
        outer.pack(fill="both", expand=True, padx=34, pady=26)
        outer.grid_columnconfigure(0, weight=1)
        outer.grid_rowconfigure(1, weight=1)

        self._build_header(outer)
        self._build_card(outer)
        self._build_footer(outer)

    def _build_header(self, parent: tk.Frame) -> None:
        header = tk.Frame(parent, bg=COLORS["base_bg"])
        header.grid(row=0, column=0, sticky="ew", pady=(0, 18))

        title = tk.Label(
            header,
            text=UI_TEXT["main_title"],
            bg=COLORS["base_bg"],
            fg=COLORS["text"],
            font=(self.font_family, 22, "bold"),
        )
        title.pack(anchor="w")

        description = tk.Label(
            header,
            text=UI_TEXT["main_description"],
            bg=COLORS["base_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 10),
        )
        description.pack(anchor="w", pady=(6, 0))

    def _build_card(self, parent: tk.Frame) -> None:
        shell = tk.Frame(parent, bg=COLORS["border"])
        shell.grid(row=1, column=0, sticky="nsew")
        shell.grid_columnconfigure(0, weight=1)
        shell.grid_rowconfigure(0, weight=1)

        card = tk.Frame(shell, bg=COLORS["card_bg"], padx=24, pady=22)
        card.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(1, weight=1)

        self._build_drop_area(card)
        self._build_file_list(card)
        self._build_form(card)
        self._build_actions(card)

    def _build_drop_area(self, parent: tk.Frame) -> None:
        self.drop_area = tk.Frame(
            parent,
            bg=COLORS["base_bg"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
            padx=18,
            pady=18,
            cursor="hand2",
        )
        self.drop_area.grid(row=0, column=0, sticky="ew")
        self.drop_area.grid_columnconfigure(0, weight=1)

        self.drop_title_label = tk.Label(
            self.drop_area,
            text=UI_TEXT["drop_title"],
            bg=COLORS["base_bg"],
            fg=COLORS["text"],
            font=(self.font_family, 13, "bold"),
        )
        self.drop_title_label.grid(row=0, column=0)

        subtitle_key = "drop_subtitle" if DND_ENABLED else "drop_subtitle_no_dnd"
        self.drop_subtitle_label = tk.Label(
            self.drop_area,
            text=UI_TEXT[subtitle_key],
            bg=COLORS["base_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 9),
        )
        self.drop_subtitle_label.grid(row=1, column=0, pady=(5, 0))

        self.drop_widgets = [
            self.drop_area,
            self.drop_title_label,
            self.drop_subtitle_label,
        ]

        for widget in self.drop_widgets:
            widget.bind("<Button-1>", self._select_pdf_files)

        if DND_ENABLED and DND_FILES is not None:
            self.drop_area.drop_target_register(DND_FILES)
            self.drop_area.dnd_bind("<<DropEnter>>", self._on_drop_enter)
            self.drop_area.dnd_bind("<<DropLeave>>", self._on_drop_leave)
            self.drop_area.dnd_bind("<<Drop>>", self._on_drop)

    def _build_file_list(self, parent: tk.Frame) -> None:
        list_header = tk.Frame(parent, bg=COLORS["card_bg"])
        list_header.grid(row=1, column=0, sticky="nsew", pady=(18, 0))
        list_header.grid_columnconfigure(0, weight=1)
        list_header.grid_rowconfigure(1, weight=1)

        label = tk.Label(
            list_header,
            text=UI_TEXT["label_pdf_list"],
            bg=COLORS["card_bg"],
            fg=COLORS["text"],
            font=(self.font_family, 10, "bold"),
        )
        label.grid(row=0, column=0, sticky="w")

        list_shell = tk.Frame(list_header, bg=COLORS["border"])
        list_shell.grid(row=1, column=0, sticky="nsew", pady=(8, 0))
        list_shell.grid_columnconfigure(0, weight=1)
        list_shell.grid_rowconfigure(0, weight=1)

        self.listbox = tk.Listbox(
            list_shell,
            activestyle="none",
            bg=COLORS["card_bg"],
            fg=COLORS["text"],
            selectbackground=COLORS["selection_bg"],
            selectforeground=COLORS["text"],
            relief="flat",
            borderwidth=0,
            height=6,
            highlightthickness=0,
            font=(self.font_family, 9),
        )
        self.listbox.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)

        scrollbar = tk.Scrollbar(list_shell, orient="vertical", command=self.listbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns", pady=1)
        self.listbox.configure(yscrollcommand=scrollbar.set)

    def _build_form(self, parent: tk.Frame) -> None:
        form = tk.Frame(parent, bg=COLORS["card_bg"])
        form.grid(row=2, column=0, sticky="ew", pady=(18, 0))
        form.grid_columnconfigure(1, weight=1)

        doc_label = tk.Label(
            form,
            text=UI_TEXT["label_doc_type"],
            bg=COLORS["card_bg"],
            fg=COLORS["text"],
            font=(self.font_family, 10, "bold"),
        )
        doc_label.grid(row=0, column=0, sticky="w", padx=(0, 14), pady=(0, 12))

        doc_combo = ttk.Combobox(
            form,
            textvariable=self.doc_type_var,
            values=UI_TEXT["doc_types"],
            state="readonly",
            style="Dake.TCombobox",
            width=28,
        )
        doc_combo.grid(row=0, column=1, sticky="ew", pady=(0, 12))

        name_label = tk.Label(
            form,
            text=UI_TEXT["label_person_name"],
            bg=COLORS["card_bg"],
            fg=COLORS["text"],
            font=(self.font_family, 10, "bold"),
        )
        name_label.grid(row=1, column=0, sticky="w", padx=(0, 14))

        name_entry = tk.Entry(
            form,
            textvariable=self.person_name_var,
            bg=COLORS["card_bg"],
            fg=COLORS["text"],
            relief="solid",
            borderwidth=1,
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            highlightcolor=COLORS["accent"],
            insertbackground=COLORS["text"],
            font=(self.font_family, 11),
        )
        name_entry.grid(row=1, column=1, sticky="ew", ipady=7)
        name_entry.bind("<KeyRelease>", self._on_name_change)

    def _build_actions(self, parent: tk.Frame) -> None:
        actions = tk.Frame(parent, bg=COLORS["card_bg"])
        actions.grid(row=3, column=0, sticky="ew", pady=(22, 0))
        actions.grid_columnconfigure(0, weight=1)

        secondary = tk.Frame(actions, bg=COLORS["card_bg"])
        secondary.grid(row=0, column=0, sticky="w")

        select_button = self._create_secondary_button(
            secondary,
            UI_TEXT["button_select_pdf"],
            self._select_pdf_files,
        )
        select_button.pack(side="left", padx=(0, 8))

        clear_button = self._create_secondary_button(
            secondary,
            UI_TEXT["button_clear"],
            self._clear_files,
        )
        clear_button.pack(side="left")

        self.execute_button = self._create_primary_button(
            actions,
            UI_TEXT["button_execute"],
            self._execute_rename,
        )
        self.execute_button.grid(row=0, column=1, sticky="e")

        status = tk.Label(
            actions,
            textvariable=self.status_var,
            bg=COLORS["card_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 9),
            anchor="w",
        )
        status.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(14, 0))

    def _build_footer(self, parent: tk.Frame) -> None:
        footer = tk.Frame(parent, bg=COLORS["base_bg"])
        footer.grid(row=2, column=0, sticky="ew", pady=(16, 0))

        left = tk.Label(
            footer,
            text=UI_TEXT["footer_left"],
            bg=COLORS["base_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 8),
        )
        left.pack(side="left")

        separator_1 = tk.Label(
            footer,
            text=UI_TEXT["footer_separator"],
            bg=COLORS["base_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 8),
        )
        separator_1.pack(side="left")

        link_1 = self._create_footer_link(footer, "footer_link_1")
        link_1.pack(side="left")

        separator_2 = tk.Label(
            footer,
            text=UI_TEXT["footer_separator"],
            bg=COLORS["base_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 8),
        )
        separator_2.pack(side="left")

        link_2 = self._create_footer_link(footer, "footer_link_2")
        link_2.pack(side="left")

        copyright_label = tk.Label(
            footer,
            text=UI_TEXT["footer_copyright"],
            bg=COLORS["base_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 8),
        )
        copyright_label.pack(side="right")

    def _create_primary_button(self, parent: tk.Widget, text: str, command) -> tk.Button:
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=COLORS["accent"],
            fg=COLORS["card_bg"],
            activebackground=COLORS["accent_hover"],
            activeforeground=COLORS["card_bg"],
            disabledforeground=COLORS["disabled_fg"],
            relief="flat",
            borderwidth=0,
            padx=22,
            pady=10,
            cursor="hand2",
            font=(self.font_family, 10, "bold"),
        )

    def _create_secondary_button(self, parent: tk.Widget, text: str, command) -> tk.Button:
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=COLORS["card_bg"],
            fg=COLORS["text"],
            activebackground=COLORS["selection_bg"],
            activeforeground=COLORS["text"],
            relief="solid",
            borderwidth=1,
            padx=14,
            pady=8,
            cursor="hand2",
            font=(self.font_family, 9),
        )

    def _create_footer_link(self, parent: tk.Widget, key: str) -> tk.Label:
        link = tk.Label(
            parent,
            text=UI_TEXT[key],
            bg=COLORS["base_bg"],
            fg=COLORS["accent"],
            font=(self.font_family, 8, "underline"),
            cursor="hand2",
        )
        link.bind("<Button-1>", lambda _event, link_key=key: open_web_link(link_key))
        return link

    def _select_pdf_files(self, _event=None) -> None:
        selected = filedialog.askopenfilenames(
            parent=self.root,
            title=UI_TEXT["file_dialog_title"],
            filetypes=(
                (UI_TEXT["filetype_pdf"], "*.pdf"),
                (UI_TEXT["filetype_all"], "*.*"),
            ),
        )
        if selected:
            self._add_files([Path(value) for value in selected])

    def _on_drop_enter(self, _event=None):
        self._set_drop_active(True)
        return "copy"

    def _on_drop_leave(self, _event=None):
        self._set_drop_active(False)
        return "copy"

    def _on_drop(self, event):
        self._set_drop_active(False)
        try:
            raw_paths = self.root.tk.splitlist(event.data)
        except Exception:
            raw_paths = event.data.split()
        self._add_files([Path(value) for value in raw_paths])
        return "copy"

    def _set_drop_active(self, is_active: bool) -> None:
        bg = COLORS["selection_bg"] if is_active else COLORS["base_bg"]
        border = COLORS["accent"] if is_active else COLORS["border"]
        if self.drop_area is not None:
            self.drop_area.configure(bg=bg, highlightbackground=border)
        for widget in self.drop_widgets:
            try:
                widget.configure(bg=bg)
            except Exception:
                pass

    def _add_files(self, paths: list[Path]) -> None:
        if not paths:
            return

        invalid_paths = [path for path in paths if path.suffix.lower() != ".pdf"]
        if invalid_paths:
            self._show_error(UI_TEXT["error_invalid_file"])
            return

        missing_paths = [path for path in paths if not path.exists()]
        if missing_paths:
            self._show_error(UI_TEXT["error_missing_file"])
            return

        existing = {normalize_path_key(path) for path in self.pdf_paths}
        for path in paths:
            resolved = path.resolve()
            path_key = normalize_path_key(resolved)
            if path_key not in existing:
                self.pdf_paths.append(resolved)
                existing.add(path_key)

        self._refresh_list()
        self._update_status()
        self._update_action_state()

    def _clear_files(self, _event=None) -> None:
        self.pdf_paths.clear()
        self._refresh_list()
        self._update_status()
        self._update_action_state()

    def _on_name_change(self, _event=None) -> None:
        self._update_action_state()

    def _execute_rename(self) -> None:
        if not self.pdf_paths:
            self._show_error(UI_TEXT["error_no_pdf"])
            return

        person_name = normalize_person_name(self.person_name_var.get())
        if not person_name:
            self._show_error(UI_TEXT["error_no_name"])
            return

        doc_type = self.doc_type_var.get().strip() or UI_TEXT["doc_types"][0]
        self._set_processing(True)
        self.root.after(80, lambda: self._run_rename(doc_type, person_name))

    def _run_rename(self, doc_type: str, person_name: str) -> None:
        try:
            results = rename_pdf_files(list(self.pdf_paths), doc_type, person_name)
        except RenameError as exc:
            self._set_processing(False)
            self._show_error(self._message_for_error(exc.code))
            return

        self.pdf_paths = [result.new_path for result in results]
        self._refresh_list()
        self._set_processing(False)
        self.status_var.set(UI_TEXT["complete_message"].format(count=len(results)))
        messagebox.showinfo(
            UI_TEXT["dialog_complete_title"],
            UI_TEXT["complete_message"].format(count=len(results)),
            parent=self.root,
        )
        if results:
            open_folder(results[0].new_path.parent)

    def _message_for_error(self, code: str) -> str:
        mapping = {
            "no_pdf": "error_no_pdf",
            "invalid_file": "error_invalid_file",
            "missing_file": "error_missing_file",
            "file_busy": "error_file_busy",
            "name_generation": "error_name_generation",
            "rename_failed": "error_rename_failed",
        }
        return UI_TEXT[mapping.get(code, "error_rename_failed")]

    def _show_error(self, message: str) -> None:
        self.status_var.set(message)
        messagebox.showwarning(UI_TEXT["dialog_error_title"], message, parent=self.root)
        self._update_action_state()

    def _set_processing(self, is_processing: bool) -> None:
        if is_processing:
            self.status_var.set(UI_TEXT["status_processing"])
            if self.execute_button is not None:
                self.execute_button.configure(
                    state="disabled",
                    bg=COLORS["disabled_bg"],
                    cursor="arrow",
                )
            self.root.update_idletasks()
            return

        self._update_action_state()
        self._update_status()

    def _update_status(self) -> None:
        count = len(self.pdf_paths)
        if count == 0:
            self.status_var.set(UI_TEXT["status_idle"])
        else:
            self.status_var.set(UI_TEXT["status_ready"].format(count=count))

    def _update_action_state(self) -> None:
        if self.execute_button is None:
            return
        is_ready = bool(self.pdf_paths) and bool(normalize_person_name(self.person_name_var.get()))
        if is_ready:
            self.execute_button.configure(
                state="normal",
                bg=COLORS["accent"],
                fg=COLORS["card_bg"],
                cursor="hand2",
            )
        else:
            self.execute_button.configure(
                state="disabled",
                bg=COLORS["disabled_bg"],
                fg=COLORS["disabled_fg"],
                cursor="arrow",
            )

    def _refresh_list(self) -> None:
        if self.listbox is None:
            return
        self.listbox.delete(0, tk.END)
        if not self.pdf_paths:
            self.listbox.insert(tk.END, UI_TEXT["list_empty"])
            self.listbox.itemconfig(0, fg=COLORS["muted"])
            return
        for path in self.pdf_paths:
            self.listbox.insert(tk.END, shorten_path(path))


def main() -> None:
    root = ROOT_CLASS()
    PdfRenameApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
