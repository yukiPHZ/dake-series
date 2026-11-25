from __future__ import annotations

import queue
import subprocess
import sys
import threading
import tkinter as tk
from dataclasses import dataclass
from enum import Enum, auto
from pathlib import Path
from tkinter import messagebox
from urllib.parse import unquote, urlparse

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except Exception as import_error:  # pragma: no cover - handled at runtime.
    DND_FILES = "DND_Files"
    TkinterDnD = None
    TKDND_IMPORT_ERROR = import_error
else:
    TKDND_IMPORT_ERROR = None


UI_TEXT = {
    "app_title": "しまりすくん",
    "drop_main": "ファイルをここに入れてください",
    "drop_sub": "投げるだけで、次が進みます",
    "drop_over_main": "ここにドロップ",
    "checking_main": "内容を確認しています…",
    "checking_sub": "",
    "confirm_sub": "必要な処理だけ選んでください",
    "running_main": "処理しています…",
    "running_sub": "終わるまでこのままお待ちください",
    "done_main": "完了しました",
    "ready_again": "もう一度ファイルを入れられます",
    "planned_main": "この処理は次段階で対応予定です",
    "planned_sub": "PDFまたは画像ファイルを入れてください",
    "decision_pdf_many": "PDFが複数ありますね → 1つにまとめます",
    "decision_pdf_one": "1つのPDFですね → ページごとに分けられます",
    "decision_image_many": "画像が複数ありますね → そろえて処理できます",
    "decision_image_one": "画像ですね → PDF化またはリサイズできます",
    "decision_mixed": "PDFと画像がありますね → どちらにそろえますか？",
    "choice_split_all": "すべて分ける",
    "choice_split_select": "選んで抽出する",
    "choice_to_pdf": "PDFにする",
    "choice_resize": "リサイズする",
    "choice_all_pdf": "全部PDFにする",
    "choice_all_images": "全部画像にする",
    "error_folder_main": "フォルダはまだ対応していません",
    "error_folder_sub": "PDFまたは画像ファイルを入れてください",
    "error_unknown_main": "対応できるファイルがありません",
    "error_unknown_sub": "PDFまたは画像ファイルを入れてください",
    "error_mixed_unknown_main": "対応できないファイルが含まれています",
    "error_mixed_unknown_sub": "PDFまたは画像ファイルだけを入れてください",
    "error_read_main": "このファイルは読み込めませんでした",
    "error_read_sub": "PDFまたは画像ファイルを入れてください",
    "error_dake_missing_main": "処理アプリが見つかりません",
    "error_dake_missing_sub": "DAKEアプリの場所を確認してください",
    "error_dake_missing_detail": "探した場所：\n{path}\n\nDAKEアプリをビルド済みか、exe名が合っているか確認してください。",
    "error_dake_launch_main": "処理アプリを起動できませんでした",
    "error_dake_launch_sub": "このDAKEアプリは、まだしまりすくん連携に対応していない可能性があります",
    "error_dnd_missing": "ドラッグ＆ドロップ部品を読み込めませんでした。\ntkinterdnd2 をインストールしてください。",
    "status_waiting": "待機中",
    "status_loading": "読み込み中",
    "status_judging": "判断中",
    "status_running": "処理中",
    "status_done": "完了しました",
    "status_ready_again": "もう一度ファイルを入れられます",
    "status_confirming": "確認待ち",
    "status_error": "やり直せます",
    "footer": "しまりすくん ｜ © 2026 しまりす不動産 - Vibe-Coded by Yukihiko Kikuta",
}

IMAGE_EXTENSIONS = {
    ".jpg",
    ".jpeg",
    ".png",
    ".webp",
    ".bmp",
    ".tif",
    ".tiff",
}
PDF_EXTENSION = ".pdf"
SHORTCUT_EXTENSION = ".lnk"

def resolve_app_dir() -> Path:
    source_dir = Path(__file__).resolve().parent
    if not getattr(sys, "frozen", False):
        return source_dir

    exe_dir = Path(sys.executable).resolve().parent
    if exe_dir.name.lower() == "dist":
        return exe_dir.parent
    return exe_dir


APP_DIR = resolve_app_dir()
DEV_ROOT = APP_DIR.parents[2]
DAKE_APPS_ROOT = DEV_ROOT / "DAKE_series" / "01_apps"

ICON_PATH = DEV_ROOT / "DAKE_series" / "02_assets" / "dake_icon.ico"

DAKE_APPS = {
    "pdf_merge": {
        "label": "PDF結合",
        "exe": DAKE_APPS_ROOT / "DAKE_PDF_Merge" / "dist" / "DakePDF_Merge.exe",
    },
    "pdf_split_one": {
        "label": "PDF分割One",
        "exe": DAKE_APPS_ROOT / "DAKE_PDF_SplitOne" / "dist" / "DakePDF_Split_One.exe",
    },
    "pdf_split_select": {
        "label": "PDF分割Select",
        "exe": DAKE_APPS_ROOT / "DAKE_PDF_SplitSelect" / "dist" / "DakePDF_Split_Select.exe",
    },
    "image_to_pdf": {
        "label": "画像PDF化",
        "exe": DAKE_APPS_ROOT / "DAKE_Image_ToPDF" / "dist" / "DakeImageToPDF.exe",
    },
    "image_resize": {
        "label": "画像リサイズ",
        "exe": DAKE_APPS_ROOT / "DAKE_Image_Resize" / "dist" / "Dake_Image_Resize.exe",
    },
    "pdf_to_images": {
        "label": "PDF画像化",
        "exe": DAKE_APPS_ROOT / "DAKE_PDF_ToImages" / "dist" / "DakePDF_to_Images.exe",
    },
}

COLORS = {
    "window_bg": "#ffffff",
    "drop_bg": "#fbfbfb",
    "drop_bg_active": "#f0f7f3",
    "drop_border": "#d7dce2",
    "drop_border_active": "#6fbf8f",
    "main_text": "#20242a",
    "sub_text": "#6b7280",
    "status_text": "#7b8190",
    "footer_text": "#9aa0aa",
    "button_bg": "#ffffff",
    "button_active": "#eef6f1",
}


class AppState(Enum):
    WAITING = auto()
    CHECKING = auto()
    JUDGING = auto()
    CONFIRMING = auto()
    RUNNING = auto()
    DONE = auto()
    ERROR = auto()


@dataclass(frozen=True)
class Choice:
    label_key: str
    action: str
    planned_only: bool = False


@dataclass(frozen=True)
class Decision:
    main_text_key: str
    inputs: tuple[Path, ...]
    auto_action: str | None = None
    choices: tuple[Choice, ...] = ()


def check_dake_app(app_key: str) -> tuple[bool, Path | None, str]:
    app = DAKE_APPS.get(app_key)
    if not app:
        return False, None, "未登録のDAKEアプリです"

    exe_path = app["exe"]
    if exe_path.exists():
        return True, exe_path, ""

    return False, exe_path, "処理アプリが見つかりません"


def is_supported_image(path: Path) -> bool:
    return path.suffix.lower() in IMAGE_EXTENSIONS


def is_pdf(path: Path) -> bool:
    return path.suffix.lower() == PDF_EXTENSION


class ShimarisuKunApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.state = AppState.WAITING
        self.event_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.current_inputs: tuple[Path, ...] = ()
        self.status_after_id: str | None = None
        self.reset_after_id: str | None = None
        self.dot_index = 0

        self.main_var = tk.StringVar(value=UI_TEXT["drop_main"])
        self.sub_var = tk.StringVar(value=UI_TEXT["drop_sub"])
        self.status_var = tk.StringVar(value=UI_TEXT["status_waiting"])

        self.choice_buttons: list[tk.Button] = []

        self._configure_root()
        self._build_ui()
        self._register_drop_targets()
        self._apply_state(AppState.WAITING)
        self._poll_worker_events()

    def _configure_root(self) -> None:
        self.root.title(UI_TEXT["app_title"])
        self.root.geometry("640x420")
        self.root.minsize(560, 360)
        self.root.configure(bg=COLORS["window_bg"])
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        icon_path = ICON_PATH
        if icon_path.exists():
            try:
                self.root.iconbitmap(str(icon_path))
            except tk.TclError:
                pass

    def _build_ui(self) -> None:
        self.content = tk.Frame(self.root, bg=COLORS["window_bg"])
        self.content.grid(row=0, column=0, sticky="nsew", padx=32, pady=(32, 16))
        self.content.grid_rowconfigure(0, weight=1)
        self.content.grid_columnconfigure(0, weight=1)

        self.drop_area = tk.Frame(
            self.content,
            bg=COLORS["drop_bg"],
            highlightthickness=2,
            highlightbackground=COLORS["drop_border"],
            highlightcolor=COLORS["drop_border"],
            bd=0,
        )
        self.drop_area.grid(row=0, column=0, sticky="nsew")
        self.drop_area.grid_rowconfigure(0, weight=1)
        self.drop_area.grid_rowconfigure(4, weight=1)
        self.drop_area.grid_columnconfigure(0, weight=1)

        self.main_label = tk.Label(
            self.drop_area,
            textvariable=self.main_var,
            bg=COLORS["drop_bg"],
            fg=COLORS["main_text"],
            font=("Yu Gothic UI", 22, "bold"),
            wraplength=500,
            justify="center",
        )
        self.main_label.grid(row=1, column=0, padx=28, pady=(12, 4), sticky="ew")

        self.sub_label = tk.Label(
            self.drop_area,
            textvariable=self.sub_var,
            bg=COLORS["drop_bg"],
            fg=COLORS["sub_text"],
            font=("Yu Gothic UI", 11),
            wraplength=500,
            justify="center",
        )
        self.sub_label.grid(row=2, column=0, padx=28, pady=(0, 12), sticky="ew")

        self.choice_frame = tk.Frame(self.drop_area, bg=COLORS["drop_bg"], height=48)
        self.choice_frame.grid(row=3, column=0, padx=28, pady=(8, 0))
        self.choice_frame.grid_propagate(False)

        self.bottom_bar = tk.Frame(self.root, bg=COLORS["window_bg"])
        self.bottom_bar.grid(row=1, column=0, sticky="ew", padx=24, pady=(0, 14))
        self.bottom_bar.grid_columnconfigure(0, weight=1)
        self.bottom_bar.grid_columnconfigure(1, weight=1)

        self.status_label = tk.Label(
            self.bottom_bar,
            textvariable=self.status_var,
            bg=COLORS["window_bg"],
            fg=COLORS["status_text"],
            font=("Yu Gothic UI", 9),
            anchor="w",
        )
        self.status_label.grid(row=0, column=0, sticky="w")

        self.footer_label = tk.Label(
            self.bottom_bar,
            text=UI_TEXT["footer"],
            bg=COLORS["window_bg"],
            fg=COLORS["footer_text"],
            font=("Yu Gothic UI", 8),
            anchor="e",
        )
        self.footer_label.grid(row=0, column=1, sticky="e")

    def _register_drop_targets(self) -> None:
        for widget in (self.root, self.drop_area, self.main_label, self.sub_label):
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<DropEnter>>", self._on_drop_enter)
            widget.dnd_bind("<<DropLeave>>", self._on_drop_leave)
            widget.dnd_bind("<<Drop>>", self._on_drop)

    def _on_drop_enter(self, _event: object) -> str:
        if self._can_accept_drop():
            self._set_drop_highlight(True)
            self.main_var.set(UI_TEXT["drop_over_main"])
        return "copy"

    def _on_drop_leave(self, _event: object) -> str:
        if self._can_accept_drop():
            self._set_drop_highlight(False)
            self._refresh_state_text()
        return "copy"

    def _on_drop(self, event: object) -> str:
        if not self._can_accept_drop():
            return "copy"

        self._cancel_reset()
        self._hide_choices()
        self._set_drop_highlight(False)
        self.current_inputs = ()
        paths = self._parse_drop_data(getattr(event, "data", ""))

        self._apply_state(
            AppState.CHECKING,
            main_text=UI_TEXT["checking_main"],
            sub_text=UI_TEXT["checking_sub"],
        )
        self.root.after(160, lambda: self._judge_paths(paths))
        return "copy"

    def _can_accept_drop(self) -> bool:
        return self.state in {AppState.WAITING, AppState.DONE, AppState.ERROR}

    def _parse_drop_data(self, data: str) -> list[Path]:
        if not data:
            return []
        try:
            raw_items = self.root.tk.splitlist(data)
        except tk.TclError:
            raw_items = (data,)

        paths: list[Path] = []
        seen: set[str] = set()
        for raw_item in raw_items:
            cleaned = str(raw_item).strip()
            if cleaned.startswith("file:"):
                parsed = urlparse(cleaned)
                cleaned = unquote(parsed.path)
                if parsed.netloc:
                    cleaned = f"//{parsed.netloc}{cleaned}"
                if len(cleaned) >= 3 and cleaned[0] == "/" and cleaned[2] == ":":
                    cleaned = cleaned[1:]
            if not cleaned:
                continue
            path = Path(cleaned)
            key = str(path).casefold()
            if key not in seen:
                seen.add(key)
                paths.append(path)
        return paths

    def _judge_paths(self, paths: list[Path]) -> None:
        self._apply_state(AppState.JUDGING, main_text=UI_TEXT["checking_main"], sub_text="")

        if not paths:
            self._show_error("error_read_main", "error_read_sub")
            return

        existing_paths = [path for path in paths if path.exists()]
        if len(existing_paths) != len(paths):
            self._show_error("error_read_main", "error_read_sub")
            return

        if any(path.is_dir() for path in existing_paths):
            self._show_error("error_folder_main", "error_folder_sub")
            return

        files = [path for path in existing_paths if path.is_file()]
        if len(files) != len(existing_paths):
            self._show_error("error_read_main", "error_read_sub")
            return

        if any(path.suffix.lower() == SHORTCUT_EXTENSION for path in files):
            self._show_error("error_mixed_unknown_main", "error_mixed_unknown_sub")
            return

        pdf_files = [path for path in files if is_pdf(path)]
        image_files = [path for path in files if is_supported_image(path)]
        supported_count = len(pdf_files) + len(image_files)
        unknown_count = len(files) - supported_count

        if supported_count == 0:
            self._show_error("error_unknown_main", "error_unknown_sub")
            return

        if unknown_count:
            self._show_error("error_mixed_unknown_main", "error_mixed_unknown_sub")
            return

        decision = self._make_decision(pdf_files, image_files)
        self.current_inputs = decision.inputs
        self._show_decision(decision)

    def _make_decision(self, pdf_files: list[Path], image_files: list[Path]) -> Decision:
        if pdf_files and image_files:
            return Decision(
                main_text_key="decision_mixed",
                inputs=tuple(pdf_files + image_files),
                choices=(
                    Choice("choice_all_pdf", "mixed_to_pdf", planned_only=True),
                    Choice("choice_all_images", "mixed_to_images", planned_only=True),
                ),
            )

        if pdf_files:
            if len(pdf_files) > 1:
                return Decision(
                    main_text_key="decision_pdf_many",
                    inputs=tuple(pdf_files),
                    auto_action="pdf_merge",
                )
            return Decision(
                main_text_key="decision_pdf_one",
                inputs=tuple(pdf_files),
                choices=(
                    Choice("choice_split_all", "pdf_split_one"),
                    Choice("choice_split_select", "pdf_split_select"),
                ),
            )

        if len(image_files) > 1:
            return Decision(
                main_text_key="decision_image_many",
                inputs=tuple(image_files),
                choices=(
                    Choice("choice_to_pdf", "image_to_pdf"),
                    Choice("choice_resize", "image_resize"),
                ),
            )

        return Decision(
            main_text_key="decision_image_one",
            inputs=tuple(image_files),
            choices=(
                Choice("choice_to_pdf", "image_to_pdf"),
                Choice("choice_resize", "image_resize"),
            ),
        )

    def _show_decision(self, decision: Decision) -> None:
        main_text = UI_TEXT[decision.main_text_key]
        if decision.auto_action:
            self._apply_state(AppState.JUDGING, main_text=main_text, sub_text="")
            self.root.after(700, lambda: self._run_action(decision.auto_action, decision.inputs))
            return

        self._apply_state(
            AppState.CONFIRMING,
            main_text=main_text,
            sub_text=UI_TEXT["confirm_sub"],
        )
        self._show_choices(decision.choices)

    def _show_choices(self, choices: tuple[Choice, ...]) -> None:
        self._hide_choices()
        for choice in choices[:2]:
            button = tk.Button(
                self.choice_frame,
                text=UI_TEXT[choice.label_key],
                command=lambda selected=choice: self._handle_choice(selected),
                bg=COLORS["button_bg"],
                activebackground=COLORS["button_active"],
                fg=COLORS["main_text"],
                relief="solid",
                bd=1,
                padx=18,
                pady=8,
                font=("Yu Gothic UI", 10, "bold"),
                cursor="hand2",
                takefocus=True,
            )
            button.pack(side="left", padx=6)
            self.choice_buttons.append(button)

    def _hide_choices(self) -> None:
        for button in self.choice_buttons:
            button.destroy()
        self.choice_buttons.clear()

    def _handle_choice(self, choice: Choice) -> None:
        if self.state != AppState.CONFIRMING:
            return
        self._hide_choices()
        if choice.planned_only:
            self._show_done("planned_main", "ready_again")
            return
        self._run_action(choice.action, self.current_inputs)

    def _run_action(self, action: str, inputs: tuple[Path, ...]) -> None:
        if not inputs:
            self._show_error("error_read_main", "error_read_sub")
            return

        self._apply_state(
            AppState.RUNNING,
            main_text=UI_TEXT["running_main"],
            sub_text=UI_TEXT["running_sub"],
        )
        worker = threading.Thread(
            target=self._run_dake_worker,
            args=(action, inputs, None),
            daemon=True,
        )
        worker.start()

    def _run_dake_worker(
        self,
        action: str,
        inputs: tuple[Path, ...],
        output_folder: Path | None,
    ) -> None:
        exists, exe_path, reason = check_dake_app(action)
        if not exists or exe_path is None:
            detail = reason
            if exe_path is not None:
                detail = UI_TEXT["error_dake_missing_detail"].format(path=exe_path)
            self.event_queue.put(("error_detail", ("error_dake_missing_main", detail, 12000)))
            return

        command = self._build_dake_command(exe_path, inputs, output_folder)
        try:
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
            completed = subprocess.run(
                command,
                cwd=str(exe_path.parent),
                check=False,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=creationflags,
            )
        except OSError:
            self.event_queue.put(("error", ("error_dake_launch_main", "error_dake_launch_sub")))
            return

        if completed.returncode == 0:
            self.event_queue.put(("done", ("done_main", "ready_again")))
        else:
            self.event_queue.put(("error", ("error_dake_launch_main", "error_dake_launch_sub")))

    @staticmethod
    def _build_dake_command(
        exe_path: Path,
        inputs: tuple[Path, ...],
        output_folder: Path | None,
    ) -> list[str]:
        command = [str(exe_path), "--inputs"]
        command.extend(str(path) for path in inputs)
        if output_folder is not None:
            command.extend(["--output", str(output_folder)])
        command.append("--from-shimarisu")
        return command

    def _poll_worker_events(self) -> None:
        try:
            while True:
                event_name, payload = self.event_queue.get_nowait()
                if event_name == "done":
                    main_key, sub_key = payload  # type: ignore[misc]
                    self._show_done(str(main_key), str(sub_key))
                elif event_name == "error":
                    main_key, sub_key = payload  # type: ignore[misc]
                    self._show_error(str(main_key), str(sub_key))
                elif event_name == "error_detail":
                    main_key, sub_text, delay_ms = payload  # type: ignore[misc]
                    self._show_error_message(str(main_key), str(sub_text), int(delay_ms))
        except queue.Empty:
            pass
        self.root.after(100, self._poll_worker_events)

    def _show_done(self, main_key: str, sub_key: str) -> None:
        self._apply_state(
            AppState.DONE,
            main_text=UI_TEXT[main_key],
            sub_text=UI_TEXT[sub_key],
        )
        self._schedule_reset(2600)

    def _show_error(self, main_key: str, sub_key: str) -> None:
        self._show_error_message(UI_TEXT[main_key], UI_TEXT[sub_key], 3600)

    def _show_error_message(self, main_text: str, sub_text: str, delay_ms: int) -> None:
        self._hide_choices()
        self._apply_state(
            AppState.ERROR,
            main_text=main_text,
            sub_text=sub_text,
        )
        self._schedule_reset(delay_ms)

    def _apply_state(
        self,
        state: AppState,
        main_text: str | None = None,
        sub_text: str | None = None,
    ) -> None:
        self.state = state
        if main_text is not None:
            self.main_var.set(main_text)
        if sub_text is not None:
            self.sub_var.set(sub_text)

        if state == AppState.WAITING:
            self._hide_choices()
            self.main_var.set(UI_TEXT["drop_main"])
            self.sub_var.set(UI_TEXT["drop_sub"])
            self.status_var.set(UI_TEXT["status_waiting"])
        elif state == AppState.CONFIRMING:
            self.status_var.set(UI_TEXT["status_confirming"])
        elif state == AppState.DONE:
            self.status_var.set(UI_TEXT["status_done"])
        elif state == AppState.ERROR:
            self.status_var.set(UI_TEXT["status_error"])

        self._set_drop_highlight(False)
        self._update_status_ticker()

    def _refresh_state_text(self) -> None:
        if self.state == AppState.WAITING:
            self.main_var.set(UI_TEXT["drop_main"])
            self.sub_var.set(UI_TEXT["drop_sub"])
        elif self.state == AppState.DONE:
            self.main_var.set(UI_TEXT["done_main"])
            self.sub_var.set(UI_TEXT["ready_again"])

    def _update_status_ticker(self) -> None:
        if self.state in {AppState.CHECKING, AppState.JUDGING, AppState.RUNNING}:
            if self.status_after_id is None:
                self.dot_index = 0
                self._tick_status()
            return

        if self.status_after_id is not None:
            self.root.after_cancel(self.status_after_id)
            self.status_after_id = None

    def _tick_status(self) -> None:
        base_key_by_state = {
            AppState.CHECKING: "status_loading",
            AppState.JUDGING: "status_judging",
            AppState.RUNNING: "status_running",
        }
        base_key = base_key_by_state.get(self.state)
        if base_key is None:
            self.status_after_id = None
            return
        dots = "." * ((self.dot_index % 3) + 1)
        self.status_var.set(f"{UI_TEXT[base_key]}{dots}")
        self.dot_index += 1
        self.status_after_id = self.root.after(420, self._tick_status)

    def _set_drop_highlight(self, is_active: bool) -> None:
        bg = COLORS["drop_bg_active"] if is_active else COLORS["drop_bg"]
        border = COLORS["drop_border_active"] if is_active else COLORS["drop_border"]
        self.drop_area.configure(bg=bg, highlightbackground=border, highlightcolor=border)
        self.main_label.configure(bg=bg)
        self.sub_label.configure(bg=bg)
        self.choice_frame.configure(bg=bg)
        for button in self.choice_buttons:
            button.configure(bg=COLORS["button_bg"])

    def _schedule_reset(self, delay_ms: int) -> None:
        self._cancel_reset()
        self.reset_after_id = self.root.after(delay_ms, self._reset_to_waiting)

    def _cancel_reset(self) -> None:
        if self.reset_after_id is not None:
            self.root.after_cancel(self.reset_after_id)
            self.reset_after_id = None

    def _reset_to_waiting(self) -> None:
        self.reset_after_id = None
        self.current_inputs = ()
        self._apply_state(AppState.WAITING)


def main() -> None:
    if TkinterDnD is None:
        fallback_root = tk.Tk()
        fallback_root.withdraw()
        messagebox.showerror(UI_TEXT["app_title"], UI_TEXT["error_dnd_missing"])
        fallback_root.destroy()
        return

    root = TkinterDnD.Tk()
    ShimarisuKunApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
