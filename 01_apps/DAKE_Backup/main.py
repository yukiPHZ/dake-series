# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import queue
import tempfile
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox

from core.backup_engine import BackupError, DiffResult, execute_backup, scan_diff
from core.logger import BackupLogger, display_timestamp
from core.settings import AppSettings, ensure_data_files, load_settings, save_settings, settings_path


APP_NAME = "DAKE_Backup"
DISPLAY_NAME = "DAKE Backup"
SUBTITLE = "消さない。静かに残す。"

COLORS = {
    "bg": "#05070D",
    "panel": "#0B111A",
    "panel_soft": "#101826",
    "field": "#060B12",
    "line": "#27364B",
    "text": "#EEF4F8",
    "muted": "#9AA8B6",
    "weak": "#627184",
    "accent": "#B99A5B",
    "accent_hover": "#D0AF70",
    "blue": "#6AA9C9",
    "success": "#7ED7B5",
    "danger": "#E27D73",
}


def now_for_settings() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


class BackupApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(DISPLAY_NAME)
        self.root.geometry("920x700")
        self.root.minsize(820, 620)
        self.root.configure(bg=COLORS["bg"])

        self.settings = load_settings()
        self.worker_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.current_diff: DiffResult | None = None
        self.busy = False

        self.source_var = tk.StringVar(value=self.settings.source_folder)
        self.destination_var = tk.StringVar(value=self.settings.destination_folder)
        self.last_saved_var = tk.StringVar(
            value=self.settings.last_saved_at or "最終保存日時: まだありません"
        )
        if self.settings.last_saved_at:
            self.last_saved_var.set(f"最終保存日時: {self.settings.last_saved_at}")

        self.summary_vars = {
            "追加": tk.StringVar(value="0"),
            "更新": tk.StringVar(value="0"),
            "退避": tk.StringVar(value="0"),
            "削除予定": tk.StringVar(value="0"),
        }

        self._build_ui()
        self._append_log("READY: Driveは記憶庫ではなく避難先。正本はローカルにある。")
        self.root.after(120, self._poll_worker_queue)

    def _build_ui(self) -> None:
        outer = tk.Frame(self.root, bg=COLORS["bg"])
        outer.pack(fill="both", expand=True, padx=28, pady=24)

        header = tk.Frame(outer, bg=COLORS["bg"])
        header.pack(fill="x")
        tk.Label(
            header,
            text=DISPLAY_NAME,
            bg=COLORS["bg"],
            fg=COLORS["text"],
            font=("Segoe UI", 28, "bold"),
        ).pack(anchor="w")
        tk.Label(
            header,
            text=SUBTITLE,
            bg=COLORS["bg"],
            fg=COLORS["muted"],
            font=("Yu Gothic UI", 12),
        ).pack(anchor="w", pady=(2, 0))

        paths_panel = self._panel(outer)
        paths_panel.pack(fill="x", pady=(24, 16))
        self._folder_row(paths_panel, "正本フォルダ", self.source_var, self._choose_source, 0)
        self._folder_row(paths_panel, "避難先フォルダ", self.destination_var, self._choose_destination, 1)

        summary_panel = self._panel(outer)
        summary_panel.pack(fill="x", pady=(0, 16))
        tk.Label(
            summary_panel,
            text="差分サマリー",
            bg=COLORS["panel"],
            fg=COLORS["text"],
            font=("Yu Gothic UI", 12, "bold"),
        ).pack(anchor="w", padx=18, pady=(16, 8))

        summary_grid = tk.Frame(summary_panel, bg=COLORS["panel"])
        summary_grid.pack(fill="x", padx=14, pady=(0, 16))
        for index, label in enumerate(["追加", "更新", "退避", "削除予定"]):
            card = tk.Frame(
                summary_grid,
                bg=COLORS["panel_soft"],
                highlightbackground=COLORS["line"],
                highlightthickness=1,
            )
            card.grid(row=0, column=index, sticky="ew", padx=5)
            summary_grid.columnconfigure(index, weight=1)
            color = COLORS["success"] if label == "追加" else COLORS["blue"]
            if label == "削除予定":
                color = COLORS["danger"]
            tk.Label(
                card,
                text=label,
                bg=COLORS["panel_soft"],
                fg=COLORS["muted"],
                font=("Yu Gothic UI", 10),
            ).pack(anchor="w", padx=14, pady=(12, 0))
            tk.Label(
                card,
                textvariable=self.summary_vars[label],
                bg=COLORS["panel_soft"],
                fg=color,
                font=("Segoe UI", 26, "bold"),
            ).pack(anchor="w", padx=14, pady=(0, 10))

        actions = tk.Frame(outer, bg=COLORS["bg"])
        actions.pack(fill="x", pady=(0, 16))
        self.diff_button = self._button(actions, "差分を見る", self.show_diff, COLORS["panel_soft"])
        self.diff_button.pack(side="left", padx=(0, 10))
        self.keep_button = self._button(actions, "残す", self.keep_backup, COLORS["accent"])
        self.keep_button.pack(side="left")
        tk.Label(
            actions,
            textvariable=self.last_saved_var,
            bg=COLORS["bg"],
            fg=COLORS["weak"],
            font=("Yu Gothic UI", 10),
        ).pack(side="right")

        log_panel = self._panel(outer)
        log_panel.pack(fill="both", expand=True)
        tk.Label(
            log_panel,
            text="Log",
            bg=COLORS["panel"],
            fg=COLORS["text"],
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w", padx=18, pady=(14, 6))
        self.log_text = tk.Text(
            log_panel,
            height=10,
            bg=COLORS["field"],
            fg=COLORS["muted"],
            insertbackground=COLORS["text"],
            relief="flat",
            borderwidth=0,
            font=("Consolas", 10),
            wrap="word",
        )
        self.log_text.pack(fill="both", expand=True, padx=18, pady=(0, 18))

    def _panel(self, parent: tk.Widget) -> tk.Frame:
        return tk.Frame(
            parent,
            bg=COLORS["panel"],
            highlightbackground=COLORS["line"],
            highlightthickness=1,
        )

    def _folder_row(
        self,
        parent: tk.Widget,
        label: str,
        variable: tk.StringVar,
        command: callable,
        row: int,
    ) -> None:
        parent.grid_columnconfigure(1, weight=1)
        tk.Label(
            parent,
            text=label,
            bg=COLORS["panel"],
            fg=COLORS["muted"],
            font=("Yu Gothic UI", 10, "bold"),
        ).grid(row=row, column=0, sticky="w", padx=(18, 12), pady=12)
        entry = tk.Entry(
            parent,
            textvariable=variable,
            bg=COLORS["field"],
            fg=COLORS["text"],
            insertbackground=COLORS["text"],
            relief="flat",
            font=("Yu Gothic UI", 10),
        )
        entry.grid(row=row, column=1, sticky="ew", pady=12, ipady=8)
        self._button(parent, "選ぶ", command, COLORS["panel_soft"]).grid(
            row=row,
            column=2,
            sticky="e",
            padx=18,
            pady=12,
        )

    def _button(self, parent: tk.Widget, text: str, command: callable, bg: str) -> tk.Button:
        hover = COLORS["accent_hover"] if bg == COLORS["accent"] else COLORS["line"]
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg=COLORS["text"],
            activebackground=hover,
            activeforeground=COLORS["text"],
            relief="flat",
            borderwidth=0,
            padx=22,
            pady=10,
            cursor="hand2",
            font=("Yu Gothic UI", 10, "bold"),
        )

    def _choose_source(self) -> None:
        folder = filedialog.askdirectory(title="正本フォルダを選択")
        if folder:
            self.source_var.set(folder)
            self._save_current_paths()

    def _choose_destination(self) -> None:
        folder = filedialog.askdirectory(title="避難先フォルダを選択")
        if folder:
            self.destination_var.set(folder)
            self._save_current_paths()

    def _save_current_paths(self) -> None:
        self.settings.source_folder = self.source_var.get().strip()
        self.settings.destination_folder = self.destination_var.get().strip()
        save_settings(self.settings)

    def _set_busy(self, busy: bool) -> None:
        self.busy = busy
        state = "disabled" if busy else "normal"
        self.diff_button.configure(state=state)
        self.keep_button.configure(state=state)

    def _append_log(self, message: str) -> None:
        self.log_text.insert("end", f"[{display_timestamp()}] {message}\n")
        self.log_text.see("end")

    def _show_error(self, message: str) -> None:
        self._append_log(f"ERROR: {message}")
        messagebox.showerror(DISPLAY_NAME, message)

    def _update_summary(self, diff: DiffResult | None) -> None:
        summary = diff.summary() if diff else {"added": 0, "updated": 0, "archive": 0, "delete": 0}
        self.summary_vars["追加"].set(str(summary.get("added", 0)))
        self.summary_vars["更新"].set(str(summary.get("updated", 0)))
        self.summary_vars["退避"].set(str(summary.get("archive", 0)))
        self.summary_vars["削除予定"].set("0")

    def show_diff(self) -> None:
        if self.busy:
            return
        self._save_current_paths()
        self._set_busy(True)
        self._append_log("差分を見ています。")
        source = self.source_var.get().strip()
        destination = self.destination_var.get().strip()
        threading.Thread(target=self._diff_worker, args=(source, destination), daemon=True).start()

    def keep_backup(self) -> None:
        if self.busy:
            return
        self._save_current_paths()
        self._set_busy(True)
        self._append_log("残します。")
        source = self.source_var.get().strip()
        destination = self.destination_var.get().strip()
        threading.Thread(target=self._backup_worker, args=(source, destination), daemon=True).start()

    def _diff_worker(self, source: str, destination: str) -> None:
        try:
            diff = scan_diff(source, destination)
        except Exception as exc:
            self.worker_queue.put(("error", str(exc)))
            return
        self.worker_queue.put(("diff", diff))

    def _backup_worker(self, source: str, destination: str) -> None:
        try:
            diff = scan_diff(source, destination)
            logger = BackupLogger()
            result = execute_backup(source, destination, diff=diff, logger=logger, timestamp=logger.timestamp)
        except Exception as exc:
            self.worker_queue.put(("error", str(exc)))
            return
        self.worker_queue.put(("backup", result))

    def _poll_worker_queue(self) -> None:
        try:
            while True:
                kind, payload = self.worker_queue.get_nowait()
                if kind == "diff":
                    diff = payload
                    if isinstance(diff, DiffResult):
                        self.current_diff = diff
                        self._update_summary(diff)
                        summary = diff.summary()
                        self._append_log(
                            "差分: "
                            f"追加 {summary['added']} / 更新 {summary['updated']} / "
                            f"退避 {summary['archive']} / 削除予定 0"
                        )
                        if summary["preserved_destination_only"]:
                            self._append_log(
                                f"避難先だけにあるファイル {summary['preserved_destination_only']} 件は残します。"
                            )
                    self._set_busy(False)
                elif kind == "backup":
                    result = payload
                    self.current_diff = result.diff
                    self._update_summary(result.diff)
                    self.settings.last_saved_at = now_for_settings()
                    save_settings(self.settings)
                    self.last_saved_var.set(f"最終保存日時: {self.settings.last_saved_at}")
                    self._append_log(
                        f"保存完了: コピー {len(result.copied)} / 退避 {len(result.archived)} / 削除予定 0"
                    )
                    self._append_log(f"ログ: {result.log_path}")
                    self._set_busy(False)
                elif kind == "error":
                    self._set_busy(False)
                    self._show_error(str(payload))
        except queue.Empty:
            pass
        self.root.after(120, self._poll_worker_queue)


def run_launch_check() -> int:
    settings = load_settings()
    save_settings(settings)
    if not settings_path().exists():
        raise RuntimeError("settings.json was not created")
    print("LAUNCH CHECK OK")
    return 0


def run_smoke_test() -> int:
    ensure_data_files()
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        source = root / "source"
        destination = root / "destination"
        source.mkdir()
        (source / "note.txt").write_text("first", encoding="utf-8")

        diff1 = scan_diff(str(source), str(destination))
        if diff1.summary()["added"] != 1 or diff1.summary()["delete"] != 0:
            raise RuntimeError("add diff check failed")
        result1 = execute_backup(str(source), str(destination), diff=diff1)
        if not (destination / "note.txt").exists() or not result1.log_path:
            raise RuntimeError("initial copy failed")

        (source / "note.txt").write_text("second", encoding="utf-8")
        diff2 = scan_diff(str(source), str(destination))
        if diff2.summary()["updated"] != 1 or diff2.summary()["archive"] != 1:
            raise RuntimeError("update diff check failed")
        execute_backup(str(source), str(destination), diff=diff2)
        archived_notes = list((destination / "backup_archive").glob("*/note.txt"))
        if not archived_notes or archived_notes[0].read_text(encoding="utf-8") != "first":
            raise RuntimeError("archive check failed")
        if (destination / "note.txt").read_text(encoding="utf-8") != "second":
            raise RuntimeError("update copy failed")

        (source / "note.txt").unlink()
        diff3 = scan_diff(str(source), str(destination))
        execute_backup(str(source), str(destination), diff=diff3)
        if not (destination / "note.txt").exists():
            raise RuntimeError("destination file was deleted")
        if diff3.summary()["delete"] != 0:
            raise RuntimeError("delete propagation was planned")

    settings = load_settings()
    save_settings(settings)
    if not settings_path().exists():
        raise RuntimeError("settings save check failed")
    if not any(settings_path().parent.joinpath("logs").glob("backup_*.log")):
        raise RuntimeError("log output check failed")
    print("SMOKE TEST OK")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=DISPLAY_NAME)
    parser.add_argument("--launch-check", action="store_true")
    parser.add_argument("--smoke-test", action="store_true")
    args = parser.parse_args()

    if args.launch_check:
        return run_launch_check()
    if args.smoke_test:
        return run_smoke_test()

    ensure_data_files()
    root = tk.Tk()
    BackupApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

