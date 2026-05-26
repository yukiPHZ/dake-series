# -*- coding: utf-8 -*-
from __future__ import annotations

import csv
import html
import json
import logging
import queue
import re
import sys
import threading
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, ttk


APP_KEY = "DAKE_Mail_Draft"
APP_NAME = "Dakeメール下書き"
WINDOW_TITLE = APP_NAME
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "display_subtitle": "CSVからOutlook下書きだけ作る",
    "main_description": "メールは送信しません。生成後はOutlookで確認してください。",
    "default_subject_template": "{会社名} {名前}様へのご連絡",
    "default_body_template": "{会社名}\n{名前} 様\n\nお世話になっております。\n\n以下の内容をご確認ください。\n\nよろしくお願いいたします。",
    "status_select_csv": "CSVを選択してください。",
    "report_pending": "レポートは生成後に表示されます。",
    "button_generate": "Outlook下書きを生成",
    "button_cancel": "キャンセル",
    "label_csv": "CSV選択",
    "button_select": "選択",
    "label_company_column": "会社名列",
    "label_name_column": "名前列",
    "label_email_column": "メール列",
    "label_subject_template": "件名テンプレート",
    "label_body_template": "本文テンプレート",
    "label_attachments": "添付ファイル",
    "button_add": "追加",
    "button_remove": "削除",
    "button_clear": "全削除",
    "label_limit": "生成件数",
    "limit_all": "全件",
    "option_keep_displayed": "生成後もOutlook下書き画面を表示したままにする",
    "preview_title": "プレビュー",
    "report_title": "レポート",
    "safety_note": "このアプリはOutlook下書き作成補助です。作成後はOutlook上で宛先・本文・添付を必ず確認してください。",
    "dialog_csv_title": "CSV名簿を選択",
    "dialog_attachment_title": "添付ファイルを選択",
    "status_csv_load_error": "CSVを読み込めませんでした: {error}",
    "dialog_csv_load_error": "CSVを読み込めませんでした。\n\n{error}",
    "status_csv_loaded": "CSVを読み込みました。{count}件 / 文字コード {encoding}",
    "preview_select_csv": "CSVを選択すると、先頭1件のプレビューを表示します。",
    "preview_select_mapping": "列マッピングを選択してください。",
    "preview_no_targets": "生成できるメールアドレスが見つかりません。スキップ候補: {count}件",
    "preview_format": "宛先: {email}\n件名: {subject}\n\n本文:\n{body}",
    "dialog_select_csv_error": "CSVファイルを選択してください。",
    "dialog_mapping_error": "会社名列・名前列・メールアドレス列を選択してください。",
    "dialog_missing_attachments": "存在しない添付ファイルがあります。\n\n{paths}",
    "dialog_no_targets": "生成対象のメールアドレスがありません。",
    "confirm_generation": "このアプリはOutlookの下書きを作成します。メールは送信しません。生成後はOutlookで内容・宛先・添付を確認してください。\n\n生成予定: {target_count}件\n事前スキップ: {skipped_count}件",
    "status_generating": "Outlook下書きを生成しています。",
    "report_generating": "生成中です。途中エラーは行ごとにレポートへ記録します。",
    "status_cancel_requested": "キャンセル要求を受け付けました。処理中の1件が終わるまで待ちます。",
    "status_progress": "生成中: {current}/{total}件 / 作成 {drafted}件 / エラー {errors}件",
    "status_completed": "完了しました。",
    "status_cancelled": "キャンセルしました。",
    "status_complete": "{prefix} 作成 {drafted}件 / スキップ {skipped}件 / エラー {errors}件",
    "report_complete": "{prefix}\n作成: {drafted}件\nスキップ: {skipped}件\nエラー: {errors}件\n\nレポート:\n{report_path}",
    "report_empty_email": "メールアドレスが空です。",
    "report_invalid_email": "メール形式が不正です。",
    "report_duplicate_email": "重複メールアドレスのため1件目のみ作成します。",
    "report_drafted": "Outlook下書きを作成しました。",
    "error_pywin32_missing": "pywin32 が見つかりません。requirements.txt を確認してください。",
    "error_outlook_connect": "Microsoft Outlook Classic に接続できませんでした。",
    "error_outlook_not_initialized": "Outlook が初期化されていません。",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_caption": "止まらない、迷わない、すぐ終わる。",
    "footer_link_1": "戸建買取査定",
    "footer_link_2": "Instagram",
    "footer_separator": "｜",
    "footer_copyright": COPYRIGHT,
}
DISPLAY_SUBTITLE = UI_TEXT["display_subtitle"]

COMPANY_CANDIDATES = ("会社名", "company", "company_name", "organization")
NAME_CANDIDATES = ("名前", "氏名", "name", "person_name")
EMAIL_CANDIDATES = ("mail 1", "mail", "email", "メール", "メールアドレス", "email_address")
PLACEHOLDERS = ("会社名", "名前", "メール", "company", "name", "email")
LIMIT_CHOICES = ("5", "10", "20", "50", "100", UI_TEXT["limit_all"])
DEFAULT_LIMIT = "50"
REPORT_FIELDS = (
    "row_number",
    "company",
    "name",
    "email",
    "status",
    "message",
    "subject",
    "created_at",
)
EMAIL_PATTERN = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
QUEUE_POLL_INTERVAL_MS = 100
FOOTER_WIDE_THRESHOLD = 900

COLORS = {
    "background": "#F5F7FA",
    "surface": "#FFFFFF",
    "surface_alt": "#F9FAFB",
    "text": "#1D2939",
    "muted": "#667085",
    "border": "#D0D5DD",
    "accent": "#2563EB",
    "accent_hover": "#1D4ED8",
    "success": "#027A48",
    "warning": "#B54708",
    "error": "#B42318",
}

DEFAULT_SUBJECT_TEMPLATE = UI_TEXT["default_subject_template"]
DEFAULT_BODY_TEMPLATE = UI_TEXT["default_body_template"]


@dataclass(frozen=True)
class CsvRecord:
    row_number: int
    values: dict[str, str]


@dataclass(frozen=True)
class DraftTarget:
    row_number: int
    company: str
    name: str
    email: str


@dataclass(frozen=True)
class ReportRecord:
    row_number: int
    company: str
    name: str
    email: str
    status: str
    message: str
    subject: str
    created_at: str

    def as_dict(self) -> dict[str, str]:
        return {
            "row_number": str(self.row_number),
            "company": self.company,
            "name": self.name,
            "email": self.email,
            "status": self.status,
            "message": self.message,
            "subject": self.subject,
            "created_at": self.created_at,
        }


def get_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APP_DIR = get_app_dir()
DATA_DIR = APP_DIR / "data"
LOG_DIR = APP_DIR / "logs"
OUTPUT_DIR = APP_DIR / "output"
ATTACHMENTS_DIR = APP_DIR / "attachments"
TEMPLATES_DIR = APP_DIR / "templates"
SETTINGS_PATH = DATA_DIR / "settings.json"
ICON_PATH = APP_DIR.parent.parent / "02_assets" / "dake_icon.ico"


def ensure_app_directories() -> None:
    for directory in (DATA_DIR, LOG_DIR, OUTPUT_DIR, ATTACHMENTS_DIR, TEMPLATES_DIR):
        directory.mkdir(parents=True, exist_ok=True)


ensure_app_directories()
logging.basicConfig(
    filename=LOG_DIR / "app.log",
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    encoding="utf-8",
)


def normalize_header(value: str) -> str:
    return re.sub(r"\s+", " ", value.strip()).lower()


def normalize_cell(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def read_csv_records(csv_path: Path) -> tuple[list[CsvRecord], list[str], str]:
    last_error: Exception | None = None
    for encoding in ("utf-8-sig", "cp932"):
        try:
            with csv_path.open("r", encoding=encoding, newline="") as handle:
                reader = csv.DictReader(handle)
                fieldnames = [str(name) for name in (reader.fieldnames or []) if name is not None]
                records = [
                    CsvRecord(
                        row_number=index,
                        values={
                            str(key): normalize_cell(value)
                            for key, value in row.items()
                            if key is not None
                        },
                    )
                    for index, row in enumerate(reader, start=2)
                ]
            return records, fieldnames, encoding
        except UnicodeDecodeError as exc:
            last_error = exc
            continue
    if last_error is not None:
        raise last_error
    raise ValueError("CSVを読み込めませんでした。")


def infer_column(fieldnames: list[str], candidates: tuple[str, ...]) -> str:
    normalized_to_original = {normalize_header(name): name for name in fieldnames}
    for candidate in candidates:
        matched = normalized_to_original.get(normalize_header(candidate))
        if matched:
            return matched
    return fieldnames[0] if fieldnames else ""


def infer_columns(fieldnames: list[str]) -> dict[str, str]:
    return {
        "company": infer_column(fieldnames, COMPANY_CANDIDATES),
        "name": infer_column(fieldnames, NAME_CANDIDATES),
        "email": infer_column(fieldnames, EMAIL_CANDIDATES),
    }


def is_valid_email(email_address: str) -> bool:
    return bool(EMAIL_PATTERN.match(email_address.strip()))


def build_placeholder_values(target: DraftTarget) -> dict[str, str]:
    return {
        "会社名": target.company,
        "名前": target.name,
        "メール": target.email,
        "company": target.company,
        "name": target.name,
        "email": target.email,
    }


def render_template(template: str, target: DraftTarget) -> str:
    rendered = template
    values = build_placeholder_values(target)
    for placeholder in PLACEHOLDERS:
        rendered = rendered.replace("{" + placeholder + "}", values[placeholder])
    return rendered


def body_text_to_html(body_text: str) -> str:
    normalized = body_text.replace("\r\n", "\n").replace("\r", "\n")
    escaped = html.escape(normalized, quote=True)
    return "<div>" + escaped.replace("\n", "<br>") + "</div>"


def make_report_record(
    target: DraftTarget,
    status: str,
    message: str,
    subject: str = "",
    created_at: str | None = None,
) -> ReportRecord:
    return ReportRecord(
        row_number=target.row_number,
        company=target.company,
        name=target.name,
        email=target.email,
        status=status,
        message=message,
        subject=subject,
        created_at=created_at or datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    )


def parse_limit(limit_label: str) -> int | None:
    if limit_label == UI_TEXT["limit_all"]:
        return None
    return int(limit_label)


def prepare_targets(
    records: list[CsvRecord],
    company_column: str,
    name_column: str,
    email_column: str,
    limit_label: str,
) -> tuple[list[DraftTarget], list[ReportRecord]]:
    limit = parse_limit(limit_label)
    targets: list[DraftTarget] = []
    reports: list[ReportRecord] = []
    seen_emails: set[str] = set()

    for record in records:
        company = normalize_cell(record.values.get(company_column))
        name = normalize_cell(record.values.get(name_column))
        email_address = normalize_cell(record.values.get(email_column))
        target = DraftTarget(record.row_number, company, name, email_address)

        if not email_address:
            reports.append(make_report_record(target, "skipped_empty_email", UI_TEXT["report_empty_email"]))
            continue

        if not is_valid_email(email_address):
            reports.append(make_report_record(target, "skipped_invalid_email", UI_TEXT["report_invalid_email"]))
            continue

        email_key = email_address.casefold()
        if email_key in seen_emails:
            reports.append(make_report_record(target, "skipped_duplicate", UI_TEXT["report_duplicate_email"]))
            continue

        if limit is not None and len(targets) >= limit:
            break

        seen_emails.add(email_key)
        targets.append(target)

    return targets, reports


def write_report(report_records: list[ReportRecord]) -> Path:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_path = LOG_DIR / f"draft_report_{timestamp}.csv"
    ordered_records = sorted(report_records, key=lambda item: item.row_number)
    with report_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=REPORT_FIELDS)
        writer.writeheader()
        for record in ordered_records:
            writer.writerow(record.as_dict())
    return report_path


def load_settings() -> dict[str, object]:
    if not SETTINGS_PATH.exists():
        return {}
    try:
        data = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        logging.warning("settings.json could not be loaded", exc_info=True)
        return {}
    return data if isinstance(data, dict) else {}


def save_settings(payload: dict[str, object]) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    SETTINGS_PATH.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


class OutlookDraftSession:
    def __init__(self, keep_displayed: bool) -> None:
        self.keep_displayed = keep_displayed
        self.pythoncom = None
        self.outlook = None

    def __enter__(self) -> "OutlookDraftSession":
        try:
            import pythoncom  # type: ignore
            import win32com.client  # type: ignore
        except ImportError as exc:
            raise RuntimeError(UI_TEXT["error_pywin32_missing"]) from exc

        self.pythoncom = pythoncom
        self.pythoncom.CoInitialize()
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
        except Exception as exc:
            raise RuntimeError(UI_TEXT["error_outlook_connect"]) from exc
        return self

    def __exit__(self, exc_type: object, exc: object, traceback: object) -> None:
        if self.pythoncom is not None:
            self.pythoncom.CoUninitialize()

    def create_draft(
        self,
        target: DraftTarget,
        subject: str,
        body_text: str,
        attachment_paths: list[Path],
    ) -> None:
        if self.outlook is None:
            raise RuntimeError(UI_TEXT["error_outlook_not_initialized"])

        mail = self.outlook.CreateItem(0)
        mail.To = target.email
        mail.Subject = subject

        for attachment_path in attachment_paths:
            mail.Attachments.Add(str(attachment_path.resolve()))

        mail.Display()
        time.sleep(0.2)
        signature_html = mail.HTMLBody or ""
        mail.HTMLBody = body_text_to_html(body_text) + signature_html
        mail.Save()

        if self.keep_displayed:
            return

        try:
            mail.GetInspector.Close(0)
        except Exception:
            logging.info("Draft inspector could not be closed automatically", exc_info=True)


def create_drafts_worker(
    event_queue: queue.Queue[dict[str, object]],
    cancel_event: threading.Event,
    targets: list[DraftTarget],
    skipped_reports: list[ReportRecord],
    subject_template: str,
    body_template: str,
    attachment_paths: list[Path],
    keep_displayed: bool,
) -> None:
    report_records = list(skipped_reports)
    drafted_count = 0
    error_count = 0

    try:
        with OutlookDraftSession(keep_displayed=keep_displayed) as outlook_session:
            for index, target in enumerate(targets, start=1):
                if cancel_event.is_set():
                    break

                subject = render_template(subject_template, target)
                body_text = render_template(body_template, target)

                try:
                    outlook_session.create_draft(target, subject, body_text, attachment_paths)
                except Exception as exc:
                    error_count += 1
                    logging.exception("Draft creation failed for row %s", target.row_number)
                    report_records.append(
                        make_report_record(target, "error", str(exc), subject=subject)
                    )
                else:
                    drafted_count += 1
                    report_records.append(
                        make_report_record(target, "drafted", UI_TEXT["report_drafted"], subject=subject)
                    )

                event_queue.put(
                    {
                        "type": "progress",
                        "current": index,
                        "total": len(targets),
                        "drafted": drafted_count,
                        "errors": error_count,
                    }
                )
    except Exception as exc:
        logging.exception("Outlook setup failed")
        for target in targets:
            subject = render_template(subject_template, target)
            report_records.append(make_report_record(target, "error", str(exc), subject=subject))
        error_count = len(targets)

    report_path = write_report(report_records)
    skipped_count = sum(1 for record in report_records if record.status.startswith("skipped_"))
    event_queue.put(
        {
            "type": "complete",
            "drafted": drafted_count,
            "skipped": skipped_count,
            "errors": error_count,
            "cancelled": cancel_event.is_set(),
            "report_path": str(report_path),
        }
    )


def choose_font_family(root: tk.Tk) -> str:
    available = set(tkfont.families(root))
    for family in ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo"):
        if family in available:
            return family
    return "TkDefaultFont"


class MailDraftApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.configure(bg=COLORS["background"])
        self.root.minsize(860, 800)
        self.root.geometry("1120x820")

        self.font_family = choose_font_family(root)
        self.fonts = {
            "title": (self.font_family, 18, "bold"),
            "subtitle": (self.font_family, 10),
            "label": (self.font_family, 10, "bold"),
            "body": (self.font_family, 10),
            "small": (self.font_family, 9),
            "footer": (self.font_family, 8),
            "button": (self.font_family, 10, "bold"),
        }

        self.event_queue: queue.Queue[dict[str, object]] = queue.Queue()
        self.cancel_event = threading.Event()
        self.worker_thread: threading.Thread | None = None
        self.csv_records: list[CsvRecord] = []
        self.csv_fieldnames: list[str] = []
        self.attachment_paths: list[Path] = []

        self.csv_path_var = tk.StringVar()
        self.company_column_var = tk.StringVar()
        self.name_column_var = tk.StringVar()
        self.email_column_var = tk.StringVar()
        self.limit_var = tk.StringVar(value=DEFAULT_LIMIT)
        self.status_var = tk.StringVar(value=UI_TEXT["status_select_csv"])
        self.report_var = tk.StringVar(value=UI_TEXT["report_pending"])
        self.keep_displayed_var = tk.BooleanVar(value=False)

        self._apply_window_icon()
        self._build_ui()
        self._load_initial_settings()
        self.root.after(QUEUE_POLL_INTERVAL_MS, self._poll_queue)

    def _apply_window_icon(self) -> None:
        if not ICON_PATH.exists():
            return
        try:
            self.root.iconbitmap(str(ICON_PATH))
            self.root.iconbitmap(default=str(ICON_PATH))
        except tk.TclError:
            return

    def _build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        header = tk.Frame(self.root, bg=COLORS["background"])
        header.grid(row=0, column=0, sticky="ew", padx=24, pady=(18, 10))
        header.columnconfigure(0, weight=1)

        tk.Label(
            header,
            text=APP_NAME,
            bg=COLORS["background"],
            fg=COLORS["text"],
            font=self.fonts["title"],
            anchor="w",
        ).grid(row=0, column=0, sticky="ew")
        tk.Label(
            header,
            text=f"{DISPLAY_SUBTITLE} / {UI_TEXT['main_description']}",
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=self.fonts["subtitle"],
            anchor="w",
        ).grid(row=1, column=0, sticky="ew", pady=(4, 0))

        main = tk.Frame(self.root, bg=COLORS["background"])
        main.grid(row=1, column=0, sticky="nsew", padx=24, pady=(0, 16))
        main.columnconfigure(0, weight=2)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(0, weight=1)

        left = tk.Frame(main, bg=COLORS["surface"], highlightbackground=COLORS["border"], highlightthickness=1)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        left.columnconfigure(1, weight=1)
        left.rowconfigure(6, weight=1)

        right = tk.Frame(main, bg=COLORS["surface"], highlightbackground=COLORS["border"], highlightthickness=1)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)
        right.rowconfigure(3, weight=1)

        self._build_input_area(left)
        self._build_preview_area(right)

        action_bar = tk.Frame(self.root, bg=COLORS["background"])
        action_bar.grid(row=2, column=0, sticky="ew", padx=24, pady=(0, 10))
        action_bar.columnconfigure(0, weight=1)
        action_bar.columnconfigure(1, weight=0)

        self.progress = ttk.Progressbar(action_bar, orient="horizontal", mode="determinate")
        self.progress.grid(row=0, column=0, sticky="ew", padx=(0, 14))

        self.generate_button = self._make_button(
            action_bar,
            text=UI_TEXT["button_generate"],
            command=self._start_generation,
            bg=COLORS["accent"],
            hover_bg=COLORS["accent_hover"],
        )
        self.generate_button.grid(row=0, column=1, sticky="e", padx=(0, 8))

        self.cancel_button = self._make_button(
            action_bar,
            text=UI_TEXT["button_cancel"],
            command=self._cancel_generation,
            bg=COLORS["muted"],
            hover_bg=COLORS["text"],
        )
        self.cancel_button.grid(row=0, column=2, sticky="e")
        self.cancel_button.configure(state=tk.DISABLED)

        self.status_label = tk.Label(
            action_bar,
            textvariable=self.status_var,
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=self.fonts["small"],
            anchor="w",
        )
        self.status_label.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        self._build_footer()

    def _build_footer(self) -> None:
        footer = tk.Frame(self.root, bg=COLORS["background"])
        footer.grid(row=3, column=0, sticky="ew", padx=24, pady=(0, 14))
        footer.columnconfigure(0, weight=1)
        footer.columnconfigure(1, weight=1)

        self.footer_frame = footer
        self.footer_left_text = (
            f"{UI_TEXT['footer_left']} {UI_TEXT['footer_separator']} "
            f"{UI_TEXT['footer_caption']}"
        )
        self.footer_right_text = (
            f"{UI_TEXT['footer_link_1']} {UI_TEXT['footer_separator']} "
            f"{UI_TEXT['footer_link_2']} {UI_TEXT['footer_separator']} "
            f"{UI_TEXT['footer_copyright']}"
        )
        self.footer_left_label = tk.Label(
            footer,
            text=self.footer_left_text,
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=self.fonts["footer"],
            anchor="center",
        )
        self.footer_right_label = tk.Label(
            footer,
            text=self.footer_right_text,
            bg=COLORS["background"],
            fg=COLORS["muted"],
            font=self.fonts["footer"],
            anchor="center",
        )
        self._footer_mode = ""
        self.root.bind("<Configure>", self._handle_footer_resize, add="+")
        self.root.after_idle(lambda: self._apply_footer_layout(self.root.winfo_width()))

    def _handle_footer_resize(self, event: tk.Event) -> None:
        if event.widget is self.root:
            self._apply_footer_layout(event.width)

    def _apply_footer_layout(self, width: int) -> None:
        required_width = (
            self.footer_left_label.winfo_reqwidth()
            + self.footer_right_label.winfo_reqwidth()
            + 48
        )
        threshold = max(FOOTER_WIDE_THRESHOLD, required_width)
        mode = "wide" if width >= threshold else "narrow"
        if self._footer_mode == mode:
            return
        self._footer_mode = mode
        self.footer_left_label.grid_forget()
        self.footer_right_label.grid_forget()
        if mode == "wide":
            self.footer_left_label.configure(anchor="w", justify="left")
            self.footer_right_label.configure(anchor="e", justify="right")
            self.footer_left_label.grid(row=0, column=0, sticky="ew")
            self.footer_right_label.grid(row=0, column=1, sticky="ew")
            return

        self.footer_left_label.configure(anchor="center", justify="center")
        self.footer_right_label.configure(anchor="center", justify="center")
        self.footer_left_label.grid(row=0, column=0, columnspan=2, sticky="ew")
        self.footer_right_label.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(2, 0))

    def _build_input_area(self, parent: tk.Frame) -> None:
        padding = {"padx": 18, "pady": 7}
        parent.grid_columnconfigure(1, weight=1)

        self._make_label(parent, UI_TEXT["label_csv"]).grid(row=0, column=0, sticky="w", **padding)
        csv_entry = tk.Entry(parent, textvariable=self.csv_path_var, font=self.fonts["body"])
        csv_entry.grid(row=0, column=1, sticky="ew", padx=(0, 8), pady=7)
        self._make_button(parent, UI_TEXT["button_select"], self._browse_csv).grid(row=0, column=2, sticky="ew", padx=(0, 18), pady=7)

        self._make_label(parent, UI_TEXT["label_company_column"]).grid(row=1, column=0, sticky="w", **padding)
        self.company_combo = ttk.Combobox(parent, textvariable=self.company_column_var, state="readonly")
        self.company_combo.grid(row=1, column=1, columnspan=2, sticky="ew", padx=(0, 18), pady=7)
        self.company_combo.bind("<<ComboboxSelected>>", lambda _event: self._refresh_preview())

        self._make_label(parent, UI_TEXT["label_name_column"]).grid(row=2, column=0, sticky="w", **padding)
        self.name_combo = ttk.Combobox(parent, textvariable=self.name_column_var, state="readonly")
        self.name_combo.grid(row=2, column=1, columnspan=2, sticky="ew", padx=(0, 18), pady=7)
        self.name_combo.bind("<<ComboboxSelected>>", lambda _event: self._refresh_preview())

        self._make_label(parent, UI_TEXT["label_email_column"]).grid(row=3, column=0, sticky="w", **padding)
        self.email_combo = ttk.Combobox(parent, textvariable=self.email_column_var, state="readonly")
        self.email_combo.grid(row=3, column=1, columnspan=2, sticky="ew", padx=(0, 18), pady=7)
        self.email_combo.bind("<<ComboboxSelected>>", lambda _event: self._refresh_preview())

        self._make_label(parent, UI_TEXT["label_subject_template"]).grid(row=4, column=0, sticky="w", **padding)
        self.subject_entry = tk.Entry(parent, font=self.fonts["body"])
        self.subject_entry.grid(row=4, column=1, columnspan=2, sticky="ew", padx=(0, 18), pady=7)
        self.subject_entry.insert(0, DEFAULT_SUBJECT_TEMPLATE)
        self.subject_entry.bind("<KeyRelease>", lambda _event: self._refresh_preview())

        body_label = self._make_label(parent, UI_TEXT["label_body_template"])
        body_label.grid(row=5, column=0, sticky="nw", **padding)
        self.body_text = tk.Text(
            parent,
            height=10,
            wrap="word",
            font=self.fonts["body"],
            undo=True,
            padx=8,
            pady=8,
        )
        self.body_text.grid(row=5, column=1, columnspan=2, sticky="nsew", padx=(0, 18), pady=7)
        self.body_text.insert("1.0", DEFAULT_BODY_TEMPLATE)
        self.body_text.bind("<KeyRelease>", lambda _event: self._refresh_preview())

        self._make_label(parent, UI_TEXT["label_attachments"]).grid(row=6, column=0, sticky="nw", **padding)
        attachment_frame = tk.Frame(parent, bg=COLORS["surface"])
        attachment_frame.grid(row=6, column=1, columnspan=2, sticky="nsew", padx=(0, 18), pady=7)
        attachment_frame.columnconfigure(0, weight=1)
        attachment_frame.rowconfigure(0, weight=1)
        self.attachment_list = tk.Listbox(attachment_frame, height=4, font=self.fonts["small"])
        self.attachment_list.grid(row=0, column=0, columnspan=3, sticky="nsew")
        self._make_button(attachment_frame, UI_TEXT["button_add"], self._add_attachments).grid(row=1, column=0, sticky="ew", pady=(7, 0), padx=(0, 6))
        self._make_button(attachment_frame, UI_TEXT["button_remove"], self._remove_selected_attachment).grid(row=1, column=1, sticky="ew", pady=(7, 0), padx=6)
        self._make_button(attachment_frame, UI_TEXT["button_clear"], self._clear_attachments).grid(row=1, column=2, sticky="ew", pady=(7, 0), padx=(6, 0))

        self._make_label(parent, UI_TEXT["label_limit"]).grid(row=7, column=0, sticky="w", **padding)
        limit_frame = tk.Frame(parent, bg=COLORS["surface"])
        limit_frame.grid(row=7, column=1, columnspan=2, sticky="ew", padx=(0, 18), pady=7)
        limit_frame.columnconfigure(0, weight=0)
        limit_frame.columnconfigure(1, weight=1)
        limit_combo = ttk.Combobox(limit_frame, values=LIMIT_CHOICES, textvariable=self.limit_var, state="readonly", width=8)
        limit_combo.grid(row=0, column=0, sticky="w")
        limit_combo.bind("<<ComboboxSelected>>", lambda _event: self._refresh_preview())
        tk.Checkbutton(
            limit_frame,
            text=UI_TEXT["option_keep_displayed"],
            variable=self.keep_displayed_var,
            bg=COLORS["surface"],
            fg=COLORS["text"],
            activebackground=COLORS["surface"],
            font=self.fonts["small"],
        ).grid(row=0, column=1, sticky="w", padx=(18, 0))

    def _build_preview_area(self, parent: tk.Frame) -> None:
        tk.Label(
            parent,
            text=UI_TEXT["preview_title"],
            bg=COLORS["surface"],
            fg=COLORS["text"],
            font=self.fonts["label"],
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=18, pady=(16, 8))

        self.preview_text = tk.Text(
            parent,
            height=16,
            wrap="word",
            font=self.fonts["small"],
            bg=COLORS["surface_alt"],
            fg=COLORS["text"],
            padx=10,
            pady=10,
            state=tk.DISABLED,
        )
        self.preview_text.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 14))

        tk.Label(
            parent,
            text=UI_TEXT["report_title"],
            bg=COLORS["surface"],
            fg=COLORS["text"],
            font=self.fonts["label"],
            anchor="w",
        ).grid(row=2, column=0, sticky="ew", padx=18, pady=(0, 8))

        self.report_text = tk.Text(
            parent,
            height=10,
            wrap="word",
            font=self.fonts["small"],
            bg=COLORS["surface_alt"],
            fg=COLORS["text"],
            padx=10,
            pady=10,
            state=tk.DISABLED,
        )
        self.report_text.grid(row=3, column=0, sticky="nsew", padx=18, pady=(0, 14))
        self._set_report_text(self.report_var.get())

        tk.Label(
            parent,
            text=UI_TEXT["safety_note"],
            bg=COLORS["surface"],
            fg=COLORS["warning"],
            font=self.fonts["small"],
            wraplength=330,
            justify="left",
            anchor="w",
        ).grid(row=4, column=0, sticky="ew", padx=18, pady=(0, 16))

    def _make_label(self, parent: tk.Widget, text: str) -> tk.Label:
        return tk.Label(parent, text=text, bg=COLORS["surface"], fg=COLORS["text"], font=self.fonts["label"])

    def _make_button(
        self,
        parent: tk.Widget,
        text: str,
        command: object,
        bg: str = COLORS["surface_alt"],
        hover_bg: str = COLORS["border"],
    ) -> tk.Button:
        fg = "#FFFFFF" if bg in {COLORS["accent"], COLORS["muted"], COLORS["text"]} else COLORS["text"]
        button = tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg=fg,
            activebackground=hover_bg,
            activeforeground=fg,
            bd=0,
            relief="flat",
            font=self.fonts["button"],
            padx=12,
            pady=7,
            cursor="hand2",
        )
        button.bind("<Enter>", lambda _event: button.configure(bg=hover_bg) if button["state"] != tk.DISABLED else None)
        button.bind("<Leave>", lambda _event: button.configure(bg=bg) if button["state"] != tk.DISABLED else None)
        return button

    def _load_initial_settings(self) -> None:
        settings = load_settings()
        csv_path = normalize_cell(settings.get("csv_path"))
        if csv_path:
            self.csv_path_var.set(csv_path)
            self._load_csv(Path(csv_path), show_message=False)

        subject = normalize_cell(settings.get("subject_template"))
        if subject:
            self.subject_entry.delete(0, tk.END)
            self.subject_entry.insert(0, subject)

        body = settings.get("body_template")
        if isinstance(body, str) and body:
            self.body_text.delete("1.0", tk.END)
            self.body_text.insert("1.0", body)

        limit = normalize_cell(settings.get("limit"))
        if limit in LIMIT_CHOICES:
            self.limit_var.set(limit)

        self.keep_displayed_var.set(bool(settings.get("keep_displayed", False)))

        attachment_values = settings.get("attachments")
        if isinstance(attachment_values, list):
            for value in attachment_values:
                path = Path(str(value))
                if path.exists():
                    self.attachment_paths.append(path)
                    self.attachment_list.insert(tk.END, str(path))

        self._refresh_preview()

    def _browse_csv(self) -> None:
        selected = filedialog.askopenfilename(
            title=UI_TEXT["dialog_csv_title"],
            filetypes=(("CSV files", "*.csv"), ("All files", "*.*")),
            initialdir=str(DATA_DIR),
        )
        if not selected:
            return
        csv_path = Path(selected)
        self.csv_path_var.set(str(csv_path))
        self._load_csv(csv_path, show_message=True)

    def _load_csv(self, csv_path: Path, show_message: bool) -> None:
        try:
            records, fieldnames, encoding = read_csv_records(csv_path)
        except Exception as exc:
            self.csv_records = []
            self.csv_fieldnames = []
            self._set_status(UI_TEXT["status_csv_load_error"].format(error=exc), is_error=True)
            if show_message:
                messagebox.showerror(APP_NAME, UI_TEXT["dialog_csv_load_error"].format(error=exc))
            self._refresh_column_combos()
            return

        self.csv_records = records
        self.csv_fieldnames = fieldnames
        self._refresh_column_combos()
        inferred = infer_columns(fieldnames)
        self.company_column_var.set(inferred["company"])
        self.name_column_var.set(inferred["name"])
        self.email_column_var.set(inferred["email"])
        self._set_status(UI_TEXT["status_csv_loaded"].format(count=len(records), encoding=encoding))
        self._refresh_preview()

    def _refresh_column_combos(self) -> None:
        for combo in (self.company_combo, self.name_combo, self.email_combo):
            combo.configure(values=self.csv_fieldnames)
            if not self.csv_fieldnames:
                combo.set("")

    def _add_attachments(self) -> None:
        selected_paths = filedialog.askopenfilenames(title=UI_TEXT["dialog_attachment_title"])
        for selected in selected_paths:
            path = Path(selected)
            if path in self.attachment_paths:
                continue
            self.attachment_paths.append(path)
            self.attachment_list.insert(tk.END, str(path))

    def _remove_selected_attachment(self) -> None:
        selected_indices = list(self.attachment_list.curselection())
        for index in reversed(selected_indices):
            self.attachment_list.delete(index)
            del self.attachment_paths[index]

    def _clear_attachments(self) -> None:
        self.attachment_paths.clear()
        self.attachment_list.delete(0, tk.END)

    def _get_subject_template(self) -> str:
        return self.subject_entry.get().strip()

    def _get_body_template(self) -> str:
        return self.body_text.get("1.0", "end-1c")

    def _current_mapping(self) -> tuple[str, str, str]:
        return (
            self.company_column_var.get().strip(),
            self.name_column_var.get().strip(),
            self.email_column_var.get().strip(),
        )

    def _refresh_preview(self) -> None:
        if not self.csv_records:
            self._set_preview_text(UI_TEXT["preview_select_csv"])
            return

        company_column, name_column, email_column = self._current_mapping()
        if not all((company_column, name_column, email_column)):
            self._set_preview_text(UI_TEXT["preview_select_mapping"])
            return

        targets, skipped_reports = prepare_targets(
            self.csv_records,
            company_column,
            name_column,
            email_column,
            "1",
        )
        if not targets:
            skipped_count = len(skipped_reports)
            self._set_preview_text(UI_TEXT["preview_no_targets"].format(count=skipped_count))
            return

        target = targets[0]
        subject = render_template(self._get_subject_template(), target)
        body = render_template(self._get_body_template(), target)
        preview = UI_TEXT["preview_format"].format(email=target.email, subject=subject, body=body)
        self._set_preview_text(preview)

    def _set_preview_text(self, text: str) -> None:
        self.preview_text.configure(state=tk.NORMAL)
        self.preview_text.delete("1.0", tk.END)
        self.preview_text.insert("1.0", text)
        self.preview_text.configure(state=tk.DISABLED)

    def _set_report_text(self, text: str) -> None:
        self.report_text.configure(state=tk.NORMAL)
        self.report_text.delete("1.0", tk.END)
        self.report_text.insert("1.0", text)
        self.report_text.configure(state=tk.DISABLED)

    def _set_status(self, text: str, is_error: bool = False) -> None:
        self.status_var.set(text)
        color = COLORS["error"] if is_error else COLORS["muted"]
        if hasattr(self, "status_label"):
            self.status_label.configure(fg=color)

    def _validate_before_generation(self) -> tuple[list[DraftTarget], list[ReportRecord], list[Path]] | None:
        csv_path = Path(self.csv_path_var.get().strip())
        if not csv_path.exists():
            messagebox.showerror(APP_NAME, UI_TEXT["dialog_select_csv_error"])
            return None

        company_column, name_column, email_column = self._current_mapping()
        if not all((company_column, name_column, email_column)):
            messagebox.showerror(APP_NAME, UI_TEXT["dialog_mapping_error"])
            return None

        if not self.csv_records:
            self._load_csv(csv_path, show_message=True)
            if not self.csv_records:
                return None

        missing_attachments = [path for path in self.attachment_paths if not path.exists()]
        if missing_attachments:
            lines = "\n".join(str(path) for path in missing_attachments[:5])
            messagebox.showerror(APP_NAME, UI_TEXT["dialog_missing_attachments"].format(paths=lines))
            return None

        targets, skipped_reports = prepare_targets(
            self.csv_records,
            company_column,
            name_column,
            email_column,
            self.limit_var.get(),
        )
        if not targets:
            messagebox.showwarning(APP_NAME, UI_TEXT["dialog_no_targets"])
            return None

        return targets, skipped_reports, list(self.attachment_paths)

    def _start_generation(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            return

        validated = self._validate_before_generation()
        if validated is None:
            return

        targets, skipped_reports, attachment_paths = validated
        confirm_text = UI_TEXT["confirm_generation"].format(
            target_count=len(targets),
            skipped_count=len(skipped_reports),
        )
        if not messagebox.askokcancel(APP_NAME, confirm_text):
            return

        self.cancel_event.clear()
        self.progress.configure(maximum=max(1, len(targets)), value=0)
        self.generate_button.configure(state=tk.DISABLED)
        self.cancel_button.configure(state=tk.NORMAL)
        self._set_status(UI_TEXT["status_generating"])
        self._set_report_text(UI_TEXT["report_generating"])

        payload = {
            "csv_path": self.csv_path_var.get().strip(),
            "company_column": self.company_column_var.get().strip(),
            "name_column": self.name_column_var.get().strip(),
            "email_column": self.email_column_var.get().strip(),
            "subject_template": self._get_subject_template(),
            "body_template": self._get_body_template(),
            "attachments": [str(path) for path in self.attachment_paths],
            "limit": self.limit_var.get(),
            "keep_displayed": self.keep_displayed_var.get(),
        }
        try:
            save_settings(payload)
        except OSError:
            logging.warning("settings.json could not be saved", exc_info=True)

        self.worker_thread = threading.Thread(
            target=create_drafts_worker,
            args=(
                self.event_queue,
                self.cancel_event,
                targets,
                skipped_reports,
                self._get_subject_template(),
                self._get_body_template(),
                attachment_paths,
                self.keep_displayed_var.get(),
            ),
            daemon=True,
        )
        self.worker_thread.start()

    def _cancel_generation(self) -> None:
        self.cancel_event.set()
        self.cancel_button.configure(state=tk.DISABLED)
        self._set_status(UI_TEXT["status_cancel_requested"])

    def _poll_queue(self) -> None:
        try:
            while True:
                event = self.event_queue.get_nowait()
                self._handle_worker_event(event)
        except queue.Empty:
            pass

        if self.root.winfo_exists():
            self.root.after(QUEUE_POLL_INTERVAL_MS, self._poll_queue)

    def _handle_worker_event(self, event: dict[str, object]) -> None:
        event_type = event.get("type")
        if event_type == "progress":
            current = int(event.get("current", 0))
            total = int(event.get("total", 0))
            drafted = int(event.get("drafted", 0))
            errors = int(event.get("errors", 0))
            self.progress.configure(value=current)
            self._set_status(UI_TEXT["status_progress"].format(current=current, total=total, drafted=drafted, errors=errors))
            return

        if event_type == "complete":
            self.generate_button.configure(state=tk.NORMAL)
            self.cancel_button.configure(state=tk.DISABLED)
            drafted = int(event.get("drafted", 0))
            skipped = int(event.get("skipped", 0))
            errors = int(event.get("errors", 0))
            cancelled = bool(event.get("cancelled", False))
            report_path = str(event.get("report_path", ""))
            prefix = UI_TEXT["status_cancelled"] if cancelled else UI_TEXT["status_completed"]
            self._set_status(UI_TEXT["status_complete"].format(prefix=prefix, drafted=drafted, skipped=skipped, errors=errors))
            self._set_report_text(
                UI_TEXT["report_complete"].format(
                    prefix=prefix,
                    drafted=drafted,
                    skipped=skipped,
                    errors=errors,
                    report_path=report_path,
                )
            )

    def run(self) -> None:
        self.root.mainloop()


def main() -> None:
    root = tk.Tk()
    MailDraftApp(root).run()


if __name__ == "__main__":
    main()
