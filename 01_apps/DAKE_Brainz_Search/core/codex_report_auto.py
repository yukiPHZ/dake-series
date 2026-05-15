from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from core.app_config import logs_dir, now_iso, read_text_safe
from core.codex_importer import CodexImportResult, SOURCE_TYPE_CODEX_REPORT_AUTO, import_codex_text
from core.db import BrainzDatabase
from core.remote_queue import destination_for, move_task_file


REPORT_EXTENSIONS = {".md", ".txt"}
REPORT_SYSTEM_DIRS = {"processed", "failed"}


@dataclass(frozen=True)
class CodexReportAutoItem:
    source_file: str
    destination_file: str
    status: str
    title: str = ""
    commit_hash: str = ""
    changed_files_count: int = 0
    skipped_duplicate: bool = False
    error: str = ""


@dataclass(frozen=True)
class CodexReportAutoResult:
    detected: int
    imported: int
    failed: int
    pending: int
    items: list[CodexReportAutoItem]
    log_path: str


def iter_report_files(report_folder: Path) -> list[Path]:
    if not report_folder.exists() or not report_folder.is_dir():
        return []

    files: list[Path] = []
    for path in sorted(report_folder.rglob("*"), key=lambda item: str(item).lower()):
        if not path.is_file():
            continue
        try:
            relative_parts = tuple(part.lower() for part in path.relative_to(report_folder).parts)
        except ValueError:
            relative_parts = tuple(part.lower() for part in path.parts)
        if any(part in REPORT_SYSTEM_DIRS for part in relative_parts):
            continue
        if path.suffix.lower() in REPORT_EXTENSIONS:
            files.append(path)
    return files


def count_pending_reports(report_folder: Path) -> int:
    return len(iter_report_files(report_folder))


def process_codex_reports_folder(
    database: BrainzDatabase,
    report_folder: Path,
    limit: int = 20,
) -> CodexReportAutoResult:
    report_folder.mkdir(parents=True, exist_ok=True)
    files = iter_report_files(report_folder)[:limit]
    items: list[CodexReportAutoItem] = []

    for source_file in files:
        try:
            raw_text = read_text_safe(source_file)
            if not raw_text.strip():
                raise ValueError("empty codex report")
            processed_destination = destination_for(report_folder, source_file, "processed")
            import_result = import_codex_report_text(
                raw_text=raw_text,
                database=database,
                source_path=source_file,
            )
            destination = move_task_file(report_folder, source_file, "processed", processed_destination)
            items.append(
                CodexReportAutoItem(
                    source_file=str(source_file.resolve()),
                    destination_file=str(destination),
                    status="processed",
                    title=import_result.title,
                    commit_hash=import_result.commit_hash,
                    changed_files_count=import_result.changed_files_count,
                    skipped_duplicate=import_result.skipped_duplicate,
                )
            )
        except Exception as exc:
            destination = move_task_file(report_folder, source_file, "failed")
            items.append(
                CodexReportAutoItem(
                    source_file=str(source_file.resolve()),
                    destination_file=str(destination),
                    status="failed",
                    error=str(exc),
                )
            )

    imported = sum(1 for item in items if item.status == "processed")
    failed = sum(1 for item in items if item.status == "failed")
    log_path = write_codex_report_log(items) if items else ""
    return CodexReportAutoResult(
        detected=len(files),
        imported=imported,
        failed=failed,
        pending=count_pending_reports(report_folder),
        items=items,
        log_path=log_path,
    )


def import_codex_report_text(
    raw_text: str,
    database: BrainzDatabase,
    source_path: Path,
) -> CodexImportResult:
    return import_codex_text(
        raw_text=raw_text,
        database=database,
        source_label=str(source_path.resolve()),
        source_type=SOURCE_TYPE_CODEX_REPORT_AUTO,
        preserve_raw_document=True,
    )


def write_codex_report_log(items: list[CodexReportAutoItem]) -> str:
    logs_dir().mkdir(parents=True, exist_ok=True)
    path = logs_dir() / f"codex_report_auto_{now_iso().replace(':', '').replace('-', '').replace('T', '_')}.log"
    lines = ["CODEX REPORT AUTO IMPORT:"]
    for item in items:
        lines.extend(
            [
                f"- source: {item.source_file}",
                f"  status: {item.status}",
                f"  destination: {item.destination_file}",
                f"  commit: {item.commit_hash}",
                f"  changed_files: {item.changed_files_count}",
                f"  skipped_duplicate: {item.skipped_duplicate}",
                f"  error: {item.error}",
            ]
        )
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return str(path)
