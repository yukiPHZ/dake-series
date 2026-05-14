from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from threading import Event

from core.app_config import logs_dir, now_iso
from core.db import BrainzDatabase
from core.file_scanner import iter_supported_files, scan_file


@dataclass(frozen=True)
class WatchScanResult:
    folder: str
    checked: int
    changed_files: list[str]
    checked_at: str
    log_path: str = ""


def detect_changed_files(
    database: BrainzDatabase,
    watch_folder: Path,
    cancel_event: Event | None = None,
    max_changes: int = 200,
) -> WatchScanResult:
    files = iter_supported_files(watch_folder, cancel_event=cancel_event)
    resolved_paths = [str(path.resolve()) for path in files]
    known_hashes = database.document_hashes_for_paths(resolved_paths)
    changed_files: list[str] = []

    for path in files:
        if cancel_event and cancel_event.is_set():
            break
        try:
            scanned = scan_file(path)
        except (OSError, UnicodeError):
            continue
        current_path = str(scanned.path)
        if known_hashes.get(current_path) != scanned.content_hash:
            changed_files.append(current_path)
            if len(changed_files) >= max_changes:
                break

    return WatchScanResult(
        folder=str(watch_folder.resolve()),
        checked=len(files),
        changed_files=changed_files,
        checked_at=now_iso(),
    )


def write_watch_log(lines: list[str]) -> Path:
    logs_dir().mkdir(parents=True, exist_ok=True)
    path = logs_dir() / f"watch_{now_iso().replace(':', '').replace('-', '').replace('T', '_')}.log"
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path
