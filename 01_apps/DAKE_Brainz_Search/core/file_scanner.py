from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from threading import Event


SUPPORTED_EXTENSIONS = {".txt", ".md", ".json"}
MAX_FILE_BYTES = 10 * 1024 * 1024
IGNORED_DIR_NAMES = {"build", "dist", "__pycache__", ".git", ".venv", "node_modules"}
IGNORED_RELATIVE_PARTS = {("data", "logs"), ("data", "exports"), ("codex_reports",)}


@dataclass(frozen=True)
class ScannedFile:
    path: Path
    title: str
    source_type: str
    created_at: str
    modified_at: str
    content: str
    content_hash: str


def format_time(timestamp: float) -> str:
    return datetime.fromtimestamp(timestamp).isoformat(timespec="seconds")


def is_ignored_path(root: Path, path: Path) -> bool:
    try:
        relative_parts = path.resolve().relative_to(root.resolve()).parts
    except ValueError:
        relative_parts = path.parts

    lowered = tuple(part.lower() for part in relative_parts)
    if any(part in IGNORED_DIR_NAMES for part in lowered):
        return True

    for ignored_parts in IGNORED_RELATIVE_PARTS:
        ignored_length = len(ignored_parts)
        for index in range(0, max(0, len(lowered) - ignored_length + 1)):
            if lowered[index : index + ignored_length] == ignored_parts:
                return True
    return False


def iter_supported_files(root: Path, cancel_event: Event | None = None) -> list[Path]:
    if not root.exists() or not root.is_dir():
        return []

    files: list[Path] = []
    for path in sorted(root.rglob("*"), key=lambda item: str(item).lower()):
        if cancel_event and cancel_event.is_set():
            break
        if not path.is_file():
            continue
        if is_ignored_path(root, path):
            continue
        if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            continue
        try:
            if path.stat().st_size > MAX_FILE_BYTES:
                continue
        except OSError:
            continue
        files.append(path)
    return files


def read_text(path: Path) -> str:
    for encoding in ("utf-8-sig", "utf-8", "cp932", "utf-16"):
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding="utf-8", errors="replace")


def normalize_content(path: Path, text: str) -> str:
    if path.suffix.lower() != ".json":
        return text
    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        return text
    return json.dumps(data, ensure_ascii=False, indent=2)


def scan_file(path: Path) -> ScannedFile:
    stat = path.stat()
    text = normalize_content(path, read_text(path))
    digest = hashlib.sha256(text.encode("utf-8", errors="replace")).hexdigest()
    return ScannedFile(
        path=path.resolve(),
        title=path.stem,
        source_type=path.suffix.lower().lstrip("."),
        created_at=format_time(stat.st_ctime),
        modified_at=format_time(stat.st_mtime),
        content=text,
        content_hash=digest,
    )
