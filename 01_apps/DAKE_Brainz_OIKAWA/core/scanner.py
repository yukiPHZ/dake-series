# -*- coding: utf-8 -*-
from __future__ import annotations

import os
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path


SUPPORTED_EXTENSIONS = {".md", ".txt"}
EXCLUDED_DIRS = {".git", "__pycache__", "build", "dist", "node_modules", "logs"}
MAX_FILE_SIZE_BYTES = 5 * 1024 * 1024


@dataclass
class MemoryDocument:
    path: Path
    relative_path: str
    text: str
    modified_at: datetime
    size: int
    title: str


def read_text_safe(path: Path) -> str:
    for encoding in ("utf-8-sig", "utf-8", "cp932", "utf-16"):
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
        except OSError:
            return ""
    try:
        return path.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return ""


def extract_title(path: Path, text: str) -> str:
    for line in text.splitlines()[:80]:
        stripped = line.strip()
        if stripped.startswith("#"):
            title = stripped.lstrip("#").strip()
            if title:
                return title
    return path.stem


def iter_memory_files(root: Path):
    for current, dirs, files in os.walk(root):
        dirs[:] = [name for name in dirs if name.lower() not in EXCLUDED_DIRS]
        current_path = Path(current)
        for file_name in files:
            path = current_path / file_name
            if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
                continue
            yield path


def scan_memory(memory_root: Path) -> tuple[list[MemoryDocument], int]:
    documents: list[MemoryDocument] = []
    skipped = 0
    root = memory_root.resolve()

    for path in iter_memory_files(root):
        try:
            stat = path.stat()
        except OSError:
            skipped += 1
            continue

        if stat.st_size > MAX_FILE_SIZE_BYTES:
            skipped += 1
            continue

        text = read_text_safe(path)
        if not text.strip():
            skipped += 1
            continue

        try:
            relative_path = str(path.resolve().relative_to(root))
        except ValueError:
            relative_path = path.name

        documents.append(
            MemoryDocument(
                path=path.resolve(),
                relative_path=relative_path,
                text=text,
                modified_at=datetime.fromtimestamp(stat.st_mtime),
                size=stat.st_size,
                title=extract_title(path, text),
            )
        )

    return documents, skipped
