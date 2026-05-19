# -*- coding: utf-8 -*-
from __future__ import annotations

from datetime import datetime
from pathlib import Path

from .settings import logs_dir


def make_timestamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def display_timestamp() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


class BackupLogger:
    def __init__(self, timestamp: str | None = None) -> None:
        self.timestamp = timestamp or make_timestamp()
        logs_dir().mkdir(parents=True, exist_ok=True)
        self.path: Path = logs_dir() / f"backup_{self.timestamp}.log"

    def write(self, message: str) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        line = f"[{display_timestamp()}] {message}\n"
        with self.path.open("a", encoding="utf-8") as handle:
            handle.write(line)

