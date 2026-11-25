from __future__ import annotations

import copy
import json
from datetime import datetime
from pathlib import Path
from typing import Any

try:
    from .atomic_io import atomic_write_text
    from .logging import write_debug_log
except ImportError:  # pragma: no cover - direct script fallback.
    from atomic_io import atomic_write_text  # type: ignore
    from logging import write_debug_log  # type: ignore


def _backup_broken_file(path: Path) -> Path | None:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = path.with_name(f"{path.name}.broken_{timestamp}")
    counter = 2
    while backup.exists():
        backup = path.with_name(f"{path.name}.broken_{timestamp}_{counter}")
        counter += 1
    try:
        path.replace(backup)
        return backup
    except OSError as exc:
        write_debug_log("failed to backup broken config", exc=exc, context={"path": path})
        return None


def safe_load_json_config(path: str | Path, default: dict[str, Any] | None = None) -> dict[str, Any]:
    config_path = Path(path)
    fallback = copy.deepcopy(default or {})
    if not config_path.exists():
        return fallback
    try:
        with config_path.open("r", encoding="utf-8") as file:
            loaded = json.load(file)
        if isinstance(loaded, dict):
            merged = copy.deepcopy(fallback)
            merged.update(loaded)
            return merged
        _backup_broken_file(config_path)
        return fallback
    except (json.JSONDecodeError, OSError, UnicodeDecodeError) as exc:
        write_debug_log("failed to load json config", exc=exc, context={"path": config_path})
        _backup_broken_file(config_path)
        return fallback


def safe_save_json_config(path: str | Path, data: dict[str, Any]) -> Path:
    text = json.dumps(data, ensure_ascii=False, indent=2, sort_keys=True)
    return atomic_write_text(path, f"{text}\n")
