from __future__ import annotations

import re
import traceback
from datetime import datetime
from pathlib import Path
from typing import Any

_SECRET_PATTERNS = (
    re.compile(r"(?i)(authorization\s*:\s*bearer\s+)[^\s]+"),
    re.compile(r"(?i)((?:api[_-]?key|token|password)\s*[=:]\s*)[^\s,;]+"),
    re.compile(r"\b[A-Za-z0-9+/]{120,}={0,2}\b"),
)


def _sanitize(value: str) -> str:
    result = value
    for pattern in _SECRET_PATTERNS:
        result = pattern.sub(r"\1[redacted]", result)
    return result


def get_log_path(log_dir: str | Path | None = None, now: datetime | None = None) -> Path:
    base = Path(log_dir) if log_dir is not None else Path.cwd() / "logs"
    current = now or datetime.now()
    return base / f"{current:%Y-%m-%d}.log"


def write_debug_log(
    message: str,
    *,
    log_dir: str | Path | None = None,
    exc: BaseException | None = None,
    context: dict[str, Any] | None = None,
) -> Path | None:
    """Append a debug log. Logging failure never raises to the app."""
    try:
        path = get_log_path(log_dir)
        path.parent.mkdir(parents=True, exist_ok=True)
        lines = [f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {_sanitize(str(message))}"]
        if context:
            for key, value in context.items():
                lines.append(f"{key}={_sanitize(str(value))}")
        if exc is not None:
            lines.append(_sanitize("".join(traceback.format_exception(exc))))
        with path.open("a", encoding="utf-8") as file:
            file.write("\n".join(lines))
            file.write("\n")
        return path
    except Exception:
        return None
