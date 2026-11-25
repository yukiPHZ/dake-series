from __future__ import annotations

import os
import tempfile
from pathlib import Path


def atomic_replace(source: str | Path, destination: str | Path) -> Path:
    """Replace destination with source without leaving a partial destination file."""
    source_path = Path(source)
    destination_path = Path(destination)
    destination_path.parent.mkdir(parents=True, exist_ok=True)
    os.replace(source_path, destination_path)
    return destination_path


def atomic_write_bytes(path: str | Path, data: bytes) -> Path:
    """Write bytes via a temporary file, then atomically replace the target."""
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    temp_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="wb",
            delete=False,
            dir=str(target.parent),
            prefix=f".{target.name}.",
            suffix=".tmp",
        ) as temp_file:
            temp_file.write(data)
            temp_file.flush()
            os.fsync(temp_file.fileno())
            temp_path = Path(temp_file.name)
        return atomic_replace(temp_path, target)
    except Exception:
        if temp_path is not None:
            try:
                temp_path.unlink(missing_ok=True)
            except OSError:
                pass
        raise


def atomic_write_text(path: str | Path, text: str, encoding: str = "utf-8") -> Path:
    """Write UTF-8 text safely without risking the previous file on failure."""
    return atomic_write_bytes(path, text.encode(encoding))
