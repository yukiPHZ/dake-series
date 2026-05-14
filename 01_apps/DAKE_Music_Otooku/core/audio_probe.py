# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import subprocess
from pathlib import Path


def probe_duration(audio_path: Path, ffprobe_command: str = "ffprobe") -> float | None:
    try:
        result = subprocess.run(
            [
                ffprobe_command,
                "-v",
                "error",
                "-show_entries",
                "format=duration",
                "-of",
                "json",
                str(audio_path),
            ],
            capture_output=True,
            text=True,
            timeout=8,
            check=False,
        )
    except Exception:
        return None
    if result.returncode != 0:
        return None
    try:
        data = json.loads(result.stdout)
        return float(data.get("format", {}).get("duration"))
    except Exception:
        return None

