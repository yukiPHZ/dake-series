from __future__ import annotations

import json
import os
import shutil
import subprocess
import urllib.error
import urllib.parse
import urllib.request
from typing import Any

CLI_TOOLS: dict[str, dict[str, str]] = {
    "ffmpeg": {"command": "ffmpeg", "label": "FFMPEG"},
    "ffprobe": {"command": "ffprobe", "label": "FFPROBE"},
    "yt-dlp": {"command": "yt-dlp", "label": "YT-DLP"},
    "gh": {"command": "gh", "label": "GH"},
    "wrangler": {"command": "wrangler", "label": "WRANGLER"},
    "ollama": {"command": "ollama", "label": "OLLAMA"},
}


def _creationflags() -> int:
    if os.name == "nt":
        return subprocess.CREATE_NO_WINDOW  # type: ignore[attr-defined]
    return 0


def run_command(args: list[str], timeout: int = 20) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        args,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
        shell=False,
        creationflags=_creationflags(),
    )


def check_cli_environment() -> dict[str, dict[str, str | None]]:
    statuses: dict[str, dict[str, str | None]] = {}
    for key, spec in CLI_TOOLS.items():
        command = spec["command"]
        path = shutil.which(command)
        if not path:
            statuses[key] = {"state": "MISSING", "path": None, "detail": ""}
            continue

        if key == "ollama":
            if is_ollama_api_ready():
                statuses[key] = {"state": "LOCAL READY", "path": path, "detail": "localhost:11434"}
            else:
                statuses[key] = {"state": "CLI ONLINE", "path": path, "detail": "API not responding"}
            continue

        version_args = [path, "--version"]
        if key in {"ffmpeg", "ffprobe"}:
            version_args = [path, "-version"]
        try:
            completed = run_command(version_args, timeout=8)
            first_line = (completed.stdout or completed.stderr).splitlines()[0] if (completed.stdout or completed.stderr) else ""
            state = "ONLINE" if completed.returncode == 0 else "FOUND"
            statuses[key] = {"state": state, "path": path, "detail": first_line[:120]}
        except Exception as exc:
            statuses[key] = {"state": "FOUND", "path": path, "detail": str(exc)[:120]}
    return statuses


def is_ollama_api_ready(timeout: float = 1.5) -> bool:
    try:
        with urllib.request.urlopen("http://localhost:11434/api/tags", timeout=timeout) as response:
            return 200 <= response.status < 300
    except Exception:
        return False


def check_nvenc(ffmpeg_path: str | None) -> dict[str, str]:
    if not ffmpeg_path:
        return {"state": "UNAVAILABLE", "detail": "FFmpeg missing"}
    try:
        completed = run_command([ffmpeg_path, "-hide_banner", "-encoders"], timeout=12)
    except Exception as exc:
        return {"state": "UNAVAILABLE", "detail": str(exc)[:120]}
    output = f"{completed.stdout}\n{completed.stderr}".lower()
    if "h264_nvenc" in output:
        return {"state": "ONLINE", "detail": "h264_nvenc detected"}
    return {"state": "UNAVAILABLE", "detail": "h264_nvenc not listed"}


def fetch_youtube_metadata(url: str, ytdlp_path: str) -> dict[str, Any]:
    parsed = urllib.parse.urlparse(url)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise ValueError("Use a valid http(s) URL.")
    args = [
        ytdlp_path,
        "--dump-single-json",
        "--skip-download",
        "--no-playlist",
        "--no-warnings",
        url,
    ]
    completed = run_command(args, timeout=30)
    if completed.returncode != 0:
        message = completed.stderr.strip() or "yt-dlp metadata fetch failed."
        raise RuntimeError(message[:500])
    try:
        return json.loads(completed.stdout)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"yt-dlp returned invalid JSON: {exc}") from exc
