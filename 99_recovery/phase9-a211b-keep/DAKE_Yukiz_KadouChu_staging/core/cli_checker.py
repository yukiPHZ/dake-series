from __future__ import annotations

import json
import os
import shutil
import subprocess
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any

from core.app_config import outputs_dir

CLI_TOOLS: dict[str, dict[str, str]] = {
    "ffmpeg": {"command": "ffmpeg", "label": "FFMPEG"},
    "ffprobe": {"command": "ffprobe", "label": "FFPROBE"},
    "yt-dlp": {"command": "yt-dlp", "label": "YT-DLP"},
    "gh": {"command": "gh", "label": "GH"},
    "wrangler": {"command": "wrangler", "label": "WRANGLER"},
    "npm": {"command": "npm", "label": "NPM"},
    "ollama": {"command": "ollama", "label": "OLLAMA"},
}

TOOL_PURPOSES = {
    "ffmpeg": "プレビュークリップ作成と動画処理に使います。",
    "ffprobe": "動画の尺、解像度、fps、codec確認に使います。",
    "yt-dlp": "YouTube LIVE URLのメタデータ取得に使います。",
    "gh": "将来のGitHub操作補助に使います。認証は gh auth login で行います。",
    "wrangler": "将来のCloudflare操作補助に使います。認証は wrangler login で行います。",
    "ollama": "ローカル補助脳コメントとメタデータ案の生成に使います。",
    "nvidia-smi": "GPU名とVRAM表示に使います。NVIDIAドライバに含まれる場合があります。",
}

SYSTEM_CHECK_KEYS = ["ffmpeg", "ffprobe", "yt-dlp", "gh", "wrangler", "npm", "ollama"]
INSTALL_TARGET_KEYS = ["ffmpeg", "ffprobe", "yt-dlp", "gh", "wrangler"]

INSTALL_TOOL_PURPOSES = {
    "ffmpeg": "動画情報取得、切り出し、NVENCエンコードに使います。",
    "ffprobe": "動画の尺、解像度、fps、codec確認に使います。",
    "yt-dlp": "YouTube LIVEアーカイブ等の取得補助に使います。",
    "gh": "GitHub Releaseや認証状態確認に使います。",
    "wrangler": "Cloudflare Pages / Workers 操作確認に使います。",
    "npm": "WranglerにはNode.js / npmが必要です。",
    "nvidia-smi": "GPU名とVRAM表示に使います。",
    "nvenc": "NVIDIA GPUエンコード確認に使います。",
}

INSTALL_COMMANDS = {
    "ffmpeg": "winget install Gyan.FFmpeg",
    "yt-dlp": "winget install yt-dlp.yt-dlp",
    "gh": "winget install GitHub.cli",
    "nodejs": "winget install OpenJS.NodeJS.LTS",
    "wrangler": "npm install -g wrangler",
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


def _first_line(text: str, limit: int = 120) -> str:
    line = text.strip().splitlines()[0] if text.strip() else ""
    return line[:limit]


def _status(state: str, path: str | None = None, detail: str = "") -> dict[str, str | None]:
    return {"state": state, "path": path, "detail": detail}


def _check_version(path: str, key: str) -> dict[str, str | None]:
    args = [path, "--version"]
    if key in {"ffmpeg", "ffprobe"}:
        args = [path, "-version"]
    try:
        completed = run_command(args, timeout=8)
    except Exception as exc:
        return _status("UNAVAILABLE", path, str(exc)[:120])
    detail = _first_line(completed.stdout or completed.stderr)
    if completed.returncode == 0:
        return _status("ONLINE", path, detail)
    return _status("UNAVAILABLE", path, detail)


def _check_gh(path: str) -> dict[str, str | None]:
    try:
        version = run_command([path, "--version"], timeout=8)
    except Exception as exc:
        return _status("UNAVAILABLE", path, str(exc)[:120])
    if version.returncode != 0:
        return _status("UNAVAILABLE", path, _first_line(version.stderr or version.stdout))
    try:
        completed = run_command([path, "auth", "status"], timeout=15)
    except Exception as exc:
        return _status("UNAUTHORIZED", path, str(exc)[:120])
    detail = _first_line(completed.stdout or completed.stderr)
    if completed.returncode == 0:
        return _status("AUTHORIZED", path, detail or "gh auth status OK")
    return _status("UNAUTHORIZED", path, detail or "gh auth status failed")


def _check_wrangler(path: str) -> dict[str, str | None]:
    try:
        version = run_command([path, "--version"], timeout=12)
    except Exception as exc:
        return _status("UNAVAILABLE", path, str(exc)[:120])
    if version.returncode != 0:
        return _status("UNAVAILABLE", path, _first_line(version.stderr or version.stdout))
    try:
        completed = run_command([path, "whoami"], timeout=20)
    except Exception as exc:
        return _status("UNAUTHORIZED", path, str(exc)[:120])
    detail = _first_line(completed.stdout or completed.stderr)
    if completed.returncode == 0:
        return _status("AUTHORIZED", path, detail or "wrangler whoami OK")
    return _status("UNAUTHORIZED", path, detail or "wrangler whoami failed")


def get_ollama_models(timeout: float = 1.5) -> list[str]:
    try:
        with urllib.request.urlopen("http://localhost:11434/api/tags", timeout=timeout) as response:
            if not 200 <= response.status < 300:
                return []
            payload = json.loads(response.read().decode("utf-8", errors="replace"))
    except Exception:
        return []
    models = payload.get("models", [])
    names = [str(model.get("name")) for model in models if model.get("name")]
    return names[:3]


def _check_ollama(path: str | None) -> dict[str, str | None]:
    models = get_ollama_models()
    if models:
        return _status("READY", path, "models: " + ", ".join(models))
    if is_ollama_api_ready():
        return _status("READY", path, "localhost:11434")
    return _status("MISSING", path, "API not responding")


def check_cli_environment() -> dict[str, dict[str, str | None]]:
    statuses: dict[str, dict[str, str | None]] = {}
    for key, spec in CLI_TOOLS.items():
        command = spec["command"]
        path = shutil.which(command)
        if key == "ollama":
            statuses[key] = _check_ollama(path)
            continue
        if not path:
            statuses[key] = _status("MISSING")
            continue
        if key == "gh":
            statuses[key] = _check_gh(path)
        elif key == "wrangler":
            statuses[key] = _check_wrangler(path)
        else:
            statuses[key] = _check_version(path, key)
    return statuses


def is_ollama_api_ready(timeout: float = 1.5) -> bool:
    try:
        with urllib.request.urlopen("http://localhost:11434/api/tags", timeout=timeout) as response:
            return 200 <= response.status < 300
    except Exception:
        return False


def check_nvenc(ffmpeg_path: str | None) -> dict[str, str]:
    if not ffmpeg_path:
        return {"state": "CHECK SKIPPED", "detail": "FFmpeg missing"}
    try:
        completed = run_command([ffmpeg_path, "-hide_banner", "-encoders"], timeout=12)
    except Exception as exc:
        return {"state": "UNAVAILABLE", "detail": str(exc)[:120]}
    output = f"{completed.stdout}\n{completed.stderr}".lower()
    h264_ready = "h264_nvenc" in output
    hevc_ready = "hevc_nvenc" in output
    if h264_ready or hevc_ready:
        details = []
        details.append("H264_NVENC READY" if h264_ready else "H264_NVENC UNAVAILABLE")
        details.append("HEVC_NVENC READY" if hevc_ready else "HEVC_NVENC UNAVAILABLE")
        return {
            "state": "ONLINE",
            "detail": " / ".join(details),
            "h264_nvenc": "READY" if h264_ready else "UNAVAILABLE",
            "hevc_nvenc": "READY" if hevc_ready else "UNAVAILABLE",
        }
    return {"state": "UNAVAILABLE", "detail": "h264_nvenc / hevc_nvenc not listed"}


def check_gpu_info() -> dict[str, str]:
    path = shutil.which("nvidia-smi")
    if not path:
        return {"state": "SKIPPED", "detail": "nvidia-smi not found"}
    try:
        completed = run_command(
            [path, "--query-gpu=name,memory.total", "--format=csv,noheader"],
            timeout=8,
        )
    except Exception as exc:
        return {"state": "SKIPPED", "detail": str(exc)[:120]}
    if completed.returncode != 0:
        return {"state": "SKIPPED", "detail": _first_line(completed.stderr or completed.stdout)}
    line = _first_line(completed.stdout, limit=180)
    if not line:
        return {"state": "SKIPPED", "detail": "No GPU reported"}
    parts = [part.strip() for part in line.split(",", 1)]
    name = parts[0]
    vram = parts[1].replace(" MiB", " MB") if len(parts) > 1 else ""
    detail = f"{name} / VRAM {vram}" if vram else name
    return {"state": "READY", "detail": detail, "name": name, "vram": vram}


def run_system_check() -> dict[str, Any]:
    statuses = check_cli_environment()
    ffmpeg_path = statuses.get("ffmpeg", {}).get("path")
    nvenc = check_nvenc(ffmpeg_path)
    gpu = check_gpu_info()
    install_commands = get_install_commands(statuses)
    install_guide = write_install_guide(statuses, nvenc, gpu)
    return {
        "cli": statuses,
        "nvenc": nvenc,
        "gpu": gpu,
        "install_commands": install_commands,
        "install_guide": str(install_guide) if install_guide else "",
    }


def _state_of(statuses: dict[str, dict[str, str | None]], key: str) -> str:
    return str(statuses.get(key, {}).get("state") or "UNKNOWN")


def _detail_of(statuses: dict[str, dict[str, str | None]], key: str) -> str:
    return str(statuses.get(key, {}).get("detail") or "")


def _needs_install(state: str) -> bool:
    return state in {"MISSING", "UNAVAILABLE"}


def _status_line(label: str, state: str, detail: str = "") -> str:
    if detail and state not in detail:
        return f"- {label}: {state} / {detail}"
    return f"- {label}: {state}"


def get_install_commands(statuses: dict[str, dict[str, str | None]]) -> list[str]:
    commands: list[str] = []
    ffmpeg_state = _state_of(statuses, "ffmpeg")
    ffprobe_state = _state_of(statuses, "ffprobe")
    if _needs_install(ffmpeg_state) or _needs_install(ffprobe_state):
        commands.append(INSTALL_COMMANDS["ffmpeg"])
    if _needs_install(_state_of(statuses, "yt-dlp")):
        commands.append(INSTALL_COMMANDS["yt-dlp"])
    if _needs_install(_state_of(statuses, "gh")):
        commands.append(INSTALL_COMMANDS["gh"])
    if _needs_install(_state_of(statuses, "wrangler")):
        if _needs_install(_state_of(statuses, "npm")):
            commands.append(INSTALL_COMMANDS["nodejs"])
        commands.append(INSTALL_COMMANDS["wrangler"])
    return commands


def write_install_guide(
    statuses: dict[str, dict[str, str | None]],
    nvenc: dict[str, str],
    gpu: dict[str, str],
) -> Path:
    guide_dir = outputs_dir() / "system_check"
    guide_dir.mkdir(parents=True, exist_ok=True)
    path = guide_dir / "install_guide.txt"
    commands = get_install_commands(statuses)
    gpu_state = gpu.get("state", "UNKNOWN")
    gpu_detail = gpu.get("detail", "")
    gpu_line = str(gpu_detail) if gpu_state == "READY" and gpu_detail else str(gpu_state)
    ollama_state = _state_of(statuses, "ollama")
    ollama_detail = _detail_of(statuses, "ollama")
    memo = "OllamaとGPUは認識済みです。" if ollama_state == "READY" and gpu_state == "READY" else "OllamaとGPUは環境により認識状態が変わります。"
    lines = [
        "# Dakeユキズ稼働中｜CLI導入補助",
        "",
        "## 現在の状態",
        "",
        f"- GPU: {gpu_line}",
        _status_line("OLLAMA", ollama_state, ollama_detail),
        _status_line("FFMPEG", _state_of(statuses, "ffmpeg"), _detail_of(statuses, "ffmpeg")),
        _status_line("FFPROBE", _state_of(statuses, "ffprobe"), _detail_of(statuses, "ffprobe")),
        _status_line("YT-DLP", _state_of(statuses, "yt-dlp"), _detail_of(statuses, "yt-dlp")),
        _status_line("GH", _state_of(statuses, "gh"), _detail_of(statuses, "gh")),
        _status_line("WRANGLER", _state_of(statuses, "wrangler"), _detail_of(statuses, "wrangler")),
        _status_line("NPM", _state_of(statuses, "npm"), _detail_of(statuses, "npm")),
        _status_line("NVENC", nvenc.get("state", "UNKNOWN"), nvenc.get("detail", "")),
        "",
        "## 入れるとできること",
        "",
        "### FFmpeg / FFprobe",
        INSTALL_TOOL_PURPOSES["ffmpeg"],
        "",
        "### yt-dlp",
        INSTALL_TOOL_PURPOSES["yt-dlp"],
        "",
        "### GitHub CLI",
        INSTALL_TOOL_PURPOSES["gh"],
        "",
        "### Wrangler",
        INSTALL_TOOL_PURPOSES["wrangler"],
        "WranglerにはNode.js / npm が必要です。",
        "npmがない場合は先にNode.jsを入れてください。",
        "",
        "## 推奨インストール方法",
        "",
        "※このアプリは自動インストールしません。",
        "※実行前に内容を確認してください。",
        "",
        "### FFmpeg",
        INSTALL_COMMANDS["ffmpeg"],
        "",
        "### yt-dlp",
        INSTALL_COMMANDS["yt-dlp"],
        "",
        "### GitHub CLI",
        INSTALL_COMMANDS["gh"],
        "",
        "### Node.js / npm",
        INSTALL_COMMANDS["nodejs"],
        "",
        "### Wrangler",
        INSTALL_COMMANDS["wrangler"],
        "",
        "## 今回コピーされる候補",
        "",
        *(commands if commands else ["不足している導入対象CLIはありません。"]),
        "",
        "## インストール後",
        "",
        "1. PowerShellを開き直す",
        "2. Dakeユキズ稼働中を再起動する",
        "3. Run System Check を押す",
        "",
        "## 確認コマンド",
        "",
        "ffmpeg -version",
        "ffprobe -version",
        "yt-dlp --version",
        "gh --version",
        "wrangler --version",
        "npm --version",
        "",
        "## 補助脳メモ",
        "",
        memo,
        "次はFFmpegがONLINEになると、動画解析とNVENC確認へ進めます。",
        "",
    ]
    path.write_text("\n".join(lines).strip() + "\n", encoding="utf-8")
    return path


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
