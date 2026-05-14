# -*- coding: utf-8 -*-
from __future__ import annotations

import importlib.util
from dataclasses import dataclass
from typing import Iterable

from .app_config import OLLAMA_BASE_URL, load_ollama_model_name, resolve_tool_command
from .subprocess_utils import run_hidden


@dataclass(frozen=True)
class ToolStatus:
    key: str
    label: str
    state: str
    detail: str = ""

    @property
    def display(self) -> str:
        return f"{self.label} {self.state}"


@dataclass(frozen=True)
class EnvironmentReport:
    statuses: tuple[ToolStatus, ...]
    ollama_models: tuple[str, ...] = ()

    def status_for(self, key: str) -> ToolStatus | None:
        for status in self.statuses:
            if status.key == key:
                return status
        return None


def command_available(command: str) -> bool:
    return resolve_tool_command(command) is not None


def _command_version(command: str) -> str:
    resolved = resolve_tool_command(command)
    if not resolved:
        return ""
    try:
        result = run_hidden(
            [resolved, "-version"],
            capture_output=True,
            text=True,
            timeout=4,
            check=False,
        )
    except Exception:
        return ""
    first_line = (result.stdout or result.stderr).splitlines()
    return first_line[0].strip() if first_line else ""


def _check_command(key: str, label: str, command: str) -> ToolStatus:
    resolved = resolve_tool_command(command)
    if not resolved:
        return ToolStatus(key=key, label=label, state="OFFLINE")
    return ToolStatus(key=key, label=label, state="ONLINE", detail=_command_version(command))


def _check_ollama(base_url: str = OLLAMA_BASE_URL) -> tuple[ToolStatus, tuple[str, ...]]:
    try:
        import requests
    except Exception:
        return ToolStatus("ollama", "OLLAMA", "UNAVAILABLE", "requests import failed"), ()

    model_name = load_ollama_model_name()
    payload = {
        "model": model_name,
        "prompt": "Return only: ok",
        "stream": False,
        "options": {
            "num_predict": 4,
            "temperature": 0,
        },
    }
    try:
        response = requests.post(f"{base_url.rstrip('/')}/api/generate", json=payload, timeout=12)
        if response.status_code != 200:
            return ToolStatus("ollama", "OLLAMA", "LOCAL OFFLINE", f"{model_name}: HTTP {response.status_code}"), ()
        return ToolStatus("ollama", "OLLAMA", "LOCAL READY", model_name), (model_name,)
    except Exception as exc:
        return ToolStatus("ollama", "OLLAMA", "LOCAL OFFLINE", f"{model_name}: {exc}"), ()


def _check_audiocraft() -> ToolStatus:
    if importlib.util.find_spec("audiocraft") is None:
        return ToolStatus("musicgen", "MUSICGEN", "UNAVAILABLE")
    return ToolStatus("musicgen", "MUSICGEN", "IMPORT READY")


def _check_cuda() -> ToolStatus:
    if importlib.util.find_spec("torch") is None:
        return ToolStatus("cuda", "CUDA CHECK", "SKIPPED", "torch not installed")
    try:
        import torch

        if torch.cuda.is_available():
            device_name = torch.cuda.get_device_name(0)
            return ToolStatus("cuda", "CUDA", "AVAILABLE", device_name)
        return ToolStatus("cuda", "CUDA", "UNAVAILABLE")
    except Exception as exc:
        return ToolStatus("cuda", "CUDA CHECK", "SKIPPED", str(exc))


def _check_uvr_candidates(candidates: Iterable[str] = ("uvr", "uvr5", "ultimatevocalremover")) -> ToolStatus:
    for command in candidates:
        if command_available(command):
            return ToolStatus("uvr", "UVR", "FOUND", command)
    return ToolStatus("uvr", "UVR", "CHECK ONLY")


def check_environment(base_url: str = OLLAMA_BASE_URL) -> EnvironmentReport:
    ffmpeg = _check_command("ffmpeg", "FFMPEG", "ffmpeg")
    ffprobe = _check_command("ffprobe", "FFPROBE", "ffprobe")
    ollama, models = _check_ollama(base_url)
    statuses = (
        ffmpeg,
        ffprobe,
        ollama,
        _check_audiocraft(),
        _check_cuda(),
        _check_uvr_candidates(),
    )
    return EnvironmentReport(statuses=statuses, ollama_models=models)
