# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import urllib.error
import urllib.request
from dataclasses import dataclass
from typing import Callable

from app.constants import OLLAMA_BASE_URL


@dataclass(frozen=True)
class OllamaStatus:
    reachable: bool
    models: tuple[str, ...] = ()
    version: str = ""
    error: str = ""


class OllamaClient:
    def __init__(self, base_url: str = OLLAMA_BASE_URL, timeout: float = 5.0) -> None:
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout

    def status(self) -> OllamaStatus:
        try:
            version_payload = self._get_json("/api/version")
            tags_payload = self._get_json("/api/tags")
            models = tuple(
                str(model.get("name") or model.get("model"))
                for model in tags_payload.get("models", [])
                if isinstance(model, dict)
            )
            return OllamaStatus(True, models=models, version=str(version_payload.get("version") or ""))
        except Exception as exc:
            return OllamaStatus(False, error=str(exc))

    def has_model(self, model_name: str) -> bool:
        status = self.status()
        return status.reachable and any(name == model_name or name.startswith(f"{model_name}:") for name in status.models)

    def pull_model(self, model_name: str, progress: Callable[[str], None] | None = None) -> None:
        request = self._post_request("/api/pull", {"model": model_name, "stream": True})
        with urllib.request.urlopen(request, timeout=None) as response:  # noqa: S310
            for raw_line in response:
                if not raw_line:
                    continue
                try:
                    payload = json.loads(raw_line.decode("utf-8"))
                except Exception:
                    continue
                status = str(payload.get("status") or "")
                if progress and status:
                    progress(status)
                if payload.get("error"):
                    raise RuntimeError(str(payload["error"]))

    def generate(self, model_name: str, prompt: str) -> str:
        request = self._post_request(
            "/api/generate",
            {
                "model": model_name,
                "prompt": prompt,
                "stream": False,
                "options": {"temperature": 0.2},
            },
        )
        with urllib.request.urlopen(request, timeout=None) as response:  # noqa: S310
            payload = json.loads(response.read().decode("utf-8"))
        return str(payload.get("response") or "").strip()

    def _get_json(self, path: str) -> dict:
        with urllib.request.urlopen(f"{self.base_url}{path}", timeout=self.timeout) as response:  # noqa: S310
            return json.loads(response.read().decode("utf-8"))

    def _post_request(self, path: str, payload: dict) -> urllib.request.Request:
        data = json.dumps(payload).encode("utf-8")
        return urllib.request.Request(
            f"{self.base_url}{path}",
            data=data,
            headers={"Content-Type": "application/json"},
            method="POST",
        )

