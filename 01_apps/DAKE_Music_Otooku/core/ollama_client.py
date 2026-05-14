# -*- coding: utf-8 -*-
from __future__ import annotations

from .app_config import OLLAMA_BASE_URL, load_ollama_model_name
from .prompt_builder import MusicDirection, parse_direction_response


def pick_model(available_models: tuple[str, ...]) -> str:
    if available_models:
        return available_models[0]
    return load_ollama_model_name()


def generate_direction(
    user_text: str,
    available_models: tuple[str, ...],
    base_url: str = OLLAMA_BASE_URL,
    timeout: int = 24,
) -> MusicDirection:
    try:
        import requests
    except Exception as exc:
        raise RuntimeError("requests import failed") from exc

    model = pick_model(available_models)
    system_prompt = (
        "You are a local assistant for a DAKE desktop app called 音を置く. "
        "Create a compact sound-material direction from the user's Japanese words. "
        "This is not a DAW and not a full song request. "
        "Return only one JSON object with keys: mood, bpm, key, instruments, "
        "loop_length, musicgen_prompt, usage_note. "
        "Avoid copyrighted songs, famous artists, imitation, vocals, and publishing advice. "
        "The musicgen_prompt must be in concise English for an original short BGM loop."
    )
    prompt = f"{system_prompt}\n\nUser words:\n{user_text.strip()}"
    payload = {
        "model": model,
        "prompt": prompt,
        "stream": False,
        "options": {
            "temperature": 0.35,
            "num_predict": 500,
        },
    }
    response = requests.post(f"{base_url.rstrip('/')}/api/generate", json=payload, timeout=timeout)
    response.raise_for_status()
    data = response.json()
    body = str(data.get("response", "")).strip()
    if not body:
        raise RuntimeError("Ollama returned an empty response")
    return parse_direction_response(body, user_text, source=f"ollama:{model}")

