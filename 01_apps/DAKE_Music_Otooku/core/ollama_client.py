# -*- coding: utf-8 -*-
from __future__ import annotations

from .app_config import OLLAMA_BASE_URL, load_ollama_model_name
from .presets import MusicPreset, build_brain_input
from .prompt_builder import MusicDirection, parse_direction_response


def pick_model(available_models: tuple[str, ...]) -> str:
    if available_models:
        return available_models[0]
    return load_ollama_model_name()


def generate_direction(
    user_text: str,
    available_models: tuple[str, ...],
    preset: MusicPreset | None = None,
    base_url: str = OLLAMA_BASE_URL,
    timeout: int = 90,
) -> MusicDirection:
    try:
        import requests
    except Exception as exc:
        raise RuntimeError("requests import failed") from exc

    model = pick_model(available_models)
    system_prompt = (
        "You are a quiet local assistant for a DAKE desktop app called 音を置く. "
        "The user provides a Japanese sound image. Propose a small BGM material direction, "
        "not a finished song. Keep it quiet, ambient, minimal, practical, and easy to place "
        "under video, work scenes, streams, websites, or Shorts. Avoid over-designed ideas, "
        "vocals, famous melodies, copyrighted songs, and artist imitation. "
        "Prefer midnight feeling, air, low temperature, texture, and '置ける音'. "
        "If a selected preset is provided, use it as supplemental atmosphere only. "
        "Prioritize the user's input words over the preset. "
        "Output must be concise. No markdown. No long explanation. "
        "Return exactly these labels, one per section:\n"
        "Mood:\n"
        "BPM:\n"
        "Key:\n"
        "Texture:\n"
        "Instruments:\n"
        "Loop Length:\n"
        "Music Direction:\n"
        "MusicGen Prompt:\n"
        "Usage Idea:"
    )
    prompt = f"{system_prompt}\n\nInput words:\n{build_brain_input(user_text, preset)}"
    payload = {
        "model": model,
        "prompt": prompt,
        "stream": False,
        "options": {
            "temperature": 0.25,
            "num_predict": 420,
        },
    }
    response = requests.post(f"{base_url.rstrip('/')}/api/generate", json=payload, timeout=timeout)
    response.raise_for_status()
    data = response.json()
    body = str(data.get("response", "")).strip()
    if not body:
        raise RuntimeError("Ollama returned an empty response")
    return parse_direction_response(body, user_text, source=f"ollama:{model}")
