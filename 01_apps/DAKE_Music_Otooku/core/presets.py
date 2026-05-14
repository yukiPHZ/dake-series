# -*- coding: utf-8 -*-
from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path

from .app_config import PRESETS_DIR
from .prompt_builder import MusicDirection, fallback_direction


MUSIC_PRESETS_PATH = PRESETS_DIR / "music_presets.json"


@dataclass(frozen=True)
class MusicPreset:
    name: str
    mood: str
    bpm: str
    texture: str
    tags: tuple[str, ...]
    use: str

    def to_prompt_context(self) -> str:
        return "\n".join(
            [
                f"Selected Preset: {self.name}",
                f"Preset Mood: {self.mood}",
                f"Preset Texture: {self.texture}",
                f"Preset Tags: {', '.join(self.tags)}",
                f"Preset Use: {self.use}",
            ]
        )

    def to_note_lines(self) -> list[str]:
        return [
            "Preset:",
            self.name,
            "",
            "Preset Mood:",
            self.mood,
            "",
            "Preset Texture:",
            self.texture,
            "",
            "Preset Tags:",
            ", ".join(self.tags),
            "",
            "Preset Use:",
            self.use,
        ]


DEFAULT_PRESETS = (
    MusicPreset(
        name="BORINEF",
        mood="ember / quiet / low heat",
        bpm="60-72",
        texture="low drone, soft noise, distant warmth",
        tags=("borinef", "quiet", "ember"),
        use="静かな余熱、夜、内省、長尺背景",
    ),
    MusicPreset(
        name="holiday-jinja",
        mood="shrine / morning / air",
        bpm="64-80",
        texture="light bell, air pad, soft field ambience",
        tags=("shrine", "quiet", "morning"),
        use="神社、空、朝、短い映像",
    ),
    MusicPreset(
        name="YUKIZ稼働中",
        mood="work / code / sewing / midnight",
        bpm="68-88",
        texture="machine hum, soft beat, quiet focus",
        tags=("work", "midnight", "yukiz"),
        use="作業動画、配信、Shorts、制作ログ",
    ),
    MusicPreset(
        name="quiet work",
        mood="focus / simple / clean",
        bpm="72-92",
        texture="minimal pulse, soft pad",
        tags=("work", "quiet"),
        use="作業用BGM、説明動画、デスク作業",
    ),
    MusicPreset(
        name="blue memory",
        mood="blue / memory / slow",
        bpm="56-70",
        texture="soft piano, airy pad, distance",
        tags=("blue", "memory", "calm"),
        use="Japan Memory Lane、静かな写真、余韻",
    ),
)


def _preset_from_dict(data: dict) -> MusicPreset:
    raw_tags = data.get("tags", ())
    if isinstance(raw_tags, str):
        tags = tuple(tag.strip() for tag in raw_tags.split(",") if tag.strip())
    else:
        tags = tuple(str(tag).strip() for tag in raw_tags if str(tag).strip())
    return MusicPreset(
        name=str(data.get("name", "")).strip(),
        mood=str(data.get("mood", "")).strip(),
        bpm=str(data.get("bpm", "")).strip(),
        texture=str(data.get("texture", "")).strip(),
        tags=tags,
        use=str(data.get("use", "")).strip(),
    )


def load_music_presets(path: Path = MUSIC_PRESETS_PATH) -> tuple[MusicPreset, ...]:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        presets = tuple(
            preset
            for preset in (_preset_from_dict(item) for item in data.get("presets", ()))
            if preset.name
        )
        if presets:
            return presets
    except Exception:
        pass
    return DEFAULT_PRESETS


def find_preset(name: str, presets: tuple[MusicPreset, ...] | None = None) -> MusicPreset | None:
    for preset in presets or load_music_presets():
        if preset.name == name:
            return preset
    return None


def build_brain_input(user_text: str, preset: MusicPreset | None) -> str:
    clean_text = user_text.strip()
    if not preset:
        return clean_text
    return "\n\n".join(
        [
            clean_text,
            preset.to_prompt_context(),
            "Preset note: prioritize the user's input words. The preset is only an air memo, not a fixed command.",
        ]
    ).strip()


def fallback_direction_with_preset(user_text: str, preset: MusicPreset | None) -> MusicDirection:
    base = fallback_direction(user_text)
    if not preset:
        return base
    musicgen_prompt = (
        f"{base.musicgen_prompt}, preset mood: {preset.mood}, texture: {preset.texture}, "
        f"tags: {', '.join(preset.tags)}, use: {preset.use}"
    )
    return MusicDirection(
        mood=preset.mood or base.mood,
        bpm=preset.bpm or base.bpm,
        key=base.key,
        texture=preset.texture or base.texture,
        instrumentation=base.instrumentation,
        loop_length=base.loop_length,
        music_direction=f"{base.music_direction}\nPreset air memo: {preset.use}",
        musicgen_prompt=musicgen_prompt,
        negative_notes=base.negative_notes,
        usage_idea=preset.use or base.usage_idea,
        source=f"template:{preset.name}",
    )


def merge_preset_tags(tags: tuple[str, ...], preset: MusicPreset | None) -> tuple[str, ...]:
    merged: list[str] = []
    if preset:
        merged.extend(preset.tags)
    merged.extend(tags)

    clean_tags: list[str] = []
    seen: set[str] = set()
    for tag in merged:
        value = tag.strip()
        if not value or value.lower() in seen:
            continue
        clean_tags.append(value)
        seen.add(value.lower())
    return tuple(clean_tags or ("quiet",))
