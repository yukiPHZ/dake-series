# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import re
from dataclasses import dataclass


@dataclass(frozen=True)
class MusicDirection:
    mood: str
    bpm: str
    key: str
    texture: str
    instrumentation: str
    loop_length: str
    music_direction: str
    musicgen_prompt: str
    negative_notes: str
    usage_idea: str
    source: str = "template"

    def to_direction_text(self, user_text: str) -> str:
        lines = [
            "# 音の設計メモ",
            "",
            f"Input: {user_text}",
            f"Source: {self.source}",
            "",
            "Mood:",
            self.mood,
            "",
            "BPM:",
            self.bpm,
            "",
            "Key:",
            self.key,
            "",
            "Texture:",
            self.texture,
            "",
            "Instruments:",
            self.instrumentation,
            "",
            "Loop Length:",
            self.loop_length,
            "",
            "Music Direction:",
            self.music_direction,
            "",
            "MusicGen Prompt:",
            self.musicgen_prompt,
            "",
            "Usage Idea:",
            self.usage_idea,
            "",
            "Negative Notes:",
            self.negative_notes,
        ]
        return "\n".join(lines).strip() + "\n"

    def to_usage_note(self) -> str:
        lines = [
            "# 使い方メモ",
            "",
            self.usage_idea,
            "",
            "この素材は、動画・配信・サイト・Shorts用のオリジナル音素材として整理する想定です。",
            "既存曲や特定アーティストの完全再現を目的にしないでください。",
            "生成物の利用条件は、使用したモデル・素材・配布条件を確認してください。",
        ]
        return "\n".join(lines).strip() + "\n"


def fallback_direction(user_text: str) -> MusicDirection:
    lowered = user_text.lower()
    quiet_words = ("深夜", "夜", "雨", "静", "余白", "黒", "low", "quiet", "ambient")
    bright_words = ("朝", "休日", "神社", "空", "shorts", "holiday")

    if any(word in lowered for word in quiet_words):
        bpm = "72"
        mood = "quiet industrial ambient"
        key = "D minor / A minor"
        texture = "soft low drone / midnight machine hum"
        instrumentation = "soft pad, low muted bass, light texture, small bell or noise layer"
        music_direction = "深夜の作業音に寄り添う、低温で薄いループ素材。主張を抑えて、映像の下に置ける音。"
    elif any(word in lowered for word in bright_words):
        bpm = "86"
        mood = "quiet morning ambient"
        key = "G major / D major"
        texture = "soft air / distant light percussion"
        instrumentation = "warm pluck, soft marimba, light percussion, airy pad"
        music_direction = "朝の余白を壊さない、軽く明るい短尺BGM素材。"
    else:
        bpm = "80"
        mood = "minimal calm background"
        key = "C major / A minor"
        texture = "gentle pulse / low soft pad"
        instrumentation = "minimal synth pad, soft pulse, simple sub bass, gentle texture"
        music_direction = "作業や説明映像の邪魔をしない、短く置ける背景ループ。"

    clean_text = " ".join(user_text.split())
    prompt = (
        "original quiet ambient minimal background music loop, no vocals, no famous melody, "
        f"{mood}, {texture}, {instrumentation}, {bpm} BPM, seamless 12 second loop, "
        f"practical video background material, concept: {clean_text}"
    )
    return MusicDirection(
        mood=mood,
        bpm=bpm,
        key=key,
        texture=texture,
        instrumentation=instrumentation,
        loop_length="12 seconds",
        music_direction=music_direction,
        musicgen_prompt=prompt,
        negative_notes="no vocals, no copyrighted melody, no artist imitation, no loud lead, no sudden drop, not over-designed",
        usage_idea="深夜のコード作業や静かな制作風景の背景向け。小さな音量で、映像の空気だけを支える使い方に向いています。",
        source="template",
    )


def _extract_json_object(text: str) -> str:
    match = re.search(r"\{.*\}", text.strip(), flags=re.DOTALL)
    if not match:
        raise ValueError("JSON object was not found")
    return match.group(0)


def _value_from_json(data: dict, key: str, default: str, *aliases: str) -> str:
    raw = data.get(key)
    if raw is None:
        for alias in aliases:
            raw = data.get(alias)
            if raw is not None:
                break
    if raw is None:
        raw = default
    if isinstance(raw, (list, tuple)):
        return ", ".join(str(item) for item in raw if item)
    return str(raw).strip() or default


def _parse_sections(text: str) -> dict[str, str]:
    label_map = {
        "mood": "mood",
        "bpm": "bpm",
        "key": "key",
        "texture": "texture",
        "instruments": "instruments",
        "instrumentation": "instruments",
        "loop length": "loop_length",
        "loop_length": "loop_length",
        "music direction": "music_direction",
        "music_direction": "music_direction",
        "musicgen prompt": "musicgen_prompt",
        "musicgen_prompt": "musicgen_prompt",
        "usage idea": "usage_idea",
        "usage_note": "usage_idea",
        "usage idea.": "usage_idea",
    }
    sections: dict[str, list[str]] = {}
    current_key: str | None = None
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        match = re.match(r"^([A-Za-z _]+):\s*(.*)$", line)
        if match:
            label = match.group(1).strip().lower()
            mapped = label_map.get(label)
            if mapped:
                current_key = mapped
                sections.setdefault(current_key, [])
                value = match.group(2).strip()
                if value:
                    sections[current_key].append(value)
                continue
        if current_key:
            sections.setdefault(current_key, []).append(line)

    return {key: " ".join(value).strip() for key, value in sections.items() if value}


def parse_direction_response(text: str, user_text: str, source: str = "ollama") -> MusicDirection:
    fallback = fallback_direction(user_text)
    try:
        data = json.loads(_extract_json_object(text))
        return MusicDirection(
            mood=_value_from_json(data, "mood", fallback.mood),
            bpm=_value_from_json(data, "bpm", fallback.bpm),
            key=_value_from_json(data, "key", fallback.key),
            texture=_value_from_json(data, "texture", fallback.texture),
            instrumentation=_value_from_json(data, "instruments", fallback.instrumentation, "instrumentation"),
            loop_length=_value_from_json(data, "loop_length", fallback.loop_length),
            music_direction=_value_from_json(data, "music_direction", fallback.music_direction),
            musicgen_prompt=_value_from_json(data, "musicgen_prompt", fallback.musicgen_prompt),
            negative_notes=_value_from_json(data, "negative_notes", fallback.negative_notes),
            usage_idea=_value_from_json(data, "usage_note", fallback.usage_idea, "usage_idea"),
            source=source,
        )
    except Exception:
        sections = _parse_sections(text)
        if not sections:
            raise
        return MusicDirection(
            mood=sections.get("mood", fallback.mood),
            bpm=sections.get("bpm", fallback.bpm),
            key=sections.get("key", fallback.key),
            texture=sections.get("texture", fallback.texture),
            instrumentation=sections.get("instruments", fallback.instrumentation),
            loop_length=sections.get("loop_length", fallback.loop_length),
            music_direction=sections.get("music_direction", fallback.music_direction),
            musicgen_prompt=sections.get("musicgen_prompt", fallback.musicgen_prompt),
            negative_notes=fallback.negative_notes,
            usage_idea=sections.get("usage_idea", fallback.usage_idea),
            source=source,
        )

