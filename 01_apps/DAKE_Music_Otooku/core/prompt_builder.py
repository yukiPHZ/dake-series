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
    instrumentation: str
    loop_length: str
    musicgen_prompt: str
    negative_notes: str
    usage_idea: str
    source: str = "template"

    def to_direction_text(self, user_text: str) -> str:
        lines = [
            "# 音の設計メモ",
            "",
            f"入力: {user_text}",
            f"source: {self.source}",
            "",
            f"mood: {self.mood}",
            f"BPM: {self.bpm}",
            f"key: {self.key}",
            f"instruments: {self.instrumentation}",
            f"loop length: {self.loop_length}",
            "",
            "musicgen prompt:",
            self.musicgen_prompt,
            "",
            "negative notes:",
            self.negative_notes,
            "",
            "usage note:",
            self.usage_idea,
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
        mood = "静かな稼働感、低い温度、余白がある"
        key = "D minor / A minor"
        instrumentation = "soft pad, low muted bass, light texture, small bell or noise layer"
    elif any(word in lowered for word in bright_words):
        bpm = "86"
        mood = "明るいが控えめ、朝の空気、近すぎない距離"
        key = "G major / D major"
        instrumentation = "warm pluck, soft marimba, light percussion, airy pad"
    else:
        bpm = "80"
        mood = "落ち着いた背景、作業の邪魔をしない、短いループ向き"
        key = "C major / A minor"
        instrumentation = "minimal synth pad, soft pulse, simple sub bass, gentle texture"

    clean_text = " ".join(user_text.split())
    prompt = (
        "original short background music loop, no vocals, no famous melody, "
        f"{mood}, {instrumentation}, {bpm} BPM, seamless 12 second loop, "
        f"made for video background material, concept: {clean_text}"
    )
    return MusicDirection(
        mood=mood,
        bpm=bpm,
        key=key,
        instrumentation=instrumentation,
        loop_length="12 seconds",
        musicgen_prompt=prompt,
        negative_notes="no vocals, no copyrighted melody, no artist imitation, no loud lead, no sudden drop",
        usage_idea="短い動画の背景、待機画面、サイトの空気付け、配信用の低音量ループに向いています。",
        source="template",
    )


def _extract_json_object(text: str) -> str:
    stripped = text.strip()
    match = re.search(r"\{.*\}", stripped, flags=re.DOTALL)
    if not match:
        raise ValueError("JSON object was not found")
    return match.group(0)


def parse_direction_response(text: str, user_text: str, source: str = "ollama") -> MusicDirection:
    data = json.loads(_extract_json_object(text))
    fallback = fallback_direction(user_text)

    def value(key: str, default: str, *aliases: str) -> str:
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

    return MusicDirection(
        mood=value("mood", fallback.mood),
        bpm=value("bpm", fallback.bpm),
        key=value("key", fallback.key),
        instrumentation=value("instruments", fallback.instrumentation, "instrumentation"),
        loop_length=value("loop_length", fallback.loop_length),
        musicgen_prompt=value("musicgen_prompt", fallback.musicgen_prompt),
        negative_notes=value("negative_notes", fallback.negative_notes),
        usage_idea=value("usage_note", fallback.usage_idea, "usage_idea"),
        source=source,
    )

