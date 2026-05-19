from __future__ import annotations

import re
from dataclasses import dataclass


HEAT_TERMS = (
    "熾火",
    "巡り",
    "側に",
    "在る",
    "握らない強さ",
    "火照り",
    "解ける",
    "静かな青",
    "還す",
    "未完",
    "未達成",
    "DAKE",
    "BRAINZ",
    "Codex",
    "JapanMemoryLane",
    "実務直感",
    "違和感",
    "判断疲れ",
)

HASHTAG_HEAT_TERMS = {
    "#aru": "在る",
    "#embers": "熾火",
    "#thought": "実務直感",
    "#brainz": "BRAINZ",
}


@dataclass(frozen=True)
class EmberMetadata:
    heat_tags: str
    temperature: float
    unfinished_score: float
    reignition_score: float
    related_terms: str
    excerpt: str


def default_ember_query() -> str:
    return " ".join(HEAT_TERMS)


def detect_heat_terms(text: str) -> list[str]:
    target = text or ""
    target_lower = target.lower()
    terms: list[str] = []
    for term in HEAT_TERMS:
        if term.lower() in target_lower and term not in terms:
            terms.append(term)
    for hashtag, term in HASHTAG_HEAT_TERMS.items():
        if hashtag in target_lower and term not in terms:
            terms.append(term)
    return terms


def build_ember_metadata(text: str) -> EmberMetadata:
    terms = detect_heat_terms(text)
    unfinished_score = unfinished_score_for(text, terms)
    temperature = min(1.0, 0.18 + (len(terms) * 0.08)) if terms else 0.0
    reignition_score = min(1.0, temperature + (unfinished_score * 0.25))
    terms_text = ", ".join(terms)
    return EmberMetadata(
        heat_tags=terms_text,
        temperature=round(temperature, 3),
        unfinished_score=round(unfinished_score, 3),
        reignition_score=round(reignition_score, 3),
        related_terms=terms_text,
        excerpt=excerpt_for(text),
    )


def unfinished_score_for(text: str, terms: list[str]) -> float:
    target = text or ""
    score = 0.0
    for marker in ("未完", "未達成", "違和感", "判断疲れ", "まだ", "途中", "保留"):
        if marker in target:
            score += 0.18
    if "未完" in terms or "未達成" in terms:
        score += 0.24
    return min(1.0, score)


def excerpt_for(text: str, limit: int = 260) -> str:
    clean = re.sub(r"\s+", " ", (text or "").strip())
    if len(clean) <= limit:
        return clean
    return clean[: limit - 3].rstrip() + "..."
