# -*- coding: utf-8 -*-
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path

from core.scanner import MemoryDocument


HEAT_WORDS = [
    "熾火",
    "巡り",
    "側に",
    "在る",
    "握らない強さ",
    "余白",
    "静か",
    "止まらない",
    "再巡回",
    "熱",
    "補助脳",
    "BRAINZ",
    "BORINEF",
    "DAKE",
    "記憶",
    "痕跡",
    "ゴースト",
    "OIKAWA",
]


@dataclass
class HeatTrace:
    word: str
    score: int
    count: int
    file_count: int
    recent_count: int
    heading_count: int
    filename_count: int
    cooccurring_words: list[str]


@dataclass
class RelatedFragment:
    heat_word: str
    title: str
    path: Path
    relative_path: str
    excerpt: str
    modified_at: datetime
    score: int


@dataclass
class AnalysisResult:
    generated_at: datetime
    traces: list[HeatTrace]
    fragments: list[RelatedFragment]
    suggestion: str
    scanned_files: int
    skipped_files: int


def analyze_documents(
    documents: list[MemoryDocument],
    memory_root: Path,
    skipped_files: int = 0,
    now: datetime | None = None,
) -> AnalysisResult:
    generated_at = now or datetime.now()
    recent_cutoff = generated_at - timedelta(days=7)
    stats = _empty_stats()
    doc_hits: list[tuple[int, str, MemoryDocument]] = []

    for document in documents:
        text = document.text
        headings = _heading_lines(text)
        filename = document.path.name
        present_words: list[str] = []

        for word in HEAT_WORDS:
            body_count = text.count(word)
            heading_count = sum(line.count(word) for line in headings)
            filename_count = filename.count(word)
            if body_count == 0 and heading_count == 0 and filename_count == 0:
                continue

            word_stats = stats[word]
            word_stats["count"] += body_count
            word_stats["heading_count"] += heading_count
            word_stats["filename_count"] += filename_count
            word_stats["files"].add(document.path)
            if document.modified_at >= recent_cutoff:
                word_stats["recent_count"] += 1
            present_words.append(word)

            match_score = body_count + (heading_count * 3) + (filename_count * 5)
            if document.modified_at >= recent_cutoff:
                match_score += 3
            doc_hits.append((match_score, word, document))

        _record_cooccurrence(stats, text, present_words)

    traces = _build_traces(stats)
    fragments = _build_fragments(doc_hits, traces)
    suggestion = _build_suggestion(traces)

    return AnalysisResult(
        generated_at=generated_at,
        traces=traces,
        fragments=fragments,
        suggestion=suggestion,
        scanned_files=len(documents),
        skipped_files=skipped_files,
    )


def _empty_stats() -> dict[str, dict[str, object]]:
    return {
        word: {
            "count": 0,
            "heading_count": 0,
            "filename_count": 0,
            "recent_count": 0,
            "files": set(),
            "cooccurring": set(),
        }
        for word in HEAT_WORDS
    }


def _heading_lines(text: str) -> list[str]:
    return [line.strip() for line in text.splitlines() if line.lstrip().startswith("#")]


def _record_cooccurrence(stats: dict[str, dict[str, object]], text: str, words: list[str]) -> None:
    if len(words) < 2:
        return
    for index, word in enumerate(words):
        cooccurring = stats[word]["cooccurring"]
        assert isinstance(cooccurring, set)
        for other in words[index + 1 :]:
            if _appears_near(text, word, other):
                cooccurring.add(other)
                reverse = stats[other]["cooccurring"]
                assert isinstance(reverse, set)
                reverse.add(word)


def _appears_near(text: str, word: str, other: str, window: int = 180) -> bool:
    start = text.find(word)
    while start >= 0:
        left = max(0, start - window)
        right = min(len(text), start + len(word) + window)
        if other in text[left:right]:
            return True
        start = text.find(word, start + len(word))
    return False


def _build_traces(stats: dict[str, dict[str, object]]) -> list[HeatTrace]:
    traces: list[HeatTrace] = []
    for word, data in stats.items():
        count = int(data["count"])
        heading_count = int(data["heading_count"])
        filename_count = int(data["filename_count"])
        recent_count = int(data["recent_count"])
        files = data["files"]
        cooccurring = data["cooccurring"]
        assert isinstance(files, set)
        assert isinstance(cooccurring, set)

        if count == 0 and heading_count == 0 and filename_count == 0:
            continue

        file_count = len(files)
        score = count + (heading_count * 3) + (filename_count * 5) + (recent_count * 3)
        if file_count >= 2:
            score += 5
        score += min(4, len(cooccurring))

        traces.append(
            HeatTrace(
                word=word,
                score=score,
                count=count,
                file_count=file_count,
                recent_count=recent_count,
                heading_count=heading_count,
                filename_count=filename_count,
                cooccurring_words=sorted(cooccurring)[:5],
            )
        )

    traces.sort(key=lambda item: (item.score, item.file_count, item.count), reverse=True)
    return traces


def _build_fragments(
    doc_hits: list[tuple[int, str, MemoryDocument]],
    traces: list[HeatTrace],
    limit: int = 5,
) -> list[RelatedFragment]:
    top_words = {trace.word for trace in traces[:6]}
    fragments: list[RelatedFragment] = []
    seen: set[tuple[str, str]] = set()

    for score, word, document in sorted(doc_hits, key=lambda item: (item[0], item[2].modified_at), reverse=True):
        if word not in top_words or score <= 0:
            continue
        key = (str(document.path), word)
        if key in seen:
            continue
        seen.add(key)
        fragments.append(
            RelatedFragment(
                heat_word=word,
                title=_fragment_title(document, word),
                path=document.path,
                relative_path=document.relative_path,
                excerpt=_excerpt(document.text, word),
                modified_at=document.modified_at,
                score=score,
            )
        )
        if len(fragments) >= limit:
            break

    return fragments


def _fragment_title(document: MemoryDocument, word: str) -> str:
    date_text = document.modified_at.strftime("%Y-%m-%d")
    title = re.sub(r"\s+", " ", document.title).strip() or document.path.stem
    if word in title:
        return f"{date_text} {title}"
    return f"{date_text} {title}【{word}】"


def _excerpt(text: str, word: str, width: int = 120) -> str:
    compact = re.sub(r"\s+", " ", text).strip()
    if not compact:
        return ""
    index = compact.find(word)
    if index < 0:
        return compact[:width]
    left = max(0, index - width // 2)
    right = min(len(compact), index + len(word) + width // 2)
    excerpt = compact[left:right].strip()
    if left > 0:
        excerpt = "..." + excerpt
    if right < len(compact):
        excerpt = excerpt + "..."
    return excerpt


def _build_suggestion(traces: list[HeatTrace]) -> str:
    if not traces:
        return (
            "強い熱語の再出現はまだ見つかっていません。\n\n"
            "BRAINZは静かな記憶層として保たれています。次の巡回で、別の痕跡が浮上する可能性があります。"
        )

    words = "・".join(trace.word for trace in traces[:3])
    if any(trace.word in {"BRAINZ", "記憶", "再巡回", "痕跡"} for trace in traces[:5]):
        return (
            f"{words}の系統が再出現しています。\n\n"
            "BRAINZは検索機能だけでなく、記憶を再巡回するUIとして育てる余地があります。"
        )

    return (
        f"{words}の系統が再出現しています。\n\n"
        "いまは断片を急いで結論にせず、近い記憶を並べて眺める段階がよさそうです。"
    )
