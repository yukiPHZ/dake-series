from __future__ import annotations

from pathlib import Path
from typing import Sequence

from core.app_config import exports_dir, now_iso
from core.db import SearchResult


CHATGPT_SOURCE_TYPE = "chatgpt_export"


def timestamp_for_file() -> str:
    return now_iso().replace("-", "").replace(":", "").replace("T", "_")[:13]


def write_chatgpt_handoff(query: str, results: Sequence[SearchResult], export_dir: Path | None = None) -> Path:
    export_dir = export_dir or exports_dir()
    export_dir.mkdir(parents=True, exist_ok=True)
    path = export_dir / f"chatgpt_handoff_{timestamp_for_file()}.md"
    path.write_text(build_chatgpt_handoff(query, results), encoding="utf-8")
    return path


def write_codex_handoff(query: str, results: Sequence[SearchResult], export_dir: Path | None = None) -> Path:
    export_dir = export_dir or exports_dir()
    export_dir.mkdir(parents=True, exist_ok=True)
    path = export_dir / f"codex_handoff_{timestamp_for_file()}.md"
    path.write_text(build_codex_handoff(query, results), encoding="utf-8")
    return path


def build_chatgpt_handoff(query: str, results: Sequence[SearchResult]) -> str:
    chatgpt_results, file_results = split_chatgpt_results(results)
    lines = [
        "# ChatGPT用まとめ",
        "",
        f"- 検索クエリ: {query}",
        f"- 生成日時: {now_iso()}",
        "",
    ]
    if chatgpt_results:
        lines.extend(
            [
                "## 過去のChatGPT会話から関連する内容",
                chatgpt_summary(chatgpt_results),
                "",
            ]
        )
    lines.extend(
        [
            "## 関連ファイル",
            *result_lines(file_results or results),
            "",
            "## 要約",
            summarize_results(results),
            "",
            "## 次にChatGPTへ聞くための文章",
            "以下の過去メモ・仕様・ChatGPT会話を踏まえて、次に決めるべき論点を整理してください。必要なら選択肢と判断材料を分けて提案してください。",
            "",
            "## 抜粋",
            excerpts(results),
            "",
        ]
    )
    return "\n".join(lines)


def build_codex_handoff(query: str, results: Sequence[SearchResult]) -> str:
    chatgpt_results, file_results = split_chatgpt_results(results)
    lines = [
        "# Codex用素材",
        "",
        f"- 検索クエリ: {query}",
        f"- 生成日時: {now_iso()}",
        "",
        "## 関連仕様",
    ]
    lines.extend(result_lines(file_results or results))
    if chatgpt_results:
        lines.extend(
            [
                "",
                "## 仕様検討メモ",
                chatgpt_summary(chatgpt_results),
            ]
        )
    lines.extend(
        [
            "",
            "## 関連README",
            related_by_keyword(results, "readme"),
            "",
            "## 修正対象候補",
            related_targets(results),
            "",
            "## Codexに渡す指示素材",
            "下記の関連資料と仕様検討メモを前提に、既存方針を壊さず、対象ファイルの差分を小さく保って実装してください。自動送信は行わず、この素材を必要に応じて貼り付けて使います。",
            "",
            "## 抜粋",
            excerpts(results),
            "",
        ]
    )
    return "\n".join(lines)


def split_chatgpt_results(results: Sequence[SearchResult]) -> tuple[list[SearchResult], list[SearchResult]]:
    chatgpt_results = [result for result in results if result.source_type == CHATGPT_SOURCE_TYPE]
    file_results = [result for result in results if result.source_type != CHATGPT_SOURCE_TYPE]
    return chatgpt_results, file_results


def result_lines(results: Sequence[SearchResult]) -> list[str]:
    if not results:
        return ["- なし"]
    return [result_line(result) for result in results[:10]]


def result_line(result: SearchResult) -> str:
    if result.source_type == CHATGPT_SOURCE_TYPE:
        return (
            f"- {result.conversation_title or result.title} ({result.source_type} / {result.role})\n"
            f"  - conversation_id: {result.conversation_id}\n"
            f"  - message_index: {result.message_index}"
        )
    return f"- {result.title} ({result.source_type})\n  - {result.path}"


def chatgpt_summary(results: Sequence[SearchResult]) -> str:
    if not results:
        return "なし"
    lines = []
    for result in results[:8]:
        title = result.conversation_title or result.title
        lines.append(
            "\n".join(
                [
                    f"- 会話タイトル: {title}",
                    f"  - role: {result.role or 'unknown'}",
                    f"  - source_type: {result.source_type}",
                    f"  - conversation_id: {result.conversation_id or result.path}",
                    f"  - 抜粋: {compact(result.snippet or result.content, 260)}",
                ]
            )
        )
    return "\n".join(lines)


def summarize_results(results: Sequence[SearchResult]) -> str:
    if not results:
        return "関連する記憶はまだ見つかっていません。"
    snippets = []
    for result in results[:5]:
        label = result.conversation_title if result.source_type == CHATGPT_SOURCE_TYPE else result.title
        suffix = f" / {result.role}" if result.role else ""
        snippets.append(f"- {label}{suffix}: {compact(result.snippet or result.content, 220)}")
    return "\n".join(snippets)


def excerpts(results: Sequence[SearchResult]) -> str:
    if not results:
        return "なし"
    blocks = []
    for result in results[:5]:
        heading = result.conversation_title or result.title
        metadata = []
        if result.source_type:
            metadata.append(f"source_type: {result.source_type}")
        if result.role:
            metadata.append(f"role: {result.role}")
        if result.conversation_id:
            metadata.append(f"conversation_id: {result.conversation_id}")
        meta_text = "\n".join(metadata)
        body = compact(result.content, 900)
        blocks.append(f"### {heading}\n\n```text\n{meta_text}\n\n{body}\n```")
    return "\n\n".join(blocks)


def related_by_keyword(results: Sequence[SearchResult], keyword: str) -> str:
    matched = [result for result in results if keyword.lower() in result.path.lower() or keyword.lower() in result.title.lower()]
    if not matched:
        return "検索結果内にREADMEらしいファイルは見つかっていません。"
    return "\n".join(f"- {result.path}" for result in matched[:8])


def related_targets(results: Sequence[SearchResult]) -> str:
    if not results:
        return "未定"
    candidates = []
    for result in results[:8]:
        if result.source_type in {"md", "txt", "json"}:
            candidates.append(f"- {result.path}")
    return "\n".join(candidates) if candidates else "未定"


def compact(text: str, limit: int) -> str:
    clean = " ".join(text.replace("\r", "\n").split())
    if len(clean) <= limit:
        return clean
    return clean[: max(0, limit - 3)] + "..."
