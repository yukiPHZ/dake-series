from __future__ import annotations

import json
from pathlib import Path
from typing import Sequence

from core.app_config import exports_dir, now_iso
from core.db import SearchResult


CHATGPT_SOURCE_TYPE = "chatgpt_export"
CODEX_SOURCE_TYPE = "codex_result"


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
    chatgpt_results, codex_results, file_results = split_results(results)
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
    if codex_results:
        lines.extend(
            [
                "## 過去のCodex実装結果",
                codex_summary_for_chatgpt(codex_results),
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
            "以下の過去メモ・仕様・ChatGPT会話・Codex実装履歴を踏まえて、次に決めるべき論点を整理してください。必要なら選択肢と判断材料を分けて提案してください。",
            "",
            "## 抜粋",
            excerpts(results),
            "",
        ]
    )
    return "\n".join(lines)


def build_codex_handoff(query: str, results: Sequence[SearchResult]) -> str:
    chatgpt_results, codex_results, file_results = split_results(results)
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
    if codex_results:
        lines.extend(
            [
                "",
                "## 既存実装履歴",
                codex_summary_for_codex(codex_results),
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
            "下記の関連資料、仕様検討メモ、既存実装履歴を前提に、既存方針を壊さず、対象ファイルの差分を小さく保って実装してください。自動送信は行わず、この素材を必要に応じて貼り付けて使います。",
            "",
            "## 抜粋",
            excerpts(results),
            "",
        ]
    )
    return "\n".join(lines)


def split_results(results: Sequence[SearchResult]) -> tuple[list[SearchResult], list[SearchResult], list[SearchResult]]:
    chatgpt_results = [result for result in results if result.source_type == CHATGPT_SOURCE_TYPE]
    codex_results = [result for result in results if result.source_type == CODEX_SOURCE_TYPE]
    file_results = [result for result in results if result.source_type not in {CHATGPT_SOURCE_TYPE, CODEX_SOURCE_TYPE}]
    return chatgpt_results, codex_results, file_results


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
    if result.source_type == CODEX_SOURCE_TYPE:
        return (
            f"- {result.title} ({result.source_type})\n"
            f"  - commit_hash: {result.commit_hash}\n"
            f"  - changed_files: {', '.join(json_list(result.changed_files_json)[:6])}"
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


def codex_summary_for_chatgpt(results: Sequence[SearchResult]) -> str:
    if not results:
        return "なし"
    lines = []
    for result in results[:8]:
        lines.append(
            "\n".join(
                [
                    f"- title: {result.title}",
                    f"  - summary: {compact(result.codex_summary or result.snippet or result.content, 260)}",
                    f"  - changed_files: {', '.join(json_list(result.changed_files_json)[:8])}",
                    f"  - test_results: {compact(result.test_results, 220)}",
                    f"  - commit_hash: {result.commit_hash}",
                    f"  - push_result: {result.push_result}",
                    f"  - 注意点: {compact(result.phase_notes or result.git_status, 220)}",
                ]
            )
        )
    return "\n".join(lines)


def codex_summary_for_codex(results: Sequence[SearchResult]) -> str:
    if not results:
        return "なし"
    lines = []
    for result in results[:8]:
        lines.append(
            "\n".join(
                [
                    f"- 直近commit: {result.commit_hash or '未検出'}",
                    f"  - title: {result.title}",
                    f"  - 変更済みファイル: {', '.join(json_list(result.changed_files_json)[:10])}",
                    f"  - 確認済みテスト: {compact(result.test_results or result.build_results, 260)}",
                    f"  - 未実装Phase候補: {compact(result.phase_notes, 220)}",
                    f"  - 次に触ってよい範囲: {', '.join(json_list(result.changed_files_json)[:5]) or '検索結果の関連範囲'}",
                ]
            )
        )
    return "\n".join(lines)


def summarize_results(results: Sequence[SearchResult]) -> str:
    if not results:
        return "関連する記憶はまだ見つかっていません。"
    snippets = []
    for result in results[:5]:
        if result.source_type == CHATGPT_SOURCE_TYPE:
            label = result.conversation_title or result.title
            suffix = f" / {result.role}" if result.role else ""
        elif result.source_type == CODEX_SOURCE_TYPE:
            label = result.title
            suffix = f" / {result.commit_hash[:7]}" if result.commit_hash else ""
        else:
            label = result.title
            suffix = ""
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
        if result.commit_hash:
            metadata.append(f"commit_hash: {result.commit_hash}")
        if result.changed_files_json:
            metadata.append(f"changed_files: {', '.join(json_list(result.changed_files_json)[:8])}")
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
        elif result.source_type == CODEX_SOURCE_TYPE:
            candidates.extend(json_list(result.changed_files_json)[:8])
    unique = []
    for candidate in candidates:
        if candidate and candidate not in unique:
            unique.append(candidate)
    return "\n".join(f"- {candidate}" for candidate in unique) if unique else "未定"


def json_list(value: str) -> list[str]:
    try:
        data = json.loads(value or "[]")
    except json.JSONDecodeError:
        return []
    if not isinstance(data, list):
        return []
    return [str(item) for item in data if str(item).strip()]


def compact(text: str, limit: int) -> str:
    clean = " ".join((text or "").replace("\r", "\n").split())
    if len(clean) <= limit:
        return clean
    return clean[: max(0, limit - 3)] + "..."
