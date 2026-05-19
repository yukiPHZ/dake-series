# -*- coding: utf-8 -*-
from __future__ import annotations

from pathlib import Path

from core.heat_engine import AnalysisResult


def write_suggestion(memory_root: Path, result: AnalysisResult) -> Path:
    output_dir = memory_root / "OIKAWA" / "suggestions"
    output_dir.mkdir(parents=True, exist_ok=True)
    file_name = result.generated_at.strftime("%Y%m%d_%H%M%S_oikawa.md")
    output_path = output_dir / file_name
    output_path.write_text(render_suggestion_markdown(result), encoding="utf-8")
    return output_path


def render_suggestion_markdown(result: AnalysisResult) -> str:
    lines: list[str] = [
        "# OIKAWA 観測記録",
        "",
        f"- generated_at: {result.generated_at.strftime('%Y-%m-%d %H:%M:%S')}",
        "- source: DAKE_Brainz_OIKAWA",
        "- mode: local_scan",
        "",
        "## 浮上した痕跡",
        "",
    ]

    if result.traces:
        for trace in result.traces[:8]:
            lines.append(f"- {trace.word}")
    else:
        lines.append("- 強い熱の痕跡は見つかりませんでした")

    lines.extend(["", "## 関連断片", ""])
    if result.fragments:
        for index, fragment in enumerate(result.fragments, start=1):
            lines.extend(
                [
                    f"### {index}. {fragment.title}",
                    "",
                    "該当ファイル：",
                    f"`{fragment.path}`",
                    "",
                    "抜粋：",
                    f"> {fragment.excerpt}",
                    "",
                ]
            )
    else:
        lines.extend(["関連断片はまだ浮上していません。", ""])

    lines.extend(
        [
            "## OIKAWA提案",
            "",
            result.suggestion,
            "",
            "## 注意",
            "",
            "この提案はローカル記憶の機械的な巡回結果です。",
            "断定ではありません。",
            "",
        ]
    )
    return "\n".join(lines)
