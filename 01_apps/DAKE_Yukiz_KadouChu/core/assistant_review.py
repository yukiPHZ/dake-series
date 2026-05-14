from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

from core.app_config import LOG_TEXT
from core.ollama_client import generate_ollama_text
from core.posting_package import packages_dir

LogCallback = Callable[[str], None]

REVIEW_FILE_NAME = "assistant_review.md"


def find_latest_package_dir() -> Path | None:
    root = packages_dir()
    if not root.exists():
        return None
    candidates = [path for path in root.iterdir() if path.is_dir()]
    if not candidates:
        return None
    return max(candidates, key=lambda path: path.stat().st_mtime)


def _read_text(path: Path, limit: int | None = None) -> str:
    if not path.exists() or not path.is_file():
        return ""
    text = path.read_text(encoding="utf-8", errors="replace").strip()
    if limit is not None and len(text) > limit:
        return text[:limit].rstrip() + "\n..."
    return text


def _read_json(path: Path) -> Any:
    if not path.exists() or not path.is_file():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8", errors="replace"))
    except Exception:
        return None


def _sample_long_text(text: str, chunk_size: int = 900, limit: int = 3000) -> str:
    text = text.strip()
    if len(text) <= limit:
        return text
    middle_start = max(0, (len(text) // 2) - (chunk_size // 2))
    parts = [
        "[start]\n" + text[:chunk_size].strip(),
        "[middle]\n" + text[middle_start : middle_start + chunk_size].strip(),
        "[end]\n" + text[-chunk_size:].strip(),
    ]
    return "\n\n---\n\n".join(parts)


def _read_package_context(package_dir: Path) -> dict[str, Any]:
    transcript = _read_text(package_dir / "transcript.txt")
    transcript_unavailable = _read_text(package_dir / "transcript_unavailable.txt", limit=1000)
    shorts = _read_json(package_dir / "shorts_candidates.json")
    media_info = _read_json(package_dir / "media_info.json")
    media_unavailable = _read_json(package_dir / "media_info_unavailable.json")
    metadata_dir = package_dir / "metadata"

    paths = [
        package_dir / "transcript.txt",
        package_dir / "transcript_unavailable.txt",
        package_dir / "shorts_candidates.json",
        metadata_dir / "title_ideas.txt",
        metadata_dir / "description_draft.txt",
        metadata_dir / "tags.txt",
        metadata_dir / "upload_notes.txt",
        package_dir / "media_info.json",
        package_dir / "media_info_unavailable.json",
    ]
    files_read = [str(path.relative_to(package_dir)) for path in paths if path.exists()]
    missing_files = [str(path.relative_to(package_dir)) for path in paths if not path.exists()]

    return {
        "package_name": package_dir.name,
        "transcript": transcript,
        "transcript_excerpt": _sample_long_text(transcript),
        "transcript_unavailable": transcript_unavailable,
        "shorts_candidates": shorts if isinstance(shorts, list) else [],
        "title_ideas": _read_text(metadata_dir / "title_ideas.txt", limit=2000),
        "description_draft": _read_text(metadata_dir / "description_draft.txt", limit=2200),
        "tags": _read_text(metadata_dir / "tags.txt", limit=1200),
        "upload_notes": _read_text(metadata_dir / "upload_notes.txt", limit=1600),
        "media_info": media_info if isinstance(media_info, dict) else None,
        "media_unavailable": media_unavailable if isinstance(media_unavailable, dict) else None,
        "files_read": files_read,
        "missing_files": missing_files,
    }


def _title_lines(text: str) -> list[str]:
    lines: list[str] = []
    for raw in text.splitlines():
        line = raw.strip().lstrip("-").strip()
        if line:
            lines.append(line)
    return lines[:3]


def _candidate_summary(candidate: dict[str, Any]) -> str:
    start = str(candidate.get("start") or "--")
    end = str(candidate.get("end") or "--")
    reason = str(candidate.get("reason") or "補助脳：候補として確認できます。")
    return f"- {start} - {end}: {reason}"


def _fallback_review(context: dict[str, Any]) -> str:
    media = context.get("media_info") or {}
    duration = media.get("duration_timecode") or media.get("duration") or "unknown"
    transcript = str(context.get("transcript") or "")
    transcript_state = "文字起こしがあります。" if transcript else "文字起こしは未取得です。"
    shorts = context.get("shorts_candidates") if isinstance(context.get("shorts_candidates"), list) else []
    titles = _title_lines(str(context.get("title_ideas") or ""))
    if not titles:
        titles = ["稼働中。夜に作る。", "止まらず作る。", "静かな作業机。"]

    recommended = []
    for item in shorts[:3]:
        if isinstance(item, dict):
            recommended.append(_candidate_summary(item))
    if not recommended:
        recommended.append("- Shorts候補は未生成です。先に投稿パッケージ生成を確認してください。")

    title_direction = "\n".join(f"- {title}" for title in titles[:3])
    return (
        "# 補助脳レビュー\n\n"
        "## Summary\n"
        f"{context['package_name']} の投稿前レビューです。動画尺は {duration}。{transcript_state} "
        "投稿用メタデータとShorts候補を確認できる状態です。\n\n"
        "## Atmosphere\n"
        "- quiet work\n"
        "- calm process\n"
        "- production base\n\n"
        "## Recommended Shorts\n"
        + "\n".join(recommended)
        + "\n\n"
        "## Title Direction\n"
        + title_direction
        + "\n\n"
        "## Description Notes\n"
        "- PEAKHEADZ / DAKE / GitHub のリンク欄を公開前に確認してください。\n"
        "- 動画内で何を進めたか、1文だけ足すと見やすくなります。\n"
        "- Shorts候補を採用する場合は、切り出し開始位置を人間の目で確認してください。\n\n"
        "## Before Publish\n"
        "- サムネ確認\n"
        "- タイトル確認\n"
        "- BGM未適用\n"
        "- Shorts候補は自動抽出のため要確認\n"
        "- 自動公開はしていません\n\n"
        "## Assistant Note\n"
        "出せる形にはなっています。最後だけ、菊田さんが握ってください。\n"
    )


def _ollama_prompt(context: dict[str, Any]) -> str:
    payload = {
        "package_name": context["package_name"],
        "media_info": context.get("media_info"),
        "media_unavailable": context.get("media_unavailable"),
        "transcript_excerpt": context.get("transcript_excerpt"),
        "transcript_unavailable": context.get("transcript_unavailable"),
        "shorts_candidates": context.get("shorts_candidates")[:5],
        "title_ideas": context.get("title_ideas"),
        "description_draft": context.get("description_draft"),
        "tags": context.get("tags"),
        "upload_notes": context.get("upload_notes"),
    }
    return (
        "You are the local assistant brain for Dakeユキズ稼働中.\n"
        "Create a concise Japanese Markdown review for a YouTube posting package.\n"
        "The assistant only proposes. The human makes the final decision.\n"
        "Do not claim upload is automatic. Do not use external API assumptions.\n"
        "Use exactly these section headings:\n"
        "# 補助脳レビュー\n"
        "## Summary\n"
        "## Atmosphere\n"
        "## Recommended Shorts\n"
        "## Title Direction\n"
        "## Description Notes\n"
        "## Before Publish\n"
        "## Assistant Note\n\n"
        "Recommended Shorts should pick 1 to 3 candidates if candidates exist.\n"
        "Atmosphere may use short English tags like quiet work or late night build.\n"
        "Assistant Note should be short and calm.\n\n"
        "Package context JSON:\n"
        f"{json.dumps(payload, ensure_ascii=False, indent=2)}"
    )


def _normalize_review(markdown: str, fallback: str) -> str:
    text = markdown.strip()
    if "# 補助脳レビュー" not in text:
        return fallback
    if not text.startswith("# 補助脳レビュー"):
        text = text[text.find("# 補助脳レビュー") :].strip()
    return text.rstrip() + "\n"


def run_assistant_review(
    package_dir: Path,
    ollama_ready: bool,
    log: LogCallback | None = None,
) -> dict[str, Any]:
    if not package_dir.exists() or not package_dir.is_dir():
        raise FileNotFoundError(str(package_dir))

    def emit(message: str) -> None:
        if log:
            log(message)

    emit(LOG_TEXT["review_package_read"])
    context = _read_package_context(package_dir)
    emit(LOG_TEXT["review_atmosphere"])
    fallback = _fallback_review(context)
    used_ollama = False
    ollama_model = ""

    review_text = fallback
    if ollama_ready:
        response = generate_ollama_text(_ollama_prompt(context), timeout=60)
        if response.get("ok"):
            review_text = _normalize_review(str(response.get("text") or ""), fallback)
            used_ollama = review_text != fallback
            ollama_model = str(response.get("model") or "")
        else:
            emit(LOG_TEXT["review_template_fallback"])
    else:
        emit(LOG_TEXT["review_template_fallback"])

    emit(LOG_TEXT["review_shorts"])
    review_path = package_dir / REVIEW_FILE_NAME
    header = (
        "<!--\n"
        f"created_at: {datetime.now().isoformat(timespec='seconds')}\n"
        f"package: {package_dir.name}\n"
        f"ollama: {'used' if used_ollama else 'template fallback'}\n"
        "-->\n\n"
    )
    review_path.write_text(header + review_text, encoding="utf-8")
    emit(LOG_TEXT["review_created"])

    return {
        "status": "COMPLETED",
        "package_dir": str(package_dir),
        "review_path": str(review_path),
        "used_ollama": used_ollama,
        "ollama_model": ollama_model,
        "files_read": context["files_read"],
        "missing_files": context["missing_files"],
    }
