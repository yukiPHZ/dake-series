from __future__ import annotations

import json
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any


NOTE_URL_PREFIX = "https://note.com/"

CHANNEL_RULES: dict[str, dict[str, object]] = {
    "brainz-inbox": {
        "qpsc_type": "inbox",
        "tags": ["inbox"],
        "series": ["inbox"],
    },
    "brainz-aru": {
        "qpsc_type": "aru",
        "tags": ["在る", "BORINEF"],
        "series": ["在る", "BORINEF"],
        "borinef_candidate": True,
    },
    "brainz-note": {
        "qpsc_type": "note_published",
        "tags": ["note", "BORINEF"],
        "series": ["note", "BORINEF"],
        "borinef_candidate": True,
    },
    "brainz-codex": {
        "qpsc_type": "codex_log",
        "tags": ["Codex", "正本"],
        "series": ["Codex"],
    },
    "brainz-reaction": {
        "qpsc_type": "reaction",
        "tags": ["reaction", "熱反応"],
        "series": ["reaction"],
    },
}

BORINEF_KEYWORDS = (
    "在る",
    "側に",
    "握らない",
    "線",
    "境界線",
    "ニュートラル",
    "たい",
    "れる",
    "熾火",
    "巡り",
    "円環",
    "呼吸",
    "静けさ",
)
BUSINESS_KEYWORDS = ("不動産", "土地", "売主", "買主", "価格", "越境", "解体", "道路", "役所", "査定", "契約", "重説")
CODEX_KEYWORDS = ("Codex", "commit", "push", "README", "Git", "build", "release", "Cloudflare")
MEDIUM_HEAT_KEYWORDS = ("やばい", "これだ", "違和感", "残る", "戻る", "苦しい", "悔しい", "大事", "重要", "本質", "震える")
SYSTEM_KEYWORDS: dict[str, tuple[str, ...]] = {
    "線系": ("線", "境界線", "輪郭", "側に", "ニュートラル", "触れられる", "離れられる", "俺とお前"),
    "熱系": ("熱", "ドラマ", "事実と熱", "たい", "れる", "冷静", "熾火", "熱が静まる"),
    "円環系": ("円環", "巡り", "戻る", "深くなる", "Hello World", "宇宙", "古いのに新しい", "未来に帰る"),
    "原本系": ("在る用", "AI対話", "実務ログ", "思想断片", "ChatGPT", "Codex", "生成途中"),
    "実務系": ("会社", "人間関係", "聞かれたら答える", "巻き込まれない", "正しさ", "役割", "距離感", "不動産"),
}


@dataclass(frozen=True)
class QpscEnrichment:
    qpsc_type: str
    tags: list[str]
    qpsc_series: list[str]
    qpsc_heat: str
    borinef_candidate: bool
    borinef_systems: list[str]
    borinef_depth: str
    oikawa_notify: bool
    oikawa_heat: bool
    oikawa_revisit: bool


@dataclass(frozen=True)
class QpscEnrichmentResult:
    text: str
    changed: bool
    metadata: QpscEnrichment


def normalize_channel_name(value: str) -> str:
    clean = str(value or "").strip().lstrip("#").lower()
    if clean in {"slack inbox", "slack_inbox"}:
        return "brainz-inbox"
    if clean in {"aru inbox", "aru_inbox"}:
        return "brainz-aru"
    return clean


def strip_frontmatter(markdown: str) -> str:
    _frontmatter, body, _had_frontmatter = split_frontmatter(markdown)
    return body


def split_frontmatter(markdown: str) -> tuple[list[str], str, bool]:
    text = markdown or ""
    lines = text.splitlines()
    if not lines or lines[0].strip() != "---":
        return [], text, False
    for index in range(1, len(lines)):
        if lines[index].strip() == "---":
            body = "\n".join(lines[index + 1 :])
            if text.endswith("\n"):
                body += "\n"
            return lines[1:index], body, True
    return [], text, False


def analyze_qpsc_markdown(
    markdown: str,
    channel_name: str = "",
    source_type: str = "",
    enable_oikawa_notify: bool = True,
    enable_heat: bool = True,
) -> QpscEnrichment:
    body = strip_frontmatter(markdown)
    channel_key = normalize_channel_name(channel_name)
    source_key = str(source_type or "").strip().lower()
    rule = dict(CHANNEL_RULES.get(channel_key, {}))
    if not rule:
        if "codex" in source_key:
            rule = dict(CHANNEL_RULES["brainz-codex"])
        elif source_key in {"aru"}:
            rule = dict(CHANNEL_RULES["brainz-aru"])
        elif "note" in source_key or NOTE_URL_PREFIX in body:
            rule = dict(CHANNEL_RULES["brainz-note"])
        else:
            rule = dict(CHANNEL_RULES["brainz-inbox"])

    qpsc_type = str(rule.get("qpsc_type") or "inbox")
    tags = list(rule.get("tags") or [])
    qpsc_series = list(rule.get("series") or [qpsc_type])
    borinef_candidate = bool(rule.get("borinef_candidate", False))

    if NOTE_URL_PREFIX in body:
        qpsc_type = "note_published"
        tags.extend(["note", "BORINEF"])
        qpsc_series.extend(["note", "BORINEF"])
        borinef_candidate = True

    if _contains_any(body, BORINEF_KEYWORDS):
        tags.extend(["BORINEF", "在る"])
        qpsc_series.append("BORINEF")
        borinef_candidate = True
    if _contains_any(body, BUSINESS_KEYWORDS):
        tags.extend(["実務", "不動産"])
    if _contains_any(body, CODEX_KEYWORDS):
        tags.extend(["Codex", "開発"])

    systems = [name for name, keywords in SYSTEM_KEYWORDS.items() if _contains_any(body, keywords)]
    if systems:
        borinef_candidate = True
        tags.append("BORINEF")
        qpsc_series.extend(systems)

    qpsc_heat = "medium" if _contains_any(body, MEDIUM_HEAT_KEYWORDS) else "low"
    return QpscEnrichment(
        qpsc_type=qpsc_type,
        tags=_unique(tags),
        qpsc_series=_unique(qpsc_series),
        qpsc_heat=qpsc_heat,
        borinef_candidate=borinef_candidate,
        borinef_systems=_unique(systems),
        borinef_depth="seed",
        oikawa_notify=bool(enable_oikawa_notify),
        oikawa_heat=bool(enable_heat),
        oikawa_revisit=True,
    )


def enrich_qpsc_markdown_text(
    markdown: str,
    channel_name: str = "",
    channel_id: str = "",
    source_type: str = "",
    enable_oikawa_notify: bool = True,
    enable_heat: bool = True,
) -> QpscEnrichmentResult:
    metadata = analyze_qpsc_markdown(
        markdown,
        channel_name=channel_name,
        source_type=source_type,
        enable_oikawa_notify=enable_oikawa_notify,
        enable_heat=enable_heat,
    )
    frontmatter, body, had_frontmatter = split_frontmatter(markdown)
    lines = list(frontmatter)
    if source_type and not top_key_exists(lines, "source"):
        lines.append(f"source: {source_type}")
    if channel_name and not top_key_exists(lines, "channel"):
        lines.append(f"channel: {normalize_channel_name(channel_name) or channel_name}")
    if channel_id and not top_key_exists(lines, "channel_id"):
        lines.append(f"channel_id: {channel_id}")
    ensure_scalar(lines, "qpsc_type", metadata.qpsc_type)
    ensure_scalar(lines, "qpsc_heat", metadata.qpsc_heat)
    ensure_list(lines, "qpsc_series", metadata.qpsc_series)
    ensure_list(lines, "tags", metadata.tags)
    ensure_nested_scalar(lines, "oikawa", "notify", metadata.oikawa_notify)
    ensure_nested_scalar(lines, "oikawa", "heat", metadata.oikawa_heat)
    ensure_nested_scalar(lines, "oikawa", "revisit", metadata.oikawa_revisit)
    ensure_nested_scalar(lines, "borinef", "candidate", metadata.borinef_candidate)
    ensure_nested_list(lines, "borinef", "systems", metadata.borinef_systems)
    ensure_nested_scalar(lines, "borinef", "depth", metadata.borinef_depth)
    ensure_nested_list(lines, "borinef", "circular_links", [])

    if had_frontmatter:
        new_text = "---\n" + "\n".join(lines).rstrip() + "\n---\n" + body.lstrip("\n")
    else:
        new_text = "---\n" + "\n".join(lines).rstrip() + "\n---\n\n" + markdown.lstrip("\n")
    if markdown.endswith("\n") and not new_text.endswith("\n"):
        new_text += "\n"
    return QpscEnrichmentResult(text=new_text, changed=new_text != markdown, metadata=metadata)


def enrich_qpsc_markdown_file(
    path: Path,
    channel_name: str = "",
    channel_id: str = "",
    source_type: str = "",
    enable_oikawa_notify: bool = True,
    enable_heat: bool = True,
) -> QpscEnrichmentResult:
    original = path.read_text(encoding="utf-8")
    result = enrich_qpsc_markdown_text(
        original,
        channel_name=channel_name,
        channel_id=channel_id,
        source_type=source_type,
        enable_oikawa_notify=enable_oikawa_notify,
        enable_heat=enable_heat,
    )
    if result.changed:
        path.write_text(result.text, encoding="utf-8")
    return result


def top_key_exists(lines: list[str], key: str) -> bool:
    prefix = f"{key}:"
    return any(not line.startswith((" ", "\t")) and line.strip().startswith(prefix) for line in lines)


def ensure_scalar(lines: list[str], key: str, value: object) -> None:
    if top_key_exists(lines, key):
        return
    lines.append(f"{key}: {format_yaml_value(value)}")


def ensure_list(lines: list[str], key: str, values: list[str]) -> None:
    values = _unique(values)
    if not values and top_key_exists(lines, key):
        return
    start, end = top_block_range(lines, key)
    if start is None:
        lines.append(f"{key}:")
        lines.extend(f"  - {value}" for value in values)
        return
    if ":" in lines[start] and lines[start].split(":", 1)[1].strip():
        return
    existing = set(list_values(lines[start + 1 : end], indent=2))
    insert_at = end
    for value in values:
        if value not in existing:
            lines.insert(insert_at, f"  - {value}")
            insert_at += 1
            existing.add(value)


def ensure_nested_scalar(lines: list[str], parent: str, key: str, value: object) -> None:
    start, end = ensure_parent_block(lines, parent)
    nested_prefix = f"{key}:"
    for line in lines[start + 1 : end]:
        if line.startswith("  ") and line.strip().startswith(nested_prefix):
            return
    lines.insert(end, f"  {key}: {format_yaml_value(value)}")


def ensure_nested_list(lines: list[str], parent: str, key: str, values: list[str]) -> None:
    values = _unique(values)
    parent_start, parent_end = ensure_parent_block(lines, parent)
    nested_start = None
    nested_end = parent_end
    for index in range(parent_start + 1, parent_end):
        line = lines[index]
        if line.startswith("  ") and not line.startswith("    ") and line.strip().startswith(f"{key}:"):
            nested_start = index
            nested_end = parent_end
            for end_index in range(index + 1, parent_end):
                nested_line = lines[end_index]
                if nested_line.startswith("  ") and not nested_line.startswith("    ") and ":" in nested_line.strip():
                    nested_end = end_index
                    break
            break
    if nested_start is None:
        lines.insert(parent_end, f"  {key}:")
        insert_at = parent_end + 1
        for value in values:
            lines.insert(insert_at, f"    - {value}")
            insert_at += 1
        return
    if ":" in lines[nested_start] and lines[nested_start].split(":", 1)[1].strip():
        return
    existing = set(list_values(lines[nested_start + 1 : nested_end], indent=4))
    insert_at = nested_end
    for value in values:
        if value not in existing:
            lines.insert(insert_at, f"    - {value}")
            insert_at += 1
            existing.add(value)


def ensure_parent_block(lines: list[str], parent: str) -> tuple[int, int]:
    start, end = top_block_range(lines, parent)
    if start is None:
        lines.append(f"{parent}:")
        start = len(lines) - 1
        end = len(lines)
    return start, end


def top_block_range(lines: list[str], key: str) -> tuple[int | None, int]:
    prefix = f"{key}:"
    for index, line in enumerate(lines):
        if not line.startswith((" ", "\t")) and line.strip().startswith(prefix):
            end = len(lines)
            for end_index in range(index + 1, len(lines)):
                candidate = lines[end_index]
                if candidate.strip() and not candidate.startswith((" ", "\t")):
                    end = end_index
                    break
            return index, end
    return None, len(lines)


def list_values(lines: list[str], indent: int) -> list[str]:
    prefix = " " * indent + "-"
    result: list[str] = []
    for line in lines:
        if line.startswith(prefix):
            value = line.split("-", 1)[1].strip()
            if value:
                result.append(value.strip('"').strip("'"))
    return result


def format_yaml_value(value: object) -> str:
    if isinstance(value, bool):
        return "true" if value else "false"
    if value is None:
        return ""
    text = str(value)
    if not text:
        return '""'
    if any(char in text for char in [":", "#", "[", "]", "{", "}", "\n"]):
        return json.dumps(text, ensure_ascii=False)
    return text


def _contains_any(text: str, keywords: tuple[str, ...]) -> bool:
    lowered = text.lower()
    return any(keyword.lower() in lowered for keyword in keywords)


def _unique(values: list[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for value in values:
        clean = str(value or "").strip()
        if not clean or clean in seen:
            continue
        seen.add(clean)
        result.append(clean)
    return result


def sidecar_metadata(metadata: QpscEnrichment) -> dict[str, Any]:
    return {
        "tags": metadata.tags,
        "qpsc_heat": metadata.qpsc_heat,
        "qpsc_type": metadata.qpsc_type,
        "borinef_candidate": metadata.borinef_candidate,
        "systems": metadata.borinef_systems,
        "oikawa_notify": metadata.oikawa_notify,
        "updated_at": datetime.now().isoformat(timespec="seconds"),
    }
