from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any


PEAKHEADZ_ROOT_NAME = "PEAKHEADZ_ROOT"
LEGACY_BRAINZ_MEMORY_NAME = "brainz_memory"


QPSC_MEMORY_FOLDER_PATHS = (
    "00_inbox",
    "10_slack/brainz-inbox",
    "10_slack/brainz-note",
    "10_slack/brainz-codex",
    "10_slack/brainz-aru",
    "10_slack/brainz-reaction",
    "20_chatgpt",
    "30_codex",
    "40_borinef/note/published",
    "40_borinef/note/reactions",
    "40_borinef/note/source_fragments",
    "90_system/logs",
    "90_system/config",
)


DEFAULT_SLACK_CHANNEL_ROUTES: tuple[dict[str, Any], ...] = (
    {
        "channel_name": "#brainz-inbox",
        "channel_id": "",
        "enabled": True,
        "purpose": "unsettled_heat",
        "save_folder": "10_slack/brainz-inbox",
        "default_tags": ["inbox"],
        "qpsc_source": "slack",
        "enable_oikawa_notify": True,
        "enable_heat": True,
    },
    {
        "channel_name": "#brainz-note",
        "channel_id": "",
        "enabled": True,
        "purpose": "published_note",
        "save_folder": "10_slack/brainz-note",
        "default_tags": ["BORINEF", "note"],
        "qpsc_source": "slack",
        "enable_oikawa_notify": True,
        "enable_heat": True,
    },
    {
        "channel_name": "#brainz-codex",
        "channel_id": "",
        "enabled": True,
        "purpose": "codex_reports",
        "save_folder": "10_slack/brainz-codex",
        "default_tags": ["codex"],
        "qpsc_source": "slack",
        "enable_oikawa_notify": True,
        "enable_heat": True,
    },
    {
        "channel_name": "#brainz-aru",
        "channel_id": "",
        "enabled": True,
        "purpose": "aru_fragments",
        "save_folder": "10_slack/brainz-aru",
        "default_tags": ["aru"],
        "qpsc_source": "aru",
        "enable_oikawa_notify": True,
        "enable_heat": True,
    },
    {
        "channel_name": "#brainz-reaction",
        "channel_id": "",
        "enabled": False,
        "purpose": "reactions",
        "save_folder": "10_slack/brainz-reaction",
        "default_tags": ["reaction"],
        "qpsc_source": "slack",
        "enable_oikawa_notify": False,
        "enable_heat": False,
    },
)


def documents_root() -> Path:
    return Path.home() / "Documents"


def recommended_peakheadz_root(documents_dir: Path | None = None) -> Path:
    return (documents_dir or documents_root()) / PEAKHEADZ_ROOT_NAME


def legacy_brainz_memory_root(documents_dir: Path | None = None) -> Path:
    return (documents_dir or documents_root()) / LEGACY_BRAINZ_MEMORY_NAME


def existing_directory(path: str | Path) -> Path | None:
    if not path:
        return None
    target = Path(os.path.expandvars(str(path))).expanduser()
    try:
        resolved = target.resolve()
    except OSError:
        resolved = target
    if resolved.exists() and resolved.is_dir():
        return resolved
    return None


def resolve_qpsc_memory_root(config_memory_folder: str = "", documents_dir: Path | None = None) -> tuple[Path | None, str]:
    configured = existing_directory(config_memory_folder)
    if configured is not None:
        return configured, "configured"
    if str(config_memory_folder or "").strip():
        return None, "configured_missing"

    peakheadz = existing_directory(recommended_peakheadz_root(documents_dir))
    if peakheadz is not None:
        return peakheadz, "peakheadz_root"

    legacy = existing_directory(legacy_brainz_memory_root(documents_dir))
    if legacy is not None:
        return legacy, "legacy_brainz_memory"

    return None, "unset"


def legacy_brainz_memory_exists(documents_dir: Path | None = None) -> bool:
    return existing_directory(legacy_brainz_memory_root(documents_dir)) is not None


def ensure_peakheadz_root_structure(documents_dir: Path | None = None) -> tuple[Path, list[Path]]:
    root = recommended_peakheadz_root(documents_dir)
    root.mkdir(parents=True, exist_ok=True)
    return root.resolve(), ensure_qpsc_memory_structure(root)


def ensure_qpsc_memory_structure(memory_folder: Path) -> list[Path]:
    root = Path(memory_folder).expanduser()
    created: list[Path] = []
    for relative_path in QPSC_MEMORY_FOLDER_PATHS:
        path = root / Path(relative_path)
        path.mkdir(parents=True, exist_ok=True)
        created.append(path.resolve())
    return created


def default_slack_channel_routes() -> list[dict[str, Any]]:
    return [dict(route) for route in DEFAULT_SLACK_CHANNEL_ROUTES]


def normalize_slack_channel_routes(value: Any) -> list[dict[str, Any]]:
    source_routes = value if isinstance(value, list) else []
    route_map: dict[str, dict[str, Any]] = {}
    route_order: list[str] = []
    for default_route in default_slack_channel_routes():
        key = normalize_channel_key(str(default_route.get("channel_name") or ""))
        route_map[key] = dict(default_route)
        route_order.append(key)
    for item in source_routes:
        if not isinstance(item, dict):
            continue
        channel_name = str(item.get("channel_name") or "").strip()
        if not channel_name:
            continue
        key = normalize_channel_key(channel_name)
        merged = dict(route_map.get(key, {}))
        merged.update(item)
        route_map[key] = merged
        if key not in route_order:
            route_order.append(key)

    normalized: list[dict[str, Any]] = []
    for key in route_order:
        item = route_map.get(key, {})
        channel_name = str(item.get("channel_name") or "").strip()
        save_folder = str(item.get("save_folder") or "").strip()
        if not channel_name or not save_folder:
            continue
        tags = item.get("default_tags")
        normalized.append(
            {
                "channel_name": channel_name,
                "channel_id": str(item.get("channel_id") or "").strip(),
                "enabled": bool(item.get("enabled", True)),
                "purpose": str(item.get("purpose") or "").strip(),
                "save_folder": normalize_relative_folder(save_folder),
                "default_tags": [str(tag) for tag in tags] if isinstance(tags, list) else [],
                "qpsc_source": str(item.get("qpsc_source") or "slack").strip() or "slack",
                "enable_oikawa_notify": bool(item.get("enable_oikawa_notify", True)),
                "enable_heat": bool(item.get("enable_heat", True)),
                "last_imported_at": str(item.get("last_imported_at") or "").strip(),
                "last_ts": str(item.get("last_ts") or "").strip(),
            }
        )
    return normalized or default_slack_channel_routes()


def load_slack_channel_routes(config_value: str) -> list[dict[str, Any]]:
    if not str(config_value or "").strip():
        return default_slack_channel_routes()
    try:
        payload = json.loads(config_value)
    except json.JSONDecodeError:
        return default_slack_channel_routes()
    return normalize_slack_channel_routes(payload)


def dump_slack_channel_routes(routes: list[dict[str, Any]] | None = None) -> str:
    normalized = normalize_slack_channel_routes(routes or default_slack_channel_routes())
    return json.dumps(normalized, ensure_ascii=False, indent=2)


def find_slack_channel_route(channel: str, routes: list[dict[str, Any]]) -> dict[str, Any] | None:
    clean_channel = normalize_channel_key(channel)
    clean_raw = str(channel or "").strip().lower()
    for route in routes:
        route_channel = normalize_channel_key(str(route.get("channel_name") or ""))
        route_channel_id = str(route.get("channel_id") or "").strip().lower()
        if clean_channel and clean_channel == route_channel:
            return route
        if clean_raw and route_channel_id and clean_raw == route_channel_id:
            return route
    return None


def route_summary_lines(routes: list[dict[str, Any]]) -> list[str]:
    lines: list[str] = []
    for route in normalize_slack_channel_routes(routes):
        status = "on" if route.get("enabled", True) and route.get("channel_id") else "unset"
        lines.append(f"{route['channel_name']} [{status}] -> {route['save_folder']}")
    return lines


def normalize_channel_key(value: str) -> str:
    return str(value or "").strip().lstrip("#").lower()


def normalize_relative_folder(value: str) -> str:
    parts: list[str] = []
    for part in str(value or "").replace("\\", "/").split("/"):
        clean = "".join(ch if ch.isalnum() or ch in {"_", "-", "."} else "_" for ch in part.strip()).strip("._")
        if clean and clean not in {".", ".."}:
            parts.append(clean)
    return "/".join(parts) or "00_inbox"


def codex_sessions_default_path() -> Path:
    return Path.home() / ".codex" / "sessions"


def codex_watch_status(path: str | Path) -> tuple[bool, str]:
    target = Path(os.path.expandvars(str(path))).expanduser()
    return target.exists() and target.is_dir(), str(target)
