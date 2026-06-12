# -*- coding: utf-8 -*-
"""Generate DAKE release social posts and optionally create Buffer posts.

The only public link included in generated copy is the dakeapp.com detail page.
BOOTH URLs and GitHub Release URLs are intentionally not placed in SNS body
text. Buffer API credentials are read from ``BUFFER_API_KEY`` and never written
to logs or JSON.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass, field
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Any

from release_source_policy import app_url_for, find_app as source_find_app, read_app_source


ROOT = Path(__file__).resolve().parents[1]
APPS_DIR = ROOT / "01_apps"
DEFAULT_SITE_ROOT = Path(os.environ.get("DAKEAPP_SITE_ROOT", r"C:\Users\yukiz\devlop\dakeapp-site"))
JST = timezone(timedelta(hours=9))
PLATFORMS = ("x", "threads", "instagram")
BUFFER_API_BASE = os.environ.get("BUFFER_API_BASE", "https://api.bufferapp.com/1").rstrip("/")


@dataclass
class SocialOutcome:
    app_dir: Path
    meta: dict[str, Any]
    app_url: str
    posts: dict[str, str]
    buffer: dict[str, dict[str, Any]] = field(default_factory=dict)
    stage: str = "failed"
    reason: str = ""


def now_iso() -> str:
    return datetime.now(JST).isoformat(timespec="seconds")


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return path.read_text(encoding="utf-8", errors="replace")


def write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text.rstrip() + "\n", encoding="utf-8", newline="\n")


def write_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8", newline="\n")


def extract_json_section(text: str, heading: str) -> dict[str, Any] | None:
    match = re.search(rf"(?s)^\s*##\s+{re.escape(heading)}\s*```json\s*(.*?)\s*```", text, re.MULTILINE)
    if not match:
        return None
    loaded = json.loads(match.group(1))
    return loaded if isinstance(loaded, dict) else None



def read_meta(app_dir: Path) -> tuple[dict[str, Any], str]:
    source = read_app_source(app_dir, ROOT)
    return source.meta, source.error


def find_app(identifier: str) -> Path:
    return source_find_app(APPS_DIR, identifier, ROOT)

def safe_sentence(value: str, fallback: str) -> str:
    text = re.sub(r"\s+", " ", value or "").strip()
    if not text:
        return fallback
    return text



def trim_for_x(text: str) -> str:
    if len(text) <= 280:
        return text
    lines = text.splitlines()
    url = lines[-1]
    head = "\n".join(lines[:-1])
    allowance = 280 - len(url) - 1
    if allowance < 20:
        return text[:277] + "..."
    return head[: allowance - 1].rstrip("。、") + "...\n" + url


def build_posts(meta: dict[str, Any], app_url: str) -> dict[str, str]:
    name = str(meta.get("site_title") or meta.get("display_name") or meta.get("launcher_title") or "DAKEアプリ")
    feature = safe_sentence(str(meta.get("launcher_description") or meta.get("site_description") or ""), name)
    background = safe_sentence(str(meta.get("update_summary") or meta.get("site_description") or ""), feature)
    posts = {
        "x": trim_for_x(f"{name} を公開しました。\n{feature}\n{app_url}"),
        "threads": (
            f"{name} を公開しました。\n\n"
            f"{background}\n"
            "配布物だけでなく、動いている姿を確認できるところまで整えて出荷しています。\n\n"
            f"{app_url}"
        ),
        "instagram": (
            f"{name}\n\n"
            f"{feature}\n\n"
            f"詳細: {app_url}\n\n"
            "#DAKE #Windowsアプリ #作業効率化"
        ),
    }
    for platform, text in posts.items():
        lowered = text.lower()
        if "github.com" in lowered or "booth.pm" in lowered:
            raise ValueError(f"{platform} post contains forbidden direct release/shop URL")
    return posts


def write_posts_markdown(path: Path, outcome: SocialOutcome) -> None:
    name = str(outcome.meta.get("display_name") or outcome.app_dir.name)
    lines = [
        f"# DAKE SNS投稿文: {name}",
        "",
        f"- detail_url: {outcome.app_url}",
        "- links: dakeapp.com detail page only",
        "",
    ]
    for platform in PLATFORMS:
        lines.extend([f"## {platform.upper()}", "", outcome.posts[platform], ""])
    write_text(path, "\n".join(lines))

def request_json(url: str, api_key: str, data: dict[str, Any] | None = None) -> dict[str, Any]:
    headers = {
        "User-Agent": "DAKE-release-social/1.0",
        "Authorization": f"Bearer {api_key}",
    }
    body: bytes | None = None
    if data is not None:
        encoded = urllib.parse.urlencode(data, doseq=True)
        body = encoded.encode("utf-8")
        headers["Content-Type"] = "application/x-www-form-urlencoded"
    request = urllib.request.Request(url, data=body, headers=headers)
    with urllib.request.urlopen(request, timeout=20) as response:
        raw = response.read().decode("utf-8", errors="replace")
    return json.loads(raw) if raw else {}


def list_buffer_profiles(api_key: str) -> list[dict[str, Any]]:
    try:
        loaded = request_json(f"{BUFFER_API_BASE}/profiles.json", api_key)
    except urllib.error.HTTPError:
        url = f"{BUFFER_API_BASE}/profiles.json?{urllib.parse.urlencode({'access_token': api_key})}"
        loaded = request_json(url, api_key)
    if isinstance(loaded, list):
        return [item for item in loaded if isinstance(item, dict)]
    if isinstance(loaded, dict) and isinstance(loaded.get("profiles"), list):
        return [item for item in loaded["profiles"] if isinstance(item, dict)]
    return []


def profile_id_from_env(platform: str) -> str:
    env_names = {
        "x": ["BUFFER_PROFILE_X_ID", "BUFFER_PROFILE_TWITTER_ID"],
        "threads": ["BUFFER_PROFILE_THREADS_ID"],
        "instagram": ["BUFFER_PROFILE_INSTAGRAM_ID"],
    }
    for name in env_names[platform]:
        value = os.environ.get(name, "").strip()
        if value:
            return value
    return ""


def discover_profile_id(platform: str, profiles: list[dict[str, Any]]) -> str:
    service_names = {
        "x": {"x", "twitter"},
        "threads": {"threads"},
        "instagram": {"instagram", "instagram_business"},
    }[platform]
    for profile in profiles:
        service = str(profile.get("service") or profile.get("service_type") or "").lower()
        if service in service_names:
            return str(profile.get("id") or profile.get("_id") or "")
    return ""


def extract_buffer_update_id(data: dict[str, Any]) -> str:
    for key in ("id", "update_id", "buffer_update_id"):
        value = data.get(key)
        if value:
            return str(value)
    update = data.get("update")
    if isinstance(update, dict):
        for key in ("id", "_id", "update_id"):
            if update.get(key):
                return str(update[key])
    updates = data.get("updates")
    if isinstance(updates, list) and updates:
        first = updates[0]
        if isinstance(first, dict):
            for key in ("id", "_id", "update_id"):
                if first.get(key):
                    return str(first[key])
    return ""


def safe_error(exc: Exception, api_key: str) -> str:
    message = str(exc)
    return message.replace(api_key, "[BUFFER_API_KEY]") if api_key else message


def create_buffer_post(platform: str, text: str, api_key: str, profile_id: str, thumbnail_url: str) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "access_token": api_key,
        "profile_ids[]": profile_id,
        "text": text,
        "now": "false",
        "shorten": "false",
    }
    if platform == "instagram" and thumbnail_url:
        payload["media[photo]"] = thumbnail_url
    response = request_json(f"{BUFFER_API_BASE}/updates/create.json", api_key, payload)
    update_id = extract_buffer_update_id(response)
    return {
        "platform": platform,
        "success": bool(update_id),
        "buffer_update_id": update_id,
    }



def create_social(app_dir: Path, args: argparse.Namespace) -> SocialOutcome:
    source = read_app_source(app_dir, ROOT)
    meta = source.meta
    meta_error = source.error
    app_url = app_url_for(app_dir, meta, args.site_root)
    posts = build_posts(meta, app_url)
    outcome = SocialOutcome(app_dir=app_dir, meta=meta, app_url=app_url, posts=posts)
    release_dir = app_dir / "release_artifacts"
    release_dir.mkdir(parents=True, exist_ok=True)
    posts_path = release_dir / "social_posts.md"
    release_path = release_dir / "social_release.json"
    write_posts_markdown(posts_path, outcome)

    reasons: list[str] = []
    if meta_error:
        reasons.append(meta_error)
    if args.post_to_buffer:
        api_key = os.environ.get("BUFFER_API_KEY", "").strip()
        if not api_key:
            reasons.append("BUFFER_API_KEY is not set")
        else:
            profiles: list[dict[str, Any]] = []
            try:
                profiles = list_buffer_profiles(api_key)
            except Exception as exc:
                profiles = []
                reasons.append(f"Buffer profile discovery failed: {safe_error(exc, api_key)}")
            for platform in PLATFORMS:
                profile_id = profile_id_from_env(platform) or discover_profile_id(platform, profiles)
                if not profile_id:
                    outcome.buffer[platform] = {"platform": platform, "success": False, "reason": "Buffer profile id not found"}
                    reasons.append(f"{platform}: Buffer profile id not found")
                    continue
                try:
                    outcome.buffer[platform] = create_buffer_post(platform, posts[platform], api_key, profile_id, args.thumbnail_url)
                    if not outcome.buffer[platform].get("buffer_update_id"):
                        reasons.append(f"{platform}: Buffer update id missing")
                except Exception as exc:
                    outcome.buffer[platform] = {"platform": platform, "success": False, "reason": safe_error(exc, api_key)}
                    reasons.append(f"{platform}: Buffer create failed")
    else:
        for platform in PLATFORMS:
            outcome.buffer[platform] = {"platform": platform, "success": False, "buffer_update_id": "", "reason": "dry run"}

    has_all_ids = all(outcome.buffer.get(platform, {}).get("buffer_update_id") for platform in PLATFORMS)
    if args.post_to_buffer:
        outcome.stage = "complete" if has_all_ids and not reasons else "failed"
        outcome.reason = "" if outcome.stage == "complete" else "; ".join(reasons)
    else:
        outcome.stage = "dry_run"
        outcome.reason = "Buffer posting not requested; generated social_posts.md only"

    record = {
        "app_key": str(meta.get("app_key") or app_dir.name),
        "display_name": str(meta.get("display_name") or meta.get("site_title") or app_dir.name),
        "app_url": app_url,
        "created_at": now_iso(),
        "stage": outcome.stage,
        "reason": outcome.reason,
        "social_posts_path": "release_artifacts/social_posts.md",
        "link_policy": "dakeapp.com detail page only",
        "dry_run": not args.post_to_buffer,
        "posts": outcome.posts,
        "buffer": outcome.buffer,
        "source_policy": {
            "source_kind": source.source_kind,
            "source_path": source.source_label,
            "original_missing": source.original_missing,
            "meta_derivative_mismatch": source.derivative_mismatches,
        },
        "tool": {"name": "tools/release_social.py", "version": 1},
    }
    write_json(release_path, record)
    return outcome

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate DAKE SNS posts and optionally create Buffer posts.")
    parser.add_argument("--app", action="append", required=True)
    parser.add_argument("--site-root", type=Path, default=DEFAULT_SITE_ROOT)
    parser.add_argument("--post-to-buffer", action="store_true")
    parser.add_argument("--thumbnail-url", default="", help="public thumbnail URL for Instagram Buffer media")
    parser.add_argument("--fail-on-incomplete", action="store_true")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    outcomes = [create_social(find_app(item), args) for item in args.app]
    for outcome in outcomes:
        print(f"{outcome.app_dir.name}: {outcome.stage} {outcome.reason}")
    if args.fail_on_incomplete and any(outcome.stage != "complete" for outcome in outcomes):
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
