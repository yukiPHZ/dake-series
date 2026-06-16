# -*- coding: utf-8 -*-
"""Generate DAKE release social posts and manage Buffer draft evidence.

Normal shipping creates Buffer drafts only. It never publishes immediately,
schedules posts, or intentionally queues posts unless a schedule plan is
explicitly applied with confirmation.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import urllib.error
import urllib.parse
import urllib.request
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

from release_source_policy import (
    app_url_for,
    find_app as source_find_app,
    find_product as source_find_product,
    product_dirs as source_product_dirs,
    product_url_for,
    read_product_source,
    store_image_url_for_product,
)


ROOT = Path(__file__).resolve().parents[1]
APPS_DIR = ROOT / "01_apps"
PACKS_DIR = ROOT / "04_packs"
DEFAULT_SITE_ROOT = Path(os.environ.get("DAKEAPP_SITE_ROOT", r"C:\Users\yukiz\devlop\dakeapp-site"))
DEFAULT_STORE_SITE_ROOT = Path(os.environ.get("DAKE_STORE_SITE_ROOT", r"C:\Users\yukiz\devlop\dake-store-site"))
DEFAULT_SCHEDULE_PLAN = ROOT / "tools" / "reports" / "social_schedule_plan.json"
CENTRAL_ARTIFACT_ROOT = ROOT / "tools" / "reports" / "release_artifacts"
JST = timezone(timedelta(hours=9))
PLATFORMS = ("x", "threads", "instagram")
CHANNEL_LABELS = {"x": "X", "threads": "Threads", "instagram": "Instagram"}
SCHEDULE_STATUS_VALUES = {"not_requested", "draft", "scheduled", "published", "skipped", "failed", "unknown"}
MAX_BATCH_LIMIT = 10
SHIPPING_STATUS = "available"
BUFFER_GRAPHQL_ENDPOINT = os.environ.get("BUFFER_GRAPHQL_ENDPOINT", "https://api.buffer.com").rstrip("/")
GRAPHQL_TIMEOUT = float(os.environ.get("BUFFER_GRAPHQL_TIMEOUT", "20"))
SERVICE_NAMES = {
    "x": {"twitter", "x"},
    "threads": {"threads"},
    "instagram": {"instagram"},
}
CHANNEL_ENV = {
    "x": "BUFFER_CHANNEL_ID_X",
    "threads": "BUFFER_CHANNEL_ID_THREADS",
    "instagram": "BUFFER_CHANNEL_ID_INSTAGRAM",
}

GET_ORGANIZATIONS_QUERY = """
query GetOrganizations {
  account {
    organizations {
      id
    }
  }
}
"""

GET_CHANNELS_QUERY = """
query GetChannels($organizationId: OrganizationId!) {
  channels(input: { organizationId: $organizationId }) {
    id
    service
    isDisconnected
    isLocked
    isQueuePaused
  }
}
"""

CREATE_DRAFT_POST_MUTATION = """
mutation CreateDraftPost($input: CreatePostInput!) {
  createPost(input: $input) {
    __typename
    ... on PostActionSuccess {
      post {
        id
      }
    }
    ... on MutationError {
      message
    }
  }
}
"""

EDIT_POST_MUTATION = """
mutation EditScheduledPost($input: EditPostInput!) {
  editPost(input: $input) {
    __typename
    ... on PostActionSuccess {
      post {
        id
      }
    }
    ... on MutationError {
      message
    }
  }
}
"""


@dataclass
class SocialOutcome:
    app_dir: Path
    meta: dict[str, Any]
    app_url: str
    posts: dict[str, str]
    requested_channels: list[str]
    buffer: dict[str, dict[str, Any]] = field(default_factory=dict)
    stage: str = "failed"
    reason: str = ""


@dataclass
class ScheduleItem:
    app_dir: Path
    app_key: str
    channels: list[str]
    scheduled_at: str
    due_at: str


class BufferGraphQLError(Exception):
    def __init__(self, code: str, detail: str = "") -> None:
        super().__init__(code)
        self.code = code
        self.detail = detail


class SchedulePlanError(Exception):
    def __init__(self, code: str) -> None:
        super().__init__(code)
        self.code = code


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


def read_secret_env(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if value or os.name != "nt":
        return value
    try:
        import winreg
    except Exception:
        return ""
    locations = (
        (winreg.HKEY_CURRENT_USER, r"Environment"),
        (winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment"),
    )
    for root, subkey in locations:
        try:
            with winreg.OpenKey(root, subkey) as handle:
                found, _ = winreg.QueryValueEx(handle, name)
        except OSError:
            continue
        value = str(found).strip()
        if value:
            return value
    return ""


def read_plain_env(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if value or os.name != "nt":
        return value
    try:
        import winreg
    except Exception:
        return ""
    for root, subkey in (
        (winreg.HKEY_CURRENT_USER, r"Environment"),
        (winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment"),
    ):
        try:
            with winreg.OpenKey(root, subkey) as handle:
                found, _ = winreg.QueryValueEx(handle, name)
        except OSError:
            continue
        value = str(found).strip()
        if value:
            return value
    return ""


def read_meta(app_dir: Path) -> tuple[dict[str, Any], str]:
    source = read_product_source(app_dir, ROOT)
    return source.meta, source.error


def find_app(identifier: str) -> Path:
    return source_find_product(APPS_DIR, PACKS_DIR, identifier, ROOT)


def app_dirs() -> list[Path]:
    return source_product_dirs(APPS_DIR, PACKS_DIR)


def product_id_for(app_dir: Path, meta: dict[str, Any]) -> str:
    return str(meta.get("app_key") or meta.get("folder_name") or app_dir.name)


def product_type_for(meta: dict[str, Any]) -> str:
    return str(meta.get("product_type") or "app")


def social_artifact_dir(app_dir: Path, meta: dict[str, Any]) -> Path:
    if product_type_for(meta) == "pack":
        return CENTRAL_ARTIFACT_ROOT / product_id_for(app_dir, meta)
    return app_dir / "release_artifacts"


def normalize_channel(value: str) -> str:
    normalized = value.strip().lower()
    if normalized in {"twitter"}:
        return "x"
    if normalized not in PLATFORMS:
        raise ValueError(f"unknown channel: {value}")
    return normalized


def parse_channels(value: str | None) -> list[str]:
    if not value:
        return list(PLATFORMS)
    channels: list[str] = []
    for item in value.split(","):
        if not item.strip():
            continue
        channel = normalize_channel(item)
        if channel not in channels:
            channels.append(channel)
    if not channels:
        raise ValueError("channels must not be empty")
    return channels


def channels_from_record(record: dict[str, Any]) -> list[str]:
    raw = record.get("requested_channels")
    if isinstance(raw, list):
        channels: list[str] = []
        for item in raw:
            try:
                channel = normalize_channel(str(item))
            except ValueError:
                continue
            if channel not in channels:
                channels.append(channel)
        if channels:
            return channels
    return list(PLATFORMS)


def safe_sentence(value: str, fallback: str) -> str:
    text = re.sub(r"\s+", " ", value or "").strip()
    return text or fallback


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


def build_posts(meta: dict[str, Any], app_url: str, requested_channels: list[str]) -> dict[str, str]:
    name = str(meta.get("site_title") or meta.get("display_name") or meta.get("launcher_title") or "DAKEアプリ")
    feature = safe_sentence(str(meta.get("launcher_description") or meta.get("site_description") or ""), name)
    background = safe_sentence(str(meta.get("update_summary") or meta.get("site_description") or ""), feature)
    product_id = str(meta.get("app_key") or meta.get("folder_name") or "")
    if str(meta.get("product_type") or "") == "pack":
        if product_id == "DAKE_Pack_Mail":
            all_posts = {
                "x": trim_for_x(
                    "DAKE メール準備パックを公開しました。\n"
                    "メールを集める、整える、下書きにする。\n"
                    "送信前までの小さな実務をまとめたWindows向けPackです。\n"
                    f"{app_url}"
                ),
                "threads": (
                    "DAKE メール準備パックを公開しました。\n\n"
                    "Outlookメールから連絡先をCSVにし、メールアドレスを整え、CSVから個別メールの下書きを作る3本をまとめています。\n\n"
                    "メールは自動送信しません。下書きを確認してから、自分で送信できます。\n\n"
                    f"{app_url}"
                ),
                "instagram": (
                    "DAKE メール準備パック\n\n"
                    "集める。\n"
                    "整える。\n"
                    "下書きにする。\n\n"
                    "送信前までのメール実務をまとめた\n"
                    "Windows向けPackです。\n\n"
                    f"{app_url}\n\n"
                    "#DAKE #Windowsアプリ #Outlook #仕事効率化"
                ),
            }
        else:
            all_posts = {
                "x": trim_for_x(f"{name} を公開しました。\n{feature}\n{app_url}"),
                "threads": f"{name} を公開しました。\n\n{background}\n\n{app_url}",
                "instagram": f"{name}\n\n{feature}\n\n詳細: {app_url}\n\n#DAKE #Windowsアプリ #仕事効率化",
            }
        posts = {channel: all_posts[channel] for channel in requested_channels}
        for platform, text in posts.items():
            lowered = text.lower()
            if "github.com" in lowered or "booth.pm" in lowered:
                raise ValueError(f"{platform} post contains forbidden direct release/shop URL")
        return posts
    all_posts = {
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
    posts = {channel: all_posts[channel] for channel in requested_channels}
    for platform, text in posts.items():
        lowered = text.lower()
        if "github.com" in lowered or "booth.pm" in lowered:
            raise ValueError(f"{platform} post contains forbidden direct release/shop URL")
    return posts


def detail_slug_from_url(app_url: str) -> str:
    parts = app_url.rstrip("/").split("/")
    return parts[-1] if parts else ""


def public_site_image_url(app_dir: Path, app_url: str, site_root: Path) -> str:
    override = read_plain_env("BUFFER_INSTAGRAM_IMAGE_URL")
    if override.startswith("https://"):
        return override
    parsed = urllib.parse.urlparse(app_url)
    if parsed.netloc == "store.dakeapp.com":
        product_id = urllib.parse.parse_qs(parsed.query).get("id", [""])[0]
        if product_id:
            return store_image_url_for_product(product_id)
    slug = detail_slug_from_url(app_url)
    if not slug:
        return ""
    image_path = site_root / "public" / "assets" / "images" / "apps" / slug / "screenshot-01.webp"
    if image_path.exists():
        return f"https://dakeapp.com/assets/images/apps/{slug}/screenshot-01.webp"
    local_screenshot = app_dir / "assets" / "screenshot.webp"
    if local_screenshot.exists():
        return f"https://dakeapp.com/assets/images/apps/{slug}/screenshot-01.webp"
    return ""


def instagram_asset(image_url: str, display_name: str) -> dict[str, Any]:
    return {
        "image": {
            "url": image_url,
            "thumbnailUrl": image_url,
            "metadata": {"altText": f"{display_name} screenshot"},
        }
    }


def read_existing_social_posts(path: Path) -> dict[str, str]:
    if not path.exists():
        return {}
    text = read_text(path)
    sections: dict[str, str] = {}
    pattern = re.compile(r"(?ms)^##\s+(X|THREADS|INSTAGRAM)\s*\n(.*?)(?=^##\s+(?:X|THREADS|INSTAGRAM)\s*\n|\Z)")
    for match in pattern.finditer(text):
        channel = normalize_channel(match.group(1).lower())
        sections[channel] = match.group(2).strip()
    return sections


def write_posts_markdown(path: Path, outcome: SocialOutcome, existing_sections: dict[str, str]) -> None:
    name = str(outcome.meta.get("display_name") or outcome.app_dir.name)
    sections = dict(existing_sections)
    for channel, text in outcome.posts.items():
        sections[channel] = text
    ordered = [channel for channel in PLATFORMS if channel in sections]
    lines = [
        f"# DAKE SNS投稿文: {name}",
        "",
        f"- detail_url: {outcome.app_url}",
        "- links: dakeapp.com detail page only",
        f"- requested_channels: {', '.join(outcome.requested_channels)}",
        "",
    ]
    for platform in ordered:
        lines.extend([f"## {platform.upper()}", "", sections[platform], ""])
    write_text(path, "\n".join(lines))


def write_social_release_markdown(path: Path, outcome: SocialOutcome, record: dict[str, Any]) -> None:
    rows = [
        f"- product_id: {record.get('product_id') or record.get('app_key')}",
        f"- product_type: {record.get('product_type')}",
        f"- stage: {record.get('stage')}",
        f"- detail_url: {record.get('app_url')}",
        f"- requested_channels: {', '.join(record.get('requested_channels') or [])}",
        f"- published: {record.get('published')}",
        f"- scheduled: {record.get('scheduled')}",
    ]
    for platform in PLATFORMS:
        item = record.get("buffer", {}).get(platform, {})
        post_id = item.get("buffer_post_id") if isinstance(item, dict) else ""
        status = item.get("status") if isinstance(item, dict) else ""
        rows.append(f"- {platform}: {status or '-'} {post_id or ''}".rstrip())
    lines = ["# Social Release", "", *rows, "", "## Posts", ""]
    for platform in outcome.requested_channels:
        lines.extend([f"### {platform}", "", outcome.posts.get(platform, ""), ""])
    write_text(path, "\n".join(lines))


def classify_graphql_error(messages: list[str]) -> str:
    joined = " ".join(messages).lower()
    if "unauthor" in joined or "invalid token" in joined or "authentication" in joined:
        return "graphql_unauthorized"
    if "forbidden" in joined or "permission" in joined:
        return "graphql_forbidden"
    if "rate" in joined and "limit" in joined:
        return "graphql_rate_limited"
    return "graphql_error"


def graphql_request(api_key: str, query: str, variables: dict[str, Any] | None = None) -> dict[str, Any]:
    body = json.dumps({"query": query, "variables": variables or {}}, ensure_ascii=False).encode("utf-8")
    request = urllib.request.Request(
        BUFFER_GRAPHQL_ENDPOINT,
        data=body,
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
            "User-Agent": "DAKE-release-social/3.0",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=GRAPHQL_TIMEOUT) as response:
            status = int(getattr(response, "status", 200))
            raw = response.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as exc:
        if exc.code == 401:
            raise BufferGraphQLError("http_unauthorized") from exc
        if exc.code == 403:
            raise BufferGraphQLError("graphql_forbidden") from exc
        if exc.code == 429:
            raise BufferGraphQLError("graphql_rate_limited") from exc
        raise BufferGraphQLError("http_error", f"HTTP {exc.code}") from exc
    except Exception as exc:
        raise BufferGraphQLError("network_error") from exc
    if not 200 <= status < 300:
        raise BufferGraphQLError("http_error", f"HTTP {status}")
    try:
        loaded = json.loads(raw) if raw else {}
    except json.JSONDecodeError as exc:
        raise BufferGraphQLError("response_parse_error") from exc
    if not isinstance(loaded, dict):
        raise BufferGraphQLError("response_parse_error")
    errors = loaded.get("errors")
    if errors:
        messages = []
        if isinstance(errors, list):
            for item in errors:
                if isinstance(item, dict):
                    messages.append(str(item.get("message") or ""))
        raise BufferGraphQLError(classify_graphql_error(messages))
    data = loaded.get("data")
    if not isinstance(data, dict):
        raise BufferGraphQLError("response_parse_error")
    return data


def select_organization(api_key: str) -> tuple[str, int]:
    data = graphql_request(api_key, GET_ORGANIZATIONS_QUERY)
    orgs = data.get("account", {}).get("organizations", [])
    if not isinstance(orgs, list):
        raise BufferGraphQLError("response_parse_error")
    if not orgs:
        raise BufferGraphQLError("organization_not_found")
    override = read_plain_env("BUFFER_ORGANIZATION_ID")
    if len(orgs) == 1:
        org_id = str(orgs[0].get("id") or "")
        if not org_id:
            raise BufferGraphQLError("organization_not_found")
        return org_id, len(orgs)
    if override:
        for org in orgs:
            org_id = str(org.get("id") or "")
            if org_id == override:
                return org_id, len(orgs)
    raise BufferGraphQLError("multiple_organizations")


def service_for_channel(channel: dict[str, Any]) -> str:
    return str(channel.get("service") or "").strip().lower()


def usable_channel(channel: dict[str, Any]) -> bool:
    return not bool(channel.get("isDisconnected")) and not bool(channel.get("isLocked"))


def discover_channels(api_key: str, organization_id: str) -> tuple[dict[str, str], dict[str, dict[str, Any]]]:
    data = graphql_request(api_key, GET_CHANNELS_QUERY, {"organizationId": organization_id})
    channels = data.get("channels", [])
    if not isinstance(channels, list):
        raise BufferGraphQLError("response_parse_error")
    selected: dict[str, str] = {}
    summary: dict[str, dict[str, Any]] = {}
    for platform in PLATFORMS:
        service_names = SERVICE_NAMES[platform]
        matches = [item for item in channels if isinstance(item, dict) and service_for_channel(item) in service_names]
        usable = [item for item in matches if usable_channel(item)]
        override = read_plain_env(CHANNEL_ENV[platform])
        chosen: dict[str, Any] | None = None
        status = "unavailable"
        if override:
            for item in usable:
                if str(item.get("id") or "") == override:
                    chosen = item
                    break
            status = "found" if chosen else "env_channel_not_available"
        elif len(usable) == 1:
            chosen = usable[0]
            status = "found"
        elif len(usable) > 1:
            status = "multiple_channels"
        elif matches:
            status = "channel_unavailable"
        else:
            status = "channel_not_found"
        if chosen:
            channel_id = str(chosen.get("id") or "")
            if channel_id:
                selected[platform] = channel_id
            else:
                status = "channel_not_found"
            queue_paused = bool(chosen.get("isQueuePaused"))
        else:
            queue_paused = False
        summary[platform] = {
            "status": status,
            "queue_paused": queue_paused,
            "match_count": len(matches),
            "usable_count": len(usable),
        }
    return selected, summary


def discover_buffer(api_key: str) -> tuple[str, dict[str, str], dict[str, dict[str, Any]], int]:
    organization_id, organization_count = select_organization(api_key)
    channels, channel_summary = discover_channels(api_key, organization_id)
    return organization_id, channels, channel_summary, organization_count


def print_discovery(summary: dict[str, dict[str, Any]], organization_count: int, requested_channels: list[str]) -> None:
    print("Buffer GraphQL discovery")
    print("api_key: present")
    print("auth: success")
    print(f"organization_count: {organization_count}")
    for platform in requested_channels:
        item = summary.get(platform, {})
        status = item.get("status", "unavailable")
        queue = "yes" if item.get("queue_paused") else "no"
        print(f"{platform}: {status} queue_paused={queue}")


def create_draft_post(
    api_key: str,
    channel_id: str,
    text: str,
    platform: str,
    image_url: str = "",
    display_name: str = "",
) -> str:
    # Buffer requires a ShareMode enum even for drafts; saveToDraft keeps this out of queue/publish flow.
    post_input: dict[str, Any] = {
        "text": text,
        "channelId": channel_id,
        "schedulingType": "automatic",
        "mode": "addToQueue",
        "saveToDraft": True,
        "assets": [],
    }
    if platform == "instagram":
        if not image_url:
            raise BufferGraphQLError("instagram_asset_url_not_found")
        post_input["metadata"] = {"instagram": {"type": "post", "shouldShareToFeed": True}}
        post_input["assets"] = [instagram_asset(image_url, display_name or "DAKE app")]
    variables = {"input": post_input}
    data = graphql_request(api_key, CREATE_DRAFT_POST_MUTATION, variables)
    result = data.get("createPost")
    if not isinstance(result, dict):
        raise BufferGraphQLError("mutation_error")
    if result.get("message"):
        raise BufferGraphQLError("mutation_error")
    post = result.get("post")
    if not isinstance(post, dict):
        raise BufferGraphQLError("mutation_error")
    post_id = str(post.get("id") or "")
    if not post_id:
        raise BufferGraphQLError("mutation_error")
    return post_id


def edit_scheduled_post(api_key: str, post_id: str, due_at: str) -> str:
    variables = {
        "input": {
            "id": post_id,
            "mode": "customScheduled",
            "dueAt": due_at,
            "saveToDraft": False,
        }
    }
    data = graphql_request(api_key, EDIT_POST_MUTATION, variables)
    result = data.get("editPost")
    if not isinstance(result, dict):
        raise BufferGraphQLError("mutation_error")
    if result.get("message"):
        raise BufferGraphQLError("mutation_error")
    post = result.get("post")
    if not isinstance(post, dict):
        raise BufferGraphQLError("mutation_error")
    returned_id = str(post.get("id") or "")
    if not returned_id:
        raise BufferGraphQLError("mutation_error")
    return returned_id


def existing_buffer_post_id(record: dict[str, Any], platform: str) -> str:
    item = record.get("buffer", {}).get(platform, {})
    if not isinstance(item, dict):
        return ""
    for key in ("buffer_post_id", "buffer_update_id", "update_id", "id"):
        value = item.get(key)
        if value:
            return str(value)
    return ""


def read_existing_record(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    try:
        loaded = json.loads(read_text(path))
    except Exception:
        return {}
    return loaded if isinstance(loaded, dict) else {}


def reason_from_platforms(items: dict[str, dict[str, Any]], requested_channels: list[str], prefix: str = "") -> str:
    parts: list[str] = []
    for platform in requested_channels:
        item = items.get(platform, {})
        if item.get("buffer_post_id"):
            continue
        reason = str(item.get("reason") or item.get("status") or "missing_buffer_post_id")
        parts.append(f"{platform}: {reason}")
    return (prefix + "; " if prefix and parts else prefix) + "; ".join(parts)


def safe_buffer_item(platform: str, status: str, *, success: bool = False, post_id: str = "", reason: str = "") -> dict[str, Any]:
    item: dict[str, Any] = {
        "channel": platform,
        "status": status,
        "success": success,
        "buffer_post_id": post_id,
    }
    if reason:
        item["reason"] = reason
    return item


def schedule_status_for(platform: str, requested_channels: list[str], existing: dict[str, Any], buffer_item: dict[str, Any]) -> str:
    existing_status = existing.get("social_schedule_status", {})
    if isinstance(existing_status, dict):
        value = str(existing_status.get(platform) or "")
        if value in SCHEDULE_STATUS_VALUES and value not in {"not_requested", ""}:
            return value
    if platform not in requested_channels:
        return "not_requested"
    if buffer_item.get("buffer_post_id"):
        return "draft"
    return "failed" if buffer_item.get("status") == "failed" else "draft"


def build_record(
    app_dir: Path,
    source: Any,
    outcome: SocialOutcome,
    existing: dict[str, Any],
    args: argparse.Namespace,
) -> dict[str, Any]:
    product_id = product_id_for(app_dir, outcome.meta)
    product_type = product_type_for(outcome.meta)
    completed_channels = [
        platform for platform in outcome.requested_channels if outcome.buffer.get(platform, {}).get("buffer_post_id")
    ]
    skipped_channels = [platform for platform in PLATFORMS if platform not in outcome.requested_channels]
    schedule_status = {
        platform: schedule_status_for(platform, outcome.requested_channels, existing, outcome.buffer.get(platform, {}))
        for platform in PLATFORMS
    }
    record = {
        "product_id": product_id,
        "product_type": product_type,
        "app_key": str(outcome.meta.get("app_key") or app_dir.name),
        "display_name": str(outcome.meta.get("display_name") or outcome.meta.get("site_title") or app_dir.name),
        "app_url": outcome.app_url,
        "created_at": now_iso(),
        "stage": outcome.stage,
        "reason": outcome.reason,
        "social_posts_path": "release_artifacts/social_posts.md",
        "link_policy": "dakeapp.com detail page only",
        "dry_run": not args.post_to_buffer,
        "save_to_draft": True,
        "published": False,
        "scheduled": False,
        "requested_channels": outcome.requested_channels,
        "completed_channels": completed_channels,
        "skipped_channels": skipped_channels,
        "social_schedule_status": schedule_status,
        "posts": {**existing.get("posts", {}), **outcome.posts} if isinstance(existing.get("posts"), dict) else outcome.posts,
        "buffer": outcome.buffer,
        "source_policy": {
            "source_kind": source.source_kind,
            "source_path": source.source_label,
            "original_missing": source.original_missing,
            "meta_derivative_mismatch": source.derivative_mismatches,
        },
        "tool": {"name": "tools/release_social.py", "version": 4},
    }
    existing_schedule = existing.get("scheduled_posts")
    if isinstance(existing_schedule, dict):
        record["scheduled_posts"] = existing_schedule
    return record


def create_social(app_dir: Path, args: argparse.Namespace) -> SocialOutcome:
    requested_channels = args.requested_channels
    source = read_product_source(app_dir, ROOT)
    meta = source.meta
    meta_error = source.error
    app_url = product_url_for(app_dir, meta, args.site_root)
    posts = build_posts(meta, app_url, requested_channels)
    display_name = str(meta.get("display_name") or meta.get("site_title") or app_dir.name)
    instagram_image_url = args.thumbnail_url if str(args.thumbnail_url).startswith("https://") else public_site_image_url(app_dir, app_url, args.site_root)
    outcome = SocialOutcome(app_dir=app_dir, meta=meta, app_url=app_url, posts=posts, requested_channels=requested_channels)
    release_dir = social_artifact_dir(app_dir, meta)
    release_dir.mkdir(parents=True, exist_ok=True)
    posts_path = release_dir / "social_posts.md"
    release_path = release_dir / "social_release.json"
    release_md_path = release_dir / "social_release.md"
    existing = read_existing_record(release_path)
    existing_posts = read_existing_social_posts(posts_path)

    if not args.no_write:
        write_posts_markdown(posts_path, outcome, existing_posts)

    reasons: list[str] = []
    if meta_error:
        reasons.append(meta_error)

    if args.discover_only:
        api_key = read_secret_env("BUFFER_API_KEY")
        if not api_key:
            raise BufferGraphQLError("missing_buffer_api_key")
        _, _, summary, organization_count = discover_buffer(api_key)
        print_discovery(summary, organization_count, requested_channels)
        outcome.stage = "dry_run"
        outcome.reason = "discover_only"
        return outcome

    existing_buffer = existing.get("buffer") if isinstance(existing.get("buffer"), dict) else {}
    for platform in PLATFORMS:
        item = existing_buffer.get(platform, {}) if isinstance(existing_buffer, dict) else {}
        if isinstance(item, dict) and platform not in requested_channels:
            outcome.buffer[platform] = item
        elif platform not in requested_channels:
            outcome.buffer[platform] = safe_buffer_item(platform, "not_requested", reason="not_requested")

    if args.post_to_buffer:
        if not args.save_to_draft:
            reasons.append("save_to_draft_required")
        api_key = read_secret_env("BUFFER_API_KEY")
        if not api_key:
            reasons.append("missing_buffer_api_key")
        else:
            try:
                _, channels, channel_summary, organization_count = discover_buffer(api_key)
                print_discovery(channel_summary, organization_count, requested_channels)
            except BufferGraphQLError as exc:
                channels = {}
                channel_summary = {}
                reasons.append(exc.code)
        for platform in requested_channels:
            existing_id = existing_buffer_post_id(existing, platform)
            if existing_id and not args.force_repost:
                outcome.buffer[platform] = safe_buffer_item(platform, "skipped_existing", success=True, post_id=existing_id)
                continue
            if not api_key:
                outcome.buffer[platform] = safe_buffer_item(platform, "failed", reason="missing_buffer_api_key")
                continue
            channel_id = channels.get(platform, "")
            if not channel_id:
                status = channel_summary.get(platform, {}).get("status", "channel_not_found") if "channel_summary" in locals() else "channel_not_found"
                outcome.buffer[platform] = safe_buffer_item(platform, "failed", reason=str(status))
                continue
            try:
                post_id = create_draft_post(
                    api_key,
                    channel_id,
                    posts[platform],
                    platform,
                    instagram_image_url,
                    display_name,
                )
                outcome.buffer[platform] = safe_buffer_item(platform, "draft_created", success=True, post_id=post_id)
            except BufferGraphQLError as exc:
                outcome.buffer[platform] = safe_buffer_item(platform, "failed", reason=exc.code)
                reasons.append(f"{platform}: {exc.code}")
    else:
        for platform in requested_channels:
            existing_id = existing_buffer_post_id(existing, platform)
            if existing_id:
                existing_item = existing_buffer.get(platform, {}) if isinstance(existing_buffer, dict) else {}
                status = str(existing_item.get("status") or "skipped_existing") if isinstance(existing_item, dict) else "skipped_existing"
                outcome.buffer[platform] = safe_buffer_item(platform, status, success=True, post_id=existing_id)
            else:
                outcome.buffer[platform] = safe_buffer_item(platform, "dry_run", reason="dry_run")

    has_requested_ids = all(outcome.buffer.get(platform, {}).get("buffer_post_id") for platform in requested_channels)
    if args.post_to_buffer:
        outcome.stage = "complete" if has_requested_ids and not reasons else "failed"
        outcome.reason = "" if outcome.stage == "complete" else reason_from_platforms(outcome.buffer, requested_channels, "; ".join(reasons))
    else:
        outcome.stage = "complete" if existing.get("stage") == "complete" and has_requested_ids else "dry_run"
        outcome.reason = "" if outcome.stage == "complete" else "Buffer posting not requested; generated selected social post draft text only"

    record = build_record(app_dir, source, outcome, existing, args)
    if not args.no_write:
        write_json(release_path, record)
        write_social_release_markdown(release_md_path, outcome, record)
    return outcome


def available_app(app_dir: Path) -> bool:
    source = read_product_source(app_dir, ROOT)
    return str(source.meta.get("status") or "unknown") == SHIPPING_STATUS


def social_ready_for(app_dir: Path, requested_channels: list[str]) -> bool:
    source = read_product_source(app_dir, ROOT)
    record = read_existing_record(social_artifact_dir(app_dir, source.meta) / "social_release.json")
    if not record:
        return False
    channels = channels_from_record(record) if "requested_channels" in record else requested_channels
    channels = [channel for channel in requested_channels if channel in channels] or requested_channels
    return str(record.get("stage") or "") == "complete" and all(existing_buffer_post_id(record, channel) for channel in channels)


def all_v2_pending_targets(requested_channels: list[str], limit: int) -> list[Path]:
    targets: list[Path] = []
    for app_dir in app_dirs():
        if not available_app(app_dir):
            continue
        if social_ready_for(app_dir, requested_channels):
            continue
        targets.append(app_dir)
    return targets[:limit]


def print_batch_targets(targets: list[Path], requested_channels: list[str], post_to_buffer: bool) -> None:
    mode = "buffer_draft_create" if post_to_buffer else "dry_run_list_only"
    print("DAKE social batch targets")
    print(f"mode: {mode}")
    print(f"requested_channels: {','.join(requested_channels)}")
    print(f"target_count: {len(targets)}")
    for index, app_dir in enumerate(targets, 1):
        print(f"{index}. {app_dir.name}")


def parse_schedule_datetime(value: str, timezone_name: str) -> tuple[datetime | None, str]:
    if not value:
        return None, "scheduled_at_missing"
    try:
        parsed = datetime.fromisoformat(value)
    except ValueError:
        return None, "scheduled_at_invalid"
    try:
        zone = ZoneInfo(timezone_name)
    except Exception:
        return None, "timezone_invalid"
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=zone)
    local = parsed.astimezone(zone)
    if local <= datetime.now(zone):
        return None, "scheduled_at_past"
    return local, ""


def load_schedule_plan(path: Path) -> tuple[str, list[dict[str, Any]]]:
    if not path.exists():
        raise SchedulePlanError("schedule_plan_not_found")
    try:
        loaded = json.loads(read_text(path))
    except Exception as exc:
        raise SchedulePlanError(f"schedule_plan_invalid_json: {exc}") from exc
    if not isinstance(loaded, dict):
        raise SchedulePlanError("schedule_plan_not_object")
    timezone_name = str(loaded.get("timezone") or "")
    if not timezone_name:
        raise SchedulePlanError("timezone_missing")
    items = loaded.get("items")
    if not isinstance(items, list):
        raise SchedulePlanError("items_missing")
    return timezone_name, [item for item in items if isinstance(item, dict)]


def schedule_record_for(record: dict[str, Any], platform: str) -> dict[str, Any]:
    scheduled = record.get("scheduled_posts", {})
    if not isinstance(scheduled, dict):
        return {}
    item = scheduled.get(platform, {})
    return item if isinstance(item, dict) else {}


def validate_schedule_plan(path: Path, confirm: bool, site_root: Path) -> tuple[list[ScheduleItem], list[str]]:
    timezone_name, raw_items = load_schedule_plan(path)
    errors: list[str] = []
    items: list[ScheduleItem] = []
    seen_pairs: set[tuple[str, str]] = set()
    time_counts: Counter[str] = Counter()
    for index, raw in enumerate(raw_items, 1):
        app_key = str(raw.get("app_key") or "").strip()
        if not app_key:
            errors.append(f"item {index}: app_key_missing")
            continue
        try:
            app_dir = find_app(app_key)
        except Exception:
            errors.append(f"item {index}: app_not_found")
            continue
        try:
            channels = [normalize_channel(str(channel)) for channel in raw.get("channels", [])]
        except Exception as exc:
            errors.append(f"item {index}: {exc}")
            continue
        channels = list(dict.fromkeys(channels))
        if not channels:
            errors.append(f"item {index}: channels_missing")
            continue
        scheduled_at_raw = str(raw.get("scheduled_at") or "")
        scheduled_at, error = parse_schedule_datetime(scheduled_at_raw, timezone_name)
        if error or scheduled_at is None:
            errors.append(f"item {index}: {error}")
            continue
        due_at = scheduled_at.astimezone(timezone.utc).isoformat(timespec="seconds")
        source = read_product_source(app_dir, ROOT)
        release_path = social_artifact_dir(app_dir, source.meta) / "social_release.json"
        record = read_existing_record(release_path)
        app_url = product_url_for(app_dir, source.meta, site_root)
        for channel in channels:
            pair = (app_dir.name, channel)
            if pair in seen_pairs:
                errors.append(f"item {index}: duplicate_app_channel:{app_dir.name}:{channel}")
                continue
            seen_pairs.add(pair)
            time_counts[scheduled_at.isoformat(timespec="minutes")] += 1
            post_id = existing_buffer_post_id(record, channel)
            if not post_id:
                errors.append(f"item {index}: missing_buffer_post_id:{app_dir.name}:{channel}")
            if channel == "instagram" and not public_site_image_url(app_dir, app_url, site_root):
                errors.append(f"item {index}: instagram_asset_url_not_found:{app_dir.name}")
            current_schedule = schedule_record_for(record, channel)
            current_due_at = str(current_schedule.get("due_at") or "")
            if current_due_at and current_due_at == due_at:
                pass
            items.append(ScheduleItem(app_dir=app_dir, app_key=app_key, channels=[channel], scheduled_at=scheduled_at.isoformat(), due_at=due_at))
    for scheduled_at, count in time_counts.items():
        if count > 3:
            errors.append(f"schedule_time_overconcentrated:{scheduled_at}:{count}")
    if confirm:
        errors.append("schedule_capacity_check_unavailable")
    return items, errors


def print_schedule_plan(items: list[ScheduleItem], errors: list[str], confirm: bool) -> None:
    print("DAKE social schedule plan")
    print(f"mode: {'confirm_requested' if confirm else 'dry_run'}")
    print(f"items: {len(items)}")
    for item in items:
        print(f"- {item.app_dir.name}: {','.join(item.channels)} at {item.scheduled_at}")
    if errors:
        print("errors:")
        for error in errors:
            print(f"- {error}")
    if not confirm:
        print("no Buffer mutation: --confirm-schedule was not supplied")


def apply_schedule_plan(path: Path, args: argparse.Namespace) -> int:
    items, errors = validate_schedule_plan(path, args.confirm_schedule, args.site_root)
    print_schedule_plan(items, errors, args.confirm_schedule)
    if not args.confirm_schedule:
        return 0 if not errors else 1
    if errors:
        print("schedule not applied")
        return 1
    api_key = read_secret_env("BUFFER_API_KEY")
    if not api_key:
        print("missing_buffer_api_key")
        return 1
    # Current safety policy stops before mutation unless schedule capacity can be verified.
    print("schedule not applied: schedule_capacity_check_unavailable")
    return 1


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate DAKE SNS posts and optionally create Buffer draft posts.")
    parser.add_argument("--app", action="append", default=[])
    parser.add_argument("--product", action="append", default=[], help="Product id or folder name; accepts apps and packs.")
    parser.add_argument("--all-v2-pending", action="store_true")
    parser.add_argument("--limit", type=int)
    parser.add_argument("--channels", default="")
    parser.add_argument("--site-root", type=Path, default=DEFAULT_SITE_ROOT)
    parser.add_argument("--post-to-buffer", action="store_true")
    parser.add_argument("--save-to-draft", action=argparse.BooleanOptionalAction, default=True)
    parser.add_argument("--discover-only", action="store_true")
    parser.add_argument("--force-repost", action="store_true")
    parser.add_argument("--thumbnail-url", default="", help="reserved for future media support")
    parser.add_argument("--fail-on-incomplete", action="store_true")
    parser.add_argument("--apply-schedule", type=Path)
    parser.add_argument("--confirm-schedule", action="store_true")
    parser.add_argument("--no-write", action="store_true", help="dry-run without updating social artifacts")
    args = parser.parse_args()
    args.requested_channels = parse_channels(args.channels)
    if args.confirm_schedule and not args.apply_schedule:
        parser.error("--confirm-schedule requires --apply-schedule")
    if args.apply_schedule and (args.app or args.product or args.all_v2_pending or args.post_to_buffer or args.discover_only):
        parser.error("--apply-schedule cannot be combined with app draft generation options")
    if args.all_v2_pending:
        if args.limit is None:
            parser.error("--all-v2-pending requires --limit")
        if args.limit < 1 or args.limit > MAX_BATCH_LIMIT:
            parser.error(f"--limit must be between 1 and {MAX_BATCH_LIMIT}")
    if not args.apply_schedule and not args.all_v2_pending and not args.app and not args.product:
        parser.error("--app, --product, --all-v2-pending, or --apply-schedule is required")
    if args.post_to_buffer and not args.save_to_draft:
        parser.error("normal shipping requires --save-to-draft")
    return args


def main() -> int:
    args = parse_args()
    if args.apply_schedule:
        return apply_schedule_plan(args.apply_schedule, args)

    app_targets: list[Path]
    if args.all_v2_pending:
        app_targets = all_v2_pending_targets(args.requested_channels, args.limit)
        print_batch_targets(app_targets, args.requested_channels, args.post_to_buffer)
        if not args.post_to_buffer:
            print("no Buffer mutation: --post-to-buffer was not supplied")
            return 0
    else:
        app_targets = [find_app(item) for item in [*args.app, *args.product]]

    outcomes: list[SocialOutcome] = []
    try:
        for app_dir in app_targets:
            outcome = create_social(app_dir, args)
            outcomes.append(outcome)
            print(f"{outcome.app_dir.name}: {outcome.stage} {outcome.reason}")
    except BufferGraphQLError as exc:
        print(exc.code)
        return 1
    if args.fail_on_incomplete and any(outcome.stage != "complete" for outcome in outcomes):
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
