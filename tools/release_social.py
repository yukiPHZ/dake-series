# -*- coding: utf-8 -*-
"""Generate DAKE release social posts and optionally create Buffer draft posts.

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
import urllib.request
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

from release_source_policy import app_url_for, find_app as source_find_app, read_app_source


ROOT = Path(__file__).resolve().parents[1]
APPS_DIR = ROOT / "01_apps"
DEFAULT_SITE_ROOT = Path(os.environ.get("DAKEAPP_SITE_ROOT", r"C:\Users\yukiz\devlop\dakeapp-site"))
JST = timezone(timedelta(hours=9))
PLATFORMS = ("x", "threads", "instagram")
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


@dataclass
class SocialOutcome:
    app_dir: Path
    meta: dict[str, Any]
    app_url: str
    posts: dict[str, str]
    buffer: dict[str, dict[str, Any]] = field(default_factory=dict)
    stage: str = "failed"
    reason: str = ""


class BufferGraphQLError(Exception):
    def __init__(self, code: str, detail: str = "") -> None:
        super().__init__(code)
        self.code = code
        self.detail = detail


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
        except FileNotFoundError:
            continue
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
        except FileNotFoundError:
            continue
        except OSError:
            continue
        value = str(found).strip()
        if value:
            return value
    return ""


def read_meta(app_dir: Path) -> tuple[dict[str, Any], str]:
    source = read_app_source(app_dir, ROOT)
    return source.meta, source.error


def find_app(identifier: str) -> Path:
    return source_find_app(APPS_DIR, identifier, ROOT)


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


def detail_slug_from_url(app_url: str) -> str:
    parts = app_url.rstrip("/").split("/")
    return parts[-1] if parts else ""


def public_site_image_url(app_dir: Path, app_url: str, site_root: Path) -> str:
    override = read_plain_env("BUFFER_INSTAGRAM_IMAGE_URL")
    if override.startswith("https://"):
        return override
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
            "User-Agent": "DAKE-release-social/2.0",
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


def create_draft_post(
    api_key: str,
    channel_id: str,
    text: str,
    platform: str,
    image_url: str = "",
    display_name: str = "",
) -> str:
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
    variables = {
        "input": post_input
    }
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


def reason_from_platforms(items: dict[str, dict[str, Any]], prefix: str = "") -> str:
    parts: list[str] = []
    for platform in PLATFORMS:
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


def discover_buffer(api_key: str) -> tuple[str, dict[str, str], dict[str, dict[str, Any]], int]:
    organization_id, organization_count = select_organization(api_key)
    channels, channel_summary = discover_channels(api_key, organization_id)
    return organization_id, channels, channel_summary, organization_count


def print_discovery(summary: dict[str, dict[str, Any]], organization_count: int) -> None:
    print("Buffer GraphQL discovery")
    print("api_key: present")
    print("auth: success")
    print(f"organization_count: {organization_count}")
    for platform in PLATFORMS:
        item = summary.get(platform, {})
        status = item.get("status", "unavailable")
        queue = "yes" if item.get("queue_paused") else "no"
        print(f"{platform}: {status} queue_paused={queue}")


def create_social(app_dir: Path, args: argparse.Namespace) -> SocialOutcome:
    source = read_app_source(app_dir, ROOT)
    meta = source.meta
    meta_error = source.error
    app_url = app_url_for(app_dir, meta, args.site_root)
    posts = build_posts(meta, app_url)
    display_name = str(meta.get("display_name") or meta.get("site_title") or app_dir.name)
    instagram_image_url = public_site_image_url(app_dir, app_url, args.site_root)
    outcome = SocialOutcome(app_dir=app_dir, meta=meta, app_url=app_url, posts=posts)
    release_dir = app_dir / "release_artifacts"
    release_dir.mkdir(parents=True, exist_ok=True)
    posts_path = release_dir / "social_posts.md"
    release_path = release_dir / "social_release.json"
    write_posts_markdown(posts_path, outcome)

    reasons: list[str] = []
    if meta_error:
        reasons.append(meta_error)

    if args.discover_only:
        api_key = read_secret_env("BUFFER_API_KEY")
        if not api_key:
            raise BufferGraphQLError("missing_buffer_api_key")
        _, _, summary, organization_count = discover_buffer(api_key)
        print_discovery(summary, organization_count)
        outcome.stage = "dry_run"
        outcome.reason = "discover_only"
        return outcome

    existing = read_existing_record(release_path)
    if args.post_to_buffer:
        if not args.save_to_draft:
            reasons.append("save_to_draft_required")
        api_key = read_secret_env("BUFFER_API_KEY")
        if not api_key:
            reasons.append("missing_buffer_api_key")
        else:
            try:
                _, channels, channel_summary, organization_count = discover_buffer(api_key)
                print_discovery(channel_summary, organization_count)
            except BufferGraphQLError as exc:
                channels = {}
                channel_summary = {}
                reasons.append(exc.code)
            for platform in PLATFORMS:
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
        for platform in PLATFORMS:
            outcome.buffer[platform] = safe_buffer_item(platform, "dry_run", reason="dry_run")

    has_all_ids = all(outcome.buffer.get(platform, {}).get("buffer_post_id") for platform in PLATFORMS)
    if args.post_to_buffer:
        outcome.stage = "complete" if has_all_ids and not reasons else "failed"
        outcome.reason = "" if outcome.stage == "complete" else reason_from_platforms(outcome.buffer, "; ".join(reasons))
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
        "save_to_draft": True,
        "posts": outcome.posts,
        "buffer": outcome.buffer,
        "source_policy": {
            "source_kind": source.source_kind,
            "source_path": source.source_label,
            "original_missing": source.original_missing,
            "meta_derivative_mismatch": source.derivative_mismatches,
        },
        "tool": {"name": "tools/release_social.py", "version": 2},
    }
    write_json(release_path, record)
    return outcome


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate DAKE SNS posts and optionally create Buffer draft posts.")
    parser.add_argument("--app", action="append", required=True)
    parser.add_argument("--site-root", type=Path, default=DEFAULT_SITE_ROOT)
    parser.add_argument("--post-to-buffer", action="store_true")
    parser.add_argument("--save-to-draft", action=argparse.BooleanOptionalAction, default=True)
    parser.add_argument("--discover-only", action="store_true")
    parser.add_argument("--force-repost", action="store_true")
    parser.add_argument("--thumbnail-url", default="", help="reserved for future media support")
    parser.add_argument("--fail-on-incomplete", action="store_true")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    outcomes: list[SocialOutcome] = []
    try:
        outcomes = [create_social(find_app(item), args) for item in args.app]
    except BufferGraphQLError as exc:
        print(exc.code)
        return 1
    for outcome in outcomes:
        print(f"{outcome.app_dir.name}: {outcome.stage} {outcome.reason}")
    if args.fail_on_incomplete and any(outcome.stage != "complete" for outcome in outcomes):
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
