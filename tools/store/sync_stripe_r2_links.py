"""Batch-sync DAKE Store Wave 1 (47 items) Stripe Payment Links for R2 delivery
and archive discontinued (C-list) Stripe Payment Links.

Safety rules:
- Default mode is --dry-run (zero Stripe writes, works offline without API key).
- Live writes happen ONLY when --apply is explicitly specified.
- STRIPE_SECRET_KEY is read strictly from os.environ["STRIPE_SECRET_KEY"] and
  is never written to files, Git, or stdout/stderr logs.
- Zero third-party dependencies (uses Python stdlib urllib.request).
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[2]
DEFAULT_STORE_SITE = ROOT.parent / "dake-store-site"
DEFAULT_DAKEAPP_SITE = ROOT.parent / "dakeapp-site"
DEFAULT_SHIMARISU_ROOT = ROOT.parent / "SHIMARISU"
DEFAULT_REPORT_PATH = ROOT / "tmp" / "stripe_sync_report.json"

STRIPE_API_BASE = "https://api.stripe.com/v1"
R2_REDIRECT_URL = "https://store.dakeapp.com/download/?session_id={CHECKOUT_SESSION_ID}"
DELIVERY_METHOD_R2 = "stripe_session_private_r2"
DEFAULT_TAX_CODE = "txcd_10202003"
JST = timezone(timedelta(hours=9))
SECRET_PATTERN = re.compile(r"(?:sk|rk)_(?:live|test)_[A-Za-z0-9_]{8,}|whsec_[A-Za-z0-9_]{8,}")

# Pre-unification prices of existing buy.stripe.com Payment Links.
# Used ONLY for offline --dry-run planning when Stripe API is not queried.
# In --apply (or --dry-run --online), live prices are read directly from Stripe API.
LEGACY_STRIPE_PRICES: dict[str, int] = {
    "https://buy.stripe.com/eVq00j43D9kjc4Me250gw06": 500,   # DAKE_Backup
    "https://buy.stripe.com/aFacN56bL3ZZ3yg7DH0gw07": 500,   # DAKE_Git_Memo
    "https://buy.stripe.com/6oU6oH8jT3ZZd8Q1fj0gw08": 500,   # DAKE_Image_PasteA4
    "https://buy.stripe.com/fZu3cvbw57cbc4Mf690gw09": 300,   # DAKE_Mail_Address_Format
    "https://buy.stripe.com/aFabJ16bL3ZZ5GobTX0gw0W": 300,   # DAKE_Mail_Draft (R2 live 300 JPY -> 500 JPY)
    "https://buy.stripe.com/4gMbJ143D2VVd8Q9LP0gw0b": 300,   # DAKE_Maji_Memo
    "https://buy.stripe.com/9B65kD2Zzaon9WE0bf0gw0c": 300,   # DAKE_Sticky_Memo
    "https://buy.stripe.com/3cIaEXcA9543b0I8HL0gw0d": 300,   # DAKE_Yesterday_Task_Memo
    "https://buy.stripe.com/7sY6oHbw5gMLc4M4rv0gw0f": 300,   # dake_column_memo
    "https://buy.stripe.com/fZubJ12Zzbsrd8Q3nr0gw0g": 500,   # dake_document_cover
    "https://buy.stripe.com/4gM00j9nX1RRb0I5vz0gw0h": 500,   # dake_fax_cover
    "https://buy.stripe.com/14A5kDcA9cwv4Ck2jn0gw0i": 300,   # dake_folder_list
    "https://buy.stripe.com/4gMfZhas1gML4CkaPT0gw0j": 500,   # dake_image_heictojpg
    "https://buy.stripe.com/cNicN5fMlgML9WE1fj0gw0k": 500,   # dake_image_iphonetopc
    "https://buy.stripe.com/5kQ5kD7fP8gf7Owf690gw03": 500,   # dake_image_resize
    "https://buy.stripe.com/8x25kDas13ZZd8Q7DH0gw0m": 500,   # dake_image_topdf
    "https://buy.stripe.com/5kQ9ATgQp0NN5Go6zD0gw0o": 300,   # dake_mail_allstaff
    "https://buy.stripe.com/aFaeVd8jTdAz4Ck2jn0gw0q": 300,   # dake_mail_list
    "https://buy.stripe.com/bJefZh57H7cb7OwcY10gw0r": 500,   # dake_mansion_schedule
    "https://buy.stripe.com/fZu7sLfMl1RRb0I2jn0gw0s": 500,   # dake_pdf_checkstamp
    "https://buy.stripe.com/aFa6oHeIh3ZZ6Kse250gw02": 500,   # dake_pdf_compress
    "https://buy.stripe.com/00w7sL43D6877Ow4rv0gw0t": 500,   # dake_pdf_crop
    "https://buy.stripe.com/bJe9AT6bL2VV0m46zD0gw0u": 500,   # dake_pdf_lookhere
    "https://buy.stripe.com/fZu8wP2ZzeED3yg3nr0gw0v": 500,   # dake_pdf_marker
    "https://buy.stripe.com/9B6fZh0Rr2VV6Ksf690gw01": 500,   # dake_pdf_merge
    "https://buy.stripe.com/28E28r8jTfIHc4M6zD0gw0R": 500,   # dake_pdf_overview_rename
    "https://buy.stripe.com/aFa7sLgQp9kj1q85vz0gw0x": 500,   # dake_pdf_rename
    "https://buy.stripe.com/14AcN5fMldAzc4Mgad0gw0y": 500,   # dake_pdf_reorder
    "https://buy.stripe.com/28E14ndEd7cb0m47DH0gw0z": 500,   # dake_pdf_splitone
    "https://buy.stripe.com/3cI00j6bL8gf8SAcY10gw0A": 500,   # dake_pdf_splitselect
    "https://buy.stripe.com/aFadR96bL5435GobTX0gw0B": 500,   # dake_pdf_toimages
    "https://buy.stripe.com/cNi8wPcA9dAzc4M0bf0gw0C": 500,   # dake_pdf_viewer
    "https://buy.stripe.com/dRm3cvcA9543d8QcY10gw0D": 300,   # dake_price_apportionment
    "https://buy.stripe.com/bJefZh9nX7cbfgYgad0gw0E": 300,   # dake_price_fixedtax
    "https://buy.stripe.com/00waEXcA9687ecU0bf0gw0F": 500,   # dake_reform_progress
    "https://buy.stripe.com/3cIaEXas1cwv0m49LP0gw0G": 300,   # dake_screen_webp
    "https://buy.stripe.com/dRmeVd57H5435GobTX0gw0H": 500,   # dake_screenshot_print
    "https://buy.stripe.com/28EeVdas1eED1q8e250gw0I": 300,   # dake_two_person_memo
    "https://buy.stripe.com/eVq7sLdEd9kjb0I1fj0gw0J": 500,   # dake_work_calendar
    "https://buy.stripe.com/3cIfZh6bL2VV5Go9LP0gw0K": 300,   # dake_year_age
    "https://buy.stripe.com/28E7sLas1aonecU5vz0gw0L": 300,   # dake_year_notice
    "https://buy.stripe.com/5kQdR9eIh8gfgl25vz0gw00": 300,   # time_advanced_timer
    "https://buy.stripe.com/aFa7sL6bL9kj6Ks9LP0gw0O": 1480,  # DAKE_Pack_Document
    "https://buy.stripe.com/7sY14ncA90NN1q84rv0gw0Q": 780,   # DAKE_Pack_Mail
    "https://buy.stripe.com/4gM9AT57H7cbc4M3nr0gw04": 3000,  # SHIMARISU_Pack
}

# Discontinued (C-list) items with active Stripe Payment Links to archive.
DISCONTINUED_STRIPE_ITEMS: tuple[dict[str, str], ...] = (
    {
        "id": "dake_mail_kikuta",
        "title": "Dake菊田メール",
        "stripe_payment_link": "https://buy.stripe.com/4gM9ATas1cwvc4M4rv0gw0p",
        "reason": "C判定（個人専用・出品停止）",
    },
    {
        "id": "dake_booth_assist",
        "title": "BOOTHアシスト",
        "stripe_payment_link": "https://buy.stripe.com/dRmcN557H8gf0m4gad0gw0e",
        "reason": "C判定（制作者内部ツール・出品停止）",
    },
    {
        "id": "DAKE_Pack_Memo",
        "title": "DAKE メモと記録パック",
        "stripe_payment_link": "https://buy.stripe.com/14A4gz9nX5438SAcY10gw0P",
        "reason": "C判定（パック販売終了・構成単体は存続）",
    },
    {
        "id": "dake_pdf_merge_mini",
        "title": "DakePDF結合mini",
        "stripe_payment_link": "https://buy.stripe.com/00w6oH57H6870m48HL0gw0w",
        "reason": "C判定（フル版無料化に伴い停止）",
    },
    {
        "id": "dake_image_receiver",
        "title": "DakeImage_Receiver",
        "stripe_payment_link": "https://buy.stripe.com/28E4gz6bLdAz3yg1fj0gw0l",
        "reason": "C判定（iPhoneToPCへ一本化）",
    },
    {
        "id": "DAKE_App_Doko",
        "title": "アプリどこ",
        "stripe_payment_link": "https://buy.stripe.com/8x200j6bL1RRecU0bf0gw05",
        "reason": "C判定（開発環境寄りツールのため停止）",
    },
    {
        "id": "dake_launcher",
        "title": "Dakeランチャー",
        "stripe_payment_link": "https://buy.stripe.com/8x228r8jTgML5GoaPT0gw0n",
        "reason": "C判定（単体販売対象外）",
    },
    {
        "id": "game_alien_road",
        "title": "DakeAlien Road",
        "stripe_payment_link": "https://buy.stripe.com/8x2bJ1eIhfIH9WE7DH0gw0M",
        "reason": "C判定（ゲーム枠終了）",
    },
    {
        "id": "game_diver_catch",
        "title": "Dake潜って捕る",
        "stripe_payment_link": "https://buy.stripe.com/dRm5kD43DeED6Ks3nr0gw0N",
        "reason": "C判定（ゲーム枠終了）",
    },
)

PAYMENT_LINK_LINE_RE = re.compile(
    r"^(\s*[-*]\s*(?:Stripe Payment Link|stripe_payment_link|Payment Link|Stripe URL)\s*[:：]\s*)(.*?)\s*$",
    re.MULTILINE,
)
PAYMENT_STATUS_LINE_RE = re.compile(
    r"^(\s*[-*]\s*(?:payment_status|Store販売状態)\s*[:：]\s*)(.*?)\s*$",
    re.MULTILINE,
)


class StripeSyncError(RuntimeError):
    """Raised on Stripe API or validation failure."""


def now_jst() -> str:
    return datetime.now(JST).isoformat(timespec="seconds")


def sanitize_text(value: str) -> str:
    return SECRET_PATTERN.sub("[REDACTED_SECRET]", str(value))


def sanitize_obj(obj: Any) -> Any:
    if isinstance(obj, str):
        return sanitize_text(obj)
    if isinstance(obj, list):
        return [sanitize_obj(v) for v in obj]
    if isinstance(obj, dict):
        return {sanitize_text(str(k)): sanitize_obj(v) for k, v in obj.items()}
    return obj


def flatten_form_params(data: Any, prefix: str = "") -> list[tuple[str, str]]:
    """Encode nested dicts/lists into Stripe application/x-www-form-urlencoded pairs."""
    pairs: list[tuple[str, str]] = []
    if isinstance(data, dict):
        for key, value in data.items():
            if value is None:
                continue
            next_prefix = f"{prefix}[{key}]" if prefix else str(key)
            pairs.extend(flatten_form_params(value, next_prefix))
    elif isinstance(data, (list, tuple)):
        for index, value in enumerate(data):
            if value is None:
                continue
            next_prefix = f"{prefix}[{index}]"
            pairs.extend(flatten_form_params(value, next_prefix))
    elif isinstance(data, bool):
        pairs.append((prefix, "true" if data else "false"))
    else:
        pairs.append((prefix, str(data)))
    return pairs


def make_idempotency_key(prefix: str, item_id: str, payload: dict[str, Any]) -> str:
    raw = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    digest = hashlib.sha256(raw.encode("utf-8")).hexdigest()[:16]
    safe = re.sub(r"[^A-Za-z0-9_-]+", "_", item_id).strip("_") or "item"
    return f"dake-r2-{prefix}-{safe}-{digest}"


class StripeClient:
    """Minimal zero-dependency Stripe REST API client using urllib.request."""

    def __init__(self, secret_key: str, *, api_base: str = STRIPE_API_BASE, timeout: int = 30) -> None:
        key = secret_key.strip()
        if not key:
            raise StripeSyncError("STRIPE_SECRET_KEY environment variable is empty.")
        if key.startswith("sk_test_"):
            raise StripeSyncError("Refusing to run with sk_test_ key. Set a live sk_live_ STRIPE_SECRET_KEY.")
        if not key.startswith("sk_live_"):
            raise StripeSyncError("STRIPE_SECRET_KEY must start with sk_live_.")
        self._secret_key = key
        self.api_base = api_base.rstrip("/")
        self.timeout = timeout

    def _request(
        self,
        method: str,
        path: str,
        *,
        params: dict[str, Any] | None = None,
        idempotency_key: str | None = None,
    ) -> dict[str, Any]:
        url = f"{self.api_base}{path}"
        body_bytes: bytes | None = None
        headers = {
            "Authorization": f"Bearer {self._secret_key}",
            "User-Agent": "DAKE-Store-R2-Sync/1.0",
        }
        if method.upper() == "GET" and params:
            query = urllib.parse.urlencode(flatten_form_params(params))
            url = f"{url}?{query}"
        elif method.upper() == "POST":
            encoded = urllib.parse.urlencode(flatten_form_params(params or {}))
            body_bytes = encoded.encode("utf-8")
            headers["Content-Type"] = "application/x-www-form-urlencoded"
            if idempotency_key:
                headers["Idempotency-Key"] = idempotency_key

        last_error = ""
        for attempt in range(1, 4):
            req = urllib.request.Request(url, data=body_bytes, headers=headers, method=method.upper())
            try:
                with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                    payload = resp.read().decode("utf-8")
                    return json.loads(payload)
            except urllib.error.HTTPError as exc:
                raw = exc.read().decode("utf-8", errors="replace")
                last_error = sanitize_text(f"HTTP {exc.code} {method} {path}: {raw[:400]}")
                if exc.code in {429, 500, 502, 503, 504} and attempt < 3:
                    time.sleep(1.5 * attempt)
                    continue
                raise StripeSyncError(last_error) from None
            except urllib.error.URLError as exc:
                last_error = sanitize_text(f"Network error {method} {path}: {exc.reason}")
                if attempt < 3:
                    time.sleep(1.5 * attempt)
                    continue
                raise StripeSyncError(last_error) from None
        raise StripeSyncError(last_error)

    def paginate(self, path: str, params: dict[str, Any] | None = None) -> list[dict[str, Any]]:
        results: list[dict[str, Any]] = []
        query = dict(params or {})
        query.setdefault("limit", 100)
        while True:
            page = self._request("GET", path, params=query)
            data = page.get("data") or []
            results.extend(data)
            if not page.get("has_more") or not data:
                break
            query["starting_after"] = data[-1]["id"]
        return results

    def list_payment_links(self, *, active: bool | None = None) -> list[dict[str, Any]]:
        params: dict[str, Any] = {"limit": 100}
        if active is not None:
            params["active"] = active
        return self.paginate("/payment_links", params)

    def list_payment_link_line_items(self, payment_link_id: str) -> list[dict[str, Any]]:
        return self.paginate(f"/payment_links/{urllib.parse.quote(payment_link_id)}/line_items", {"limit": 100})

    def list_products(self) -> list[dict[str, Any]]:
        return self.paginate("/products", {"limit": 100})

    def list_prices(self, product_id: str) -> list[dict[str, Any]]:
        return self.paginate("/prices", {"product": product_id, "active": True, "limit": 100})

    def create_product(self, payload: dict[str, Any], idempotency_key: str) -> dict[str, Any]:
        return self._request("POST", "/products", params=payload, idempotency_key=idempotency_key)

    def create_price(self, payload: dict[str, Any], idempotency_key: str) -> dict[str, Any]:
        return self._request("POST", "/prices", params=payload, idempotency_key=idempotency_key)

    def create_payment_link(self, payload: dict[str, Any], idempotency_key: str) -> dict[str, Any]:
        return self._request("POST", "/payment_links", params=payload, idempotency_key=idempotency_key)

    def update_payment_link(self, payment_link_id: str, payload: dict[str, Any]) -> dict[str, Any]:
        return self._request("POST", f"/payment_links/{urllib.parse.quote(payment_link_id)}", params=payload)


def parse_private_download_ids(private_download_path: Path) -> set[str]:
    """Parse DESKTOP_DELIVERIES allowlist keys from functions/_lib/private-download.mjs."""
    text = private_download_path.read_text(encoding="utf-8")
    start_match = re.search(r"export\s+const\s+DESKTOP_DELIVERIES\s*=\s*Object\.freeze\(\{", text)
    if not start_match:
        raise StripeSyncError(f"DESKTOP_DELIVERIES not found in {private_download_path}")
    end_idx = text.find("});", start_match.end())
    if end_idx == -1:
        raise StripeSyncError(f"DESKTOP_DELIVERIES closing delimiter not found in {private_download_path}")
    block = text[start_match.end():end_idx]
    keys = set(re.findall(r'^\s*(?:"([A-Za-z0-9_]+)"|([A-Za-z0-9_]+))\s*:\s*Object\.freeze\(', block, re.MULTILINE))
    return {a or b for a, b in keys if (a or b)}


def resolve_original_path(source_original: str, *, repo_root: Path = ROOT) -> Path:
    candidate = Path(source_original)
    if candidate.is_absolute():
        return candidate
    return (repo_root / source_original).resolve()


def write_back_original_payment_link(original_path: Path, new_url: str) -> bool:
    """Atomically write back new Stripe Payment Link and stripe_ready status to ORIGINAL.md."""
    if not new_url.startswith("https://buy.stripe.com/") or "test_" in new_url:
        raise StripeSyncError(f"Refusing to write non-live Stripe Payment Link URL: {new_url}")
    if not original_path.exists():
        raise StripeSyncError(f"ORIGINAL.md not found for write-back: {original_path}")

    original = original_path.read_text(encoding="utf-8")
    updated = original

    if PAYMENT_LINK_LINE_RE.search(updated):
        updated = PAYMENT_LINK_LINE_RE.sub(lambda m: f"{m.group(1)}{new_url}", updated)
    else:
        sep = "" if updated.endswith("\n") else "\n"
        section = "\n## Commerce\n\n" if "## Commerce" not in updated else "\n"
        updated = f"{updated}{sep}{section}- Stripe Payment Link: {new_url}\n"

    if PAYMENT_STATUS_LINE_RE.search(updated):
        updated = PAYMENT_STATUS_LINE_RE.sub(lambda m: f"{m.group(1)}stripe_ready", updated)
    else:
        sep = "" if updated.endswith("\n") else "\n"
        updated = f"{updated}{sep}- payment_status: stripe_ready\n"

    if updated == original:
        return False

    tmp_path = original_path.with_suffix(".md.tmp")
    tmp_path.write_text(updated, encoding="utf-8", newline="\n")
    tmp_path.replace(original_path)
    return True


def expected_payment_link_metadata(item_id: str) -> dict[str, str]:
    return {
        "dake_item_id": item_id,
        "purchase_delivery_method": DELIVERY_METHOD_R2,
    }


def payment_link_already_synced(link: dict[str, Any], item_id: str) -> bool:
    after = link.get("after_completion") or {}
    redirect = after.get("redirect") or {}
    if after.get("type") != "redirect" or redirect.get("url") != R2_REDIRECT_URL:
        return False
    meta = link.get("metadata") or {}
    expected = expected_payment_link_metadata(item_id)
    return all(str(meta.get(k, "")) == v for k, v in expected.items())


def extract_live_price_info(line_items: list[dict[str, Any]]) -> tuple[int | None, str, str]:
    """Return (unit_amount, currency, product_id) from Payment Link line items."""
    if len(line_items) != 1:
        return None, "", ""
    price = line_items[0].get("price")
    if not isinstance(price, dict):
        return None, "", ""
    amount = price.get("unit_amount")
    currency = str(price.get("currency") or "").lower()
    product = price.get("product")
    product_id = product.get("id") if isinstance(product, dict) else str(product or "")
    return (int(amount) if amount is not None else None), currency, product_id


def plan_offline(
    catalog_items: list[dict[str, Any]],
    discontinued_items: tuple[dict[str, str], ...],
    *,
    archive_discontinued: bool,
) -> dict[str, Any]:
    """Build deterministic dry-run plan purely from local files (no Stripe API call)."""
    product_plans: list[dict[str, Any]] = []
    counts = {
        "create_new_link": 0,
        "update_existing_link": 0,
        "noop": 0,
        "archive_replaced_old_link": 0,
        "archive_discontinued_link": 0,
        "errors": 0,
    }

    for item in catalog_items:
        item_id = str(item["id"])
        target_price = int(item["price"])
        current_url = item.get("stripe_payment_link")
        title = str(item.get("title") or item_id)

        if not current_url:
            counts["create_new_link"] += 1
            product_plans.append({
                "id": item_id,
                "title": title,
                "type": item.get("type"),
                "target_price": target_price,
                "previous_price": None,
                "current_url": None,
                "action": "create_new_link",
                "reason": "no_existing_payment_link",
                "r2_redirect_url": R2_REDIRECT_URL,
                "metadata": expected_payment_link_metadata(item_id),
            })
            continue

        legacy_price = LEGACY_STRIPE_PRICES.get(current_url)
        if legacy_price is None:
            counts["update_existing_link"] += 1
            product_plans.append({
                "id": item_id,
                "title": title,
                "type": item.get("type"),
                "target_price": target_price,
                "previous_price": "unknown_offline_verify_on_apply",
                "current_url": current_url,
                "action": "update_existing_link_or_replace_if_price_differs",
                "reason": "unknown_legacy_price_check_live",
                "r2_redirect_url": R2_REDIRECT_URL,
                "metadata": expected_payment_link_metadata(item_id),
            })
        elif legacy_price != target_price:
            counts["create_new_link"] += 1
            counts["archive_replaced_old_link"] += 1
            product_plans.append({
                "id": item_id,
                "title": title,
                "type": item.get("type"),
                "target_price": target_price,
                "previous_price": legacy_price,
                "current_url": current_url,
                "action": "create_new_link_and_archive_old",
                "reason": f"price_changed_{legacy_price}_to_{target_price}",
                "r2_redirect_url": R2_REDIRECT_URL,
                "metadata": expected_payment_link_metadata(item_id),
            })
        else:
            counts["update_existing_link"] += 1
            product_plans.append({
                "id": item_id,
                "title": title,
                "type": item.get("type"),
                "target_price": target_price,
                "previous_price": legacy_price,
                "current_url": current_url,
                "action": "update_existing_link",
                "reason": f"price_unchanged_{target_price}_apply_r2_redirect_and_metadata",
                "r2_redirect_url": R2_REDIRECT_URL,
                "metadata": expected_payment_link_metadata(item_id),
            })

    discontinued_plans: list[dict[str, Any]] = []
    if archive_discontinued:
        for entry in discontinued_items:
            counts["archive_discontinued_link"] += 1
            discontinued_plans.append({
                "id": entry["id"],
                "title": entry["title"],
                "current_url": entry["stripe_payment_link"],
                "action": "archive_discontinued_link",
                "reason": entry["reason"],
            })

    return {
        "mode": "dry-run (offline)",
        "generated_at": now_jst(),
        "r2_redirect_url": R2_REDIRECT_URL,
        "summary": {
            "active_products_total": len(catalog_items),
            "new_links_to_create": counts["create_new_link"],
            "existing_links_to_update": counts["update_existing_link"],
            "noop": counts["noop"],
            "replaced_old_links_to_archive": counts["archive_replaced_old_link"],
            "discontinued_links_to_archive": counts["archive_discontinued_link"],
            "total_links_to_archive": counts["archive_replaced_old_link"] + counts["archive_discontinued_link"],
            "errors": counts["errors"],
        },
        "products": product_plans,
        "discontinued": discontinued_plans,
    }


def reconcile_live(
    client: StripeClient,
    catalog_items: list[dict[str, Any]],
    discontinued_items: tuple[dict[str, str], ...],
    *,
    apply: bool,
    archive_discontinued: bool,
    repo_root: Path = ROOT,
) -> dict[str, Any]:
    """Reconcile against live Stripe API (read-only if apply=False, mutating if apply=True)."""
    all_active_links = client.list_payment_links(active=True)
    links_by_url: dict[str, dict[str, Any]] = {
        str(link.get("url") or ""): link for link in all_active_links if link.get("url")
    }
    links_by_item_id: dict[str, list[dict[str, Any]]] = {}
    for link in all_active_links:
        meta_id = str((link.get("metadata") or {}).get("dake_item_id") or "")
        if meta_id:
            links_by_item_id.setdefault(meta_id, []).append(link)

    all_products = client.list_products()
    products_by_item_id: dict[str, dict[str, Any]] = {}
    for prod in all_products:
        meta_id = str((prod.get("metadata") or {}).get("dake_item_id") or "")
        if meta_id and prod.get("active", True):
            products_by_item_id[meta_id] = prod

    product_results: list[dict[str, Any]] = []
    discontinued_results: list[dict[str, Any]] = []
    counts = {
        "new_links_created": 0,
        "existing_links_updated": 0,
        "noop": 0,
        "replaced_old_links_archived": 0,
        "discontinued_links_archived": 0,
        "originals_updated": 0,
        "errors": 0,
    }

    for item in catalog_items:
        item_id = str(item["id"])
        title = str(item.get("title") or item_id)
        description = str(item.get("catch") or item.get("description") or title)[:450]
        target_price = int(item["price"])
        current_url = str(item.get("stripe_payment_link") or "").strip()
        source_original = str(item.get("source_original") or "")
        expected_meta = expected_payment_link_metadata(item_id)

        try:
            existing_link = links_by_url.get(current_url) if current_url else None
            if existing_link is None and item_id in links_by_item_id:
                # Fallback: pick active link matching metadata.dake_item_id if URL not yet written back
                existing_link = links_by_item_id[item_id][0]

            live_amount: int | None = None
            live_currency = ""
            live_product_id = ""
            if existing_link is not None:
                line_items = client.list_payment_link_line_items(existing_link["id"])
                live_amount, live_currency, live_product_id = extract_live_price_info(line_items)

            # Case 1: Existing active link has matching price -> update in-place or noop
            if existing_link is not None and live_amount == target_price and live_currency == "jpy":
                if payment_link_already_synced(existing_link, item_id):
                    counts["noop"] += 1
                    result_entry: dict[str, Any] = {
                        "id": item_id,
                        "title": title,
                        "target_price": target_price,
                        "live_price": live_amount,
                        "payment_link_id": existing_link["id"],
                        "payment_link_url": existing_link.get("url"),
                        "action": "noop",
                        "reason": "already_synced_with_r2_and_price",
                    }
                    # Ensure ORIGINAL.md points to this URL if it was missing
                    if apply and current_url != existing_link.get("url") and source_original:
                        orig_path = resolve_original_path(source_original, repo_root=repo_root)
                        if write_back_original_payment_link(orig_path, str(existing_link["url"])):
                            counts["originals_updated"] += 1
                            result_entry["original_updated"] = True
                    product_results.append(result_entry)
                    continue

                update_payload = {
                    "after_completion": {
                        "type": "redirect",
                        "redirect": {"url": R2_REDIRECT_URL},
                    },
                    "metadata": expected_meta,
                    "payment_intent_data": {
                        "metadata": expected_meta,
                    },
                }
                if apply:
                    client.update_payment_link(existing_link["id"], update_payload)
                counts["existing_links_updated"] += 1
                product_results.append({
                    "id": item_id,
                    "title": title,
                    "target_price": target_price,
                    "live_price": live_amount,
                    "payment_link_id": existing_link["id"],
                    "payment_link_url": existing_link.get("url"),
                    "action": "updated_existing_link" if apply else "would_update_existing_link",
                })
                continue

            # Case 2: Price changed OR no active link -> create new Price + new Payment Link, then archive old
            stripe_product_id = live_product_id or (products_by_item_id[item_id]["id"] if item_id in products_by_item_id else "")
            if not apply:
                counts["new_links_created"] += 1
                if existing_link is not None:
                    counts["replaced_old_links_archived"] += 1
                product_results.append({
                    "id": item_id,
                    "title": title,
                    "target_price": target_price,
                    "live_price": live_amount,
                    "stripe_product_id": stripe_product_id or "(would_create_product)",
                    "old_payment_link_id": existing_link["id"] if existing_link else None,
                    "action": "would_create_new_link_and_archive_old" if existing_link else "would_create_new_link",
                })
                continue

            if not stripe_product_id:
                prod_payload = {
                    "name": title,
                    "description": description,
                    "active": True,
                    "tax_code": DEFAULT_TAX_CODE,
                    "metadata": {"dake_item_id": item_id},
                }
                created_prod = client.create_product(
                    prod_payload,
                    make_idempotency_key("prod", item_id, prod_payload),
                )
                stripe_product_id = str(created_prod["id"])

            # Reuse matching active JPY one-time price on this product if one already exists
            existing_prices = client.list_prices(stripe_product_id)
            matched_price = next(
                (
                    p
                    for p in existing_prices
                    if p.get("active") is True
                    and p.get("type", "one_time") == "one_time"
                    and not p.get("recurring")
                    and str(p.get("currency", "")).lower() == "jpy"
                    and p.get("unit_amount") == target_price
                ),
                None,
            )
            if matched_price is not None:
                price_id = str(matched_price["id"])
            else:
                price_payload = {
                    "product": stripe_product_id,
                    "currency": "jpy",
                    "unit_amount": target_price,
                    "metadata": {"dake_item_id": item_id},
                }
                created_price = client.create_price(
                    price_payload,
                    make_idempotency_key("price", item_id, price_payload),
                )
                price_id = str(created_price["id"])

            link_payload = {
                "line_items": [{"price": price_id, "quantity": 1}],
                "after_completion": {
                    "type": "redirect",
                    "redirect": {"url": R2_REDIRECT_URL},
                },
                "metadata": expected_meta,
                "payment_intent_data": {
                    "metadata": expected_meta,
                },
            }
            created_link = client.create_payment_link(
                link_payload,
                make_idempotency_key("plink", item_id, link_payload),
            )
            new_link_id = str(created_link["id"])
            new_link_url = str(created_link["url"])

            # Write back new URL to ORIGINAL.md BEFORE archiving old link
            original_updated = False
            if source_original:
                orig_path = resolve_original_path(source_original, repo_root=repo_root)
                original_updated = write_back_original_payment_link(orig_path, new_link_url)
                if original_updated:
                    counts["originals_updated"] += 1

            archived_old_id = None
            if existing_link is not None and existing_link["id"] != new_link_id:
                client.update_payment_link(existing_link["id"], {"active": False})
                archived_old_id = existing_link["id"]
                counts["replaced_old_links_archived"] += 1

            counts["new_links_created"] += 1
            product_results.append({
                "id": item_id,
                "title": title,
                "target_price": target_price,
                "previous_price": live_amount,
                "stripe_product_id": stripe_product_id,
                "price_id": price_id,
                "payment_link_id": new_link_id,
                "payment_link_url": new_link_url,
                "archived_old_payment_link_id": archived_old_id,
                "original_updated": original_updated,
                "action": "created_new_link",
            })
        except Exception as exc:  # noqa: BLE001
            counts["errors"] += 1
            product_results.append({
                "id": item_id,
                "title": title,
                "target_price": target_price,
                "action": "error",
                "error": sanitize_text(str(exc)),
            })

    # Archive C-list discontinued Payment Links
    if archive_discontinued:
        for entry in discontinued_items:
            item_id = entry["id"]
            url = entry["stripe_payment_link"]
            try:
                active_link = links_by_url.get(url)
                if active_link is None:
                    discontinued_results.append({
                        "id": item_id,
                        "title": entry["title"],
                        "current_url": url,
                        "action": "already_inactive",
                    })
                    continue
                if apply:
                    client.update_payment_link(active_link["id"], {"active": False})
                counts["discontinued_links_archived"] += 1
                discontinued_results.append({
                    "id": item_id,
                    "title": entry["title"],
                    "payment_link_id": active_link["id"],
                    "current_url": url,
                    "action": "archived_discontinued_link" if apply else "would_archive_discontinued_link",
                })
            except Exception as exc:  # noqa: BLE001
                counts["errors"] += 1
                discontinued_results.append({
                    "id": item_id,
                    "title": entry["title"],
                    "current_url": url,
                    "action": "error",
                    "error": sanitize_text(str(exc)),
                })

    return {
        "mode": "apply" if apply else "dry-run (online)",
        "generated_at": now_jst(),
        "r2_redirect_url": R2_REDIRECT_URL,
        "summary": {
            "active_products_total": len(catalog_items),
            "new_links_created": counts["new_links_created"],
            "existing_links_updated": counts["existing_links_updated"],
            "noop": counts["noop"],
            "replaced_old_links_archived": counts["replaced_old_links_archived"],
            "discontinued_links_archived": counts["discontinued_links_archived"],
            "total_links_archived": counts["replaced_old_links_archived"] + counts["discontinued_links_archived"],
            "originals_updated": counts["originals_updated"],
            "errors": counts["errors"],
        },
        "products": product_results,
        "discontinued": discontinued_results,
    }


def run_post_apply_sync(store_site: Path, dakeapp_site: Path) -> None:
    sync_script = ROOT / "tools" / "store" / "sync_store_to_site.py"
    subprocess.run(
        [sys.executable, str(sync_script), "--store-site", str(store_site)],
        cwd=str(ROOT),
        check=True,
    )
    store_json = store_site / "public" / "assets" / "data" / "store_products.generated.json"
    dakeapp_json = dakeapp_site / "public" / "assets" / "data" / "store_products.generated.json"
    if store_json.exists() and dakeapp_json.parent.exists():
        shutil.copy2(store_json, dakeapp_json)


def save_report(report: dict[str, Any], report_path: Path) -> None:
    safe_report = sanitize_obj(report)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(
        json.dumps(safe_report, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def print_summary(report: dict[str, Any], report_path: Path) -> None:
    s = report["summary"]
    print(f"=== DAKE Store Stripe R2 Sync [{report['mode']}] ===")
    print(f"対象商品数 (第1弾): {s['active_products_total']} 件")
    if "new_links_to_create" in s:
        print(f"新規作成予定件数   : {s['new_links_to_create']} 件")
        print(f"既存更新予定件数   : {s['existing_links_to_update']} 件")
        print(f"変更不要(noop)件数 : {s['noop']} 件")
        print(
            f"アーカイブ予定件数 : {s['total_links_to_archive']} 件 "
            f"(価格変更に伴う旧リンク置換: {s['replaced_old_links_to_archive']}件 / C判定停止: {s['discontinued_links_to_archive']}件)"
        )
        print(f"エラー件数         : {s['errors']} 件")
    else:
        print(f"新規作成件数       : {s['new_links_created']} 件")
        print(f"既存更新件数       : {s['existing_links_updated']} 件")
        print(f"変更不要(noop)件数 : {s['noop']} 件")
        print(
            f"アーカイブ件数     : {s['total_links_archived']} 件 "
            f"(旧リンク置換: {s['replaced_old_links_archived']}件 / C判定停止: {s['discontinued_links_archived']}件)"
        )
        print(f"ORIGINAL書戻し件数 : {s['originals_updated']} 件")
        print(f"エラー件数         : {s['errors']} 件")
    print(f"レポート出力先     : {report_path}")


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Sync DAKE Store 47 Wave 1 products with Stripe Payment Links (R2 redirect & metadata) and archive C-list links."
    )
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument(
        "--dry-run",
        action="store_true",
        default=True,
        help="Plan changes without writing to Stripe (default).",
    )
    mode.add_argument(
        "--apply",
        action="store_true",
        help="Execute live writes against Stripe API using STRIPE_SECRET_KEY.",
    )
    parser.add_argument(
        "--online",
        action="store_true",
        help="During --dry-run, read current state from live Stripe API (GET only, no writes).",
    )
    parser.add_argument(
        "--archive-discontinued",
        action="store_true",
        default=True,
        help="Deactivate C-list discontinued Payment Links (enabled by default).",
    )
    parser.add_argument(
        "--skip-archive-discontinued",
        action="store_true",
        help="Skip deactivating C-list discontinued Payment Links.",
    )
    parser.add_argument(
        "--only",
        help="Comma-separated list of product IDs to process (e.g. dake_pdf_extract,DAKE_Image_BatchPDF).",
    )
    parser.add_argument(
        "--store-site",
        type=Path,
        default=DEFAULT_STORE_SITE,
        help="Path to dake-store-site repository.",
    )
    parser.add_argument(
        "--dakeapp-site",
        type=Path,
        default=DEFAULT_DAKEAPP_SITE,
        help="Path to dakeapp-site repository.",
    )
    parser.add_argument(
        "--report-path",
        type=Path,
        default=DEFAULT_REPORT_PATH,
        help="Output path for JSON sync report (gitignored under tmp/).",
    )
    parser.add_argument(
        "--skip-sync",
        action="store_true",
        help="Skip running sync_store_to_site.py after --apply.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    store_site = args.store_site.resolve()
    dakeapp_site = args.dakeapp_site.resolve()
    catalog_path = store_site / "public" / "assets" / "data" / "store_products.generated.json"
    private_download_path = store_site / "functions" / "_lib" / "private-download.mjs"

    if not catalog_path.exists():
        print(f"[ERROR] Catalog JSON not found: {catalog_path}", file=sys.stderr)
        return 1
    if not private_download_path.exists():
        print(f"[ERROR] private-download.mjs not found: {private_download_path}", file=sys.stderr)
        return 1

    catalog = json.loads(catalog_path.read_text(encoding="utf-8"))
    catalog_items: list[dict[str, Any]] = list(catalog.get("items") or [])
    r2_ids = parse_private_download_ids(private_download_path)

    catalog_ids = {str(it["id"]) for it in catalog_items}
    missing_in_r2 = sorted(catalog_ids - r2_ids)
    if missing_in_r2:
        print(
            f"[ERROR] Catalog IDs missing from private-download.mjs DESKTOP_DELIVERIES: {missing_in_r2}",
            file=sys.stderr,
        )
        return 1

    if args.only:
        selected = {part.strip() for part in args.only.split(",") if part.strip()}
        catalog_items = [it for it in catalog_items if str(it["id"]) in selected]
        discontinued_items = tuple(d for d in DISCONTINUED_STRIPE_ITEMS if d["id"] in selected)
    else:
        discontinued_items = DISCONTINUED_STRIPE_ITEMS

    archive_discontinued = args.archive_discontinued and not args.skip_archive_discontinued

    if args.apply or args.online:
        secret_key = os.environ.get("STRIPE_SECRET_KEY", "")
        if not secret_key:
            print(
                "[ERROR] STRIPE_SECRET_KEY environment variable is required for --apply / --online.",
                file=sys.stderr,
            )
            return 1
        try:
            client = StripeClient(secret_key)
        except StripeSyncError as exc:
            print(f"[ERROR] {sanitize_text(str(exc))}", file=sys.stderr)
            return 1
        report = reconcile_live(
            client,
            catalog_items,
            discontinued_items,
            apply=args.apply,
            archive_discontinued=archive_discontinued,
        )
        if args.apply and not args.skip_sync and report["summary"]["errors"] == 0:
            run_post_apply_sync(store_site, dakeapp_site)
    else:
        report = plan_offline(
            catalog_items,
            discontinued_items,
            archive_discontinued=archive_discontinued,
        )

    save_report(report, args.report_path)
    print_summary(report, args.report_path)
    return 0 if report["summary"]["errors"] == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
