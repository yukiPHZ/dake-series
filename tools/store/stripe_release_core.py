from __future__ import annotations

import copy
import csv
import hashlib
import json
import os
import re
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Callable


ROOT = Path(__file__).resolve().parents[2]
STORE_PRODUCTS_JSON = ROOT / "tools" / "generated" / "store_products.generated.json"
ROLLOUT_REVIEW_CSV = ROOT / "tools" / "reports" / "stripe_payment_link_rollout_review.csv"
PACK_READY_CSV = ROOT / "tools" / "reports" / "stripe_pack2_manual_delivery_ready.csv"
ARTIFACT_ROOT = ROOT / "tools" / "reports" / "release_artifacts"

DRY_RUN_NOTICE = "DRY RUN ONLY. NO STRIPE API CALL. NO LIVE OBJECT IS CREATED."
PRODUCT_PLACEHOLDER = "__PRODUCT_ID_FROM_LIVE_PRODUCT__"
PRICE_PLACEHOLDER = "__PRICE_ID_FROM_LIVE_PRICE__"
STORE_PRODUCT_BASE_URL = "https://store.dakeapp.com/product/"
PACK_TAX_CODE_CANDIDATE = "txcd_10202003"
JST = timezone(timedelta(hours=9))
SECRET_PATTERN = re.compile(r"sk_(?:test|live)_[A-Za-z0-9]{16,}|whsec_[A-Za-z0-9]{16,}")


class SafetyStop(RuntimeError):
    pass


def now_jst() -> str:
    return datetime.now(JST).isoformat(timespec="seconds")


def read_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def write_json_atomic(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(f"{path.name}.{os.getpid()}.tmp")
    with temp.open("w", encoding="utf-8", newline="\n") as handle:
        handle.write(json.dumps(data, ensure_ascii=False, indent=2) + "\n")
        handle.flush()
        os.fsync(handle.fileno())
    os.replace(temp, path)


def write_text_atomic(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(f"{path.name}.{os.getpid()}.tmp")
    with temp.open("w", encoding="utf-8", newline="\n") as handle:
        handle.write(text)
        handle.flush()
        os.fsync(handle.fileno())
    os.replace(temp, path)


def read_csv(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def row_map(path: Path) -> dict[str, dict[str, str]]:
    return {row.get("id", ""): row for row in read_csv(path) if row.get("id")}


def canonical_json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def sha256_payload(value: Any) -> str:
    return hashlib.sha256(canonical_json(value).encode("utf-8")).hexdigest()


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def contains_secret_like_value(value: Any) -> bool:
    return bool(SECRET_PATTERN.search(canonical_json(value)))


def relative_to_root(path: Path) -> str:
    try:
        return path.resolve().relative_to(ROOT).as_posix()
    except ValueError:
        return str(path)


def safe_id(item_id: str) -> str:
    safe = re.sub(r"[^A-Za-z0-9_-]+", "_", item_id).strip("_")
    return safe or "item"


def idempotency_key(kind: str, item_id: str, payload_hash: str) -> str:
    return f"dake-release-{kind}-v1-{safe_id(item_id)}-{payload_hash[:12]}"


def store_url(item_id: str) -> str:
    return f"{STORE_PRODUCT_BASE_URL}?id={item_id}"


def normalize_space(value: object) -> str:
    return " ".join(str(value or "").replace("\r\n", "\n").replace("\r", "\n").split()).strip()


def normalize_original_value(value: str) -> str:
    text = value.strip()
    if len(text) >= 2 and text.startswith("`") and text.endswith("`"):
        text = text[1:-1].strip()
    return text.strip()


def metadata_line(text: str, key: str) -> str:
    pattern = re.compile(rf"^\s*[-*]\s*{re.escape(key)}\s*:\s*(.+?)\s*$", re.MULTILINE)
    match = pattern.search(text)
    return normalize_original_value(match.group(1)) if match else ""


def json_like_value(text: str, key: str) -> str:
    pattern = re.compile(rf'"{re.escape(key)}"\s*:\s*"([^"]*)"')
    match = pattern.search(text)
    return normalize_original_value(match.group(1)) if match else ""


def first_url_after(text: str, label: str) -> str:
    pattern = re.compile(rf"{re.escape(label)}[^\n]*?(https?://[^\s`]+)")
    match = pattern.search(text)
    return match.group(1).strip() if match else ""


def parse_price_from_text(text: str) -> int | None:
    for label in ["price", "萓｡譬ｼ"]:
        pattern = re.compile(rf"{re.escape(label)}\s*[:：]\s*([0-9][0-9,]*)")
        match = pattern.search(text)
        if match:
            return int_price(match.group(1))
    return None


def clean_description(value: object, fallback: str, limit: int = 240) -> str:
    text = normalize_space(value) or normalize_space(fallback)
    if len(text) > limit:
        return text[: limit - 3].rstrip() + "..."
    return text


def discover_originals() -> list[Path]:
    return list((ROOT / "01_apps").glob("**/ORIGINAL.md")) + list((ROOT / "04_packs").glob("**/ORIGINAL.md"))


def validate_product_id(product_id: str) -> None:
    if not product_id or not product_id.strip():
        raise SafetyStop("product_id is required")
    if any(part in product_id for part in ["..", "/", "\\"]):
        raise SafetyStop("product_id must not contain path separators or traversal")
    if any(ord(char) < 32 for char in product_id):
        raise SafetyStop("product_id must not contain control characters")


def parse_original_item(path: Path) -> dict[str, Any]:
    text = path.read_text(encoding="utf-8")
    relative = relative_to_root(path)
    if relative.startswith("04_packs/"):
        product_type = "pack"
        product_id = metadata_line(text, "pack_id") or json_like_value(text, "folder_name") or path.parent.name
        title = (
            metadata_line(text, "pack_title")
            or metadata_line(text, "title")
            or json_like_value(text, "display_name")
            or product_id
        )
        payment_status = metadata_line(text, "payment_status") or "booth_only"
        stripe_payment_link = metadata_line(text, "stripe_payment_link")
        if stripe_payment_link.lower() in {"not set", "none", "null", "譛ｪ險ｭ螳・"}:
            stripe_payment_link = ""
        return {
            "id": product_id,
            "type": product_type,
            "source_repo": "DAKE_series",
            "source_original": relative,
            "title": title,
            "price": parse_price_from_text(text),
            "currency": "JPY",
            "booth_url": metadata_line(text, "booth_url") or json_like_value(text, "booth_url") or first_url_after(text, "BOOTH URL"),
            "github_release_url": "",
            "stripe_payment_link": stripe_payment_link or None,
            "payment_status": payment_status,
            "description": "",
        }

    product_id = path.parent.name
    return {
        "id": product_id,
        "type": "app",
        "source_repo": "DAKE_series",
        "source_original": relative,
        "title": metadata_line(text, "title") or path.parent.name,
        "price": parse_price_from_text(text),
        "currency": "JPY",
        "booth_url": metadata_line(text, "booth_url") or first_url_after(text, "BOOTH URL"),
        "github_release_url": metadata_line(text, "release_url") or first_url_after(text, "Release"),
        "stripe_payment_link": metadata_line(text, "stripe_payment_link") or None,
        "payment_status": metadata_line(text, "payment_status") or "",
        "description": "",
    }


def discover_original_items() -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    duplicates: dict[str, list[str]] = {}
    for path in discover_originals():
        item = parse_original_item(path)
        item_id = str(item.get("id") or "")
        if not item_id:
            continue
        if item_id in result:
            duplicates.setdefault(item_id, [result[item_id]["source_original"]]).append(item["source_original"])
        result[item_id] = item
    if duplicates:
        detail = "; ".join(f"{item_id}: {', '.join(paths)}" for item_id, paths in sorted(duplicates.items()))
        raise SafetyStop("duplicate product id in ORIGINAL.md files: " + detail)
    return result


def load_store_items() -> dict[str, dict[str, Any]]:
    if not STORE_PRODUCTS_JSON.exists():
        return {}
    try:
        data = read_json(STORE_PRODUCTS_JSON)
    except json.JSONDecodeError:
        return {}
    items = data.get("items")
    if not isinstance(items, list):
        raise SafetyStop("tools/generated/store_products.generated.json must contain items")
    result: dict[str, dict[str, Any]] = {}
    duplicates: set[str] = set()
    for item in items:
        if not isinstance(item, dict) or not item.get("id"):
            continue
        item_id = str(item["id"])
        if item_id in result:
            duplicates.add(item_id)
        result[item_id] = item
    if duplicates:
        raise SafetyStop("duplicate product id in generated store JSON: " + ", ".join(sorted(duplicates)))
    return result


def resolve_product(product_id: str) -> dict[str, Any]:
    validate_product_id(product_id)
    original_items = discover_original_items()
    store_items = load_store_items()
    generated_item = store_items.get(product_id)
    direct_item = original_items.get(product_id)
    if direct_item is None and generated_item is not None:
        generated_source = str(generated_item.get("source_original") or "")
        for item in original_items.values():
            if item.get("source_original") == generated_source:
                direct_item = {**item, "id": product_id}
                break
    if direct_item is None:
        raise SafetyStop(f"Product id not found: {product_id}")
    item = {**direct_item}
    if generated_item:
        for key, value in generated_item.items():
            if item.get(key) in (None, "", []) and value not in (None, "", []):
                item[key] = value
        item["id"] = product_id
        item["source_original"] = direct_item["source_original"]
    source_original = str(item.get("source_original") or "")
    if not source_original:
        raise SafetyStop(f"{product_id}: source_original is missing")
    source_path = ROOT / source_original
    if not source_path.exists():
        raise SafetyStop(f"{product_id}: source ORIGINAL.md does not exist: {source_original}")
    discovered = {path.resolve() for path in discover_originals()}
    if source_path.resolve() not in discovered:
        raise SafetyStop(f"{product_id}: source_original is outside supported ORIGINAL.md search targets")
    product_type = str(item.get("type") or "")
    if product_type not in {"app", "pack"}:
        raise SafetyStop(f"{product_id}: unsupported product type: {product_type}")
    return {
        "item": item,
        "product_id": product_id,
        "product_type": product_type,
        "source_original": source_original,
        "source_path": source_path,
        "original_text": source_path.read_text(encoding="utf-8"),
    }


def validate_payment_state(product_id: str, item: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    status = str(item.get("payment_status") or "")
    payment_link = str(item.get("stripe_payment_link") or "")
    if status == "stripe_ready" or payment_link:
        errors.append(f"{product_id}: already stripe_ready or has a Stripe Payment Link")
    if status not in {"booth_only", "preparing", ""}:
        errors.append(f"{product_id}: unsupported payment_status before release: {status}")
    if status == "preparing":
        errors.append(f"{product_id}: payment_status is preparing")
    return errors


def int_price(value: Any) -> int | None:
    if isinstance(value, int) and value > 0:
        return value
    text = str(value or "").replace(",", "").strip()
    if not text:
        return None
    try:
        price = int(text)
    except ValueError:
        return None
    return price if price > 0 else None


def currency_jpy(value: Any) -> str:
    return "jpy" if str(value or "").strip().lower() == "jpy" else str(value or "").strip().lower()


def base_payload(
    *,
    product_id: str,
    product_type: str,
    title: str,
    description: str,
    price: int,
    currency: str,
    source_original: str,
    booth_url: str,
    github_release_url: str,
    tax_code_candidate: str,
    purchase_delivery_method: str,
) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any]]:
    metadata = {
        "dake_item_id": product_id,
        "dake_type": product_type,
        "source_repo": "DAKE_series",
        "source_original": source_original,
        "store_url": store_url(product_id),
        "booth_url": booth_url,
    }
    if github_release_url:
        metadata["github_release_url"] = github_release_url
    if purchase_delivery_method:
        metadata["purchase_delivery_method"] = purchase_delivery_method
        metadata["delivery_policy"] = purchase_delivery_method

    payment_metadata = {
        "dake_item_id": product_id,
        "dake_type": product_type,
        "source_original": source_original,
    }
    if purchase_delivery_method:
        payment_metadata["purchase_delivery_method"] = purchase_delivery_method
        payment_metadata["delivery_policy"] = purchase_delivery_method

    product_payload = {
        "name": title,
        "description": description,
        "active": True,
        "tax_code": tax_code_candidate,
        "metadata": {key: str(value) for key, value in metadata.items() if str(value) != ""},
    }
    price_payload = {
        "currency": currency,
        "unit_amount": price,
        "product": PRODUCT_PLACEHOLDER,
        "metadata": {"dake_item_id": product_id},
    }
    payment_link_payload = {
        "line_items": [
            {
                "price": PRICE_PLACEHOLDER,
                "quantity": 1,
            }
        ],
        "metadata": payment_metadata,
        "payment_intent_data": {
            "metadata": payment_metadata,
        },
    }
    return product_payload, price_payload, payment_link_payload


def finalize_payload_hashes(product_id: str, data: dict[str, Any]) -> None:
    product_hash = sha256_payload(data["product_payload"])
    price_hash = sha256_payload(data["price_payload"])
    link_hash = sha256_payload(data["payment_link_payload"])
    data["product_payload_sha256"] = product_hash
    data["price_payload_sha256"] = price_hash
    data["payment_link_payload_sha256"] = link_hash
    data["product_idempotency_key"] = idempotency_key("product", product_id, product_hash)
    data["price_idempotency_key"] = idempotency_key("price", product_id, price_hash)
    data["payment_link_idempotency_key"] = idempotency_key("link", product_id, link_hash)


def build_pack_release(context: dict[str, Any]) -> dict[str, Any]:
    product_id = context["product_id"]
    item = context["item"]
    original_text = context["original_text"]
    source_original = context["source_original"]
    source_path = context["source_path"]
    rollout = row_map(ROLLOUT_REVIEW_CSV).get(product_id, {})
    pack_ready = row_map(PACK_READY_CSV).get(product_id, {})
    manifest_path = source_path.parent / "pack_manifest.json"
    errors = validate_payment_state(product_id, item)
    if not manifest_path.exists():
        errors.append(f"{product_id}: pack_manifest.json is missing")
        manifest: dict[str, Any] = {}
    else:
        manifest = read_json(manifest_path)

    title = pack_ready.get("title") or item.get("title") or manifest.get("display_name") or product_id
    price = int_price(item.get("price") or manifest.get("price"))
    currency = currency_jpy(item.get("currency") or "JPY")
    booth_url = str(item.get("booth_url") or manifest.get("booth_url") or "")
    tax_code = rollout.get("tax_code_candidate") or PACK_TAX_CODE_CANDIDATE
    pack_zip = str(manifest.get("pack_zip") or pack_ready.get("distribution_path") or "")
    zip_path = ROOT / pack_zip if pack_zip else None
    zip_exists = bool(zip_path and zip_path.exists())
    zip_size = zip_path.stat().st_size if zip_exists and zip_path else None
    zip_sha256 = file_sha256(zip_path) if zip_exists and zip_path else ""
    expected_size = manifest.get("pack_zip_size")
    expected_sha256 = str(manifest.get("pack_zip_sha256") or "").lower()

    manual_checks = {
        "manual_delivery_method": "manual_email_private_download" in original_text,
        "purchase_delivery_ready": "purchase_delivery_ready: yes" in original_text,
        "manual_dashboard_ready": "stripe_creation_method: manual_dashboard_ready" in original_text,
        "review_ready": "review_result: ready" in original_text,
        "delivery_window": "next business day" in original_text or "次営業日以内" in original_text,
        "buyer_notice": "Buyer notice" in original_text,
        "resend_policy": "Resend and failure handling" in original_text,
        "rule_reference": "00_core/DAKE_PACK_MANUAL_DELIVERY_RULE.md" in original_text,
    }
    for key, ok in manual_checks.items():
        if not ok:
            errors.append(f"{product_id}: Pack delivery gate failed: {key}")
    if price is None:
        errors.append(f"{product_id}: price is missing or invalid")
        price = 0
    if currency != "jpy":
        errors.append(f"{product_id}: currency must be JPY")
    if not booth_url:
        errors.append(f"{product_id}: booth_url is missing")
    if not zip_exists:
        errors.append(f"{product_id}: Pack ZIP is missing")
    if expected_size is not None and zip_size != expected_size:
        errors.append(f"{product_id}: Pack ZIP size mismatch")
    if expected_sha256 and zip_sha256.lower() != expected_sha256:
        errors.append(f"{product_id}: Pack ZIP SHA256 mismatch")
    if tax_code != PACK_TAX_CODE_CANDIDATE:
        errors.append(f"{product_id}: unexpected tax code candidate: {tax_code}")

    description = clean_description(
        "",
        f"{title}。本商品は自動ダウンロードではありません。Stripe決済確認後、購入時のメールアドレス宛に次営業日以内にダウンロード方法をご案内します。",
    )
    product_payload, price_payload, payment_link_payload = base_payload(
        product_id=product_id,
        product_type="pack",
        title=str(title),
        description=description,
        price=price,
        currency=currency,
        source_original=source_original,
        booth_url=booth_url,
        github_release_url="",
        tax_code_candidate=tax_code,
        purchase_delivery_method="manual_email_private_download",
    )

    data = {
        "mode": "dry-run",
        "notice": DRY_RUN_NOTICE,
        "created_at": now_jst(),
        "product_id": product_id,
        "product_type": "pack",
        "source_original": source_original,
        "title": str(title),
        "price": price,
        "currency": currency,
        "tax_code_candidate": tax_code,
        "tax_candidate_review": "operator_confirmation_required",
        "payment_status_before": item.get("payment_status"),
        "stripe_payment_link_before": item.get("stripe_payment_link"),
        "purchase_delivery_ready": "yes" if not errors else "no",
        "purchase_delivery_method": "manual_email_private_download",
        "distribution_file": Path(pack_zip).name if pack_zip else "",
        "distribution_path": pack_zip,
        "distribution_file_size": zip_size,
        "distribution_file_sha256": zip_sha256.lower(),
        "product_payload": product_payload,
        "price_payload": price_payload,
        "payment_link_payload": payment_link_payload,
        "ready_for_live_execution": "yes" if not errors else "no",
        "errors": errors,
        "secret_read": "no",
        "live_api_called": "no",
        "safety": [
            "No Stripe API call is made in dry-run mode.",
            "No Stripe Secret Key is read in dry-run mode.",
            "No Product, Price, or Payment Link is created.",
            "No Payment Link URL is written back to ORIGINAL.md.",
            "No generated JSON or Store file is updated.",
            "No buyer information is stored.",
            "No private download URL is stored.",
            "No absolute local Pack ZIP path is exposed.",
        ],
    }
    finalize_payload_hashes(product_id, data)
    if contains_secret_like_value(data):
        data["errors"].append(f"{product_id}: secret-like value detected in dry-run payload")
        data["ready_for_live_execution"] = "no"
    return data


def build_app_release(context: dict[str, Any]) -> dict[str, Any]:
    product_id = context["product_id"]
    item = context["item"]
    source_original = context["source_original"]
    rollout = row_map(ROLLOUT_REVIEW_CSV).get(product_id, {})
    errors = validate_payment_state(product_id, item)
    title = rollout.get("stripe_product_name") or rollout.get("title") or item.get("title") or product_id
    price = int_price(item.get("price") or rollout.get("price"))
    currency = currency_jpy(item.get("currency") or rollout.get("currency") or "JPY")
    booth_url = str(item.get("booth_url") or rollout.get("booth_url") or "")
    github_release_url = str(item.get("github_release_url") or rollout.get("github_release_url") or "")
    tax_code = rollout.get("tax_code_candidate") or ""
    if not rollout:
        errors.append(f"{product_id}: missing rollout review row")
    if rollout and not (
        rollout.get("review_result") == "create"
        and rollout.get("creation_method") == "api_candidate"
        and rollout.get("price_check") == "price_ok"
        and rollout.get("metadata_ready") == "yes"
    ):
        errors.append(f"{product_id}: rollout review row is not an API creation candidate")
    if price is None:
        errors.append(f"{product_id}: price is missing or invalid")
        price = 0
    if currency != "jpy":
        errors.append(f"{product_id}: currency must be JPY")
    if not booth_url:
        errors.append(f"{product_id}: booth_url is missing")
    if not github_release_url:
        errors.append(f"{product_id}: github_release_url is missing")
    if not tax_code:
        errors.append(f"{product_id}: tax_code_candidate is missing")

    description = clean_description(item.get("description"), f"{title} / {item.get('category') or ''}")
    product_payload, price_payload, payment_link_payload = base_payload(
        product_id=product_id,
        product_type="app",
        title=str(title),
        description=description,
        price=price,
        currency=currency,
        source_original=source_original,
        booth_url=booth_url,
        github_release_url=github_release_url,
        tax_code_candidate=tax_code,
        purchase_delivery_method="",
    )
    data = {
        "mode": "dry-run",
        "notice": DRY_RUN_NOTICE,
        "created_at": now_jst(),
        "product_id": product_id,
        "product_type": "app",
        "source_original": source_original,
        "title": str(title),
        "price": price,
        "currency": currency,
        "tax_code_candidate": tax_code,
        "tax_candidate_review": "operator_confirmation_required",
        "payment_status_before": item.get("payment_status"),
        "stripe_payment_link_before": item.get("stripe_payment_link"),
        "purchase_delivery_ready": "not_applicable",
        "purchase_delivery_method": "",
        "distribution_file": "",
        "distribution_path": "",
        "distribution_file_size": None,
        "distribution_file_sha256": "",
        "product_payload": product_payload,
        "price_payload": price_payload,
        "payment_link_payload": payment_link_payload,
        "ready_for_live_execution": "yes" if not errors else "no",
        "errors": errors,
        "secret_read": "no",
        "live_api_called": "no",
        "safety": [
            "No Stripe API call is made in dry-run mode.",
            "No Stripe Secret Key is read in dry-run mode.",
            "No Product, Price, or Payment Link is created.",
            "No Payment Link URL is written back to ORIGINAL.md.",
            "No generated JSON or Store file is updated.",
            "No buyer information is stored.",
        ],
    }
    finalize_payload_hashes(product_id, data)
    if contains_secret_like_value(data):
        data["errors"].append(f"{product_id}: secret-like value detected in dry-run payload")
        data["ready_for_live_execution"] = "no"
    return data


def build_release_payload(product_id: str) -> dict[str, Any]:
    context = resolve_product(product_id)
    if context["product_type"] == "pack":
        return build_pack_release(context)
    return build_app_release(context)


def artifact_dir(product_id: str) -> Path:
    return ARTIFACT_ROOT / safe_id(product_id)


def dry_run_json_path(product_id: str) -> Path:
    return artifact_dir(product_id) / "stripe_release_dry_run.json"


def dry_run_md_path(product_id: str) -> Path:
    return artifact_dir(product_id) / "stripe_release_dry_run.md"


def state_json_path(product_id: str) -> Path:
    return artifact_dir(product_id) / "stripe_release_state.json"


def result_json_path(product_id: str) -> Path:
    return artifact_dir(product_id) / "stripe_release_result.json"


def result_md_path(product_id: str) -> Path:
    return artifact_dir(product_id) / "stripe_release_result.md"


def md(value: object) -> str:
    text = "" if value is None else str(value)
    text = text.replace("|", "\\|").replace("\r\n", "\n").replace("\r", "\n")
    return "<br>".join(part.strip() for part in text.split("\n") if part.strip()) or "-"


def table(headers: list[str], rows: list[list[object]]) -> str:
    lines = [
        "| " + " | ".join(headers) + " |",
        "| " + " | ".join("---" for _ in headers) + " |",
    ]
    if not rows:
        rows = [["-" for _ in headers]]
    for row in rows:
        lines.append("| " + " | ".join(md(cell) for cell in row) + " |")
    return "\n".join(lines)


def write_dry_run_files(payload: dict[str, Any]) -> tuple[Path, Path]:
    product_id = payload["product_id"]
    json_path = dry_run_json_path(product_id)
    md_path = dry_run_md_path(product_id)
    write_json_atomic(json_path, payload)
    write_text_atomic(md_path, dry_run_markdown(payload, json_path),)
    return json_path, md_path


def dry_run_markdown(payload: dict[str, Any], json_path: Path) -> str:
    product_rows = [
        ["product_id", payload["product_id"]],
        ["product_type", payload["product_type"]],
        ["title", payload["title"]],
        ["price", f"{payload['price']} {payload['currency']}"],
        ["tax_code_candidate", payload["tax_code_candidate"]],
    ]
    delivery_rows = [
        ["purchase_delivery_ready", payload["purchase_delivery_ready"]],
        ["purchase_delivery_method", payload["purchase_delivery_method"]],
        ["distribution_file", payload["distribution_file"]],
        ["distribution_file_sha256", payload["distribution_file_sha256"]],
    ]
    current_rows = [
        ["payment_status_before", payload["payment_status_before"]],
        ["stripe_payment_link_before", payload["stripe_payment_link_before"]],
    ]
    hash_rows = [
        ["product_payload_sha256", payload["product_payload_sha256"]],
        ["price_payload_sha256", payload["price_payload_sha256"]],
        ["payment_link_payload_sha256", payload["payment_link_payload_sha256"]],
    ]
    key_rows = [
        ["product_idempotency_key", payload["product_idempotency_key"]],
        ["price_idempotency_key", payload["price_idempotency_key"]],
        ["payment_link_idempotency_key", payload["payment_link_idempotency_key"]],
    ]
    error_rows = [[error] for error in payload["errors"]]
    return f"""# Product Stripe Release Dry Run

## Product

{table(['field', 'value'], product_rows)}

## Source of Truth

- `{payload['source_original']}`

## Delivery Readiness

{table(['field', 'value'], delivery_rows)}

## Current Payment State

{table(['field', 'value'], current_rows)}

## Stripe Product

{table(['field', 'value'], [['name', payload['product_payload']['name']], ['description', payload['product_payload']['description']], ['tax_code', payload['product_payload']['tax_code']]])}

## Stripe Price

{table(['field', 'value'], [['currency', payload['price_payload']['currency']], ['unit_amount', payload['price_payload']['unit_amount']], ['product', payload['price_payload']['product']]])}

## Stripe Payment Link

{table(['field', 'value'], [['price', payload['payment_link_payload']['line_items'][0]['price']], ['quantity', payload['payment_link_payload']['line_items'][0]['quantity']]])}

## Metadata

{table(['key', 'value'], [[key, value] for key, value in payload['product_payload']['metadata'].items()])}

## Tax Code Candidate

`{payload['tax_code_candidate']}` is a candidate. It must be confirmed before live execution.

## Payload Hashes

{table(['field', 'value'], hash_rows)}

## Idempotency Keys

{table(['field', 'value'], key_rows)}

## Safety Checks

- mode: {payload['mode']}
- ready_for_live_execution: {payload['ready_for_live_execution']}
- secret_read: {payload['secret_read']}
- live_api_called: {payload['live_api_called']}
- buyer_information_stored: no
- private_download_url_stored: no
- output_json: `{relative_to_root(json_path)}`

## Live Execution Readiness

{payload['ready_for_live_execution']}

## Errors

{table(['error'], error_rows)}

## Next Command

```powershell
python tools\\release_product.py {payload['product_id']} --execute-live --confirm-product-id {payload['product_id']} --confirm-tax-code --confirmation-text "CREATE LIVE PAYMENT LINK {payload['product_id']}"
```
"""


def validate_dry_run_payload(payload: dict[str, Any]) -> list[str]:
    errors = list(payload.get("errors") or [])
    if payload.get("mode") != "dry-run":
        errors.append("mode must be dry-run")
    if payload.get("secret_read") != "no":
        errors.append("secret_read must be no")
    if payload.get("live_api_called") != "no":
        errors.append("live_api_called must be no")
    for key in ["product_payload", "price_payload", "payment_link_payload"]:
        if not isinstance(payload.get(key), dict):
            errors.append(f"{key} must be an object")
    for payload_key, hash_key in [
        ("product_payload", "product_payload_sha256"),
        ("price_payload", "price_payload_sha256"),
        ("payment_link_payload", "payment_link_payload_sha256"),
    ]:
        part = payload.get(payload_key)
        if isinstance(part, dict) and sha256_payload(part) != payload.get(hash_key):
            errors.append(f"{hash_key} mismatch")
    if contains_secret_like_value(payload):
        errors.append("secret-like value detected")
    if payload.get("ready_for_live_execution") != "yes":
        errors.append("ready_for_live_execution is not yes")
    return errors


def validate_live_confirmation(args: Any, payload: dict[str, Any]) -> None:
    product_id = payload["product_id"]
    if args.confirm_product_id != product_id:
        raise SafetyStop(f"--confirm-product-id must be {product_id}")
    if not args.confirm_tax_code:
        raise SafetyStop("--confirm-tax-code is required")
    expected = f"CREATE LIVE PAYMENT LINK {product_id}"
    if args.confirmation_text != expected:
        raise SafetyStop(f'--confirmation-text must be "{expected}"')


def validate_live_secret_key() -> str:
    value = os.environ.get("STRIPE_SECRET_KEY", "").strip()
    if not value:
        raise SafetyStop("STRIPE_SECRET_KEY is not set")
    if value.startswith("sk_test_"):
        raise SafetyStop("Refusing to run: test mode secret keys are not allowed for live execution")
    if not value.startswith("sk_live_"):
        raise SafetyStop("Refusing to run: STRIPE_SECRET_KEY must start with sk_live_")
    return value


def object_to_dict(obj: Any) -> Any:
    to_dict = getattr(obj, "to_dict", None)
    if callable(to_dict):
        return object_to_dict(to_dict())
    legacy_to_dict_recursive = getattr(obj, "to_dict_recursive", None)
    if callable(legacy_to_dict_recursive):
        return object_to_dict(legacy_to_dict_recursive())
    if isinstance(obj, dict):
        return {key: object_to_dict(value) for key, value in obj.items()}
    if isinstance(obj, list):
        return [object_to_dict(value) for value in obj]
    if isinstance(obj, tuple):
        return [object_to_dict(value) for value in obj]
    return obj


def import_stripe_module() -> Any:
    try:
        import stripe  # type: ignore[import-not-found]
    except ImportError as exc:
        raise SafetyStop("stripe Python SDK is not installed. Install with: pip install stripe") from exc
    return stripe


def initial_state(payload: dict[str, Any], payload_file_hash: str) -> dict[str, Any]:
    return {
        "mode": "live",
        "product_id": payload["product_id"],
        "product_type": payload["product_type"],
        "source_original": payload["source_original"],
        "input_payload_file": relative_to_root(dry_run_json_path(payload["product_id"])),
        "input_payload_file_sha256": payload_file_hash,
        "started_at": now_jst(),
        "updated_at": now_jst(),
        "status": "started",
        "product_id_on_stripe": None,
        "price_id": None,
        "payment_link_id": None,
        "payment_link_url": None,
        "livemode": None,
        "manual_resolution_required": False,
        "error": None,
    }


def load_or_create_state(payload: dict[str, Any], resume: bool) -> dict[str, Any]:
    path = state_json_path(payload["product_id"])
    payload_hash = file_sha256(dry_run_json_path(payload["product_id"]))
    if path.exists():
        if not resume:
            raise SafetyStop(f"state file exists; pass --resume after manual review: {relative_to_root(path)}")
        state = read_json(path)
        if state.get("input_payload_file_sha256") != payload_hash:
            raise SafetyStop("input payload hash differs from existing state; refusing to resume")
        if state.get("status") in {"product_created", "price_created", "failed", "existing_detected"}:
            raise SafetyStop(f"state requires manual resolution before resume: {state.get('status')}")
        if state.get("status") == "completed":
            raise SafetyStop("state is already completed; refusing to run")
        return state
    if resume:
        raise SafetyStop("--resume was passed but state file does not exist")
    state = initial_state(payload, payload_hash)
    write_json_atomic(path, state)
    return state


def save_state(product_id: str, state: dict[str, Any], status: str | None = None) -> None:
    if status:
        state["status"] = status
    state["updated_at"] = now_jst()
    write_json_atomic(state_json_path(product_id), state)


def sanitize_error(exc: BaseException, failed_step: str) -> dict[str, Any]:
    message = str(exc)
    if contains_secret_like_value(message):
        message = "redacted secret-like value"
    return {
        "error_type": type(exc).__name__,
        "safe_message": message[:500],
        "failed_step": failed_step,
        "occurred_at": now_jst(),
        "stripe_request_id": getattr(exc, "request_id", None),
        "stripe_error_code": getattr(exc, "code", None),
    }


def create_with_idempotency(create_func: Callable[..., Any], payload: dict[str, Any], idempotency_key_value: str) -> dict[str, Any]:
    params = copy.deepcopy(payload)
    return object_to_dict(create_func(**params, idempotency_key=idempotency_key_value))


def find_existing_products(stripe: Any, product_id: str) -> list[dict[str, Any]]:
    matches: list[dict[str, Any]] = []
    for product in stripe.Product.list(limit=100).auto_paging_iter():
        data = object_to_dict(product)
        metadata = data.get("metadata") or {}
        if isinstance(metadata, dict) and metadata.get("dake_item_id") == product_id:
            matches.append(data)
    return matches


def find_prices_for_product(stripe: Any, stripe_product_id: str) -> list[str]:
    price_ids: list[str] = []
    for price in stripe.Price.list(product=stripe_product_id, active=True, limit=100).auto_paging_iter():
        data = object_to_dict(price)
        if data.get("id"):
            price_ids.append(str(data["id"]))
    return price_ids


def find_payment_links_for_prices(stripe: Any, price_ids: set[str]) -> list[str]:
    if not price_ids:
        return []
    found: list[str] = []
    for link in stripe.PaymentLink.list(limit=100).auto_paging_iter():
        link_data = object_to_dict(link)
        link_id = link_data.get("id")
        if not link_id:
            continue
        line_items = stripe.PaymentLink.list_line_items(str(link_id), limit=100)
        for line_item in line_items.auto_paging_iter():
            line_data = object_to_dict(line_item)
            price = line_data.get("price") or {}
            price_id = price.get("id") if isinstance(price, dict) else price
            if price_id in price_ids:
                found.append(str(link_id))
                break
    return found


def contains_placeholder(value: Any) -> bool:
    text = canonical_json(value)
    return PRODUCT_PLACEHOLDER in text or PRICE_PLACEHOLDER in text


def write_result(payload: dict[str, Any], state: dict[str, Any]) -> None:
    product_id = payload["product_id"]
    if state.get("status") != "completed":
        raise SafetyStop("result can only be written after completed state")
    result = {
        "mode": "live",
        "product_id": product_id,
        "product_type": payload["product_type"],
        "product_id_on_stripe": state.get("product_id_on_stripe"),
        "price_id": state.get("price_id"),
        "payment_link_id": state.get("payment_link_id"),
        "payment_link_url": state.get("payment_link_url"),
        "livemode": state.get("livemode"),
        "metadata": payload["product_payload"].get("metadata", {}),
        "completed_at": now_jst(),
        "errors": [],
    }
    write_json_atomic(result_json_path(product_id), result)
    write_text_atomic(
        result_md_path(product_id),
        f"""# Product Stripe Release Result

## Summary

- product_id: {product_id}
- product_type: {payload['product_type']}
- livemode: {state.get('livemode')}
- product_id_on_stripe: {state.get('product_id_on_stripe')}
- price_id: {state.get('price_id')}
- payment_link_id: {state.get('payment_link_id')}
- payment_link_url: {state.get('payment_link_url')}
- errors: 0

## Source of Truth

Payment Link URL still must be written back to `{payload['source_original']}` in a separate reviewed phase.
""",
    )


def execute_live_release(payload: dict[str, Any], args: Any) -> int:
    dry_errors = validate_dry_run_payload(payload)
    if dry_errors:
        raise SafetyStop("dry-run payload is not live-ready: " + "; ".join(dry_errors))
    validate_live_confirmation(args, payload)
    secret_key = validate_live_secret_key()
    stripe = import_stripe_module()
    stripe.api_key = secret_key

    product_id = payload["product_id"]
    state = load_or_create_state(payload, args.resume)
    failed_step = "preflight_existing_product"
    try:
        existing = find_existing_products(stripe, product_id)
        if len(existing) == 1:
            product_id_on_stripe = existing[0].get("id")
            price_ids = find_prices_for_product(stripe, str(product_id_on_stripe))
            link_ids = find_payment_links_for_prices(stripe, set(price_ids))
            state.update(
                {
                    "status": "existing_detected",
                    "existing_product_id": product_id_on_stripe,
                    "existing_price_ids": price_ids,
                    "existing_payment_link_ids": link_ids,
                    "manual_resolution_required": True,
                    "error": sanitize_error(SafetyStop("existing live Product detected; manual resolution required"), failed_step),
                }
            )
            save_state(product_id, state)
            return 1
        if len(existing) > 1:
            raise SafetyStop("multiple existing live Products with matching metadata.dake_item_id")

        failed_step = "create_product"
        product = create_with_idempotency(
            stripe.Product.create,
            payload["product_payload"],
            payload["product_idempotency_key"],
        )
        if product.get("livemode") is not True:
            raise SafetyStop("created Product returned livemode=false")
        state["product_id_on_stripe"] = product["id"]
        save_state(product_id, state, "product_created")

        failed_step = "create_price"
        price_payload = copy.deepcopy(payload["price_payload"])
        price_payload["product"] = product["id"]
        if contains_placeholder(price_payload):
            raise SafetyStop("Price payload still contains placeholder")
        price = create_with_idempotency(
            stripe.Price.create,
            price_payload,
            payload["price_idempotency_key"],
        )
        if price.get("livemode") is not True:
            raise SafetyStop("created Price returned livemode=false")
        state["price_id"] = price["id"]
        save_state(product_id, state, "price_created")

        failed_step = "create_payment_link"
        link_payload = copy.deepcopy(payload["payment_link_payload"])
        link_payload["line_items"] = [
            {**line, "price": price["id"]}
            for line in link_payload.get("line_items", [])
        ]
        if contains_placeholder(link_payload):
            raise SafetyStop("Payment Link payload still contains placeholder")
        payment_link = create_with_idempotency(
            stripe.PaymentLink.create,
            link_payload,
            payload["payment_link_idempotency_key"],
        )
        if payment_link.get("livemode") is not True:
            raise SafetyStop("created Payment Link returned livemode=false")
        state["payment_link_id"] = payment_link["id"]
        state["payment_link_url"] = payment_link.get("url")
        state["livemode"] = True
        save_state(product_id, state, "completed")
        write_result(payload, state)
        return 0
    except Exception as exc:
        state["error"] = sanitize_error(exc, failed_step)
        save_state(product_id, state, "failed")
        return 1
