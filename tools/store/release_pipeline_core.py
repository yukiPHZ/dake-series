from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[2]
DEFAULT_STORE_SITE_ROOT = ROOT.parent / "dake-store-site"
ARTIFACT_ROOT = ROOT / "tools" / "reports" / "release_artifacts"
GENERATED_JSON = ROOT / "tools" / "generated" / "store_products.generated.json"
STORE_SITE_JSON = DEFAULT_STORE_SITE_ROOT / "public" / "assets" / "data" / "store_products.generated.json"
COMPLETION_REPORT = ROOT / "tools" / "reports" / "pack2_stripe_release_completion.md"
JST = timezone(timedelta(hours=9))

SOURCE_ROOTS = ("01_apps", "04_packs")
BUY_STRIPE_PREFIX = "https://buy.stripe.com/"

STAGES = [
    "SOURCE_INVALID",
    "PREPARING_BLOCKED",
    "BOOTH_REGISTRATION_PENDING",
    "SOURCE_READY",
    "STRIPE_DRY_RUN_READY",
    "STRIPE_LIVE_COMPLETED",
    "CHECKOUT_REVIEW_PENDING",
    "CHECKOUT_REVIEW_PASSED",
    "SOURCE_FINALIZED",
    "STORE_GENERATED",
    "STORE_SYNC_PENDING",
    "STORE_SYNCED",
    "PRODUCTION_VERIFICATION_PENDING",
    "RELEASE_COMPLETE",
    "LEGACY_COMPLETE",
    "INCONSISTENT",
]

NEXT_ACTIONS = {
    "SOURCE_INVALID": "fix source of truth",
    "PREPARING_BLOCKED": "complete product preparation",
    "BOOTH_REGISTRATION_PENDING": "register product on BOOTH and record the product URL",
    "SOURCE_READY": "run Stripe dry-run",
    "STRIPE_DRY_RUN_READY": "run Stripe live execution with explicit confirmation",
    "STRIPE_LIVE_COMPLETED": "record Checkout browser review",
    "CHECKOUT_REVIEW_PENDING": "record Checkout browser review",
    "CHECKOUT_REVIEW_PASSED": "finalize source of truth",
    "SOURCE_FINALIZED": "regenerate Store JSON",
    "STORE_GENERATED": "sync Store JSON",
    "STORE_SYNC_PENDING": "sync Store JSON",
    "STORE_SYNCED": "verify production Store",
    "PRODUCTION_VERIFICATION_PENDING": "verify production Store",
    "RELEASE_COMPLETE": None,
    "LEGACY_COMPLETE": None,
    "INCONSISTENT": "manual investigation",
}


class SafetyStop(RuntimeError):
    pass


@dataclass
class ProductSource:
    product_id: str
    product_type: str
    source_original: str
    path: Path
    text: str
    title: str
    price: int | None
    status: str
    payment_status: str
    stripe_payment_link: str
    booth_url: str
    github_release_url: str


def now_jst() -> str:
    return datetime.now(JST).isoformat(timespec="seconds")


def normalize_value(value: object) -> str:
    text = str(value or "").strip()
    if len(text) >= 2 and text.startswith("`") and text.endswith("`"):
        text = text[1:-1].strip()
    return text


def is_unset(value: object) -> bool:
    return normalize_value(value).lower() in {"", "not set", "none", "null", "-", "unset", "未設定"}


def metadata_line(text: str, key: str) -> str:
    match = re.search(rf"^\s*[-*]\s*{re.escape(key)}\s*:\s*(.+?)\s*$", text, re.MULTILINE)
    return normalize_value(match.group(1)) if match else ""


def json_like_value(text: str, key: str) -> str:
    match = re.search(rf'"{re.escape(key)}"\s*:\s*"([^"]*)"', text)
    return normalize_value(match.group(1)) if match else ""


def first_url_after(text: str, *labels: str) -> str:
    for label in labels:
        match = re.search(rf"{re.escape(label)}[^\n]*?(https?://[^\s`]+)", text)
        if match:
            return match.group(1).rstrip("`),。")
    return ""


def parse_int(value: object) -> int | None:
    if isinstance(value, int):
        return value if value > 0 else None
    text = str(value or "").replace(",", "")
    match = re.search(r"([0-9][0-9]*)", text)
    if not match:
        return None
    number = int(match.group(1))
    return number if number > 0 else None


def read_json(path: Path) -> dict[str, Any]:
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise SafetyStop(f"JSON root must be an object: {path}")
    return data


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


def relative_to(root: Path, path: Path) -> str:
    try:
        return path.resolve().relative_to(root.resolve()).as_posix()
    except ValueError:
        return str(path)


def safe_artifact_id(product_id: str) -> str:
    safe = re.sub(r"[^A-Za-z0-9_-]+", "_", product_id).strip("_")
    return safe or "item"


def validate_product_id(product_id: str) -> list[str]:
    errors: list[str] = []
    if not product_id or not product_id.strip():
        errors.append("product_id is required")
    if product_id != product_id.strip():
        errors.append("product_id must not contain surrounding whitespace")
    if any(part in product_id for part in ["..", "/", "\\"]):
        errors.append("product_id must not contain path separators or traversal")
    if Path(product_id).is_absolute():
        errors.append("product_id must not be an absolute path")
    if any(ord(char) < 32 for char in product_id):
        errors.append("product_id must not contain control characters")
    return errors


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


class ReleasePipeline:
    def __init__(self, root: Path | None = None, store_site_root: Path | None = None) -> None:
        self.root = (root or ROOT).resolve()
        self.store_site_root = (store_site_root or self.root.parent / "dake-store-site").resolve()
        self.artifact_root = self.root / "tools" / "reports" / "release_artifacts"
        self.generated_json = self.root / "tools" / "generated" / "store_products.generated.json"
        self.store_site_json = self.store_site_root / "public" / "assets" / "data" / "store_products.generated.json"
        self.completion_report = self.root / "tools" / "reports" / "pack2_stripe_release_completion.md"

    def artifact_dir(self, product_id: str) -> Path:
        return self.artifact_root / safe_artifact_id(product_id)

    def original_paths(self) -> list[Path]:
        paths: list[Path] = []
        for root_name in SOURCE_ROOTS:
            source_root = self.root / root_name
            if source_root.exists():
                paths.extend(source_root.glob("**/ORIGINAL.md"))
        return paths

    def parse_original(self, path: Path) -> ProductSource:
        text = path.read_text(encoding="utf-8")
        relative = relative_to(self.root, path)
        is_pack = relative.startswith("04_packs/")
        if is_pack:
            product_id = metadata_line(text, "pack_id") or json_like_value(text, "folder_name") or path.parent.name
            product_type = "pack"
            title = metadata_line(text, "pack_title") or metadata_line(text, "title") or json_like_value(text, "display_name")
        else:
            product_id = metadata_line(text, "app_id") or json_like_value(text, "app_key") or path.parent.name
            product_type = "app"
            title = metadata_line(text, "title") or json_like_value(text, "display_name")
        stripe_payment_link = (
            metadata_line(text, "stripe_payment_link")
            or first_url_after(text, "Stripe Payment Link", "stripe_payment_link")
        )
        if is_unset(stripe_payment_link):
            stripe_payment_link = ""
        booth_url = metadata_line(text, "booth_url") or first_url_after(text, "BOOTH URL", "booth_url")
        if is_unset(booth_url):
            booth_url = ""
        return ProductSource(
            product_id=product_id,
            product_type=product_type,
            source_original=relative,
            path=path,
            text=text,
            title=title or product_id,
            price=parse_int(metadata_line(text, "price")),
            status=metadata_line(text, "status") or json_like_value(text, "status"),
            payment_status=metadata_line(text, "payment_status"),
            stripe_payment_link=stripe_payment_link,
            booth_url=booth_url,
            github_release_url=metadata_line(text, "release_url") or first_url_after(text, "GitHub Release", "release_url"),
        )

    def discover_sources(self) -> tuple[dict[str, ProductSource], dict[str, list[str]]]:
        sources: dict[str, ProductSource] = {}
        duplicates: dict[str, list[str]] = {}
        for path in self.original_paths():
            source = self.parse_original(path)
            if source.product_id in sources:
                duplicates.setdefault(source.product_id, [sources[source.product_id].source_original]).append(source.source_original)
            sources[source.product_id] = source
        return sources, duplicates

    def load_items_json(self, path: Path) -> tuple[dict[str, Any] | None, dict[str, dict[str, Any]], list[str]]:
        if not path.exists():
            return None, {}, []
        try:
            data = read_json(path)
        except Exception as exc:  # noqa: BLE001 - report parser failures as pipeline errors.
            return None, {}, [f"failed to read {relative_to(self.root, path)}: {exc}"]
        items = data.get("items")
        if not isinstance(items, list):
            return data, {}, [f"{relative_to(self.root, path)} must contain items list"]
        result: dict[str, dict[str, Any]] = {}
        duplicates: list[str] = []
        for item in items:
            if not isinstance(item, dict) or not item.get("id"):
                continue
            item_id = str(item["id"])
            if item_id in result:
                duplicates.append(item_id)
            result[item_id] = item
        return data, result, [f"duplicate item in {relative_to(self.root, path)}: {item_id}" for item_id in duplicates]

    def read_optional_json(self, path: Path) -> tuple[dict[str, Any] | None, str | None]:
        if not path.exists():
            return None, None
        try:
            return read_json(path), None
        except Exception as exc:  # noqa: BLE001
            return None, f"failed to read {relative_to(self.root, path)}: {exc}"

    def validate_pack_delivery(self, source: ProductSource) -> tuple[bool, list[str], list[str]]:
        errors: list[str] = []
        warnings: list[str] = []
        manifest_path = source.path.parent / "pack_manifest.json"
        if not manifest_path.exists():
            errors.append("pack_manifest.json is missing")
            return False, errors, warnings
        manifest, manifest_error = self.read_optional_json(manifest_path)
        if manifest_error or manifest is None:
            errors.append(manifest_error or "pack_manifest.json is invalid")
            return False, errors, warnings

        pack_zip = str(manifest.get("pack_zip") or metadata_line(source.text, "delivery_path") or "")
        if not pack_zip:
            errors.append("Pack ZIP path is missing")
        if re.match(r"^[A-Za-z]:\\", pack_zip) or pack_zip.startswith("\\\\") or pack_zip.startswith("/"):
            errors.append("Pack ZIP path must be repo-relative, not absolute")
        zip_path = self.root / pack_zip if pack_zip else None
        if not zip_path or not zip_path.exists():
            errors.append("Pack ZIP is missing")
        else:
            expected_size = parse_int(manifest.get("pack_zip_size"))
            if expected_size is not None and zip_path.stat().st_size != expected_size:
                errors.append("Pack ZIP size does not match manifest")
            expected_hash = str(manifest.get("pack_zip_sha256") or "").lower()
            if expected_hash:
                actual_hash = file_sha256(zip_path).lower()
                if actual_hash != expected_hash:
                    errors.append("Pack ZIP SHA256 does not match manifest")
            else:
                errors.append("Pack ZIP SHA256 is missing from manifest")

        if metadata_line(source.text, "purchase_delivery_ready").lower() != "yes":
            errors.append("purchase_delivery_ready must be yes")
        if not metadata_line(source.text, "purchase_delivery_method"):
            errors.append("purchase_delivery_method is missing")
        if not metadata_line(source.text, "delivery_file_sha256"):
            errors.append("delivery_file_sha256 is missing")
        if "Buyer notice" not in source.text:
            errors.append("Buyer notice is missing")
        if "next business day" not in source.text:
            errors.append("next-business-day delivery notice is missing")
        if "Resend and failure handling" not in source.text:
            errors.append("resend policy is missing")
        if "Manual delivery procedure" not in source.text:
            warnings.append("manual delivery procedure section was not found")
        return not errors, errors, warnings

    def validate_booth_assets(self, source: ProductSource) -> tuple[bool, list[str]]:
        errors: list[str] = []
        pack_dir = source.path.parent
        ready_dir = pack_dir / "pack_ready"
        required_paths = [
            pack_dir / "README.md",
            ready_dir / "README.txt",
            ready_dir / "注意事項.txt",
            pack_dir / "assets" / "booth_thumbnail.jpg",
        ]
        for path in required_paths:
            if not path.exists():
                errors.append(f"missing BOOTH asset: {relative_to(self.root, path)}")
        if not ((pack_dir / "booth_product.txt").exists() or (ready_dir / "booth_product.txt").exists()):
            errors.append("missing BOOTH product text")
        pack_zip_path = ready_dir / f"{source.product_id}.zip"
        if not pack_zip_path.exists():
            errors.append(f"missing Pack ZIP: {relative_to(self.root, pack_zip_path)}")
        return not errors, errors

    def validate_dry_run(self, product_id: str, dry_run: dict[str, Any] | None) -> tuple[bool, list[str]]:
        if dry_run is None:
            return False, []
        errors: list[str] = []
        if dry_run.get("product_id") != product_id:
            errors.append("dry-run product_id does not match target")
        if dry_run.get("mode") != "dry-run":
            errors.append("dry-run mode must be dry-run")
        if dry_run.get("ready_for_live_execution") != "yes":
            errors.append("dry-run is not ready_for_live_execution=yes")
        if dry_run.get("errors") not in ([], None):
            errors.append("dry-run errors must be empty")
        if dry_run.get("secret_read") != "no":
            errors.append("dry-run secret_read must be no")
        if dry_run.get("live_api_called") != "no":
            errors.append("dry-run live_api_called must be no")
        return not errors, errors

    def validate_live(self, product_id: str, result: dict[str, Any] | None, state: dict[str, Any] | None) -> tuple[bool, str, list[str]]:
        if result is None and state is None:
            return False, "", []
        errors: list[str] = []
        if result is None or state is None:
            return False, "", ["live result and state must both exist"]
        if result.get("product_id") != product_id:
            errors.append("result product_id does not match target")
        if state.get("product_id") != product_id:
            errors.append("state product_id does not match target")
        if result.get("livemode") is not True:
            errors.append("result livemode must be true")
        if state.get("livemode") is not True:
            errors.append("state livemode must be true")
        if result.get("errors") not in ([], None):
            errors.append("result errors must be empty")
        if state.get("status") != "completed":
            errors.append("state status must be completed")
        for key in ["product_id_on_stripe", "price_id", "payment_link_id", "payment_link_url"]:
            if not result.get(key):
                errors.append(f"result {key} is missing")
            if not state.get(key):
                errors.append(f"state {key} is missing")
            if result.get(key) and state.get(key) and result.get(key) != state.get(key):
                errors.append(f"result/state mismatch: {key}")
        url = str(result.get("payment_link_url") or "")
        if url and not url.startswith(BUY_STRIPE_PREFIX):
            errors.append("payment_link_url must start with https://buy.stripe.com/")
        if "test_" in url:
            errors.append("payment_link_url must not contain test_")
        return not errors, url, errors

    def validate_checkout(self, product_id: str, product_type: str, review: dict[str, Any] | None) -> tuple[bool, bool, list[str]]:
        if review is None:
            return False, False, []
        errors: list[str] = []
        failed = review.get("review_status") == "failed"
        if review.get("product_id") != product_id:
            errors.append("checkout_review product_id does not match target")
        if review.get("review_status") != "passed":
            errors.append("checkout_review review_status must be passed")
        expected_common: dict[str, Any] = {
            "livemode": True,
            "actual_payment_completed": False,
            "product_name_visible": True,
            "price_visible": True,
            "email_field_present": True,
            "email_required": True,
            "quantity_change_ui": False,
            "shipping_address_required": False,
            "test_url_detected": False,
            "private_download_url_exposed": False,
            "local_path_exposed": False,
        }
        for key, value in expected_common.items():
            if review.get(key) != value:
                errors.append(f"checkout_review {key} must be {value!r}")
        if product_type == "pack":
            for key in ["manual_delivery_notice_visible", "next_business_day_notice_visible"]:
                if review.get(key) is not True:
                    errors.append(f"checkout_review {key} must be True")
        if int(review.get("console_errors") or 0) != 0:
            errors.append("checkout_review console_errors must be 0")
        return not errors, failed, errors

    def production_verified_from_report(self, product_id: str, product_type: str) -> bool:
        if product_type != "pack":
            return False
        if not self.completion_report.exists():
            return False
        text = self.completion_report.read_text(encoding="utf-8", errors="replace")
        lower_text = text.lower()
        return (
            product_id in text
            and "Production Store Review" in text
            and ("console error" in lower_text or "console_errors" in lower_text)
        )

    def validate_production_review(self, product_id: str, product_type: str, review: dict[str, Any] | None) -> tuple[bool, list[str]]:
        if review is None:
            return self.production_verified_from_report(product_id, product_type), []
        errors: list[str] = []
        if review.get("product_id") != product_id:
            errors.append("production_review product_id does not match target")
        if review.get("review_status") != "passed":
            errors.append("production_review review_status must be passed")
        expected = {
            "page_status": 200,
            "product_name_visible": True,
            "price_visible": True,
            "stripe_button_visible": True,
            "payment_link_matches_source": True,
            "booth_link_visible": True,
            "test_url_detected": False,
        }
        for key, value in expected.items():
            if review.get(key) != value:
                errors.append(f"production_review {key} must be {value!r}")
        if int(review.get("console_errors") or 0) != 0:
            errors.append("production_review console_errors must be 0")
        return not errors, errors

    def git_status(self, path: Path) -> str:
        if not (path / ".git").exists():
            return "unknown"
        try:
            result = subprocess.run(
                ["git", "-C", str(path), "status", "--short"],
                check=False,
                capture_output=True,
                text=True,
                timeout=10,
            )
        except Exception:
            return "unknown"
        if result.returncode != 0:
            return "unknown"
        return "clean" if not result.stdout.strip() else "dirty"

    def status(self, product_id: str) -> dict[str, Any]:
        warnings: list[str] = []
        errors: list[str] = validate_product_id(product_id)
        base = {
            "product_id": product_id,
            "product_type": None,
            "current_stage": "SOURCE_INVALID",
            "payment_status": "",
            "stripe_payment_link": "",
            "next_action": NEXT_ACTIONS["SOURCE_INVALID"],
            "checks": {},
            "errors": errors,
            "warnings": warnings,
        }
        if errors:
            return base

        sources, duplicates = self.discover_sources()
        if product_id in duplicates:
            base["errors"].append("duplicate product id in ORIGINAL.md files: " + ", ".join(duplicates[product_id]))
            return base
        source = sources.get(product_id)
        if source is None:
            base["errors"].append(f"Product id not found: {product_id}")
            return base

        generated_root, generated_items, generated_errors = self.load_items_json(self.generated_json)
        site_root, site_items, site_errors = self.load_items_json(self.store_site_json)
        generated_item = generated_items.get(product_id, {})
        site_item = site_items.get(product_id, {})
        payment_status = source.payment_status or str(generated_item.get("payment_status") or "")
        stripe_payment_link = source.stripe_payment_link or str(generated_item.get("stripe_payment_link") or "")
        if is_unset(stripe_payment_link):
            stripe_payment_link = ""
        price = source.price or parse_int(generated_item.get("price"))

        artifact_dir = self.artifact_dir(product_id)
        dry_run, dry_run_error = self.read_optional_json(artifact_dir / "stripe_release_dry_run.json")
        result, result_error = self.read_optional_json(artifact_dir / "stripe_release_result.json")
        state, state_error = self.read_optional_json(artifact_dir / "stripe_release_state.json")
        checkout, checkout_error = self.read_optional_json(artifact_dir / "checkout_review.json")
        finalize_result, finalize_error = self.read_optional_json(artifact_dir / "finalize_release_result.json")
        production_review, production_error = self.read_optional_json(artifact_dir / "production_review.json")

        errors.extend(generated_errors + site_errors)
        for optional_error in [dry_run_error, result_error, state_error, checkout_error, finalize_error, production_error]:
            if optional_error:
                errors.append(optional_error)

        delivery_ready = True
        booth_assets_ready = False
        if source.product_type == "pack":
            delivery_ready, delivery_errors, delivery_warnings = self.validate_pack_delivery(source)
            errors.extend(delivery_errors)
            warnings.extend(delivery_warnings)
            booth_assets_ready, booth_asset_errors = self.validate_booth_assets(source)
            if payment_status == "preparing":
                errors.extend(booth_asset_errors)

        if price is None and payment_status != "preparing":
            errors.append("price is missing or invalid")
        if source.product_type not in {"app", "pack"}:
            errors.append(f"unsupported product type: {source.product_type}")

        dry_run_ready, dry_run_errors = self.validate_dry_run(product_id, dry_run)
        live_completed, live_url, live_errors = self.validate_live(product_id, result, state)
        checkout_passed, checkout_failed, checkout_errors = self.validate_checkout(product_id, source.product_type, checkout)
        production_verified, production_errors = self.validate_production_review(product_id, source.product_type, production_review)
        artifact_validation_errors = dry_run_errors + live_errors + checkout_errors + production_errors

        source_finalized = payment_status == "stripe_ready" and bool(stripe_payment_link)
        generated_status = str(generated_item.get("payment_status") or "")
        generated_link = str(generated_item.get("stripe_payment_link") or "")
        if is_unset(generated_link):
            generated_link = ""
        site_status = str(site_item.get("payment_status") or "")
        site_link = str(site_item.get("stripe_payment_link") or "")
        if is_unset(site_link):
            site_link = ""

        store_generated = False
        store_synced = False
        inconsistent: list[str] = []
        if payment_status == "stripe_ready" and not stripe_payment_link:
            inconsistent.append("source payment_status is stripe_ready but stripe_payment_link is missing")
        inconsistent.extend(artifact_validation_errors)
        if live_url and stripe_payment_link and live_url != stripe_payment_link:
            inconsistent.append("source stripe_payment_link does not match live result URL")
        if checkout_failed and source_finalized:
            inconsistent.append("checkout_review failed but source is finalized")
        if finalize_result and finalize_result.get("errors") not in ([], None):
            inconsistent.append("finalize result contains errors")

        if source_finalized:
            if not generated_item:
                store_generated = False
            elif generated_link == stripe_payment_link and generated_status == "stripe_ready":
                store_generated = True
            elif generated_link and generated_link != stripe_payment_link:
                inconsistent.append("generated Store JSON stripe_payment_link does not match source")
            elif generated_status and generated_status != payment_status:
                inconsistent.append("generated Store JSON payment_status does not match source")
        elif payment_status == "preparing":
            store_generated = generated_status == "preparing"
        elif generated_item and not stripe_payment_link:
            store_generated = True

        if store_generated and site_item:
            if site_link and generated_link and site_link != generated_link:
                inconsistent.append("dake-store-site JSON stripe_payment_link does not match generated JSON")
            elif site_status and generated_status and site_status != generated_status:
                inconsistent.append("dake-store-site JSON payment_status does not match generated JSON")
            else:
                store_synced = site_item == generated_item

        legacy_complete = (
            source.product_type == "app"
            and not dry_run
            and not result
            and source_finalized
            and store_generated
            and store_synced
        )
        release_complete = (
            live_completed
            and checkout_passed
            and source_finalized
            and store_generated
            and store_synced
            and production_verified
        )

        checks = {
            "source_valid": not errors and not inconsistent,
            "delivery_ready": delivery_ready if source.product_type == "pack" else "not_applicable",
            "pack_zip_ready": delivery_ready if source.product_type == "pack" else "not_applicable",
            "booth_assets_ready": booth_assets_ready if source.product_type == "pack" else "not_applicable",
            "stripe_dry_run_ready": dry_run_ready,
            "stripe_live_completed": live_completed,
            "checkout_review_passed": checkout_passed,
            "source_finalized": source_finalized,
            "store_generated": store_generated,
            "store_synced": store_synced,
            "production_verified": production_verified,
            "legacy_complete": legacy_complete,
            "dake_series_git": self.git_status(self.root),
            "dake_store_site_git": self.git_status(self.store_site_root),
        }

        stage = "SOURCE_READY"
        if errors:
            stage = "SOURCE_INVALID"
        elif inconsistent:
            errors.extend(inconsistent)
            stage = "INCONSISTENT"
        elif (
            source.product_type == "pack"
            and payment_status == "preparing"
            and delivery_ready
            and booth_assets_ready
            and not source.booth_url
            and not source.stripe_payment_link
        ):
            stage = "BOOTH_REGISTRATION_PENDING"
        elif payment_status == "preparing":
            stage = "PREPARING_BLOCKED"
        elif legacy_complete:
            stage = "LEGACY_COMPLETE"
        elif release_complete:
            stage = "RELEASE_COMPLETE"
        elif live_completed and checkout_passed and source_finalized and store_generated and store_synced:
            stage = "PRODUCTION_VERIFICATION_PENDING"
        elif live_completed and checkout_passed and source_finalized and store_generated:
            stage = "STORE_SYNC_PENDING" if site_item else "STORE_GENERATED"
        elif live_completed and checkout_passed and source_finalized:
            stage = "SOURCE_FINALIZED"
        elif live_completed and checkout_passed:
            stage = "CHECKOUT_REVIEW_PASSED"
        elif live_completed:
            stage = "CHECKOUT_REVIEW_PENDING"
        elif dry_run_ready:
            stage = "STRIPE_DRY_RUN_READY"
        else:
            stage = "SOURCE_READY"

        return {
            "product_id": product_id,
            "product_type": source.product_type,
            "source_original": source.source_original,
            "title": source.title,
            "price": price,
            "current_stage": stage,
            "payment_status": payment_status,
            "stripe_payment_link": "present" if stripe_payment_link else "missing",
            "stripe_payment_link_url": stripe_payment_link,
            "booth_url": "present" if source.booth_url else "missing",
            "booth_url_value": source.booth_url,
            "stripe_result": "completed" if live_completed else "missing",
            "checkout_review": "passed" if checkout_passed else ("failed" if checkout_failed else "missing"),
            "next_action": NEXT_ACTIONS[stage],
            "checks": checks,
            "errors": errors,
            "warnings": warnings,
            "paths": {
                "artifact_dir": relative_to(self.root, artifact_dir),
                "generated_json": relative_to(self.root, self.generated_json),
                "store_site_json": str(self.store_site_json),
            },
        }

    def save_status_report(self, status: dict[str, Any]) -> tuple[Path, Path]:
        product_id = str(status["product_id"])
        artifact_dir = self.artifact_dir(product_id)
        json_path = artifact_dir / "pipeline_status.json"
        md_path = artifact_dir / "pipeline_status.md"
        report = {**status, "saved_at": now_jst()}
        write_json_atomic(json_path, report)
        write_text_atomic(md_path, render_status_markdown(report))
        return json_path, md_path

    def advance(self, product_id: str) -> dict[str, Any]:
        status = self.status(product_id)
        stage = status["current_stage"]
        if stage in {"SOURCE_INVALID", "PREPARING_BLOCKED", "INCONSISTENT"}:
            return {"advanced": False, "stage": stage, "message": "advance refused: " + str(status["next_action"])}
        if stage in {"RELEASE_COMPLETE", "LEGACY_COMPLETE"}:
            return {"advanced": False, "stage": stage, "message": "no action needed"}
        if stage == "SOURCE_READY":
            command = [sys.executable, str(self.root / "tools" / "release_product.py"), product_id]
            result = subprocess.run(command, cwd=self.root, check=False)
            return {"advanced": result.returncode == 0, "stage": stage, "command": command, "returncode": result.returncode}
        if stage == "STRIPE_DRY_RUN_READY":
            command_text = (
                f'python tools\\release_product.py {product_id} --execute-live '
                f'--confirm-product-id {product_id} --confirm-tax-code '
                f'--confirmation-text "CREATE LIVE PAYMENT LINK {product_id}"'
            )
            return {"advanced": False, "stage": stage, "message": "manual live execution required", "command": command_text}
        if stage == "BOOTH_REGISTRATION_PENDING":
            return {
                "advanced": False,
                "stage": stage,
                "message": "manual BOOTH registration required",
                "command": f"python tools\\record_booth_registration.py {product_id} --booth-url <BOOTH商品URL>",
            }
        if stage in {"STRIPE_LIVE_COMPLETED", "CHECKOUT_REVIEW_PENDING"}:
            return {"advanced": False, "stage": stage, "message": "manual Checkout review required", "command": f"python tools\\record_checkout_review.py {product_id}"}
        if stage == "CHECKOUT_REVIEW_PASSED":
            command = [sys.executable, str(self.root / "tools" / "finalize_product_release.py"), product_id]
            result = subprocess.run(command, cwd=self.root, check=False)
            return {"advanced": result.returncode == 0, "stage": stage, "command": command, "returncode": result.returncode}
        if stage == "SOURCE_FINALIZED":
            command = [sys.executable, str(self.root / "tools" / "store" / "generate_store_products.py")]
            result = subprocess.run(command, cwd=self.root, check=False)
            return {"advanced": result.returncode == 0, "stage": stage, "command": command, "returncode": result.returncode}
        if stage in {"STORE_GENERATED", "STORE_SYNC_PENDING"}:
            return {"advanced": False, "stage": stage, "message": "manual Store sync required", "command": "python tools\\store\\sync_store_to_site.py"}
        if stage in {"STORE_SYNCED", "PRODUCTION_VERIFICATION_PENDING"}:
            return {"advanced": False, "stage": stage, "message": "manual production verification required"}
        return {"advanced": False, "stage": stage, "message": "no advance rule for stage"}


def render_status_markdown(status: dict[str, Any]) -> str:
    errors = "\n".join(f"- {error}" for error in status.get("errors", [])) or "- none"
    warnings = "\n".join(f"- {warning}" for warning in status.get("warnings", [])) or "- none"
    checks = "\n".join(f"- {key}: {value}" for key, value in status.get("checks", {}).items()) or "- none"
    return f"""# Product Release Pipeline Status

## Summary

- product_id: {status.get('product_id')}
- product_type: {status.get('product_type')}
- current_stage: {status.get('current_stage')}
- payment_status: {status.get('payment_status')}
- stripe_payment_link: {status.get('stripe_payment_link')}
- next_action: {status.get('next_action') or 'none'}

## Checks

{checks}

## Errors

{errors}

## Warnings

{warnings}
"""


def print_status(status: dict[str, Any]) -> None:
    print(f"product_id: {status.get('product_id')}")
    print(f"product_type: {status.get('product_type')}")
    print(f"current_stage: {status.get('current_stage')}")
    print(f"payment_status: {status.get('payment_status')}")
    print(f"stripe_payment_link: {status.get('stripe_payment_link')}")
    print(f"booth_url: {status.get('booth_url', 'unknown')}")
    checks = status.get("checks", {})
    for key in [
        "delivery_ready",
        "pack_zip_ready",
        "booth_assets_ready",
        "stripe_dry_run_ready",
        "stripe_live_completed",
        "checkout_review_passed",
        "source_finalized",
        "store_generated",
        "store_synced",
        "production_verified",
    ]:
        if key in checks:
            print(f"{key}: {checks[key]}")
    print(f"DAKE_series: {checks.get('dake_series_git', 'unknown')}")
    print(f"dake-store-site: {checks.get('dake_store_site_git', 'unknown')}")
    print(f"next_action: {status.get('next_action') or 'none'}")
    if status.get("errors"):
        print("errors:")
        for error in status["errors"]:
            print(f"- {error}")
    if status.get("warnings"):
        print("warnings:")
        for warning in status["warnings"]:
            print(f"- {warning}")


def cli_json(status: dict[str, Any]) -> str:
    return json.dumps(status, ensure_ascii=False, indent=2)


def add_common_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("product_id", help="DAKE product id")
    parser.add_argument("command", nargs="?", default="status", choices=["status", "next", "advance"], help="Pipeline command")
    parser.add_argument("--json", action="store_true", help="Print status as JSON")
    parser.add_argument("--save-report", action="store_true", help="Save pipeline status report under release_artifacts")
