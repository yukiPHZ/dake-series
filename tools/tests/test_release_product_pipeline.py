from __future__ import annotations

import hashlib
import json
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "tools"))

import record_booth_registration  # noqa: E402
from store.release_pipeline_core import ReleasePipeline  # noqa: E402


def write_json(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def make_store_json(path: Path, items: list[dict]) -> None:
    write_json(
        path,
        {
            "generated_at": "2026-06-16T00:00:00+09:00",
            "source_policy": "test fixture",
            "schema_version": "1.0.0",
            "do_not_edit": True,
            "items": items,
        },
    )


def write_pack(
    root: Path,
    product_id: str,
    payment_status: str = "booth_only",
    stripe_link: str = "not set",
    booth_url: str = "https://peakheadz.booth.pm/items/synthetic-test",
) -> None:
    pack_dir = root / "04_packs" / product_id
    pack_ready = pack_dir / "pack_ready"
    pack_ready.mkdir(parents=True, exist_ok=True)
    zip_relative = f"04_packs/{product_id}/pack_ready/{product_id}.zip"
    zip_path = root / zip_relative
    with zipfile.ZipFile(zip_path, "w") as archive:
        archive.writestr("README.txt", "Synthetic fixture.\n")
    zip_hash = sha256_file(zip_path)
    write_json(
        pack_dir / "pack_manifest.json",
        {
            "schema": "dake_pack_manifest_v1",
            "folder_name": product_id,
            "display_name": "Synthetic Pack",
                "status": "available",
                "price": 1234,
                "booth_url": booth_url,
            "included_apps": [],
            "pack_zip": zip_relative,
            "pack_zip_size": zip_path.stat().st_size,
            "pack_zip_sha256": zip_hash,
        },
    )
    (pack_dir / "ORIGINAL.md").write_text(
        f"""# ORIGINAL.md

## Basic

- pack_id: `{product_id}`
- title: Synthetic Pack
- status: available
- price: 1234
- booth_url: {booth_url}

## Stripe manual delivery operation

- payment_status: {payment_status}
- stripe_payment_link: {stripe_link}
- purchase_delivery_method: manual_email_private_download
- purchase_delivery_ready: yes
- delivery_file_sha256: `{zip_hash}`

### Buyer notice

This Pack is a digital product. After Stripe payment is confirmed, DAKE sends download instructions to the email address entered at purchase.

This is not automatic download. The standard delivery window is within the next business day after payment confirmation.

### Manual delivery procedure

Confirm payment and send the private download instructions by email.

### Resend and failure handling

If the buyer requests resend, verify the original payment, Pack, buyer email address, and previous delivery record before resending.
""",
        encoding="utf-8",
        newline="\n",
    )


def write_booth_assets(root: Path, product_id: str) -> None:
    pack_dir = root / "04_packs" / product_id
    ready_dir = pack_dir / "pack_ready"
    (pack_dir / "assets").mkdir(parents=True, exist_ok=True)
    (pack_dir / "README.md").write_text("# Synthetic Pack\n", encoding="utf-8")
    (pack_dir / "booth_product.txt").write_text("Synthetic product text\n", encoding="utf-8")
    (pack_dir / "assets" / "booth_thumbnail.jpg").write_bytes(b"fake-jpg-for-pipeline-test")
    (ready_dir / "README.txt").write_text("Synthetic README\n", encoding="utf-8")
    (ready_dir / "注意事項.txt").write_text("Synthetic notice\n", encoding="utf-8")


def write_dummy_release_product(root: Path) -> Path:
    marker = root / "dummy_release_product_called.txt"
    script = root / "tools" / "release_product.py"
    script.parent.mkdir(parents=True, exist_ok=True)
    script.write_text(
        "from pathlib import Path\n"
        f"Path({str(marker)!r}).write_text('called', encoding='utf-8')\n"
        "raise SystemExit(0)\n",
        encoding="utf-8",
    )
    return marker


def write_live_artifacts(root: Path, product_id: str, url: str, state_status: str = "completed", checkout_status: str = "passed") -> None:
    artifact = root / "tools" / "reports" / "release_artifacts" / product_id
    write_json(
        artifact / "stripe_release_result.json",
        {
            "mode": "live",
            "status": "completed",
            "product_id": product_id,
            "product_type": "pack",
            "product_id_on_stripe": "prod_fixture",
            "price_id": "price_fixture",
            "payment_link_id": "plink_fixture",
            "payment_link_url": url,
            "livemode": True,
            "errors": [],
        },
    )
    write_json(
        artifact / "stripe_release_state.json",
        {
            "mode": "live",
            "product_id": product_id,
            "product_type": "pack",
            "product_id_on_stripe": "prod_fixture",
            "price_id": "price_fixture",
            "payment_link_id": "plink_fixture",
            "payment_link_url": url,
            "livemode": True,
            "status": state_status,
        },
    )
    write_json(
        artifact / "checkout_review.json",
        {
            "product_id": product_id,
            "review_status": checkout_status,
            "livemode": True,
            "actual_payment_completed": False,
            "product_name_visible": True,
            "price_visible": True,
            "email_field_present": True,
            "email_required": True,
            "quantity_change_ui": False,
            "shipping_address_required": False,
            "manual_delivery_notice_visible": True,
            "next_business_day_notice_visible": True,
            "test_url_detected": False,
            "private_download_url_exposed": False,
            "local_path_exposed": False,
            "console_errors": 0,
        },
    )


def write_production_review(root: Path, product_id: str) -> None:
    artifact = root / "tools" / "reports" / "release_artifacts" / product_id
    write_json(
        artifact / "production_review.json",
        {
            "product_id": product_id,
            "review_status": "passed",
            "page_status": 200,
            "product_name_visible": True,
            "price_visible": True,
            "stripe_button_visible": True,
            "payment_link_matches_source": True,
            "booth_link_visible": True,
            "booth_link_correct": True,
            "manual_delivery_notice_visible": True,
            "next_business_day_notice_visible": True,
            "test_url_detected": False,
            "private_url_exposed": False,
            "local_path_exposed": False,
            "zip_url_exposed": False,
            "actual_payment_completed": False,
            "console_errors": 0,
        },
    )


def write_social_artifact(root: Path, product_id: str) -> None:
    artifact = root / "tools" / "reports" / "release_artifacts" / product_id
    write_json(
        artifact / "social_release.json",
        {
            "product_id": product_id,
            "product_type": "pack",
            "app_key": product_id,
            "stage": "complete",
            "published": False,
            "scheduled": False,
            "requested_channels": ["x", "threads", "instagram"],
            "buffer": {
                "x": {"status": "draft_created", "buffer_post_id": "x_fixture"},
                "threads": {"status": "draft_created", "buffer_post_id": "threads_fixture"},
                "instagram": {"status": "draft_created", "buffer_post_id": "instagram_fixture"},
            },
        },
    )


def test_real_repo_compatibility() -> None:
    pipeline = ReleasePipeline()
    expected = {
        "DAKE_Pack_Document": "RELEASE_COMPLETE",
        "DAKE_Pack_Memo": "RELEASE_COMPLETE",
        "DAKE_Pack_Mail": "RELEASE_COMPLETE",
        "dake_pdf_viewer": "LEGACY_COMPLETE",
        "video_shorts_cut": "PREPARING_BLOCKED",
    }
    for product_id, stage in expected.items():
        status = pipeline.status(product_id)
        assert status["current_stage"] == stage, (product_id, status)
    mail_status = pipeline.status("DAKE_Pack_Mail")
    assert mail_status["commerce_status"] == "complete", mail_status
    assert mail_status["formal_release_status"] in {"v2_closed", "v2_pending"}, mail_status

    assert pipeline.status("DOES_NOT_EXIST")["current_stage"] == "SOURCE_INVALID"
    assert pipeline.status(r"..\..\example")["current_stage"] == "SOURCE_INVALID"


def test_generated_counts() -> None:
    data = json.loads((ROOT / "tools" / "generated" / "store_products.generated.json").read_text(encoding="utf-8"))
    items = data["items"]
    assert len(items) == 54
    counts: dict[str, int] = {}
    for item in items:
        counts[item.get("payment_status") or ""] = counts.get(item.get("payment_status") or "", 0) + 1
    assert counts.get("stripe_ready") == 53
    assert counts.get("booth_only", 0) == 0
    assert counts.get("preparing") == 1


def test_synthetic_pack_discovery_and_advance_dispatch() -> None:
    with tempfile.TemporaryDirectory() as temp:
        root = Path(temp) / "DAKE_series"
        store = Path(temp) / "dake-store-site"
        write_pack(root, "Synthetic_Pack_Example")
        marker = write_dummy_release_product(root)
        pipeline = ReleasePipeline(root=root, store_site_root=store)
        status = pipeline.status("Synthetic_Pack_Example")
        assert status["current_stage"] == "SOURCE_READY", status
        assert status["next_action"] == "run Stripe dry-run"
        result = pipeline.advance("Synthetic_Pack_Example")
        assert result["advanced"] is True, result
        assert marker.read_text(encoding="utf-8") == "called"


def run_booth_command(root: Path, store: Path, argv: list[str]) -> int:
    old_argv = sys.argv[:]
    old_pipeline = record_booth_registration.ReleasePipeline
    try:
        record_booth_registration.ReleasePipeline = lambda: ReleasePipeline(root=root, store_site_root=store)  # type: ignore[assignment]
        sys.argv = ["record_booth_registration.py", *argv]
        return record_booth_registration.main()
    finally:
        sys.argv = old_argv
        record_booth_registration.ReleasePipeline = old_pipeline


def test_booth_registration_pending_and_record_command() -> None:
    with tempfile.TemporaryDirectory() as temp:
        root = Path(temp) / "DAKE_series"
        store = Path(temp) / "dake-store-site"
        product_id = "Pending_Mail_Pack"
        write_pack(root, product_id, payment_status="preparing", stripe_link="未設定", booth_url="未設定")
        write_booth_assets(root, product_id)
        pipeline = ReleasePipeline(root=root, store_site_root=store)
        status = pipeline.status(product_id)
        assert status["current_stage"] == "BOOTH_REGISTRATION_PENDING", status

        booth_url = "https://peakheadz.booth.pm/items/1234567"
        assert run_booth_command(root, store, [product_id, "--booth-url", booth_url]) == 0
        original_path = root / "04_packs" / product_id / "ORIGINAL.md"
        before = original_path.read_text(encoding="utf-8")
        assert "payment_status: preparing" in before
        assert run_booth_command(root, store, [product_id, "--booth-url", booth_url, "--apply"]) == 1
        assert "payment_status: preparing" in original_path.read_text(encoding="utf-8")
        assert run_booth_command(
            root,
            store,
            [
                product_id,
                "--booth-url",
                booth_url,
                "--apply",
                "--confirm-product-id",
                product_id,
                "--confirmation-text",
                f"RECORD BOOTH REGISTRATION {product_id}",
            ],
        ) == 0
        after = original_path.read_text(encoding="utf-8")
        assert f"booth_url: {booth_url}" in after
        assert "status: available" in after
        assert "payment_status: booth_only" in after
        assert "stripe_payment_link: 未設定" in after


def test_booth_registration_rejections() -> None:
    bad_urls = [
        "http://peakheadz.booth.pm/items/123",
        "https://example.com/items/123",
        "https://peakheadz.booth.pm/product/123",
        "https://peakheadz.booth.pm/items/not-number",
        "https://peakheadz.booth.pm/items/123?x=1",
    ]
    for bad_url in bad_urls:
        try:
            record_booth_registration.validate_booth_url(bad_url)
        except record_booth_registration.SafetyStop:
            pass
        else:
            raise AssertionError(f"bad URL accepted: {bad_url}")

    with tempfile.TemporaryDirectory() as temp:
        root = Path(temp) / "DAKE_series"
        store = Path(temp) / "dake-store-site"
        write_pack(root, "Target_Pack", payment_status="preparing", stripe_link="未設定", booth_url="未設定")
        write_booth_assets(root, "Target_Pack")
        write_pack(root, "Owner_Pack", payment_status="booth_only", stripe_link="未設定", booth_url="https://peakheadz.booth.pm/items/7654321")
        write_booth_assets(root, "Owner_Pack")
        assert run_booth_command(root, store, ["Target_Pack", "--booth-url", "https://peakheadz.booth.pm/items/7654321"]) == 1
        assert run_booth_command(root, store, ["Owner_Pack", "--booth-url", "https://peakheadz.booth.pm/items/9999999"]) == 1


def test_inconsistent_fixtures() -> None:
    with tempfile.TemporaryDirectory() as temp:
        root = Path(temp) / "DAKE_series"
        store = Path(temp) / "dake-store-site"
        pipeline = ReleasePipeline(root=root, store_site_root=store)

        write_pack(root, "Ready_No_Link", payment_status="stripe_ready", stripe_link="not set")
        assert pipeline.status("Ready_No_Link")["current_stage"] == "INCONSISTENT"

        write_pack(root, "Url_Mismatch", payment_status="stripe_ready", stripe_link="https://buy.stripe.com/source")
        write_live_artifacts(root, "Url_Mismatch", "https://buy.stripe.com/result")
        assert pipeline.status("Url_Mismatch")["current_stage"] == "INCONSISTENT"

        write_pack(root, "State_Failed")
        write_live_artifacts(root, "State_Failed", "https://buy.stripe.com/statefailed", state_status="failed")
        assert pipeline.status("State_Failed")["current_stage"] == "INCONSISTENT"

        write_pack(root, "Checkout_Failed", payment_status="stripe_ready", stripe_link="https://buy.stripe.com/checkoutfailed")
        write_live_artifacts(root, "Checkout_Failed", "https://buy.stripe.com/checkoutfailed", checkout_status="failed")
        assert pipeline.status("Checkout_Failed")["current_stage"] == "INCONSISTENT"

        write_pack(root, "Store_Wrong", payment_status="stripe_ready", stripe_link="https://buy.stripe.com/source")
        make_store_json(
            root / "tools" / "generated" / "store_products.generated.json",
            [
                {
                    "id": "Store_Wrong",
                    "type": "pack",
                    "payment_status": "stripe_ready",
                    "stripe_payment_link": "https://buy.stripe.com/wrong",
                    "price": 1234,
                }
            ],
        )
        assert pipeline.status("Store_Wrong")["current_stage"] == "INCONSISTENT"


def test_duplicate_product_id() -> None:
    with tempfile.TemporaryDirectory() as temp:
        root = Path(temp) / "DAKE_series"
        store = Path(temp) / "dake-store-site"
        write_pack(root, "Duplicate_Item")
        duplicate = root / "04_packs" / "Duplicate_Item_Copy"
        shutil.copytree(root / "04_packs" / "Duplicate_Item", duplicate)
        status = ReleasePipeline(root=root, store_site_root=store).status("Duplicate_Item")
        assert status["current_stage"] == "SOURCE_INVALID", status
        assert "duplicate product id" in "\n".join(status["errors"])


def test_formal_release_status_from_social_artifact() -> None:
    with tempfile.TemporaryDirectory() as temp:
        root = Path(temp) / "DAKE_series"
        store = Path(temp) / "dake-store-site"
        product_id = "Formal_Pack"
        stripe_url = "https://buy.stripe.com/formalpack"
        booth_url = "https://peakheadz.booth.pm/items/1234567"
        write_pack(root, product_id, payment_status="stripe_ready", stripe_link=stripe_url, booth_url=booth_url)
        write_booth_assets(root, product_id)
        write_live_artifacts(root, product_id, stripe_url)
        write_production_review(root, product_id)
        item = {
            "id": product_id,
            "type": "pack",
            "title": "Formal Pack",
            "price": 1234,
            "payment_status": "stripe_ready",
            "stripe_payment_link": stripe_url,
            "booth_url": booth_url,
        }
        make_store_json(root / "tools" / "generated" / "store_products.generated.json", [item])
        make_store_json(store / "public" / "assets" / "data" / "store_products.generated.json", [item])

        pipeline = ReleasePipeline(root=root, store_site_root=store)
        pending = pipeline.status(product_id)
        assert pending["current_stage"] == "RELEASE_COMPLETE", pending
        assert pending["commerce_status"] == "complete", pending
        assert pending["social_status"] == "buffer_drafts_pending", pending
        assert pending["formal_release_status"] == "v2_pending", pending
        assert pending["next_formal_action"] == "create Buffer drafts", pending

        write_social_artifact(root, product_id)
        closed = pipeline.status(product_id)
        assert closed["current_stage"] == "RELEASE_COMPLETE", closed
        assert closed["social_status"] == "buffer_drafts_complete", closed
        assert closed["formal_release_status"] == "v2_closed", closed
        assert closed["next_formal_action"] is None, closed


def main() -> int:
    tests = [
        test_real_repo_compatibility,
        test_generated_counts,
        test_synthetic_pack_discovery_and_advance_dispatch,
        test_booth_registration_pending_and_record_command,
        test_booth_registration_rejections,
        test_inconsistent_fixtures,
        test_duplicate_product_id,
        test_formal_release_status_from_social_artifact,
    ]
    for test in tests:
        test()
        print(f"{test.__name__}=passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
