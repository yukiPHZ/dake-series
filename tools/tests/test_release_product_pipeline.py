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


def write_pack(root: Path, product_id: str, payment_status: str = "booth_only", stripe_link: str = "not set") -> None:
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
            "booth_url": "https://peakheadz.booth.pm/items/synthetic-test",
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
- booth_url: https://peakheadz.booth.pm/items/synthetic-test

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


def test_real_repo_compatibility() -> None:
    pipeline = ReleasePipeline()
    expected = {
        "DAKE_Pack_Document": "RELEASE_COMPLETE",
        "DAKE_Pack_Memo": "RELEASE_COMPLETE",
        "dake_pdf_viewer": "LEGACY_COMPLETE",
        "video_shorts_cut": "PREPARING_BLOCKED",
    }
    for product_id, stage in expected.items():
        status = pipeline.status(product_id)
        assert status["current_stage"] == stage, (product_id, status)

    assert pipeline.status("DOES_NOT_EXIST")["current_stage"] == "SOURCE_INVALID"
    assert pipeline.status(r"..\..\example")["current_stage"] == "SOURCE_INVALID"


def test_generated_counts() -> None:
    data = json.loads((ROOT / "tools" / "generated" / "store_products.generated.json").read_text(encoding="utf-8"))
    items = data["items"]
    assert len(items) == 53
    counts: dict[str, int] = {}
    for item in items:
        counts[item.get("payment_status") or ""] = counts.get(item.get("payment_status") or "", 0) + 1
    assert counts.get("stripe_ready") == 52
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


def main() -> int:
    tests = [
        test_real_repo_compatibility,
        test_generated_counts,
        test_synthetic_pack_discovery_and_advance_dispatch,
        test_inconsistent_fixtures,
        test_duplicate_product_id,
    ]
    for test in tests:
        test()
        print(f"{test.__name__}=passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
