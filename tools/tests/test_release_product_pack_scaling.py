from __future__ import annotations

import hashlib
import json
import shutil
import sys
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

from tools.store.stripe_release_core import build_release_payload

PACK_ID = "Synthetic_Pack_Example"
PACK_DIR = ROOT / "04_packs" / PACK_ID
ZIP_RELATIVE = f"04_packs/{PACK_ID}/pack_ready/{PACK_ID}.zip"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def write_synthetic_pack() -> None:
    pack_ready = PACK_DIR / "pack_ready"
    pack_ready.mkdir(parents=True, exist_ok=True)
    zip_path = ROOT / ZIP_RELATIVE
    with zipfile.ZipFile(zip_path, "w") as archive:
        archive.writestr("README.txt", "Synthetic pack test artifact.\n")
    zip_size = zip_path.stat().st_size
    zip_hash = sha256_file(zip_path)

    (PACK_DIR / "pack_manifest.json").write_text(
        json.dumps(
            {
                "schema": "dake_pack_manifest_v1",
                "folder_name": PACK_ID,
                "display_name": "Synthetic Pack Example",
                "status": "available",
                "price": 1234,
                "booth_url": "https://peakheadz.booth.pm/items/synthetic-test",
                "included_apps": [],
                "pack_zip": ZIP_RELATIVE,
                "pack_zip_size": zip_size,
                "pack_zip_sha256": zip_hash,
            },
            ensure_ascii=False,
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )
    (PACK_DIR / "ORIGINAL.md").write_text(
        f"""# ORIGINAL.md

## Basic

- pack_id: `{PACK_ID}`
- title: Synthetic Pack Example
- status: available
- price: 1234
- booth_url: https://peakheadz.booth.pm/items/synthetic-test

## Stripe manual delivery operation

- payment_status: booth_only
- stripe_payment_link: not set
- purchase_delivery_method: manual_email_private_download
- purchase_delivery_ready: yes
- stripe_creation_method: manual_dashboard_ready
- review_result: ready
- delivery_rule: `00_core/DAKE_PACK_MANUAL_DELIVERY_RULE.md`

### Buyer notice

This Pack is a digital product. After Stripe payment is confirmed, DAKE sends download instructions to the email address entered at purchase.

This is not automatic download. The standard delivery window is within the next business day after payment confirmation.

### Resend and failure handling

If the buyer requests resend, verify the original payment, Pack, buyer email address, and previous delivery record before resending.
""",
        encoding="utf-8",
        newline="\n",
    )


def main() -> int:
    if PACK_DIR.exists():
        shutil.rmtree(PACK_DIR)
    try:
        write_synthetic_pack()
        payload = build_release_payload(PACK_ID)
        assert payload["product_id"] == PACK_ID
        assert payload["product_type"] == "pack"
        assert payload["price"] == 1234
        assert payload["currency"] == "jpy"
        assert payload["purchase_delivery_ready"] == "yes"
        assert payload["purchase_delivery_method"] == "manual_email_private_download"
        assert payload["distribution_file"] == f"{PACK_ID}.zip"
        assert payload["ready_for_live_execution"] == "yes"
        assert payload["errors"] == []
        assert payload["secret_read"] == "no"
        assert payload["live_api_called"] == "no"
        assert "dake_item_id" in payload["product_payload"]["metadata"]
        print("synthetic_pack_scaling=passed")
        return 0
    finally:
        if PACK_DIR.exists():
            shutil.rmtree(PACK_DIR)


if __name__ == "__main__":
    raise SystemExit(main())
