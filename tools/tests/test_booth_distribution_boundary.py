from __future__ import annotations

import sys
import tempfile
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "tools"))

import check_booth_ready  # noqa: E402
import make_booth_ready  # noqa: E402
import record_booth_registration  # noqa: E402
import release_product_pipeline  # noqa: E402


def write_original(app_dir: Path) -> None:
    app_dir.mkdir(parents=True, exist_ok=True)
    (app_dir / "ORIGINAL.md").write_text(
        """# ORIGINAL.md

## DAKE_META生成用情報

```json
{
  "app_key": "synthetic_app",
  "folder_name": "DAKE_Synthetic",
  "display_name": "Synthetic DAKE",
  "status": "available",
  "exe_name": "Synthetic.exe"
}
```
""",
        encoding="utf-8",
    )


def write_booth_ready(app_dir: Path, *, include_zip: bool = True) -> None:
    ready = app_dir / "booth_ready"
    ready.mkdir(parents=True, exist_ok=True)
    (ready / "booth_product.txt").write_text(
        "# 商品名\nSynthetic DAKE\n\n# URL\n\n",
        encoding="utf-8",
    )
    (ready / "booth_thumbnail.jpg").write_bytes(b"synthetic thumbnail")
    (ready / "README.txt").write_text("Synthetic instructions.\n", encoding="utf-8")
    if include_zip:
        with zipfile.ZipFile(ready / "Synthetic.zip", "w") as archive:
            archive.writestr("Synthetic.exe", b"synthetic exe fixture")


def test_booth_prep_ignores_ship_social_and_commerce() -> None:
    with tempfile.TemporaryDirectory() as temp:
        root = Path(temp) / "DAKE_series"
        app_dir = root / "01_apps" / "DAKE_Synthetic"
        write_original(app_dir)
        write_booth_ready(app_dir)
        old_root = check_booth_ready.ROOT
        try:
            check_booth_ready.ROOT = root
            result = check_booth_ready.check_app(app_dir)
        finally:
            check_booth_ready.ROOT = old_root
        assert result.booth_prep_ready is True
        assert result.booth_url_state == "empty"
        assert not any("SNS" in action or "Stripe" in action or "Release" in action for action in result.actions)


def test_booth_prep_requires_registration_materials() -> None:
    with tempfile.TemporaryDirectory() as temp:
        root = Path(temp) / "DAKE_series"
        app_dir = root / "01_apps" / "DAKE_Synthetic"
        write_original(app_dir)
        write_booth_ready(app_dir, include_zip=False)
        old_root = check_booth_ready.ROOT
        try:
            check_booth_ready.ROOT = root
            result = check_booth_ready.check_app(app_dir)
        finally:
            check_booth_ready.ROOT = old_root
        assert result.booth_prep_ready is False
        assert any("ZIP" in action for action in result.actions)


def test_booth_generator_uses_original_and_existing_ready_assets() -> None:
    with tempfile.TemporaryDirectory() as temp:
        root = Path(temp) / "DAKE_series"
        app_dir = root / "01_apps" / "DAKE_Synthetic"
        write_original(app_dir)
        write_booth_ready(app_dir)
        old_root = make_booth_ready.ROOT
        try:
            make_booth_ready.ROOT = root
            result = make_booth_ready.process_app(app_dir, dry_run=True)
        finally:
            make_booth_ready.ROOT = old_root
        assert result.ok is True
        assert result.price_source == "inferred"
        assert any("existing verified booth_ready ZIP" in warning for warning in result.warnings)


def test_recording_booth_url_changes_only_booth_url() -> None:
    before = """# ORIGINAL.md

- booth_url: 未設定
- status: draft
- payment_status: stripe_ready
- stripe_payment_link: https://buy.stripe.com/example
- distribution: Store only
"""
    after = record_booth_registration.build_updated_original(
        before,
        "https://peakheadz.booth.pm/items/1234567",
    )
    assert "booth_url: https://peakheadz.booth.pm/items/1234567" in after
    assert "status: draft" in after
    assert "payment_status: stripe_ready" in after
    assert "stripe_payment_link: https://buy.stripe.com/example" in after
    assert "distribution: Store only" in after


def test_policy_and_tools_stop_before_booth_registration() -> None:
    spec = (ROOT / "00_core" / "DAKE_BOOTH_REGISTER_SPEC.md").read_text(encoding="utf-8")
    maker = (ROOT / "tools" / "make_booth_ready.py").read_text(encoding="utf-8")
    checker = (ROOT / "tools" / "check_booth_ready.py").read_text(encoding="utf-8")
    pipeline = (ROOT / "tools" / "release_product_pipeline.py").read_text(encoding="utf-8")
    assert "Codexは売る準備まで。BOOTH登録は人間。" in spec
    assert "BOOTH_PREP = READY" in spec
    assert "formal shipping includes" not in maker
    assert "social_buffer_ready" not in checker
    assert "release_capture_ready" not in checker
    assert "--allow-legacy-commerce-advance" in pipeline


def test_legacy_pipeline_advance_is_gated() -> None:
    class DummyPipeline:
        def status(self, product_id: str) -> dict[str, object]:
            return {
                "product_id": product_id,
                "current_stage": "SOURCE_READY",
                "next_action": "legacy action",
                "next_formal_action": None,
            }

        def advance(self, product_id: str) -> dict[str, object]:
            raise AssertionError(f"advance must stay disabled for {product_id}")

    old_argv = sys.argv[:]
    old_pipeline = release_product_pipeline.ReleasePipeline
    try:
        sys.argv = ["release_product_pipeline.py", "Synthetic_Product", "advance"]
        release_product_pipeline.ReleasePipeline = DummyPipeline  # type: ignore[assignment]
        assert release_product_pipeline.main() == 1
    finally:
        sys.argv = old_argv
        release_product_pipeline.ReleasePipeline = old_pipeline


def main() -> int:
    tests = [
        test_booth_prep_ignores_ship_social_and_commerce,
        test_booth_prep_requires_registration_materials,
        test_booth_generator_uses_original_and_existing_ready_assets,
        test_recording_booth_url_changes_only_booth_url,
        test_policy_and_tools_stop_before_booth_registration,
        test_legacy_pipeline_advance_is_gated,
    ]
    for test in tests:
        test()
        print(f"{test.__name__}=passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
