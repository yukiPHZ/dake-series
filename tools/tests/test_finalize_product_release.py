from tools.finalize_product_release import parse_original, validate_checkout_review


def _base_review() -> dict[str, object]:
    return {
        "product_id": "example_app",
        "review_status": "passed",
        "livemode": True,
        "actual_payment_completed": False,
        "email_field_present": True,
        "email_required": True,
        "quantity_change_ui": False,
        "shipping_address_required": False,
        "test_url_detected": False,
        "private_download_url_exposed": False,
        "local_path_exposed": False,
        "console_errors": 0,
        "manual_delivery_notice_visible": "not_applicable",
        "next_business_day_notice_visible": "not_applicable",
    }


def test_app_checkout_does_not_require_manual_pack_delivery_notices() -> None:
    review = _base_review()
    assert validate_checkout_review("example_app", "app", review) == []


def test_pack_checkout_still_requires_manual_delivery_notices() -> None:
    review = _base_review()
    review["product_id"] = "example_pack"
    errors = validate_checkout_review("example_pack", "pack", review)
    assert "checkout_review manual_delivery_notice_visible must be True" in errors
    assert "checkout_review next_business_day_notice_visible must be True" in errors


def test_app_original_uses_declared_app_id(tmp_path) -> None:
    app_dir = tmp_path / "01_apps" / "FolderName"
    app_dir.mkdir(parents=True)
    original = app_dir / "ORIGINAL.md"
    original.write_text("- app_id: declared_app_id\n", encoding="utf-8")
    parsed = parse_original(original)
    assert parsed["id"] == "declared_app_id"
    assert parsed["product_type"] == "app"
