from __future__ import annotations

import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "tools"))

from store.stripe_release_core import build_checkout_submit_message  # noqa: E402
from update_payment_link_checkout_notice import validate_live_payment_link, validate_notice  # noqa: E402


PRODUCT_NOTICE = (
    "Windows向けです。Dakeメール下書きはWindows版Microsoft Outlook Classicを使用します。"
    "New Outlook / Web Outlookでは動作しない場合があります。メールは自動送信されません。"
    "作成された下書きの宛先・件名・本文・添付を確認してから、利用者自身で送信してください。"
)


def test_notice_generation_with_product_specific_notice() -> None:
    original = f"""# ORIGINAL

- checkout_notice_required: yes
- checkout_product_notice: {PRODUCT_NOTICE}
"""
    message, errors = build_checkout_submit_message(original, "manual_email_private_download")
    assert errors == []
    assert len(message) < 1200
    for term in [
        "Windows",
        "Microsoft Outlook Classic",
        "New Outlook",
        "Web Outlook",
        "自動送信されません",
        "下書き",
        "確認してから",
        "自動ダウンロードではありません",
        "購入時に入力されたメールアドレス",
        "次営業日以内",
    ]:
        assert term in message
    assert validate_notice(message) == []


def test_notice_generation_without_product_specific_notice_keeps_manual_delivery_only() -> None:
    message, errors = build_checkout_submit_message("# ORIGINAL\n", "manual_email_private_download")
    assert errors == []
    assert "自動ダウンロードではありません" in message
    assert "次営業日以内" in message
    assert "Microsoft Outlook Classic" not in message


def test_notice_validation_rejects_missing_required_terms_and_too_long_text() -> None:
    assert validate_notice("自動ダウンロードではありません") != []
    long_notice = (
        PRODUCT_NOTICE
        + " 本商品は自動ダウンロードではありません。決済確認後、購入時に入力されたメールアドレス宛に、次営業日以内にダウンロード方法をご案内します。"
        + "x" * 1200
    )
    assert any("exceeds" in error for error in validate_notice(long_notice))


def test_payment_link_validation_accepts_expected_live_shape() -> None:
    plan = {
        "product_id": "DAKE_Pack_Mail",
        "payment_link_id": "plink_live",
        "payment_link_url": "https://buy.stripe.com/live",
        "price_id": "price_live",
    }
    link = {
        "id": "plink_live",
        "livemode": True,
        "active": True,
        "url": "https://buy.stripe.com/live",
        "metadata": {"dake_item_id": "DAKE_Pack_Mail"},
    }
    line_items = [{"price": {"id": "price_live"}, "quantity": 1}]
    assert validate_live_payment_link(link, line_items, plan) == []


def test_payment_link_validation_rejects_test_mode_and_mutated_line_item() -> None:
    plan = {
        "product_id": "DAKE_Pack_Mail",
        "payment_link_id": "plink_live",
        "payment_link_url": "https://buy.stripe.com/live",
        "price_id": "price_live",
    }
    link = {
        "id": "plink_live",
        "livemode": False,
        "active": True,
        "url": "https://buy.stripe.com/live",
        "metadata": {"dake_item_id": "DAKE_Pack_Mail"},
    }
    line_items = [{"price": {"id": "price_other"}, "quantity": 2}]
    errors = validate_live_payment_link(link, line_items, plan)
    assert "Payment Link livemode must be true" in errors
    assert "Payment Link Price ID mismatch" in errors
    assert "Payment Link quantity must be 1" in errors


def main() -> int:
    tests = [
        test_notice_generation_with_product_specific_notice,
        test_notice_generation_without_product_specific_notice_keeps_manual_delivery_only,
        test_notice_validation_rejects_missing_required_terms_and_too_long_text,
        test_payment_link_validation_accepts_expected_live_shape,
        test_payment_link_validation_rejects_test_mode_and_mutated_line_item,
    ]
    for test in tests:
        test()
        print(f"{test.__name__}=passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
