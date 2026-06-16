# Product Stripe Release Dry Run

## Product

| field | value |
| --- | --- |
| product_id | DAKE_Pack_Mail |
| product_type | pack |
| title | DAKE メール準備パック |
| price | 780 jpy |
| tax_code_candidate | txcd_10202003 |

## Source of Truth

- `04_packs/DAKE_Pack_Mail/ORIGINAL.md`

## Delivery Readiness

| field | value |
| --- | --- |
| purchase_delivery_ready | yes |
| purchase_delivery_method | manual_email_private_download |
| checkout_notice_required | True |
| checkout_submit_message_length | 250 |
| distribution_file | DAKE_Pack_Mail.zip |
| distribution_file_sha256 | dfc972b91529161bbf688fbe4fb5bf91b5e27956afe058486a0b5d79ab293ad4 |

## Current Payment State

| field | value |
| --- | --- |
| payment_status_before | booth_only |
| stripe_payment_link_before | - |

## Stripe Product

| field | value |
| --- | --- |
| name | DAKE メール準備パック |
| description | DAKE メール準備パック。本商品は自動ダウンロードではありません。Stripe決済確認後、購入時のメールアドレス宛に次営業日以内にダウンロード方法をご案内します。 |
| tax_code | txcd_10202003 |

## Stripe Price

| field | value |
| --- | --- |
| currency | jpy |
| unit_amount | 780 |
| product | __PRODUCT_ID_FROM_LIVE_PRODUCT__ |

## Stripe Payment Link

| field | value |
| --- | --- |
| price | __PRICE_ID_FROM_LIVE_PRICE__ |
| quantity | 1 |

## Checkout Submit Notice

【動作環境・重要事項】
Windows向けです。Dakeメール下書きはWindows版Microsoft Outlook Classicを使用します。New Outlook / Web Outlookでは動作しない場合があります。メールは自動送信されません。作成された下書きの宛先・件名・本文・添付を確認してから、利用者自身で送信してください。

【お渡し方法】
本商品は自動ダウンロードではありません。決済確認後、購入時に入力されたメールアドレス宛に、次営業日以内にダウンロード方法をご案内します。

## Metadata

| key | value |
| --- | --- |
| dake_item_id | DAKE_Pack_Mail |
| dake_type | pack |
| source_repo | DAKE_series |
| source_original | 04_packs/DAKE_Pack_Mail/ORIGINAL.md |
| store_url | https://store.dakeapp.com/product/?id=DAKE_Pack_Mail |
| booth_url | https://peakheadz.booth.pm/items/8457085 |
| purchase_delivery_method | manual_email_private_download |
| delivery_policy | manual_email_private_download |

## Tax Code Candidate

`txcd_10202003` is a candidate. It must be confirmed before live execution.

## Payload Hashes

| field | value |
| --- | --- |
| product_payload_sha256 | 45f9ea00d0952c117ac2fb022d31a911d0abb1f759f7011dc3c95b50a364701a |
| price_payload_sha256 | 1deff06049e71b55dac19ca0e6e578196b523d763941aa86d913e854311e83f3 |
| payment_link_payload_sha256 | bcb94817368da233252cd07973666b6b91635446c86a47f7fd5e45792f0d648c |

## Idempotency Keys

| field | value |
| --- | --- |
| product_idempotency_key | dake-release-product-v1-DAKE_Pack_Mail-45f9ea00d095 |
| price_idempotency_key | dake-release-price-v1-DAKE_Pack_Mail-1deff06049e7 |
| payment_link_idempotency_key | dake-release-link-v1-DAKE_Pack_Mail-bcb94817368d |

## Safety Checks

- mode: dry-run
- ready_for_live_execution: yes
- secret_read: no
- live_api_called: no
- buyer_information_stored: no
- private_download_url_stored: no
- output_json: `tools/reports/release_artifacts/DAKE_Pack_Mail/stripe_release_dry_run.json`

## Live Execution Readiness

yes

## Errors

| error |
| --- |
| - |

## Next Command

```powershell
python tools\release_product.py DAKE_Pack_Mail --execute-live --confirm-product-id DAKE_Pack_Mail --confirm-tax-code --confirmation-text "CREATE LIVE PAYMENT LINK DAKE_Pack_Mail"
```
