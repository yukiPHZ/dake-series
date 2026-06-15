# Product Stripe Release Dry Run

## Product

| field | value |
| --- | --- |
| product_id | DAKE_Pack_Document |
| product_type | pack |
| title | DAKE 書類整理パック |
| price | 1480 jpy |
| tax_code_candidate | txcd_10202003 |

## Source of Truth

- `04_packs/DAKE_Pack_Document/ORIGINAL.md`

## Delivery Readiness

| field | value |
| --- | --- |
| purchase_delivery_ready | yes |
| purchase_delivery_method | manual_email_private_download |
| distribution_file | DAKE_Pack_Document.zip |
| distribution_file_sha256 | 83b4be666bafd907de79884a698fce5e98101123d7fef5a942eaf6d0ea3f72b0 |

## Current Payment State

| field | value |
| --- | --- |
| payment_status_before | booth_only |
| stripe_payment_link_before | - |

## Stripe Product

| field | value |
| --- | --- |
| name | DAKE 書類整理パック |
| description | DAKE 書類整理パック。本商品は自動ダウンロードではありません。Stripe決済確認後、購入時のメールアドレス宛に次営業日以内にダウンロード方法をご案内します。 |
| tax_code | txcd_10202003 |

## Stripe Price

| field | value |
| --- | --- |
| currency | jpy |
| unit_amount | 1480 |
| product | __PRODUCT_ID_FROM_LIVE_PRODUCT__ |

## Stripe Payment Link

| field | value |
| --- | --- |
| price | __PRICE_ID_FROM_LIVE_PRICE__ |
| quantity | 1 |

## Metadata

| key | value |
| --- | --- |
| dake_item_id | DAKE_Pack_Document |
| dake_type | pack |
| source_repo | DAKE_series |
| source_original | 04_packs/DAKE_Pack_Document/ORIGINAL.md |
| store_url | https://store.dakeapp.com/product/?id=DAKE_Pack_Document |
| booth_url | https://peakheadz.booth.pm/items/8448353 |
| purchase_delivery_method | manual_email_private_download |
| delivery_policy | manual_email_private_download |

## Tax Code Candidate

`txcd_10202003` is a candidate. It must be confirmed before live execution.

## Payload Hashes

| field | value |
| --- | --- |
| product_payload_sha256 | b69082d8ad4f5695fd3d82f98b1021e009896d665950fa1c637ca08578789ff8 |
| price_payload_sha256 | 97623feb5254113ebdb5b4f1ff0a3412e84b08d4ecb3333b0dc1be8751acd8ed |
| payment_link_payload_sha256 | 98e3852b04492a90dd1c44764a7550e2330254aa8a81619f6761b47a171a4921 |

## Idempotency Keys

| field | value |
| --- | --- |
| product_idempotency_key | dake-release-product-v1-DAKE_Pack_Document-b69082d8ad4f |
| price_idempotency_key | dake-release-price-v1-DAKE_Pack_Document-97623feb5254 |
| payment_link_idempotency_key | dake-release-link-v1-DAKE_Pack_Document-98e3852b0449 |

## Safety Checks

- mode: dry-run
- ready_for_live_execution: yes
- secret_read: no
- live_api_called: no
- buyer_information_stored: no
- private_download_url_stored: no
- output_json: `tools/reports/release_artifacts/DAKE_Pack_Document/stripe_release_dry_run.json`

## Live Execution Readiness

yes

## Errors

| error |
| --- |
| - |

## Next Command

```powershell
python tools\release_product.py DAKE_Pack_Document --execute-live --confirm-product-id DAKE_Pack_Document --confirm-tax-code --confirmation-text "CREATE LIVE PAYMENT LINK DAKE_Pack_Document"
```
