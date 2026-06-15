# Product Stripe Release Dry Run

## Product

| field | value |
| --- | --- |
| product_id | DAKE_Pack_Memo |
| product_type | pack |
| title | DAKE メモと記録パック |
| price | 980 jpy |
| tax_code_candidate | txcd_10202003 |

## Source of Truth

- `04_packs/DAKE_Pack_Memo/ORIGINAL.md`

## Delivery Readiness

| field | value |
| --- | --- |
| purchase_delivery_ready | yes |
| purchase_delivery_method | manual_email_private_download |
| distribution_file | DAKE_Pack_Memo.zip |
| distribution_file_sha256 | 62b2656fc2b2941bcb01c070d9338f1dde3ed24798c79fc8db24ca71b45f21bc |

## Current Payment State

| field | value |
| --- | --- |
| payment_status_before | booth_only |
| stripe_payment_link_before | - |

## Stripe Product

| field | value |
| --- | --- |
| name | DAKE メモと記録パック |
| description | DAKE メモと記録パック。本商品は自動ダウンロードではありません。Stripe決済確認後、購入時のメールアドレス宛に次営業日以内にダウンロード方法をご案内します。 |
| tax_code | txcd_10202003 |

## Stripe Price

| field | value |
| --- | --- |
| currency | jpy |
| unit_amount | 980 |
| product | __PRODUCT_ID_FROM_LIVE_PRODUCT__ |

## Stripe Payment Link

| field | value |
| --- | --- |
| price | __PRICE_ID_FROM_LIVE_PRICE__ |
| quantity | 1 |

## Metadata

| key | value |
| --- | --- |
| dake_item_id | DAKE_Pack_Memo |
| dake_type | pack |
| source_repo | DAKE_series |
| source_original | 04_packs/DAKE_Pack_Memo/ORIGINAL.md |
| store_url | https://store.dakeapp.com/product/?id=DAKE_Pack_Memo |
| booth_url | https://peakheadz.booth.pm/items/8449208 |
| purchase_delivery_method | manual_email_private_download |
| delivery_policy | manual_email_private_download |

## Tax Code Candidate

`txcd_10202003` is a candidate. It must be confirmed before live execution.

## Payload Hashes

| field | value |
| --- | --- |
| product_payload_sha256 | d89060ec22539f19b6ad05c35df49bc531567bf2d3cfdf7932f25762e40c5e42 |
| price_payload_sha256 | 83195943907b48ee45e6729d3054877792f9edbe106c41ee1428c7839f928097 |
| payment_link_payload_sha256 | f47fd9190716555147a7e3ad403fe254673b499dbf9c8bb61525a095e7299aef |

## Idempotency Keys

| field | value |
| --- | --- |
| product_idempotency_key | dake-release-product-v1-DAKE_Pack_Memo-d89060ec2253 |
| price_idempotency_key | dake-release-price-v1-DAKE_Pack_Memo-83195943907b |
| payment_link_idempotency_key | dake-release-link-v1-DAKE_Pack_Memo-f47fd9190716 |

## Safety Checks

- mode: dry-run
- ready_for_live_execution: yes
- secret_read: no
- live_api_called: no
- buyer_information_stored: no
- private_download_url_stored: no
- output_json: `tools/reports/release_artifacts/DAKE_Pack_Memo/stripe_release_dry_run.json`

## Live Execution Readiness

yes

## Errors

| error |
| --- |
| - |

## Next Command

```powershell
python tools\release_product.py DAKE_Pack_Memo --execute-live --confirm-product-id DAKE_Pack_Memo --confirm-tax-code --confirmation-text "CREATE LIVE PAYMENT LINK DAKE_Pack_Memo"
```
