# Product Stripe Release Dry Run

## Product

| field | value |
| --- | --- |
| product_id | dake_pdf_overview_rename |
| product_type | app |
| title | DakePDF俯瞰名前変更 |
| price | 500 jpy |
| tax_code_candidate | txcd_10202003 |

## Source of Truth

- `01_apps/DAKE_PDF_OverviewRename/ORIGINAL.md`

## Delivery Readiness

| field | value |
| --- | --- |
| purchase_delivery_ready | not_applicable |
| purchase_delivery_method | - |
| checkout_notice_required | not_applicable |
| checkout_submit_message_length | 0 |
| distribution_file | - |
| distribution_file_sha256 | - |

## Current Payment State

| field | value |
| --- | --- |
| payment_status_before | booth_only |
| stripe_payment_link_before | 未設定 |

## Stripe Product

| field | value |
| --- | --- |
| name | DakePDF俯瞰名前変更 |
| description | フォルダ内のPDFを1ページ目のサムネイルで俯瞰しながら、PDFごとに新しい名前を入力し、変更分だけまとめて反映するWindows向けアプリです。 |
| tax_code | txcd_10202003 |

## Stripe Price

| field | value |
| --- | --- |
| currency | jpy |
| unit_amount | 500 |
| product | __PRODUCT_ID_FROM_LIVE_PRODUCT__ |

## Stripe Payment Link

| field | value |
| --- | --- |
| price | __PRICE_ID_FROM_LIVE_PRICE__ |
| quantity | 1 |

## Checkout Submit Notice

-

## Metadata

| key | value |
| --- | --- |
| dake_item_id | dake_pdf_overview_rename |
| dake_type | app |
| source_repo | DAKE_series |
| source_original | 01_apps/DAKE_PDF_OverviewRename/ORIGINAL.md |
| store_url | https://store.dakeapp.com/product/?id=dake_pdf_overview_rename |
| booth_url | https://peakheadz.booth.pm/items/8798555 |
| github_release_url | https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_PDF_OverviewRename_v1.0.0 |

## Tax Code Candidate

`txcd_10202003` is a candidate. It must be confirmed before live execution.

## Payload Hashes

| field | value |
| --- | --- |
| product_payload_sha256 | e5b0b9ed08c8e471cea4468e960ceb1ccae9f173eb1948c174406b62da95847f |
| price_payload_sha256 | 178344fe7c825e9d9b18af6fb0b7dab1a87d6226dd2989d23ecaa3066fcda500 |
| payment_link_payload_sha256 | f4aade9ca5d549448c2d0b3829d3684a5c10bf95e2e2ac6e324b9c5b9e2a2361 |

## Idempotency Keys

| field | value |
| --- | --- |
| product_idempotency_key | dake-release-product-v1-dake_pdf_overview_rename-e5b0b9ed08c8 |
| price_idempotency_key | dake-release-price-v1-dake_pdf_overview_rename-178344fe7c82 |
| payment_link_idempotency_key | dake-release-link-v1-dake_pdf_overview_rename-f4aade9ca5d5 |

## Safety Checks

- mode: dry-run
- ready_for_live_execution: yes
- secret_read: no
- live_api_called: no
- buyer_information_stored: no
- private_download_url_stored: no
- output_json: `tools/reports/release_artifacts/dake_pdf_overview_rename/stripe_release_dry_run.json`

## Live Execution Readiness

yes

## Errors

| error |
| --- |
| - |

## Next Command

```powershell
python tools\release_product.py dake_pdf_overview_rename --execute-live --confirm-product-id dake_pdf_overview_rename --confirm-tax-code --confirmation-text "CREATE LIVE PAYMENT LINK dake_pdf_overview_rename"
```
