# Stripe Payment Link Pilot10 Payloads

## Purpose

Build dry-run payloads for Stripe Product / Price / Payment Link creation for the Phase 13A pilot 10 items.

## Dry-run Notice

DRY RUN ONLY. This file does not call Stripe API.

No Stripe API call is made. No Stripe Secret Key is read. No Payment Link is created. No Product is created. No Price is created.

## Input

- `tools/reports/stripe_payment_link_pilot10_selection.csv`
- `tools/generated/store_products.generated.json`

## Output

- `tools/reports/stripe_payment_link_pilot10_payloads.json`
- `tools/reports/stripe_payment_link_pilot10_payloads.md`

## Count

- items: 10
- errors: 0

## Items

| id | title | price | currency | tax_code | metadata |
| --- | --- | --- | --- | --- | --- |
| dake_pdf_viewer | DakePDF見る | 500 | jpy | txcd_10202003 | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url |
| dake_pdf_reorder | DakePDFページ並べ替え | 500 | jpy | txcd_10202003 | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url |
| dake_pdf_splitone | DakePDF分割One | 500 | jpy | txcd_10202003 | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url |
| dake_image_heictojpg | HEIC→JPG変換 | 500 | jpy | txcd_10202003 | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url |
| dake_image_topdf | DakeImageToPDF | 500 | jpy | txcd_10202003 | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url |
| DAKE_Sticky_Memo | 付箋メモ | 300 | jpy | txcd_10202003 | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url |
| DAKE_Mail_Draft | Dakeメール下書き | 300 | jpy | txcd_10202003 | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url |
| DAKE_Backup | Dakeバックアップ | 500 | jpy | txcd_10202003 | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url |
| dake_folder_list | Dakeフォルダ一覧 | 300 | jpy | txcd_10202003 | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url |
| dake_year_age | Dake築年数 | 300 | jpy | txcd_10202003 | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url |

## Errors

| error |
| --- |
| - |

## Safety Notes

- No Stripe API call is made.
- No Stripe Secret Key is read.
- No Payment Link is created.
- No Product is created.
- No Price is created.
- Metadata must not contain personal information, card information, buyer information, Stripe Secret Key, Webhook Secret, or internal tokens.
- `tax_code` is a candidate and should be reviewed before live execution.
- Stripe Payment Links are Stripe-hosted checkout URLs. The payload here only describes what would be sent in a later phase.
- Metadata stores structured key-value information on Stripe objects. Payment Link `payment_intent_data.metadata` is included so generated Payment Intents can carry DAKE identifiers.

## Next Phase

Use this JSON as the reviewed input for a test mode implementation phase. That next phase should still default to dry-run, require explicit human approval before any execution path, and keep Stripe Secret Key in environment variables only.
