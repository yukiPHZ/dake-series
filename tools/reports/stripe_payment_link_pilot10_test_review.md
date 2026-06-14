# Stripe Payment Link Pilot10 Test Review

## Purpose

Review the 10 Stripe test mode Product, Price, and Payment Link objects created in Phase 13C using read-only Stripe API retrieval.

## Safety Conditions

- Stripe Secret Key is read only from `STRIPE_SECRET_KEY`.
- `sk_live_` keys are rejected.
- This script calls only retrieve APIs for Product, Price, and Payment Link.
- No Stripe object is created, updated, disabled, or deleted.
- `ORIGINAL.md`, generated JSON, dake-store-site, and Store production are not updated.
- tax_code is checked only as a configured candidate; tax business review remains pending.

## Target

- review count: 10
- local result mode: test
- expected items: 10

## Local Structure Check

| local_error |
| --- |
| - |

## Stripe Test Mode Retrieval Summary

- Product retrieved: 10
- Price retrieved: 10
- Payment Link retrieved: 10
- price ok: 10
- one_time ok: 10
- metadata ok: 10
- browser ok: 0
- browser manual_review: 10
- technical_ready yes: 10
- live_ready conditional: 0
- live_ready no: 10

## Item Review

| id | expected title | actual product name | expected price | actual amount | currency | metadata | tax_code | browser | technical | live_ready | notes |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| dake_pdf_viewer | DakePDF見る | DakePDF見る | 500 | 500 | jpy | ok | actual_matches_candidate | manual_review | yes | no | browser checkout page is not verified by this script |
| dake_pdf_reorder | DakePDFページ並べ替え | DakePDFページ並べ替え | 500 | 500 | jpy | ok | actual_matches_candidate | manual_review | yes | no | browser checkout page is not verified by this script |
| dake_pdf_splitone | DakePDF分割One | DakePDF分割One | 500 | 500 | jpy | ok | actual_matches_candidate | manual_review | yes | no | browser checkout page is not verified by this script |
| dake_image_heictojpg | HEIC→JPG変換 | HEIC→JPG変換 | 500 | 500 | jpy | ok | actual_matches_candidate | manual_review | yes | no | browser checkout page is not verified by this script |
| dake_image_topdf | DakeImageToPDF | DakeImageToPDF | 500 | 500 | jpy | ok | actual_matches_candidate | manual_review | yes | no | browser checkout page is not verified by this script |
| DAKE_Sticky_Memo | 付箋メモ | 付箋メモ | 300 | 300 | jpy | ok | actual_matches_candidate | manual_review | yes | no | browser checkout page is not verified by this script |
| DAKE_Mail_Draft | Dakeメール下書き | Dakeメール下書き | 300 | 300 | jpy | ok | actual_matches_candidate | manual_review | yes | no | browser checkout page is not verified by this script |
| DAKE_Backup | Dakeバックアップ | Dakeバックアップ | 500 | 500 | jpy | ok | actual_matches_candidate | manual_review | yes | no | browser checkout page is not verified by this script |
| dake_folder_list | Dakeフォルダ一覧 | Dakeフォルダ一覧 | 300 | 300 | jpy | ok | actual_matches_candidate | manual_review | yes | no | browser checkout page is not verified by this script |
| dake_year_age | Dake築年数 | Dake築年数 | 300 | 300 | jpy | ok | actual_matches_candidate | manual_review | yes | no | browser checkout page is not verified by this script |

## Product Name And Description Review

Product names are compared with the pilot selection Stripe product name. Descriptions must be non-empty, reasonably short, and free of secret-like tokens.

## Price And Billing Review

Each Price must be active, `currency=jpy`, `type=one_time`, `recurring=null`, and linked to the expected Product.

## Payment Link Review

Each Payment Link must be active, test mode only, and its URL must match the Phase 13C local result JSON.

## Metadata Review

Product and Payment Link metadata must match the full DAKE metadata. Price metadata is checked against the Phase 13B Price payload, which contains `dake_item_id`.

## Tax Code Review

Expected candidate: `txcd_10202003`. This is not a final tax determination. `tax_business_review` remains `pending`.

## Browser Manual Review

The script does not open checkout pages. Rows remain `manual_review` unless a separate browser pass confirms the pages.

## Secret Leak Check

The secret value is not written to this report. Use the repository regex check to confirm no key-like token appears in report files.

## Corrections Before Live

Rows with `technical_ready=no`, `browser_check!=ok`, or `tax_code_check` not matching the candidate must be reviewed before any live-mode rollout.

## Judgment

`live_ready` is `conditional` only when technical checks pass, browser check is `ok`, and the candidate tax code matches. This phase never marks a row as final live approval.

## Next Phase Proposal

After human browser review and tax review, decide whether to write selected test Payment Link URLs back to the source planning files or proceed to a live-mode dry-run plan.

## Not Done In This Phase

- No Stripe object creation.
- No Stripe object update.
- No Stripe object deletion.
- No live mode API call.
- No `sk_live_` use.
- No Store production update.
- No `ORIGINAL.md` or generated JSON update.
