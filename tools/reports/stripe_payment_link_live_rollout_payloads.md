# Stripe Payment Link Live Rollout Dry Run

## Purpose

Generate live-mode Stripe Product, Price, and Payment Link payloads for the 45 API creation candidates without calling Stripe.

## Safety Notice

DRY RUN ONLY. NO STRIPE API CALL. NO LIVE OBJECT IS CREATED.

- No Stripe Secret Key is read.
- No Product, Price, or Payment Link is created.
- No Stripe Product list is retrieved.
- No `ORIGINAL.md`, generated JSON, dake-store-site, or Store production update is performed.

## Inputs

- `tools/reports/stripe_payment_link_rollout_review.csv`
- `tools/generated/store_products.generated.json`
- `tools/reports/stripe_payment_link_pilot10_final_review.csv`

## Target Conditions

Rows must satisfy `review_result=create`, `creation_method=api_candidate`, `price_check=price_ok`, and `metadata_ready=yes`.

## Excluded

Pack products and preparing products are not included in this live dry-run.

| id | title | review_result | creation_method | reason |
| --- | --- | --- | --- | --- |
| video_shorts_cut | Dakeショート切り出し | hold | hold | not create/api_candidate/price_ok/metadata_ready |
| DAKE_Pack_Document | DAKE 書類整理パック | manual | manual_dashboard | not create/api_candidate/price_ok/metadata_ready |
| DAKE_Pack_Memo | DAKE メモと記録パック | manual | manual_dashboard | not create/api_candidate/price_ok/metadata_ready |

## Summary

- total candidates: 45
- payloads generated: 45
- live_dry_run_ready yes: 45
- live_dry_run_ready no: 0
- normal apps: 43
- game apps: 2
- price errors: 0
- metadata errors: 0
- tax candidate review required: 45
- payload hash missing: 0
- idempotency key duplicates: 0
- errors: 0

## Target 45

| id | title | price | tax_code | ready | notes |
| --- | --- | --- | --- | --- | --- |
| DAKE_App_Doko | アプリどこ | 300 | txcd_10202003 | yes | ok |
| DAKE_Backup | Dakeバックアップ | 500 | txcd_10202003 | yes | ok |
| DAKE_Git_Memo | DakeGitメモ | 500 | txcd_10202003 | yes | ok |
| DAKE_Image_PasteA4 | 貼る | 500 | txcd_10202003 | yes | ok |
| DAKE_Mail_Address_Format | Dakeメールアドレス整形 | 300 | txcd_10202003 | yes | ok |
| DAKE_Mail_Draft | Dakeメール下書き | 300 | txcd_10202003 | yes | ok |
| DAKE_Maji_Memo | マジでメモ | 300 | txcd_10202003 | yes | ok |
| DAKE_Sticky_Memo | 付箋メモ | 300 | txcd_10202003 | yes | ok |
| DAKE_Yesterday_Task_Memo | Dake昨日タスクメモ | 300 | txcd_10202003 | yes | ok |
| dake_booth_assist | BOOTHアシスト | 300 | txcd_10202003 | yes | ok |
| dake_column_memo | ずっとメモ | 300 | txcd_10202003 | yes | ok |
| dake_document_cover | Dake書類送付状 | 500 | txcd_10202003 | yes | ok |
| dake_fax_cover | DakeFAX送付状 | 500 | txcd_10202003 | yes | ok |
| dake_folder_list | Dakeフォルダ一覧 | 300 | txcd_10202003 | yes | ok |
| dake_image_heictojpg | HEIC→JPG変換 | 500 | txcd_10202003 | yes | ok |
| dake_image_iphonetopc | Dake画像iPhoneToPC | 500 | txcd_10202003 | yes | ok |
| dake_image_receiver | DakeImage_Receiver | 500 | txcd_10202003 | yes | ok |
| dake_image_topdf | DakeImageToPDF | 500 | txcd_10202003 | yes | ok |
| dake_launcher | Dakeランチャー | 300 | txcd_10202003 | yes | ok |
| dake_mail_allstaff | Dake全社員メール起動 | 300 | txcd_10202003 | yes | ok |
| dake_mail_kikuta | Dake菊田メール | 300 | txcd_10202003 | yes | ok |
| dake_mail_list | Dakeメールリスト | 300 | txcd_10202003 | yes | ok |
| dake_mansion_schedule | マンション工程表 | 500 | txcd_10202003 | yes | ok |
| dake_pdf_checkstamp | Dake確認印 | 500 | txcd_10202003 | yes | ok |
| dake_pdf_crop | DakePDFトリミング | 500 | txcd_10202003 | yes | ok |
| dake_pdf_lookhere | DakePDFここ見て | 500 | txcd_10202003 | yes | ok |
| dake_pdf_marker | DakePDFマーカー | 500 | txcd_10202003 | yes | ok |
| dake_pdf_merge_mini | DakePDF結合mini | 500 | txcd_10202003 | yes | ok |
| dake_pdf_rename | DakePDFファイル名整理 | 500 | txcd_10202003 | yes | ok |
| dake_pdf_reorder | DakePDFページ並べ替え | 500 | txcd_10202003 | yes | ok |
| dake_pdf_splitone | DakePDF分割One | 500 | txcd_10202003 | yes | ok |
| dake_pdf_splitselect | DakePDF分割Select | 500 | txcd_10202003 | yes | ok |
| dake_pdf_toimages | DakePDFto画像 | 500 | txcd_10202003 | yes | ok |
| dake_pdf_viewer | DakePDF見る | 500 | txcd_10202003 | yes | ok |
| dake_price_apportionment | Dake価格按分 | 300 | txcd_10202003 | yes | ok |
| dake_price_fixedtax | Dake固都税計算 | 300 | txcd_10202003 | yes | ok |
| dake_reform_progress | リフォーム進捗管理 | 500 | txcd_10202003 | yes | ok |
| dake_screen_webp | DakeScreen_WebP | 300 | txcd_10202003 | yes | ok |
| dake_screenshot_print | Dakeスクショ印刷 | 500 | txcd_10202003 | yes | ok |
| dake_two_person_memo | Dake二人メモ | 300 | txcd_10202003 | yes | ok |
| dake_work_calendar | Dake工程カレンダー | 500 | txcd_10202003 | yes | ok |
| dake_year_age | Dake築年数 | 300 | txcd_10202003 | yes | ok |
| dake_year_notice | Dake今年の注意点 | 300 | txcd_10202003 | yes | ok |
| game_alien_road | DakeAlien Road | 300 | txcd_10201000 | yes | ok |
| game_diver_catch | Dake潜って捕る | 300 | txcd_10201000 | yes | ok |

## Product Payload Policy

Products are built with `name`, a concise generated-store description, `active=true`, candidate `tax_code`, and full DAKE metadata.

## Price Payload Policy

Prices use `currency=jpy`, integer `unit_amount`, one-time billing by omitting recurring fields, and `metadata.dake_item_id`.

## Payment Link Payload Policy

Payment Links use one line item with quantity 1. Link metadata and payment intent metadata contain `dake_item_id`, `dake_type`, and `source_original`.

## Metadata Policy

Product metadata includes `dake_item_id`, `dake_type`, `source_repo`, `source_original`, `store_url`, `booth_url`, and `github_release_url`. No personal information or secret values are included.

## Payload Hash

Each Product, Price, and Payment Link payload is canonicalized with sorted JSON keys and hashed with SHA-256.

## Idempotency Keys

Idempotency keys are generated from the safe item id and the first 12 characters of each payload hash. These keys are not secrets and are not sent to Stripe in this phase.

## Duplicate Avoidance Policy

Next phase should retrieve live Products, match by `metadata.dake_item_id`, create only when there is no match, reuse or stop when one match exists, and stop on multiple matches.

## Tax Code Candidate

Tax codes come from `stripe_payment_link_rollout_review.csv`. `tax_candidate_review=operator_confirmation_required` for every item; this dry-run does not make a final tax determination.

## Stop Conditions Before Live Execution

- Target count is not 45.
- Any `live_dry_run_ready=no`.
- Price, currency, metadata, source, BOOTH URL, GitHub Release URL, payload hash, or idempotency key errors exist.
- Duplicate idempotency keys or duplicate DAKE item ids exist.
- Any secret-like value appears in output files.

## Next Phase Proposal

Use this dry-run as the review artifact for a separate live-mode execution script with explicit operator approval and duplicate checks.

## Not Done In This Phase

- No Stripe API call.
- No Stripe Secret Key read.
- No `sk_test_` or `sk_live_` use.
- No Product, Price, or Payment Link creation.
- No Payment Link URL write-back.
- No Store production update.

## Errors

| error |
| --- |
| - |
