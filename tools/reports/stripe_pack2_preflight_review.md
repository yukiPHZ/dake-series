# Stripe Pack2 Preflight Review

## Purpose

Review the two DAKE Pack products before Stripe Payment Link registration.

## Safety Scope

- No Stripe API call is made.
- No Stripe Secret Key is read.
- No Product, Price, or Payment Link is created, updated, or deleted.
- ORIGINAL.md, generated JSON, dake-store-site, BOOTH, and Pack ZIP files are not updated.

## Inputs

- `tools/generated/store_products.generated.json`
- `tools/reports/stripe_payment_link_rollout_review.csv`
- `04_packs/DAKE_Pack_Document/pack_manifest.json`
- `04_packs/DAKE_Pack_Memo/pack_manifest.json`

## Outputs

- `tools/reports/stripe_pack2_preflight_review.csv`
- `tools/reports/stripe_pack2_preflight_review.md`

## Summary

- reviewed_packs: 2
- ready: 0
- conditional: 0
- hold: 2
- generated_at: 2026-06-15T17:55:29

## Pack Review

| id | title | price | booth | distribution | delivery_ready | review |
|---|---|---:|---|---|---|---|
| DAKE_Pack_Document | DAKE 書類整理パック | 1480 JPY | yes | DAKE_Pack_Document.zip | no | hold |
| DAKE_Pack_Memo | DAKE メモと記録パック | 980 JPY | yes | DAKE_Pack_Memo.zip | no | hold |

## Details

### DAKE_Pack_Document

- title: DAKE 書類整理パック
- price: 1480 JPY
- source_original: `04_packs/DAKE_Pack_Document/ORIGINAL.md`
- booth_url: https://peakheadz.booth.pm/items/8448353
- github_release_url: (empty)
- distribution_file: `DAKE_Pack_Document.zip`
- distribution_path: `04_packs/DAKE_Pack_Document/pack_ready/DAKE_Pack_Document.zip`
- purchase_delivery_method: BOOTH delivery exists; Stripe post-payment fulfillment not confirmed
- purchase_delivery_ready: no
- tax_code_candidate: txcd_10202003
- stripe_creation_method: hold
- review_result: hold
- notes: pack_zip_exists=True; pack_zip_size_ok=True; pack_zip_sha256_ok=True; pack_zip_git_tracked=False; included_apps=DAKE_PDF_Merge, DAKE_Image_ToPDF, DAKE_Image_Resize, DAKE_Image_PasteA4; Stripe post-payment delivery route is not confirmed in ORIGINAL.md; Define manual fulfillment or a private download route before live Stripe registration

### DAKE_Pack_Memo

- title: DAKE メモと記録パック
- price: 980 JPY
- source_original: `04_packs/DAKE_Pack_Memo/ORIGINAL.md`
- booth_url: https://peakheadz.booth.pm/items/8449208
- github_release_url: (empty)
- distribution_file: `DAKE_Pack_Memo.zip`
- distribution_path: `04_packs/DAKE_Pack_Memo/pack_ready/DAKE_Pack_Memo.zip`
- purchase_delivery_method: BOOTH delivery exists; Stripe post-payment fulfillment not confirmed
- purchase_delivery_ready: no
- tax_code_candidate: txcd_10202003
- stripe_creation_method: hold
- review_result: hold
- notes: pack_zip_exists=True; pack_zip_size_ok=True; pack_zip_sha256_ok=True; pack_zip_git_tracked=False; included_apps=DAKE_Sticky_Memo, DAKE_Maji_Memo, DAKE_Git_Memo, DAKE_Yesterday_Task_Memo; Stripe post-payment delivery route is not confirmed in ORIGINAL.md; Define manual fulfillment or a private download route before live Stripe registration

## Pre-live Required Actions

- Define the post-Stripe purchase delivery route for each Pack.
- If fulfillment is manual, document the operator workflow and purchase message before live registration.
- If fulfillment uses a private download URL or GitHub Release asset, add the confirmed route to the source of truth before Store reflection.
- Review the tax code candidate before live execution. Current candidate: `txcd_10202003`.

## Recommendation

Both Pack ZIPs and BOOTH routes exist, but Stripe post-payment fulfillment is not confirmed in the source of truth. Keep both Packs on hold for live Stripe registration until the delivery route is defined.

## Next Phase

After the delivery route is confirmed, create manual Dashboard Payment Links or a dedicated Pack payload flow, then write confirmed Payment Link URLs back to ORIGINAL.md.
