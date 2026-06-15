# DAKE Mail Pack BOOTH and Stripe Dry-run Review

## BOOTH Registration

- product_id: DAKE_Pack_Mail
- product_name: DAKE メール準備パック
- booth_url: https://peakheadz.booth.pm/items/8457085
- registration_status: published
- display_price_jpy: 780
- registered_zip_name: DAKE_Pack_Mail.zip
- duplicate_url_check: passed. The URL appears only in DAKE_Pack_Mail files.

## Source of Truth Writeback

- source: `04_packs/DAKE_Pack_Mail/ORIGINAL.md`
- status: available
- payment_status: booth_only
- booth_url: https://peakheadz.booth.pm/items/8457085
- stripe_payment_link: 未設定
- price: 780円

The Pack title, price, included apps, ZIP name, ZIP size, ZIP SHA256, delivery method, product text, and notice text were kept unchanged except for BOOTH registration metadata.

## Pipeline Transition

- before BOOTH writeback: BOOTH_REGISTRATION_PENDING
- after BOOTH writeback: SOURCE_READY
- after Stripe dry-run: STRIPE_DRY_RUN_READY
- next_action: run Stripe live execution with explicit confirmation

Next live command shown by the pipeline:

```powershell
python tools\release_product.py DAKE_Pack_Mail --execute-live --confirm-product-id DAKE_Pack_Mail --confirm-tax-code --confirmation-text "CREATE LIVE PAYMENT LINK DAKE_Pack_Mail"
```

## Stripe Dry-run

- dry_run_json: `tools/reports/release_artifacts/DAKE_Pack_Mail/stripe_release_dry_run.json`
- mode: dry-run
- price: 780
- currency: jpy
- unit_amount: 780
- quantity: 1
- tax_code_candidate: txcd_10202003
- purchase_delivery_ready: yes
- purchase_delivery_method: manual_email_private_download
- ready_for_live_execution: yes
- errors: 0
- secret_read: no
- live_api_called: no

Payload hashes:

- product_payload_sha256: 45f9ea00d0952c117ac2fb022d31a911d0abb1f759f7011dc3c95b50a364701a
- price_payload_sha256: 1deff06049e71b55dac19ca0e6e578196b523d763941aa86d913e854311e83f3
- payment_link_payload_sha256: 0a2ebc3a48f545d5dea3833d85ee23f798489005a908116f4a2f2e18a8cd7faf

Planned idempotency keys:

- product_idempotency_key: dake-release-product-v1-DAKE_Pack_Mail-45f9ea00d095
- price_idempotency_key: dake-release-price-v1-DAKE_Pack_Mail-1deff06049e7
- payment_link_idempotency_key: dake-release-link-v1-DAKE_Pack_Mail-0a2ebc3a48f5

## Pack Integrity

- distribution_file: DAKE_Pack_Mail.zip
- distribution_path: `04_packs/DAKE_Pack_Mail/pack_ready/DAKE_Pack_Mail.zip`
- distribution_file_size: 56950902
- distribution_file_sha256: dfc972b91529161bbf688fbe4fb5bf91b5e27956afe058486a0b5d79ab293ad4
- included_apps: DAKE_Mail_List, DAKE_Mail_Address_Format, DAKE_Mail_Draft

## Existing Product Regression

- DAKE_Pack_Document: RELEASE_COMPLETE
- DAKE_Pack_Memo: RELEASE_COMPLETE
- dake_pdf_viewer: LEGACY_COMPLETE
- video_shorts_cut: PREPARING_BLOCKED

## Store Boundary

- generated JSON was not regenerated.
- dake-store-site was not synced.
- Store production was not changed.
- Store generated JSON remains total=53, stripe_ready=52, booth_only=0, preparing=1.
- DAKE_Pack_Mail is not listed in generated Store JSON.

## Security Review

- Stripe API was not called.
- Stripe Secret Key was not read.
- BOOTH API was not called.
- No live Product, Price, or Payment Link was created.
- No Checkout was opened.
- No buyer information is stored.
- No private download URL is stored.
- No BOOTH login, cookie, or session value is stored.
- No Pack ZIP absolute local path is exposed in public-facing artifacts.

## Next Live Command

```powershell
python tools\release_product.py DAKE_Pack_Mail --execute-live --confirm-product-id DAKE_Pack_Mail --confirm-tax-code --confirmation-text "CREATE LIVE PAYMENT LINK DAKE_Pack_Mail"
```

## Conclusion

READY_FOR_STRIPE_LIVE_EXECUTION
