# Release Product Live Execution Self Review

## 目的

`python tools\release_product.py <product_id>` を、初回Stripe本番実行前にセルフレビューした。

対象は直接には以下の2 Pack。

- DAKE_Pack_Document
- DAKE_Pack_Memo

Stripe API、Stripe Secret、Product / Price / Payment Link作成、Store同期、generated JSON再生成、Payment Link書き戻しは行っていない。

## 対象

- `tools/release_product.py`
- `tools/store/stripe_release_core.py`
- `tools/store/create_stripe_live_rollout.py`
- `tools/store/apply_stripe_live_results_to_originals.py`
- `04_packs/DAKE_Pack_Document/ORIGINAL.md`
- `04_packs/DAKE_Pack_Memo/ORIGINAL.md`
- `00_core/DAKE_PACK_MANUAL_DELIVERY_RULE.md`
- Pack dry-run artifacts under `tools/reports/release_artifacts/`

## 結論

READY_FOR_PACK_LIVE_EXECUTION

## Blocking Findings

0

## Fixed Findings

1. Source discovery no longer depends on generated JSON.
   `release_product.py` now resolves products by scanning `01_apps/**/ORIGINAL.md` and `04_packs/**/ORIGINAL.md` first. Generated JSON is only supplemental for aliases and existing Store-state cross-checks.

2. Pack readiness no longer depends on the Phase 16B CSV report.
   Pack gate now checks each target `ORIGINAL.md`, `pack_manifest.json`, and Pack ZIP directly.

3. Atomic dry-run writes no longer share a fixed `.tmp` file name.
   Temporary output files include the process id, avoiding same-product concurrent dry-run collisions.

## Accepted Risks

- The Payment Link payload does not include a private download URL or buyer data. That is intentional.
- Buyer email collection must be confirmed in the Stripe-hosted Checkout page immediately after the first live Payment Link is created and before public use.
- `tax_code_candidate=txcd_10202003` remains a candidate, not a tax determination. Live execution requires `--confirm-tax-code`.

## Source of Truth Discovery

The release command scans `ORIGINAL.md` directly.

- `01_apps/**/ORIGINAL.md`
- `04_packs/**/ORIGINAL.md`

Generated JSON is used only as:

- alias support for existing Store IDs such as `dake_pdf_viewer`
- existing payment-state cross-check where present
- fallback for missing display fields

If generated JSON is missing or stale, a new Pack can still be discovered from `ORIGINAL.md`.

Duplicate product IDs in scanned `ORIGINAL.md` files stop execution.

## Future Pack Scalability

Verified with a temporary synthetic Pack:

- `Synthetic_Pack_Example`
- not added to generated JSON
- temporary `ORIGINAL.md`
- temporary Pack ZIP
- temporary `pack_manifest.json`

Result: `synthetic_pack_scaling=passed`.

The temporary Pack directory is deleted by the test.

## Hardcoded Product Review

No live logic is hardcoded to `DAKE_Pack_Document`, `DAKE_Pack_Memo`, `1480`, or `980`.

Allowed occurrence:

- CLI help example text
- generated dry-run artifacts
- review/report text

Pack name, price, ZIP path, SHA256, and delivery method are read from `ORIGINAL.md`, `pack_manifest.json`, or generated artifacts for the target product.

## Pack Delivery Gate

Each Pack must satisfy:

- `product_type=pack`
- price exists
- currency is `jpy`
- payment status is `booth_only`
- no existing `stripe_payment_link`
- BOOTH URL exists
- Pack ZIP exists
- Pack ZIP size matches manifest
- Pack ZIP SHA256 matches manifest
- `purchase_delivery_ready=yes`
- `purchase_delivery_method=manual_email_private_download`
- delivery window exists
- buyer notice exists
- resend/failure policy exists
- common manual delivery rule reference exists

## Checkout Email Requirement

Manual Pack delivery depends on buyer email availability in Stripe Dashboard.

The payload does not store or hardcode buyer email. Before public use of the first live Pack Payment Link, the operator must confirm in the Stripe-hosted Checkout page and Dashboard that the buyer email is collected and visible for fulfillment.

## Stripe SDK Version

Installed SDK:

- stripe-python: 15.2.1

The implementation uses the legacy resource style consistently:

- `stripe.Product.create(..., idempotency_key=...)`
- `stripe.Price.create(..., idempotency_key=...)`
- `stripe.PaymentLink.create(..., idempotency_key=...)`

No mixed `StripeClient` path is introduced.

## Idempotency Request Options

Idempotency keys are not inserted into Stripe payload dictionaries.

They are passed as request options/keyword arguments to the SDK call.

Keys are deterministic:

- same product + same payload -> same key
- different payload -> different key
- product / price / link use separate keys
- no timestamp or random value

## Payload Hashes

Hashes are recalculated from canonical JSON:

- key sorted
- UTF-8
- compact separators

Generated fields:

- `product_payload_sha256`
- `price_payload_sha256`
- `payment_link_payload_sha256`

The live path refuses execution if hashes do not match.

## Live Guard Order

Local checks happen before `STRIPE_SECRET_KEY` is read.

Guard order:

1. Parse CLI arguments.
2. Validate product id.
3. Discover `ORIGINAL.md`.
4. Validate payment state.
5. Validate Pack delivery gate.
6. Validate ZIP and SHA256.
7. Build payload.
8. Recalculate payload hashes.
9. Confirm product id.
10. Confirm tax code.
11. Confirm exact confirmation text.
12. Read `STRIPE_SECRET_KEY`.
13. Validate `sk_live_`.
14. Import Stripe SDK.
15. Existing Product preflight.
16. Create Product.

## Existing Product Preflight

Live execution searches all live Products by:

- `metadata.dake_item_id`

Behavior:

- 0 matches: creation may proceed
- 1 match: stop for manual resolution
- 2+ matches: stop as duplicate live Product state

Existing Products are not updated, deleted, or automatically reused.

## Product / Price / Payment Link Flow

Creation order:

1. Product
2. save state
3. Price using the created Product ID
4. save state
5. Payment Link using the created Price ID
6. save state
7. completed result

Placeholders are checked before sending dependent payloads:

- `__PRODUCT_ID_FROM_LIVE_PRODUCT__`
- `__PRICE_ID_FROM_LIVE_PRICE__`

Payment Link line item uses:

- live Price ID
- `quantity=1`

## State / Resume

State path:

- `tools/reports/release_artifacts/<product_id>/stripe_release_state.json`

State includes:

- product id
- product type
- source original
- input hash
- status
- Stripe object IDs when created
- safe error object

Partial states stop automatic resume:

- `product_created`
- `price_created`
- `failed`
- `existing_detected`

## Result / Writeback Boundary

Result files are written only after complete success.

This phase does not write Payment Link URLs back to `ORIGINAL.md` and does not change `payment_status`.

Writeback remains a separate reviewed phase after live creation and result review.

## Path Safety

`product_id` rejects:

- `..`
- `/`
- `\`
- control characters
- empty values

Artifact output paths are always under:

- `tools/reports/release_artifacts/`

## Error Sanitization

Saved errors use:

- `error_type`
- `safe_message`
- `stripe_error_code`
- `stripe_request_id`
- `failed_step`
- `occurred_at`

Secret-like values are redacted.

## Regression Tests

Passed:

- `py_compile tools/release_product.py`
- `py_compile tools/store/stripe_release_core.py`
- `py_compile tools/tests/test_release_product_pack_scaling.py`
- `python tools/tests/test_release_product_pack_scaling.py`
- `python tools/release_product.py DAKE_Pack_Document`
- `python tools/release_product.py DAKE_Pack_Memo`
- `python tools/release_product.py DOES_NOT_EXIST`
- `python tools/release_product.py "..\..\example"`
- `python tools/release_product.py dake_pdf_viewer`
- `python tools/release_product.py DAKE_Pack_Document --execute-live`
- full confirmation with `STRIPE_SECRET_KEY` unset
- `python tools/store/create_stripe_live_rollout.py`

Observed:

- Pack dry-runs: `ready_for_live_execution=yes`, `errors=0`, `secret_read=no`, `live_api_called=no`
- existing Stripe item `dake_pdf_viewer`: blocked
- missing product: blocked
- path traversal: blocked
- missing live confirmations: blocked before Secret read
- full live confirmations without Secret: `STRIPE_SECRET_KEY is not set`, `live_api_called=no`
- existing 45-item rollout: `candidate_count=45`, `errors=0`, `live_api_called=no`, `secret_read=no`

## Secret and Personal Data Review

No matches for:

- `sk_test_`
- `sk_live_`
- `whsec_`
- buyer email patterns
- local absolute Pack ZIP paths

Dry-run payloads contain only relative Pack ZIP paths.

## Live Execution Recommendation

READY_FOR_PACK_LIVE_EXECUTION

Before public release of each live Payment Link:

1. Create the live Payment Link with explicit confirmations.
2. Review `stripe_release_result.json`.
3. Open the Stripe-hosted Checkout page.
4. Confirm buyer email is collected.
5. Confirm product name, price, JPY currency, and manual delivery description.
6. Only then proceed to write the Payment Link URL back to `ORIGINAL.md` in a separate reviewed phase.
