# Stripe Payment Link Live Execution Self Review

## Purpose

Final self review of `tools/store/create_stripe_live_rollout.py` before operator live execution.

This review did not call the Stripe API, did not read `STRIPE_SECRET_KEY`, and did not create Product, Price, or Payment Link objects.

## Target

- `tools/store/create_stripe_live_rollout.py`
- `tools/reports/stripe_payment_link_live_execution_plan.md`
- `tools/reports/stripe_payment_link_live_rollout_payloads.json`

## SDK Version

- stripe-python: `15.2.1`
- Current implementation uses the legacy resource style, not `StripeClient`.
- Idempotency keys are passed as request options via keyword argument, for example `stripe.Product.create(**payload, idempotency_key=key)`.

## Conclusion

`READY_FOR_OPERATOR_LIVE_EXECUTION`

The conclusion is conditional on the operator intentionally providing a live Stripe Secret Key and confirming the tax-code candidate gate. This review is not a tax or legal determination.

## Blocking Findings

None after the fixes in this phase.

## Fixed Findings

- Changed existing Product handling from per-item just-in-time checks to all-target preflight before any write call.
- Added atomic JSON writes for state/result JSON using a temporary file, flush, fsync, and `os.replace`.
- Added input payload file SHA-256 to state and resume validation.
- Tightened resume behavior so `product_created`, `price_created`, `existing_object_detected`, and `failed` require manual resolution.
- Clarified idempotency-key passing so keys are not inserted into the Stripe payload object.
- Added placeholder checks after replacing Product and Price IDs.
- Added sanitized error records with `error_type`, `safe_message`, request/code fields, failed step, and timestamp.
- Added result-file guard so final result is written only when all 45 items are completed and `livemode=true`.

## Accepted Risks

- Live execution has not been performed in this phase.
- Existing live Stripe objects can only be conclusively checked during the real live execution preflight.
- Tax code values remain configured candidates and require operator confirmation.
- The script is still DAKE-oriented; reusable extraction is a next-step design task, not part of this phase.

## Guard Order Review

Expected order for `--execute-live`:

1. Parse command-line arguments.
2. Read the dry-run payload.
3. Validate count, hashes, readiness, duplicate IDs, duplicate idempotency keys, price/currency, metadata, and secret-like values.
4. Validate `--confirm-count 45`.
5. Validate `--confirm-tax-codes`.
6. Validate exact confirmation text.
7. Read `STRIPE_SECRET_KEY`.
8. Reject missing, `sk_test_`, and non-`sk_live_` keys.
9. Import and initialize Stripe SDK.
10. Reach read-only live preflight and then write calls only if every prior gate passed.

Dry-run does not read Secret, import Stripe, call Stripe API, or write state/result files.

## Stripe v15 / Idempotency Review

The installed SDK is `stripe-python 15.2.1`. The script keeps the existing resource-style implementation and passes idempotency keys as keyword request options:

- Product: `stripe.Product.create(**params, idempotency_key=product_key)`
- Price: `stripe.Price.create(**params, idempotency_key=price_key)`
- Payment Link: `stripe.PaymentLink.create(**params, idempotency_key=link_key)`

The script does not mutate the original payload with `idempotency_key`.

## Payload Hash Review

The script canonicalizes each Product, Price, and Payment Link payload with sorted JSON keys and compact separators, then checks:

- `product_payload_sha256`
- `price_payload_sha256`
- `payment_link_payload_sha256`

Any mismatch stops before Secret read and before Stripe API calls.

## Existing Product Preflight Review

Live execution lists live Products with `auto_paging_iter()` and indexes them by exact `metadata.dake_item_id`.

Before any create call, the script checks all target items:

- 0 matches: eligible for creation.
- 1 match: state records `existing_object_detected`, related Price/Payment Link IDs if found, and stops.
- 2+ matches: state records failure and stops.

It does not normalize ID case. `DAKE_Sticky_Memo` and `dake_sticky_memo` remain distinct.

## Pagination Review

Product, Price, Payment Link, and Payment Link line item listing use `auto_paging_iter()`, so the preflight does not depend on only the first page.

## State Atomicity Review

State/result JSON writes use a same-directory temporary file, flush, fsync, and `os.replace`.

State includes:

- `mode=live`
- `input_payload_file`
- `input_payload_file_sha256`
- `expected_count=45`
- timestamps
- per-item safe status and object IDs

State does not store Secret values, request headers, Stripe client objects, or raw exception tracebacks.

## Resume Review

Rules:

- state exists without `--resume`: stop.
- `--resume` without state: stop.
- completed state: stop.
- payload hash mismatch: stop.
- `product_created`, `price_created`, `existing_object_detected`, or `failed`: stop for manual resolution.
- completed items are not recreated.

The script does not infer continuation from partial object creation states.

## Per-item Execution Review

Each item is processed as:

1. Product create.
2. Save Product ID to state.
3. Replace Price payload Product placeholder with the real Product ID.
4. Price create.
5. Save Price ID to state.
6. Replace Payment Link payload Price placeholder with the real Price ID.
7. Payment Link create.
8. Save Payment Link ID/URL and mark completed.

Placeholders are checked before Price and Payment Link calls.

## Error Handling Review

On error:

- the current item is marked failed
- a sanitized error object is stored
- execution stops
- remaining items are not skipped over

Stored error fields are safe operational fields, not tracebacks or request headers.

## Secret Leak Review

The script does not print, log, or write Secret values.

Secret-like value scan over the script and execution plan found no real key-shaped token.

## Result File Review

Result files are written only after all 45 items are completed and every completed item has `livemode=true`.

Partial success is not written as final success.

## Future release_product.py Extraction

Reusable as-is or close to reusable:

- payload loading
- payload/hash/idempotency validation
- confirmation gates
- Secret validation
- StripeObject to plain dict conversion
- existing Product indexing
- state save/load
- per-item Product/Price/Payment Link creation
- sanitized errors

DAKE-specific parts to extract later:

- expected count `45`
- confirmation text
- game item list
- fixed file paths
- DAKE metadata names and source file expectations

No large refactor was done in this phase to keep the live-execution safety diff small.

## Dry-run Verification

Commands run without setting `STRIPE_SECRET_KEY`:

```powershell
python -m py_compile tools\store\create_stripe_live_rollout.py
python tools\store\create_stripe_live_rollout.py
python tools\store\create_stripe_live_rollout.py --execute-live
python tools\store\create_stripe_live_rollout.py --execute-live --confirm-count 45 --confirm-tax-codes --confirmation-text "CREATE 45 DAKE LIVE PAYMENT LINKS"
```

Observed:

- dry-run: `candidate_count=45`
- dry-run: `errors=0`
- dry-run: `live_api_called=no`
- dry-run: `secret_read=no`
- `--execute-live` only: stopped at confirmation gate
- full confirmation without Secret: `STRIPE_SECRET_KEY is not set`
- no state/result file was generated
- no Stripe API was called

## Live Execution Recommendation

Proceed only if the operator intentionally sets a live key, accepts the configured tax-code candidates, and is ready to stop for manual resolution if existing Stripe Products are detected.

Recommended command remains:

```powershell
python tools\store\create_stripe_live_rollout.py `
  --execute-live `
  --confirm-count 45 `
  --confirm-tax-codes `
  --confirmation-text "CREATE 45 DAKE LIVE PAYMENT LINKS"
```

## Not Done

- No Stripe API call.
- No Stripe Secret Key read.
- No Product, Price, or Payment Link creation.
- No Stripe object update or deletion.
- No `ORIGINAL.md`, generated JSON, dake-store-site, or Store production update.
- No Payment Link URL write-back.
- No tax or legal determination.
