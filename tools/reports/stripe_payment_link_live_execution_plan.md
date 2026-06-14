# Stripe Payment Link Live Execution Plan

## Purpose

Prepare the script and operating rules for creating 45 DAKE Stripe live Product, Price, and Payment Link objects from the Phase 14B live dry-run payloads.

This phase implements the script and validates dry-run behavior only. It does not create live Stripe objects.

## Inputs

- `tools/reports/stripe_payment_link_live_rollout_payloads.json`
- `tools/reports/stripe_payment_link_live_rollout_payloads.csv`
- `tools/reports/stripe_payment_link_live_rollout_payloads.md`
- `tools/reports/stripe_payment_link_pilot10_final_review.csv`

## Dry-Run Target

- candidate count: 45
- normal DAKE apps: 43
- game apps: 2
- Product create planned: 45
- Price create planned: 45
- Payment Link create planned: 45
- live API called: no
- secret read: no

## Safety Gates

The default command is dry-run only:

```powershell
python tools\store\create_stripe_live_rollout.py
```

Live execution requires all explicit confirmations:

```powershell
python tools\store\create_stripe_live_rollout.py `
  --execute-live `
  --confirm-count 45 `
  --confirm-tax-codes `
  --confirmation-text "CREATE 45 DAKE LIVE PAYMENT LINKS"
```

This Phase does not run the live command.

## Secret Management

`STRIPE_SECRET_KEY` is read only in `--execute-live` mode after all confirmation gates pass.

Accepted:

- `sk_live_...`

Rejected:

- missing key
- `sk_test_...`
- any other format

The secret value must never be printed, logged, written to JSON/CSV/Markdown, or committed.

## Tax Code Confirmation

The dry-run payload contains configured candidates only:

- normal DAKE apps: `txcd_10202003`
- game apps: `txcd_10201000`

`--confirm-tax-codes` confirms only that the operator accepts these configured candidates for execution. It is not a tax or legal determination.

## Payload Hash Verification

Before live execution, the script canonicalizes each Product, Price, and Payment Link payload and verifies the SHA-256 hashes from Phase 14B:

- `product_payload_sha256`
- `price_payload_sha256`
- `payment_link_payload_sha256`

Any mismatch stops execution.

## Idempotency Keys

The script uses the Phase 14B keys as request idempotency keys for POST calls:

- `product_idempotency_key`
- `price_idempotency_key`
- `payment_link_idempotency_key`

The script does not regenerate keys.

## Existing Object Detection

Before creating a Product, live execution lists existing live Products and matches by:

- `metadata.dake_item_id`

Rules:

- 0 matches: create candidate may proceed.
- 1 match: stop and record `existing_object_detected`; manual resolution is required.
- 2 or more matches: stop as abnormal.

Existing Products are not automatically reused because Price, tax code, and metadata must be reviewed by a human first.

## One-Item Creation Order

Live execution is designed to complete one item at a time:

1. Existing Product check
2. Product create
3. Save state
4. Price create
5. Save state
6. Payment Link create
7. Save state
8. Mark item completed
9. Move to the next item

The script does not create all Products first and then all Prices.

## State Storage

Live execution writes incremental state to:

```txt
tools/reports/stripe_payment_link_live_execution_state.json
```

The state file stores object IDs, URLs, statuses, and errors only. It does not store Stripe Secret Key values.

## Interruption And Resume

`--resume` is supported for a stopped execution.

The script refuses to resume automatically from partial states such as `product_created` or `price_created`, because those cases require manual reconciliation against Stripe before continuing.

Completed items are not recreated.

## Error Handling

If one item fails:

- the error is written to state
- the script stops
- remaining items are not skipped over

This keeps failure blast radius small.

## Result Files

On successful live completion, the script is designed to write:

- `tools/reports/stripe_payment_link_live_execution_result.json`
- `tools/reports/stripe_payment_link_live_execution_result.csv`
- `tools/reports/stripe_payment_link_live_execution_result.md`

These files must not contain Secret values.

## Live Execution Command Draft

```powershell
cd C:\Users\yukiz\devlop\DAKE_series
$env:STRIPE_SECRET_KEY="sk_live_..."
python tools\store\create_stripe_live_rollout.py `
  --execute-live `
  --confirm-count 45 `
  --confirm-tax-codes `
  --confirmation-text "CREATE 45 DAKE LIVE PAYMENT LINKS"
Remove-Item Env:STRIPE_SECRET_KEY
```

## Not Done In This Phase

- No Stripe API call.
- No Stripe Product list retrieval.
- No Stripe Secret Key read.
- No `sk_live_` or `sk_test_` use.
- No Product, Price, or Payment Link creation.
- No Stripe object update or deletion.
- No `ORIGINAL.md`, generated JSON, dake-store-site, or Store production update.
- No Payment Link URL write-back.
- No Pack or preparing item creation.

## Next Phase Completion Conditions

Before running live execution, an operator must confirm:

- candidate count is 45
- all payload hashes pass
- idempotency keys are unique
- tax candidates are accepted for execution
- `STRIPE_SECRET_KEY` is a live key
- no live Product already exists for the target `dake_item_id`
