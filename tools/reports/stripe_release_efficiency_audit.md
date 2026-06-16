# Stripe Release Efficiency Audit

Generated: 2026-06-16

## Current Flow

The current DAKE commerce release flow is product-id based:

1. Prepare `ORIGINAL.md` and shipping assets.
2. Record BOOTH URL when BOOTH registration is complete.
3. Run Stripe dry-run payload generation.
4. Run live Stripe creation only with explicit confirmation.
5. Record human Checkout review.
6. Finalize the Payment Link URL back to `ORIGINAL.md`.
7. Regenerate Store JSON.
8. Sync `dake-store-site`.
9. Record production Store review.
10. Create Buffer drafts for X, Threads, and Instagram.
11. Close formal release as `v2_closed`.

## Required Safety Gates

KEEP:

- Dry-run is the default for Stripe payload work.
- Live Stripe execution requires `STRIPE_SECRET_KEY`, live key validation, product-id confirmation, tax-code confirmation, and exact confirmation text.
- Checkout browser review remains a human gate.
- Store production review remains a human gate.
- Secrets are never written to generated JSON, Markdown reports, public JS, or Store static files.
- Stripe object mutation is never performed by pipeline status inspection.
- Buffer normal shipping creates drafts only. It does not publish, schedule, or intentionally queue posts.

## Duplicate Work

MERGE:

- Stripe Product, Price, and Payment Link payload construction was duplicated inside the DAKE adapter. The payment-link payload core is now in `tools/commerce/stripe_payment_link_core.py`.
- App and Pack social source lookup was split. `release_source_policy.py` now exposes product-level lookup while preserving app-only helpers.
- Pipeline completion and formal social completion were previously conflated in operator notes. The pipeline now reports separate `commerce_status`, `social_status`, and `formal_release_status`.

KEEP:

- `tools/store/stripe_release_core.py` remains the DAKE adapter. It owns DAKE source discovery, `ORIGINAL.md`, Pack ZIP checks, BOOTH fields, DAKE metadata, and DAKE Store URLs.
- Existing CLI entrypoints remain compatible.

## Unnecessary Artifacts

REMOVE from normal future requests:

- Repeated full narrative reports for every Stripe step after the release command has stable status output.
- Timestamp-only generated JSON churn.
- Separate manual status write-ups when `tools/release_product_pipeline.py <product_id> --save-report` already records status.

DEFER:

- A broad rewrite of every release script into one common library. The current split is sufficient for v1 and avoids churn.
- Automatic social scheduling. Draft creation is enough for normal release closure.

## KEEP

- Explicit confirmations before live Stripe creation.
- Human Checkout review.
- Human production Store review.
- Source-of-truth write-back to `ORIGINAL.md`.
- Generated Store JSON as a derived view.
- State files for live Stripe execution and resume safety.
- Idempotency keys and payload hashes.
- Buffer draft evidence.

## MERGE

- Generic Stripe payload building: moved to `tools/commerce/stripe_payment_link_core.py`.
- Product social source lookup: shared app/pack lookup through `release_source_policy.py`.
- Formal completion reporting: added directly to the release pipeline status output.

## REMOVE

- None removed in Phase 18. The safe change was consolidation and status clarity, not deleting working gates.

## DEFER

- Subscription, recurring prices, usage billing, customer portal, automatic entitlement, webhook fulfillment, refund automation, and provisioning.
- Store-wide image asset sync automation. Phase 18 only added the missing DAKE Mail Pack thumbnail that was required for Instagram draft creation.
- Full social rollout for the existing 51 incomplete available products.

## Codex Usage Reduction

Target normal Pack release after Phase 18:

- Before BOOTH registration: 1 Codex request to create Pack assets and BOOTH-ready material.
- After BOOTH registration: 1 Codex request to run Stripe release, finalize Store, production review handoff, Buffer drafts, and v2 closeout.
- Exception requests only when Checkout, Store production review, Stripe live execution, or Buffer draft creation fails.

This keeps human safety gates while reducing repeated status narration and duplicated manual checks.

## Recommended Normal Release

Recommended command path:

```powershell
python tools\release_product_pipeline.py <product_id>
python tools\release_product_pipeline.py <product_id> next
python tools\release_product_pipeline.py <product_id> advance
```
Use `tools/release_product.py <product_id>` for Stripe dry-run and explicit live execution, then return to the pipeline for status. Formal release is complete when:

```text
current_stage=RELEASE_COMPLETE
commerce_status=complete
social_status=buffer_drafts_complete
formal_release_status=v2_closed
next_action=none
next_formal_action=none
```
