# Stripe Commerce v1 Closeout

Generated: 2026-06-16

## Current Architecture

DAKE commerce now uses `ORIGINAL.md` as the source of truth. Stripe Product, Price, Payment Link, Store JSON, Store pages, BOOTH text, and release reports are derived views or external results that must be written back to the source after review.

The Stripe layer is split into:

- Generic Stripe core: `tools/commerce/stripe_payment_link_core.py`
- DAKE adapter: `tools/store/stripe_release_core.py`
- Product pipeline controller: `tools/store/release_pipeline_core.py`
- CLI entrypoints: `tools/release_product.py` and `tools/release_product_pipeline.py`

## Waste Audit

- Payload generation was the best candidate for reuse and was split into a pure commerce core.
- Safety gates were not removed. They are load-bearing.
- Full release status is now available from the pipeline, reducing repeated manual summaries.
- Social completion is separated from commerce completion.

## Changes Applied

- Added reusable Stripe Payment Link payload core.
- Kept DAKE-specific source discovery and metadata in the DAKE adapter.
- Added Pack support to `tools/release_social.py`.
- Added Pack support to `tools/check_release_social.py`.
- Added `commerce_status`, `social_status`, `formal_release_status`, and `next_formal_action` to pipeline status.
- Created Buffer drafts for `DAKE_Pack_Mail`.
- Added the missing DAKE Mail Pack Store thumbnail at the public Store image path.
- Added fixture tests for generic commerce payloads and formal release status.

## Required Safety Gates

- No live Stripe API mutation was performed in Phase 18.
- No existing Stripe object was changed.
- Buffer mutation was limited to draft creation.
- No Buffer publish, schedule, or queue operation was performed.
- No Stripe Secret Key, Buffer token, Authorization header, buyer information, card information, or private URL was written to repo files.
- Checkout and production Store reviews remain separate human gates.

## DAKE Adapter

`tools/store/stripe_release_core.py` remains responsible for:

- `01_apps/**/ORIGINAL.md`
- `04_packs/**/ORIGINAL.md`
- DAKE metadata
- BOOTH URL and GitHub Release URL
- Pack ZIP, size, and SHA256 validation
- DAKE Store metadata
- DAKE-specific Stripe metadata

## Generic Stripe Core

`tools/commerce/stripe_payment_link_core.py` is responsible for:

- Stripe Product payload
- Stripe Price payload
- Stripe Payment Link payload
- One-time price validation
- Metadata safety checks
- Checkout notice length checks
- Redirect URL public-safety checks
- Stable idempotency-key helper

It does not know DAKE folders, BOOTH rules, Pack ZIP structure, GitHub Release URLs, or Store generation.

## Supported Product Matrix

| Product mode | Status |
|---|---|
| DAKE app one-time | supported |
| DAKE Pack manual delivery | supported |
| generic digital product one-time | supported |
| generic web service one-time manual fulfillment | supported |
| generic web service one-time redirect | supported |
| subscription SaaS | not supported in v1 |
| automatic provisioning | not supported in v1 |
| webhook fulfillment | not supported in v1 |

## App Support

Existing DAKE apps remain compatible with the release pipeline. Legacy apps that were completed before the product pipeline can still show `LEGACY_COMPLETE`; they are not forced to recreate Stripe artifacts.

## Pack Support

Pack commerce is supported through manual private download delivery:

- Pack ZIP validation
- Manual delivery checkout notice
- BOOTH URL record
- Stripe dry-run and live creation path
- Checkout review
- Store production review
- Buffer draft social release
- formal `v2_closed`

`DAKE_Pack_Mail` is the first Pack to complete the full v2 path.

## Web Service Support

Commerce v1 supports one-time web service products at the payload level:

- `web_service_one_time_manual_fulfillment`
- `web_service_one_time_redirect`

The core rejects private or local redirect URLs. Automatic account provisioning is not part of v1.

## Unsupported Subscription Features

The following remain out of scope:

- recurring prices
- usage billing
- subscription lifecycle
- plan upgrade or downgrade
- cancellation workflow
- failed payment recovery
- customer portal
- webhook fulfillment
- automatic entitlement
- refund automation

Passing `price_model=recurring` stops with `unsupported_in_commerce_v1`.

## Buffer Pack Support

`DAKE_Pack_Mail` Buffer drafts were created for:

| Channel | Draft ID |
|---|---|
| X | `6a310df504f576fe42b04034` |
| Threads | `6a310df6a8215e3dea50f650` |
| Instagram | `6a310df7924339d68c6ec98b` |

Evidence:

- `tools/reports/release_artifacts/DAKE_Pack_Mail/social_release.json`
- `tools/reports/release_artifacts/DAKE_Pack_Mail/social_release.md`
- `tools/reports/release_artifacts/DAKE_Pack_Mail/social_posts.md`

The Store image URL used for Instagram is:

```text
https://store.dakeapp.com/assets/images/products/DAKE_Pack_Mail/thumbnail.jpg
```

The URL returned `Content-Type: image/jpeg` and `Content-Length: 101582`.

## Formal Release Status

Current DAKE Mail Pack status:

```text
current_stage=RELEASE_COMPLETE
commerce_status=complete
social_status=buffer_drafts_complete
formal_release_status=v2_closed
next_action=none
next_formal_action=none
```

## Codex Usage Reduction

Normal Pack release should now require two high-level Codex requests:

1. Create the Pack and BOOTH-ready material.
2. After BOOTH registration, run product release completion through Stripe, Store, production review handoff, Buffer drafts, and v2 closeout.

Additional requests should be exceptions for failed gates, not normal procedure.

## Regression

Executed:

```powershell
python -m py_compile tools\commerce\stripe_payment_link_core.py tools\store\stripe_release_core.py tools\release_source_policy.py tools\release_social.py tools\check_release_social.py tools\store\release_pipeline_core.py tools\release_product_pipeline.py tools\tests\test_commerce_stripe_payment_link_core.py tools\tests\test_release_product_pipeline.py tools\tests\test_release_product_pack_scaling.py
python tools\tests\test_commerce_stripe_payment_link_core.py
python tools\tests\test_release_product_pipeline.py
python tools\tests\test_release_product_pack_scaling.py
python tools\check_release_social.py --only-available --report-dir tools\reports\release_artifacts
python tools\release_product_pipeline.py DAKE_Pack_Mail --json
```

## Security

- Stripe Secret Key was not read in Phase 18.
- No Stripe live mutation was executed.
- Buffer API token was read only by `tools/release_social.py` during draft creation.
- Buffer token was not displayed or saved.
- No secret-like values were found in DAKE Mail Pack social artifacts.
- No buyer data was stored.
- No private download URL was exposed.
- No actual payment was performed.

## Deferred Work

- Full social rollout for remaining available products.
- Store thumbnail sync automation.
- Subscription and webhook commerce v2.
- Generic non-DAKE product adapters for BORINEF Labs, games, templates, or other digital products.

## Conclusion

STRIPE_COMMERCE_V1_CLOSED
