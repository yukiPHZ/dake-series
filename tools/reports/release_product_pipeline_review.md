# Release Product Pipeline Review

## 目的

Phase 17A adds a product-id based release pipeline controller that detects the current release stage from `ORIGINAL.md` and existing artifacts. It does not create Stripe objects, read Stripe secrets, change Store production, or commit automatically.

## 新しい入口

```powershell
python tools\release_product_pipeline.py <product_id>
python tools\release_product_pipeline.py <product_id> status --json
python tools\release_product_pipeline.py <product_id> next
python tools\release_product_pipeline.py <product_id> advance
```

`status` is read-only by default. Reports are written only with `--save-report`.

## Source of Truth

The pipeline directly searches:

- `01_apps/**/ORIGINAL.md`
- `04_packs/**/ORIGINAL.md`

Generated JSON is used only as a derived-view consistency check. It is not required for product discovery.

## Stage Model

Implemented stages:

`SOURCE_INVALID`, `PREPARING_BLOCKED`, `SOURCE_READY`, `STRIPE_DRY_RUN_READY`, `STRIPE_LIVE_COMPLETED`, `CHECKOUT_REVIEW_PENDING`, `CHECKOUT_REVIEW_PASSED`, `SOURCE_FINALIZED`, `STORE_GENERATED`, `STORE_SYNC_PENDING`, `STORE_SYNCED`, `PRODUCTION_VERIFICATION_PENDING`, `RELEASE_COMPLETE`, `LEGACY_COMPLETE`, `INCONSISTENT`.

## Status Detection

The controller checks product id safety, direct `ORIGINAL.md` discovery, duplicate ids, product type, payment status, Pack delivery gates, dry-run artifacts, live state/result, Checkout review, finalize result, generated Store JSON, site Store JSON, and production review evidence.

## Advance Safety

`advance` handles only one boundary at a time.

- `SOURCE_READY`: runs `python tools\release_product.py <product_id>` dry-run only.
- `STRIPE_DRY_RUN_READY`: prints the live execution command and stops.
- `CHECKOUT_REVIEW_PENDING`: prints the Checkout review command and stops.
- `CHECKOUT_REVIEW_PASSED`: runs finalize dry-run only.
- `SOURCE_FINALIZED`: runs Store JSON generation only.
- Store sync and production verification remain manual stops in Phase 17A.

## Checkout Review Recording

`tools/record_checkout_review.py` records human browser review evidence under:

- `tools/reports/release_artifacts/<product_id>/checkout_review.json`
- `tools/reports/release_artifacts/<product_id>/checkout_review.md`

The command asks for product name, price, live URL, email field, quantity UI, delivery notice, private URL exposure, local path exposure, console errors, and actual payment status. It marks `passed` only when every required answer passes. Existing non-failed reviews are not overwritten.

## Production Review

The pipeline recognizes `production_review.json` when present. For the already completed two Pack releases, it safely uses `tools/reports/pack2_stripe_release_completion.md` as existing production review evidence instead of inventing new artifacts.

## Existing Product Compatibility

Observed stages:

- `DAKE_Pack_Document`: `RELEASE_COMPLETE`
- `DAKE_Pack_Memo`: `RELEASE_COMPLETE`
- `dake_pdf_viewer`: `LEGACY_COMPLETE`
- `video_shorts_cut`: `PREPARING_BLOCKED`

## Legacy Complete

Existing Stripe-ready apps without the newer per-product live state/result/Checkout artifacts are treated as `LEGACY_COMPLETE` when their source, generated JSON, and Store JSON agree. The pipeline does not backfill fake evidence.

## Pack Scaling

A synthetic Pack fixture was discovered from `04_packs/**/ORIGINAL.md` without generated JSON registration. Its stage was `SOURCE_READY`, and `advance` dispatched to `release_product.py` dry-run without requiring a code change or a hard-coded Pack id.

## Inconsistency Detection

Fixture checks covered:

- source `stripe_ready` with missing URL
- source URL and live result URL mismatch
- completed result with failed state
- failed Checkout review after source finalization
- generated Store JSON URL mismatch
- duplicate product id in two `ORIGINAL.md` files

All were rejected as `INCONSISTENT` or `SOURCE_INVALID` as appropriate.

## Secret and Personal Data

The pipeline does not read `STRIPE_SECRET_KEY`, environment variable listings, buyer email, buyer name, card information, private download URLs, or absolute Pack ZIP paths. Public Payment Link URLs are allowed.

## Git Boundary

The pipeline does not run git commit or git push. Status output reports DAKE_series and dake-store-site clean/dirty state when a git checkout is available.

## Future Hooks

The stage model can be extended for GitHub Release, BOOTH, Store publication, SNS announcement, Buffer draft, market observation, and post-sale delivery without replacing the existing pipeline entrypoint.

## Test Results

Commands run:

```powershell
python -m py_compile tools\release_product_pipeline.py tools\store\release_pipeline_core.py tools\record_checkout_review.py tools\tests\test_release_product_pipeline.py
python tools\tests\test_release_product_pipeline.py
python tools\release_product_pipeline.py DAKE_Pack_Document
python tools\release_product_pipeline.py DAKE_Pack_Memo
python tools\release_product_pipeline.py dake_pdf_viewer
python tools\release_product_pipeline.py video_shorts_cut
python tools\release_product_pipeline.py DOES_NOT_EXIST
python tools\release_product_pipeline.py "..\..\example"
python tools\store\create_stripe_live_rollout.py
python tools\store\generate_store_products.py
```

Results:

- unit/fixture tests: passed
- Pack 2 status: `RELEASE_COMPLETE`
- legacy app status: `LEGACY_COMPLETE`
- preparing status: `PREPARING_BLOCKED`
- unknown/traversal ids: safe stop
- live rollout default: dry-run, `live_api_called=no`, `secret_read=no`
- Store generation counts: total `53`, `stripe_ready=52`, `booth_only=0`, `preparing=1`
- generated JSON timestamp-only test diff was discarded and not included in Phase 17A changes

## Conclusion

READY_FOR_PIPELINE_PILOT
