# encoding: utf-8

# DAKE Series by しまりす不動産

Simple, fast, single-purpose desktop apps for real work.

Vibe-coded by Yukihiko Kikuta.

## Concept

DAKE Series is a collection of lightweight desktop tools designed to remove small but real work friction.

The core idea is simple:

- one app = one job
- fast to launch
- easy to understand
- stable in actual use
- no unnecessary UI
- no heavy effects
- local-first and practical

## Source Policy

The project uses `ORIGINAL.md` as the true source of truth for each app, Pack, and product.

`README.md`, `DAKE_META`, `release_body.md`, `booth_product.txt`, generated JSON, BOOTH pages, and Store pages are derived views.

If an important fact exists only in a derived view, it should be reviewed and moved back to `ORIGINAL.md` when appropriate.

## Store

`store.dakeapp.com` is a sales view, not the source of truth.

Store data is generated from `ORIGINAL.md` into `store_products.generated.json` and synced to `dake-store-site`.

Store sync command:

```powershell
python tools\store\sync_store_to_site.py
```

## Shipping

Formal DAKE shipping includes:

- GitHub Release
- BOOTH
- dakeapp.com
- store.dakeapp.com
- payment_status confirmation
- Cloudflare confirmation

GitHub Release alone is not treated as formal shipping completion.

## Payment Links

Stripe Payment Link can be used for Store purchase flow.

Stripe Secret, API keys, and Webhook Secret must not be stored in the public repo, generated JSON, or Store JavaScript.

## Project Structure

```text
00_core/
01_apps/
02_assets/
03_docs/
04_packs/
tools/
```

## Copyright

© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta

Unauthorized reproduction or redistribution is prohibited.
