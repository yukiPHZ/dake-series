# Release Product Pack Support Review

## 目的

`python tools\release_product.py <product_id>` をDAKEの単品正式出荷入口として追加し、Pack商品も同じ入口からStripe dry-runできる状態にした。

## 新しいCLI

- `python tools\release_product.py DAKE_Pack_Document`
- `python tools\release_product.py DAKE_Pack_Memo`

デフォルトはdry-run。Stripe APIは呼ばず、Secretも読まない。

live実行を行う場合は、少なくとも以下を要求する設計にした。

- `--execute-live`
- `--confirm-product-id <product_id>`
- `--confirm-tax-code`
- `--confirmation-text "CREATE LIVE PAYMENT LINK <product_id>"`

## 対応商品種別

- `app`
- `pack`

Packでは追加で手動配布ゲートを確認する。

## 正本解決

商品IDは `tools/generated/store_products.generated.json` を索引として使用し、最終的には `source_original` の `ORIGINAL.md` 実体を読む。

探索対象は以下。

- `01_apps/**/ORIGINAL.md`
- `04_packs/**/ORIGINAL.md`

## Pack Delivery Gate

Packでは以下を確認する。

- Pack ZIPが存在する。
- Pack ZIPサイズが `pack_manifest.json` と一致する。
- Pack ZIP SHA256が `pack_manifest.json` と一致する。
- `purchase_delivery_ready=yes`
- `purchase_delivery_method=manual_email_private_download`
- `stripe_creation_method=manual_dashboard_ready`
- 配布期限が正本に記録されている。
- 購入者向け案内、再送、不達対応、共通運用ルール参照がある。

## Stripe Payload

Product / Price / Payment Link の3 payloadを生成する。

- Product: `name`, `description`, `active`, `tax_code`, `metadata`
- Price: `currency=jpy`, `unit_amount`, `product=__PRODUCT_ID_FROM_LIVE_PRODUCT__`
- Payment Link: `line_items`, `metadata`, `payment_intent_data.metadata`

## Metadata

Pack metadataには以下を含める。

- `dake_item_id`
- `dake_type=pack`
- `source_repo=DAKE_series`
- `source_original`
- `store_url`
- `booth_url`
- `purchase_delivery_method`
- `delivery_policy`

購入者メール、非公開配布URL、Pack ZIPの絶対パス、Secretは含めない。

## Payload Hash

各payloadをJSON key sort済みの正規化文字列にしてSHA-256化する。

- `product_payload_sha256`
- `price_payload_sha256`
- `payment_link_payload_sha256`

## Idempotency

payload hashから決定的に生成する。

- `product_idempotency_key`
- `price_idempotency_key`
- `payment_link_idempotency_key`

既存45件の `dake-live-*` と衝突しないよう、単品コマンドでは `dake-release-*` 名前空間を使う。

## Live Execution Guards

live実行は確認引数とtax code確認を通過するまでSecretを読まない。

Secretは `STRIPE_SECRET_KEY` のみ。`sk_live_` 以外は拒否する。

## Existing Product Preflight

live実行時は、Stripe Product一覧を `metadata.dake_item_id` で照合する設計。

- 一致0件: 作成可能
- 一致1件: 自動再利用せず停止
- 一致2件以上: 異常として停止

## State / Resume

live実行時は商品ごとに以下へstateを保存する。

- `tools/reports/release_artifacts/<product_id>/stripe_release_state.json`

`product_created` / `price_created` / `failed` / `existing_detected` は人間確認なしで自動resumeしない。

## Result / Writeback

成功時は以下に結果を保存する設計。

- `stripe_release_result.json`
- `stripe_release_result.md`

Phase 16CではPayment Link URLの `ORIGINAL.md` 書き戻しは行わない。

## Current Pack Dry Runs

| product_id | price | currency | delivery_ready | method | distribution_file | ready | errors |
|---|---:|---|---|---|---|---|---:|
| DAKE_Pack_Document | 1480 | jpy | yes | manual_email_private_download | DAKE_Pack_Document.zip | yes | 0 |
| DAKE_Pack_Memo | 980 | jpy | yes | manual_email_private_download | DAKE_Pack_Memo.zip | yes | 0 |

## Regression Review

- 既存45件live dry-runは `candidate_count=45` / `errors=0` / `live_api_called=no` / `secret_read=no` を維持。
- generated JSONは再生成していない。
- 現在の集計は `total=53`, `stripe_ready=50`, `booth_only=2`, `preparing=1`。
- `dake_pdf_viewer` は既に `stripe_ready` のため安全停止する。

## Future Pack Scaling

今後Packが増えた場合も、Packごとに `release_product.py <product_id>` を実行する。

一括Pack処理を作る場合も、この単品コマンドを順番に呼び出し、Stripe作成ロジックを重複実装しない。

## Cross-brand Extraction

DAKE固有部分は `ORIGINAL.md` 解決、Store URL、Pack手動配布ルールに閉じる。

Stripe payload、hash、idempotency、Secret管理、state/result保存は将来SHIMARISUや他ブランドへ切り出せる構成に寄せた。

## Conclusion

READY_FOR_PACK_LIVE_EXECUTION_REVIEW
