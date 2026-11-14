# DAKE Commerce Rule

## Scope

CommerceはStripe決済だけを扱います。SHIP、DISTRIBUTION、BOOTH登録、Store同期・公開、Cloudflare、SNS、Bufferへ自動的に進みません。商品の正本は各 `ORIGINAL.md` で、generated JSONは派生ビューです。

## Normal flow

通常入口は `tools/release_product.py` の1本です。

```powershell
python tools\release_product.py <product_id>
python tools\release_product.py <product_id> --apply-live
```

1. `ORIGINAL.md` の商品名、金額、通貨、税区分、商品固有のfulfillment条件を検証する。
2. デフォルトのdry-runで計画を作る。Secretを読まず、Stripe APIを呼ばない。
3. 人間がlive操作を明示する場合だけ `--apply-live` を使う。この1回のintentをlive承認とし、追加のconfirmation textや二回目のFinalize承認は要求しない。
4. `metadata.dake_item_id` を軸にlive Productを照合する。0件なら作成、完全一致1件なら再利用、複数または不一致なら停止する。
5. Product配下のone-time Priceを通貨、金額、metadata等で照合する。完全一致1件を再利用し、0件なら新規Priceを作る。既存Priceを変更しない。完全一致が複数なら停止する。
6. 対象Priceとmetadataに一致するactive Payment Linkを照合する。完全一致1件を再利用し、0件なら作成し、複数なら停止する。
7. Product、Price、Payment LinkをStripe APIから再取得し、商品名、金額、通貨、対応関係、metadata、live/active状態を検証する。
8. 全read-back成功後だけ、同じ処理内で `stripe_payment_link` と `payment_status: stripe_ready` を `ORIGINAL.md` へ原子的に書き戻し、再読込で一致を確認する。
9. ここで `COMMERCE = READY`。Checkout目視、Store production review、Store同期・公開はCommerce完了条件ではない。

既存stateは復旧用の観測記録です。stateの存在、`product_created`、`price_created` 等だけで手動resumeを要求しません。毎回Stripe側をread-backし、完全一致なら安全に続行します。

## Safety invariants

- dry-runをデフォルトにし、dry-runでは `STRIPE_SECRET_KEY` を読まず、Stripe APIを呼ばない。
- liveは `sk_live_` だけを受け付け、test keyや不明な形式を拒否する。
- Secretらしい値をrepo、JSON、Markdown、HTML、JavaScript、CLI出力へ保存・表示しない。
- Product、Price、Payment Linkのcreateには安定したidempotency keyを使う。
- `metadata.dake_item_id`、商品固有metadata、Product / Price / Payment Linkの対応を照合する。
- 完全一致なら再利用する。不一致、複数候補、read-back不一致は停止し、既存live objectを変更しない。
- 価格変更は既存Priceの変更ではなく、新しいone-time Priceを作れる構造を保つ。
- `ORIGINAL.md` は全Stripe read-back成功後だけ更新する。失敗時は `stripe_ready` にしない。
- Pack等のmanual fulfillment noticeは商品固有の検証として残せる。通常商品へ一律のCheckout review gateを追加しない。
- Codexはユーザーからlive実行を明示されない限り `--apply-live` を実行しない。

通常artifactは `commerce_plan.json`、`commerce_result.json`、障害復旧用 `commerce_state.json` に限定します。同じ内容のMarkdownを二重生成しません。

## Legacy compatibility

`tools/release_product_pipeline.py` と `tools/store/release_pipeline_core.py` は、DISTRIBUTION、Commerce、Store、production review、socialを含んだ旧大型stage machineです。通常Commerce入口には使いません。

`tools/finalize_product_release.py`、`tools/record_checkout_review.py`、`tools/record_production_review.py` は履歴再現・旧artifact互換用です。通常Commerceの必須工程ではありません。

次の初回大量展開用ツールもLegacyです。通常の商品追加では使いません。

- `select_stripe_pilot10.py`
- `review_stripe_pilot10_test_results.py`
- `finalize_stripe_pilot10_review.py`
- `create_stripe_payment_links_test.py`
- `build_stripe_live_rollout_payloads.py`
- `create_stripe_live_rollout.py`
- `review_stripe_pack2_preflight.py`
- `check_stripe_pack2_manual_delivery_ready.py`

Legacy toolは履歴互換のため残しますが、通常フローへgateやartifactを持ち込みません。
