# ORIGINAL Phase5-2 Stripe Payment Link Ready

## 目的

`ORIGINAL.md` 由来の Store 表示用 generated JSON に Stripe Payment Link 用フィールドを追加し、Store UI の購入導線出し分け準備を確認する。

## 生成スクリプト

- `tools/store/generate_store_products.py`

## 生成ファイル

- `tools/generated/store_products.generated.json`
- `tools/generated/README.md`

## 参照したORIGINAL

- discovered: 53

## 生成対象件数

- generated_at: 2026-06-14T22:23:39+09:00
- items: 53

## type別件数

- app: 50
- pack: 2
- shimarisu_pack: 1

## skipped件数

- skipped: 0

## payment_status別件数

- stripe_ready: 50
- booth_only: 2
- preparing: 1
- free_download: 0
- not_for_sale: 0

## 主な未確定項目

- download_url: 53
- store_url: 53
- support_policy: 53
- stripe_price_id: 53
- stripe_payment_link: 3

## SHIMARISU Pack参照結果

- result: included
- source: `C:/Users/yukiz/devlop/SHIMARISU/ORIGINAL.md`

## 注意点

- `store_products.generated.json` は正本ではなく、手編集禁止の生成物。
- Store表示を変更する場合は、各商品の `ORIGINAL.md` を修正する。
- 未確定値はJSON上では `null` に正規化した。
- Stripe Payment Linkが未設定の商品には、Stripe購入ボタンを出さない。
- Stripe Checkout API、Pages Functions、Webhook、R2、download_url確定は今回未実施。

## 次Phase提案

1. Stripe Payment Linkを付ける商品を絞る。
2. Payment Linkを `ORIGINAL.md` へ戻す運用を決める。
3. Store本番反映前にPayment Linkあり商品のみ `Stripeで購入` を目視確認する。
4. Checkout API / Webhook / R2 は別Phaseで検討する。
