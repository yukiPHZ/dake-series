# ORIGINAL Phase4-7 Store Generated Build

## 目的

`ORIGINAL.md` 由来の Store 表示用 generated JSON を実生成し、Store構築前にデータ生成基盤を確認する。

## 生成スクリプト

- `tools/store/generate_store_products.py`

## 生成ファイル

- `tools/generated/store_products.generated.json`
- `tools/generated/README.md`

## 参照したORIGINAL

- discovered: 53

## 生成対象件数

- generated_at: 2026-06-05T13:10:07+09:00
- items: 53

## type別件数

- app: 50
- pack: 2
- shimarisu_pack: 1

## skipped件数

- skipped: 0

## 主な未確定項目

- download_url: 53
- store_url: 53
- support_policy: 53
- stripe_price_id: 53

## SHIMARISU Pack参照結果

- result: included
- source: `C:/Users/yukiz/devlop/SHIMARISU/ORIGINAL.md`

## 注意点

- `store_products.generated.json` は正本ではなく、手編集禁止の生成物。
- Store表示を変更する場合は、各商品の `ORIGINAL.md` を修正する。
- 未確定値はJSON上では `null` に正規化した。
- Storeサイト、Stripe、Cloudflare Pages、download_url確定は今回未実施。

## 次Phase提案

1. Store側でこのJSONを読む最小UIを作る。
2. 画像パスをStore公開URLへ変換する方針を決める。
3. `download_type` の正式値と表示文言を固める。
4. Stripe接続情報を generated に含めるか、Store設定に逃がすかを決める。
