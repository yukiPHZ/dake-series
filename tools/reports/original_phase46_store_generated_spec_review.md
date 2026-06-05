# ORIGINAL Phase4-6 Store Generated Spec Review

## 目的

A1単品アプリ、DAKE Pack、SHIMARISU Pack の `ORIGINAL.md` 整備が進んだため、Store構築前に `store_products.generated.json` の位置づけとデータ構造を整理した。

今回はStoreサイト本体、Stripe、生成スクリプトは作成していない。

## 参照したファイル

- `00_core/DAKE_ORIGINAL_RULE.md`
- `00_core/DAKE_ORIGINAL_TEMPLATE_APP.md`
- `00_core/DAKE_ORIGINAL_TEMPLATE_PACK.md`
- `00_core/CHATGPT_CODEX_WORKFLOW.md`
- `tools/reports/original_phase45_app_inventory.md`
- `tools/reports/original_phase45j2_app_batch_rollout.md`
- `tools/reports/original_phase45h_pack_rollout.md`
- `tools/reports/original_phase45i_shimarisu_pack_rollout.md`

## 作成した仕様

| file | result |
|---|---|
| `00_core/DAKE_STORE_GENERATED_SPEC.md` | 作成 |

## 仕様化した内容

- `store_products.generated.json` はStore表示用の生成データであり、正本ではない。
- Store専用の商品正本は作らない。
- Store表示は `ORIGINAL.md` 由来の情報を読む。
- MVPでは `store_products.generated.json` 1ファイルに `items` 配列としてまとめる。
- `type` は `app` / `pack` / `shimarisu_pack` を定義した。
- 未確定値は generated JSON では `null` 推奨とした。
- Stripe Secret は絶対に含めない。
- Stripe Price ID はMVP必須にせず、任意項目 `stripe_price_id: null` とした。
- `download_url` はMVPでは未確定でよく、GitHub Release / BOOTH / R2 / Store購入後URLを候補として残した。

## 正本と生成物の関係

```text
ORIGINAL.md
↓
store_products.generated.json
↓
store.dakeapp.com
```

`store_products.generated.json` は手編集禁止。

Store表示を直す場合は、generated JSONではなく、該当商品の `ORIGINAL.md` を直す。

## 単品アプリの扱い

DAKE単品アプリは、各アプリの以下を正本とする。

```text
DAKE_series/01_apps/{app}/ORIGINAL.md
```

generated item の `type` は `app`。

## DAKE Packの扱い

DAKE Packは、各Packの以下を正本とする。

```text
DAKE_series/04_packs/{pack}/ORIGINAL.md
```

generated item の `type` は `pack`。

## SHIMARISU Packの扱い

SHIMARISU Packは、DAKE_series外の以下を正本とする。

```text
C:/Users/yukiz/devlop/SHIMARISU/ORIGINAL.md
```

generated item の `type` は `shimarisu_pack`。

`shimarisu-pack-release` は正本ではなく、配布repo / 公開ビューとして扱う。

## 今回決めた未確定値方針

- generated JSONでは `null` を推奨。
- 表示側で `準備中`、`BOOTHで見る`、`近日追加` などに変換する。
- `ORIGINAL.md` やMarkdownでは人間向けに `未確定` と書いてよい。

## 今回やらなかったこと

- `store_products.generated.json` の実生成
- 生成スクリプト作成
- Storeサイト構築
- Stripe実装
- Cloudflare Pages設定
- 各 `ORIGINAL.md` 更新
- README / booth_product / pack_ready 更新
- SHIMARISU repo更新

## 次Phase提案

1. `store_products.generated.json` 生成スクリプトの仕様を決める。
2. 単品アプリ、DAKE Pack、SHIMARISU Packから数件だけサンプル生成する。
3. 画像パスの公開URL変換方針を決める。
4. `download_type` の正式値を実装側で固定する。
5. Stripe接続情報を generated に含めるか、Store実装側設定に逃がすかを決める。
