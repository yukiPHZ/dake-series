# DAKE Store Generated Data Spec

## 目的

`store.dakeapp.com` が読む Store 表示用データの形式を定義する。

この仕様は、DAKE単品アプリ、DAKE Pack、SHIMARISU Pack を Store へ接続する前段として、`ORIGINAL.md` 由来の情報を静的サイトやビルド処理が読みやすい形へ変換するためのものである。

今回は Store サイト本体、Stripe 決済、Cloudflare Pages 設定、生成スクリプトは作らない。

## 基本思想

Store専用の商品正本は作らない。

真の正本は `ORIGINAL.md` である。

```text
ORIGINAL.md
↓
store_products.generated.json
↓
store.dakeapp.com
```

`store_products.generated.json` は Store 表示用の生成データであり、正本ではない。

いつでも `ORIGINAL.md` から再生成できる状態にする。

## 正本と生成物の関係

```text
ORIGINAL.md
= 真の正本

store_products.generated.json
= Store表示用の生成データ
= ORIGINAL.md由来
= 正本ではない
= 手編集禁止
```

Storeに表示する文言、価格、画像、注意事項、販売方針の元情報は、各商品の `ORIGINAL.md` に置く。

`store_products.generated.json` にだけ存在する判断情報を増やしてはいけない。

generated 側にだけ必要な接続情報が出た場合は、次のどちらかに分類する。

- Store実装のための接続情報
- ORIGINAL.mdへ戻すべき正本情報

## 生成ファイル

MVPでは以下の1ファイルを推奨する。

```text
store_products.generated.json
```

将来的に分ける場合の候補。

```text
store_apps.generated.json
store_packs.generated.json
store_index.generated.json
```

ただし、初期Storeでは1ファイルに `items` 配列としてまとめる。

## 生成対象

Store生成対象は、原則として以下。

- `status: available`
- BOOTH登録済み
- market系
- Store掲載候補
- DAKE Pack商品
- SHIMARISU Pack商品

ただし、最終的な掲載可否は各 `ORIGINAL.md` の Store表示用情報を優先する。

## 除外対象

通常のStore生成対象から除外する候補。

- `draft`
- `frozen`
- `prototype`
- `internal`
- `experimental`
- Store非掲載
- 正本が未整備のもの
- 販売対象ではない内部ツール

除外対象であっても、将来の移行や検証目的で別レポートに一覧化してよい。

## JSON構造

MVPの推奨構造。

```json
{
  "generated_at": "2026-XX-XXT00:00:00+09:00",
  "source_policy": "ORIGINAL.md is the source of truth. This file is generated and must not be edited manually.",
  "schema_version": "1.0.0",
  "items": [
    {
      "id": "",
      "type": "app",
      "source_repo": "DAKE_series",
      "source_original": "01_apps/DAKE_Time_AdvancedTimer/ORIGINAL.md",
      "title": "",
      "short_title": "",
      "catch": "",
      "description": "",
      "price": null,
      "currency": "JPY",
      "status": "available",
      "category": "",
      "tags": [],
      "image": null,
      "thumbnail": null,
      "download_type": null,
      "download_url": null,
      "booth_url": null,
      "github_release_url": null,
      "store_url": null,
      "support_policy": null,
      "disclaimer": "",
      "included_items": [],
      "stripe_payment_link": null,
      "payment_status": "booth_only",
      "stripe_price_id": null,
      "source_kind": "original",
      "is_generated": true
    }
  ]
}
```

## フィールド定義

| field | required | description |
|---|---:|---|
| `generated_at` | yes | 生成日時。ISO 8601形式、原則Asia/Tokyoのオフセット付き。 |
| `source_policy` | yes | `ORIGINAL.md` が正本であり、このJSONが手編集禁止の生成物であることを示す文。 |
| `schema_version` | yes | generated JSON構造のバージョン。 |
| `items` | yes | Storeに表示する商品配列。 |
| `id` | yes | 商品ID。単品アプリは app_id、Packは pack_id を優先。 |
| `type` | yes | `app` / `pack` / `shimarisu_pack` などの商品種別。 |
| `source_repo` | yes | 正本があるrepoまたは母艦。例: `DAKE_series`, `SHIMARISU`。 |
| `source_original` | yes | 参照元 `ORIGINAL.md` のパス。 |
| `title` | yes | Store表示用の商品名。 |
| `short_title` | no | 短縮名。カード表示や一覧で使う。 |
| `catch` | yes | 商品カードや詳細冒頭で使う短い説明。 |
| `description` | yes | 商品詳細本文。 |
| `price` | yes | 表示価格。未確定なら `null`。 |
| `currency` | yes | 通貨。初期値は `JPY`。 |
| `status` | yes | `available` などの状態。 |
| `category` | yes | PDF、画像、メモ、Packなどの分類。 |
| `tags` | no | 表示・検索用タグ。 |
| `image` | no | 詳細ページ向け画像。未確定なら `null`。 |
| `thumbnail` | no | 一覧カード向け画像。未確定なら `null`。 |
| `download_type` | no | `github_release`, `booth`, `r2`, `store_purchase`, `external`, `none` など。 |
| `download_url` | no | 取得導線。未確定なら `null`。 |
| `booth_url` | no | BOOTH商品URL。未登録なら `null`。 |
| `github_release_url` | no | GitHub Release URL。未作成なら `null`。 |
| `store_url` | no | Store内URL。生成前なら `null`。 |
| `support_policy` | no | サポート方針。未確定なら `null`。 |
| `disclaimer` | no | 注意事項・免責。 |
| `included_items` | no | Packなどの構成物。単品アプリでは空配列でよい。 |
| `stripe_payment_link` | no | Stripe Payment Link。未設定なら `null`。SecretやAPIキーは絶対に含めない。 |
| `payment_status` | yes | Storeの購入導線状態。`stripe_ready` / `booth_only` / `preparing` / `free_download` / `not_for_sale`。 |
| `stripe_price_id` | no | Stripe接続用の任意項目。MVPでは `null`。 |
| `source_kind` | yes | 原則 `original`。 |
| `is_generated` | yes | 常に `true`。 |

## type定義

```text
app
= 単品アプリ

pack
= DAKE PackなどのPack商品

shimarisu_pack
= SHIMARISU Pack
```

将来追加できる候補。

```text
bundle
template
note
asset
```

追加時も、正本がどこにあるかを必ず明記する。

## 未確定値の扱い

以下は未確定になり得る。

- Store URL
- Storeダウンロード導線
- Storeサポート方針
- Stripe Payment Link
- Stripe Price ID
- Store用画像
- `download_url`

未確定値は `null` を推奨する。

```json
{
  "download_url": null,
  "store_url": null,
  "support_policy": null
}
```

人間向け表示では、Store側で `null` を以下のような表示へ変換してよい。

- 準備中
- BOOTHで見る
- GitHub Releaseで見る
- 近日追加

`"未確定"` 文字列は、人間用Markdownでは許容するが、generated JSONでは機械処理しやすい `null` を優先する。

## Stripe情報の扱い

Stripe Secret は絶対に含めない。

Stripe Price ID はMVP仕様では必須にしない。

決済実装Phaseで必要になった場合のみ、generated 側に任意項目として持たせる。

```json
{
  "stripe_price_id": null
}
```

`ORIGINAL.md` に Stripe 内部IDを混ぜることは避ける。

Stripe接続情報は、Store実装側の環境変数、設定、または別の安全な接続層で扱う。

### Stripe Payment Link

Stripe Payment Link は、Store静的MVPで使える販売導線として generated JSON に任意項目で持たせる。

```json
{
  "stripe_payment_link": null,
  "payment_status": "booth_only"
}
```

`stripe_payment_link` が `null` の場合、Store側はStripe購入ボタンを出してはいけない。

Stripe Secret Key、Webhook Secret、APIキーは、generated JSON、静的HTML、JavaScript、`ORIGINAL.md` に含めない。

### payment_status

`payment_status` は、Store UI の購入導線を安全に出し分けるための状態値である。

```text
stripe_ready
= Stripe Payment Link があり、StoreからStripe購入へ進める。

booth_only
= Stripe Payment Link は未設定だが、BOOTH URL がある。

preparing
= Stripe Payment Link も BOOTH URL も未設定で、販売準備中。

free_download
= 無料配布想定。

not_for_sale
= Store販売対象外。
```

MVPの自動判定。

```text
stripe_payment_link あり
→ stripe_ready

stripe_payment_link なし + booth_url あり
→ booth_only

stripe_payment_link なし + booth_url なし
→ preparing
```


## download_urlの扱い

MVPでは `download_url` は未確定でよい。

想定候補。

- GitHub Release
- BOOTH
- Cloudflare R2
- Store購入後URL

`ORIGINAL.md` の Store表示用情報でダウンロード導線が未確定の場合、generated 側も `null` にする。

Store表示では、`download_type` と `download_url` の組み合わせで導線を決める。

例:

```json
{
  "download_type": "booth",
  "download_url": "https://peakheadz.booth.pm/items/0000000"
}
```

`download_url` が `null` の場合、Store側は購入ボタンやダウンロードボタンを出さず、準備中または外部導線を表示する。

## 画像の扱い

画像はまず既存の出荷素材を参照する。

候補。

- `assets/booth_thumbnail.jpg`
- `assets/screenshot.webp`
- `assets/screenshot.jpg`
- `booth_ready/booth_thumbnail.jpg`
- `pack_ready/images`
- `SHIMARISU/booth_ready/images`

初期Storeでは、一覧カード向けに `thumbnail`、詳細ページ向けに `image` を分ける。

Store専用画像が未確定の場合は、以下の優先順位を候補とする。

1. `assets/booth_thumbnail.jpg`
2. `booth_ready/booth_thumbnail.jpg`
3. `assets/screenshot.webp`
4. `assets/screenshot.jpg`
5. PackまたはSHIMARISU側の画像
6. `null`

ただし、実際の生成スクリプトでパスの存在確認を行う。

## SHIMARISU Packの扱い

SHIMARISU Pack は DAKE_series 外に真の正本がある。

```json
{
  "type": "shimarisu_pack",
  "source_repo": "SHIMARISU",
  "source_original": "C:/Users/yukiz/devlop/SHIMARISU/ORIGINAL.md"
}
```

`SHIMARISU/ORIGINAL.md` はSHIMARISU Pack商品の真の正本である。

`shimarisu-pack-release` は正本ではなく、配布repo / 公開ビューとして扱う。

Store生成スクリプトで絶対パスを使うか、設定ファイルで外部正本の場所を指定するかは次Phaseで決める。

Store generated JSONには、SHIMARISU Packも `items` の1件として含められる構造にする。

## 手編集禁止ルール

`store_products.generated.json` は手編集禁止。

ファイル冒頭または同階層のREADMEに、以下の趣旨を必ず残す。

```text
This file is generated from ORIGINAL.md.
Do not edit manually.
```

Store表示を直したい場合は、まず該当商品の `ORIGINAL.md` を直す。

generated だけを直しても、次回生成で消える。

## Store実装Phaseで決めること

次Phase以降で決めること。

- `store_products.generated.json` の出力先
- 生成スクリプト名
- DAKE_series外の `SHIMARISU/ORIGINAL.md` 参照方法
- 画像ファイルのコピー先・公開URL
- `download_type` の正式値
- Stripe Price IDをどこで管理するか
- Store内URL設計
- Cloudflare Pages / R2 / Functions の役割
- Store公開前の検証チェック

## 今回やらないこと

今回のPhase 4-6では以下を行わない。

- `store_products.generated.json` の実生成
- 生成スクリプト作成
- `store.dakeapp.com` 構築
- Stripe実装
- Cloudflare Pages設定
- 各 `ORIGINAL.md` 更新
- README再生成
- booth_product更新
- pack_ready更新
- SHIMARISU更新

今回は、Store用 generated データの仕様定義のみを行う。
