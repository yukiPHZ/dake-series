# DAKE Store Operation Rule

## 目的

DAKE Store の公開条件、購入後導線、サポート、返金の運用方針を定義する。

StoreはStripeで購入可能な商品だけを静かに置く販売棚です。BOOTHは独立した別販売経路です。

## Storeの位置づけ

DAKE Store は正本ではありません。

Store は、各商品・各アプリの `ORIGINAL.md` から生成された `store_products.generated.json` を読む販売ビューです。

```text
ORIGINAL.md
↓
store_products.generated.json
↓
store.dakeapp.com
```

商品情報、価格、販売文、注意事項、サポート方針などの正本化すべき情報は、必要に応じて `ORIGINAL.md` または共通ルールへ戻します。

公開Storeの一覧・商品詳細に表示できる条件は次の3つすべてです。

```text
status == available
AND payment_status == stripe_ready
AND stripe_payment_link が存在する
```

generated JSONに存在すること自体は公開条件ではありません。`booth_only`、`preparing`、`free_download`、`not_for_sale` は公開Storeへ表示しません。

## 正本と販売ビューの関係

- `ORIGINAL.md`: 真の正本
- `store_products.generated.json`: Store表示用の生成データ。手編集禁止
- `store.dakeapp.com`: 販売ビュー
- Stripe Payment Link: 決済導線
- BOOTH: Storeとは独立した別販売経路

Store側だけに商品正本を作りません。

## 決済方式

公開Storeの決済基盤はStripe Payment Linkです。

Stripe Checkout API、Pages Functions、Webhook、Cloudflare R2、購入後自動URL発行は後続Phaseで検討します。

Stripe Secret Key、Webhook Secret、APIキーは、静的Store、generated JSON、ORIGINAL.mdに保存しません。

## Stripe Payment Linkの扱い

`payment_status == stripe_ready` かつ `stripe_payment_link` がある商品だけをStoreへ公開し、購入ボタン「購入する」のリンク先に使います。

Stripe Payment Link は販売導線であり、商品正本ではありません。

Payment LinkのURLを追加・変更する場合は、対象商品の `ORIGINAL.md` に戻し、`store_products.generated.json` を再生成します。

## BOOTH導線の扱い

BOOTHは独立した別販売経路です。`booth_url` があってもStripe-readyでない商品をStoreへ公開せず、Store上の購入代替ボタンにも使いません。

BOOTH側の商品説明やURLも、可能な範囲で各商品の `ORIGINAL.md` に戻します。

## 購入後導線

MVPでは、Stripe Payment Link購入後の自動ダウンロード発行は未実装です。

購入後の戻り先は、商品ごとの詳細ページを基本とします。

例:

```text
https://store.dakeapp.com/product/?id=DAKE_Time_AdvancedTimer
```

購入者には、Stripe決済または商品固有の購入後案内に従ってもらいます。

## ダウンロード導線

`download_url`、GitHub Release、BOOTH URLはStore掲載条件でも購入代替導線でもありません。購入後downloadの仕組みは今回変更しません。

## 問い合わせ先

問い合わせ導線は、DAKEシリーズ公式サイト、BOOTH、GitHub Release、または各商品ページの案内に従います。

MVPでは、Store内に専用問い合わせフォームを置きません。

## サポート範囲

不具合報告、ダウンロード不備、ファイル不備には可能な範囲で対応します。

個別環境での完全な動作保証、業務結果の保証、個別業務判断の代行は行いません。

## 返金方針

デジタル商品の性質上、購入後の返金は原則として個別確認とします。

重複購入、明らかな誤購入、ファイル不備、ダウンロード不備などがある場合はお問い合わせください。

StripeまたはBOOTHの仕組みに従う必要がある場合は、各サービス側の手順を優先します。

## 再配布禁止

購入・ダウンロードしたファイルの無断転載、無断再配布、再販売を禁止します。

チーム内利用、社内利用などの扱いは、商品ごとの説明または今後のライセンス方針で整理します。

## 免責

各ツールは作業補助を目的としたものです。

重要なファイルは事前にバックアップしてください。

税務・法務・不動産判断に関わる内容は、必要に応じて専門家へ確認してください。

## レイヤー分離

Stripe Payment Linkの作成と `payment_status = stripe_ready` への更新はCOMMERCEです。generated JSONの生成、Store repoへの同期、公開確認はDISTRIBUTIONです。

```text
COMMERCE
↓
DISTRIBUTION
```

`sync_store_to_site.py` はgenerated JSON生成・同期・検証・次のGit手順表示だけを担当し、Stripe APIを呼びません。

## 今後の拡張

後続Phaseで以下を検討します。

- Stripe Checkout API
- Cloudflare Pages Functions
- Webhook
- Cloudflare R2
- 購入後ダウンロードURL発行
- 購入完了ページ
- 商品ごとのライセンス表示
- Store専用サポート導線

## 今回やらないこと

- Stripe Checkout API実装
- Pages Functions実装
- Webhook実装
- Cloudflare R2連携
- 購入後自動ダウンロードURL発行
- 全商品Stripe化
- Payment Link追加
- 商品価格変更
- Storeデザイン大幅変更
