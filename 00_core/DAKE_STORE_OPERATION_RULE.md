# DAKE Store Operation Rule

## 目的

DAKE Store の購入後導線、ダウンロード案内、サポート、返金、BOOTH併用の運用方針を定義する。

このルールは、静的 Store + Stripe Payment Link + BOOTH 併用の MVP 運用を安全に続けるためのものです。

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

## 正本と販売ビューの関係

- `ORIGINAL.md`: 真の正本
- `store_products.generated.json`: Store表示用の生成データ。手編集禁止
- `store.dakeapp.com`: 販売ビュー
- Stripe Payment Link: 決済導線
- BOOTH: 併用する販売・配布導線

Store側だけに商品正本を作りません。

## 決済方式

MVPでは Stripe Payment Link を優先します。

Stripe Checkout API、Pages Functions、Webhook、Cloudflare R2、購入後自動URL発行は後続Phaseで検討します。

Stripe Secret Key、Webhook Secret、APIキーは、静的Store、generated JSON、ORIGINAL.mdに保存しません。

## Stripe Payment Linkの扱い

`stripe_payment_link` がある商品は、Store上で「Stripeで購入」を表示します。

Stripe Payment Link は販売導線であり、商品正本ではありません。

Payment LinkのURLを追加・変更する場合は、対象商品の `ORIGINAL.md` に戻し、`store_products.generated.json` を再生成します。

## BOOTH導線の扱い

`stripe_payment_link` が未設定で `booth_url` がある商品は、Store上で「BOOTHで見る」を表示します。

BOOTHは、DAKE Store MVP期間中の販売・配布導線として併用します。

BOOTH側の商品説明やURLも、可能な範囲で各商品の `ORIGINAL.md` に戻します。

## 購入後導線

MVPでは、Stripe Payment Link購入後の自動ダウンロード発行は未実装です。

購入後の戻り先は、商品ごとの詳細ページを基本とします。

例:

```text
https://store.dakeapp.com/product/?id=DAKE_Time_AdvancedTimer
```

購入者には、商品詳細、GitHub Release、BOOTH、または商品ごとの案内に従ってもらいます。

## ダウンロード導線

MVPでは、以下を優先します。

1. GitHub Release
2. BOOTH
3. Store購入後URL（未実装）

`download_url` が未確定の商品は、`null` のまま扱います。

表示側では、必要に応じて「準備中」「BOOTHで見る」「GitHub Release」などへ変換します。

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

## download_url未確定商品の扱い

`download_url` が未確定でも、BOOTH URL、GitHub Release URL、Stripe Payment Linkがある場合は販売・案内を継続できます。

ただし、購入者が迷わないように、商品詳細ページや販売文で現在の配布導線を明確にします。

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