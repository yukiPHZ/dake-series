# DAKE Store Operation Rule

## 目的

DAKE Store の公開条件、購入後導線、サポート、返金の運用方針を定義する。

StoreはStripeで購入可能な商品だけを静かに置く販売棚です。BOOTHは独立した別販売経路です。

## Storeの位置づけ

DAKE Store は正本ではありません。

Store は、各商品・各アプリの `ORIGINAL.md` から生成された `store_products.generated.json` を読む販売ビューです。

Store掲載はDISTRIBUTIONです。BOOTH登録の完了やStripe準備を掲載そのものの前提条件にしません。BOOTH URLとStripe Payment Linkは、各商品で利用可能な購入導線として独立に扱います。

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

## 決済・自動配信方式（現行正式運用）

公開Storeの決済基盤は **Stripe Payment Link**、購入後配信基盤は **Cloudflare Pages Functions + プライベート Cloudflare R2（`dake-desktop-releases`）** です。

```text
Stripe Payment Link 決済完了
↓
https://store.dakeapp.com/download/?session_id={CHECKOUT_SESSION_ID}
↓
Cloudflare Pages Functions（/api/download-status, /api/download）
  - STRIPE_SECRET_KEY（Cloudflare Secret）で Stripe Checkout Session / Payment Link を検証
  - livemode=true, payment_status=paid, purchase_delivery_method=stripe_session_private_r2 を確認
  - DESKTOP_DELIVERIES ホワイトリスト（functions/_lib/private-download.mjs）照合
↓
プライベート R2 バケット dake-desktop-releases（バインディング: DAKE_DESKTOP_RELEASES）
  - desktop/<id>/current.zip をダウンロード配信
```

- 無料配布枠は `DakePDF結合`（`DAKE_PDF_Merge`）のみとし、GitHub Release に公開アセットを残しつつ、Store には同機能の応援購入版（500円・R2自動配信）を併設する。
- 有料アプリおよび Pack は Store（Stripe Payment Link → R2 `desktop/<id>/current.zip` 自動配信）＋ BOOTH 並行販売を正本とし、GitHub Releases には公開アセット（`.exe` / `.zip`）を置かずタグ・履歴テキストのみ残す。
- Stripe Secret Key、Webhook Secret、APIキーは、静的Store、generated JSON、ORIGINAL.md、wrangler.toml に保存しない（`STRIPE_SECRET_KEY` は Cloudflare Secret のみ）。

## Stripe Payment Linkの扱い

`payment_status == stripe_ready` かつ `stripe_payment_link` がある商品だけをStoreへ公開し、購入ボタン「購入する」のリンク先に使います。

Stripe Payment Link は販売導線であり、商品正本ではありません。

Payment LinkのURLを追加・変更する場合は、対象商品の `ORIGINAL.md` に戻し、`store_products.generated.json` を再生成します。

StripeはCOMMERCEです。Store同期からProduct、Price、Payment Link、live API、checkout確認、`payment_status` 変更へ自動的に進みません。

## BOOTH導線の扱い

BOOTHは独立した別販売経路です。`booth_url` があってもStripe-readyでない商品をStoreへ公開せず、Store上の購入代替ボタンにも使いません。

BOOTH側の商品説明やURLも、可能な範囲で各商品の `ORIGINAL.md` に戻します。

BOOTH URLが未設定でもStore掲載データは生成できます。その場合は購入導線を「準備中」等として表示し、Store全体をPENDINGにはしません。

## 同期ツールの境界

`tools/store/sync_store_to_site.py` はgenerated JSON生成、Store repoへの同期、検証、次のGit手順表示までを担当します。自動commit、自動push、Cloudflare deployは行いません。公開反映は明示依頼がある場合だけ別工程で実施します。

## 購入後導線・ダウンロード導線

Stripe Payment Link の決済完了後は、自動的に DAKE Store のダウンロード専用ページへ遷移します。

```text
https://store.dakeapp.com/download/?session_id={CHECKOUT_SESSION_ID}
```

- サーバー側（Cloudflare Pages Functions）で Stripe Checkout Session と Payment Link のメタデータ（`dake_item_id`、`purchase_delivery_method=stripe_session_private_r2`）を検証し、プライベート R2 バケット `dake-desktop-releases` の `desktop/<id>/current.zip` を直接配信します。
- 有料商品の公開 `download_url` や GitHub Release 公開アセットは発行・掲載しません。
- BOOTH は独立した並行販売経路として継続し、Store 上の購入代替ボタンとしては表示しません。
- 決済完了画面を閉じた・再ダウンロードが必要になった等の個別対応は、問い合わせ窓口で決済情報を確認のうえ対応します。

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

今後のPhaseで必要に応じて以下を検討します。

- Webhook 連携や購入履歴・再ダウンロード案内メールの自動化
- 商品ごとのライセンス・法人利用範囲表示の整備
- Store専用サポート導線の強化

## 運用上の禁止事項（やらないこと）

- 静的ファイル・generated JSON・Git・ORIGINAL.md への Stripe Secret Key / APIキー / 購入者個人情報の保存
- 有料アプリ・Pack の `.exe` / `.zip` の GitHub Release 公開アセット配置（無料配布枠 `DakePDF結合` 除く）
- `store_products.generated.json` の手編集
- 人間の明示依頼なしの Stripe live 決済作成・価格変更・公開操作
