# DAKE Release Definition v2 With Store

## 目的

DAKE正式出荷の完了条件に、store.dakeapp.comへの掲載・同期・本番確認を追加する。

## 正本と派生ビュー

- `ORIGINAL.md`: 真の正本。
- `README.md`: GitHub公開用ビュー。
- `DAKE_META`: Launcher / サイト / 管理ツール向け機械利用ビュー。
- `release_body.md`: GitHub Release用ビュー。
- `booth_product.txt`: BOOTH登録用ビュー。
- `store_products.generated.json`: Store表示用の生成データ。
- `store.dakeapp.com`: 自社Storeの販売ビュー。

Storeは正本ではない。ただし、Store公開はDAKE正式出荷の一部である。

## DAKE正式出荷定義 v2

DAKE正式出荷は今後以下を含む。

1. `ORIGINAL.md` 更新。
2. `README.md` 更新または整合確認。
3. `DAKE_META` 更新または整合確認。
4. `release_body.md` 更新。
5. `booth_product.txt` / `booth_ready` 更新。
6. GitHub Release 作成・確認。
7. BOOTH登録またはBOOTH導線確認。
8. dakeapp.com 掲載確認。
9. `store_products.generated.json` 再生成。
10. `dake-store-site` へ同期。
11. store.dakeapp.com 商品詳細確認。
12. Stripe / BOOTH / 準備中の `payment_status` 確認。

## Store同期手順

Store反映時は以下を使う。

```powershell
cd C:\Users\yukiz\devlop\DAKE_series
python tools\store\sync_store_to_site.py
```

同期スクリプトでは以下を確認する。

- items件数
- type別件数
- `payment_status` 件数
- `stripe_ready` 件数
- `booth_only` 件数
- `preparing` 件数
- `source_policy`
- `do_not_edit: true`
- `shimarisu_pack`

## payment_statusの扱い

- `stripe_ready`: Stripe Payment Linkあり。StoreでStripe購入導線を表示できる。
- `booth_only`: Stripe Payment Linkは未設定で、BOOTH導線がある。
- `preparing`: Stripe Payment LinkもBOOTH URLも未設定で準備中。
- `not_for_sale`: 販売対象外。

全商品にStripe Payment Linkを必須にはしない。

Stripe Secret、APIキー、Webhook Secretは、公開repo、generated JSON、Store JavaScriptへ入れない。

## 今後のCodex完了報告に含める項目

- `ORIGINAL.md` 更新有無
- README / release_body / booth_product 整合確認
- GitHub Release URL
- BOOTH URL
- dakeapp.com 掲載URL
- store.dakeapp.com 商品詳細URL
- `payment_status`
- Stripe Payment Link有無
- BOOTH導線有無
- `store_products.generated.json` 再生成結果
- `dake-store-site` 同期結果
- Cloudflare Pages反映確認
- final git status

## 今回やらなかったこと

- Store UI修正
- Stripe Payment Link追加
- Stripe API自動登録
- Webhook実装
- R2実装
- 商品情報変更
- 全アプリのORIGINAL再更新
- プロジェクト情報源そのものの更新

## 更新した情報源

- `00_core/DAKE_RELEASE_FLOW.md`
- `00_core/DAKE_COMMON_SPEC.md`
- `00_core/DAKE_REVIEW_CHECKLIST.md`
- `00_core/CHATGPT_CODEX_WORKFLOW.md`
- `00_core/DAKE_STORE_GENERATED_SPEC.md`

## 結論

DAKE正式出荷は、GitHub Release、BOOTH、dakeapp.comだけでなく、store.dakeapp.com掲載・同期・本番確認まで含めて扱う。

Storeは正本ではなく、`ORIGINAL.md` 由来の販売ビューである。
