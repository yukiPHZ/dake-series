# Codex Workflow v2 Update

## 更新理由

DAKEシリーズは `ORIGINAL.md` を真の正本として運用する段階へ移行しました。Store公開、Stripe運用、Store同期フロー、DAKE出荷定義v2まで完了したため、今後CodexがREADME正本時代の前提で作業しないよう、標準作業ルールを更新しました。

## ORIGINAL.md優先ルール

`ORIGINAL.md` が存在する対象では、`ORIGINAL.md` を最優先で確認します。README、DAKE_META、release_body、booth_product、Store表示は派生ビューであり、正本ではありません。

矛盾がある場合の優先順位は次の通りです。

```text
ORIGINAL.md
↓
README.md
↓
DAKE_META
↓
release_body.md
↓
booth_product.txt
```

READMEにしか存在しない重要情報を見つけた場合は、その情報を `ORIGINAL.md` へ戻すべきか確認します。

## 作業開始時確認順

Codexは対象アプリ・サイト・Packの作業前に、原則として次の順で確認します。

1. `ORIGINAL.md`
2. `README.md`
3. `DAKE_META`
4. `release_body.md`
5. `booth_product.txt`
6. 関連仕様ファイル

存在しない場合のみ次へ進みます。

## 出荷確認順

DAKE正式出荷では、次の流れを標準確認順とします。

```text
ORIGINAL.md確認
↓
README整合確認
↓
release_body整合確認
↓
booth_product整合確認
↓
Store generated JSON更新
↓
Store同期
↓
GitHub Release
↓
BOOTH
↓
dakeapp.com
↓
store.dakeapp.com
↓
出荷完了
```

## Store確認項目

Storeは正本ではなく、`ORIGINAL.md` 由来の generated JSON を読む販売ビューです。出荷確認では次を標準項目にします。

- `store_products.generated.json` 更新確認
- Store同期確認
- Store商品詳細URL確認
- `payment_status` 確認
- Stripe導線確認
- BOOTH導線確認
- Cloudflare反映確認

Store側で商品名、価格、説明文、Stripe Payment Link、BOOTH URLを直接編集しません。

## 完了報告追加項目

今後、正式出荷・Store反映が関係するCodex完了報告には次を含めます。

- `ORIGINAL.md` 確認有無
- README / release_body / booth_product 整合確認
- Store同期結果
- Store商品詳細URL
- `payment_status`
- Stripe導線
- BOOTH導線
- Cloudflare反映確認
- final git status

## 過渡期ルール

`ORIGINAL.md` がまだ存在しない既存アプリでは、移行完了まで `README.md` を暫定参照してよいものとします。ただし、今後の移行対象であれば `ORIGINAL.md` 作成を提案します。

## 今回やらなかったこと

- Store UI修正
- Stripe追加
- Webhook実装
- R2実装
- DAKE出荷定義の再変更
- 情報源更新
- 全アプリ再レビュー
