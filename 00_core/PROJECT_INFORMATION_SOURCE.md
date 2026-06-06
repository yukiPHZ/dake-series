# DAKE Project Information Source

## 目的

このファイルは、DAKEプロジェクト全体の現在の正式運用状態をまとめる情報源である。

思想を増やすための文書ではなく、Phase 1〜13で確定した運用前提を、未来の自分と未来のCodexが迷わないように固定するための文書である。

## 現在の正式運用状態

DAKEシリーズは、`ORIGINAL.md` を真の正本とする運用へ移行している。

README、DAKE_META、release_body、booth_product、Store表示は派生ビューである。

Storeは公開済みであり、Stripe Payment Link、BOOTH導線、Store同期フローを含む正式出荷先として扱う。

## 正本

DAKEシリーズの真の正本は `ORIGINAL.md` である。

`ORIGINAL.md` は、対象アプリ、Pack、商品、販売表示の最後に信じる場所である。

矛盾がある場合は `ORIGINAL.md` を優先する。

## 派生ビュー

次のものは派生ビューであり、正本ではない。

- `README.md`: GitHub公開ビュー
- `DAKE_META`: Launcher / サイト / 管理ツール向け機械利用ビュー
- `release_body.md`: GitHub Releaseビュー
- `booth_product.txt`: BOOTH登録ビュー
- Store表示: 販売ビュー

派生ビュー側にしかない重要情報を見つけた場合は、`ORIGINAL.md` へ戻すべき情報か確認する。

## generated

generatedファイルは正本ではない。

`store_products.generated.json` は `ORIGINAL.md` 由来のStore表示用生成データである。

手編集しない。修正する場合は正本へ戻る。

## Store

`store.dakeapp.com` は販売ビューである。

正本ではない。

ただし、DAKE正式出荷先の一つである。

Store側で商品名、価格、説明文、Stripe Payment Link、BOOTH URLを直接編集しない。

```text
ORIGINAL.md
↓
store_products.generated.json
↓
dake-store-site
↓
store.dakeapp.com
```

## 出荷定義

DAKE正式出荷には、次を含める。

- GitHub Release
- BOOTH
- dakeapp.com
- store.dakeapp.com

GitHub Releaseだけでは正式出荷完了としない。

正本から派生ビューが生成または整合され、必要な出荷先へ届いている状態を正式出荷とする。

## Store同期

Store同期は次のスクリプトを使用する。

```powershell
python tools/store/sync_store_to_site.py
```

この同期では、少なくとも次を確認する。

- `store_products.generated.json` の再生成
- dake-store-siteへの同期
- items件数
- type別件数
- `payment_status` 件数
- `source_policy`
- `do_not_edit: true`
- `shimarisu_pack` の存在

## payment_status

Storeの購入導線状態は `payment_status` で管理する。

正式値は次の通り。

- `stripe_ready`: Stripe Payment Linkあり
- `booth_only`: BOOTH導線のみ
- `preparing`: 準備中
- `not_for_sale`: 販売対象外

Store確認では、商品詳細URLと `payment_status` を必ず確認する。

## Stripe

Stripe Payment Linkは利用できる。

ただし、Stripeは正本ではない。

次の情報は公開repo、generated JSON、Store JavaScriptへ保存しない。

- Stripe Secret
- APIキー
- Webhook Secret

Stripe Payment Linkを追加・変更する場合は、対象商品の `ORIGINAL.md` に戻し、generated JSONを再生成してStoreへ同期する。

## Codex標準確認順

Codexは対象アプリ、サイト、Packの作業前に、原則として次の順で確認する。

1. `ORIGINAL.md`
2. `README.md`
3. `DAKE_META`
4. `release_body.md`
5. `booth_product.txt`

`ORIGINAL.md` が存在しない既存対象では、移行完了までREADMEを暫定参照してよい。

ただし、今後の移行対象であれば `ORIGINAL.md` 作成を提案する。

## SHIMARISU

SHIMARISU Packの正本は次に置く。

```text
SHIMARISU/ORIGINAL.md
```

SHIMARISU Packの正本をDAKE_series内へ混ぜない。

`shimarisu-pack-release` は正本ではなく、公開・配布ビューとして扱う。

## 主要参照ファイル

- `00_core/DAKE_SOURCE_OF_TRUTH_ENGINE_V02.md`
- `00_core/DAKE_ORIGINAL_RULE.md`
- `00_core/DAKE_RELEASE_FLOW.md`
- `00_core/CHATGPT_CODEX_WORKFLOW.md`
- `00_core/DAKE_STORE_GENERATED_SPEC.md`
- `00_core/DAKE_STORE_OPERATION_RULE.md`

## 今回やらないこと

この情報源更新では、次を行わない。

- Store UI修正
- Stripe追加
- Webhook実装
- R2実装
- 商品変更
- README再生成
- 全アプリ再レビュー
