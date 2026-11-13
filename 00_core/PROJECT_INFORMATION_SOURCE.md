# DAKE Project Information Source

## 目的

このファイルは、DAKEプロジェクト全体の現在の正式運用状態をまとめる情報源である。

思想を増やすための文書ではなく、Phase 1〜13で確定した運用前提を、未来の自分と未来のCodexが迷わないように固定するための文書である。

## 現在の正式運用状態

DAKEシリーズは、`ORIGINAL.md` を真の正本とする運用へ移行している。

README、DAKE_META、release_body、booth_product、Store表示は派生ビューである。

Storeは公開済みの販売ビューであり、DISTRIBUTION / COMMERCEとしてSHIPと独立して扱う。

## 正本

DAKEシリーズの真の正本は `ORIGINAL.md` である。

`ORIGINAL.md` は、対象アプリ、Pack、商品、販売表示の最後に信じる場所である。

派生ビューとの矛盾は `ORIGINAL.md` を優先する。シリーズ共通必須規則は `DAKE_COMMON_SPEC.md` が正本であり、個別ORIGINALで緩和しない。

共通仕様は新規アプリと次回更新対象へ段階適用する。作業入口はrepoルートの `AGENTS.md`。既存成果物の改修・実機検証・再出荷は一括実施せず、旧Releaseの新仕様適合を保証するものではない。

### 改定時点の実装差分（未改修・read-only確認）

- `DAKE_PDF_SplitOne/build.bat` は `--onedir`。SplitOne / SplitSelectの `main.py` は旧クレジットのまま。次回対象改修で配布方式とUI_TEXTを共通規則へ合わせる。
- `DAKE_PDF_SplitSelect/build.bat` は既に `--onefile`、ORIGINALにはv1.0.2の単独配布記録がある。今回は再ビルド・実機再検証していない。
- `tools/generate_version_info.py` はCompanyName=`PEAKHEADZ`、作者=`Yukihiko Kikuta`、LegalCopyright=`© 2026 Yukihiko Kikuta` のWindowsメタデータをbuild時に自動生成する。UIフッターとは分けて扱う。
- 出荷ツールのonedir対応や既存Releaseは互換・履歴として維持している。全アプリの実装・成果物を適合監査したわけではなく、上記以外も次回対象作業で調べる。

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

ただし、DISTRIBUTIONの販売ビューの一つである。

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

- SHIP: ORIGINAL確認、build、生成exeの主要動作、必要なGit作業、GitHub Release。GitHub Release完了で `SHIP = COMPLETE`。
- DISTRIBUTION: BOOTH、dakeapp.com、store.dakeapp.com。`COMPLETE / PARTIAL / PENDING`。
- COMMERCE: Stripe、Payment Link、`payment_status`、checkout。`READY / PENDING / NOT_REQUIRED`。

DISTRIBUTION / COMMERCEの未完了はSHIPをPENDINGへ戻さない。「正式出荷」は全レイヤー完了の総称として使わない。

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

通常作業では、repoルートの `AGENTS.md`、`DAKE_COMMON_SPEC.md`、対象の `ORIGINAL.md` の3つだけを最初に読む。

README、DAKE_META、release_body、booth_product、Storeは派生ビューであり、変更が影響する時だけ確認・更新する。`ORIGINAL.md` が存在しない既存対象は、次回その対象を修正・出荷する時だけREADMEを暫定参照して段階移行する。

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
