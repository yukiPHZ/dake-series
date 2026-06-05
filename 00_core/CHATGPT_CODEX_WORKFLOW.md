# ChatGPT / Codex Workflow

## 基本思想

チャットは記憶庫ではなく、作業机として扱う。

完了した作業は、ORIGINAL.md・README・Git・レポート・成果物に反映し、チャット自体は必要に応じてアーカイブする。

## 役割

### ChatGPT

- 思考
- 相談
- 設計
- 判断
- Codex指示書作成
- 作業方針整理

### Codex

- 実装
- 修正
- 調査
- 棚卸し
- commit
- push
- レポート作成

### ORIGINAL.md / README

- `ORIGINAL.md`: 現在の仕様、思想、状態をまとめる真の正本
- `README.md`: GitHub公開用ビュー
- `DAKE_META`: Launcher / サイト / 管理ツール向けの機械利用ビュー
- `release_body.md` / `booth_product.txt`: Release / BOOTH向けの派生ビュー
- `store_products.generated.json` / `store.dakeapp.com`: Store向けの生成データと販売ビュー

### Git

- 変更履歴
- 実装履歴
- 復元可能性
- commit単位の記録

### current_status_review

現在地を見失った時に作成する棚卸しレポート。

使うタイミング:

- 作業が多方面に広がった時
- 何が終わっているかわからなくなった時
- 大きな出荷作業後
- 新しいフェーズへ移る前

保存先:

```text
tools/reports/current_status_review_YYYYMMDD.md
```

## 新チャット運用

原則:

1テーマ1チャットを推奨する。

ただし、横断作業も可能。
その場合は、毎回以下を明示する。

- 対象repo
- 対象フォルダ
- 作業目的
- やってよいこと
- やってはいけないこと
- commit / push の有無



## Store反映作業

DAKE正式出荷でStoreへ反映する場合、Codexは `DAKE_series` で以下を実行して同期する。

```powershell
python tools\store\sync_store_to_site.py
```

Storeは正本ではない。商品情報をStore側で手編集せず、必要な変更は各商品の `ORIGINAL.md` へ戻す。

完了報告では、Store商品詳細URL、`payment_status`、Stripe Payment Link有無、BOOTH導線有無、dake-store-site同期結果、Cloudflare Pages反映確認を報告する。

## アーカイブ基準

以下を満たしたチャットはアーカイブ候補。

- ORIGINAL.mdまたはREADMEに反映済み
- Git commit済み
- push済み
- 必要なレポートが保存済み
- 次にやることが別テーマになった

## Codex作業前の確認

Codexは作業前に必要に応じて以下を確認する。

- ORIGINAL.md
- README.md
- DAKE_META
- release_body.md
- booth_product.txt
- store_products.generated.json
- tools/reports/current_status_review_*.md
- git status
- git log --oneline -5

## 禁止事項

- チャットだけを正本にしない
- 完了報告だけで終わらせない
- ORIGINAL.md、README、Gitに反映しないまま記憶に頼らない
- 対象repo不明のまま作業しない
- git add . を安易に使わない

## 運用まとめ

ChatGPTは考える場所。
Codexは動かす場所。
ORIGINAL.mdは真の正本。
READMEはGitHub公開用ビュー。
Gitは履歴。
status_reviewは現在地。
チャットは作業机。

終わったら片付ける。

## Codex標準作業ルール v2

DAKEシリーズでは、`ORIGINAL.md` を真の正本として扱います。CodexはREADME正本時代の前提で作業せず、対象アプリ・サイト・Packの作業前に、原則として次の順で確認します。

1. `ORIGINAL.md`
2. `README.md`
3. `DAKE_META`
4. `release_body.md`
5. `booth_product.txt`
6. 関連仕様ファイル

存在しないファイルがある場合のみ次へ進みます。`ORIGINAL.md` が存在する対象では、矛盾があれば必ず `ORIGINAL.md` を優先します。

優先順位は次の通りです。

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

`README.md` はGitHub公開ビューであり、真の正本ではありません。READMEにしか存在しない重要情報を見つけた場合は、その情報を `ORIGINAL.md` へ戻すべきか確認します。

派生ビューの位置づけは次の通りです。

- `README.md` = GitHub公開ビュー
- `DAKE_META` = Launcher / サイト / 管理ツール向け機械利用ビュー
- `release_body.md` = GitHub Releaseビュー
- `booth_product.txt` = BOOTH登録ビュー
- `store.dakeapp.com` = 販売ビュー

上記はすべて派生物であり、正本ではありません。Store側で商品名、価格、説明文、Stripe Payment Link、BOOTH URLを直接編集しません。

正式出荷時は、次の順で確認します。

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

Store確認では、次を標準項目とします。

- `store_products.generated.json` 更新確認
- Store同期確認
- Store商品詳細URL確認
- `payment_status` 確認
- Stripe導線確認
- BOOTH導線確認
- Cloudflare反映確認

Codexの完了報告には、正式出荷・Store反映が関係する場合、次を含めます。

- `ORIGINAL.md` 確認有無
- README / release_body / booth_product 整合確認
- Store同期結果
- Store商品詳細URL
- `payment_status`
- Stripe導線
- BOOTH導線
- Cloudflare反映確認
- final git status

過渡期ルールとして、`ORIGINAL.md` がまだ存在しない既存アプリでは `README.md` を暫定参照してよいものとします。ただし、今後の移行対象であれば `ORIGINAL.md` 作成を提案します。
