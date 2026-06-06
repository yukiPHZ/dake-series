# Project Information Source Update

## 更新内容

Phase 1〜13で確定したDAKEプロジェクトの正式運用状態を `00_core/PROJECT_INFORMATION_SOURCE.md` として整理した。

`00_core/README.md` には、情報源一覧として `PROJECT_INFORMATION_SOURCE.md` を追加した。

## 現在の正式運用状態

DAKEシリーズは、`ORIGINAL.md` を真の正本とする運用へ移行済みである。

Store、Stripe Payment Link、Store同期、DAKE出荷定義v2、Codex標準指示v2、正本主義エンジンv0.2を前提とする。

## 正本

真の正本は `ORIGINAL.md`。

矛盾がある場合は `ORIGINAL.md` を優先する。

## 派生ビュー

次を派生ビューとして整理した。

- `README.md`
- `DAKE_META`
- `release_body.md`
- `booth_product.txt`
- Store表示

## Store

`store.dakeapp.com` は正本ではなく販売ビューである。

ただし、正式出荷先として扱う。

Store同期は `tools/store/sync_store_to_site.py` を使用する。

## Stripe

Stripe Payment Linkは利用可能。

Stripe Secret、APIキー、Webhook Secretは公開repo、generated JSON、Store JavaScriptへ保存しない。

## 出荷定義

正式出荷には GitHub Release、BOOTH、dakeapp.com、store.dakeapp.com を含める。

Store確認では `payment_status` と商品詳細URLを確認する。

## Codex標準指示

Codex確認順を次の通り整理した。

1. `ORIGINAL.md`
2. `README.md`
3. `DAKE_META`
4. `release_body.md`
5. `booth_product.txt`

`ORIGINAL.md` がない既存対象はREADMEを暫定参照し、必要に応じて `ORIGINAL.md` 作成を提案する。

## 今回やらなかったこと

- Store UI修正
- Stripe追加
- Webhook実装
- R2実装
- 商品変更
- README再生成
- 全アプリ再レビュー
