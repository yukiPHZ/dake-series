# DAKE BOOTH準備仕様

## 一文での責務

Codexは売る準備まで。BOOTH登録は人間。

BOOTHはDISTRIBUTIONの一導線であり、SHIP、dakeapp.com、Store、Stripeとは独立して扱います。

## 基本フロー

```text
ORIGINAL.md
↓
BOOTH登録素材を生成・確認
↓
booth_ready/
↓
BOOTH_PREP = READY
↓
Codex STOP
↓
DakeBOOTHアシスト
↓
人間確認
↓
人間がBOOTH登録・公開
```

## Codexが行うこと

- 対象商品の `ORIGINAL.md` を正本として読む。
- `booth_product.txt` を生成・更新する。
- `booth_thumbnail.jpg` を生成する。
- 必要なスクリーンショットを参照する。
- 配布ZIPを準備する。
- `booth_ready/` に登録素材をまとめる。
- 登録に不足している素材を報告する。

正常な準備完了状態は `BOOTH_PREP = READY` です。BOOTH URLは準備完了条件ではありません。

## Codexが行わないこと

- BOOTHサイトを開く。
- Playwrightその他のブラウザ自動化でBOOTHを操作する。
- 商品登録画面へ入力する。
- 商品画像・ZIPをアップロードする。
- 商品登録・更新・公開・販売開始を行う。
- BOOTHのログイン状態を操作する。
- DakeBOOTHアシストを自動起動・自動操作する。

これらはすべて人間操作です。BOOTH登録作業をCodex workflowへ組み込みません。

## 正本と派生ビュー

- 正本: `ORIGINAL.md`
- 派生ビュー: README、DAKE_META、`release_body.md`、`booth_product.txt`、画像、ZIP、`booth_ready/`

LegacyアプリでORIGINALが未導入の場合だけREADMEを暫定参照します。変更が影響する派生ビューだけを更新します。

## booth_ready完成条件

人間がDakeBOOTHアシストで登録するため、原則として次を揃えます。

- `booth_product.txt`
- `booth_thumbnail.jpg`
- 配布ZIP
- `README.txt` または `注意事項.txt`
- 必要な場合のみ `screenshot.webp` / `screenshot.jpg`

GitHub Release、BOOTH URL、SNS、Buffer、動画、Release capture、Cloudflare、Store、Stripeは `BOOTH_PREP` の完了条件にしません。

## booth_product.txt

基本項目:

- 商品名
- 価格
- 説明文
- タグ
- GitHub Release URL（存在する場合）
- 商品画像
- 配布ZIP
- 注意事項
- 更新区分: NEW / UPDATE
- BOOTH URL（人間の登録後のみ）

BOOTH URLが未設定ならNEW候補、登録済みならUPDATE候補として人間へ示せますが、Codexが登録・更新を実行する意味ではありません。

## サムネイル

- 1200x1200 JPG
- 小サイズでも用途が分かる
- DAKE共通の軽い空気感
- 過剰演出を避け、実務道具感を優先する

## 登録後のURL記録

人間が登録を完了し、`https://peakheadz.booth.pm/items/<数字>` を取得した後だけ正本へ記録します。

`tools/record_booth_registration.py` はBOOTH登録ツールではありません。ユーザーがURLを明示し「記録して」と依頼した場合だけ、登録済みURLを `ORIGINAL.md` の `booth_url` へ記録します。商品状態、決済状態、Stripe情報を同時変更しません。

## DakeBOOTHアシスト

DakeBOOTHアシストは人間用の登録補助アプリです。アプリ内部に入力補助機能があっても、Codexの通常責務とは分離します。最終確認、登録、公開、販売開始は人間が行います。
