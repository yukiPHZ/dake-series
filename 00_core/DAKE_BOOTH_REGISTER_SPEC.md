# DAKE BOOTH登録最適化仕様

## 目的

BOOTH登録作業を、DAKE正式出荷ラインの一部として標準化する。
BOOTHは単なる後作業ではなく、ユーザーが触れる状態へ届けるための正式工程とする。

GitHub Releaseは配布物公開、BOOTH登録はユーザーが見つけて理解し取得できる状態へ近づける工程です。
DAKEでは、BOOTH readyとBOOTH登録確認までを正式出荷ラインに含める。

## 正本

BOOTH登録に使う正本は以下。

- `README.md`
- `release_body.md`
- `booth_product.txt`（通常は `booth_ready/booth_product.txt`）
- `assets/screenshot.webp`
- `assets/booth_thumbnail.jpg`
- `booth_ready/`

## booth_product.txt の推奨構成

以下の項目を標準化する。

- 商品名
- 価格
- 説明文
- タグ
- GitHub Release URL
- BOOTH URL
- 商品画像
- 配布zip
- 注意事項
- 更新区分: NEW / UPDATE

Markdown見出し形式を基本とする。
将来形式として以下のキー形式も許可する。

```text
TITLE=
PRICE=
DESCRIPTION=
TAGS=
URL=
GITHUB_RELEASE=
ZIP_PATH=
THUMBNAIL_PATH=
SCREENSHOT_PATH=
MODE=
```

## BOOTH登録補助方針

DakeBOOTHアシストは Playwright を使って入力補助する。
ただし、公開ボタンは押さない。
最終確認と公開判断は人間が行う。

## 自動化の境界

やること:

- 商品情報の読み取り
- BOOTH登録画面を開く
- 入力補助
- 商品画像選択補助
- zip選択補助
- コピー補助
- `booth_ready/` 確認

やらないこと:

- 自動公開
- 自動販売開始
- ログイン情報保存
- 決済情報操作
- 売上情報取得
- BOOTH内部APIへの直接アクセス
- 無確認の既存商品更新

## booth_ready 完成定義

`booth_ready/` に以下が揃っていること。

- 配布zip
- `booth_thumbnail.jpg`
- `booth_product.txt`
- `README.txt` または `注意事項.txt`
- 必要に応じて `screenshot.webp` または `screenshot.jpg`

zipはBOOTHへアップロードする作品ファイルです。
Git管理には含めず、必要時にローカルで生成する。

## サムネ仕様

- 1200x1200
- JPG
- 小サイズでも用途が分かる
- DAKE共通空気感
- 過剰演出禁止
- 実務道具感を優先

`assets/screenshot.webp` は実際に動く感、`assets/booth_thumbnail.jpg` は一覧で止まる感を担当する。

## NEW / UPDATE

新規登録と既存更新は分ける。

- NEW: 新規商品登録
- UPDATE: zip差し替え、説明文更新、画像差し替え

`booth_product.txt` に BOOTH URL が未登録なら NEW 候補。
BOOTH URL が登録済みなら UPDATE 候補として扱う。

## 点検チェック

各アプリについて以下を確認する。

- `README.md` がある
- `DAKE_META` がある
- `release_body.md` がある
- `assets/screenshot.webp` がある
- `assets/booth_thumbnail.jpg` がある
- `booth_product.txt` がある
- `booth_ready/` がある
- `booth_ready/` に zip がある
- `build.bat` が共通アイコンを参照している
- `main.py` で `root.iconbitmap` が共通アイコンを参照している、または安全に設定されている

## 不足時の扱い

- `release_body.md` がない場合は、READMEの `RELEASE_BODY` から生成する。
- `booth_product.txt` がない場合は、READMEの `DAKE_META` と `RELEASE_BODY` から暫定生成する。
- `assets/booth_thumbnail.jpg` がない場合は、`tools/make_booth_ready.py` で生成する。
- `assets/screenshot.webp` がない場合は、自動撮影できる場合のみ作成し、難しい場合は要スクショとして止める。
- `booth_ready/` がない場合は、`tools/make_booth_ready.py` の既存仕様に従って生成する。
- zipがない場合は、`dist/*.exe` があればBOOTH ready用zipを生成し、exeがなければ要buildとして扱う。

## DakeBOOTHアシスト

Playwrightは入力補助に使用する。
公開操作、販売開始、既存商品の無確認更新はしない。

ローカルプロファイルを使う場合は `playwright_profile/` をGit管理に含めない。
