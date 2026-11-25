# YUKIZ BLOG 投稿DAKE

YUKIZ BLOG に言葉を置くための専用アプリです。

本文を書き、`置く` を押すだけで、サイト側の静かな並びに HTML と JSON を反映します。タイトル、タグ、カテゴリ、SEO 入力欄はありません。

## 思想

このアプリは管理画面ではありません。

菊田幸彦が本文を書き、ただ置くための場所です。増やすことや整えることよりも、書く前の静けさと、置いたあとの余白を優先します。

## 使い方

1. `main.py` を起動します。
2. 中央の大きな入力エリアに本文を書きます。
3. 空気を見たい場合は `プレビュー` を押します。
4. 問題なければ `置く` を押します。

## 生成されるファイル

`置く` を押すと、YUKIZ BLOG サイト内で以下が更新されます。

- `posts/YYYYMMDD-001.html`
- `data/posts.json`
- `posts/index.html`
- `main.js`
- `sitemap.xml`
- `robots.txt`

同じ日に複数置いた場合は、`YYYYMMDD-002.html` のように連番になります。

## 削除機能がない理由

投稿DAKEは、置いた言葉を簡単に消すための道具ではありません。

緊急時は、GitHub、`data/posts.json`、記事 HTML を直接編集して対応します。初期実装では過去記事の編集機能も付けていません。

## 生成ルール

外部 AI API は使いません。

本文の最初の意味ある一文から、名詞、場所、道具、行為語に見える語を簡易ルールで抽出し、`/` 区切りの内部タイトルを作ります。抽出できない場合は本文冒頭、それも難しい場合は日付を使います。

meta description と OGP description は本文冒頭から 60〜90 文字程度で静かに生成します。

## ビルド方法

```bat
build.bat
```

PyInstaller で `dist/Dake_YukizBlog_Post.exe` を作成します。

DAKE 共通アイコン `C:\Users\yukiz\devlop\DAKE_series\02_assets\dake_icon.ico` がある場合は自動で使用します。

## 注意事項

対象サイトは初期値で `C:\Users\yukiz\devlop\yukizblog-site` です。

`sitemap.xml` の URL は初期値として `https://yukizblog-site.pages.dev/` を使います。公開 URL が変わる場合は、環境変数 `YUKIZBLOG_BASE_URL` で指定してから実行してください。
