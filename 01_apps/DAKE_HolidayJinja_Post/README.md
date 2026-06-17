# holiday-jinja 投稿DAKE

holiday-jinja 用の投稿をPC上で静かに作成・整理するための小さなデスクトップアプリです。

CMSではなく、過去写真を見返し、記憶を辿り、静かに投稿を置くための作業机として作っています。

## 使い方

1. `Dake_HolidayJinja_Post.exe` または `python main.py` で起動します。
2. 初回または未設定時に `holiday-jinja-site` フォルダを選択します。
3. `写真を選択` から `jpg` / `jpeg` / `png` / `webp` の写真を選びます。
4. `title` / `text` / `location` を入力します。
5. 生成される `id` と今日の日付を確認して、`投稿を保存` を押します。
6. 必要に応じて `ローカルプレビューを開く` から `http://127.0.0.1:4173/` を開きます。

## 更新対象フォルダ

このアプリが更新する対象は、原則として次の2つだけです。

- `holiday-jinja-site/images/`
- `holiday-jinja-site/posts.json`

サイト本体の `index.html` / `style.css` / `main.js` は変更しません。

## posts.json 追記仕様

既存の `posts.json` を読み込み、`hj-001` 形式の最大IDから次の連番を生成します。

追記例:

```json
{
  "id": "hj-004",
  "image": "images/hj-004.jpg",
  "title": "after rain.",
  "text": "雨のあと。",
  "location": "Japan",
  "date": "2026-05-03",
  "tags": []
}
```

`tags` は v0.1 では空配列です。

## 画像保存仕様

選択した写真は `holiday-jinja-site/images/` に、投稿IDと同じ名前のJPEGとして保存します。

例:

```text
images/hj-004.jpg
```

長辺が約2000pxを超える場合は縮小し、JPEG品質は高めに保ちます。

## v0.1 の方針

- AI連携はしません。
- GitHub push はしません。
- Instagram投稿機能はありません。
- タグ管理UIや投稿一覧編集UIはありません。

GitHubへの反映は、サイト側の変更確認後に手動で行ってください。

## 凍結メモ

- 凍結日: 2026-06-17
- 理由: `holiday-jinja.com` 本体は解約候補であり、コンテンツ再利用は行わない。`sky` のみ `sky.niceskill.com` へ移行する。
- 凍結対象: この投稿DAKEからの新規投稿、画像コピー、`posts.json` 追記、ローカル投稿運用。
- 自動投稿、予約投稿、投稿API、管理画面は持たない。もし外部に残っている場合は停止対象。
- 再開条件: holiday-jinja本体を維持する明確な運用理由ができ、サイトrepoのREADME/SITE_SPECを先に更新した場合のみ。

## ビルド

依存関係を入れたうえで、`build.bat` を実行します。

```bat
pip install -r requirements.txt
build.bat
```

出力:

```text
dist/Dake_HolidayJinja_Post.exe
```

## DAKE_META

```json
{
  "app_key": "dake_holidayjinja_post",
  "display_name": "holiday-jinja 投稿DAKE",
  "launcher_title": "holiday-jinja 投稿",
  "launcher_description": "holiday-jinja 用の投稿データをローカルで作成します。",
  "site_title": "holiday-jinja 投稿DAKE",
  "site_description": "holiday-jinja 用の投稿データをローカルで整える内部向けアプリです。",
  "update_summary": "内部投稿ツールとして整理しました。",
  "folder_name": "DAKE_HolidayJinja_Post",
  "exe_name": "Dake_HolidayJinja_Post.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "frozen",
  "show_in_launcher": false,
  "show_on_site": false,
  "demo_video_path": "release_artifacts/demo.mp4",
  "demo_video_url": "",
  "social_release_path": "release_artifacts/social_release.json"
}
```

## RELEASE_BODY

- holiday-jinja投稿作成アプリ
- 写真・本文・場所の入力に対応
- 投稿データをPC上で整理
- Windows向けexe
