# DakeGitメモ

書くたびに、変更が残るメモアプリです。
Gitを意識せず、ただメモを書くだけで過去の内容を残せます。

## アプリ概要

`DakeGitメモ` は、1つの plain text メモを自動保存し、保存時点の内容を履歴として残す Windows デスクトップアプリです。
Gitクライアントではなく、一般ユーザーが「過去が消えないメモ」として使えることを目的にしています。

## できること

- 1つのメモを書く
- 本文変更後に自動保存する
- `Ctrl + S` で保存する
- アプリ終了時に保存する
- 保存時点の履歴を一覧表示する
- 選んだ履歴と現在の本文の変更を見る
- 選んだ履歴の内容に戻す
- 戻す前の現在内容も履歴に残す

## できないこと

- 複数メモ管理
- タグ管理や高度な検索
- Markdownプレビュー
- カレンダー日記UI
- クラウド保存や複数ユーザー同期
- GitHub連携
- branch、merge、push、pull などのGit操作
- AI要約、感情分析、SNS共有

## 保存場所

メモ本文と履歴は、ユーザーごとのアプリ専用データフォルダに保存します。

```text
C:\Users\<user>\AppData\Local\DAKE_Git_Memo
```

作成される主なファイルとフォルダ:

- `memo.txt`: 現在のメモ本文
- `history/`: 保存時点の本文スナップショット
- `config.json`: 最終保存日時などのアプリ設定

メモ本文は plain text です。
このデータフォルダはGit管理対象ではありません。

## 使い方

1. 起動したら、中央のメモ欄にそのまま書きます。
2. 入力後しばらくすると自動で保存されます。
3. 明示的に残したい時は `保存` を押します。
4. 右側の `履歴` から過去の時点を選び、`変更を見る` または `この時点に戻す` を押します。

## ビルド

アプリフォルダで以下を実行します。

```bat
build.bat
```

`dist\DakeGit_Memo.exe` が作成されます。

## DAKE_META

```json
{
  "app_key": "DAKE_Git_Memo",
  "display_name": "DakeGitメモ",
  "launcher_title": "DakeGitメモ",
  "launcher_description": "書くたびに、変更が残るメモです。",
  "site_title": "DakeGitメモ",
  "site_description": "Gitを意識せず、変更履歴を残せるメモアプリです。",
  "update_summary": "DakeGitメモ v1.0.0 を公開しました。",
  "folder_name": "DAKE_Git_Memo",
  "exe_name": "DakeGit_Memo.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Git_Memo_v1.0.0",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## RELEASE_BODY

```text
- 書くたびに変更が残る、1メモだけの軽量メモアプリです。
- 自動保存、Ctrl + S、終了時保存に対応しています。
- 履歴から変更を確認し、必要な時点に戻せます。
- Git未インストール環境でも使えるスナップショット方式です。
- Windows向けexeです。
```
