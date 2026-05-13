# Dakeフォルダ一覧

## 目的

指定したフォルダを基準に、フォルダ構成とファイル一覧を見える化するためのDAKEシリーズアプリです。

## できること

- フォルダを選んで構成をツリー表示する
- フォルダ名、ファイル名、拡張子、ファイルサイズ、更新日時を表示する
- 表示中の一覧をクリップボードへコピーする
- 表示中の一覧を `.txt` で保存する
- 同じフォルダを再読み込みする

## できないこと

- ファイルの編集
- ファイルやフォルダの削除
- ファイルやフォルダの移動
- ファイル名やフォルダ名の変更
- ファイル内容の読み取り、検索、分類

## 使い方

1. `main.py` を実行します。
2. 「フォルダを選ぶ」から対象フォルダを選択します。
3. 表示された一覧を必要に応じて「一覧をコピー」または「txt保存」します。
4. フォルダ内容を更新したあと確認する場合は「再読み込み」を押します。

## 注意事項

このアプリはフォルダ構成とファイル一覧を表示するための補助ツールです。
ファイルの編集、削除、移動、リネームは行いません。

アクセスできないフォルダはスキップし、一覧の集計にスキップ件数を表示します。

## ビルド方法

Windows環境で以下を実行します。

```bat
build.bat
```

ビルドに成功すると、`dist/DakeFolder_List.exe` が作成されます。

## 修正・確認履歴

- 2026-05-06: DAKE共通仕様に合わせて、ヘッダー、フッター、フォント、UI_TEXT、共通アイコン、build.bat、.gitignore を確認しました。
- フッターは広幅時に左右2ブロック、狭幅時に中央寄せ2段構成へ切り替わるように調整しています。
- `build.bat` によるEXE化と、`dist/DakeFolder_List.exe` の起動確認を実施済みです。

## DAKE_META

```json
{
  "app_key": "dake_folder_list",
  "display_name": "Dakeフォルダ一覧",
  "launcher_title": "フォルダ一覧",
  "launcher_description": "フォルダ構成とファイル一覧を表示・コピー・保存します。",
  "site_title": "Dakeフォルダ一覧",
  "site_description": "指定フォルダの構成、ファイル名、拡張子、サイズ、更新日時を一覧化できるWindows向けアプリです。",
  "update_summary": "READMEメタ情報とRelease本文を整備。スクリーンショットをassets/screenshot.webpに作成。",
  "folder_name": "DAKE_Folder_List",
  "exe_name": "DakeFolder_List.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Folder_List_v1.0.0",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## RELEASE_BODY

- フォルダ一覧作成アプリ
- フォルダ構成とファイル情報を表示
- コピー・txt保存に対応
- Windows向けexe
