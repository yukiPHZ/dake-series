# Dakeずっとメモ

1本の長文メモを、画面内の列へ左から右に流して見渡せるWindowsデスクトップアプリです。

複数メモ、タブ、検索、Markdown、装飾は持たせず、通常のメモ帳のように1本のテキストだけを自動保存します。

## アプリ概要

- 表示名: `ずっとメモ`
- ウィンドウタイトル: `Dakeずっとメモ`
- exe名: `DakeColumn_Memo.exe`
- 実装: Python + Tkinter
- 保存ファイル: `DAKE_Column_Memo_config.json`
- 共通アイコン: `..\..\02_assets\dake_icon.ico`

## できること

- 1本のメモを入力
- 入力内容を自動保存
- 起動時に前回のメモを自動復元
- 段組表示を `1 / 2 / 3 / 4` 列で切り替え
- 左列から右列へ流れる段組プレビュー
- 段組プレビュー全体を1つの長文として同期スクロール
- メモ全文をクリップボードへコピー

## 使い方

1. `メモ入力` に文章を書きます。
2. `列数` で `1 / 2 / 3 / 4` を選びます。
3. 右側の `段組表示` で、長いメモを列に流して確認します。
4. 必要なときは `コピー` で全文をコピーします。

## 注意点

- 扱うメモは1本だけです。
- 保存データはローカル環境用のJSONで、Git管理しません。
- 段組表示はプレビューです。編集は左側の1本の入力欄で行います。
- 列ごとの独立スクロールはありません。
- Markdown、リッチテキスト、画像添付、PDF出力、クラウド同期、AI機能はありません。

## 実行

```bat
python main.py
```

## ビルド

```bat
build.bat
```

`dist\DakeColumn_Memo.exe` が作成されます。

## DAKE_META

```json
{
  "app_key": "dake_column_memo",
  "display_name": "ずっとメモ",
  "launcher_title": "ずっとメモ",
  "launcher_description": "1本のメモを、列に流して見渡します。",
  "site_title": "Dakeずっとメモ",
  "site_description": "長くなったメモを、列に流して見渡せる単機能メモです。",
  "update_summary": "1本の長文メモを段組表示できるメモアプリを追加しました。",
  "folder_name": "DAKE_Column_Memo",
  "exe_name": "DakeColumn_Memo.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## RELEASE_BODY

```text
- 1本のメモを列に流して表示します。
- 1～4列の段組表示に切り替えできます。
- 自動保存・起動時復元に対応しています。
- Windows向けexeです。
```
