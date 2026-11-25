# Dake専用Web入口Builder

これは指定URL専用の入口アプリを作るBuilderです。
ブラウザではありません。
公式ツールではありません。
対象Webサイトの表示や動作を保証しません。

## 概要

- 表示名、画面見出し、起動URL、exe名、推奨ブラウザを入力して、専用入口アプリ一式を作成します。
- 生成物は `output` に作成されます。
- Edge推奨サイト向けに使えます。
- 生成アプリはURLを開くだけで、サイト内容の加工、データ取得、自動操作、ログ保存は行いません。

## 生成されるもの

`output` 内に、生成アプリ用の `main.py`、`build.bat`、`requirements.txt`、`README.md` を作成します。

## ビルド

Windows環境で `build.bat` を実行すると、`dist/DakeWebOne_Builder.exe` が作成されます。
共通アイコン `..\..\02_assets\dake_icon.ico` を参照します。
