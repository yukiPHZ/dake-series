# Dake画像リサイズ

スマホで撮影した画像を、長辺1600px上限のJPEGに整えて軽量化する Windows デスクトップアプリです。

## 概要

- 出力は JPEG 固定です
- 長辺は 1600px を上限にし、比率は維持します
- 元画像は上書きしません
- 出力ファイル名の末尾に `_resizeDake` を付けて保存します
- 出力先は `DakeImageResize_Output` フォルダです
- 保存先未指定時は元画像フォルダ内に `DakeImageResize_Output` を作成します
- 保存先を選んだ場合は選択フォルダ内に `DakeImageResize_Output` を作成します
- 同名ファイルがある場合は `_2`, `_3` の連番で回避します

## 対応形式

- jpg
- jpeg
- png
- bmp

HEIC / HEIF は初期版では非対応です。

## ビルド

`build.bat` を実行すると `PyInstaller` で exe 化できます。

## アイコン

- アプリ内の個別アイコンは持たず、共通アイコン `../../02_assets/dake_icon.ico` を参照します
- アイコン反映が古い場合は `build` / `dist` / `*.spec` を削除してから再ビルドしてください
