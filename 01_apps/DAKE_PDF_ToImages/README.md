# DakePDFto画像

## アプリ概要

`DakePDFto画像` は、PDF をページごとの PNG 画像へまとめて変換する DAKE シリーズ専用のデスクトップアプリです。  
PDF を追加して保存先を選ぶだけで、各 PDF をページ単位の画像として書き出します。

## 使い方

1. `PDFを追加` を押すか、対応環境ではドラッグ＆ドロップで PDF を追加します
2. 必要なら `保存先を選ぶ` で出力先を変更します
3. `画像に変換して保存` を押します
4. 完了ダイアログで OK を押すと保存先フォルダを開きます

## 入力

- PDF ファイルのみ対応
- 複数 PDF を同時に追加可能

## 出力

- 出力形式は PNG 固定
- 1 ページ = 1 画像
- ファイル名は `page_001.png` 形式のゼロ埋め連番
- 保存先の中に PDF ごとの出力フォルダを自動作成
- 同名フォルダがある場合はタイムスタンプ付きで安全に分岐

## 保存先

- 初期値は `Downloads`
- 最後に使った保存先は設定ファイルへ保存
- 設定ファイル名: `dake_pdf_to_images_config.json`
- 保存場所: `%LOCALAPPDATA%\DAKE_PDF_ToImages\dake_pdf_to_images_config.json`

## 共通アイコン

- DAKE シリーズ共通アイコンのみを使用
- 参照パス: `..\..\02_assets\dake_icon.ico`
- 実体: `C:\Users\yukiz\devlop\DAKE_series\02_assets\dake_icon.ico`
- アプリ個別アイコンは参照しません

## ビルド方法

1. このフォルダで `build.bat` を実行します
2. 初回は `.venv` が自動作成されます
3. 依存関係を自動インストール後、PyInstaller で exe を生成します
4. 成功すると `dist` フォルダが開きます

## 生成物

- onefile 形式
- 実行ファイル: `dist\DakePDF_to_Images.exe`

## 注意事項

- `tkinterdnd2` が利用できる環境ではドラッグ＆ドロップを有効にします
- exe のアイコンは PyInstaller の `--icon` で共通 ico を使用します
- 開発時のウィンドウアイコンは `main.py` で共通 ico を安全に適用します
