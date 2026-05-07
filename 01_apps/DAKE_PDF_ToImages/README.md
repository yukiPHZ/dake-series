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

## DAKE共通仕様確認

- 確認日: 2026-05-06
- 表示名: `PDF→画像`
- exe名: `DakePDF_to_Images.exe`
- フォント: `BIZ UDPGothic` を最優先、`Yu Gothic UI` / `Meiryo` にフォールバック
- ヘッダー: 画面内にアプリ名を重複表示せず、機能タイトルと短い説明文で構成
- フッター: DAKE共通文言を `UI_TEXT` 管理し、広幅時は左右配置、狭幅時は中央寄せ2段構成
- アイコン: `..\..\02_assets\dake_icon.ico` の共通アイコンのみ参照
- ビルド確認: `build.bat` で `dist\DakePDF_to_Images.exe` を生成
- 起動確認: 生成後の exe が起動することを確認

## DAKE_META

```json
{
  "app_key": "dake_pdf_toimages",
  "display_name": "DakePDFto画像",
  "launcher_title": "PDF→画像",
  "launcher_description": "PDFの各ページをPNG画像に変換します。",
  "site_title": "DakePDFto画像",
  "site_description": "PDFを追加して保存先を選ぶだけで、各ページをPNG画像として書き出せるWindows向けアプリです。",
  "update_summary": "READMEメタ情報とRelease本文を整備。スクリーンショットをassets/screenshot.webpに作成。",
  "folder_name": "DAKE_PDF_ToImages",
  "exe_name": "DakePDF_to_Images.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## RELEASE_BODY

- PDF→画像変換アプリ
- ページごとのPNG保存に対応
- 複数PDFの追加に対応
- Windows向けexe
