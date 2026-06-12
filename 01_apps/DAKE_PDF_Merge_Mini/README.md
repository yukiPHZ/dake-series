# DakePDF結合mini

PDFを最大5ファイルまで選択し、表示順に1つのPDFへ結合する社内配布向けmini版アプリです。

## 概要

- 最大5ファイルまでのPDFを結合します。
- 大容量PDFや大量ページPDFは対象外です。
- サムネイルカードをドラッグして結合順を並び替えできます。
- 保存時は保存ダイアログで出力先を指定します。

## 制限

- PDFファイル数: 最大5件
- 1ファイルサイズ: 最大50MB
- 合計サイズ: 最大150MB
- 1PDFのページ数: 最大100ページ
- 合計ページ数: 最大200ページ

## 操作方法

1. 「PDFを追加」をクリック、または画面中央へPDFをドラッグ＆ドロップします。
2. 必要に応じてカードをドラッグして順番を変更します。
3. 「結合して保存」をクリックします。
4. 保存ダイアログで保存先を選びます。初期ファイル名は `merged.pdf` です。

## ビルド方法

このフォルダで `build.bat` を実行します。

ビルド時に以下を指定しています。

- `--onefile`
- `--noconsole`
- `--clean`
- 共通アイコン: `..\..\02_assets\dake_icon.ico`
- バージョン情報: `version_info.txt`

## exeプロパティ

`version_info.txt` により、exeプロパティへ `KIKUTA YUKIHIKO` を設定しています。

- CompanyName: `KIKUTA YUKIHIKO`
- FileDescription: `PDF結合mini`
- ProductName: `DAKE PDF Merge Mini`
- LegalCopyright: `© 2026 KIKUTA YUKIHIKO`

## Git管理対象外ファイル

以下は `.gitignore` で管理対象外です。

- `build/`
- `dist/`
- `*.spec`
- `*_config.json`
- `__pycache__/`
- `*.pyc`

## DAKE_META

```json
{
  "app_key": "dake_pdf_merge_mini",
  "display_name": "DakePDF結合mini",
  "launcher_title": "PDF結合mini",
  "launcher_description": "最大5ファイルまでのPDFを並べ替えて結合します。",
  "site_title": "DakePDF結合mini",
  "site_description": "PDFを最大5ファイルまで選択し、表示順に1つのPDFへ結合できるWindows向けminiアプリです。",
  "update_summary": "READMEメタ情報とRelease本文を整備。スクリーンショットをassets/screenshot.webpに作成。",
  "folder_name": "DAKE_PDF_Merge_Mini",
  "exe_name": "DakePDF_Merge_Mini.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_PDF_Merge_Mini_v1.0.0",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true,
  "demo_video_path": "release_artifacts/demo.mp4",
  "demo_video_url": "",
  "social_release_path": "release_artifacts/social_release.json"
}
```

## RELEASE_BODY

- PDF結合miniアプリ
- 最大5ファイルまで対応
- サムネイル並べ替え対応
- Windows向けexe
