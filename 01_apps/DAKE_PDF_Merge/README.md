# DakePDF Merge

DAKEシリーズ向けのPDF結合アプリです。追加したPDFを画面上で並べ替え、表示順どおりに1つのPDFへ結合します。

## ファイル構成

- `main.py`
- `requirements.txt`
- `build.bat`
- `README.md`

## アプリ情報

- 表示名: `PDF結合`
- exe名: `DakePDF_Merge.exe`
- 設定ファイル: `dake_pdf_merge_config.json`
- 共通アイコン: `..\..\02_assets\dake_icon.ico`

## 実行

```bat
python main.py
```

## ビルド

```bat
build.bat
```

配布対象は `dist\DakePDF_Merge.exe` です。

## 確認記録

2026-05-06:

- DAKE共通仕様に合わせて、UI文言を `UI_TEXT` 管理へ整理しました。
- フォントは `BIZ UDPGothic` を最優先、`Yu Gothic UI` / `Meiryo` をフォールバックにしています。
- ヘッダーは画面内アプリ名を出さず、機能タイトルと短い説明に整理しています。
- フッターは左にシリーズ文言、右に `戸建買取査定` / `Instagram` / コピーライトを配置しています。
- 共通アイコン `..\..\02_assets\dake_icon.ico` を使用しています。
- `build.bat` によるビルド成功を確認しました。
- `dist\DakePDF_Merge.exe` の起動を確認しました。

## DAKE_META

```json
{
  "app_key": "dake_pdf_merge",
  "display_name": "DakePDF結合",
  "launcher_title": "PDF結合",
  "launcher_description": "複数PDFを並べ替えて1つに結合します。",
  "site_title": "DakePDF結合",
  "site_description": "追加したPDFを画面上で並べ替え、表示順どおりに1つのPDFへ結合できるWindows向けアプリです。",
  "update_summary": "READMEメタ情報とRelease本文を整備。スクリーンショットをassets/screenshot.webpに作成。",
  "folder_name": "DAKE_PDF_Merge",
  "exe_name": "DakePDF_Merge.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## RELEASE_BODY

- PDF結合アプリ
- 複数PDFの追加に対応
- ドラッグ並べ替え対応
- Windows向けexe
