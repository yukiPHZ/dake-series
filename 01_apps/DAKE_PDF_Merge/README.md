# DakePDF Merge

DAKEシリーズ向けのPDF結合アプリです。追加したPDFを画面上で並べ替え、表示順どおりに1つのPDFへ結合します。

## ファイル構成

- `main.py`
- `requirements.txt`
- `build.bat`
- `README.md`

## アプリ情報

- 表示名: `PDF結合`
- version: `1.0.1`
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

2026-05-10:

- アプリ下部に `他のDAKEツール` リンクを追加しました。
- リンク先は `https://dakeapp.com/launcher/` です。
- ランチャー直接DLではなく、DAKEツール導線用LPへのリンクです。

2026-05-12:

- しまりすくん連携用に `--from-shimarisu` 指定時だけGUIを開かないCLI結合に対応しました。
- `--inputs` で受け取ったPDFを順番どおりに結合し、正常時は exit code 0、エラー時は exit code 1 を返します。
- `--output` 未指定時は最初のPDFと同じフォルダへ `merged_YYYYMMDD_HHMMSS.pdf` として保存します。

## DAKE_META

```json
{
  "app_key": "dake_pdf_merge",
  "version": "1.0.1",
  "display_name": "DakePDF結合",
  "launcher_title": "PDF結合",
  "launcher_description": "複数PDFを並べ替えて1つに結合します。",
  "site_title": "DakePDF結合",
  "site_description": "追加したPDFを画面上で並べ替え、表示順どおりに1つのPDFへ結合できるWindows向けアプリです。",
  "update_summary": "起動ウインドウを画面サイズへ最適化し、保存処理を原子的保存へ変更。高速なPDF追加・並び替え・CLI互換性を維持した品質改善版。",
  "folder_name": "DAKE_PDF_Merge",
  "exe_name": "DakePDF_Merge.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_PDF_Merge_v1.0.1",
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

# DakePDF結合 v1.0.1

- PDF結合アプリ
- 複数PDFの追加に対応
- ドラッグ並べ替え対応
- 小型画面への表示最適化・保存安全性を改善
- Windows向けexe
