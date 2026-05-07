# HEIC→JPG変換

HEIC画像をJPGに変換するDAKEシリーズの単機能アプリです。

## 使い方

1. HEIC画像、複数ファイル、またはフォルダを画面中央へドロップします。
2. 自動で変換します。
3. 元ファイルと同じフォルダへJPGを保存します。

## 保存仕様

- 保存先: 元ファイルと同じフォルダ
- ファイル名: 元ファイル名 + `.jpg`
- 重複時: `_2`, `_3` を付与

## ビルド

```bat
build.bat
```

## DAKE_META

```json
{
  "app_key": "dake_image_heictojpg",
  "display_name": "HEIC→JPG変換",
  "launcher_title": "HEIC→JPG",
  "launcher_description": "HEIC画像をJPGへ変換します。",
  "site_title": "HEIC→JPG変換",
  "site_description": "HEIC画像やフォルダを追加して、元画像と同じ場所へJPGを書き出せるWindows向けアプリです。",
  "update_summary": "READMEメタ情報とRelease本文を整備。スクリーンショットをassets/screenshot.webpに作成。",
  "folder_name": "DAKE_Image_HEICtoJPG",
  "exe_name": "DakeHEIC_JPG.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## RELEASE_BODY

- HEIC→JPG変換アプリ
- 複数ファイル・フォルダ追加に対応
- 元ファイルと同じ場所へ保存
- Windows向けexe
