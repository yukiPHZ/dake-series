# Dake画像まとめPDF

複数の画像を並べて、1つのPDFにまとめるDAKEシリーズのWindowsデスクトップアプリです。

既存の `DAKE_Image_ToPDF` は画像1枚をPDF化するアプリです。このアプリは置き換えではなく、複数画像を順番どおり1つのPDFへまとめるための新規独立アプリです。

## できること

- 複数画像をドラッグ＆ドロップまたはファイル選択で追加
- 順番番号、サムネイル、ファイル名を確認
- 選択画像を上へ / 下へ移動
- 選択画像を削除、またはすべてクリア
- `1ページに1枚` または `1ページに4枚` を選択
- A4縦のPDFとして保存

## 対応形式

JPG / JPEG / PNG / BMP / WEBP / HEIC / HEIF

HEIC / HEIFは `pillow-heif` を使って読み込みます。EXIF Orientationを補正し、透明画像は白背景へ合成してPDF化します。

## 使い方

1. 画像をドラッグ＆ドロップ、または「画像を追加」から選びます。
2. 一覧で順番を確認します。
3. 必要なら「上へ」「下へ」で並び替えます。
4. `1ページに1枚` または `1ページに4枚` を選びます。
5. 「PDFにして保存」を押し、保存先とファイル名を選びます。

元画像は削除・移動・変更しません。

## ビルド

```powershell
cd 01_apps\DAKE_Image_BatchPDF
build.bat
```

`dist/DakeImage_BatchPDF.exe` を生成します。exe、build、dist、specはGit管理しません。

## DAKE_META

```json
{
  "app_key": "DAKE_Image_BatchPDF",
  "display_name": "画像まとめPDF",
  "launcher_title": "画像まとめPDF",
  "launcher_description": "複数画像を順番どおり1つのPDFにまとめます。",
  "site_title": "Dake画像まとめPDF",
  "site_description": "複数の画像を並べて、1ページ1枚または4枚のPDFにまとめるWindows向けDAKEアプリです。",
  "update_summary": "複数画像を順番確認して1つのPDFへまとめる新規アプリを追加しました。",
  "folder_name": "DAKE_Image_BatchPDF",
  "exe_name": "DakeImage_BatchPDF.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "draft",
  "show_in_launcher": false,
  "show_on_site": false,
  "app_type": "market",
  "completion_goal": "formal_release"
}
```

## RELEASE_BODY

- 複数画像を順番どおり1つのPDFにまとめます。
- 1ページ1枚 / 1ページ4枚を選べます。
- JPG / PNG / WEBP / HEIC / HEIFなどに対応します。
- Windows向けexeです。
