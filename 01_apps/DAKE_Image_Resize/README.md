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
- webp
- bmp
- tif
- tiff

HEIC / HEIF は初期版では非対応です。

## しまりすくん連携CLI

`--from-shimarisu` がある場合のみCLIモードで起動し、GUIを表示せずに画像を順番通りリサイズします。

```bat
Dake_Image_Resize.exe --inputs "A.jpg" "B.png" --from-shimarisu
```

- `--inputs`: 画像ファイルパスを1つ以上指定します
- `--output`: 出力先フォルダを指定します。未指定時は最初の画像と同じフォルダへ保存します
- `--max-size`: 長辺の最大pxを指定します。未指定時は1600pxです
- `--silent`: しまりすくん連携用の任意フラグです
- 出力名は `photo_resized.jpg`, `photo_resized_2.jpg` のように元ファイル名を活かして保存します
- 元画像は上書き・削除しません

## ビルド

`build.bat` を実行すると `PyInstaller` で exe 化できます。

## アイコン

- アプリ内の個別アイコンは持たず、共通アイコン `../../02_assets/dake_icon.ico` を参照します
- アイコン反映が古い場合は `build` / `dist` / `*.spec` を削除してから再ビルドしてください

## DAKE_META

```json
{
  "app_key": "dake_image_resize",
  "display_name": "Dake画像リサイズ",
  "launcher_title": "画像リサイズ",
  "launcher_description": "スマホ写真を長辺1600px上限のJPEGに整えます。",
  "site_title": "Dake画像リサイズ",
  "site_description": "スマホ写真などの画像を長辺1600px上限のJPEGへ軽量化できるWindows向けアプリです。",
  "update_summary": "READMEメタ情報とRelease本文を整備。スクリーンショットをassets/screenshot.webpに作成。",
  "folder_name": "DAKE_Image_Resize",
  "exe_name": "Dake_Image_Resize.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Image_Resize_v1.0.0",
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

- 画像リサイズアプリ
- 長辺1600px上限で軽量化
- JPEG出力に対応
- Windows向けexe
