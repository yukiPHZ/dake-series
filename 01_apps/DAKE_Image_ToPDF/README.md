# DakeImageToPDF

画像を1枚入れるだけで、A4のPDFを1枚作るための DAKE アプリです。

## 特徴

- 画像1ファイルだけに対応
- ドラッグ＆ドロップまたはクリック追加に対応
- 画像の縦横を見て A4 縦 / A4 横 を自動判定
- アスペクト比を保ったまま中央配置
- 保存先は `Downloads`
- 保存後はフォルダを自動で開く

## 対応形式

- `png`
- `jpg`
- `jpeg`
- `bmp`
- `webp`

## 非対応形式

- `gif`
- `tiff`
- `heic`
- `svg`
- `pdf`

## 実行

```powershell
py main.py
```

## しまりすくん連携CLI

`--from-shimarisu` がある場合だけGUIを表示せず、指定画像を順番通りに1つのPDFへ変換します。

```powershell
DakeImageToPDF.exe --inputs "A.jpg" "B.png" --from-shimarisu
DakeImageToPDF.exe --inputs "A.jpg" "B.png" --output "C:\out\merged.pdf" --from-shimarisu --silent
```

- `--inputs`: 画像ファイルパスを1つ以上指定
- `--output`: 出力先フォルダまたはPDFファイルパス。未指定時は最初の画像と同じフォルダ
- `--from-shimarisu`: CLIモード起動フラグ
- `--silent`: 確認ダイアログ等を出さない指定
- CLI対応形式: `jpg`, `jpeg`, `png`, `webp`, `bmp`, `tif`, `tiff`
- 正常終了は exit code `0`、エラー時は exit code `1`

## ビルド

```powershell
build.bat
```

## 出力ルール

- 元画像ファイル名に `_dake` を付けて PDF を保存
- 例: `sample.png` -> `sample_dake.pdf`
- 同名ファイルがある場合は `_1`, `_2` を付けて退避

## 構成

- `main.py`
- `build.bat`
- `requirements.txt`
- `README.md`
- `app.ico`
- `icon.png`

## DAKE_META

```json
{
  "app_key": "dake_image_topdf",
  "display_name": "DakeImageToPDF",
  "launcher_title": "画像→PDF",
  "launcher_description": "画像1枚をA4のPDFに変換します。",
  "site_title": "DakeImageToPDF",
  "site_description": "画像を1枚追加するだけで、A4縦または横のPDFへ変換できるWindows向けアプリです。",
  "update_summary": "READMEメタ情報とRelease本文を整備。スクリーンショットをassets/screenshot.webpに作成。",
  "folder_name": "DAKE_Image_ToPDF",
  "exe_name": "DakeImageToPDF.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Image_ToPDF_v1.0.0",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## RELEASE_BODY

- 画像→PDF変換アプリ
- ドラッグ＆ドロップ対応
- A4縦横の自動判定
- Windows向けexe
