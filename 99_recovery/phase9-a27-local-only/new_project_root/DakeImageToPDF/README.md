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
