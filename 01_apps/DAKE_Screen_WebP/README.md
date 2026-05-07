# DAKEシリーズ｜スクショ→WebP

現在アクティブなウインドウだけを撮影し、白背景に合成してWebP形式で保存する単機能アプリです。

## できること

- 選んだウインドウだけをキャプチャ
- 横幅1200pxを上限にリサイズ
- WebP品質88で保存
- 既存ファイルを確認して連番保存
- `Ctrl + Shift + 1` のショートカットで即保存
- ボタン操作時は3秒以内に撮影対象を選択

## 使い方

ボタンで保存する場合：

1. 「選んでWebP保存」を押す
2. アプリが一時的に隠れる
3. 3秒以内に撮りたいウインドウをクリックする
4. 自動でWebP保存される

ショートカットで保存する場合：

1. 撮りたいウインドウを前面に出す
2. `Ctrl + Shift + 1` を押す
3. 即保存される

保存先：

```text
デスクトップ\DAKE_screenshots
```

## 保存先

ユーザーのデスクトップ直下に `DAKE_screenshots` フォルダを自動作成します。

例：

```text
C:\Users\ユーザー名\Desktop\DAKE_screenshots
```

## ファイル名ルール

保存先フォルダ内の最大番号 + 1 で保存します。上書きはしません。

```text
screenshot-01.webp
screenshot-02.webp
screenshot-03.webp
```

## ショートカットキー

```text
Ctrl + Shift + 1
```

## ビルド方法

必要なライブラリをインストールしてから、`build.bat` を実行します。

```bat
pip install -r requirements.txt
build.bat
```

ビルド後、`dist\DakeScreen_WebP.exe` が作成されます。

## 注意事項

- ボタン操作では、アプリが最小化してから3秒後のアクティブウインドウを保存します。
- ショートカット操作では、その時点で前面にあるウインドウを即保存します。
- 横幅が1200px未満の画像は、原則として拡大しません。
- `Ctrl + Shift + 1` が他アプリで使われている場合は、ボタンから保存してください。

## DAKE_META

```json
{
  "app_key": "dake_screen_webp",
  "display_name": "DakeScreen_WebP",
  "launcher_title": "スクショ→WebP",
  "launcher_description": "選んだウインドウをWebPで保存します。",
  "site_title": "DakeScreen_WebP",
  "site_description": "現在アクティブなウインドウだけを撮影し、横幅1200px以内のWebPとして保存できるWindows向けアプリです。",
  "update_summary": "READMEメタ情報とRelease本文を整備。スクリーンショットをassets/screenshot.webpに作成。",
  "folder_name": "DAKE_Screen_WebP",
  "exe_name": "DakeScreen_WebP.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## RELEASE_BODY

- ウインドウスクショWebP保存アプリ
- 横幅1200px以内で保存
- ショートカット操作に対応
- Windows向けexe
