# Dakeスクショ印刷

Win + Shift + S で切り取ったスクリーンショット画像を、クリップボードから取得してA4縦に印刷するDAKEシリーズの単機能アプリです。

## 使い方

1. Win + Shift + S で必要な部分だけスクショします。
2. Dakeスクショ印刷を起動します。
3. A4縦プレビューを確認します。
4. 「印刷」を押します。

## 仕様

- クリップボード内の画像だけを取得します。
- A4縦印刷専用です。
- 画像は縦横比を維持して中央に配置します。
- 履歴保存はしません。
- 画像編集、OCR、トリミング、複数画像対応はしません。

シンプルそれDAKEシリーズ / 止まらない、迷わない、すぐ終わる。

## 2026-05-06 確認

- DAKE共通仕様に合わせて、フォント、ヘッダー、フッター、UI_TEXT、共通アイコン参照を確認しました。
- `build.bat` で `DakeScreenshot_Print.exe` を再ビルドし、起動確認を行いました。
- 実印刷は未実行です。印刷処理はスクショ画像のみをA4縦へ配置する静的確認まで実施しています。

## DAKE_META

```json
{
  "app_key": "dake_screenshot_print",
  "display_name": "Dakeスクショ印刷",
  "launcher_title": "スクショ印刷",
  "launcher_description": "クリップボードのスクリーンショットをA4縦で印刷します。",
  "site_title": "Dakeスクショ印刷",
  "site_description": "Win + Shift + Sで切り取った画像をクリップボードから取得し、A4縦プレビューで印刷できるWindows向けアプリです。",
  "update_summary": "READMEメタ情報とRelease本文を整備。スクリーンショットをassets/screenshot.webpに作成。",
  "folder_name": "DAKE_Screenshot_Print",
  "exe_name": "DakeScreenshot_Print.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## RELEASE_BODY

- スクリーンショット印刷アプリ
- クリップボード画像の取得に対応
- A4縦プレビュー付き
- Windows向けexe
