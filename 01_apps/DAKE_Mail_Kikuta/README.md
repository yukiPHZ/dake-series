# Dake菊田メール

## 目的

`Dake菊田メール` は、菊田宛の新規メール作成画面を開くためのDAKEシリーズ単機能アプリです。

宛先は `kikuta@sakuratoshi.co.jp` 固定です。

## 使い方

1. アプリを起動します。
2. 件名と本文を入力します。
3. 「メールを作る」ボタンを押します。
4. 既定のメールソフトで、菊田宛の新規メール作成画面だけが開きます。

## 注意

- このアプリはメールを自動送信しません。
- SMTP送信、Gmail API、ログイン機能は使いません。
- 既定メールソフトを `mailto:` で開くだけです。
- 送信するかどうかの最終判断は、必ず人間がメールソフト側で行います。
- DAKEシリーズ共通アイコン `..\..\02_assets\dake_icon.ico` を使用します。
- 個別アイコンは作成しません。

## DAKEシリーズ

単機能、軽量、迷わない、現場で止まらないことを優先したDAKEシリーズのアプリです。

## DAKE_META

```json
{
  "app_key": "dake_mail_kikuta",
  "display_name": "Dake菊田メール",
  "launcher_title": "菊田メール",
  "launcher_description": "菊田宛のメール作成画面を開きます。",
  "site_title": "Dake菊田メール",
  "site_description": "件名と本文を入力して、菊田宛の新規メール作成画面を既定メーラーで開くWindows向けアプリです。",
  "update_summary": "READMEメタ情報とRelease本文を整備。スクリーンショットをassets/screenshot.webpに作成。",
  "folder_name": "DAKE_Mail_Kikuta",
  "exe_name": "DakeKikuta_Mail.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Mail_Kikuta_v1.0.0",
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

- 菊田宛メール起動アプリ
- 件名・本文入力に対応
- 既定メーラーで作成画面を開く
- Windows向けexe
