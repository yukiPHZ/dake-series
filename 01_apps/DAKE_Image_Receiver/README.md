# DakeImage_Receiver

スマホから画像受け取るDAKE は、PCに表示したQRコードをスマホで読み取り、画像をこのPCへ送るための入口アプリです。

## 使い方

1. アプリを起動
2. QRコードをスマホで読む
3. 画像を選んで送信
4. PC側の保存フォルダを確認

## 注意

- 同じWi-Fi内で使う
- HEICは変換しない
- 画像を受け取るだけのアプリ

## ビルド方法

`build.bat` を実行すると、`dist/DakeImage_Receiver.exe` を生成します。

## アイコン

DAKE共通アイコン `..\..\02_assets\dake_icon.ico` を使用します。アプリ個別アイコンは作成しません。

## DAKE_META

```json
{
  "app_key": "dake_image_receiver",
  "display_name": "DakeImage_Receiver",
  "launcher_title": "スマホ画像受信",
  "launcher_description": "スマホから画像をこのPCへ送ります。",
  "site_title": "DakeImage_Receiver",
  "site_description": "PCに表示したQRコードをスマホで読み取り、画像をWindows PCへ送るための受信アプリです。",
  "update_summary": "READMEメタ情報とRelease本文を整備。スクリーンショットをassets/screenshot.webpに作成。",
  "folder_name": "DAKE_Image_Receiver",
  "exe_name": "DakeImage_Receiver.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Image_Receiver_v1.0.0",
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

- スマホ画像受信アプリ
- QRコード接続に対応
- 画像をPC側へ保存
- Windows向けexe
