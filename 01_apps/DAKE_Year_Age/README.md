# Dake築年数

築年を1つ入力するだけで、現在年基準の築年数を表示するDAKEシリーズの単機能アプリです。

## 使い方

1. 入力欄に築年を入力します。
2. 入力変更またはEnterで、築年数がすぐに表示されます。

## 入力例

- `1996`
- `平成8`
- `H8`
- `令和元`
- `R1`
- `昭和63`
- `１９９６`
- `Ｒ８`

## 対応元号

- 令和: 2019年以降
- 平成: 1989年から2019年
- 昭和: 1926年から1989年

明治・大正には対応していません。

## ビルド方法

```bat
build.bat
```

`dist\DakeYear_Age.exe` が作成されます。共通アイコン `..\..\02_assets\dake_icon.ico` を使用します。

## DAKEシリーズ表記

シンプルそれDAKEシリーズ  
© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta

## DAKE_META

```json
{
  "app_key": "dake_year_age",
  "display_name": "Dake築年数",
  "launcher_title": "築年数",
  "launcher_description": "西暦・和暦から築年数を確認します。",
  "site_title": "Dake築年数",
  "site_description": "築年を入力するだけで、現在年基準の築年数を西暦・和暦どちらからでも確認できるWindows向けアプリです。",
  "update_summary": "READMEメタ情報とRelease本文を整備。スクリーンショットをassets/screenshot.webpに作成。",
  "folder_name": "DAKE_Year_Age",
  "exe_name": "DakeYear_Age.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Year_Age_v1.0.0",
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

- 築年数確認アプリ
- 西暦・和暦入力に対応
- 入力すると即時計算
- Windows向けexe
