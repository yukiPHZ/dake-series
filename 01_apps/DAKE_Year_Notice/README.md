# Dake今年の注意点

起動するだけで、今年が閏年かどうか、固定資産税の評価替え年かどうか、次回の評価替え年を確認するDAKEシリーズの単機能アプリです。

正式フォルダ名は `DAKE_Year_Notice`、実行ファイル名は `DakeYear_Notice.exe` です。

## 表示内容

- 今年の西暦と和暦
- 閏年かどうか
- 固定資産税評価替え年かどうか
- 次回の評価替え年

## 判定ルール

閏年は一般ルールで判定します。

- 4で割り切れる年は閏年
- ただし100で割り切れる年は平年
- ただし400で割り切れる年は閏年

固定資産税評価替えは、2024年を基準に3年ごととして判定します。

```text
(対象年 - 2024) % 3 == 0
```

和暦表示は令和のみ対応です。

## ビルド方法

同じフォルダ内で以下を実行します。

```bat
build.bat
```

`dist` フォルダに `DakeYear_Notice.exe` が作成されます。

## DAKEシリーズ

シンプルそれDAKEシリーズの単機能アプリです。入力不要、設定不要、起動してすぐ確認できることを優先しています。

## DAKE_META

```json
{
  "app_key": "dake_year_notice",
  "display_name": "Dake今年の注意点",
  "launcher_title": "今年の注意点",
  "launcher_description": "閏年と固定資産税評価替え年を確認します。",
  "site_title": "Dake今年の注意点",
  "site_description": "今年が閏年か、固定資産税の評価替え年か、次回評価替え年はいつかを確認できるWindows向けアプリです。",
  "update_summary": "READMEメタ情報とRelease本文を整備。スクリーンショットをassets/screenshot.webpに作成。",
  "folder_name": "DAKE_Year_Notice",
  "exe_name": "DakeYear_Notice.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## RELEASE_BODY

- 今年の注意点確認アプリ
- 閏年判定に対応
- 固定資産税評価替え年を表示
- Windows向けexe
