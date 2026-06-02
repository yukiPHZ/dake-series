# アドバンスドタイマー

アドバンスドタイマーは、集中時間と休憩時間をすぐに始めるための静かなタイマーです。
20年前の自分との約束を、いま起動できる形にしたDAKEアプリです。

## 機能

- 25分集中
- 5分休憩
- 15分休憩
- 1〜180分のカスタム時間
- 開始、一時停止、リセット
- 終了後の控えめな完了表示と「もう一回」

## ビルド方法

同じフォルダ内で以下を実行します。

```bat
build.bat
```

`dist` フォルダに `DakeAdvanced_Timer.exe` が作成されます。

## DAKE_META

```json
{
  "app_key": "time_advanced_timer",
  "display_name": "アドバンスドタイマー",
  "launcher_title": "アドバンスドタイマー",
  "launcher_description": "集中時間と休憩時間を、静かに始めるタイマーです。",
  "site_title": "アドバンスドタイマー",
  "site_description": "20年前の自分との約束を、いま起動できる形にした静かな集中タイマーです。",
  "update_summary": "集中・休憩・カスタム時間に対応した静かなタイマーを追加しました。",
  "folder_name": "DAKE_Time_AdvancedTimer",
  "exe_name": "DakeAdvanced_Timer.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## RELEASE_BODY

```text
アドバンスドタイマー
集中時間と休憩時間をすぐに始められる静かなタイマー
25分 / 5分 / 15分 / カスタム時間に対応
Windows向けexe
```
