# QPSC Dashboard

QPSC Dashboard は、QPSC 全体の現在地を確認するための最上位レイヤーです。

App Dashboard と Web Dashboard は完成形ではなく監視ノードです。QPSC Dashboard は、それらを作り直さず、既存ノードから取れる状態を束ねて一目で見せます。

## 役割

- App Dashboard の状態を取得する
- Web Dashboard の状態を取得する
- 全体要約を表示する
- 既存 App Dashboard を起動する
- 既存 Web Dashboard を起動する

## 初版の対象

```text
QPSC_Dashboard
├─ App Dashboard
└─ Web Dashboard
```

BRAINZ、OIKAWA、Slack、Git、ORBIT は将来ノードです。初版では内部機能を実装しません。

## 実装方針

既存 Dashboard の `main.py` を動的に読み込み、既存の状態取得関数を利用します。

QPSC Dashboard 側では、README 再解析、Git 取得、Cloudflare 取得、API 接続などを大量に再実装しません。

## 起動

```powershell
python main.py
```

## launch-check

```powershell
python main.py --launch-check
```

成功時は exit code 0 で終了します。

## build

```powershell
.\build.bat
.\dist\QPSC_Dashboard.exe --launch-check
```

## 内部ツール

release、GitHub Release、BOOTH、dakeapp 掲載は不要です。

## RELEASE_BODY

```text
QPSC全体の現在地を確認する司令塔
App Dashboard / Web Dashboard を監視ノードとして集約
内部QPSC用ダッシュボード
Windows向けexe
```

## DAKE_META

```json
{
  "app_key": "DAKE_QPSC_Dashboard",
  "display_name": "QPSC Dashboard",
  "launcher_title": "QPSC Dashboard",
  "launcher_description": "QPSC全体の現在地を束ねて確認する司令塔",
  "site_title": "",
  "site_description": "",
  "update_summary": "App DashboardとWeb Dashboardの状態を束ねる内部司令塔",
  "folder_name": "DAKE_QPSC_Dashboard",
  "exe_name": "QPSC_Dashboard.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "internal",
  "show_in_launcher": true,
  "show_on_site": false
}
```
