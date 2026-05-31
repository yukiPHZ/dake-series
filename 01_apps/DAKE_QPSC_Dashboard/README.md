# QPSC Dashboard

QPSC Dashboard は、QPSC 全体の現在地を確認するための最上位レイヤーです。

QPSC_Dashboardは、正本から未処理と次にやる候補を返す司令塔です。

App Dashboard / Web Dashboardは監視ノードであり、QPSC_Dashboardの下位に位置します。QPSC_Dashboardは関連アプリを自動起動せず、必要な時だけ詳細へ潜る導線を提供します。

## 役割

- App Dashboard の状態を取得する
- Web Dashboard の状態を取得する
- BOOTH未登録、Release未作成、スクショ未作成、README不足を表示する
- Cloudflare未確認、health未確認、Git未反映を表示する
- 次にやる候補を優先順位順に表示する
- 必要な時だけ詳細ノードや対象フォルダを開く

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

v0.2では未処理レーダーとして、既存ノードの正本読み取り結果を使い、未処理件数と次アクションを返します。QPSC起動時に App Dashboard / Web Dashboard / BOOTH Assist を自動起動しません。

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
  "update_summary": "正本からBOOTH未登録、Release未作成、スクショ未作成、Cloudflare未確認を集計する未処理レーダーに対応しました。",
  "folder_name": "DAKE_QPSC_Dashboard",
  "exe_name": "QPSC_Dashboard.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "internal",
  "show_in_launcher": true,
  "show_on_site": false
}
```
