# Quiet Personal Cognitive System

QPCS Dashboard は、Quiet Personal Cognitive System 全体の現在地を確認するための最上位レイヤーです。

QPCS Dashboardは、正本から未処理と次にやる候補を返す司令塔です。

App Dashboardは監視ノードであり、DAKE_Web_IndexはWeb資産索引としてQPCS Dashboardの下位に位置します。QPCS Dashboardは関連アプリを自動起動せず、必要な時だけ詳細へ潜る導線を提供します。

## QPCS名称

QPCS = Quiet Personal Cognitive System

記憶、実務、アプリ、Git、Web資産索引などの正本を読み、
現在地と次の行動を返すためのシステム。

## QPCS Dashboard位置付け

QPCS Dashboard は QPCS の司令塔です。

App Dashboard は監視ノードであり、
DAKE_Web_Index はWeb資産の索引としてQPCS Dashboardの下位レイヤーに位置します。

QPCS Dashboard は正本を読み、
現在地と次の行動を返します。

## 役割

- App Dashboard の状態を取得する
- DAKE_Web_Index からWeb資産索引を取得する
- BOOTH登録候補、Release未作成、スクショ未作成、README不足を表示する
- Web資産数と最終更新情報を表示する
- 次に見る候補を表示する
- 必要な時だけ詳細ノードや対象フォルダを開く

## 初版の対象

```text
QPCS Dashboard
├─ App Dashboard
└─ DAKE_Web_Index
```

BRAINZ、OIKAWA、Slack、Git、ORBIT は将来ノードです。初版では内部機能を実装しません。

## 実装方針

App Dashboard と DAKE_Web_Index の `main.py` を動的に読み込み、既存の状態取得関数を利用します。

QPCS Dashboard 側では、Web資産管理、README再解析、通知処理を再実装しません。

v0.2では未処理レーダーとして、既存ノードの正本読み取り結果を使い、未処理件数と次アクションを返します。QPCS起動時に App Dashboard / DAKE_Web_Index / BOOTH Assist を自動起動しません。

v0.3では、アプリ未処理と分類を分け、Web資産は索引として件数と更新情報だけを表示します。

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
QPCS全体の現在地を確認する司令塔
App Dashboard を監視ノード、DAKE_Web_Index をWeb資産索引として集約
内部QPCS用ダッシュボード
Windows向けexe
```

## Positioning

This app is a QPCS/operations app, not a market-facing standalone product. Its completion goal is `system_ready`: README, launch-check, build, and its role inside the QPCS workflow are documented and working.

## app_type / completion_goal 対応

QPCS Dashboard は README 正本の `app_type` と `completion_goal` を読み取り、市場向け・QPCS系・ユキズ専用・凍結を分けて扱います。

市場向けアプリだけを正式出荷ラインの主な未処理として扱います。
QPCS系はシステム稼働または正本提示を完成ゴールとします。
ユキズ専用アプリはローカル運用を完成ゴールとします。
凍結アプリは凍結理由と再開条件が正本化されていれば、凍結完了として扱います。

## レイアウト整理

QPCS Dashboardは、分類と未処理を分けて表示します。
市場向け、QPCS系、ユキズ専用、凍結といった分類情報と、BOOTH登録候補などの未処理情報を混在させず、Web資産は索引情報として現在地を一画面で確認できるようにしています。

## Web Sites

QPCS DashboardはDAKE_Web_Indexと連携します。

DAKE_Web_IndexはWeb資産の索引であり、
通知・管理・優先順位付けは行いません。

QPCS Dashboardでは、
Web資産数と更新情報だけを表示し、
詳細はDAKE_Web_Indexへ委譲します。

## DAKE_META
```json
{
  "app_key": "DAKE_QPSC_Dashboard",
  "display_name": "QPCS Dashboard",
  "launcher_title": "QPCS Dashboard",
  "launcher_description": "QPCS全体の現在地を束ねて確認する司令塔",
  "site_title": "",
  "site_description": "",
  "update_summary": "DAKE_Web_Dashboard連携を廃止し、DAKE_Web_IndexによるWeb資産索引表示へ変更しました。",
  "folder_name": "DAKE_QPSC_Dashboard",
  "exe_name": "QPSC_Dashboard.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "internal",
  "app_type": "qpcs",
  "completion_goal": "system_ready",
  "show_in_launcher": true,
  "show_on_site": false
}
```
