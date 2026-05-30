# DAKE Dashboard

DAKE Dashboard は、DAKEシリーズ配下のアプリフォルダを読み取り、README.md の `DAKE_META` を正本として現在の状態を一覧・集計・確認するための内部用アプリです。

一般公開、BOOTH販売、dakeapp.com掲載、GitHub Release作成は行いません。

## 役割

- DAKEシリーズ管理端末
- 補助脳BRAINZ / QPSC系の開発支援アプリ
- README正本可視化アプリ
- Notion手管理からの移行補助

## 読み取り対象

```text
C:\Users\yukiz\devlop\DAKE_series\01_apps
```

各アプリフォルダの `README.md` から `DAKE_META` JSON を読み取り、実在するファイルと照合します。

## 判定する情報

- README.md
- release_body.md
- assets/screenshot.webp
- assets/booth_thumbnail.jpg
- booth_product.txt
- booth_ready/
- dist/*.exe
- release_url

## 操作

- 再読み込み
- 起動時自動読込
- 30秒ごと自動読込
- watchdog による 01_apps 配下の変更監視
- フィルタ切替
- 検索
- アプリフォルダを開く
- README.md を開く
- Release URL を開く

## Phase2 監視対象

- README.md
- release_body.md
- booth_product.txt
- assets/
- dist/

変更検知時は対象アプリを再読込し、右下通知に `{folder} を再読込しました` を表示します。
`watchdog` が未導入の環境でもアプリは落ちず、30秒ごとの自動読込で追従します。

## Phase3 司令塔機能

- 正式出荷ライン達成率を表示
- show_on_site=false の内部アプリは「内部アプリ」として出荷率の対象外にする
- Git状態カードで branch、latest、未コミット、未追跡、push/pull待ち、Dashboard状態を表示
- 次にやる候補を最大5件表示
- 要確認、未出荷、BOOTH未準備、Release未作成をQPSC通知カードで浮上

このアプリ自身は内部アプリです。一般公開、BOOTH素材作成、dakeapp.com掲載、GitHub Release作成は行いません。

## 品質チェック

```powershell
python -m py_compile main.py
python main.py --launch-check
.\build.bat
.\dist\DakeApp_Dashboard.exe --launch-check
git diff --check
```

## DAKE_META

```json
{
  "app_key": "DAKE_App_Dashboard",
  "display_name": "DAKE Dashboard",
  "launcher_title": "DAKE Dashboard",
  "launcher_description": "DAKEアプリ群の正本状態を見る",
  "site_title": "",
  "site_description": "",
  "update_summary": "README正本からDAKEアプリ群の状態を確認する開発アシスト",
  "folder_name": "DAKE_App_Dashboard",
  "exe_name": "DakeApp_Dashboard.exe",
  "release_url": "",
  "screenshot_path": "",
  "status": "internal",
  "show_in_launcher": true,
  "show_on_site": false
}
```
