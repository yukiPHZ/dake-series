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

## Phase4 BOOTH作業導線

- 次工程がBOOTH登録のアプリを判定
- BOOTHアシストを対象アプリ付きで起動
- BOOTH作業中の対象行を静かな青紫でハイライト
- booth_ready/、booth_product.txt、booth_thumbnail、screenshot、Release URL を詳細ペインから直接開く
- 次の作業へボタンで、BOOTH登録、Release作成、スクショ作成、README整備へ移動

## Phase5 作業発射台リンク導線

- 詳細ペインから各作業対象を直接開く
- フォルダ、README.md、release_body.md、assets、screenshot、booth_product.txt、booth_ready/、dist、exe場所、Release URL に対応
- GitHubリポジトリURLをRelease URLから取得できる場合はGitHubページを開く
- 次の作業へボタンが、BOOTH登録、Release作成、スクショ作成、README整備、BOOTH素材作成に応じて作業対象を開く
- exeは直接実行せず、エクスプローラーで場所表示のみ行う

## Phase6 BOOTH状態表示の精密化

- BOOTH素材を `assets/booth_thumbnail.jpg`、`booth_product.txt`、`booth_ready/` の3点セットで判定
- アプリ一覧で BOOTH 0/3〜3/3 を表示
- 詳細ペインでBOOTH素材の有無と不足項目を表示
- 次工程判定で BOOTH素材作成 と BOOTH登録 を分離
- show_on_site=false の内部アプリはBOOTH未準備集計から除外

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
  "show_in_launcher": false,
  "show_on_site": false
}
```
