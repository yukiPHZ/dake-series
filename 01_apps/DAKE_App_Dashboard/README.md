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
- フィルタ切替
- 検索
- アプリフォルダを開く
- README.md を開く
- Release URL を開く

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
