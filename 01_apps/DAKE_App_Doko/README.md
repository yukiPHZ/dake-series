# Dakeアプリどこ

PC内に散らばったDAKE系exeを探し、忘れていたアプリへ再接続するための探索アプリです。

## 役割

DAKE_アプリどこ は、DAKE_Launcher の代替ではありません。

DAKE_Launcher は「登録済み・現役・いつも使うDAKEを起動するアプリ」です。

DAKE_アプリどこ は「PC内に散らばったDAKE系exeを探し、忘れていたアプリへ再接続するアプリ」です。

思想は「起動」ではなく「探索・発掘・再接続」です。

## DAKE_Launcherとの住み分け

- DAKE_Launcher：登録済みアプリを起動する
- DAKE_アプリどこ：PC内に散らばったDAKE系exeを探す
- DAKE_Launcherの一覧には通常表示してよいが、役割は「探索補助」
- Launcherの内部データや表示ロジックは変更しない
- Launcherへ統合しない
- Launcherの代替にしない

## 使い方

1. `アプリを探す` を押します。
2. PC内の候補フォルダから `Dake*.exe` / `DAKE*.exe` を探します。
3. 見つかったアプリの表示名、説明、exeパス、README検出、最終更新日時、最終起動日時を確認します。
4. 必要に応じて `起動`、`フォルダを開く`、`パスをコピー` を使います。
5. `最近使っていないアプリ` を押すと、未起動または最終起動日時が古いアプリを上位に表示します。

## 探索対象

初期探索対象は以下です。存在しないパスはスキップします。

- `C:\Users\yukiz\Downloads`
- `C:\Users\yukiz\Desktop`
- `C:\Users\yukiz\Documents`
- `D:\`
- `C:\Users\yukiz\devlop\DAKE_series\01_apps`

以下のフォルダは探索から除外します。

- `build`
- `__pycache__`
- `.git`
- `node_modules`
- `venv`
- `.venv`

## README正本の読み取り

exeと同階層、または親フォルダに `README.md` がある場合は読み取り、`DAKE_META` を抽出します。

`DAKE_META` がある場合は、以下を表示名や説明へ反映します。

- `display_name`
- `launcher_title`
- `launcher_description`
- `site_description`
- `exe_name`
- `folder_name`

`DAKE_META` がない場合は、exe名から表示名を推定します。

## 注意事項

- 探索は別スレッドで実行し、UIが固まらないようにしています。
- 検索中はキャンセルできます。キャンセル時は途中結果を表示します。
- 起動履歴はローカルファイル `DAKE_App_Doko_config.json` に保存します。
- `DAKE_App_Doko_config.json` はユーザー環境依存のためGit管理しません。
- Launcherの内部データや表示ロジックは変更しません。

## ビルド方法

```bat
build.bat
```

`dist\DakeApp_Doko.exe` が作成されます。

## 品質チェック

```powershell
python -m py_compile main.py
python main.py --launch-check
python ..\..\00_core\dake_quality_engine\quality_check.py --app DAKE_App_Doko
build.bat
dist\DakeApp_Doko.exe --launch-check
```

## DAKE_META
```json
{
  "app_key": "DAKE_App_Doko",
  "display_name": "アプリどこ",
  "launcher_title": "アプリどこ",
  "launcher_description": "PC内に散らばったDAKE系アプリを探します。",
  "site_title": "アプリどこ",
  "site_description": "どこに置いたか忘れたDAKE系アプリを探して、起動やフォルダ表示まで行う探索アプリです。",
  "update_summary": "PC内のDAKE系exe探索、README正本読み取り、起動履歴による発掘表示に対応。",
  "folder_name": "DAKE_App_Doko",
  "exe_name": "DakeApp_Doko.exe",
  "version": "1.0.0",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## RELEASE_BODY

```text
アプリどこ
PC内のDAKE系exeを探します
README正本を読み取って表示します
忘れていたアプリを発掘できます
Windows向けexe
```
