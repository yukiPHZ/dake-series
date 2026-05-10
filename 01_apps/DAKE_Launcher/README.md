# Dakeランチャー

Dakeランチャーは、DAKEシリーズ各アプリを迷わず1クリックで起動する入口アプリです。

ヘルプアプリでも、ダウンロード管理アプリでもありません。ローカルDAKE環境に静かに馴染み、探して、覚えて、起動するための小さなランチャーです。

## 目的

- DAKEシリーズ各アプリを一覧表示して起動します
- 一覧の正本は各アプリフォルダの `README.md` です
- ランチャー側にアプリ一覧を手入力しません
- 別台帳ファイルは作りません

## 一覧生成

`..\` 配下の各アプリフォルダにある `README.md` から `DAKE_META` を読み取り、`show_in_launcher` が `true` のアプリだけを表示します。

読み取る主な項目は以下です。

- `app_key`
- `display_name`
- `launcher_title`
- `launcher_description`
- `folder_name`
- `exe_name`
- `release_url`
- `status`
- `show_in_launcher`

`DAKE_Launcher` 自身は一覧対象から除外します。

## 起動

標準起動パスは以下です。

```text
各アプリフォルダ\dist\exe_name
```

exe が存在する場合は「起動」ボタンで実行します。

exe が未検出の場合は「場所を指定」からユーザーが exe を選択できます。指定した exe パスは以下の config に保存します。

```text
..\..\04_data\configs\DAKE_Launcher_config.json
```

次回以降は、ユーザー指定パス、標準パスの順で起動対象を探します。ユーザー指定パスが存在しなくなった場合は標準パスへ戻します。

## 最近使った

起動に成功したアプリだけを `recent_apps` に記録します。

- 最大2件
- 重複なし
- 最新順

## 守ること

- 勝手に整理しません
- 勝手に移動しません
- 勝手に削除しません
- ダウンロード管理アプリ化しません

## 実行

```bat
python main.py
```

## ビルド

```bat
build.bat
```

PyInstaller で `dist\Dake_Launcher.exe` を作成します。共通アイコンは `..\..\02_assets\dake_icon.ico` を使用します。

## DAKE_META

```json
{
  "app_key": "dake_launcher",
  "display_name": "Dakeランチャー",
  "launcher_title": "DAKEツール",
  "launcher_description": "DAKEシリーズ各アプリを1クリックで起動します。",
  "site_title": "Dakeランチャー",
  "site_description": "DAKEシリーズ各アプリを迷わず1クリックで起動する入口アプリです。",
  "update_summary": "READMEのDAKE_METAを正本として読み取り、DAKEシリーズ各アプリを起動するランチャーを追加。",
  "folder_name": "DAKE_Launcher",
  "exe_name": "Dake_Launcher.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Launcher_v1.0.0",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": false,
  "show_on_site": true
}
```

## RELEASE_BODY

- DAKEシリーズ各アプリを1クリックで起動するランチャー
- 各アプリの `README.md` にある `DAKE_META` から一覧生成
- exe 未検出時の場所指定に対応
- 最近使ったアプリを最大2件まで記録
- 勝手な整理・移動・削除は行いません
