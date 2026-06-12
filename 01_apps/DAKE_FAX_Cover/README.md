# DakeFAX送付状

DAKEシリーズの単機能Windowsデスクトップアプリです。送信先、FAX情報、送信内容、送信者情報を入力して、A4縦のFAX送付状PDFだけを作成します。

正式フォルダ名は `DAKE_FAX_Cover` です。

## できること

- 本日の日付を初期入力し、任意の日付に変更できます。
- 送信先の会社名、担当者名、FAX番号、電話番号を入力できます。
- 件名、送信枚数、送信日を入力できます。
- 送信内容は初期3行で、必要に応じて行を追加できます。
- 空の送信内容行を除外してPDFを作成します。
- 送信者情報と前回保存先を次回起動時に復元します。

## 使い方

1. 送信先、FAX情報、送信内容、メッセージ、送信者情報を入力します。
2. 必要に応じて保存先を選びます。
3. 「FAX送付状PDFを作成」を押します。
4. 完了ダイアログのOK後、保存先フォルダが開きます。

## 設定保存について

`fax_cover_config.json` に送信者情報、前回保存先、前回使用した敬称を保存します。

このファイルはユーザー環境ごとに変わるため、Git管理しません。`.gitignore` で `*_config.json` を除外しています。

## PDF出力について

- 出力形式はPDFです。
- 用紙はA4縦です。
- 白黒印刷や白黒FAXでも読みやすい、罫線中心の業務向けレイアウトです。
- 初期ファイル名は `YYYYMMDD_FAX送付状_宛先会社名.pdf` です。
- 同名ファイルがある場合は末尾に番号を付けて保存します。

## ビルド方法

事前に必要なライブラリをインストールします。

```bat
pip install reportlab pyinstaller
```

アプリフォルダで以下を実行します。

```bat
build.bat
```

生成される実行ファイル名は `DakeFAX_Cover.exe` です。共通アイコン `..\..\02_assets\dake_icon.ico` を使用します。

## 注意事項

- PDF作成に `reportlab` を使用します。
- `fax_cover_config.json` はGit管理対象外です。
- 実際のFAX送信、メール送信、Word出力、Excel出力、宛先履歴、テンプレート複数管理、クラウド同期は行いません。

## 共通仕様レビュー結果

2026-05-06 にDAKE共通仕様へ合わせて横断確認しました。

- フォントは `BIZ UDPGothic` を最優先にし、`Yu Gothic UI`、`Meiryo` の順でフォールバックします。
- ヘッダーは機能タイトルと短い説明文に整理しています。
- フッターは共通文言に統一し、広幅時は左右配置、狭幅時は中央2段配置へ切り替えます。
- UI文言は `UI_TEXT` に集約し、Tkinterの `text=""` に日本語を直接書かない方針で確認済みです。
- 処理中はボタンを無効化し、左下ステータスでドット進行を表示します。
- 共通アイコン `..\..\02_assets\dake_icon.ico` をTkinter起動時とPyInstallerビルド時に参照します。
- `build.bat` によるビルド成功を確認しました。
- `dist\DakeFAX_Cover.exe` が起動後すぐ終了しないことを確認しました。

## メッセージ欄・PDF品質改善

2026-05-07 にメッセージ欄とPDF出力品質を改善しました。

- メッセージ欄を5行表示、折り返し、縦スクロール、内側余白付きの複数行入力にしました。
- 初期メッセージを2行の定型文に変更しました。
- PDF出力時にメッセージの改行が反映されることを確認しました。
- PDFフォントは `BIZ-UDGothic` を優先して登録し、見つからない場合は日本語対応フォントへフォールバックします。
- FAX番号、件名、送信枚数、送信日の視認性を上げました。
- 白黒FAXでも読めるよう、表罫線と情報ブロックの線を薄すぎない設定にしました。
- 複数行メッセージと空欄メッセージのPDF生成を確認しました。
- `build.bat` によるexe再生成と `dist\DakeFAX_Cover.exe` の起動確認を行いました。

## DAKE_META

```json
{
  "app_key": "dake_fax_cover",
  "display_name": "DakeFAX送付状",
  "launcher_title": "FAX送付状",
  "launcher_description": "送信先と内容を入力してA4縦のFAX送付状PDFを作成します。",
  "site_title": "DakeFAX送付状",
  "site_description": "送信先、FAX情報、送信内容、送信者情報を入力して、FAX送付状PDFを作成できるWindows向けアプリです。",
  "update_summary": "READMEメタ情報とRelease本文を整備。スクリーンショットをassets/screenshot.webpに作成。",
  "folder_name": "DAKE_FAX_Cover",
  "exe_name": "DakeFAX_Cover.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_FAX_Cover_v1.0.0",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true,
  "demo_video_path": "release_artifacts/demo.mp4",
  "demo_video_url": "",
  "social_release_path": "release_artifacts/social_release.json"
}
```

## RELEASE_BODY

- FAX送付状PDF作成アプリ
- 送信先・FAX情報・送信内容に対応
- A4縦PDF出力
- Windows向けexe
