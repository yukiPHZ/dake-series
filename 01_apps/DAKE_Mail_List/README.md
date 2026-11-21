# Dakeメールリスト

Outlookで選択したメールを直接、または保存済みの `.msg` メールをドロップすると、会社名・お名前・メールアドレスだけをCSVにする小さなWindows向けアプリです。

## 使い方

1. Outlookで1通または複数のメールを選択します。保存済みの `.msg` ファイルも使用できます。
2. `Dakeメールリスト` の「メールをここへドロップ」へ直接ドラッグ＆ドロップします。
3. デスクトップ、または選択した保存先に `mail_list_YYYYMMDD_HHMMSS.csv` を保存します。

CSVはUTF-8 with BOMで保存します。列は次の3列だけです。

```csv
会社名,お名前,メールアドレス
```

## 仕様

- 複数の `.msg` ファイルをまとめてドロップできます
- Outlookで選択した複数メールをWindows native OLEで直接ドロップできます
- `.msg` 以外は無視します
- メールアドレスはFrom情報を優先し、取れない場合だけ本文から最初のメールアドレスを拾います
- お名前は送信者表示名、From表示名、現在の送信者側の署名の順に取得します
- 会社名は返信・転送履歴より前にある現在の送信者側の署名を優先します
- HTML署名の改行や表セルを保ち、会社名・部署・役職・氏名を根拠のある範囲で分離します
- 不明な項目は空欄にします
- Outlook COM/API、外部通信、AI送信は使いません
- Outlook仮想メールの一時 `.msg` は標準ランタイムTEMPへ保存し、処理後に削除します
- 実機確認用の最小診断は `%LOCALAPPDATA%\PEAKHEADZ\DakeMailList\logs\ole_dnd.log` で確認できます。本文・アドレス・氏名・会社名・件名全文は記録しません
- CSV保存先設定は `%LOCALAPPDATA%\PEAKHEADZ\DakeMailList\config.json` に保存します。exe横の旧configは必要な場合だけ新しい場所へ移行し、削除しません

## 依存ライブラリ

- extract-msg
- pywin32

`extract-msg` はOutlook `.msg` を解析し、`pywin32` はWindows native OLEの `IDropTarget`、`CF_HDROP`、`FileGroupDescriptorW`、`FileContents` を扱います。

## ビルド

```bat
build.bat
```

PyInstallerで `dist/DakeMail_List.exe` を作成します。共通アイコン `../../02_assets/dake_icon.ico` を使用します。

## 公式サイト

- DAKEシリーズ公式サイト：https://dakeapp.com
- 関連サイト：https://soredake.com

## 初回起動時の注意

GitHubからダウンロードしたWindows向けexeは、初回起動時にWindowsのセキュリティ確認画面が表示される場合があります。

表示された場合は、内容を確認したうえで、

1. 「詳細情報」をクリック
2. 「実行」をクリック

して起動してください。

これは個人開発アプリや未署名アプリで表示されることがある確認画面です。
不安な場合は、無理に起動せず、公式サイトや配布元を確認してください。

## DAKE_META

```json
{
  "app_key": "dake_mail_list",
  "display_name": "Dakeメールリスト",
  "version": "1.0.1",
  "launcher_title": "メールリスト作成",
  "launcher_description": "Outlookのメールを直接ドロップしてCSVリストにします。",
  "site_title": "Dakeメールリスト",
  "site_description": "Outlookで選択した1通または複数メールを直接ドロップし、会社名・お名前・メールアドレスだけをCSVにする小さな実務ツールです。保存済みの.msgにも対応します。",
  "update_summary": "Outlookメールの直接ドラッグ＆ドロップと複数メール処理に対応しました。",
  "folder_name": "DAKE_Mail_List",
  "exe_name": "DakeMail_List.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Mail_List_v1.0.1",
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

- Dakeメールリスト
- Outlookで選択した1通または複数メールを直接ドラッグ＆ドロップ
- 保存済みの複数.msgにも対応
- 現在の送信者側の署名を優先して会社名とお名前を抽出
- 会社名・お名前・メールアドレスだけ抽出
- UTF-8 BOM、3列固定のCSV
- ローカル完結のWindows向けexe
