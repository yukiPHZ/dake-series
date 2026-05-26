# Dakeメールリスト

Outlookから保存した `.msg` メールをドロップすると、会社名・お名前・メールアドレスだけをCSVにする小さなWindows向けアプリです。

## 使い方

1. Outlookからメールを `.msg` 形式で保存します。
2. `Dakeメールリスト` に `.msg` ファイルをドロップします。
3. デスクトップ、または選択した保存先に `mail_list_YYYYMMDD_HHMMSS.csv` を保存します。

CSVはUTF-8 with BOMで保存します。列は次の3列だけです。

```csv
会社名,お名前,メールアドレス
```

## 仕様

- 複数の `.msg` ファイルをまとめてドロップできます
- `.msg` 以外は無視します
- メールアドレスはFrom情報を優先し、取れない場合だけ本文から最初のメールアドレスを拾います
- お名前はFrom表示名から取得します
- 会社名は本文末尾の署名らしき行から、株式会社・有限会社・合同会社などを含む行だけを控えめに拾います
- 不明な項目は空欄にします
- Outlook本体、外部通信、AI送信は使いません

## 依存ライブラリ

- extract-msg
- tkinterdnd2

`extract-msg` はOutlook `.msg` から送信者・件名・本文などを抽出するPythonライブラリです。PyPI上のメタ情報ではGPLライセンス、Python 3.8以上が案内されています。exe配布時は `extract-msg` と同梱依存ライブラリのライセンス条件を確認してください。

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
  "launcher_title": "メールリスト作成",
  "launcher_description": "Outlookから保存したメールをCSVリストにします。",
  "site_title": "Dakeメールリスト",
  "site_description": "Outlookの.msgメールから、会社名・お名前・メールアドレスだけをCSVにする小さな実務ツールです。",
  "update_summary": "Outlook .msg メールから3列CSVを作成します。",
  "folder_name": "DAKE_Mail_List",
  "exe_name": "DakeMail_List.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## RELEASE_BODY

- Dakeメールリスト
- Outlookの.msgメールをCSV化
- 会社名・お名前・メールアドレスだけ抽出
- ローカル完結のWindows向けexe
