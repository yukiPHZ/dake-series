# Dake書類送付状

DAKEシリーズの単機能Windowsデスクトップアプリです。相手方、送付内容、作成者情報を入力して、A4縦の一般的な書類送付状PDFだけを作成します。FAX送付状、Word出力、Excel出力、メール送信などは扱いません。

## できること

- 本日の日付を初期入力し、任意の日付に変更できます。
- 宛先、送付内容、作成者情報を入力できます。
- 送付内容は最低5行で、必要に応じて行を追加できます。
- 本文欄に自由コメントを入力できます。
- 送付内容がない場合でも、本文だけでPDFを作成できます。
- 空の送付内容行を除外してPDFを作成します。
- 作成者情報、前回保存先、前回使用した敬称を次回起動時に復元します。

## 使い方

1. 宛先、本文、作成者情報を入力します。
2. 必要に応じて保存先を選びます。
3. 「PDFを作成」を押します。
4. 完了ダイアログのOK後、保存先フォルダが開きます。

## 設定保存について

`document_cover_config.json` に作成者情報、前回保存先、前回使用した敬称を保存します。

このファイルはユーザー環境ごとに変わるため、Git管理しません。`.gitignore` で `*_config.json` を除外しています。

## PDF出力について

- 出力形式はPDFです。
- 用紙はA4縦です。
- 白背景、黒文字中心の実務向けレイアウトです。
- PDFにはWindowsの日本語フォントを登録して使用します。
- 送付内容が空の場合、送付内容表は出力しません。
- 初期ファイル名は `YYYYMMDD_書類送付状_宛先名.pdf` です。
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

生成される実行ファイル名は `DakeDocument_Cover.exe` です。共通アイコン `..\..\02_assets\dake_icon.ico` を使用します。

## 注意事項

- PDF作成に `reportlab` を使用します。
- `document_cover_config.json` はGit管理対象外です。
- 既存のFAX送付状アプリとは別アプリです。
- テンプレート選択、社判、印影、Word/Excel出力は追加していません。

## DAKE共通仕様レビュー

2026-05-06 にDAKE共通仕様へ合わせてUIを再確認しました。

- BIZ UDPGothic を最優先にし、Yu Gothic UI / Meiryo をフォールバックにしています。
- ヘッダーは機能タイトルと短い説明のみを表示し、画面内にアプリ名を重複表示しません。
- フッターはシリーズコピー、リンク、コピーライトをDAKE共通構成にしています。
- 狭い幅ではフッターが中央寄せ2段構成に切り替わります。
- リンクは通常時に補助文字色、ホバー時のみアクセント色で表示します。
- UI文言は `UI_TEXT` に集約し、日本語の `text=""` 直書きがないことを確認しています。
- `build.bat` で `DakeDocument_Cover.exe` を再生成し、短時間起動確認済みです。

## 本文欄とPDFフォント

2026-05-07 に本文欄とPDFフォントを調整しました。

- 本文欄は必須入力です。
- 送付内容は任意です。
- 送付内容がない場合は本文だけの送付状としてPDFを作成します。
- 送付内容がある場合のみ、本文の下に送付内容表を出力します。
- PDFフォントは `BIZ UDPGothic`、`BIZ UDGothic`、`Yu Gothic UI / Yu Gothic`、`Meiryo` 系を順番に探索します。
- 日本語フォントが見つからない場合はPDF作成を止め、エラーを表示します。

## DAKE_META

```json
{
  "app_key": "dake_document_cover",
  "display_name": "Dake書類送付状",
  "launcher_title": "書類送付状",
  "launcher_description": "相手方と本文を入力してA4送付状PDFを作成します。",
  "site_title": "Dake書類送付状",
  "site_description": "宛先、送付内容、作成者情報を入力して、A4縦の書類送付状PDFを作成できるWindows向けアプリです。",
  "update_summary": "READMEメタ情報とRelease本文を整備。スクリーンショットをassets/screenshot.webpに作成。",
  "folder_name": "DAKE_Document_Cover",
  "exe_name": "DakeDocument_Cover.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Document_Cover-v1.0.0",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## RELEASE_BODY

- 書類送付状PDF作成アプリ
- 宛先・送付内容・作成者情報に対応
- A4縦PDF出力
- Windows向けexe
