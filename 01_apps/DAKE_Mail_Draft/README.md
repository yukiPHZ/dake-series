# Dakeメールドラフト

CSVからOutlook下書きだけ作る、Windows向けの小さな実務補助ツールです。

このアプリは **Outlook下書き作成補助ツール** です。メール送信はしません。作成後は必ずOutlook上で宛先・本文・添付を確認してから、利用者本人の判断で送信してください。

## できること

- CSV名簿から個別のOutlook下書きを作成
- 件名テンプレートと本文テンプレートへの差し込み
- 複数添付ファイルの追加
- Outlook Classic の標準署名を可能な限り保持
- 生成結果を `logs/draft_report_YYYYMMDD_HHMMSS.csv` に出力

## 重要な注意

- 自動送信はしません。
- SMTP送信、メールアカウント保存、パスワード保存はしません。
- CSVの個人情報、テンプレート、添付ファイルはローカルで処理します。外部送信しません。
- 利用には Windows版 Microsoft Outlook Classic が必要です。
- New Outlook / Web Outlook では動作しない可能性があります。
- 利用者は送信前に宛先・本文・添付を必ず確認してください。
- 迷惑メール、無断大量送信、法令・規約違反用途での利用は禁止です。
- 会社・組織のルールに従って利用してください。

## 使い方

1. `data/recipients.example.csv` を参考にCSVを用意します。
2. アプリでCSVを選択します。
3. 会社名列、名前列、メールアドレス列を確認します。
4. 件名テンプレートと本文テンプレートを入力します。
5. 必要な場合だけ、ローカルの添付ファイルを選択します。
6. 生成件数を選びます。まずは50件ずつの生成を推奨します。
7. プレビューで先頭1件の宛先・件名・本文を確認します。
8. 「Outlook下書きを生成」を押します。
9. 生成後、Outlook上で内容・宛先・添付を必ず確認してください。

## CSV列名例

CSVはUTF-8-SIGを推奨します。CP932のCSVも読み込みを試みます。

```csv
会社名,名前,メールアドレス
サンプル株式会社,山田 太郎,taro.yamada@example.com
テスト合同会社,佐藤 花子,hanako.sato@example.com
```

列名は以下を自動候補にします。

- 会社名: `会社名`, `company`, `company_name`, `organization`
- 名前: `名前`, `氏名`, `name`, `person_name`
- メール: `mail 1`, `mail`, `email`, `メール`, `メールアドレス`, `email_address`

空メール、不正なメール形式、重複メールアドレスはスキップされます。重複メールアドレスは既定で1件目のみ下書きを作成します。

## テンプレート

件名・本文では以下のプレースホルダーを使えます。

- `{会社名}`
- `{名前}`
- `{メール}`
- `{company}`
- `{name}`
- `{email}`

テンプレート例:

```txt
{会社名}
{名前} 様

お世話になっております。

以下の内容をご確認ください。

よろしくお願いいたします。
```

本文はHTMLエスケープし、改行を `<br>` に変換してOutlook下書きへ入れます。Outlook署名は、Outlook側で生成された署名HTMLの前に本文を挿入する方式で可能な限り保持します。

## 添付ファイル

- アプリ画面からローカルファイルを選択します。
- 複数添付に対応しています。
- 存在しないパスがある場合、生成前に警告します。
- 添付ファイルはアプリに同梱しません。

## レポート

生成後、`logs/draft_report_YYYYMMDD_HHMMSS.csv` を出力します。

列:

- `row_number`
- `company`
- `name`
- `email`
- `status`
- `message`
- `subject`
- `created_at`

status:

- `drafted`
- `skipped_empty_email`
- `skipped_invalid_email`
- `skipped_duplicate`
- `error`

## ビルド

```bat
pip install -r requirements.txt
build.bat
```

PyInstallerで `dist/DAKE_Mail_Draft.exe` を作成します。共通アイコン `../../02_assets/dake_icon.ico` がある場合は使用し、ない場合もアイコンなしでビルドを継続します。

## トラブルシュート

### Outlookが起動しない

Windows版 Microsoft Outlook Classic がインストールされ、初期設定済みか確認してください。New Outlook / Web Outlook では動作しない場合があります。

### 署名が入らない

Outlook Classic 側で既定署名が設定されているか確認してください。このアプリは下書きを一度表示し、Outlook側に署名を生成させてから本文を挿入します。

### 添付が付かない

添付ファイルのパスが存在するか、ファイルが他のアプリでロックされていないか確認してください。

### pywin32エラー

以下を実行してください。

```bat
pip install -r requirements.txt
```

改善しない場合は、PythonとOutlookのビット数やWindows環境を確認してください。

### New Outlookで動かない

このアプリはOutlook ClassicのCOM操作を利用します。New Outlook / Web Outlook では動作しない場合があります。Outlook Classicに切り替えて確認してください。

## Outlook実機テスト手順

1. Outlook Classicを起動し、既定署名が必要なら設定します。
2. `data/recipients.example.csv` を選択します。
3. 生成件数を `5` にします。
4. 添付なしで下書き生成し、Outlookの下書きフォルダーを確認します。
5. ローカルのテスト用ファイルを1つ添付し、再度1件だけ生成します。
6. 宛先・件名・本文・署名・添付・レポートCSVを確認します。
7. 作成されたメールはOutlook上で内容確認後、不要なら削除してください。

## BOOTH商品説明案

商品名:

DAKE Mail Draft｜CSVからOutlook下書きだけ作る

説明:

CSV名簿をもとに、Outlook Classic の個別メール下書きをまとめて作成する小さな実務補助ツールです。会社名・名前・メールアドレスを件名や本文に差し込み、添付ファイル付きの下書きを作れます。

このアプリはメールを送信しません。作成後は必ずOutlook上で宛先・本文・添付を確認してから送信してください。

BOOTH掲載時に使わない表現:

- メール一括送信
- 自動営業
- スパム送信
- 大量送信ツール
- リスト収集
- 自動送信

## DAKE_META

```json
{
  "app_key": "DAKE_Mail_Draft",
  "display_name": "DAKE Mail Draft",
  "launcher_title": "Mail Draft",
  "launcher_description": "CSVからOutlookの個別下書きだけ作る",
  "site_title": "DAKE Mail Draft",
  "site_description": "CSV名簿から、Outlookの下書きメールを個別に作成する実務補助ツール。",
  "update_summary": "CSV差し込み、Outlook下書き作成、添付、署名保持、生成レポートに対応。",
  "folder_name": "DAKE_Mail_Draft",
  "exe_name": "DAKE_Mail_Draft.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "active",
  "show_in_launcher": true,
  "show_on_site": true
}
```
