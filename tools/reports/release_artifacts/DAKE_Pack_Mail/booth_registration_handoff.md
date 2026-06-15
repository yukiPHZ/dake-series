# DAKE_Pack_Mail BOOTH Registration Handoff

## 商品名

DAKE メール準備パック

## 価格

780円

## 商品説明

メールを集める、整える、下書きにする。

送信前までの小さな実務をまとめたWindows向けPackです。

Outlookメールから連絡先をCSVにし、メールアドレスを使いやすい形へ整え、CSVから個別メールの下書きを作る3本をまとめました。

収録アプリ:

- Dakeメールリスト
- Dakeメールアドレス整形
- Dakeメール下書き

メールは自動送信しません。作成された下書きをOutlookで確認してから送信できます。

Dakeメール下書きは、Windows版Microsoft Outlook Classicを使用します。New Outlook / Web Outlookでは動作しない場合があります。

## タグ

- メール
- Outlook
- CSV
- Windows
- 実務
- 仕事効率化
- 下書き
- ツール

## 作品ファイル

`04_packs/DAKE_Pack_Mail/pack_ready/DAKE_Pack_Mail.zip`

Pack ZIP:

- size: 56950902 bytes
- sha256: `dfc972b91529161bbf688fbe4fb5bf91b5e27956afe058486a0b5d79ab293ad4`

## メイン画像

`04_packs/DAKE_Pack_Mail/assets/booth_thumbnail.jpg`

## 補助画像

`04_packs/DAKE_Pack_Mail/pack_ready/booth_thumbnail.jpg`

## Outlook Classic注意

Dakeメール下書きはWindows版Microsoft Outlook Classicを前提とします。

New Outlook / Web Outlookでは動作しない場合があります。

メールは自動送信しません。作成された下書きの宛先・件名・本文・添付を確認してから送信してください。

## 登録後に必要な情報

登録後、以下をCodexへ戻してください。

- BOOTH商品URL
- 公開状態
- 表示価格
- 登録ZIP名

BOOTH URLを正本へ記録するコマンド:

```powershell
python tools\record_booth_registration.py DAKE_Pack_Mail --booth-url <BOOTH商品URL>
python tools\record_booth_registration.py DAKE_Pack_Mail --booth-url <BOOTH商品URL> --apply --confirm-product-id DAKE_Pack_Mail --confirmation-text "RECORD BOOTH REGISTRATION DAKE_Pack_Mail"
```
