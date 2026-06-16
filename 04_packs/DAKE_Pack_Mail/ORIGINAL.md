# ORIGINAL.md

## 正本宣言

このファイルは `DAKE_Pack_Mail` のPack商品の真の正本です。
README.md、PACK_META、pack_manifest.json、booth_product.txt、pack_ready、BOOTH登録情報、Store表示などは、このファイルから派生するビュー、またはこのファイルを補助する参照情報です。

## 基本情報

- pack_id: `DAKE_Pack_Mail`
- title: DAKE メール準備パック
- short_title: DAKE メール準備パック
- category: メール実務Pack
- pack_type: mail_pack
- status: available
- version: 1.0.0
- price: 780円
- currency: JPY
- target_platform: Windows
- booth_url: https://peakheadz.booth.pm/items/8457085
- distribution: BOOTH / manual private download
## Packの目的

メール作業のうち、送信前に発生する小さな準備作業をまとめて扱えるようにする。

## キャッチ

集める、整える、下書きにする。

送信前までを静かにまとめるPack。

## 商品説明

Outlookメールから連絡先を一覧にし、メールアドレスを整え、CSVから個別メールの下書きを作る3本をまとめたPackです。

メールは自動送信しません。作成された下書きをOutlookで確認してから送信できます。

## 対象ユーザー

- WindowsでOutlookを使う実務担当者
- メールの宛先整理を手作業で行っている人
- 複数人分の個別メール下書きを安全に準備したい人
- 自動送信ではなく、送信前に人間確認を残したい人

## 解決する困りごと

```txt
Outlookメールから連絡先をCSVにする
↓
宛先表記を使いやすい形へ整える
↓
CSVからOutlookの個別下書きを作る
↓
人間が内容を確認して送信する
```

## Packとしての価値

メール送信そのものではなく、送信前の準備をまとめるPackです。宛先整理、CSV作成、下書き作成を分けて使えるため、最終送信の判断を人間に残せます。

## 同梱アプリ・構成物

- Dakeメールリスト (DAKE_Mail_List): Outlookの.msgメールから会社名・氏名・メールアドレスをCSV化する
  - ZIP: `DakeMail_List.zip` / 32925063 bytes / SHA256 `1012c378bb28d62b1c7fec9d50f5de966c606a9ff44a74b60e20854bfebf8bc1`
  - BOOTH: https://peakheadz.booth.pm/items/8448014
  - GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Mail_List_v1.0.0
- Dakeメールアドレス整形 (DAKE_Mail_Address_Format): 名前付き表記、改行、セミコロン等が混在したメールアドレスを抽出・整形する
  - ZIP: `DakeMail_Address_Format.zip` / 10106704 bytes / SHA256 `3fe241b58b95de37ae4b9d08b9fb6096e485fbf72d28d9926cb43120267d53de`
  - BOOTH: https://peakheadz.booth.pm/items/8447996
  - GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Mail_Address_Format_v1.0.0
- Dakeメール下書き (DAKE_Mail_Draft): CSV名簿からOutlookの個別下書きを作成する
  - ZIP: `DakeMail_Draft.zip` / 13902128 bytes / SHA256 `d6d691d890eb8d5b3fa903765f756a6e66d666400a228eebdb3135174a55724e`
  - BOOTH: https://peakheadz.booth.pm/items/8448001
  - GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Mail_Draft_v1.0.0

## 単品販売との差分

単品3本を別々に探す代わりに、メール準備の流れとしてまとめて導入できます。価格は単品合計900円に対し、Pack価格780円です。

## 含めるもの

- Dakeメールリスト
- Dakeメールアドレス整形
- Dakeメール下書き
- README.txt
- 注意事項.txt

## 含めないもの

- DAKE_Mail_AllStaff: 組織固有の用途を含むため、汎用販売Packには収録しない。
- DAKE_Mail_Kikuta: 利用者固有の用途を含むため、汎用販売Packには収録しない。
- ソースコード、build、dist、spec、venv、APIキー、Secret、個人情報、実メールアドレス、実CSV名簿、Outlookアカウント情報、ログ、一時ファイル。

## ZIP構成

```txt
DAKE_Pack_Mail.zip
├─ README.txt
├─ 注意事項.txt
└─ apps/
   ├─ DAKE_Mail_List/
   │  └─ DakeMail_List.zip
   ├─ DAKE_Mail_Address_Format/
   │  └─ DakeMail_Address_Format.zip
   └─ DAKE_Mail_Draft/
      └─ DakeMail_Draft.zip
```

- zip名: `DAKE_Pack_Mail.zip`
- Pack zip path: `04_packs/DAKE_Pack_Mail/pack_ready/DAKE_Pack_Mail.zip`
- Pack zip size: 56950902 bytes
- Pack zip sha256: `dfc972b91529161bbf688fbe4fb5bf91b5e27956afe058486a0b5d79ab293ad4`

## 公開用説明

メールを集める、整える、下書きにする。送信前までの小さな実務をまとめたWindows向けPackです。

## README生成用情報

- 概要: Outlookメールから連絡先を一覧にし、メールアドレスを整え、CSVから個別メールの下書きを作る3本をまとめたPackです。
- 使い方: Pack ZIPを解凍し、apps内の各アプリZIPを解凍して使います。
- 注意: メールは自動送信しません。下書きは送信前に必ず人間が確認してください。
- Outlook要件: Dakeメール下書きはWindows版Microsoft Outlook Classic前提です。New Outlook / Web Outlookでは動作しない場合があります。

## PACK_META生成用情報

```json
{
  "folder_name": "DAKE_Pack_Mail",
  "display_name": "DAKE メール準備パック",
  "product_type": "pack",
  "pack_type": "mail_pack",
  "status": "available",
  "price": 780,
  "currency": "JPY",
  "booth_url": "https://peakheadz.booth.pm/items/8457085",
  "included_apps": ["DAKE_Mail_List", "DAKE_Mail_Address_Format", "DAKE_Mail_Draft"],
  "show_in_booth_assist": true,
  "show_on_dashboard": true,
  "tags": ["メール", "Outlook", "CSV", "Windows", "実務", "仕事効率化", "下書き", "ツール"],
  "summary": "Outlookメールから連絡先を集め、宛先を整え、CSVから個別メール下書きを作るPackです。",
  "copyright": "PEAKHEADZ"
}
```

## booth_product生成用情報

- 商品名: DAKE メール準備パック
- 価格: 780円
- 商品紹介文: Outlookメールから連絡先をCSVにし、メールアドレスを使いやすい形へ整え、CSVから個別メールの下書きを作る3本をまとめました。
- タグ: メール / Outlook / CSV / Windows / 実務 / 仕事効率化 / 下書き / ツール
- 商品画像: assets/booth_thumbnail.jpg
- 作品ファイル: `04_packs/DAKE_Pack_Mail/pack_ready/DAKE_Pack_Mail.zip`

## Store表示用情報

Storeにはまだ掲載しない。BOOTH登録後、BOOTH URLを正本へ記録してから次工程で派生ビューを更新する。

## 価格・販売方針

- Pack価格: 780円
- 単品合計: 900円
- 初期販売状態: BOOTH登録待ち
- 初期payment_status: preparing
- 初期stripe_payment_link: 未設定

## 配布・ダウンロード方針

- 配布物: `04_packs/DAKE_Pack_Mail/pack_ready/DAKE_Pack_Mail.zip`
- BOOTH登録後はBOOTH商品ページで配布する。
- Stripe販売を始める場合は、既存Packと同じ手動配布ルールを使用する。

## 免責・注意事項

- Windows向けです。
- Dakeメール下書きはWindows版Microsoft Outlook Classicを前提とします。
- New Outlook / Web Outlookでは動作しない場合があります。
- メールは自動送信しません。
- 生成した下書きの宛先・件名・本文・添付は送信前に必ず確認してください。
- 迷惑メール、無断送信、法令・規約違反用途には使用しないでください。
- CSV、メールデータ、連絡先情報の管理は利用者の責任で行ってください。
- 第三者への再配布は禁止します。

## Checkout購入前表示

- checkout_notice_required: yes
- checkout_notice_target: submit
- checkout_notice_version: 1
- checkout_product_notice: Windows向けです。Dakeメール下書きはWindows版Microsoft Outlook Classicを使用します。New Outlook / Web Outlookでは動作しない場合があります。メールは自動送信されません。作成された下書きの宛先・件名・本文・添付を確認してから、利用者自身で送信してください。

## 手動配布方針

Stripe販売を開始する場合は、既存Packと同じ共通ルールを使用する。

- purchase_delivery_method: manual_email_private_download
- purchase_delivery_ready: yes
- delivery_sla: 決済確認後、次営業日以内
- delivery_rule: `00_core/DAKE_PACK_MANUAL_DELIVERY_RULE.md`
- delivery_email_template: `tools/templates/stripe_pack_manual_delivery_email.txt`
- delivery_log_template: `tools/templates/stripe_pack_manual_delivery_log.example.csv`

### Buyer notice

本商品は自動ダウンロードではありません。

Stripeでの決済完了後、購入時に入力されたメールアドレス宛に、次営業日以内にダウンロード方法をご案内します。

This is not automatic download. The standard delivery window is within the next business day after payment confirmation.

購入者から再送依頼があった場合は、購入Pack名、購入時メールアドレス、購入日時、Stripe決済を識別できる情報を確認して対応します。カード番号、セキュリティコード、パスワードなどの機微情報は求めません。

### Manual delivery procedure

1. Confirm the Stripe payment completion in Stripe Dashboard.
2. Match the payment to this Pack and the expected price.
3. Confirm the current delivery file path, size, and SHA256.
4. Send download instructions to the purchase-time email address.
5. Do not expose private download URLs in public files.
6. Record delivery status in a secure local delivery log outside Git.

### Resend and failure handling

If the buyer requests resend, verify the original payment, Pack, buyer email address, and previous delivery record before resending. Do not send public download URLs.

If email delivery fails, record `delivery_failed`, confirm the email address and payment information, and handle refund or individual support according to existing DAKE Store policy.

Do not store buyer email addresses, payment IDs, private download URLs, Stripe Secret Keys, Webhook Secrets, card information, or real delivery logs in Git, generated JSON, Store static files, Markdown reports, or public folders.

## Stripe情報

- payment_status: stripe_ready
- stripe_payment_link: https://buy.stripe.com/7sY14ncA90NN1q84rv0gw0Q
- stripe_creation_method: manual_dashboard_ready
- review_result: ready
- tax_code_candidate: txcd_10202003

## 配布ファイル証跡

- pack_id: `DAKE_Pack_Mail`
- pack_title: DAKE メール準備パック
- delivery_file: `DAKE_Pack_Mail.zip`
- delivery_path: `04_packs/DAKE_Pack_Mail/pack_ready/DAKE_Pack_Mail.zip`
- delivery_file_size: 56950902 bytes
- delivery_file_sha256: `dfc972b91529161bbf688fbe4fb5bf91b5e27956afe058486a0b5d79ab293ad4`
- included_apps: DAKE_Mail_List, DAKE_Mail_Address_Format, DAKE_Mail_Draft

## 更新履歴

- 2026-06-16: Phase 17B-1でBOOTH登録待ちPackとして新規作成。
