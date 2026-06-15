# ORIGINAL.md

## 正本宣言

このファイルは `DAKE_Pack_Document` のPack商品の真の正本です。

README.md、PACK_META、pack_manifest.json、booth_product.txt、pack_ready、Store表示などは、このファイルから派生するビュー、またはこのファイルを補助する参照情報です。

単品アプリ側の `ORIGINAL.md` は、各アプリ自体の真の正本です。

Pack側の `ORIGINAL.md` は、複数アプリ・素材・説明・価格・zip構成を束ねる商品としての真の正本です。

## 基本情報

- pack_id: `DAKE_Pack_Document`
- title: DAKE 書類整理パック
- short_title: DAKE 書類整理パック
- category: 書類整理Pack
- status: available
- version: 未設定（既存Pack情報に明示バージョンなし）
- price: 1480円
- distribution: BOOTH / Store（Storeは未確定）
- target_platform: Windows
- pack_type: document_pack

## Packの目的

PDFや画像の小さな書類整理作業を、ひとつの入口から扱いやすくするためのPack。

## 対象ユーザー

日常業務でPDFや画像の小さな書類整理をまとめて処理したい方。

## Packとして解決する困りごと

PDF結合、画像PDF化、画像リサイズ、A4貼り付けを個別に探す手間を減らし、書類整理の流れを止めにくくする。

## Packとしての価値

- 書類整理で一緒に使いやすい4本をまとめる。
- PDFと画像まわりの軽作業をPackとして導入しやすくする。
- 各アプリの詳細は単品アプリ側ORIGINALに任せ、Pack側では構成と販売導線を正本化する。

## 同梱アプリ・構成物

| item | type | folder | role | original |
|---|---|---|---|---|
| DakePDF結合 | app | `DAKE_PDF_Merge` | 複数PDFをひとつにまとめる。 | `01_apps/DAKE_PDF_Merge/ORIGINAL.md` |
| DakeImageToPDF | app | `DAKE_Image_ToPDF` | 画像ファイルをPDF化する。 | `01_apps/DAKE_Image_ToPDF/ORIGINAL.md` |
| Dake画像リサイズ | app | `DAKE_Image_Resize` | 画像サイズをまとめて調整する。 | `01_apps/DAKE_Image_Resize/ORIGINAL.md` |
| 貼る | app | `DAKE_Image_PasteA4` | 画像をA4用紙に配置して印刷しやすくする。 | 未作成（当面は `01_apps/DAKE_Image_PasteA4/README.md` を暫定参照） |

## 単品販売との差分

- 単品販売との価格差: 要確認（単品価格の合計値はこのPack正本では再計算しない）
- Pack限定の同梱物: 収録アプリZIPをまとめたPack ZIP、Pack用README、Pack用注意事項、Pack商品画像
- Pack限定の説明: Packとしての概要、収録アプリ一覧、Pack ZIPの使い方
- Pack限定ではないもの: 各アプリの機能、各アプリのRelease URL、各アプリのBOOTH URL

## Packに含めるもの / 含めないもの

### 含めるもの

- 収録アプリの配布zip
- Pack用README
- Pack用注意事項
- Pack商品画像

### 含めないもの

- ソースコード
- build/
- dist/
- *.spec
- 個人設定ファイル
- APIキーや個人情報

## zip構成

- zip名: `DAKE_Pack_Document.zip`
- Pack zip path: `04_packs/DAKE_Pack_Document/pack_ready/DAKE_Pack_Document.zip`
- Pack zip size: 124859888 bytes
- Pack zip sha256: `83b4be666bafd907de79884a698fce5e98101123d7fef5a942eaf6d0ea3f72b0`
- 同梱exe: 収録アプリZIP内のexeを使用
- 同梱README: `README.txt`
- 注意事項: `注意事項.txt`
- 入れないもの: ソースコード、build、dist、spec、個人設定ファイル、APIキー、個人情報

### 収録アプリZIP

- `apps/DAKE_PDF_Merge/DakePDF_Merge.zip`
- `apps/DAKE_Image_ToPDF/DakeImageToPDF.zip`
- `apps/DAKE_Image_Resize/Dake_Image_Resize.zip`
- `apps/DAKE_Image_PasteA4/DakeImage_PasteA4.zip`

## 公開用説明の元情報

PDF結合、画像PDF化、画像リサイズ、A4貼り付けをまとめた書類整理向けパックです。

PDFの結合、画像からPDF作成、画像リサイズ、画像のA4貼り付けをまとめたパックです。
書類整理でよくある小さな作業を、ひとつずつ軽く処理できます。

### できること

- 複数PDFをひとつにまとめる
- 画像ファイルをPDFにする
- 画像サイズをまとめて調整する
- 画像をA4用紙に配置して印刷しやすくする

### 使い方

1. パックZIPを解凍します。
2. 使いたいアプリのZIPを解凍します。
3. 各アプリの `.exe` を起動して使います。

### 注意事項

- Windows向けです。
- 大切なファイルは事前にバックアップしてください。
- 各アプリの詳しい使い方は、収録アプリ側のREADMEを確認してください。

## README生成用情報

- 概要: PDF結合、画像PDF化、画像リサイズ、A4貼り付けをまとめた書類整理向けパックです。
- 収録内容: DakePDF結合, DakeImageToPDF, Dake画像リサイズ, 貼る
- 使い方: Pack ZIPを解凍し、appsフォルダ内の各アプリZIPを解凍して、使いたいアプリのexeを起動する。
- 必要なもの: Windows環境
- 注意: 各アプリの詳しい使い方は、収録アプリ側のREADMEまたはORIGINALを確認する。

## PACK_META生成用情報

```json
{
  "folder_name": "DAKE_Pack_Document",
  "display_name": "DAKE 書類整理パック",
  "product_type": "pack",
  "status": "available",
  "version": "未設定",
  "price": 1480,
  "booth_url": "https://peakheadz.booth.pm/items/8448353",
  "included_apps": [
    "DAKE_PDF_Merge",
    "DAKE_Image_ToPDF",
    "DAKE_Image_Resize",
    "DAKE_Image_PasteA4"
  ],
  "show_in_booth_assist": true,
  "show_on_dashboard": true,
  "tags": [
    "PDF",
    "画像",
    "書類整理",
    "Windows",
    "実務",
    "ツール"
  ],
  "summary": "PDF結合、画像PDF化、画像リサイズ、A4貼り付けをまとめた書類整理向けパックです。",
  "copyright": "PEAKHEADZ"
}
```

## booth_product生成用情報

- 商品名: DAKE 書類整理パック
- 価格案: 1480円
- 商品紹介文: PDF結合、画像PDF化、画像リサイズ、A4貼り付けをまとめた書類整理向けパックです。
- 補足紹介文:

```text
PDFの結合、画像からPDF作成、画像リサイズ、画像のA4貼り付けをまとめたパックです。
書類整理でよくある小さな作業を、ひとつずつ軽く処理できます。

### できること

- 複数PDFをひとつにまとめる
- 画像ファイルをPDFにする
- 画像サイズをまとめて調整する
- 画像をA4用紙に配置して印刷しやすくする

### 使い方

1. パックZIPを解凍します。
2. 使いたいアプリのZIPを解凍します。
3. 各アプリの `.exe` を起動して使います。

### 注意事項

- Windows向けです。
- 大切なファイルは事前にバックアップしてください。
- 各アプリの詳しい使い方は、収録アプリ側のREADMEを確認してください。
```

- タグ:
  - PDF
  - 画像
  - 書類整理
  - Windows
  - 実務
  - ツール
- 商品画像: `assets/booth_thumbnail.jpg`
- 補助画像: `pack_ready/booth_thumbnail.jpg`
- 作品ファイル: `04_packs/DAKE_Pack_Document/pack_ready/DAKE_Pack_Document.zip`
- BOOTH URL: https://peakheadz.booth.pm/items/8448353

## Store表示用情報

- 商品名: DAKE 書類整理パック
- キャッチ: PDFと画像の書類整理を、まとめて静かに進めるPack。
- 説明: PDF結合、画像PDF化、画像リサイズ、A4貼り付けをまとめた書類整理向けパックです。
- 価格: 1480円
- 画像: `assets/booth_thumbnail.jpg`
- ダウンロード導線: 未確定
- サポート方針: 未確定

## 価格・販売方針

Pack価格は 1480円。

BOOTHではPack商品として公開済み。Store掲載時の価格、決済、ダウンロード導線は未確定。

Pack価格はPack側の正本情報であり、構成アプリ側のORIGINALへ書き戻さない。

## 配布・ダウンロード方針

- BOOTH URL: https://peakheadz.booth.pm/items/8448353
- Pack ZIP: `04_packs/DAKE_Pack_Document/pack_ready/DAKE_Pack_Document.zip`
- Store: 未確定
- 構成アプリ単品の配布導線:
- DakePDF結合: BOOTH `https://peakheadz.booth.pm/items/8448196` / Release `https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_PDF_Merge_v1.0.0`
- DakeImageToPDF: BOOTH `https://peakheadz.booth.pm/items/8397610` / Release `https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Image_ToPDF_v1.0.0`
- Dake画像リサイズ: BOOTH `https://peakheadz.booth.pm/items/8397622` / Release `https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Image_Resize_v1.0.0`
- 貼る: BOOTH `https://peakheadz.booth.pm/items/8397628` / Release `https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Image_PasteA4_v1.0.0`

## 免責・注意事項

```text
【注意事項】

・このZIPはDAKEシリーズのパック商品です
・収録アプリは、それぞれのZIPを解凍してから起動してください
・Windows向けアプリです
・大切なファイルは事前にバックアップしてください
・本パックおよび収録アプリの無断転載・再配布を禁止します
・環境によっては起動時にWindowsの警告が表示される場合があります

PEAKHEADZ
https://peakheadz.com
```

## 同梱ファイル方針

- apps/: 収録アプリZIPを配置する。
- README.txt: Pack概要、収録アプリ、使い方を短く記載する。
- 注意事項.txt: Windows向け、バックアップ推奨、無断転載禁止、Windows警告の可能性を記載する。
- 画像: Pack商品画像をPack素材として保持する。
- 入れないもの: ソースコード、build、dist、spec、個人設定、APIキー、個人情報。

## スクリーンショット・画像方針

- assets/booth_thumbnail.jpg: Pack商品画像の正本候補。
- pack_ready/booth_thumbnail.jpg: BOOTH登録時に実使用するPack商品画像。
- 補助画像: 既存ファイルに記載なし。
- Store用画像: 未確定。

## 更新時の扱い

- 構成アプリ更新時: 各アプリのORIGINAL、Release、BOOTH readyを確認してからPack再生成を検討する。
- Pack zip更新時: `pack_manifest.json`、Pack ZIP、pack_ready、booth_productの整合を確認する。
- BOOTH更新時: BOOTH URLはPack側booth_productとPack ORIGINALへ戻す。
- Store更新時: Store用generatedはPack ORIGINAL由来にする。
- バージョン表記: 未設定。次回Pack更新時にPack versionの扱いを決める。

## 構成アプリ側 ORIGINAL.md との関係

Pack側はPack商品の正本であり、構成アプリ自体の正本ではない。

- Pack側で持つ情報: Pack名、Pack価格、Pack価値、構成、Pack ZIP、販売文、Pack画像、Pack更新方針
- 各アプリORIGINALへ戻すべき情報: 個別アプリの詳細機能、注意事項、Release URL、BOOTH URL、スクリーンショット方針
- 重複させない情報: 各アプリの細かい機能説明や操作説明
- 参照するORIGINAL: 上記の同梱アプリ・構成物テーブルを参照
- ORIGINAL未作成の構成アプリ: DAKE_Image_PasteA4

## Pack側で正本化してよい情報

- Pack名
- Pack価格
- Packとしての価値
- Packに含める構成物
- Pack zip構成
- Pack販売文
- Pack画像方針
- Pack更新方針

## 各アプリORIGINALへ戻すべき情報

- 個別アプリの詳しい機能説明
- 個別アプリの注意事項
- 個別アプリのRelease URL
- 個別アプリのBOOTH URL
- 個別アプリのスクリーンショット方針
- 個別アプリの今後の改善予定

## やらないこと / 非ゴール

- Pack側で構成アプリの機能説明を再定義しない。
- Packにソースコードを同梱しない。
- Pack価格を各アプリ側ORIGINALへ書かない。
- SHIMARISU Packの情報をこのPackに混ぜない。

## 今後の改善予定

- Pack versionの扱いを決める。
- Store掲載時のダウンロード導線とサポート方針を決める。
- ORIGINAL未作成の構成アプリがある場合は、単品アプリ側でORIGINALを導入する。

## Codex作業時の注意

- 触ってよい: Pack ORIGINAL、Pack正本から派生させるビュー、Packレポート
- 触らない: Pack ZIP、pack_ready、assets、構成アプリ本体、SHIMARISU関連、Store/Stripe実装
- 外部公開しない: Packに含めない内部メモ、個人情報、APIキー、ソースコード
- 自動操作しない: BOOTH公開ボタン、Store決済、外部アップロード

## 派生物一覧

- README.md: PackのGitHub公開用ビュー。
- PACK_META: README内のPack機械利用ビュー。
- pack_manifest.json: Pack生成結果・同梱ZIP・ハッシュ確認用ビュー。
- booth_product.txt: BOOTH登録用ビュー。現状は `pack_ready/booth_product.txt` が実使用ファイル。
- pack_ready: BOOTH/配布用Pack成果物フォルダ。
- Store: 未構築。Store専用正本は作らず、Pack ORIGINAL由来のgeneratedを読む方針。

## 参照した既存ファイル

- `README.md`
- `pack_manifest.json`
- `pack_ready/booth_product.txt`
- `pack_ready/README.txt`
- `pack_ready/注意事項.txt`
- `assets/booth_thumbnail.jpg`
- `pack_ready/booth_thumbnail.jpg`
- `pack_ready/DAKE_Pack_Document.zip`

## Stripe manual delivery operation

This section defines the temporary manual fulfillment policy for Stripe sales of this Pack.

- payment_status: booth_only
- stripe_payment_link: not set
- purchase_delivery_method: manual_email_private_download
- purchase_delivery_ready: yes
- stripe_creation_method: manual_dashboard_ready
- review_result: ready
- delivery_rule: `00_core/DAKE_PACK_MANUAL_DELIVERY_RULE.md`
- delivery_email_template: `tools/templates/stripe_pack_manual_delivery_email.txt`
- delivery_log_template: `tools/templates/stripe_pack_manual_delivery_log.example.csv`

### Delivery file

- pack_id: `DAKE_Pack_Document`
- pack_title: DAKE 書類整理パック
- delivery_file: `DAKE_Pack_Document.zip`
- delivery_path: `04_packs/DAKE_Pack_Document/pack_ready/DAKE_Pack_Document.zip`
- delivery_file_size: 124859888 bytes
- delivery_file_sha256: `83b4be666bafd907de79884a698fce5e98101123d7fef5a942eaf6d0ea3f72b0`
- included_apps: DAKE_PDF_Merge, DAKE_Image_ToPDF, DAKE_Image_Resize, DAKE_Image_PasteA4

### Buyer notice

This Pack is a digital product. After Stripe payment is confirmed, DAKE sends download instructions to the email address entered at purchase.

This is not automatic download. The standard delivery window is within the next business day after payment confirmation.

If the delivery email does not arrive, the buyer should contact DAKE with the purchased Pack name, purchase-time email address, purchase date/time, and Stripe payment information that can identify the payment. Card numbers, security codes, passwords, and other sensitive information must not be requested.

### Payment confirmation

Before sending download instructions, confirm all of the following in Stripe Dashboard.

- Payment succeeded.
- Payment is in live mode.
- Purchased product is `DAKE 書類整理パック`.
- Paid amount is 1480 JPY.
- Currency is JPY.
- Buyer email address is available.
- Payment is not refunded or canceled.

### Manual delivery procedure

1. Confirm the payment in Stripe Dashboard.
2. Match the payment to this Pack and the expected price.
3. Confirm the current delivery file path, size, and SHA256.
4. Generate or prepare a private download instruction for this buyer only.
5. Send the instruction to the buyer email address using the official template.
6. Record the delivery status in a secure local delivery log outside Git.

### Resend and failure handling

If the buyer requests resend, verify the original payment, Pack, buyer email address, and previous delivery record before resending. Do not send public download URLs.

If email delivery fails, record `delivery_failed`, confirm the email address and payment information, and handle refund or individual support according to existing DAKE Store policy.

### Prohibited storage

Do not store buyer email addresses, payment IDs, private download URLs, Stripe Secret Keys, Webhook Secrets, card information, or real delivery logs in Git, generated JSON, Store static files, Markdown reports, or public folders.
