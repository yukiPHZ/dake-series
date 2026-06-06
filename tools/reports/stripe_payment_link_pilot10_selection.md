# Stripe Payment Link Pilot 10 Selection

Generated: 2026-06-06 12:50:40 +0900

## 目的

Phase 12で分類したStripe Payment Link API作成候補45件から、次Phaseのtest mode / dry-run検証に使う先行10件を選定する。今回はStripe API、Stripe Secret Key、Product / Price / Payment Link作成は行わない。

## 入力ファイル

- `tools/reports/stripe_payment_link_rollout_review.csv`
- `tools/reports/stripe_payment_link_candidates.csv`
- `tools/reports/stripe_payment_links_full_rollout_plan.md`
- `tools/generated/store_products.generated.json`

## 選定方針

- `review_result=create`
- `creation_method=api_candidate`
- `price_check=price_ok`
- `booth_url` あり
- `metadata_ready=yes`
- Pack、preparing、ゲーム系、tax_codeレビュー必須の商品は除外
- PDF、画像、メモ、メール、作業補助、ファイル整理、年数/計算系に分散
- 計算系は先行10件では1件に留める

## 除外したもの

| id | type | title | reason |
| --- | --- | --- | --- |
| DAKE_Pack_Document | pack | DAKE 書類整理パック | Pack商品は手動作成候補。github_release_urlがなく、Pack ZIP/pack_ready/BOOTH導線/購入後案内を個別確認する。 |
| DAKE_Pack_Memo | pack | DAKE メモと記録パック | Pack商品は手動作成候補。github_release_urlがなく、Pack ZIP/pack_ready/BOOTH導線/購入後案内を個別確認する。 |
| video_shorts_cut | app | Dakeショート切り出し | payment_status=preparing。BOOTH URLなし、販売導線未確定のため保留。 |
| game_alien_road | app | DakeAlien Road | ゲーム系はtax_code要確認のため後続確認へ回す。 |
| game_diver_catch | app | Dake潜って捕る | ゲーム系はtax_code要確認のため後続確認へ回す。 |

## 先行10件

| id | type | title | price | category | reason |
| --- | --- | --- | --- | --- | --- |
| dake_pdf_viewer | app | DakePDF見る | 500 JPY | PDF | PDF閲覧で内容が分かりやすく、既存Stripe対応PDF系と近い通常アプリ。 |
| dake_pdf_reorder | app | DakePDFページ並べ替え | 500 JPY | PDF | PDFページ並び替えで用途が明確。PDF系の別操作パターンを確認できる。 |
| dake_pdf_splitone | app | DakePDF分割One | 500 JPY | PDF | PDF分割系の代表として選定。既存PDF結合/圧縮とは異なる操作カテゴリ。 |
| dake_image_heictojpg | app | HEIC→JPG変換 | 500 JPY | 画像 | HEICからJPGへの変換で商品内容が直感的。画像変換系の代表。 |
| dake_image_topdf | app | DakeImageToPDF | 500 JPY | 画像/PDF | 画像PDF化でPDF/画像の横断カテゴリを確認できる。 |
| DAKE_Sticky_Memo | app | 付箋メモ | 300 JPY | メモ / 付箋 | 付箋メモは軽量で用途が明確。メモ系の代表として適している。 |
| DAKE_Mail_Draft | app | Dakeメール下書き | 300 JPY | メール | メール下書きは自動送信しない補助ツールで、説明・免責が比較的整理しやすい。 |
| DAKE_Backup | app | Dakeバックアップ | 500 JPY | バックアップ / 保全 | バックアップ補助は実務用途が明確。購入前注意書きは維持する。 |
| dake_folder_list | app | Dakeフォルダ一覧 | 300 JPY | ファイル / 一覧 | フォルダ一覧化で配布導線・商品説明が分かりやすい。 |
| dake_year_age | app | Dake築年数 | 300 JPY | 年数・日付 | 年数計算系を1件だけ入れる。高度な税務判断ではなく、比較的低リスクな計算枠。 |

## ジャンル分散

| genre | count |
| --- | --- |
| PDF系 | 3 |
| 画像系 | 2 |
| メモ系 | 1 |
| メール系 | 1 |
| 作業補助系 | 1 |
| ファイル整理系 | 1 |
| 年数/計算系 | 1 |

## tax_code確認

選定10件はすべて `tax_code_review_required=no` の通常候補。tax_code候補は全件 `txcd_10202003`。ゲーム系2件は今回除外し、後続のtax_code確認へ回す。

| id | type | title | candidate | reason |
| --- | --- | --- | --- | --- |
| game_alien_road | app | DakeAlien Road | txcd_10201000 | ゲーム系はtax_code要確認のため後続確認へ回す。 |
| game_diver_catch | app | Dake潜って捕る | txcd_10201000 | ゲーム系はtax_code要確認のため後続確認へ回す。 |

## metadata確認

選定10件はすべて `metadata_ready=yes`。維持するmetadataキーは `dake_item_id`, `dake_type`, `source_repo`, `source_original`, `store_url`, `booth_url`, `github_release_url`。

## 次Phaseで使う想定

- Stripe test mode用dry-runスクリプトの入力リストとして `tools/reports/stripe_payment_link_pilot10_selection.csv` を使う。
- dry-runではProduct / Price / Payment Link作成予定だけを表示し、Stripe APIは本実行しない。
- metadataで `dake_item_id` を紐付け、既存Product重複を避ける。
- 作成後に保存するのはPayment Link URLと必要最小限のStripe IDだけにする。

## 今回やらなかったこと

- Stripe API実行
- Stripe Secret Key使用
- Product作成
- Price作成
- Payment Link作成
- `ORIGINAL.md`更新
- generated JSON更新
- dake-store-site同期
- Store本番反映
- Pack商品のStripe作成
- ゲーム系商品のStripe作成

## 次Phase提案

1. `stripe_payment_link_pilot10_selection.csv` を入力に、test mode + dry-run専用スクリプトを設計する。
2. dry-run出力でProduct名、価格、currency、tax_code候補、metadataを確認する。
3. Stripe API本実行はさらに次のPhaseで、人間レビュー後に限定して行う。
4. Pack 2件はDashboard手動作成で先に導線を確認する。
5. ゲーム系2件はtax_code候補を確認してから別枠で進める。
