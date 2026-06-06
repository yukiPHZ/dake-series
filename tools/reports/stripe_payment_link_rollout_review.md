# Stripe Payment Link Rollout Review

Generated: 2026-06-06 12:44:03 +0900

## 目的

Phase 11 の `stripe_payment_link_candidates.csv` をレビューし、Stripe Payment Link作成対象を `create` / `manual` / `hold` / `exclude` に分類する。今回はStripe API、Stripe Secret Key、Payment Link / Product / Price作成は行わない。

## 現状

- Stripe未対応候補: 48
- API作成候補: 45
- 手動作成候補: 2
- 保留: 1
- 対象外: 0
- price_missing: 0
- tax_code_review_required: 2

## 入力ファイル

- `tools/reports/stripe_payment_link_candidates.csv`
- `tools/generated/store_products.generated.json`
- `tools/reports/stripe_payment_links_full_rollout_plan.md`
- `00_core/DAKE_STORE_GENERATED_SPEC.md`
- `00_core/DAKE_STORE_OPERATION_RULE.md`

## 出力ファイル

- `tools/reports/stripe_payment_link_rollout_review.csv`
- `tools/reports/stripe_payment_link_rollout_review.md`

## 集計

| key | count |
| --- | --- |
| Stripe未対応 | 48 |
| API作成候補 | 45 |
| 手動作成候補 | 2 |
| 保留 | 1 |
| 対象外 | 0 |
| creation_method.api_candidate | 45 |
| creation_method.manual_dashboard | 2 |
| creation_method.hold | 1 |
| price_ok | 48 |
| price_missing | 0 |
| price_review | 0 |

## API作成候補

| id | type | title | price | tax_code | memo |
| --- | --- | --- | --- | --- | --- |
| DAKE_App_Doko | app | アプリどこ | 300 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| DAKE_Backup | app | Dakeバックアップ | 500 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| DAKE_Git_Memo | app | DakeGitメモ | 500 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| DAKE_Image_PasteA4 | app | 貼る | 500 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| DAKE_Mail_Address_Format | app | Dakeメールアドレス整形 | 300 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| DAKE_Mail_Draft | app | Dakeメール下書き | 300 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| DAKE_Maji_Memo | app | マジでメモ | 300 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| DAKE_Sticky_Memo | app | 付箋メモ | 300 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| DAKE_Yesterday_Task_Memo | app | Dake昨日タスクメモ | 300 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_booth_assist | app | BOOTHアシスト | 300 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_column_memo | app | ずっとメモ | 300 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_document_cover | app | Dake書類送付状 | 500 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_fax_cover | app | DakeFAX送付状 | 500 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_folder_list | app | Dakeフォルダ一覧 | 300 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_image_heictojpg | app | HEIC→JPG変換 | 500 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_image_iphonetopc | app | Dake画像iPhoneToPC | 500 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_image_receiver | app | DakeImage_Receiver | 500 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_image_topdf | app | DakeImageToPDF | 500 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_launcher | app | Dakeランチャー | 300 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_mail_allstaff | app | Dake全社員メール起動 | 300 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_mail_kikuta | app | Dake菊田メール | 300 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_mail_list | app | Dakeメールリスト | 300 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_mansion_schedule | app | マンション工程表 | 500 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_pdf_checkstamp | app | Dake確認印 | 500 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_pdf_crop | app | DakePDFトリミング | 500 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_pdf_lookhere | app | DakePDFここ見て | 500 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_pdf_marker | app | DakePDFマーカー | 500 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_pdf_merge_mini | app | DakePDF結合mini | 500 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_pdf_rename | app | DakePDFファイル名整理 | 500 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_pdf_reorder | app | DakePDFページ並べ替え | 500 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_pdf_splitone | app | DakePDF分割One | 500 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_pdf_splitselect | app | DakePDF分割Select | 500 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_pdf_toimages | app | DakePDFto画像 | 500 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_pdf_viewer | app | DakePDF見る | 500 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_price_apportionment | app | Dake価格按分 | 300 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_price_fixedtax | app | Dake固都税計算 | 300 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_reform_progress | app | リフォーム進捗管理 | 500 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_screen_webp | app | DakeScreen_WebP | 300 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_screenshot_print | app | Dakeスクショ印刷 | 500 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_two_person_memo | app | Dake二人メモ | 300 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_work_calendar | app | Dake工程カレンダー | 500 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_year_age | app | Dake築年数 | 300 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| dake_year_notice | app | Dake今年の注意点 | 300 JPY | txcd_10202003 | single DAKE app; API candidate with price, BOOTH URL, and source_original present; source memo: ready for Stripe Payment Link planning |
| game_alien_road | app | DakeAlien Road | 300 JPY | txcd_10201000 | single DAKE game app; API candidate, but tax_code review is required; source memo: ready for Stripe Payment Link planning |
| game_diver_catch | app | Dake潜って捕る | 300 JPY | txcd_10201000 | single DAKE game app; API candidate, but tax_code review is required; source memo: ready for Stripe Payment Link planning |

## 手動作成候補

| id | type | title | reason |
| --- | --- | --- | --- |
| DAKE_Pack_Document | pack | DAKE 書類整理パック | Pack item; create manually first because github_release_url is empty and pack_ready, BOOTH flow, and post-purchase guidance need individual review; source memo: github_release_url missing; acceptable for packs if delivery policy is separate |
| DAKE_Pack_Memo | pack | DAKE メモと記録パック | Pack item; create manually first because github_release_url is empty and pack_ready, BOOTH flow, and post-purchase guidance need individual review; source memo: github_release_url missing; acceptable for packs if delivery policy is separate |

## 保留候補

| id | type | title | reason |
| --- | --- | --- | --- |
| video_shorts_cut | app | Dakeショート切り出し | payment_status=preparing; BOOTH URL and sales flow are not confirmed; source memo: preparing item; do not create Stripe link yet |

## 対象外候補

| id | type | title | reason |
| --- | --- | --- | --- |
| - | - | - | - |

## tax_codeレビュー対象

| id | type | title | candidate | reason |
| --- | --- | --- | --- | --- |
| game_alien_road | app | DakeAlien Road | txcd_10201000 | single DAKE game app; API candidate, but tax_code review is required; source memo: ready for Stripe Payment Link planning |
| game_diver_catch | app | Dake潜って捕る | txcd_10201000 | single DAKE game app; API candidate, but tax_code review is required; source memo: ready for Stripe Payment Link planning |

## Pack商品の扱い

`DAKE_Pack_Document` と `DAKE_Pack_Memo` は手動作成候補にする。理由は、`github_release_url` がなく、Pack ZIP、`pack_ready/`、BOOTH導線、購入後案内を個別確認した方が安全なため。最初のPack 2件はStripe Dashboardで手動作成し、動線が固まった後にAPI化を検討する。

## metadata方針

維持するmetadataキー:

- `dake_item_id`
- `dake_type`
- `source_repo`
- `source_original`
- `store_url`
- `booth_url`
- `github_release_url`

空欄の値は空欄またはnull相当でよい。ただしSecret、個人情報、購入者情報、内部トークンはmetadataに入れない。

## API作成前の安全ルール

- Stripe Secret Keyは環境変数のみで扱う。
- Secretをpublic JS、generated JSON、GitHub repo、Store静的ファイル、`ORIGINAL.md` に入れない。
- Store側にカード情報、購入者DB、Stripe Secretを置かない。
- 最初は必ずStripe test modeで実行する。
- dry-runで作成予定Product / Price / Payment Linkを出し、人間レビュー後に本実行する。
- `metadata.dake_item_id` で既存Product重複を避ける。
- 作成後はPayment Link URLを対象商品の `ORIGINAL.md` に戻し、generated JSONを再生成する。

## 次Phase提案

1. API作成候補45件から先行10件を選ぶ。
2. 先行10件についてStripe test mode + dry-run用スクリプトを設計する。
3. Pack 2件はStripe Dashboardで手動作成し、購入後案内と配布導線を確認する。
4. `video_shorts_cut` はBOOTH URLと販売導線が確定するまで保留する。
5. Payment Link作成後、URLだけを各 `ORIGINAL.md` へ戻し、Store生成・同期は別Phaseで行う。
