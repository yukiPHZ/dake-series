# Stripe Payment Links Full Rollout Plan

Generated: 2026-06-06 09:03:03 +0900

## 目的

DAKE Store の全商品を Stripe Payment Link へ展開する前に、`ORIGINAL.md` 由来の `store_products.generated.json` を基準として、Stripe登録候補、登録用フィールド、metadata、安全ルール、手動/API作成の境界を整理する。

このPhaseでは Stripe API 実行、Payment Link作成、Product作成、Price作成、`ORIGINAL.md` 更新、Store同期は行わない。

## 現状

| metric | count |
| --- | --- |
| items | 53 |
| available | 53 |
| type.app | 50 |
| type.pack | 2 |
| type.shimarisu_pack | 1 |
| stripe_ready | 5 |
| booth_only | 47 |
| preparing | 1 |
| available and not stripe_ready | 48 |

参照元:

- DAKE_series: `tools/generated/store_products.generated.json`
- Store site mirror: `C:/Users/yukiz/devlop/dake-store-site/public/assets/data/store_products.generated.json`
- SHIMARISU source: `C:/Users/yukiz/devlop/SHIMARISU/ORIGINAL.md`
- Stripe Payment Links API: https://docs.stripe.com/api/payment_links/payment_links
- Stripe metadata: https://docs.stripe.com/metadata
- Stripe product tax codes: https://docs.stripe.com/tax/tax-codes

## Stripe対応済み商品

| id | type | title | price | stripe_payment_link | source_original |
| --- | --- | --- | --- | --- | --- |
| dake_image_resize | app | Dake画像リサイズ | 500 JPY | https://buy.stripe.com/5kQ5kD7fP8gf7Owf690gw03 | 01_apps/DAKE_Image_Resize/ORIGINAL.md |
| dake_pdf_compress | app | DakePDF圧縮 | 500 JPY | https://buy.stripe.com/aFa6oHeIh3ZZ6Kse250gw02 | 01_apps/DAKE_PDF_Compress/ORIGINAL.md |
| dake_pdf_merge | app | DakePDF結合 | 500 JPY | https://buy.stripe.com/9B6fZh0Rr2VV6Ksf690gw01 | 01_apps/DAKE_PDF_Merge/ORIGINAL.md |
| time_advanced_timer | app | Dakeアドバンスドタイマー | 300 JPY | https://buy.stripe.com/5kQdR9eIh8gfgl25vz0gw00 | 01_apps/DAKE_Time_AdvancedTimer/ORIGINAL.md |
| SHIMARISU_Pack | shimarisu_pack | しまりすくん 実務判断Pack | 3,000 JPY | https://buy.stripe.com/4gM9AT57H7cbc4M3nr0gw04 | C:/Users/yukiz/devlop/SHIMARISU/ORIGINAL.md |

## Stripe未対応商品一覧

対象条件: `status == available` かつ `payment_status != stripe_ready`。`preparing` は別枠で扱う。

| id | type | title | price | payment_status | classification |
| --- | --- | --- | --- | --- | --- |
| DAKE_App_Doko | app | アプリどこ | 300 JPY | booth_only | stripe_candidate |
| DAKE_Backup | app | Dakeバックアップ | 500 JPY | booth_only | stripe_candidate |
| DAKE_Git_Memo | app | DakeGitメモ | 500 JPY | booth_only | stripe_candidate |
| DAKE_Image_PasteA4 | app | 貼る | 500 JPY | booth_only | stripe_candidate |
| DAKE_Mail_Address_Format | app | Dakeメールアドレス整形 | 300 JPY | booth_only | stripe_candidate |
| DAKE_Mail_Draft | app | Dakeメール下書き | 300 JPY | booth_only | stripe_candidate |
| DAKE_Maji_Memo | app | マジでメモ | 300 JPY | booth_only | stripe_candidate |
| DAKE_Sticky_Memo | app | 付箋メモ | 300 JPY | booth_only | stripe_candidate |
| DAKE_Yesterday_Task_Memo | app | Dake昨日タスクメモ | 300 JPY | booth_only | stripe_candidate |
| dake_booth_assist | app | BOOTHアシスト | 300 JPY | booth_only | stripe_candidate |
| dake_column_memo | app | ずっとメモ | 300 JPY | booth_only | stripe_candidate |
| dake_document_cover | app | Dake書類送付状 | 500 JPY | booth_only | stripe_candidate |
| dake_fax_cover | app | DakeFAX送付状 | 500 JPY | booth_only | stripe_candidate |
| dake_folder_list | app | Dakeフォルダ一覧 | 300 JPY | booth_only | stripe_candidate |
| dake_image_heictojpg | app | HEIC→JPG変換 | 500 JPY | booth_only | stripe_candidate |
| dake_image_iphonetopc | app | Dake画像iPhoneToPC | 500 JPY | booth_only | stripe_candidate |
| dake_image_receiver | app | DakeImage_Receiver | 500 JPY | booth_only | stripe_candidate |
| dake_image_topdf | app | DakeImageToPDF | 500 JPY | booth_only | stripe_candidate |
| dake_launcher | app | Dakeランチャー | 300 JPY | booth_only | stripe_candidate |
| dake_mail_allstaff | app | Dake全社員メール起動 | 300 JPY | booth_only | stripe_candidate |
| dake_mail_kikuta | app | Dake菊田メール | 300 JPY | booth_only | stripe_candidate |
| dake_mail_list | app | Dakeメールリスト | 300 JPY | booth_only | stripe_candidate |
| dake_mansion_schedule | app | マンション工程表 | 500 JPY | booth_only | stripe_candidate |
| dake_pdf_checkstamp | app | Dake確認印 | 500 JPY | booth_only | stripe_candidate |
| dake_pdf_crop | app | DakePDFトリミング | 500 JPY | booth_only | stripe_candidate |
| dake_pdf_lookhere | app | DakePDFここ見て | 500 JPY | booth_only | stripe_candidate |
| dake_pdf_marker | app | DakePDFマーカー | 500 JPY | booth_only | stripe_candidate |
| dake_pdf_merge_mini | app | DakePDF結合mini | 500 JPY | booth_only | stripe_candidate |
| dake_pdf_rename | app | DakePDFファイル名整理 | 500 JPY | booth_only | stripe_candidate |
| dake_pdf_reorder | app | DakePDFページ並べ替え | 500 JPY | booth_only | stripe_candidate |
| dake_pdf_splitone | app | DakePDF分割One | 500 JPY | booth_only | stripe_candidate |
| dake_pdf_splitselect | app | DakePDF分割Select | 500 JPY | booth_only | stripe_candidate |
| dake_pdf_toimages | app | DakePDFto画像 | 500 JPY | booth_only | stripe_candidate |
| dake_pdf_viewer | app | DakePDF見る | 500 JPY | booth_only | stripe_candidate |
| dake_price_apportionment | app | Dake価格按分 | 300 JPY | booth_only | stripe_candidate |
| dake_price_fixedtax | app | Dake固都税計算 | 300 JPY | booth_only | stripe_candidate |
| dake_reform_progress | app | リフォーム進捗管理 | 500 JPY | booth_only | stripe_candidate |
| dake_screen_webp | app | DakeScreen_WebP | 300 JPY | booth_only | stripe_candidate |
| dake_screenshot_print | app | Dakeスクショ印刷 | 500 JPY | booth_only | stripe_candidate |
| dake_two_person_memo | app | Dake二人メモ | 300 JPY | booth_only | stripe_candidate |
| dake_work_calendar | app | Dake工程カレンダー | 500 JPY | booth_only | stripe_candidate |
| dake_year_age | app | Dake築年数 | 300 JPY | booth_only | stripe_candidate |
| dake_year_notice | app | Dake今年の注意点 | 300 JPY | booth_only | stripe_candidate |
| game_alien_road | app | DakeAlien Road | 300 JPY | booth_only | stripe_candidate |
| game_diver_catch | app | Dake潜って捕る | 300 JPY | booth_only | stripe_candidate |
| video_shorts_cut | app | Dakeショート切り出し | 500 JPY | preparing | preparing |
| DAKE_Pack_Document | pack | DAKE 書類整理パック | 1,480 JPY | booth_only | stripe_candidate |
| DAKE_Pack_Memo | pack | DAKE メモと記録パック | 980 JPY | booth_only | stripe_candidate |

## Stripe追加候補

現時点では、`booth_only` の available 商品47件をStripe追加候補とする。BOOTH URLがあり、価格・通貨・source_original が揃っているため、少なくとも登録計画の対象にできる。

| id | type | title | price | booth | source_original | memo |
| --- | --- | --- | --- | --- | --- | --- |
| DAKE_App_Doko | app | アプリどこ | 300 JPY | yes | 01_apps/DAKE_App_Doko/ORIGINAL.md | ready for Stripe Payment Link planning |
| DAKE_Backup | app | Dakeバックアップ | 500 JPY | yes | 01_apps/DAKE_Backup/ORIGINAL.md | ready for Stripe Payment Link planning |
| DAKE_Git_Memo | app | DakeGitメモ | 500 JPY | yes | 01_apps/DAKE_Git_Memo/ORIGINAL.md | ready for Stripe Payment Link planning |
| DAKE_Image_PasteA4 | app | 貼る | 500 JPY | yes | 01_apps/DAKE_Image_PasteA4/ORIGINAL.md | ready for Stripe Payment Link planning |
| DAKE_Mail_Address_Format | app | Dakeメールアドレス整形 | 300 JPY | yes | 01_apps/DAKE_Mail_Address_Format/ORIGINAL.md | ready for Stripe Payment Link planning |
| DAKE_Mail_Draft | app | Dakeメール下書き | 300 JPY | yes | 01_apps/DAKE_Mail_Draft/ORIGINAL.md | ready for Stripe Payment Link planning |
| DAKE_Maji_Memo | app | マジでメモ | 300 JPY | yes | 01_apps/DAKE_Maji_Memo/ORIGINAL.md | ready for Stripe Payment Link planning |
| DAKE_Sticky_Memo | app | 付箋メモ | 300 JPY | yes | 01_apps/DAKE_Sticky_Memo/ORIGINAL.md | ready for Stripe Payment Link planning |
| DAKE_Yesterday_Task_Memo | app | Dake昨日タスクメモ | 300 JPY | yes | 01_apps/DAKE_Yesterday_Task_Memo/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_booth_assist | app | BOOTHアシスト | 300 JPY | yes | 01_apps/DAKE_BOOTH_Assist/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_column_memo | app | ずっとメモ | 300 JPY | yes | 01_apps/DAKE_Column_Memo/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_document_cover | app | Dake書類送付状 | 500 JPY | yes | 01_apps/DAKE_Document_Cover/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_fax_cover | app | DakeFAX送付状 | 500 JPY | yes | 01_apps/DAKE_FAX_Cover/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_folder_list | app | Dakeフォルダ一覧 | 300 JPY | yes | 01_apps/DAKE_Folder_List/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_image_heictojpg | app | HEIC→JPG変換 | 500 JPY | yes | 01_apps/DAKE_Image_HEICtoJPG/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_image_iphonetopc | app | Dake画像iPhoneToPC | 500 JPY | yes | 01_apps/DAKE_Image_iPhoneToPC/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_image_receiver | app | DakeImage_Receiver | 500 JPY | yes | 01_apps/DAKE_Image_Receiver/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_image_topdf | app | DakeImageToPDF | 500 JPY | yes | 01_apps/DAKE_Image_ToPDF/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_launcher | app | Dakeランチャー | 300 JPY | yes | 01_apps/DAKE_Launcher/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_mail_allstaff | app | Dake全社員メール起動 | 300 JPY | yes | 01_apps/DAKE_Mail_AllStaff/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_mail_kikuta | app | Dake菊田メール | 300 JPY | yes | 01_apps/DAKE_Mail_Kikuta/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_mail_list | app | Dakeメールリスト | 300 JPY | yes | 01_apps/DAKE_Mail_List/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_mansion_schedule | app | マンション工程表 | 500 JPY | yes | 01_apps/DAKE_Mansion_Schedule/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_pdf_checkstamp | app | Dake確認印 | 500 JPY | yes | 01_apps/DAKE_PDF_CheckStamp/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_pdf_crop | app | DakePDFトリミング | 500 JPY | yes | 01_apps/DAKE_PDF_Crop/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_pdf_lookhere | app | DakePDFここ見て | 500 JPY | yes | 01_apps/DAKE_PDF_LookHere/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_pdf_marker | app | DakePDFマーカー | 500 JPY | yes | 01_apps/DAKE_PDF_Marker/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_pdf_merge_mini | app | DakePDF結合mini | 500 JPY | yes | 01_apps/DAKE_PDF_Merge_Mini/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_pdf_rename | app | DakePDFファイル名整理 | 500 JPY | yes | 01_apps/DAKE_PDF_Rename/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_pdf_reorder | app | DakePDFページ並べ替え | 500 JPY | yes | 01_apps/DAKE_PDF_Reorder/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_pdf_splitone | app | DakePDF分割One | 500 JPY | yes | 01_apps/DAKE_PDF_SplitOne/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_pdf_splitselect | app | DakePDF分割Select | 500 JPY | yes | 01_apps/DAKE_PDF_SplitSelect/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_pdf_toimages | app | DakePDFto画像 | 500 JPY | yes | 01_apps/DAKE_PDF_ToImages/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_pdf_viewer | app | DakePDF見る | 500 JPY | yes | 01_apps/DAKE_PDF_Viewer/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_price_apportionment | app | Dake価格按分 | 300 JPY | yes | 01_apps/DAKE_Price_Apportionment/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_price_fixedtax | app | Dake固都税計算 | 300 JPY | yes | 01_apps/DAKE_Price_FixedTax/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_reform_progress | app | リフォーム進捗管理 | 500 JPY | yes | 01_apps/DAKE_Reform_Progress/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_screen_webp | app | DakeScreen_WebP | 300 JPY | yes | 01_apps/DAKE_Screen_WebP/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_screenshot_print | app | Dakeスクショ印刷 | 500 JPY | yes | 01_apps/DAKE_Screenshot_Print/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_two_person_memo | app | Dake二人メモ | 300 JPY | yes | 01_apps/DAKE_TwoPerson_Memo/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_work_calendar | app | Dake工程カレンダー | 500 JPY | yes | 01_apps/DAKE_Work_Calendar/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_year_age | app | Dake築年数 | 300 JPY | yes | 01_apps/DAKE_Year_Age/ORIGINAL.md | ready for Stripe Payment Link planning |
| dake_year_notice | app | Dake今年の注意点 | 300 JPY | yes | 01_apps/DAKE_Year_Notice/ORIGINAL.md | ready for Stripe Payment Link planning |
| game_alien_road | app | DakeAlien Road | 300 JPY | yes | 01_apps/DAKE_Game_Alien_Road/ORIGINAL.md | ready for Stripe Payment Link planning |
| game_diver_catch | app | Dake潜って捕る | 300 JPY | yes | 01_apps/DAKE_Game_Diver_Catch/ORIGINAL.md | ready for Stripe Payment Link planning |
| DAKE_Pack_Document | pack | DAKE 書類整理パック | 1,480 JPY | yes | 04_packs/DAKE_Pack_Document/ORIGINAL.md | github_release_url missing; acceptable for packs if delivery policy is separate |
| DAKE_Pack_Memo | pack | DAKE メモと記録パック | 980 JPY | yes | 04_packs/DAKE_Pack_Memo/ORIGINAL.md | github_release_url missing; acceptable for packs if delivery policy is separate |

## BOOTHのみでよい候補

現時点では、積極的にBOOTH限定のまま残す候補はない。BOOTH限定維持にする場合は、価格、サポート負荷、配布導線、税務上の扱いを商品ごとに明示してから除外する。

| id | type | title | reason |
| --- | --- | --- | --- |
| - | - | - | - |

## 準備中

| id | type | title | reason |
| --- | --- | --- | --- |
| video_shorts_cut | app | Dakeショート切り出し | preparing item; do not create Stripe link yet |

## 確認必要

CSV上の `needs_review` は0件。Pack 2件は `github_release_url` が空だが、BOOTH URLと価格はあり、Pack商品の配布導線を別途定義する前提でStripe候補に含める。

| id | title | booth_url | github_release_url | memo |
| --- | --- | --- | --- | --- |
| DAKE_Pack_Document | DAKE 書類整理パック | https://peakheadz.booth.pm/items/8448353 | - | github_release_url missing; acceptable for packs if delivery policy is separate |
| DAKE_Pack_Memo | DAKE メモと記録パック | https://peakheadz.booth.pm/items/8449208 | - | github_release_url missing; acceptable for packs if delivery policy is separate |

## Stripe登録用フィールド案

詳細CSV: `tools/reports/stripe_payment_link_candidates.csv`

| store_id | type | stripe_product_name | stripe_price_amount | currency | tax_code_candidate | Payment Link target | metadata keys | memo |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| DAKE_App_Doko | app | アプリどこ | 300 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| DAKE_Backup | app | Dakeバックアップ | 500 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| DAKE_Git_Memo | app | DakeGitメモ | 500 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| DAKE_Image_PasteA4 | app | 貼る | 500 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| DAKE_Mail_Address_Format | app | Dakeメールアドレス整形 | 300 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| DAKE_Mail_Draft | app | Dakeメール下書き | 300 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| DAKE_Maji_Memo | app | マジでメモ | 300 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| DAKE_Sticky_Memo | app | 付箋メモ | 300 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| DAKE_Yesterday_Task_Memo | app | Dake昨日タスクメモ | 300 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_booth_assist | app | BOOTHアシスト | 300 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_column_memo | app | ずっとメモ | 300 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_document_cover | app | Dake書類送付状 | 500 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_fax_cover | app | DakeFAX送付状 | 500 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_folder_list | app | Dakeフォルダ一覧 | 300 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_image_heictojpg | app | HEIC→JPG変換 | 500 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_image_iphonetopc | app | Dake画像iPhoneToPC | 500 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_image_receiver | app | DakeImage_Receiver | 500 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_image_topdf | app | DakeImageToPDF | 500 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_launcher | app | Dakeランチャー | 300 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_mail_allstaff | app | Dake全社員メール起動 | 300 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_mail_kikuta | app | Dake菊田メール | 300 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_mail_list | app | Dakeメールリスト | 300 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_mansion_schedule | app | マンション工程表 | 500 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_pdf_checkstamp | app | Dake確認印 | 500 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_pdf_crop | app | DakePDFトリミング | 500 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_pdf_lookhere | app | DakePDFここ見て | 500 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_pdf_marker | app | DakePDFマーカー | 500 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_pdf_merge_mini | app | DakePDF結合mini | 500 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_pdf_rename | app | DakePDFファイル名整理 | 500 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_pdf_reorder | app | DakePDFページ並べ替え | 500 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_pdf_splitone | app | DakePDF分割One | 500 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_pdf_splitselect | app | DakePDF分割Select | 500 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_pdf_toimages | app | DakePDFto画像 | 500 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_pdf_viewer | app | DakePDF見る | 500 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_price_apportionment | app | Dake価格按分 | 300 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_price_fixedtax | app | Dake固都税計算 | 300 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_reform_progress | app | リフォーム進捗管理 | 500 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_screen_webp | app | DakeScreen_WebP | 300 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_screenshot_print | app | Dakeスクショ印刷 | 500 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_two_person_memo | app | Dake二人メモ | 300 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_work_calendar | app | Dake工程カレンダー | 500 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_year_age | app | Dake築年数 | 300 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| dake_year_notice | app | Dake今年の注意点 | 300 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| game_alien_road | app | DakeAlien Road | 300 | JPY | txcd_10201000 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| game_diver_catch | app | Dake潜って捕る | 300 | JPY | txcd_10201000 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | ready for Stripe Payment Link planning |
| DAKE_Pack_Document | pack | DAKE 書類整理パック | 1480 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | github_release_url missing; acceptable for packs if delivery policy is separate |
| DAKE_Pack_Memo | pack | DAKE メモと記録パック | 980 | JPY | txcd_10202003 | yes | dake_item_id, dake_type, source_repo, source_original, store_url, booth_url, github_release_url | github_release_url missing; acceptable for packs if delivery policy is separate |

### 共通フィールド方針

- `stripe_product_name`: Store表示名と同じ `title` を初期値にする。
- `stripe_price_amount`: JPYはゼロ小数通貨なので、`price` の整数値をそのまま `unit_amount` 候補にする。
- `description`: `store_products.generated.json.description` をProduct説明案として使う。長すぎる場合のみDashboard入力時に短縮する。
- `currency`: 全件 `JPY`。
- `tax_code_candidate`: DAKE実務アプリ/Packは `txcd_10202003` を候補にする。ゲーム系は `txcd_10201000` を候補にする。
- tax code は候補であり、Stripe Tax有効化前または本番登録前に税務・販売地域・個人/事業用途の扱いを確認する。

## metadata案

Product と Payment Link の両方に、最低限以下を入れる案とする。Payment Linkのmetadataは、そのリンクから作成されるCheckout Sessionへコピーされるため、後続の照合に使える。

| metadata key | value source | purpose |
| --- | --- | --- |
| dake_item_id | item.id | Store itemとの照合 |
| dake_type | item.type | app / pack / shimarisu_pack の識別 |
| source_repo | item.source_repo | 正本repoの識別 |
| source_original | item.source_original | ORIGINAL.mdへの戻し先 |
| store_url | generated Store URL | 購入導線との照合 |
| booth_url | item.booth_url | BOOTH併用導線との照合 |
| github_release_url | item.github_release_url | 配布・Release導線との照合 |

metadataにはSecret、個人情報、購入者情報、内部トークンを入れない。

## API自動化する場合の安全ルール

- Stripe Secret Keyは環境変数でのみ扱う。
- Secret Keyをpublic JS、generated JSON、GitHub repo、Store静的ファイル、`ORIGINAL.md` に保存しない。
- Store側にカード情報、購入者DB、Stripe Secretを置かない。
- 最初は必ずStripe test modeで実行する。
- dry-runモードを必須にし、作成予定Product/Price/Payment Link一覧を人間が確認してから本実行する。
- 既存Product重複回避は `metadata.dake_item_id` と必要に応じてProduct検索で行う。
- Product、Price、Payment Linkを作る場合も、作成後に保存するのはPayment Link URLと必要最小限のStripe IDだけにする。
- Payment Link URLは対象商品の `ORIGINAL.md` へ戻し、generated JSONは再生成する。
- BOOTH公開、Store同期、Cloudflare deploy、Payment Link本番作成を同一Phaseに混ぜない。

## 手動作成とAPI作成の境界

- 少数追加: Stripe Dashboardで手動作成。作成後、Payment Link URLだけを `ORIGINAL.md` に戻す。
- 10件以上: API / CLI / ローカル補助スクリプトを検討。ただしdry-runとtest modeを必須にする。
- 全商品一括: 必ず一覧レビュー、test mode作成、差分確認、本番作成の順に分ける。
- Pack商品: 配布導線と購入後案内が単品と異なるため、最初の2件は手動作成で挙動確認してから一括化する。

## 次Phase提案

1. `stripe_payment_link_candidates.csv` をレビューし、47件のStripe追加候補から先行10件を選ぶ。
2. Pack 2件の購入後案内と配布導線を決める。
3. Stripe Dashboardで数件を手動作成し、`ORIGINAL.md` へPayment Link URLを戻す。
4. 10件以上を進める場合、test mode + dry-run専用の作成スクリプトを別Phaseで設計する。
5. `generate_store_products.py` と `sync_store_to_site.py` でStore表示へ反映する。
