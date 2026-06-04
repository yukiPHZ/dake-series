# ORIGINAL Phase4-5J App Batch Rollout

## 目的

Phase 4-5Iまでで、単品アプリ、DAKE Pack、SHIMARISU Packの正本構造が整理された。

今回はA1単品アプリのうち、まだ `ORIGINAL.md` が未作成だったものから15件を選び、単品アプリ用テンプレートで横展開した。

## 対象範囲

- 対象: `01_apps/` 配下のA1単品アプリ
- 対象条件: available / BOOTH登録済みまたはbooth_readyあり / Releaseあり / ORIGINAL.md未作成
- 対象外: `04_packs/`、SHIMARISU Pack、shimarisu-pack-release、prototype、frozen、draft、internal、既にORIGINAL.mdがあるアプリ

## 選定したアプリ

| folder | title | reason | source files |
|---|---|---|---|
| DAKE_App_Doko | アプリどこ | 補助ツール系。A1単品アプリ、ORIGINAL未作成、BOOTH/Release情報あり。 | README.md, release_body.md, booth_product.txt, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/screenshot.webp, assets/booth_thumbnail.jpg |
| DAKE_BOOTH_Assist | BOOTHアシスト | 補助ツール系。A1単品アプリ、ORIGINAL未作成、BOOTH/Release情報あり。 | README.md, release_body.md, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/screenshot.webp, assets/booth_thumbnail.jpg |
| DAKE_Column_Memo | ずっとメモ | メモ系。A1単品アプリ、ORIGINAL未作成、BOOTH/Release情報あり。 | README.md, release_body.md, booth_product.txt, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/screenshot.webp, assets/booth_thumbnail.jpg |
| DAKE_Document_Cover | Dake書類送付状 | 書類系。A1単品アプリ、ORIGINAL未作成、BOOTH/Release情報あり。 | README.md, release_body.md, booth_product.txt, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/screenshot.webp, assets/booth_thumbnail.jpg |
| DAKE_FAX_Cover | DakeFAX送付状 | 書類系。A1単品アプリ、ORIGINAL未作成、BOOTH/Release情報あり。 | README.md, release_body.md, booth_product.txt, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/screenshot.webp, assets/booth_thumbnail.jpg |
| DAKE_Image_iPhoneToPC | Dake画像iPhoneToPC | 画像系。A1単品アプリ、ORIGINAL未作成、BOOTH/Release情報あり。 | README.md, release_body.md, booth_product.txt, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/screenshot.webp, assets/booth_thumbnail.jpg |
| DAKE_Image_PasteA4 | 貼る | 画像系。A1単品アプリ、ORIGINAL未作成、BOOTH/Release情報あり。 DAKE Pack構成要素でもあり、Pack側の暫定参照解消候補。 | README.md, release_body.md, booth_product.txt, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/screenshot.webp, assets/booth_thumbnail.jpg |
| DAKE_Image_Receiver | DakeImage_Receiver | 画像系。A1単品アプリ、ORIGINAL未作成、BOOTH/Release情報あり。 | README.md, release_body.md, booth_product.txt, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/screenshot.webp, assets/booth_thumbnail.jpg |
| DAKE_Launcher | Dakeランチャー | 補助ツール系。A1単品アプリ、ORIGINAL未作成、BOOTH/Release情報あり。 | README.md, release_body.md, booth_product.txt, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/screenshot.webp, assets/booth_thumbnail.jpg |
| DAKE_Mail_Address_Format | Dakeメールアドレス整形 | メール系。A1単品アプリ、ORIGINAL未作成、BOOTH/Release情報あり。 | README.md, release_body.md, booth_product.txt, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/screenshot.webp, assets/booth_thumbnail.jpg |
| DAKE_Mail_AllStaff | Dake全社員メール起動 | メール系。A1単品アプリ、ORIGINAL未作成、BOOTH/Release情報あり。 | README.md, release_body.md, booth_product.txt, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/screenshot.webp, assets/booth_thumbnail.jpg |
| DAKE_Mail_Kikuta | Dake菊田メール | メール系。A1単品アプリ、ORIGINAL未作成、BOOTH/Release情報あり。 | README.md, release_body.md, booth_product.txt, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/screenshot.webp, assets/booth_thumbnail.jpg |
| DAKE_PDF_CheckStamp | Dake確認印 | PDF系。A1単品アプリ、ORIGINAL未作成、BOOTH/Release情報あり。 | README.md, release_body.md, booth_product.txt, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/screenshot.webp, assets/booth_thumbnail.jpg |
| DAKE_PDF_Crop | DakePDFトリミング | PDF系。A1単品アプリ、ORIGINAL未作成、BOOTH/Release情報あり。 | README.md, release_body.md, booth_product.txt, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/screenshot.webp, assets/booth_thumbnail.jpg |
| DAKE_Yesterday_Task_Memo | Dake昨日タスクメモ | メモ系。A1単品アプリ、ORIGINAL未作成、BOOTH/Release情報あり。 DAKE Pack構成要素でもあり、Pack側の暫定参照解消候補。 | README.md, release_body.md, booth_product.txt, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/screenshot.webp, assets/booth_thumbnail.jpg |

## 作成した ORIGINAL.md

| folder | original_path | missing_info | memo |
|---|---|---|---|
| DAKE_App_Doko | `01_apps/DAKE_App_Doko/ORIGINAL.md` | Store URL / Storeダウンロード導線 / Storeサポート方針 / Store用画像 | 既存派生ビューは未変更。 |
| DAKE_BOOTH_Assist | `01_apps/DAKE_BOOTH_Assist/ORIGINAL.md` | Store URL / Storeダウンロード導線 / Storeサポート方針 / Store用画像 | 既存派生ビューは未変更。 |
| DAKE_Column_Memo | `01_apps/DAKE_Column_Memo/ORIGINAL.md` | Store URL / Storeダウンロード導線 / Storeサポート方針 / Store用画像 | 既存派生ビューは未変更。 |
| DAKE_Document_Cover | `01_apps/DAKE_Document_Cover/ORIGINAL.md` | Store URL / Storeダウンロード導線 / Storeサポート方針 / Store用画像 | 既存派生ビューは未変更。 |
| DAKE_FAX_Cover | `01_apps/DAKE_FAX_Cover/ORIGINAL.md` | Store URL / Storeダウンロード導線 / Storeサポート方針 / Store用画像 | 既存派生ビューは未変更。 |
| DAKE_Image_iPhoneToPC | `01_apps/DAKE_Image_iPhoneToPC/ORIGINAL.md` | Store URL / Storeダウンロード導線 / Storeサポート方針 / Store用画像 | 既存派生ビューは未変更。 |
| DAKE_Image_PasteA4 | `01_apps/DAKE_Image_PasteA4/ORIGINAL.md` | Store URL / Storeダウンロード導線 / Storeサポート方針 / Store用画像 | 既存派生ビューは未変更。 |
| DAKE_Image_Receiver | `01_apps/DAKE_Image_Receiver/ORIGINAL.md` | Store URL / Storeダウンロード導線 / Storeサポート方針 / Store用画像 | 既存派生ビューは未変更。 |
| DAKE_Launcher | `01_apps/DAKE_Launcher/ORIGINAL.md` | Store URL / Storeダウンロード導線 / Storeサポート方針 / Store用画像 | 既存派生ビューは未変更。 |
| DAKE_Mail_Address_Format | `01_apps/DAKE_Mail_Address_Format/ORIGINAL.md` | Store URL / Storeダウンロード導線 / Storeサポート方針 / Store用画像 | 既存派生ビューは未変更。 |
| DAKE_Mail_AllStaff | `01_apps/DAKE_Mail_AllStaff/ORIGINAL.md` | Store URL / Storeダウンロード導線 / Storeサポート方針 / Store用画像 | 既存派生ビューは未変更。 |
| DAKE_Mail_Kikuta | `01_apps/DAKE_Mail_Kikuta/ORIGINAL.md` | Store URL / Storeダウンロード導線 / Storeサポート方針 / Store用画像 | 既存派生ビューは未変更。 |
| DAKE_PDF_CheckStamp | `01_apps/DAKE_PDF_CheckStamp/ORIGINAL.md` | Store URL / Storeダウンロード導線 / Storeサポート方針 / Store用画像 | 既存派生ビューは未変更。 |
| DAKE_PDF_Crop | `01_apps/DAKE_PDF_Crop/ORIGINAL.md` | Store URL / Storeダウンロード導線 / Storeサポート方針 / Store用画像 | 既存派生ビューは未変更。 |
| DAKE_Yesterday_Task_Memo | `01_apps/DAKE_Yesterday_Task_Memo/ORIGINAL.md` | Store URL / Storeダウンロード導線 / Storeサポート方針 / Store用画像 | 既存派生ビューは未変更。 |

## 共通してうまくいった点

- README内DAKE_META、release_body.md、booth_product、booth_readyから、基本情報・Release情報・BOOTH情報を集約できた。
- Store関連は未確定として残し、推測で埋めなかった。
- Pack構成要素の `DAKE_Image_PasteA4` と `DAKE_Yesterday_Task_Memo` も、Pack情報を混ぜず単品アプリ正本として作成できた。

## アプリごとに詰まった点

- メール系は送信有無やOutlook連携の細部を、既存ファイルだけでは断定しない形にした。
- `DAKE_BOOTH_Assist` と `DAKE_Launcher` はFactory/導線補助色が強いため、Store掲載時の販売文は慎重に確認する必要がある。
- 対応形式や外部ツールは、実装確認なしに断定しないため、多くを未記載・要確認として残した。

## 未確定として残した情報

- Store URL
- Storeダウンロード導線
- Storeサポート方針
- Store用画像
- 一部アプリの対応形式・外部ツール・保存方針の詳細

## 次Batchへの注意

- PDF系の残件が多いため、次BatchではPDF操作ごとの違いをREADME/booth_productから丁寧に拾う。
- Store仕様が固まるまでは、Store関連は未確定のまま残す。
- 既存派生ビューは更新しない。

## 残りA1単品アプリ

- `DAKE_PDF_LookHere`
- `DAKE_PDF_Marker`
- `DAKE_PDF_Merge_Mini`
- `DAKE_PDF_Reorder`
- `DAKE_PDF_SplitOne`
- `DAKE_PDF_SplitSelect`
- `DAKE_PDF_ToImages`
- `DAKE_PDF_Viewer`
- `DAKE_Reform_Progress`
- `DAKE_Screen_WebP`
- `DAKE_Screenshot_Print`
- `DAKE_TwoPerson_Memo`
- `DAKE_Year_Age`
- `DAKE_Year_Notice`

残り件数: 14

## Store接続前の注意

- Storeは単品アプリORIGINAL、DAKE Pack ORIGINAL、SHIMARISU Pack ORIGINALの3系統を読める必要がある。
- Store専用の商品正本は作らない。
- Pack構成要素でも、単品アプリとしてStore掲載する場合は各アプリORIGINALを優先する。

## 次Phase提案

1. Phase 4-5J-2: 残りA1単品アプリ 14 件へ `ORIGINAL.md` を展開する。
2. Store用generatedデータ形式を定義する。
3. ORIGINAL由来のREADME / release_body / booth_product生成方針を検討する。
