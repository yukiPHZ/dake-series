# ORIGINAL Phase4-5E App Batch Rollout

## 目的

Phase 4-5Dで単品アプリとPack商品の正本構造を分けたうえで、A1単品アプリ優先対象から15件を選び、改善済み単品アプリ用テンプレートで `ORIGINAL.md` を横展開した。

今回はPack商品、SHIMARISU Pack、prototype / frozen / internal系は対象外とし、`01_apps` 配下の単品アプリだけを扱った。

## 対象範囲

- 対象: `01_apps/` 配下のA1単品アプリ
- 対象条件: available / BOOTH登録済みまたはbooth_readyあり / Releaseあり / ORIGINAL.md未作成
- 対象外: `04_packs/`、SHIMARISU Pack、shimarisu-pack-release、prototype、frozen、draft、internal、既にORIGINAL.mdがあるアプリ

## 選定したアプリ

| folder | title | reason | source files |
|---|---|---|---|
| DAKE_PDF_Merge | DakePDF結合 | PDF系。Pack構成要素でもあり、単品商品としても優先度が高い。 | README.md, README内DAKE_META, release_body.md, booth_product.txt, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/screenshot.webp, assets/booth_thumbnail.jpg |
| DAKE_PDF_Rename | DakePDFファイル名整理 | PDF系。ファイル名整理という非破壊・実務補助の観点を確認できる。 | README.md, README内DAKE_META, release_body.md, booth_product.txt, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/screenshot.webp, assets/booth_thumbnail.jpg |
| DAKE_Image_ToPDF | DakeImageToPDF | 画像系。Pack構成要素でもあり、画像からPDFへの変換方針を整理できる。 | README.md, README内DAKE_META, release_body.md, booth_product.txt, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/screenshot.webp, assets/booth_thumbnail.jpg |
| DAKE_Image_HEICtoJPG | HEIC→JPG変換 | 画像系。対応形式・非対応形式の欄が効く変換アプリ。 | README.md, README内DAKE_META, release_body.md, booth_product.txt, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/screenshot.webp, assets/booth_thumbnail.jpg |
| DAKE_Mail_Draft | Dakeメール下書き | メール系。Outlook下書き作成のみ、送信しない方針を正本化できる。 | README.md, README内DAKE_META, release_body.md, booth_product.txt, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/screenshot.webp, assets/booth_thumbnail.jpg |
| DAKE_Mail_List | Dakeメールリスト | メール系。.msgからCSV整理という形式変換と非送信方針を整理できる。 | README.md, README内DAKE_META, release_body.md, booth_product.txt, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/screenshot.webp, assets/booth_thumbnail.jpg |
| DAKE_Maji_Memo | マジでメモ | メモ系。軽量アプリの非ゴール整理を確認できる。 | README.md, README内DAKE_META, release_body.md, booth_product.txt, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/screenshot.webp, assets/booth_thumbnail.jpg |
| DAKE_Git_Memo | DakeGitメモ | メモ/開発記録系。Git作業メモとしてStore説明へ展開しやすい。 | README.md, README内DAKE_META, release_body.md, booth_product.txt, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/screenshot.webp, assets/booth_thumbnail.jpg |
| DAKE_Work_Calendar | Dake工程カレンダー | 作業補助系。工程・日付管理の説明整理を確認できる。 | README.md, README内DAKE_META, release_body.md, booth_product.txt, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/screenshot.webp, assets/booth_thumbnail.jpg |
| DAKE_Mansion_Schedule | マンション工程表 | 工程管理系。業務カレンダー系と別の現場向け日程管理を確認できる。 | README.md, README内DAKE_META, release_body.md, booth_product.txt, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/screenshot.webp, assets/booth_thumbnail.jpg |
| DAKE_Price_FixedTax | Dake固都税計算 | 計算系。価格・税額計算アプリの正本化を確認できる。 | README.md, README内DAKE_META, release_body.md, booth_product.txt, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/screenshot.webp, assets/booth_thumbnail.jpg |
| DAKE_Price_Apportionment | Dake価格按分 | 計算系。按分計算アプリとして入力・出力の整理を確認できる。 | README.md, README内DAKE_META, release_body.md, booth_product.txt, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/screenshot.webp, assets/booth_thumbnail.jpg |
| DAKE_Game_Alien_Road | DakeAlien Road | ゲーム系。単品ゲームアプリをPackと分けて扱う確認ができる。 | README.md, README内DAKE_META, release_body.md, booth_product.txt, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/screenshot.webp, assets/booth_thumbnail.jpg |
| DAKE_Game_Diver_Catch | Dake潜って捕る | ゲーム系。市場向けゲームアプリの正本化を確認できる。 | README.md, README内DAKE_META, release_body.md, booth_product.txt, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/screenshot.webp, assets/booth_thumbnail.jpg |
| DAKE_Video_Shorts_Cut | Dakeショート切り出し | 動画系。外部ツールや対応形式の整理が必要になりやすい。 | README.md, README内DAKE_META, release_body.md, booth_product.txt, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/screenshot.webp, assets/booth_thumbnail.jpg |

## 作成した ORIGINAL.md

| folder | original_path | missing_info | memo |
|---|---|---|---|
| DAKE_PDF_Merge | 01_apps/DAKE_PDF_Merge/ORIGINAL.md | Store URL / Storeダウンロード導線 / Storeサポート方針 / Store用画像 | PDF系。既存派生ビューは未変更。 |
| DAKE_PDF_Rename | 01_apps/DAKE_PDF_Rename/ORIGINAL.md | Store URL / Storeダウンロード導線 / Storeサポート方針 / Store用画像 | PDF系。既存派生ビューは未変更。 |
| DAKE_Image_ToPDF | 01_apps/DAKE_Image_ToPDF/ORIGINAL.md | Store URL / Storeダウンロード導線 / Storeサポート方針 / Store用画像 | 画像/PDF系。既存派生ビューは未変更。 |
| DAKE_Image_HEICtoJPG | 01_apps/DAKE_Image_HEICtoJPG/ORIGINAL.md | Store URL / Storeダウンロード導線 / Storeサポート方針 / Store用画像 | 画像系。既存派生ビューは未変更。 |
| DAKE_Mail_Draft | 01_apps/DAKE_Mail_Draft/ORIGINAL.md | Store URL / Storeダウンロード導線 / Storeサポート方針 / Store用画像 | メール系。既存派生ビューは未変更。 |
| DAKE_Mail_List | 01_apps/DAKE_Mail_List/ORIGINAL.md | Store URL / Storeダウンロード導線 / Storeサポート方針 / Store用画像 | メール系。既存派生ビューは未変更。 |
| DAKE_Maji_Memo | 01_apps/DAKE_Maji_Memo/ORIGINAL.md | Store URL / Storeダウンロード導線 / Storeサポート方針 / Store用画像 | メモ系。既存派生ビューは未変更。 |
| DAKE_Git_Memo | 01_apps/DAKE_Git_Memo/ORIGINAL.md | Store URL / Storeダウンロード導線 / Storeサポート方針 / Store用画像 | メモ/開発記録系。既存派生ビューは未変更。 |
| DAKE_Work_Calendar | 01_apps/Dake_Work_Calendar/ORIGINAL.md | Store URL / Storeダウンロード導線 / Storeサポート方針 / Store用画像 | 作業補助系。既存派生ビューは未変更。 |
| DAKE_Mansion_Schedule | 01_apps/DAKE_Mansion_Schedule/ORIGINAL.md | Store URL / Storeダウンロード導線 / Storeサポート方針 / Store用画像 | 作業補助系。既存派生ビューは未変更。 |
| DAKE_Price_FixedTax | 01_apps/DAKE_Price_FixedTax/ORIGINAL.md | Store URL / Storeダウンロード導線 / Storeサポート方針 / Store用画像 | 計算系。既存派生ビューは未変更。 |
| DAKE_Price_Apportionment | 01_apps/DAKE_Price_Apportionment/ORIGINAL.md | Store URL / Storeダウンロード導線 / Storeサポート方針 / Store用画像 | 計算系。既存派生ビューは未変更。 |
| DAKE_Game_Alien_Road | 01_apps/DAKE_Game_Alien_Road/ORIGINAL.md | Store URL / Storeダウンロード導線 / Storeサポート方針 / Store用画像 | ゲーム系。既存派生ビューは未変更。 |
| DAKE_Game_Diver_Catch | 01_apps/DAKE_Game_Diver_Catch/ORIGINAL.md | Store URL / Storeダウンロード導線 / Storeサポート方針 / Store用画像 | ゲーム系。既存派生ビューは未変更。 |
| DAKE_Video_Shorts_Cut | 01_apps/DAKE_Video_Shorts_Cut/ORIGINAL.md | Store URL / Storeダウンロード導線 / Storeサポート方針 / Store用画像 / BOOTH URL | 動画系。既存派生ビューは未変更。 |

## 共通してうまくいった点

- README内DAKE_META、release_body.md、booth_ready/booth_product.txt から、基本情報・Release情報・BOOTH情報を集約できた。
- 改善済みテンプレートの任意セクションにより、CLI連携、対応形式、非ゴール、保存方針、非破壊方針を無理なく記録できた。
- Store関連の未確定情報を、推測で埋めずに未確定として残せた。
- Pack構成要素になっている単品アプリも、Pack情報を混ぜずに単品アプリ正本として扱えた。

## アプリごとに詰まった点

- メール系はOutlook連携や送信しない方針を、今後より明確な任意セクションにするとよい。
- 動画系は外部ツールや対応形式の確認が重要になりやすく、実装確認なしに断定しない運用が必要。
- ゲーム系は実務ツールとは違うため、対象ユーザー・プレイ体験・販売文のセクションを少し変えた方がよい可能性がある。
- 計算系は入力項目・計算式・免責の扱いを、専用の任意欄として整理できるとさらに安全。

## テンプレート改善候補

- `送信しない / 外部操作しない` など、外部サービス連携系の非ゴール欄を明示してもよい。
- ゲーム系向けに `遊び方 / 操作方法 / 難易度 / スコア方針` の任意セクションを検討する。
- 計算系向けに `入力項目 / 計算結果 / 免責` の任意セクションを検討する。
- 動画系向けに `外部ツール同梱 / PATH依存 / 対応コーデック` の任意セクションを検討する。

## 次Batchへの注意

- A1残件も、Pack構成要素であっても単品アプリとして `ORIGINAL.md` を作る。
- Store URL、Storeダウンロード導線、Storeサポート方針は、Store仕様が固まるまで未確定のまま残す。
- BOOTH URLやGitHub Release URLが既存ファイルにない場合は推測しない。
- 既存派生ビューは、Phase 4-5Eでは更新しない。

## Pack対応への影響

今回作成した単品アプリ `ORIGINAL.md` の一部は、DAKE Packの構成アプリにもなる。

ただし、Pack商品の価格・同梱構成・Pack販売文は単品アプリORIGINALには入れず、次Phase以降のPack用ORIGINALで扱う。

Pack側では、構成アプリの機能説明を再定義せず、各単品アプリORIGINALを参照する方針が妥当。

## 次Phase提案

1. Phase 4-5F: Pack用 `ORIGINAL.md` テンプレートを `00_core` に作成する。
2. Phase 4-5G: `DAKE_Pack_Document`、`DAKE_Pack_Memo`、SHIMARISU PackへPack用ORIGINALを導入する。
3. Phase 4-5H: A1単品アプリの残りへ、同じ方式でバッチ展開する。
4. Phase 4-5I: Store用generatedデータ形式を定義する。
