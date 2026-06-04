# ORIGINAL Phase4-5D Pack and SHIMARISU Plan

## 目的

Phase 4-5Cで改善した単品アプリ用 `ORIGINAL.md` テンプレートを、A優先対象へ横展開する前に、単品アプリ、Pack商品、SHIMARISU関連商品の正本構造を分けて整理する。

単品アプリ用 `ORIGINAL.md` をPack商品へ無理に当てず、StoreやBOOTHで扱う商品単位ごとに、どこへ真の正本を置くべきかを明確にする。

## 調査対象

- `C:\Users\yukiz\devlop\DAKE_series\01_apps`
- `C:\Users\yukiz\devlop\DAKE_series\04_packs`
- `C:\Users\yukiz\devlop\SHIMARISU`
- `C:\Users\yukiz\devlop\SHIMARISU\booth_ready`
- `C:\Users\yukiz\devlop\shimarisu-pack-release`

確認した主な情報源:

- `00_core/DAKE_ORIGINAL_RULE.md`
- `00_core/DAKE_ORIGINAL_TEMPLATE_APP.md`
- `tools/reports/original_phase45_app_inventory.md`
- `tools/reports/original_phase45b_pilot_rollout.md`
- `04_packs/DAKE_Pack_Document/README.md`
- `04_packs/DAKE_Pack_Document/pack_manifest.json`
- `04_packs/DAKE_Pack_Document/pack_ready/booth_product.txt`
- `04_packs/DAKE_Pack_Memo/README.md`
- `04_packs/DAKE_Pack_Memo/pack_manifest.json`
- `04_packs/DAKE_Pack_Memo/pack_ready/booth_product.txt`
- `SHIMARISU/booth_ready/booth_product.txt`
- `shimarisu-pack-release/README.md`

## 単品アプリ ORIGINAL 横展開方針

単品アプリは、各アプリフォルダ内の `ORIGINAL.md` を真の正本とする。

対象は原則として `01_apps/` 配下の、`status: available`、BOOTH登録済み、Releaseあり、Storeに単品商品として載せる可能性があるもの。

単品アプリの `ORIGINAL.md` は、そのアプリ単体の目的・機能・対象ユーザー・配布方針・注意事項・派生ビュー情報を持つ。Pack全体の価格、Pack構成、Pack専用の販売文は持たない。

Phase 4-5Bまでに、以下6アプリは `ORIGINAL.md` 導入済み。

- `DAKE_Time_AdvancedTimer`
- `DAKE_PDF_Compress`
- `DAKE_Image_Resize`
- `DAKE_Folder_List`
- `DAKE_Backup`
- `DAKE_Sticky_Memo`

## Pack商品 ORIGINAL 方針

Pack商品は、単品アプリとは別の商品単位として `ORIGINAL.md` を持つ。

Packの `ORIGINAL.md` は、Pack名、Pack ID、価格、同梱アプリ、zip構成、Packとしての価値、単品販売との差分、BOOTH/Store用説明、更新時の扱いを正本化する。

各構成アプリの機能説明や注意事項は、原則として各アプリ側 `ORIGINAL.md` を参照する。Pack側には、Packとして束ねる理由、同梱物、購入者がどう使い始めるか、Pack価格と販売導線だけを置く。

Pack側に持つ情報:

- Pack商品名
- Pack価格
- Packに含まれるアプリ・素材
- Pack zip構成
- Packとしての価値・使い始め方
- Pack用BOOTH/Store説明
- Pack更新時の扱い

各アプリ側へ戻すべき情報:

- 個別アプリの細かい機能説明
- 個別アプリの注意事項
- 個別アプリのRelease URL
- 個別アプリのBOOTH URL
- 個別アプリのスクリーンショット方針

## SHIMARISU関連 ORIGINAL 方針

SHIMARISU関連は、単品アプリ、Pack、キャラクター要素、ブランドサイト導線が混ざるため、DAKE単品アプリとは別枠で扱う。

`SHIMARISU/booth_ready/booth_product.txt` には、`しまりすくん 実務判断Pack` の商品情報、価格、同梱内容、画像、BOOTH URLが存在する。これは単品DAKEアプリではなく、SHIMARISU Pack商品の正本候補である。

`shimarisu-pack-release` はPack ZIP公開用repoであり、README上も「SHIMARISU Pack ZIPのみを公開するrepo」とされている。商品正本そのものではなく、配布物置き場・公開導線として扱う。

`DAKE_Game_ShimarisuRealEstate` は `status: prototype` のゲームアプリであり、現時点ではStore前の優先単品商品ではない。将来販売・公開する場合は、単品ゲームアプリとして `ORIGINAL.md` を持ち、SHIMARISU Packとは別の商品として扱う。

## A1：単品アプリ優先対象

| folder | title | original | booth | store_candidate | memo |
|---|---|---|---|---|---|
| DAKE_App_Doko | アプリどこ | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_Backup | Dakeバックアップ | yes | url | yes | Phase 4-5BでORIGINAL導入済み。 |
| DAKE_BOOTH_Assist | BOOTHアシスト | no | url | yes | 単品アプリだがFactory補助ツール色が強く、Store文言は慎重に扱う。 |
| DAKE_Column_Memo | ずっとメモ | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_Document_Cover | Dake書類送付状 | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_FAX_Cover | DakeFAX送付状 | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_Folder_List | Dakeフォルダ一覧 | yes | url | yes | Phase 4-5BでORIGINAL導入済み。 |
| DAKE_Game_Alien_Road | DakeAlien Road | no | url | yes | 単品ゲーム。A1対象。Packとは分ける。 |
| DAKE_Game_Diver_Catch | Dake潜って捕る | no | url | yes | 単品ゲーム。A1対象。Packとは分ける。 |
| DAKE_Git_Memo | DakeGitメモ | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_Image_HEICtoJPG | HEIC→JPG変換 | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_Image_iPhoneToPC | Dake画像iPhoneToPC | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_Image_PasteA4 | 貼る | no | url | yes | 単品アプリ。DAKE_Pack_Documentの構成要素でもある。 |
| DAKE_Image_Receiver | DakeImage_Receiver | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_Image_Resize | Dake画像リサイズ | yes | url | yes | Phase 4-5BでORIGINAL導入済み。DAKE_Pack_Documentの構成要素でもある。 |
| DAKE_Image_ToPDF | DakeImageToPDF | no | url | yes | 単品アプリ。DAKE_Pack_Documentの構成要素でもある。 |
| DAKE_Launcher | Dakeランチャー | no | url | yes | 配布導線の中核。Storeで単品扱いするかは要判断。 |
| DAKE_Mail_Address_Format | Dakeメールアドレス整形 | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_Mail_AllStaff | Dake全社員メール起動 | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_Mail_Draft | Dakeメール下書き | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_Mail_Kikuta | Dake菊田メール | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_Mail_List | Dakeメールリスト | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_Maji_Memo | マジでメモ | no | url | yes | 単品アプリ。DAKE_Pack_Memoの構成要素でもある。 |
| DAKE_Mansion_Schedule | マンション工程表 | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_PDF_CheckStamp | Dake確認印 | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_PDF_Compress | DakePDF圧縮 | yes | url | yes | Phase 4-5BでORIGINAL導入済み。 |
| DAKE_PDF_Crop | DakePDFトリミング | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_PDF_LookHere | DakePDFここ見て | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_PDF_Marker | DakePDFマーカー | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_PDF_Merge | DakePDF結合 | no | url | yes | 単品アプリ。DAKE_Pack_Documentの構成要素でもある。 |
| DAKE_PDF_Merge_Mini | DakePDF結合mini | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_PDF_Rename | DakePDFファイル名整理 | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_PDF_Reorder | DakePDFページ並べ替え | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_PDF_SplitOne | DakePDF分割One | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_PDF_SplitSelect | DakePDF分割Select | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_PDF_ToImages | DakePDFto画像 | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_PDF_Viewer | DakePDF見る | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_Price_Apportionment | Dake価格按分 | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_Price_FixedTax | Dake固都税計算 | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_Reform_Progress | リフォーム進捗管理 | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_Screen_WebP | DakeScreen_WebP | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_Screenshot_Print | Dakeスクショ印刷 | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_Sticky_Memo | 付箋メモ | yes | url | yes | Phase 4-5BでORIGINAL導入済み。DAKE_Pack_Memoの構成要素でもある。 |
| DAKE_Time_AdvancedTimer | Dakeアドバンスドタイマー | yes | url | yes | Phase 3でORIGINAL導入済み。 |
| DAKE_TwoPerson_Memo | Dake二人メモ | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_Video_Shorts_Cut | Dakeショート切り出し | no | ready | yes | BOOTH公開待ち。単品アプリとしてA1対象。 |
| DAKE_Work_Calendar | Dake工程カレンダー | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_Year_Age | Dake築年数 | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_Year_Notice | Dake今年の注意点 | no | url | yes | 単品アプリ。A1対象。 |
| DAKE_Yesterday_Task_Memo | Dake昨日タスクメモ | no | url | yes | 単品アプリ。DAKE_Pack_Memoの構成要素でもある。 |

## A2：Pack商品優先対象

| folder | title | original | booth | included_items | memo |
|---|---|---|---|---|---|
| 04_packs/DAKE_Pack_Document | DAKE 書類整理パック | no | https://peakheadz.booth.pm/items/8448353 | DAKE_PDF_Merge, DAKE_Image_ToPDF, DAKE_Image_Resize, DAKE_Image_PasteA4 | `PACK_META`、`pack_manifest.json`、`pack_ready/booth_product.txt`あり。Pack用ORIGINAL対象。 |
| 04_packs/DAKE_Pack_Memo | DAKE メモと記録パック | no | https://peakheadz.booth.pm/items/8449208 | DAKE_Sticky_Memo, DAKE_Maji_Memo, DAKE_Git_Memo, DAKE_Yesterday_Task_Memo | `PACK_META`、`pack_manifest.json`、`pack_ready/booth_product.txt`あり。Pack用ORIGINAL対象。 |

## A3：SHIMARISU関連優先対象

| folder | title | type | original | booth | memo |
|---|---|---|---|---|---|
| C:\Users\yukiz\devlop\SHIMARISU\booth_ready | しまりすくん 実務判断Pack | pack | no | https://peakheadz.booth.pm/items/8449321 | SHIMARISU Pack v1.0。`SHIMARISU_Pack.zip`、画像一式、booth_productあり。Pack用ORIGINAL対象。 |
| C:\Users\yukiz\devlop\shimarisu-pack-release | SHIMARISU Pack Release | release_repo | no | n/a | Pack ZIP公開用repo。商品正本ではなく、配布物置き場として扱う。 |
| 01_apps/DAKE_Game_ShimarisuRealEstate | Dakeしまりす不動産 | prototype_game | no | n/a | しまりすキャラクター系ゲーム。現状 `status: prototype` でStore前必須ではない。公開時は単品ゲームORIGINAL対象。 |

## Pack用 ORIGINAL.md テンプレート案

Pack用テンプレートは、単品アプリ用テンプレートとは別に作成する。

想定セクション:

```md
# ORIGINAL.md

## 正本宣言

このファイルは、このPack商品の真の正本です。
構成アプリ側ORIGINAL.md、README.md、PACK_META、pack_manifest.json、booth_product.txt、Store表示などは、このファイルから派生または参照されるビューです。

## 基本情報

- pack_id:
- product_type: pack
- title:
- short_title:
- status:
- version:
- price:
- distribution:
- target_platform:

## Packの目的

## 対象ユーザー

## Packとしての価値

## 単品販売との差分

## 含まれるアプリ・素材

| item | type | source | role | original |
|---|---|---|---|---|

## 各構成物のフォルダ

## 構成アプリ側ORIGINALとの関係

## Pack側で正本化してよい情報

## 各アプリORIGINALへ戻すべき情報

## BOOTH商品情報の元情報

## Store表示用情報

## 価格・販売方針

## 配布・ダウンロード方針

## 同梱ファイル一覧

## zip構成

## 画像方針

## サポート方針

## 更新時の扱い

## 免責・注意事項

## Codex作業時の注意

## 派生物一覧

- README.md:
- PACK_META:
- pack_manifest.json:
- booth_product.txt:
- pack_ready:
- Store:
```

## 単品アプリ ORIGINAL と Pack ORIGINAL の関係

```text
単品アプリ ORIGINAL
= そのアプリ自体の正本

Pack ORIGINAL
= 複数アプリ・素材を束ねた商品としての正本
```

Pack ORIGINALは、構成アプリの機能説明を再定義しない。構成アプリの機能・注意事項・更新方針は各アプリの `ORIGINAL.md` を参照する。

Pack ORIGINALは、Packとしての価値、価格、同梱物、zip構成、Pack販売文、購入後の使い始め方、更新時のPack再生成方針を持つ。

Pack内に含まれる単品アプリも、それぞれ単品アプリとしてStore掲載される可能性がある。そのため、Packと単品アプリは親子ではなく、商品単位の別正本として扱う。

## Store接続前に必須のORIGINAL

Storeへ掲載する商品だけを必須とする。

優先:

1. A1のうち、Store初期掲載する単品アプリ
2. A2のうち、Store初期掲載するDAKE Pack商品
3. A3のうち、Store初期掲載するSHIMARISU Pack

現時点でStore前に先行して整えるべきもの:

- A1単品アプリ: まず10〜15件単位でバッチ展開
- A2 Pack商品: `DAKE_Pack_Document`, `DAKE_Pack_Memo`
- A3 SHIMARISU Pack: `しまりすくん 実務判断Pack`

`DAKE_Game_ShimarisuRealEstate` は現状prototypeのため、Store初期掲載対象に含めない。

## Store後でもよいORIGINAL

- frozen / draft / experimental / internal / prototype
- Store初期掲載しない単品アプリ
- 過去参照用Pack
- privateまたは社内運用色の強いもの
- SHIMARISU関連でも、商品として未確定のキャラクター素材や試作ゲーム

## 次Phase提案

1. Phase 4-5E: A1単品アプリから10〜15件を選び、改善済みテンプレートで `ORIGINAL.md` を作成する。
2. Phase 4-5F: Pack用 `ORIGINAL.md` テンプレートを `00_core` に作成する。
3. Phase 4-5G: `DAKE_Pack_Document`、`DAKE_Pack_Memo`、`SHIMARISU Pack` へPack用 `ORIGINAL.md` を導入する。
4. Phase 4-5H: A1単品アプリの残りへ `ORIGINAL.md` を展開する。
5. Phase 4-5I: Store用 generated データ形式を定義する。

## 注意点

- 今回は調査・整理のみで、`ORIGINAL.md` は新規作成していない。
- 単品アプリ用テンプレートをPack商品へ使い回さない。
- Pack側は、構成アプリの説明を丸ごと正本化しない。
- PackのBOOTH URLはPack側の販売結果情報として保持する。
- Pack内アプリのBOOTH URLやRelease URLは、各アプリ側 `ORIGINAL.md` へ戻す。
- `SHIMARISU` 本体repoはprivate性があるため、公開repoへ転記する情報は `booth_ready` と公開用READMEに限定して扱う。
- `shimarisu-pack-release` は商品正本ではなく、Pack ZIP公開用repoとして扱う。
- Storeでは、単品商品ページとPack商品ページを分ける。
