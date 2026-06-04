# ORIGINAL Phase4-5H Pack Rollout

## 目的

Phase 4-5Gで作成したPack用 `ORIGINAL.md` テンプレートを、DAKE Pack 2件へ適用できるか確認する。

今回は `04_packs` 配下のPack商品のみを対象にし、SHIMARISU Packや単品アプリは変更しない。

## 対象Pack

| folder | title | source files | result |
|---|---|---|---|
| `04_packs/DAKE_Pack_Document` | DAKE 書類整理パック | README.md, pack_manifest.json, pack_ready/booth_product.txt, pack_ready/README.txt, pack_ready/注意事項.txt, assets/booth_thumbnail.jpg, pack_ready/booth_thumbnail.jpg, pack_ready/DAKE_Pack_Document.zip | ORIGINAL.md作成 |
| `04_packs/DAKE_Pack_Memo` | DAKE メモと記録パック | README.md, pack_manifest.json, pack_ready/booth_product.txt, pack_ready/README.txt, pack_ready/注意事項.txt, assets/booth_thumbnail.jpg, pack_ready/booth_thumbnail.jpg, pack_ready/DAKE_Pack_Memo.zip | ORIGINAL.md作成 |

## 作成した ORIGINAL.md

| folder | original_path | missing_info | memo |
|---|---|---|---|
| `04_packs/DAKE_Pack_Document` | `04_packs/DAKE_Pack_Document/ORIGINAL.md` | version, DAKE_Image_PasteA4 ORIGINAL | Pack価格・構成・BOOTH URL・zip構成をPack側に集約。 |
| `04_packs/DAKE_Pack_Memo` | `04_packs/DAKE_Pack_Memo/ORIGINAL.md` | version, DAKE_Yesterday_Task_Memo ORIGINAL | Pack価格・構成・BOOTH URL・zip構成をPack側に集約。 |

## Pack用テンプレートが有効だった点

- Pack名、価格、BOOTH URL、Pack ZIP、同梱アプリ、zip構成を単品アプリ情報と分けて整理できた。
- `Packとしての価値` と `構成アプリ側 ORIGINAL.md との関係` によって、Pack側がアプリ機能説明を再定義しない形にできた。
- `pack_manifest.json` のハッシュやZIPサイズは、Pack生成結果の参照情報として整理できた。
- `booth_product.txt` の商品紹介文、タグ、BOOTH URLをPack商品ビューとして扱えた。

## 詰まった点

- 既存Pack情報にPack versionが明示されていないため、`version` は未設定として残した。
- `DAKE_Image_PasteA4` と `DAKE_Yesterday_Task_Memo` は構成アプリ側 `ORIGINAL.md` が未作成のため、当面はREADME暫定参照扱いにした。
- Storeのダウンロード導線、サポート方針、Store用画像は未確定として残した。

## 構成アプリORIGINALとの関係

- Pack ORIGINALはPack商品の正本であり、構成アプリ自体の正本ではない。
- Pack側に置いた情報は、Pack価格、Pack価値、構成、Pack ZIP、Pack販売文、Pack画像、Pack更新方針。
- 各アプリの詳細機能、注意事項、Release URL、BOOTH URL、スクリーンショット方針は各アプリ側 `ORIGINAL.md` を優先する。
- ORIGINAL未作成の構成アプリは、次の単品アプリ横展開で作成する候補として残す。

## SHIMARISU Pack導入前の注意

- SHIMARISU PackはDAKE_series外の `SHIMARISU/booth_ready` を参照するため、Pack ORIGINALの置き場所を先に決める。
- SHIMARISU本体repoのprivate性に注意し、DAKE_seriesへZIPや内部情報を持ち込まない。
- `shimarisu-pack-release` は配布物置き場であり、商品正本そのものではない。
- SHIMARISU Packでは、しまりすくん本体、Pack ZIP、画像、販売文、サイト導線の境界を明記する。

## 次Phase提案

1. SHIMARISU Pack用 `ORIGINAL.md` の置き場所を決める。
2. SHIMARISU PackへPack用ORIGINALを導入する。
3. `DAKE_Image_PasteA4` と `DAKE_Yesterday_Task_Memo` を含め、残りA1単品アプリへORIGINALを展開する。
4. Store用generatedデータ形式を、単品アプリORIGINALとPack ORIGINALの両方を読める形で定義する。
