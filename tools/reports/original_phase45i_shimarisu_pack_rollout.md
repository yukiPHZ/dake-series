# ORIGINAL Phase4-5I SHIMARISU Pack Rollout

## 目的

DAKE Pack 2件でPack用 `ORIGINAL.md` が成立したため、SHIMARISU PackにもPack商品用の真の正本を導入する。

## 正本の置き場所

SHIMARISU Packの真の正本は以下に置いた。

```text
C:\Users\yukiz\devlop\SHIMARISU\ORIGINAL.md
```

理由:

- `SHIMARISU/` はSHIMARISU本体、booth_ready、Pack素材を持つ母艦。
- `shimarisu-pack-release/` はPack ZIP公開用repoであり、商品正本ではない。
- `DAKE_series/` はDAKEアプリ群の母艦であり、SHIMARISU Packそのものの正本置き場ではない。

## 参照したファイル

- `DAKE_series/00_core/DAKE_ORIGINAL_RULE.md`
- `DAKE_series/00_core/DAKE_ORIGINAL_TEMPLATE_PACK.md`
- `DAKE_series/tools/reports/original_phase45d_pack_and_shimarisu_plan.md`
- `DAKE_series/tools/reports/original_phase45h_pack_rollout.md`
- `SHIMARISU/README.md`
- `SHIMARISU/release_body.md`
- `SHIMARISU/booth_ready/booth_product.txt`
- `SHIMARISU/booth_ready/README.txt`
- `SHIMARISU/booth_ready/注意事項.txt`
- `SHIMARISU/booth_ready/SHIMARISU_Pack.zip`
- `SHIMARISU/booth_ready/images/`
- `shimarisu-pack-release/README.md`

## 作成した ORIGINAL.md

| path | result |
|---|---|
| `C:\Users\yukiz\devlop\SHIMARISU\ORIGINAL.md` | 作成 |

## SHIMARISU Packとして整理した内容

- 商品名: しまりすくん 実務判断Pack
- 英字補助: SHIMARISU Pack v1.0
- 価格案: 3,000円
- BOOTH URL: https://peakheadz.booth.pm/items/8449321
- Pack ZIP: `booth_ready/SHIMARISU_Pack.zip`
- Pack ZIP size: 199722019 bytes
- 商品画像: `booth_ready/images/booth_thumbnail.jpg`
- 追加画像: 01_start.jpg, 02_checking.jpg, 03_decision.jpg, booth_thumbnail.jpg, screenshot.jpg
- 同梱内容: ShimarisuKun.exe, apps/DakePDF_Merge.exe, apps/DakePDF_Compress.exe, apps/DakePDF_ToImages.exe, apps/DakeImage_ToPDF.exe, apps/DakeImage_Resize.exe, README.txt, START_HERE.txt

## shimarisu-pack-releaseとの関係

`shimarisu-pack-release` はPack ZIP公開用repoであり、商品正本ではない。

SHIMARISU Packの情報は `SHIMARISU/ORIGINAL.md` を正本とし、`shimarisu-pack-release` は配布・公開用ビューとして扱う。

## DAKE_seriesとの関係

- DAKE_seriesはDAKEアプリ群の母艦。
- SHIMARISU PackはDAKE_seriesに正本を置かない。
- 同梱DAKE exeやCLI連携仕様はDAKE側に由来するが、Pack商品としての価格・販売文・画像・zip構成はSHIMARISU Pack ORIGINALで扱う。
- DAKE_Game_ShimarisuRealEstateはprototype/別枠であり、SHIMARISU Packへ混ぜない。

## 未確定として残した情報

- Store URL
- Storeダウンロード導線
- Storeサポート方針
- Store用画像
- Pack versionの次回更新ルール
- 社内利用範囲、複数拠点、外部委託先、別法人への共有可否
- 返金対応の最終文言

## 触らなかったもの

- `SHIMARISU/booth_ready/`
- `SHIMARISU/booth_ready/SHIMARISU_Pack.zip`
- `SHIMARISU/booth_ready/images/`
- `SHIMARISU/README.md`
- `SHIMARISU/release_body.md`
- `shimarisu-pack-release/`
- `DAKE_series/01_apps`
- `DAKE_series/04_packs`
- Store / Stripe関連

## 次Phase提案

1. 残りA1単品アプリへ `ORIGINAL.md` を展開する。
2. Store用generatedデータ形式を、単品アプリORIGINAL、DAKE Pack ORIGINAL、SHIMARISU Pack ORIGINALの3系統を読める形で定義する。
3. shimarisu.dakeapp.comへBOOTH購入導線を追加する場合は、SHIMARISU ORIGINAL由来の文言を使う。
