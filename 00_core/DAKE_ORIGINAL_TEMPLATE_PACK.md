# ORIGINAL.md

## 正本宣言

このファイルは、このPack商品の真の正本です。

README.md、PACK_META、pack_manifest.json、booth_product.txt、pack_ready、Store表示などは、このファイルから派生するビュー、またはこのファイルを補助する参照情報です。

単品アプリ側の `ORIGINAL.md` は、各アプリ自体の真の正本です。

Pack側の `ORIGINAL.md` は、複数アプリ・素材・説明・価格・zip構成を束ねる商品としての真の正本です。

## 基本情報

- pack_id:
- title:
- short_title:
- category:
- status:
- version:
- price:
- distribution:
- target_platform:
- pack_type:

## Packの目的

このPack商品が何のために存在するかを書く。

単品アプリをただ集めたものではなく、Packとしてどの作業や体験をまとめるのかを書く。

## 対象ユーザー

このPackを買う・使う人を書く。

## Packとして解決する困りごと

単品アプリを個別に探す手間、組み合わせる迷い、導入時のつまずきなど、Packだから減らせる困りごとを書く。

## Packとしての価値

Packとして束ねる理由を書く。

例:

- よく一緒に使うアプリをまとめる
- 導入時の迷いを減らす
- 価格を単品購入より分かりやすくする
- 同じ作業領域の道具をひとつの入口にする

## 同梱アプリ・構成物

| item | type | folder | role | original |
|---|---|---|---|---|
|  | app / asset / document / other |  |  |  |

## 単品販売との差分

単品アプリとして買う場合と、Packとして買う場合の違いを書く。

- 単品販売との価格差:
- Pack限定の同梱物:
- Pack限定の説明:
- Pack限定ではないもの:

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

- zip名:
- 同梱exe:
- 同梱README:
- 注意事項:
- 入れないもの:

## 公開用説明の元情報

GitHub、BOOTH、Store、dakeapp.comなどで共通して使えるPack説明の元情報を書く。

## README生成用情報

README.mdへ出す内容を書く。

- 概要:
- 収録内容:
- 使い方:
- 必要なもの:
- 注意:

## PACK_META生成用情報

PACK_METAやpack_manifest.jsonへ変換するための情報を書く。

```json
{
  "folder_name": "",
  "display_name": "",
  "product_type": "pack",
  "status": "available",
  "version": "1.0.0",
  "price": 0,
  "booth_url": "",
  "included_apps": [],
  "show_in_booth_assist": true,
  "show_on_dashboard": true,
  "tags": [],
  "summary": "",
  "copyright": "PEAKHEADZ"
}
```

## booth_product生成用情報

BOOTH登録用ビューへ変換するための情報を書く。

- 商品名:
- 価格案:
- 商品紹介文:
- 補足紹介文:
- タグ:
- 商品画像:
- 補助画像:
- 作品ファイル:
- BOOTH URL:

## Store表示用情報

store.dakeapp.com へ出すPack商品情報の元情報を書く。

- 商品名:
- キャッチ:
- 説明:
- 価格:
- 画像:
- ダウンロード導線:
- サポート方針:

## 価格・販売方針

Pack価格、単品合計との差分、BOOTH、Store、自社販売、応援購入などの方針を書く。

## 配布・ダウンロード方針

BOOTH、Store、Pack用Release、dakeapp.comなどのどこから取得できるかを書く。

## 免責・注意事項

Packとして公開前に必ず出す注意事項を書く。

例:

- Windows向けPackです。
- 各アプリの詳細な使い方は、収録アプリ側のREADMEを確認してください。
- 大切なファイルは事前にバックアップを推奨します。
- 本Packおよび収録アプリの無断転載・再配布を禁止します。
- 環境によっては起動時にWindowsの警告が表示される場合があります。

## 同梱ファイル方針

Pack zipや配布物へ何を入れるかを書く。

- apps/:
- README.txt:
- 注意事項.txt:
- 画像:
- 入れないもの:

## スクリーンショット・画像方針

Pack商品画像の方針を書く。

- assets/booth_thumbnail.jpg:
- pack_ready/booth_thumbnail.jpg:
- 補助画像:
- Store用画像:

## 更新時の扱い

- 構成アプリ更新時:
- Pack zip更新時:
- BOOTH更新時:
- Store更新時:
- バージョン表記:

## 構成アプリ側 ORIGINAL.md との関係

Pack側はPack商品の正本であり、構成アプリ自体の正本ではない。

- Pack側で持つ情報:
- 各アプリORIGINALへ戻すべき情報:
- 重複させない情報:
- 参照するORIGINAL:

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

Pack商品として、あえてやらないことを書く。

- 例: 構成アプリの機能説明をPack側で再定義しない
- 例: Packにソースコードを同梱しない
- 例: Pack価格を各アプリ側ORIGINALへ書かない

## 今後の改善予定

未実装、将来候補、やらないことを書く。

## Codex作業時の注意

Codexが触ってよい範囲、触ってはいけない範囲を書く。

- 触ってよい:
- 触らない:
- 外部公開しない:
- 自動操作しない:

## 派生物一覧

- README.md:
- PACK_META:
- pack_manifest.json:
- booth_product.txt:
- pack_ready:
- Store:
