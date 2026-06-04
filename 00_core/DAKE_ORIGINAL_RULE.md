# DAKE ORIGINAL Rule

## 目的

DAKEシリーズでは、アプリ、サイト、記事、商品、販売ページなどの一次情報を `ORIGINAL.md` に集約する。

このファイルは、`ORIGINAL.md` を真の正本として扱うための共通ルールである。

## 1. ORIGINAL.md の定義

`ORIGINAL.md` は、対象アプリ・サイト・記事・商品の真の正本である。

編集すべき一次情報は `ORIGINAL.md` に集約する。

- 何のためのものか
- 誰のためのものか
- 何を解決するか
- どのように配布・販売するか
- どの説明文を公開に使うか
- どの注意事項を添えるか
- 今後どこを直すか

これらは、原則として `ORIGINAL.md` を起点に判断する。

## 2. README.md の位置づけ

`README.md` はGitHub公開用ビューであり、真の正本ではない。

`ORIGINAL.md` と `README.md` に矛盾がある場合は、`ORIGINAL.md` を優先する。

READMEには、GitHub上で必要な説明、ビルド方法、Release案内、公開可能な情報だけを置く。

## 3. DAKE_META の位置づけ

`DAKE_META` は機械利用ビューであり、真の正本ではない。

DAKE Launcher、dakeapp.com、管理ツール、レビュー用スクリプトなどが読みやすい形式に整えた派生物として扱う。

`DAKE_META` に正本になりそうな判断情報が増えた場合は、まず `ORIGINAL.md` へ戻す。

## 4. release_body.md の位置づけ

`release_body.md` はGitHub Release用ビューであり、真の正本ではない。

GitHub Release説明欄へ貼る短い文だけを置く。

リリース文の元情報、機能説明、注意事項、更新理由は `ORIGINAL.md` に置く。

## 5. booth_product.txt の位置づけ

`booth_product.txt` はBOOTH登録用ビューであり、真の正本ではない。

BOOTH用の商品名、価格案、説明文、タグ、商品画像、配布zip、公開後URLなどを登録画面に貼りやすい形にした派生物として扱う。

BOOTH用説明文の元情報は `ORIGINAL.md` に置く。

BOOTH公開後に `# URL` 欄へ戻したURLも、販売ビューの結果情報であり、真の正本そのものではない。

## 6. Storeの位置づけ

`store.dakeapp.com` は自社販売ビューであり、Store専用の商品正本を作らない。

Storeは可能な限り `ORIGINAL.md` 由来の情報を読む。

静的生成などの都合で `store_products.generated.json` のような生成物を作る場合も、それは正本ではない。

生成物は、いつでも `ORIGINAL.md` から作り直せる状態にする。

## 7. 新しいファイルが生まれる時の判断基準

新しいファイル、JSON、Markdown、CSV、HTML、登録用テキストが生まれる時は、必ず次を確認する。

```text
これは正本か？
それとも ORIGINAL.md から抽出されたビューか？
```

正本になりそうな情報は `ORIGINAL.md` に戻す。

ビューでよい情報は、README、DAKE_META、release_body、booth_product、Store生成物などへ分けてよい。

## 8. 過渡期ルール

まだ `ORIGINAL.md` がない既存アプリでは、暫定的に `README.md` を参照してよい。

ただし、移行対象アプリでは必ず `ORIGINAL.md` を作成し、以後は `ORIGINAL.md` を真の正本として扱う。

```text
移行前：README.md を暫定正本として扱う
移行後：ORIGINAL.md を真の正本として扱う
```

既存の出荷ツールやレビュー項目がREADMEを読んでいる場合も、移行後は `ORIGINAL.md` 由来の情報へ寄せる。

## 9. 公開情報と非公開情報

`ORIGINAL.md` には、内部メモ、販売戦略、未公開構想、Codex作業注意などが入る可能性がある。

そのため、公開repoに置いてよい内容かを必ず確認する。

- 公開できる内容だけなら、対象repo内に置いてよい。
- 公開できない内容を含む場合は、private repo またはローカル正本置き場で管理する。
- 公開用READMEへ転記する時は、内部メモや未公開構想を混ぜない。

## 10. Codex作業時の基本ルール

Codexは作業前に、原則として以下の順で確認する。

```text
1. ORIGINAL.md
2. README.md
3. DAKE_META
4. release_body.md
5. booth_product.txt
6. 関連仕様ファイル
```

優先順位は必ず以下とする。

```text
ORIGINAL.md
↓
その他の派生ファイル
```

`ORIGINAL.md` が存在しない既存対象では、過渡期ルールとしてREADMEを暫定参照してよい。

## 11. 派生ビューの基本対応

```text
ORIGINAL.md
↓
README.md              GitHub公開用ビュー
DAKE_META              Launcher / サイト / 管理ツール向け機械利用ビュー
release_body.md        GitHub Release用ビュー
booth_product.txt      BOOTH登録用ビュー
store.dakeapp.com      自社販売ビュー
```

派生ビューを直接編集した場合は、必要に応じて `ORIGINAL.md` へ戻す。

## 12. 今回のPhase 1でやらないこと

このルール策定Phaseでは、以下を行わない。

- 既存アプリへの `ORIGINAL.md` 一括導入
- README再生成
- release_body再生成
- booth_product再生成
- Store構築
- Stripe実装
- 全アプリ一括移行

まず `ORIGINAL.md` の定義を固定し、次Phaseで代表アプリ1本へ試験導入する。
