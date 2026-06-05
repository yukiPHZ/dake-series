# DAKE_RELEASE_FLOW

## DAKE正式出荷ライン

DAKEの正式出荷ラインは以下で固定します。

```text
ORIGINAL.md
↓
README.md
↓
release_body.md
↓
assets/screenshot.webp
↓
assets/booth_thumbnail.jpg
↓
booth_product.txt
↓
build
↓
dist/*.exe
↓
GitHub Release
↓
BOOTH ready
↓
BOOTH
↓
dakeapp.com
↓
store_products.generated.json 再生成
↓
dake-store-site 同期
↓
store.dakeapp.com 商品詳細確認
↓
payment_status 確認
↓
Cloudflare反映確認
↓
正式出荷完了
```

GitHub Release公開のみでは正式出荷完了とはしません。
以下が揃って初めて正式出荷とします。

- `ORIGINAL.md` が最新（未導入の既存アプリでは `README.md` を暫定参照）
- `README.md` がGitHub公開用ビューとして最新
- `release_body.md` が最新
- `assets/screenshot.webp` が存在する
- `assets/booth_thumbnail.jpg` が存在する
- `booth_product.txt` が存在する
- `booth_ready/` が生成されている
- `dist/*.exe` が生成されている
- GitHub Release が作成され、exeが添付されている
- README の `DAKE_META.release_url` が派生ビューとして更新済み
- BOOTH掲載準備が完了している
- dakeapp.com に反映されている
- `tools/store/sync_store_to_site.py` で `store_products.generated.json` を再生成し、dake-store-siteへ同期している
- store.dakeapp.comの商品詳細ページを確認している
- `payment_status`（`stripe_ready` / `booth_only` / `preparing` / `not_for_sale`）を確認している
- Cloudflare反映確認済み

上記が揃って初めて「正式出荷完了」とします。
GitHub Release作成のみで出荷完了と報告しません。

## Codex出荷テンプレ必須項目

今後のCodex出荷指示では、以下を必ず確認対象に含めます。

- `assets/screenshot.webp` を生成・確認する。
- `assets/booth_thumbnail.jpg` を生成・確認する。
- `booth_ready/booth_product.txt` を生成・確認する。
- `booth_ready/` を生成・確認する。
- `dist/*.exe` をGitHub Releaseへ添付する。
- READMEの `DAKE_META.release_url` をRelease作成後に更新する。
- README / DAKE_META / release_body / booth_product を直接編集する場合は、`ORIGINAL.md` に戻すべき一次情報かを確認する。
- dakeapp.com掲載に必要な `display_name`、`site_title`、`site_description`、`update_summary`、`release_url`、`screenshot_path` を確認する。
- dakeapp.com反映後、Cloudflare上の本番URLで200、スクショ、Releaseリンクを確認する。
- Store反映時は `tools/store/sync_store_to_site.py` を実行し、`store_products.generated.json` 再生成、dake-store-site同期、store.dakeapp.com商品詳細、`payment_status` を確認する。

BOOTH登録最適化の正本は `DAKE_BOOTH_REGISTER_SPEC.md` です。
DakeBOOTHアシストは登録入力補助に使えますが、公開ボタンは押さず、最終公開は人間が確認します。

DAKEアプリ完成後の公開手順です。

## 前提

- GitHub Release本体の編集は、明示された公開作業時だけ行う。
- Release説明文の元情報は `ORIGINAL.md`。
- `ORIGINAL.md` 未導入の既存アプリでは、各アプリREADMEの `RELEASE_BODY` を暫定参照する。
- `release_body.md` はGitHub Releaseへ貼るための派生ビュー。
- `release_url` はRelease作成後にREADMEへ後入れする。

## 1. 完成確認

- `main.py` が起動する。
- `build.bat` が通る。
- `dist/<exe_name>.exe` が生成される。
- 共通アイコンが反映される。
- 文字化けがない。
- UI_TEXT管理が守られている。
- 起動直後の画面が整っている。

## 2. README整備

各アプリREADMEはGitHub公開用ビューとして以下を置く:

- `DAKE_META`
- `RELEASE_BODY`
- 必要な使い方、仕様、注意事項

`DAKE_META.status`:

- `available`: exe、README、screenshotが揃っている。
- `missing_exe`: dist内exeがない。
- `build_failed`: buildまたは起動、スクショ取得に失敗した。

## 3. release_body.md生成

- `ORIGINAL.md` のRelease用情報を優先する。未導入アプリではREADMEの `RELEASE_BODY` を読み取る。
- 各アプリフォルダ直下に `release_body.md` を生成する。
- 3〜5行の箇条書きにする。
- 長い説明は書かない。

## 4. スクリーンショット作成

- `dist` 内exeを起動する。
- 起動直後のアプリウインドウだけを撮る。
- `assets/screenshot.webp` に保存する。
- 横幅1200pxを超える場合だけ縮小する。
- 引き延ばしは禁止。

詳細は `DAKE_SCREENSHOT_RULE.md` を参照する。

## 5. Git確認

- ソース、README、release_body、screenshotを確認する。
- `build/`、`dist/`、`*.spec`、`*.exe` は原則Gitに含めない。
- GitHub Releaseへ添付するexeは、Git管理ではなくRelease成果物として扱う。

## 6. GitHubへpush

- 変更内容を確認する。
- 必要なファイルだけcommitする。
- GitHubへpushする。
- Release作成前にREADMEやスクショの最終状態を確認する。

## 7. GitHub Release作成

Releaseごとに確認する項目:

- 対象アプリ名。
- 添付するexe。
- `release_body.md` の内容。
- バージョンまたは日付。
- 既存Releaseとの重複。

Release説明欄には `release_body.md` を貼る。

## 8. exe添付

- `dist/<exe_name>.exe` を添付する。
- 添付後、ダウンロードできることを確認する。
- 必要ならexe名とReleaseタイトルを揃える。

## 9. release_url後入れ

Release作成後:

- 各アプリREADMEの `DAKE_META.release_url` にRelease URLを入れる。
- `ORIGINAL.md` 導入済みアプリでは、Release URLや公開状態が正本へ戻すべき情報か確認する。
- `update_summary` に公開内容を短く反映する。
- `status` は問題なければ `available` のままにする。

## 10. dakeapp.com反映

- READMEの `DAKE_META` を機械利用ビューとして読み込む。
- `site_title`、`site_description`、`release_url`、`screenshot_path` を反映する。
- `show_on_site: true` のアプリだけ掲載する。
- スクリーンショットは `assets/screenshot.webp` を使う。

## 11. Cloudflare反映確認

- `git push` 後、dakeapp.com本番URLの反映を確認する。
- `/apps/`、`/update/`、`/sitemap.xml`、対象アプリ詳細ページが200で返ることを確認する。
- 対象アプリ詳細ページにRelease URL、説明文、スクリーンショットが反映されていることを確認する。
- Cloudflare反映確認まで終わってから、正式出荷完了として報告する。

## 公開前の最終チェック

- Release本文が短い。
- exeが対象アプリのもの。
- スクリーンショットが最新。
- READMEの `exe_name` と添付exeが一致。
- `release_url` はRelease作成後に更新されている。
- BOOTH ready素材が生成済み。
- dakeapp.comに表示されている。
- Cloudflare反映確認済み。
- GitHub Releaseのみで出荷完了と扱っていない。

## 正式出荷ライン v2

通常出荷候補は `DAKE_META.status: available` のアプリのみに限定する。

以下の状態は通常出荷チェックから除外し、別枠で扱う。

- `frozen`
- `draft`
- `experimental`
- `private`

BOOTH公開後は `booth_ready/booth_product.txt` の `# URL` 欄へBOOTH URLを戻す。
BOOTH URL未記入のままCLOSED扱いしない。

`show_in_launcher` / `show_on_site` は `available` の公開対象にのみ true を付ける。
`frozen`、`draft`、`experimental`、`private` が true を持つ場合は、出荷ライン上の矛盾として点検する。


## Formal Shipping Line v3

Normal shipping candidates are limited to `DAKE_META.status: available`.

Excluded statuses:

- `frozen`
- `draft`
- `experimental`
- `private`
- `internal`

`internal` means an operations or management tool. It is excluded from normal distribution, BOOTH, dakeapp.com, and Launcher listing.
`show_in_launcher` and `show_on_site` should be true only for public `available` apps.
If a non-available app has either flag set to true, treat it as a formal shipping line conflict.

### v3 exe launch check

For `status: available` apps, the actual `dist/*.exe` launch check is required before treating the app as ready for release.

Priority:

1. Run `dist/*.exe --launch-check`.
2. If `--launch-check` is not implemented yet, run a short GUI smoke launch and close it.

New DAKE apps should implement `--launch-check` by default. The command must only verify launchability, return quickly, and exit with code `0` on success. It must not run file conversion, publishing, sending, browser automation, BOOTH operations, or other external side effects. Failures should exit with code `1` and a short stderr message.


## DAKE正式出荷定義 v2（Store対応）

今後のDAKE正式出荷は、GitHub Release、BOOTH、dakeapp.comに加えて、store.dakeapp.comへの掲載・同期・本番確認までを含める。

Storeは正本ではない。真の正本は各商品の `ORIGINAL.md` である。

```text
ORIGINAL.md
↓
store_products.generated.json
↓
dake-store-site
↓
store.dakeapp.com
```

正式出荷時は以下を確認する。

1. `ORIGINAL.md` 更新。
2. `README.md` / `DAKE_META` / `release_body.md` / `booth_product.txt` の整合確認。
3. GitHub Release URL確認。
4. BOOTH URLまたはBOOTH導線確認。
5. dakeapp.com掲載URL確認。
6. `python tools/store/sync_store_to_site.py` によるStore generated JSON再生成・同期。
7. store.dakeapp.comの商品詳細URL確認。
8. `payment_status` 確認。
9. Stripe Payment Link有無、BOOTH導線有無、準備中表示確認。
10. Cloudflare Pages反映確認。

`payment_status` の扱い:

- `stripe_ready`: Stripe Payment Linkあり。
- `booth_only`: BOOTH導線のみ。
- `preparing`: 準備中。
- `not_for_sale`: 販売対象外。

Stripe Secret、APIキー、Webhook Secretは、公開repo、generated JSON、Store JavaScriptへ絶対に入れない。

Codexの正式出荷完了報告には、store.dakeapp.com商品詳細URL、`payment_status`、Stripe Payment Link有無、BOOTH導線有無、Store同期結果、Cloudflare Pages反映確認を含める。
