# DAKE_COMMON_SPEC

## DAKE正式出荷ライン

DAKEシリーズの正式出荷は、GitHub Release公開だけでは完了としません。
ORIGINAL.md正本運用、README公開ビュー、BOOTH販売素材、dakeapp.com掲載、store.dakeapp.com掲載までを含めて、初めて正式出荷とします。

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

## Releaseと正式出荷の定義

- Release = GitHub上の配布物公開。
- 出荷 = ユーザーが触れる状態。
- GitHub Release公開のみでは、DAKEの正式出荷完了とは扱わない。
- ユーザーが一覧で見つけ、スクショで理解し、BOOTH、dakeapp.com、store.dakeapp.comから迷わず取得できる状態までを正式出荷とする。

正式出荷に必要なもの:

- `ORIGINAL.md`（移行済みアプリ。未導入の既存アプリでは `README.md` を暫定参照）
- `README.md`
- `release_body.md`
- `assets/screenshot.webp`
- `assets/booth_thumbnail.jpg`
- `booth_product.txt`（通常は `booth_ready/booth_product.txt`）
- `booth_ready/`
- `dist/*.exe`
- GitHub Release
- BOOTH ready
- BOOTH掲載または掲載準備
- dakeapp.com掲載状態
- store.dakeapp.com掲載状態
- `store_products.generated.json` 再生成・dake-store-site同期
- Storeの商品詳細URLと `payment_status` 確認
- Cloudflare反映確認

Releaseだけ、exeだけ、READMEだけの状態は正式出荷完了とは扱いません。

DAKEシリーズの最上位共通仕様です。個別仕様や実装判断で迷った場合は、このファイルを優先します。



## Store掲載と正式出荷

Store掲載は、DAKE正式出荷の一部である。

Storeは正本ではなく、`ORIGINAL.md` 由来の `store_products.generated.json` を読む販売ビューとして扱う。

正式出荷時は、`tools/store/sync_store_to_site.py` を実行し、generated JSONの再生成、dake-store-siteへの同期、store.dakeapp.comの商品詳細確認、`payment_status` 確認を行う。

Store側で商品名、価格、説明、Stripe Payment Link、BOOTH URLを手編集しない。

## DAKEシリーズの考え方

- 現場で止まらない道具を作る。
- 1アプリは1つの仕事だけを終わらせる。
- 起動してすぐ分かる、迷わず押せる、すぐ閉じられることを重視する。
- 高機能より、安定・軽量・説明不要を優先する。
- 完璧な汎用ツールではなく、具体的な小さな困りごとを終わらせる。

## 単機能思想

- 1アプリ = 1目的。
- 目的外の便利機能を足さない。
- 設定画面、モード切替、詳細オプションは必要最小限にする。
- 迷う選択肢を増やすくらいなら、仕様を固定する。
- 複数の目的が出たら、別アプリに分けることを検討する。

## 応答性と安定性

- 起動直後に主目的が分かる初期画面にする。
- 重い処理はUIを固めない。必要に応じて別スレッド化する。
- 処理中・完了・失敗をステータス表示で伝える。
- 元ファイルは原則上書きしない。
- 保存名の衝突は連番や別名で回避する。
- 失敗時は復旧しやすい短い日本語で伝える。

## 命名

- フォルダ名は `DAKE_カテゴリ_機能` を基本にする。
- exe名はReleaseで見ても分かる英数字中心の名前にする。
- 表示名は日本語でよい。短く、用途が一瞬で分かる名前にする。
- READMEの `DAKE_META.folder_name` と実フォルダ名を一致させる。
- READMEの `DAKE_META.exe_name` と `dist` 内の配布exe名を一致させる。

## ORIGINAL.md と派生ビュー

各アプリの真の正本は `ORIGINAL.md` です。

`README.md` はGitHub公開用ビューであり、真の正本ではありません。
ただし、まだ `ORIGINAL.md` がない既存アプリでは、移行までの暫定参照として `README.md` を読んでよいものとします。

Codex作業時の確認順:

1. `ORIGINAL.md`
2. `README.md`
3. `DAKE_META`
4. `release_body.md`
5. `booth_product.txt`
6. 関連仕様ファイル

README / DAKE_META / release_body / booth_product を直接編集する場合は、その変更が `ORIGINAL.md` に戻すべき一次情報かを確認します。

必須ブロック:

- `DAKE_META`: ランチャー、dakeapp.com、Release連携用の機械利用ビュー。
- `RELEASE_BODY`: GitHub Release説明欄に貼る3〜5行の短文ビュー。

`DAKE_META` の必須項目:

```json
{
  "app_key": "",
  "display_name": "",
  "launcher_title": "",
  "launcher_description": "",
  "site_title": "",
  "site_description": "",
  "update_summary": "",
  "folder_name": "",
  "exe_name": "",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## UI_TEXT

- 画面表示の日本語は `UI_TEXT` に集約する。
- `APP_NAME`、`WINDOW_TITLE`、`COPYRIGHT` を上部に定義する。
- UI部品の `text="..."` 直書きは避ける。
- エラーメッセージ、ステータス、ボタン、フッター文言も `UI_TEXT` 管理にする。
- 日本語文字化けは放置禁止。

## 共通フッター

基本要素:

- `シンプルそれDAKEシリーズ`
- `止まらない、迷わない、すぐ終わる。`
- `戸建買取査定`
- `Instagram`
- `© 2026 しまりす不動産 / Vibe-Coded by Yukihiko Kikuta`

ルール:

- 画面下部に控えめに置く。
- 主機能より目立たせない。
- 狭い幅では中央寄せ2段などにして崩れを防ぐ。
- リンクは通常時に補助色、ホバー時にアクセント色を基本にする。

## 共通アイコン

- 全アプリは `02_assets/dake_icon.ico` を参照する。
- アプリごとの個別アイコンは作らない。
- アイコンで機能識別しない。機能は名前と画面で伝える。

## 禁止事項の要約

- 多機能化しない。
- UI演出を増やしすぎない。
- 設定項目を増やしすぎない。
- 文字化けを放置しない。
- 過去プロジェクトを丸ごと流用しない。
- Release用exeをGit管理の正本にしない。
- `ORIGINAL.md` 以外の別台帳を増やさない。
- README / DAKE_META / release_body / booth_product を真の正本として扱わない。


## DAKE_META Status Model

`DAKE_META.status` decides whether an app is allowed into the formal shipping line.

- `available`: formal shipping candidate. Release / BOOTH / dakeapp.com target.
- `draft`: work in progress. Excluded from formal shipping.
- `experimental`: exploratory app. Excluded from formal shipping. `show_in_launcher` and `show_on_site` should be false.
- `frozen`: preserved but inactive. Excluded from formal shipping. Existing assets are not treated as missing.
- `private`: personal or non-public app. Excluded from formal shipping.
- `internal`: operations or management tool. Excluded from normal distribution, BOOTH, dakeapp.com, and Launcher listing.

Formal shipping checks, BOOTH ready generation, and dakeapp.com publish candidates target `available` apps only.

## DAKE_META Role And Completion Goal

`status` describes the current state. `app_type` describes the app role. `completion_goal` describes the correct definition of done.

Add these fields to `DAKE_META`:

```json
{
  "app_type": "market",
  "completion_goal": "formal_release"
}
```

`app_type` values:

- `market`: public DAKE app for GitHub Release / BOOTH / dakeapp.com.
- `personal`: Yukiz/local-purpose app. Public release is optional, but the primary goal is local use or reference.
- `qpcs`: QPCS / BRAINZ / OIKAWA / Dashboard / Wake / operations app. Not treated as a standalone market product.
- `frozen`: frozen app preserved as history or experiment.
- `archived`: archived or legacy app, excluded from normal unresolved lanes.

`completion_goal` values:

- `formal_release`: README, release body, screenshots, BOOTH assets, build, dist exe, GitHub Release, BOOTH, dakeapp.com, and Cloudflare confirmation.
- `local_ready`: README, launch-check, build, local operation notes, and safe config handling.
- `system_ready`: README, launch-check, build, QPCS role, source/read target notes, and dashboard relationship are documented.
- `reference_ready`: README explains the reference/sample/local-document role and public-use cautions.
- `frozen_closed`: README explains frozen reason, current state, reopen conditions, and exclusion from shipping.

Dashboards should use `completion_goal` to choose the correct completion lane. Do not place `qpcs`, `personal`, `frozen`, or `archived` apps in the same unresolved lane as `market` apps waiting for `formal_release`.
