# DAKE 00_core

## SHIP / DISTRIBUTION / COMMERCE

- SHIP: build、生成exeの主要動作、必要なGit作業、GitHub Release。GitHub Release完了で `SHIP = COMPLETE`。
- DISTRIBUTION: BOOTH、dakeapp.com、store.dakeapp.com。`COMPLETE / PARTIAL / PENDING`。
- COMMERCE: Stripe、Payment Link、`payment_status`、checkout。`READY / PENDING / NOT_REQUIRED`。

BOOTH、Store、Stripe、Cloudflare、SNSはSHIP完了条件ではありません。詳細は `DAKE_RELEASE_FLOW.md` を参照します。

DAKEシリーズ全体で参照する共通仕様ファイル置き場です。

今後の正本はこのフォルダ内のMarkdownファイルです。既存の `design_rules.txt`、`dev_rules.txt`、`philosophy.txt` は削除せず残しますが、参照・更新はMarkdown側を優先します。

## 仕様ファイル一覧

| ファイル | 役割 |
| --- | --- |
| `DAKE_SOURCE_OF_TRUTH_ENGINE_V02.md` | `ORIGINAL.md` を中心に、変換・利用・販売・運用までを含めた正本主義エンジン v0.2 の上位思想です。 |
| `PROJECT_INFORMATION_SOURCE.md` | DAKEプロジェクトの現在の運用状態とLegacy境界を確認します。 |
| `DAKE_ORIGINAL_RULE.md` | `ORIGINAL.md` を真の正本とし、README / DAKE_META / release_body / booth_product / Store表示を派生ビューとして扱うルールです。 |
| `DAKE_ORIGINAL_TEMPLATE_APP.md` | DAKEアプリへ `ORIGINAL.md` を導入するためのテンプレートです。 |
| `DAKE_COMMON_SPEC.md` | DAKEシリーズの最上位共通仕様。迷ったら最初に読む正本です。 |
| `DAKE_UI_GUIDE.md` | 見た目、操作感、初期表示、余白、ボタン、フッターの基準です。 |
| `DAKE_RELEASE_FLOW.md` | SHIP / DISTRIBUTION / COMMERCEの独立した手順です。 |
| `DAKE_SCREENSHOT_RULE.md` | `assets/screenshot.webp` と `assets/booth_thumbnail.jpg` の作成・品質ルールです。 |
| `DAKE_BOOTH_REGISTER_SPEC.md` | BOOTH登録最適化、DakeBOOTHアシスト、自動化境界、booth_ready完成定義です。 |
| `DAKE_REVIEW_CHECKLIST.md` | 変更したレイヤーだけに使う短い確認表です。 |
| `DAKE_BUILD_RULE.md` | PyInstaller、`build.bat`、exe名、hidden-importのルールです。 |
| `DAKE_APP_TEMPLATE.md` | 新規アプリ作成時の基本フォルダ構成と最小テンプレです。 |
| `DAKE_UI_TEXT_RULE.md` | `APP_NAME`、`WINDOW_TITLE`、`COPYRIGHT`、`UI_TEXT` の運用ルールです。 |
| `DAKE_GIT_RULE.md` | Git管理、除外、Release配布物の扱いです。 |
| `DAKE_ICON_RULE.md` | `02_assets/dake_icon.ico` を使う統一アイコンルールです。 |
| `DAKE_FORBIDDEN_RULE.md` | 多機能化、過剰UI、文字化け放置などの禁止事項です。 |

## 読む順番

通常は `AGENTS.md`、`DAKE_COMMON_SPEC.md`、対象ORIGINALの3ファイルだけを読む。その他はAGENTSの条件付きルートに従い、対象作業が発生した時だけ参照する。

## 正本の考え方

- シリーズ共通仕様の正本: `00_core/*.md`
- アプリごとの真の正本: `01_apps/<app>/ORIGINAL.md`
- ORIGINAL.md未導入の既存アプリ: `01_apps/<app>/README.md` を暫定参照する
- `README.md`: GitHub公開用ビューであり、真の正本ではない
- `DAKE_META`: Launcher / サイト / 管理ツール向けの機械利用ビュー
- `release_body.md`: GitHub Release用ビュー
- `booth_product.txt`: BOOTH登録用ビュー
- スクリーンショット正本: `01_apps/<app>/assets/screenshot.webp`
- BOOTH商品画像: `01_apps/<app>/assets/booth_thumbnail.jpg`
- Release説明文の元情報: `ORIGINAL.md`。未導入アプリでは各アプリREADMEの `RELEASE_BODY` を暫定参照する

## BOOTH登録

- BOOTH登録はDISTRIBUTION作業です。SHIPとは独立します。
- Codexは `booth_product.txt`、`booth_ready/`、`assets/booth_thumbnail.jpg` 等の登録素材を準備し、`BOOTH_PREP = READY` で停止します。
- DakeBOOTHアシストは人間用の登録補助アプリです。Codexは自動起動・自動操作しません。
- BOOTHサイトの入力、画像・ZIPアップロード、登録、更新、公開、販売開始は人間が行います。
- 詳細は `DAKE_BOOTH_REGISTER_SPEC.md` を参照します。

## CLI標準化

| ファイル | 役割 |
| --- | --- |
| `DAKE_CLI_SPEC.md` | しまりすくんCLI対応の共通仕様。`--from-shimarisu`、引数、exit code、stderr、保存ルールを定義します。 |
| `DAKE_CLI_TEMPLATE.md` | 新規アプリへCLI入口を追加するためのargparse/GUI分岐テンプレートです。 |
| `SHIMARISU_DAKE_FLOW.md` | しまりすくんとDAKEの責務分担、呼び出し構造、実務パイプラインの考え方です。 |


## DAKE Status Model v3

`DAKE_META.status` は既存の公開候補フィルタとして残します。

- `available`: public candidate. SHIP / DISTRIBUTIONの対象になり得る。
- `draft`: work in progress.
- `experimental`: exploratory app. `show_in_launcher` and `show_on_site` should be false.
- `frozen`: preserved but inactive.
- `private`: personal or non-public app.
- `internal`: operations or management tool. Excluded from normal distribution, BOOTH, dakeapp.com, and Launcher listing.

`available` であっても、依頼されていないDISTRIBUTION / COMMERCEへ自動的に進めない。

## DAKE_META Role Model

DAKE app README files also carry role and completion metadata:

- `app_type`: role classification such as `market`, `qpcs`, `personal`, `frozen`, or `archived`.
- `completion_goal`: Legacy互換値を含む完了目標。既存の `formal_release` は新しいSHIP定義を上書きしない。

Dashboards should read these fields without assuming that every app must reach BOOTH / dakeapp.com. SHIP / DISTRIBUTION / COMMERCEは別状態として扱う。
