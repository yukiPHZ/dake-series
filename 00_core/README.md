# DAKE 00_core

## DAKE正式出荷ライン

DAKEの正式出荷は、GitHub Release公開だけでは完了としません。
README正本、スクリーンショット、BOOTH販売素材、dakeapp.com掲載までを1本の出荷ラインとして扱います。

```text
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
Cloudflare反映確認
↓
正式出荷完了
```

GitHub Release公開のみでは正式出荷完了とはしません。
DAKEでは、ユーザーが一覧で見つけ、スクショで理解し、BOOTHまたはdakeapp.comから迷わず取得できる状態までを正式出荷とします。

詳細は `DAKE_RELEASE_FLOW.md`、画像ルールは `DAKE_SCREENSHOT_RULE.md`、棚卸しは `DAKE_REVIEW_CHECKLIST.md` を参照します。

DAKEシリーズ全体で参照する共通仕様ファイル置き場です。

今後の正本はこのフォルダ内のMarkdownファイルです。既存の `design_rules.txt`、`dev_rules.txt`、`philosophy.txt` は削除せず残しますが、参照・更新はMarkdown側を優先します。

## 仕様ファイル一覧

| ファイル | 役割 |
| --- | --- |
| `DAKE_COMMON_SPEC.md` | DAKEシリーズの最上位共通仕様。迷ったら最初に読む正本です。 |
| `DAKE_UI_GUIDE.md` | 見た目、操作感、初期表示、余白、ボタン、フッターの基準です。 |
| `DAKE_RELEASE_FLOW.md` | buildからGitHub Release、BOOTH、dakeapp.com、Cloudflare反映確認までの公開手順です。 |
| `DAKE_SCREENSHOT_RULE.md` | `assets/screenshot.webp` と `assets/booth_thumbnail.jpg` の作成・品質ルールです。 |
| `DAKE_BOOTH_REGISTER_SPEC.md` | BOOTH登録最適化、DakeBOOTHアシスト、自動化境界、booth_ready完成定義です。 |
| `DAKE_REVIEW_CHECKLIST.md` | 横断レビュー時に見るチェックリストです。 |
| `DAKE_BUILD_RULE.md` | PyInstaller、`build.bat`、exe名、hidden-importのルールです。 |
| `DAKE_APP_TEMPLATE.md` | 新規アプリ作成時の基本フォルダ構成と最小テンプレです。 |
| `DAKE_UI_TEXT_RULE.md` | `APP_NAME`、`WINDOW_TITLE`、`COPYRIGHT`、`UI_TEXT` の運用ルールです。 |
| `DAKE_GIT_RULE.md` | Git管理、除外、Release配布物の扱いです。 |
| `DAKE_ICON_RULE.md` | `02_assets/dake_icon.ico` を使う統一アイコンルールです。 |
| `DAKE_FORBIDDEN_RULE.md` | 多機能化、過剰UI、文字化け放置などの禁止事項です。 |

## 読む順番

1. 新規アプリを作る: `DAKE_COMMON_SPEC.md` → `DAKE_APP_TEMPLATE.md` → `DAKE_UI_GUIDE.md` → `DAKE_BUILD_RULE.md`
2. 既存アプリをレビューする: `DAKE_REVIEW_CHECKLIST.md` → `DAKE_UI_TEXT_RULE.md` → `DAKE_SCREENSHOT_RULE.md`
3. 公開作業をする: `DAKE_RELEASE_FLOW.md` → `DAKE_GIT_RULE.md` → 各アプリの `README.md`
4. 判断に迷ったとき: `DAKE_COMMON_SPEC.md` を優先する

## 正本の考え方

- シリーズ共通仕様の正本: `00_core/*.md`
- アプリごとの正本: `01_apps/<app>/README.md`
- スクリーンショット正本: `01_apps/<app>/assets/screenshot.webp`
- BOOTH商品画像: `01_apps/<app>/assets/booth_thumbnail.jpg`
- Release説明文の生成元: 各アプリREADMEの `RELEASE_BODY`

## BOOTH登録

- BOOTH登録はDAKE正式出荷ラインの一部です。
- `booth_product.txt`、`booth_ready/`、`assets/booth_thumbnail.jpg` は必須素材として扱います。
- DakeBOOTHアシストは登録補助用です。
- Playwrightは入力補助に使用できますが、最終公開と販売開始は人間が確認して行います。
- 詳細は `DAKE_BOOTH_REGISTER_SPEC.md` を参照します。

## CLI標準化

| ファイル | 役割 |
| --- | --- |
| `DAKE_CLI_SPEC.md` | しまりすくんCLI対応の共通仕様。`--from-shimarisu`、引数、exit code、stderr、保存ルールを定義します。 |
| `DAKE_CLI_TEMPLATE.md` | 新規アプリへCLI入口を追加するためのargparse/GUI分岐テンプレートです。 |
| `SHIMARISU_DAKE_FLOW.md` | しまりすくんとDAKEの責務分担、呼び出し構造、実務パイプラインの考え方です。 |


## DAKE Status Model v3

`DAKE_META.status` controls whether an app enters the formal shipping line.

- `available`: formal shipping candidate. Release / BOOTH / dakeapp.com target.
- `draft`: work in progress. Excluded from formal shipping.
- `experimental`: exploratory app. Excluded from formal shipping. `show_in_launcher` and `show_on_site` should be false.
- `frozen`: preserved but inactive. Excluded from formal shipping. Existing assets are not treated as missing.
- `private`: personal or non-public app. Excluded from formal shipping.
- `internal`: operations or management tool. Excluded from normal distribution, BOOTH, dakeapp.com, and Launcher listing.

Only `available` apps are normal formal shipping check targets.
