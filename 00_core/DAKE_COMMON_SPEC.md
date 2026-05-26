# DAKE_COMMON_SPEC

## DAKE正式出荷ライン

DAKEシリーズの正式出荷は、GitHub Release公開だけでは完了としません。
README正本運用、BOOTH販売素材、dakeapp.com掲載までを含めて、初めて正式出荷とします。

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

## Releaseと正式出荷の定義

- Release = GitHub上の配布物公開。
- 出荷 = ユーザーが触れる状態。
- GitHub Release公開のみでは、DAKEの正式出荷完了とは扱わない。
- ユーザーが一覧で見つけ、スクショで理解し、BOOTHまたはdakeapp.comから迷わず取得できる状態までを正式出荷とする。

正式出荷に必要なもの:

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
- Cloudflare反映確認

Releaseだけ、exeだけ、READMEだけの状態は正式出荷完了とは扱いません。

DAKEシリーズの最上位共通仕様です。個別仕様や実装判断で迷った場合は、このファイルを優先します。

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
- READMEの `DAKE_META.exe_name` と `dist` 内の正本exe名を一致させる。

## README正本

各アプリの `README.md` は、そのアプリ自身が語る正本です。

必須ブロック:

- `DAKE_META`: ランチャー、dakeapp.com、Release連携用のJSON。
- `RELEASE_BODY`: GitHub Release説明欄に貼る3〜5行の短文。

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
- README以外の別台帳を増やさない。
