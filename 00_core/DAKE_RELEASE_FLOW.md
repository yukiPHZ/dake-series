# DAKE_RELEASE_FLOW

## DAKE正式出荷ライン

DAKEの正式出荷ラインは以下で固定します。

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
以下が揃って初めて正式出荷とします。

- `README.md` が最新
- `release_body.md` が最新
- `assets/screenshot.webp` が存在する
- `assets/booth_thumbnail.jpg` が存在する
- `booth_product.txt` が存在する
- `booth_ready/` が生成されている
- `dist/*.exe` が生成されている
- GitHub Release が作成され、exeが添付されている
- README の `DAKE_META.release_url` が更新済み
- BOOTH掲載準備が完了している
- dakeapp.com に反映されている
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
- dakeapp.com掲載に必要な `display_name`、`site_title`、`site_description`、`update_summary`、`release_url`、`screenshot_path` を確認する。
- dakeapp.com反映後、Cloudflare上の本番URLで200、スクショ、Releaseリンクを確認する。

BOOTH登録最適化の正本は `DAKE_BOOTH_REGISTER_SPEC.md` です。
DakeBOOTHアシストは登録入力補助に使えますが、公開ボタンは押さず、最終公開は人間が確認します。

DAKEアプリ完成後の公開手順です。

## 前提

- GitHub Release本体の編集は、明示された公開作業時だけ行う。
- Release説明文の正本は各アプリREADMEの `RELEASE_BODY`。
- `release_body.md` はREADMEから生成する貼り付け用ファイル。
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

各アプリREADMEに以下を置く:

- `DAKE_META`
- `RELEASE_BODY`
- 必要な使い方、仕様、注意事項

`DAKE_META.status`:

- `available`: exe、README、screenshotが揃っている。
- `missing_exe`: dist内exeがない。
- `build_failed`: buildまたは起動、スクショ取得に失敗した。

## 3. release_body.md生成

- READMEの `RELEASE_BODY` を読み取る。
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
- `update_summary` に公開内容を短く反映する。
- `status` は問題なければ `available` のままにする。

## 10. dakeapp.com反映

- READMEの `DAKE_META` を読み込む。
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
