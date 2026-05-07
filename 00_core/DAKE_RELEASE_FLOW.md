# DAKE_RELEASE_FLOW

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

## 公開前の最終チェック

- Release本文が短い。
- exeが対象アプリのもの。
- スクリーンショットが最新。
- READMEの `exe_name` と添付exeが一致。
- `release_url` はRelease作成後に更新されている。
