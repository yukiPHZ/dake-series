# DAKE_REVIEW_CHECKLIST

## DAKE正式出荷ライン

- [ ] `ORIGINAL.md` がある場合は真の正本として整っている。
- [ ] `ORIGINAL.md` 未導入の既存アプリでは `README.md` を暫定参照できる。
- [ ] `README.md` はGitHub公開用ビューとして整っている。
- [ ] `release_body.md` が `ORIGINAL.md` 由来、または移行前READMEの `RELEASE_BODY` と一致している。
- [ ] `assets/screenshot.webp` がある。
- [ ] `assets/booth_thumbnail.jpg` がある。
- [ ] `booth_ready/booth_product.txt` がある。
- [ ] `booth_ready/` がある。
- [ ] buildが成功している。
- [ ] `dist/*.exe` が生成されている。
- [ ] GitHub Releaseが作成され、exeが添付されている。
- [ ] Releaseにexeが添付済み。
- [ ] READMEの `DAKE_META.release_url` が更新されている。
- [ ] BOOTH ready素材が揃っている。
- [ ] BOOTH掲載準備が完了している。
- [ ] BOOTH公開後のURL欄が保持されている。
- [ ] dakeapp.com掲載に必要な項目が `ORIGINAL.md` 由来、または移行前README / DAKE_METAから読める。
- [ ] dakeapp.comに表示されている。
- [ ] Cloudflare反映確認済み。
- [ ] GitHub Releaseのみで出荷完了と扱っていない。

GitHub Release公開のみでは正式出荷完了とはしません。
BOOTH ready、BOOTH、dakeapp.com、Cloudflare反映確認まで含めて確認します。

## BOOTH素材

- [ ] `assets/booth_thumbnail.jpg` がある。
- [ ] `booth_ready/booth_thumbnail.jpg` がある。
- [ ] `booth_ready/screenshot.jpg` がある。
- [ ] `booth_ready/README.txt` がある。
- [ ] `booth_ready/注意事項.txt` がある。
- [ ] `booth_ready/booth_product.txt` がある。
- [ ] `booth_ready/booth_product.txt` に `# URL` 欄がある。
- [ ] BOOTH用zipはGit管理対象に含めていない。

DAKEアプリ横断レビュー用チェックリストです。

## 基本

- [ ] アプリの目的が1つに絞られている。
- [ ] 起動直後に用途が分かる。
- [ ] 主操作が迷わず見つかる。
- [ ] 処理中・完了・失敗が分かる。
- [ ] 元ファイルを不用意に上書きしない。

## UTF-8と文字化け

- [ ] `main.py`、README、release_bodyがUTF-8で保存されている。
- [ ] 日本語が `????` や `�` になっていない。
- [ ] コンソールやPowerShell経由の生成で文字化けしていない。

## UI_TEXT

- [ ] `APP_NAME` がある。
- [ ] `WINDOW_TITLE` がある。
- [ ] `COPYRIGHT` がある。
- [ ] 画面文言が `UI_TEXT` に集約されている。
- [ ] ボタン、ラベル、ステータス、エラー文が直書きされていない。

## UI

- [ ] フォントがWindows日本語表示に向いている。
- [ ] 色がDAKEらしく控えめ。
- [ ] 主ボタンが目立つ。
- [ ] 補助ボタンが主ボタンより目立たない。
- [ ] 入力欄とラベルの関係が分かる。
- [ ] 狭い画面で崩れない。
- [ ] スクロールが必要な画面で操作不能にならない。

## フッター

- [ ] DAKE共通コピーがある。
- [ ] `戸建買取査定` と `Instagram` リンクが必要に応じてある。
- [ ] コピーライトがある。
- [ ] フッターが主操作を邪魔していない。
- [ ] 狭い幅でも見切れない。

## アイコン

- [ ] `02_assets/dake_icon.ico` を参照している。
- [ ] `build.bat` の `--icon` が共通アイコンを指している。
- [ ] 個別アイコンを作っていない。

## build

- [ ] `build.bat` がある。
- [ ] build前に `build/`、`dist/`、`*.spec` を整理する。
- [ ] PyInstallerの `--onefile` を使う。
- [ ] GUIアプリは `--noconsole` を使う。
- [ ] exe名がREADMEの `DAKE_META.exe_name` と一致する。
- [ ] hidden-importやcollect-dataが必要なライブラリを指定している。

## README

- [ ] `DAKE_META` がある。
- [ ] `DAKE_META` がJSONとして読める。
- [ ] `DAKE_META.app_type` がある。
- [ ] `DAKE_META.completion_goal` がある。
- [ ] `status` / `app_type` / `completion_goal` の組み合わせが実態と一致している。
- [ ] `formal_release` 以外のアプリを市場向け未完了レーンに入れていない。
- [ ] `folder_name` が実フォルダ名と一致する。
- [ ] `exe_name` がdist内exeと一致する。
- [ ] `release_url` は未確定なら空文字。
- [ ] `screenshot_path` は `assets/screenshot.webp`。
- [ ] `status` が現状と一致する。

## Release本文

- [ ] READMEに `RELEASE_BODY` がある。
- [ ] 3〜5行の箇条書きになっている。
- [ ] 長すぎない。
- [ ] `release_body.md` と一致している。

## スクリーンショット

- [ ] `assets/screenshot.webp` がある。
- [ ] 起動直後のアプリウインドウだけを撮っている。
- [ ] WebP形式。
- [ ] 横幅1200px以内。
- [ ] 引き延ばしていない。

## Git除外

- [ ] `build/` は除外。
- [ ] `dist/` は除外。
- [ ] `*.spec` は除外。
- [ ] `*.exe` はGit管理しない。
- [ ] `__pycache__/` と `*.pyc` は除外。
- [ ] 個人設定ファイルは除外。

## 起動確認

- [ ] `python main.py` またはexeで起動できる。
- [ ] 初期画面が表示される。
- [ ] 閉じられる。
- [ ] 主要操作で例外が出ない。

## 正式出荷ライン v2

- [ ] 通常出荷候補は `status: available` のみに限定している。
- [ ] `frozen` / `draft` / `experimental` / `private` は通常不足扱いにしていない。
- [ ] `available` 以外で `show_in_launcher: true` になっていない。
- [ ] `available` 以外で `show_on_site: true` になっていない。
- [ ] BOOTH公開後、`booth_ready/booth_product.txt` の `# URL` にBOOTH URLを戻している。
- [ ] BOOTH URL未記入のままCLOSED扱いしていない。
- [ ] dakeapp.com掲載とCloudflare反映確認まで終えている。


## Formal Shipping Line v3

- [ ] Normal shipping candidates are limited to `status: available`.
- [ ] `frozen` / `draft` / `experimental` / `private` / `internal` are not counted as normal missing assets.
- [ ] `internal` is treated as an operations or management tool, excluded from Release / BOOTH / dakeapp.com / Launcher listing.
- [ ] Non-available apps do not have `show_in_launcher: true`.
- [ ] Non-available apps do not have `show_on_site: true`.
- [ ] Apps are not treated as CLOSED while the BOOTH URL is missing.
- [ ] Available apps have actual `dist/*.exe` launch confirmation.
- [ ] Prefer `dist/*.exe --launch-check` for launch confirmation.
- [ ] If `--launch-check` is not implemented, perform only a short GUI smoke launch.
- [ ] New apps implement `--launch-check` without file conversion, sending, publishing, browser automation, or external side effects.
