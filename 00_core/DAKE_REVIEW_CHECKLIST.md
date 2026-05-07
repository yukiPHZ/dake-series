# DAKE_REVIEW_CHECKLIST

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
