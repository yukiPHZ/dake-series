# DAKE_GIT_RULE

DAKEシリーズのGit運用と除外ルールです。

## 基本方針

- Gitで管理する主要素材は、ソース、仕様、`ORIGINAL.md`、README公開ビュー、スクリーンショット。
- アプリの真の正本は `ORIGINAL.md`。未導入の既存アプリではREADMEを暫定参照する。
- exeはGitに入れず、GitHub Releaseで配布する。
- build成果物は再生成できるものとして扱う。
- 個人設定やローカル環境依存ファイルを入れない。

## 除外するもの

`.gitignore` に含める基本:

```gitignore
# Python
__pycache__/
*.pyc

# Build
build/
dist/
*.spec

# Executables
*.exe

# Local config
*_config.json
*.local.json

# OS
Thumbs.db
desktop.ini
```

## Gitに入れるもの

- `main.py`
- `build.bat`
- `requirements.txt` 必要な場合
- `README.md`
- `release_body.md`
- `.gitignore`
- `assets/screenshot.webp`
- `00_core/*.md`
- `02_assets/dake_icon.ico`

## Release配布物

- `dist/<exe_name>.exe` はGitHub Releaseへ添付する。
- Release説明文は `release_body.md` を使う。
- Release作成後、READMEの `DAKE_META.release_url` を更新する。`ORIGINAL.md` 導入済みアプリでは、正本へ戻すべき公開情報か確認する。

## commit前確認

- 不要な `dist/` や `build/` が含まれていない。
- `*.exe` が含まれていない。
- 個人情報や環境依存パスが入っていない。
- READMEの `DAKE_META` が壊れていない。
- `release_body.md` が `ORIGINAL.md` 由来、または移行前READMEの `RELEASE_BODY` と一致している。
- `assets/screenshot.webp` が最新。

## ブランチと作業

- 横断作業は変更範囲が広くなるため、commit前に対象ファイル一覧を確認する。
- 既存のユーザー変更を戻さない。
- 仕様更新とアプリ実装更新は、可能ならcommitを分ける。
