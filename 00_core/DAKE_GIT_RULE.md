# DAKE_GIT_RULE

Git除外、成果物、履歴の扱いを変える時だけ参照します。

## 基本

- 個別アプリの正本は `ORIGINAL.md`。README等は派生ビュー。
- ソース、仕様、必要な公開ビューをGitで管理する。
- exeは通常commitせず、GitHub Release成果物として扱う。
- 既存のユーザー変更・未追跡ファイルを戻す、削除する、stashする、混在commitすることは禁止。
- 対象ファイルだけをstageし、`git add .` を安易に使わない。

## 通常の除外

```gitignore
__pycache__/
*.pyc
build/
dist/
*.spec
*.exe
*_config.json
*.local.json
Thumbs.db
desktop.ini
```

## commit前

- 対象外の差分が含まれていない。
- build / dist / exe / ローカル設定 / 秘密情報が含まれていない。
- 変更が影響する派生ビューだけを更新している。
- スクリーンショットはUIの公開表示が変わった場合だけ更新している。

## Releaseと履歴

- GitHub ReleaseはSHIP成果物の公開先。
- 過去Releaseは原則保持し、問題時は直前Releaseへ戻す。
- 専用バックアップ工程を追加しない。
- commit、push、GitHub Releaseは別権限として扱う。
