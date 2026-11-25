# PEAKHEADZ Phase3 Migration Report

作成日: 2026-06-02
実施日時: 2026-06-02 13:24:14 +09:00

## 実施内容

Phase3実コピーを実施した。

実施したのはコピーのみ。元ファイルの削除、移動、既存ファイル上書き、DAKE_Note_Inbox作成、BRAINZ/OIKAWA追加変更、remote追加は実施していない。

## 作成フォルダ

- `C:\Users\yukiz\Documents\PEAKHEADZ_ROOT\INBOX`
- `C:\Users\yukiz\Documents\PEAKHEADZ_ROOT\NOTES`
- `C:\Users\yukiz\Documents\PEAKHEADZ_ROOT\ARTICLES`
- `C:\Users\yukiz\Documents\PEAKHEADZ_ROOT\SYSTEM`

## コピー結果

素材コピー件数: 163件

| 移行先 | コピー件数 | サイズ |
| --- | ---: | ---: |
| `INBOX` | 46 | 82,513 bytes |
| `NOTES` | 44 | 88,080 bytes |
| `ARTICLES` | 51 | 87,985 bytes |
| `SYSTEM` | 22 | 190,204 bytes |

`SYSTEM` は素材22件に加えて、ログ3件を作成したため、最終ファイル数は25件。

## 作成ログ

- `C:\Users\yukiz\Documents\PEAKHEADZ_ROOT\SYSTEM\migration_log.md`
- `C:\Users\yukiz\Documents\PEAKHEADZ_ROOT\SYSTEM\duplicate_skip_log.md`
- `C:\Users\yukiz\Documents\PEAKHEADZ_ROOT\SYSTEM\excluded_files_log.md`

## 除外結果

除外件数: 61件

| 理由 | 件数 |
| --- | ---: |
| BOOTH material candidate | 1 |
| `booth_packs` など除外パス | 60 |

明示的に移行対象外:

- `brainz_memory\chatgpt`
- embedding / sqlite / vector / search index
- Remote Queue
- 通知
- 承認
- 遠隔操作
- BOOTH packs

## 重複除外

重複除外件数: 14件

| 理由 | 件数 |
| --- | ---: |
| same source_path dedupe | 9 |
| same sha256 in `40_borinef\20_sources`; one kept | 5 |

`10_slack` と `10_slack\brainz-aru` の同一source_pathは1回だけコピーした。

`40_borinef\20_sources` の同一hash重複5組は、各hashごとに1件のみコピーし、除外内容を `duplicate_skip_log.md` に記録した。

## 衝突対応

- filename衝突suffix対応件数: 0
- 既存ファイル上書き: 0
- エラー件数: 0

## Secret / Token確認

コピー前:

- Slack token実値検出: 0
- Slack webhook実値検出: 0
- API key / secret / access tokenらしき実値検出: 0

コピー後:

- `INBOX` / `NOTES` / `ARTICLES` / `SYSTEM` 内でsecret/token/webhook実値検出なし。
- ログ3種にもsecret/token/webhook実値は記載していない。

## UTF-8確認

- コピー後のtext対象でUTF-8 decodeリスクなし。
- included内の非text候補なし。

## 元ファイル確認

- コピー済みファイルの元source pathは全件存在確認済み。
- 元ファイル削除なし。
- 元ファイル移動なし。

## Git確認

`C:\Users\yukiz\Documents\PEAKHEADZ_ROOT` は `.git` が存在しないため、移行されたMarkdownとSYSTEMログのcommitは実施していない。

`C:\Users\yukiz\Documents\New project` はremote未設定のため、レポートcommit後もpushは行わない。

## Phase4へ進めるか

Phase3のコピー移行は完了。4フォルダ、素材コピー、ログ3種、検証が揃ったため、Phase4の `DAKE_Note_Inbox` 新規作成へ進める状態。

ただし、Phase4開始はユーザーの明示指示後に行う。
