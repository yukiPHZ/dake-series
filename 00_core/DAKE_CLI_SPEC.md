# DAKE_CLI_SPEC

DAKEアプリを、しまりすくんCLIから安全に呼び出すための共通仕様です。
通常利用者はGUIを使い、CLIは自動処理・実務パイプライン用の入口として最小限だけ持ちます。

## 基本方針

- DAKEは処理を担当し、判断はしまりすくん側に置く。
- GUI通常起動とCLI実行を明確に分離する。
- 1アプリ1機能を崩さない。CLI追加を理由に多機能化しない。
- UIを止めない。CLI処理はGUIのイベントループに依存しない。
- 失敗時は短く、復旧しやすい言葉で返す。
- Python Tracebackを利用者やしまりすくんへ直接見せない。

## CLI起動条件

CLIとして動作する条件は、必ず `--from-shimarisu` が渡された場合のみです。

```text
DakeApp.exe --from-shimarisu --input input.pdf --output output.pdf
```

- `--from-shimarisu` なし: GUI通常起動。
- `--from-shimarisu` あり: GUIを出さずCLI処理。
- CLI中は `messagebox`、ファイルダイアログ、ブラウザ起動、Explorer起動をしない。
- CLI中は標準入出力とexit codeだけで結果を返す。

## 必須引数

全アプリ共通で必須:

- `--from-shimarisu`: CLI実行を明示するフラグ。

ファイル処理系アプリで原則必須:

- `--input`: 入力ファイル、または入力フォルダ。
- `--output`: 出力ファイル。複数出力の場合は `--output-dir` を使う。

複数入力が自然なアプリ:

- `--inputs`: 複数ファイルを受け取る。区切り仕様はアプリREADMEに明記する。
- `--input-list`: 1行1パスのテキストファイル。大量入力はこちらを優先する。

アプリ固有で必要な引数:

- 例: `--pages`, `--name`, `--size`, `--format`。
- 追加は必要最小限にする。
- 迷うオプションはGUI側に残し、CLIでは固定値にする。

## optional引数

共通で使ってよい任意引数:

- `--silent`: 成功時の余計な表示を抑える。
- `--job-id`: しまりすくん側の処理ID。ログやstdout JSONにそのまま返してよい。
- `--output-dir`: 複数ファイル出力時の保存先。
- `--format`: 出力形式を選ぶ必要がある場合のみ。
- `--dry-run`: 実処理せず入力検証だけ行う場合のみ。

使わない方針の引数:

- `--overwrite`: 原則禁止。
- `--force`: 原則禁止。
- `--interactive`: CLIでは禁止。

## exit code

| code | 意味 | 例 |
| --- | --- | --- |
| `0` | 成功 | 出力完了 |
| `1` | 引数エラー | 必須引数不足、形式不正 |
| `2` | 入力エラー | 入力ファイルなし、読めない |
| `3` | 出力先エラー | 既存ファイルあり、保存先に書けない |
| `4` | 処理失敗 | PDF変換失敗、画像処理失敗 |
| `5` | 未対応 | 形式未対応、CLI未対応機能 |
| `10` | 予期しない失敗 | 捕捉済みの想定外エラー |

終了時は必ず `sys.exit(code)` または `raise SystemExit(code)` で返します。

## stdout方針

成功時は、しまりすくんが読める短いJSONをstdoutへ出します。

```json
{
  "ok": true,
  "app_key": "DAKE_Category_Function",
  "output": "C:/path/to/output.pdf"
}
```

- `--silent` がない場合でも、stdoutは構造化された1つのJSONを基本にする。
- 進捗表示をstdoutへ流さない。
- 利用者向けの長文説明をstdoutへ出さない。

## stderr方針

失敗時はstderrへ短く出します。

```text
[DAKE_ERROR] code=2 reason=input_not_found message=入力ファイルが見つかりません。
```

- stderrは1行目だけで原因が分かるようにする。
- Tracebackは出さない。
- 内部パスやライブラリ例外をそのまま長く出さない。
- 復旧の手がかりを短く含める。

## Traceback禁止

CLIでは、未捕捉例外をそのまま落とさない。

- `try/except` で捕捉する。
- 予期しないエラーは exit code `10`。
- 詳細ログが必要な場合でも、標準では出さない。
- デバッグ用ログを残す場合は、READMEに保存場所を明記する。

## output保存ルール

- 入力ファイルを上書きしない。
- `--output` が存在する場合は exit code `3` で止める。
- `--output-dir` が存在しない場合は作成してよい。ただし作成失敗は exit code `3`。
- 出力名はアプリ側で勝手に大きく変えない。
- 複数出力の場合は、連番や元ファイル名ベースで衝突しない名前にする。
- 一時ファイルは処理終了時に削除する。

## silentモード

`--silent` 指定時:

- 成功時stdoutは必要最小限のJSONだけ。
- stderrは失敗時のみ。
- GUI表示、完了通知、確認ダイアログは出さない。
- 進捗ログを標準出力へ流さない。

`--silent` がなくても、CLIではGUI通知を出しません。

## overwrite禁止

DAKE CLIでは、既存ファイルの上書きを標準機能にしません。

- `--overwrite` は追加しない。
- 既存ファイルがある場合は止める。
- しまりすくん側で別名を判断して再実行する。
- DAKEは勝手に削除・移動・置換しない。

## 命名ルール

- CLI関数: `cli_main(argv=None)`。
- GUI関数: `gui_main()` または既存の `main()`。
- 分岐関数: `main(argv=None)`。
- argparse生成: `build_parser()`。
- CLI処理本体: `run_cli(args)`。
- CLIフラグ: `--from-shimarisu`。
- 引数名は小文字kebab-case。
- READMEにはCLI対応時のみ、短い使用例を追加する。

## READMEへの記載

CLI対応アプリは、READMEに以下を追加します。

````markdown
## CLI

```bat
DakeApp.exe --from-shimarisu --input input.pdf --output output.pdf
```

- GUI通常起動とは分離しています。
- 既存ファイルは上書きしません。
- 失敗時はstderrとexit codeで返します。
````

## 実装判断

- CLI追加でGUIを複雑にしない。
- GUI用関数を無理にCLIへ流用しない。
- 処理本体だけを共有し、入口とエラー表示は分ける。
- 迷ったら `DAKE_COMMON_SPEC.md` を優先する。