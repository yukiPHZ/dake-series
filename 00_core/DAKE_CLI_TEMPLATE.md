# DAKE_CLI_TEMPLATE

DAKEアプリへ、しまりすくんCLI対応を追加するときの最小テンプレートです。
既存GUIを壊さず、入口だけを分けます。

## 構成の考え方

- `main()` はGUI/CLIの分岐だけを担当する。
- `gui_main()` は通常のTkinter起動を担当する。
- `cli_main()` はargparse、検証、exit code変換を担当する。
- `run_cli(args)` は処理本体を呼ぶ。
- GUIのmessageboxとCLIのstderrは混ぜない。

## argparseテンプレ

```python
import argparse
import json
import sys
from pathlib import Path

APP_KEY = "DAKE_Category_Function"
CLI_FLAG = "--from-shimarisu"

EXIT_OK = 0
EXIT_USAGE = 1
EXIT_INPUT = 2
EXIT_OUTPUT = 3
EXIT_PROCESS = 4
EXIT_UNSUPPORTED = 5
EXIT_UNEXPECTED = 10


class DakeCliError(Exception):
    def __init__(self, exit_code, reason, message):
        super().__init__(message)
        self.exit_code = exit_code
        self.reason = reason
        self.message = message


def build_parser():
    parser = argparse.ArgumentParser(
        prog=APP_KEY,
        description="DAKE app CLI for Shimarisu.",
    )
    parser.add_argument("--from-shimarisu", action="store_true", required=True)
    parser.add_argument("--input", required=True, help="Input file path.")
    parser.add_argument("--output", required=True, help="Output file path.")
    parser.add_argument("--silent", action="store_true")
    parser.add_argument("--job-id", default="")
    return parser
```

## GUI/CLI分岐テンプレ

```python
def main(argv=None):
    argv = sys.argv[1:] if argv is None else argv

    if CLI_FLAG in argv:
        return cli_main(argv)

    gui_main()
    return EXIT_OK


if __name__ == "__main__":
    raise SystemExit(main())
```

## main()構成

既存のTkinter起動が `main()` に入っている場合は、以下のように分けます。

```python
def gui_main():
    root = tk.Tk()
    app = App(root)
    root.mainloop()
```

- GUI側の表示文言は従来どおり `UI_TEXT` を使う。
- CLI側のエラー文は `DakeCliError` で管理する。
- 処理本体は `process_file(input_path, output_path)` のような純粋関数に寄せる。

## exit codeテンプレ

```python
def cli_main(argv=None):
    parser = build_parser()
    try:
        args = parser.parse_args(argv)
        result = run_cli(args)
        if not args.silent:
            pass
        print(json.dumps(result, ensure_ascii=False))
        return EXIT_OK
    except DakeCliError as exc:
        print_cli_error(exc)
        return exc.exit_code
    except SystemExit as exc:
        # argparse error. Avoid Python traceback.
        return exc.code if isinstance(exc.code, int) else EXIT_USAGE
    except Exception:
        print_cli_error(DakeCliError(
            EXIT_UNEXPECTED,
            "unexpected_error",
            "処理中に予期しないエラーが発生しました。",
        ))
        return EXIT_UNEXPECTED
```

## stderrテンプレ

```python
def print_cli_error(exc):
    print(
        f"[DAKE_ERROR] code={exc.exit_code} "
        f"reason={exc.reason} message={exc.message}",
        file=sys.stderr,
    )
```

## try/except構成

```python
def run_cli(args):
    input_path = Path(args.input)
    output_path = Path(args.output)

    if not input_path.exists():
        raise DakeCliError(EXIT_INPUT, "input_not_found", "入力ファイルが見つかりません。")

    if output_path.exists():
        raise DakeCliError(EXIT_OUTPUT, "output_exists", "出力ファイルがすでに存在します。")

    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        process_file(input_path, output_path)
    except DakeCliError:
        raise
    except Exception:
        raise DakeCliError(EXIT_PROCESS, "process_failed", "処理に失敗しました。")

    return {
        "ok": True,
        "app_key": APP_KEY,
        "job_id": args.job_id,
        "output": str(output_path),
    }
```

## sample implementation

以下は、入力テキストを別ファイルへコピーするだけの最小実装です。
実アプリでは `process_file()` の中だけを機能に合わせて置き換えます。

```python
import argparse
import json
import sys
from pathlib import Path

APP_KEY = "DAKE_Sample_Copy"
CLI_FLAG = "--from-shimarisu"

EXIT_OK = 0
EXIT_USAGE = 1
EXIT_INPUT = 2
EXIT_OUTPUT = 3
EXIT_PROCESS = 4
EXIT_UNSUPPORTED = 5
EXIT_UNEXPECTED = 10


class DakeCliError(Exception):
    def __init__(self, exit_code, reason, message):
        super().__init__(message)
        self.exit_code = exit_code
        self.reason = reason
        self.message = message


def build_parser():
    parser = argparse.ArgumentParser(prog=APP_KEY)
    parser.add_argument("--from-shimarisu", action="store_true", required=True)
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--silent", action="store_true")
    parser.add_argument("--job-id", default="")
    return parser


def print_cli_error(exc):
    print(
        f"[DAKE_ERROR] code={exc.exit_code} "
        f"reason={exc.reason} message={exc.message}",
        file=sys.stderr,
    )


def process_file(input_path, output_path):
    text = input_path.read_text(encoding="utf-8")
    output_path.write_text(text, encoding="utf-8")


def run_cli(args):
    input_path = Path(args.input)
    output_path = Path(args.output)

    if not input_path.exists():
        raise DakeCliError(EXIT_INPUT, "input_not_found", "入力ファイルが見つかりません。")
    if output_path.exists():
        raise DakeCliError(EXIT_OUTPUT, "output_exists", "出力ファイルがすでに存在します。")

    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        process_file(input_path, output_path)
    except DakeCliError:
        raise
    except Exception:
        raise DakeCliError(EXIT_PROCESS, "process_failed", "処理に失敗しました。")

    return {
        "ok": True,
        "app_key": APP_KEY,
        "job_id": args.job_id,
        "output": str(output_path),
    }


def cli_main(argv=None):
    parser = build_parser()
    try:
        args = parser.parse_args(argv)
        result = run_cli(args)
        print(json.dumps(result, ensure_ascii=False))
        return EXIT_OK
    except DakeCliError as exc:
        print_cli_error(exc)
        return exc.exit_code
    except SystemExit as exc:
        return exc.code if isinstance(exc.code, int) else EXIT_USAGE
    except Exception:
        print_cli_error(DakeCliError(
            EXIT_UNEXPECTED,
            "unexpected_error",
            "処理中に予期しないエラーが発生しました。",
        ))
        return EXIT_UNEXPECTED


def gui_main():
    # 既存のTkinter GUI起動をここへ置く。
    raise NotImplementedError("GUI implementation goes here.")


def main(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    if CLI_FLAG in argv:
        return cli_main(argv)
    gui_main()
    return EXIT_OK


if __name__ == "__main__":
    raise SystemExit(main())
```

## 既存アプリへ追加するときの確認

- GUIを通常起動できる。
- `--from-shimarisu` ありでGUIが出ない。
- 入力不足でTracebackが出ない。
- 出力先が存在する場合に上書きしない。
- 成功時stdoutがJSONで読める。
- 失敗時stderrが1行で読める。
- exit codeが仕様どおり返る。