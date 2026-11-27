# DAKE Quality Engine

DAKE Quality Engine は、DAKEシリーズ全体へ段階的に組み込むための共通品質基盤です。

目的は、アプリを多機能にすることではありません。exe化後も落ちにくくし、保存時に元ファイルを壊さず、設定ファイル破損時も起動を継続し、出荷前の逸脱を早めに見つけることです。

## Reliability Kit

- `reliability.py`
  - `safe_run()`
  - `format_user_error()`
  - `show_user_error()`
- GUIへTracebackを直接出さず、短い日本語メッセージに整えます。
- 詳細はログへ残します。

## Atomic IO

- `atomic_io.py`
  - `atomic_write_bytes()`
  - `atomic_write_text()`
  - `atomic_replace()`
- 本番ファイルへ直接書かず、一時ファイルへ保存後に `os.replace()` します。
- 保存失敗時は、既存ファイルを残す前提です。

## Config

- `config.py`
  - `safe_load_json_config()`
  - `safe_save_json_config()`
- JSON破損時は `.broken_YYYYMMDD_HHMMSS` へ退避し、デフォルト設定で起動継続します。
- `*_config.json` はユーザー環境依存のため、Git管理しない前提です。

## Logging

- `logging.py`
  - `write_debug_log()`
  - `get_log_path()`
- `logs/YYYY-MM-DD.log` へUTF-8で保存します。
- ログ出力失敗でアプリ本体を落としません。
- APIキー、個人情報、base64全文、Authorizationヘッダーなどをログへ保存しない運用を守ってください。

## Launch Check

- `launch_check.py`
  - `run_launch_check()`
- 各アプリの `--launch-check` から呼ぶ想定です。
- 成功時は exit `0`、失敗時は exit `1` と短いstderrを返します。

## UI Guard

- `ui_guard.py`
- 確定正本フッター（`戸建買取査定`、`by Yukihiko Kikuta`、`https://dakeapp.com/go/kaitori?app=<app_id>`）の欠落を `FAIL` 検出します。
- 旧フッターキー（`footer_left`、`footer_caption`、`footer_link_1`、`footer_link_2`、`footer_separator`、`footer_copyright`）、禁止表記（`Vibe-Coded`、`Instagram`、`sakurayk.notion.site`、フッター内 `しまりす不動産`）、`text="日本語"` の直書きを `FAIL` 検出します。

## Quality Checker

```powershell
python 00_core\dake_quality_engine\quality_check.py --app 01_apps\DAKE_PDF_Merge
```

確認するもの:

- `ORIGINAL.md` の存在および文字化け（置換文字 `\ufffd`・Mojibake）検知
- `README.md` / `DAKE_META` / `RELEASE_BODY` / `release_body.md`
- `assets/screenshot.webp` / `assets/booth_thumbnail.jpg` / `booth_ready/` / `booth_product.txt`
- `build.bat` / `main.py` / `dist/*.exe` / `release_url`
- アイコン検証（`.ico` の存在・収録解像度、`main.py` ウィンドウアイコン設定、`build.bat`/`.spec` の `--icon` 指定、`dist/*.exe` への PE アイコンリソース埋め込み）
- UI Guard（新フッター必須・旧フッターキー禁止・禁止表記・UI_TEXT直書き）

`FAIL` がある場合は exit `1`、`WARN` のみなら exit `0` です。

## Factory Hook構想

今後のDAKE_FACTORYでは、以下の順で組み込む想定です。

1. 代表アプリ1本へ Reliability Kit / Atomic IO / Config を導入
2. `--launch-check` を `run_launch_check()` に寄せる
3. `quality_check.py` を出荷前チェックへ組み込む
4. BOOTH ready / VersionInfo / ORIGINAL・README派生ビュー整合チェックと接続する

## exe後にも効く項目

- `safe_run()`
- `atomic_write_text()`
- `atomic_write_bytes()`
- `safe_load_json_config()`
- `safe_save_json_config()`
- `write_debug_log()`
- `run_launch_check()`

## 開発中だけ効く項目

- `ui_guard.py`
- `quality_check.py`

## 注意

- 今回は既存アプリへ一括導入しません。
- まず Quality Engine 本体を作り、単体検証し、次フェーズで代表アプリ1本へ導入します。
- 標準ライブラリ優先で、依存を増やしません。
