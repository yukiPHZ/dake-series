# DAKE_APP_TEMPLATE

新規・次回改修時は `DAKE_COMMON_SPEC.md` を継承する。Windows配布は単独exe必須。具体的なビルドと出荷判定は `DAKE_BUILD_RULE.md` / `DAKE_RELEASE_FLOW.md` を参照する。

## 新規アプリの基本構成

新規アプリは最初からSHIP / DISTRIBUTION / COMMERCEを分けます。通常のSHIPにBOOTH素材やStore情報を必須化しません。

```text
01_apps/
  DAKE_Category_Function/
    ORIGINAL.md             # 個別仕様の正本。共通仕様は参照で継承
    README.md
    release_body.md
    build.bat
    main.py
    assets/                  # UI変更やDISTRIBUTIONで必要な時だけ
      screenshot.webp
    booth_ready/             # BOOTH作業時だけ
      booth_product.txt
```

新規作成時のSHIP確認:

1. ORIGINALを確認する。
2. buildして `dist/*.exe` を生成する。
3. 生成された配布exeで起動と主要機能を確認する。
4. 対象差分だけcommit / pushする。
5. GitHub Releaseへexeを添付する。
6. `SHIP = COMPLETE` とする。

BOOTH / dakeapp.com / StoreはDISTRIBUTION、StripeはCOMMERCEとして別依頼で扱います。

新規DAKEアプリ作成時の基本テンプレです。

## フォルダ構成

```text
01_apps/
  DAKE_Category_Function/
    main.py
    build.bat
    README.md
    release_body.md
    requirements.txt        # 必要な場合のみ
    .gitignore
    assets/
      screenshot.webp       # Release/サイト用スクリーンショット
      booth_thumbnail.jpg   # BOOTH一覧用の正方形サムネイル
    booth_ready/
      booth_product.txt
```

## main.py基本要素

必ず検討する定義:

```python
APP_NAME = "Dakeアプリ名"
WINDOW_TITLE = "短いウインドウタイトル"
KAITORI_APP_ID = "app_id"
KAITORI_URL = f"https://dakeapp.com/go/kaitori?app={KAITORI_APP_ID}"

UI_TEXT = {
    "main_title": "何をするか",
    "main_description": "短い説明",
    "button_execute": "実行",
    "status_idle": "待機中",
    "status_processing": "処理中",
    "status_complete": "完了しました",
    "status_error": "処理できませんでした",
    "footer_kaitori": "戸建買取査定",
    "footer_author": "by Yukihiko Kikuta",
}
```

`戸建買取査定` だけを `KAITORI_URL` へリンクし、`by Yukihiko Kikuta` は非リンクにします。Instagram、長いコピーライト、Vibe-Coded、不要な外部リンクを新仕様へ入れません。

`screenshot_path` は `assets/screenshot.webp` を基本としますが、UIの公開表示が変わらない修正では毎回再生成しません。BOOTH素材はDISTRIBUTION作業時だけ生成します。

## build.bat基本要素

- `cd /d "%~dp0"` でアプリフォルダへ移動する。
- `build/`、`dist/`、`*.spec` を整理する。
- PyInstallerで `--onefile`、`--noconsole`、`--clean` を使う。
- `--icon=..\..\02_assets\dake_icon.ico` を使う。
- `--name` はREADMEの `exe_name` と合わせる。

## README.md基本構成

````markdown
# Dakeアプリ名

アプリの目的を1〜2文で説明します。

## 使い方

1. 入力または追加します。
2. 実行します。
3. 結果を確認します。

## DAKE_META

```json
{
  "app_key": "dake_app_key",
  "display_name": "Dakeアプリ名",
  "launcher_title": "短い名前",
  "launcher_description": "ランチャー用の短い説明。",
  "site_title": "Dakeアプリ名",
  "site_description": "サイト掲載用の説明。",
  "update_summary": "初回作成。",
  "folder_name": "DAKE_Category_Function",
  "exe_name": "DakeApp_Name.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## RELEASE_BODY

- 何のアプリか
- 主な操作
- 主な出力
- Windows向けexe
````

## .gitignore基本

```gitignore
__pycache__/
*.pyc
build/
dist/
*.spec
*.exe
*_config.json
```

## assets

- `assets/` はアプリフォルダ直下に置く。
- `assets/screenshot.webp` は起動直後のアプリウインドウを保存する。
- `assets/booth_thumbnail.jpg` はBOOTH一覧用に1200x1200で生成する。
- 画像を引き延ばさない。

## 新規作成時の流れ

1. フォルダを作る。
2. `main.py` に単機能UIを実装する。
3. `build.bat` を作る。
4. ORIGINALを整え、READMEに派生ビューの `DAKE_META` と `RELEASE_BODY` を置く。文言定義例は `DAKE_UI_TEXT_RULE.md` と同期する。
5. buildする。
6. 起動確認する。
7. `release_body.md` を生成する。
8. UIの公開表示が変わる場合だけ `assets/screenshot.webp` を更新する。
9. BOOTH作業を依頼された場合だけBOOTH素材を生成する。
10. 変更したレイヤーだけReviewチェックリストで確認する。
## launch-check

New apps should implement `--launch-check`.

Rules:

- Verify only that imports and minimal startup are possible.
- Return `0` on success and `1` on failure.
- Finish within a few seconds.
- Do not run file conversion, publishing, sending, browser automation, BOOTH operations, or other external side effects.
- If full support is not available yet, release review must use a short GUI smoke launch as a fallback.
