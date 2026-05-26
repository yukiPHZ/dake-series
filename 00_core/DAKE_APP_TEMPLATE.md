# DAKE_APP_TEMPLATE

## 正式出荷素材テンプレ

新規アプリは、GitHub ReleaseだけでなくBOOTH readyとdakeapp.com掲載までを前提にします。

```text
01_apps/
  DAKE_Category_Function/
    README.md
    release_body.md
    build.bat
    main.py
    assets/
      screenshot.webp
      screenshot.jpg
      booth_thumbnail.jpg
    booth_product.txt        # 正式出荷ライン上の項目名。通常の実体は booth_ready/booth_product.txt
    booth_ready/
      README.txt
      注意事項.txt
      booth_product.txt
      screenshot.jpg
      booth_thumbnail.jpg
```

新規作成時の出荷確認:

1. `release_body.md` をREADMEの `RELEASE_BODY` から生成する。
2. `assets/screenshot.webp` を起動中の実画面から作る。
3. `tools/make_booth_ready.py` で `assets/booth_thumbnail.jpg` と `booth_ready/` を作る。
4. buildして `dist/*.exe` を生成する。
5. GitHub Releaseへexeを添付する。
6. READMEの `release_url` を更新する。
7. BOOTH ready、BOOTH、dakeapp.com掲載状態を確認する。
8. Cloudflare反映確認まで終えて、正式出荷完了とする。

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
COPYRIGHT = "© 2026 しまりす不動産 / Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "main_title": "何をするか",
    "main_description": "短い説明",
    "button_execute": "実行",
    "status_idle": "待機中",
    "status_processing": "処理中",
    "status_complete": "完了しました",
    "status_error": "処理できませんでした",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_subtitle": "止まらない、迷わない、すぐ終わる。",
    "footer_link_1": "戸建買取査定",
    "footer_link_2": "Instagram",
    "footer_copyright": COPYRIGHT,
}
```

`screenshot_path` は `assets/screenshot.webp` を基本とし、実ファイルが存在していることを必須にします。
正式出荷前には `assets/booth_thumbnail.jpg` と `booth_ready/booth_product.txt` も生成済みにします。

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
4. READMEに `DAKE_META` と `RELEASE_BODY` を書く。
5. buildする。
6. 起動確認する。
7. `release_body.md` を生成する。
8. `assets/screenshot.webp` を作る。
9. `tools/make_booth_ready.py` でBOOTH素材を生成する。
10. Reviewチェックリストを見る。
