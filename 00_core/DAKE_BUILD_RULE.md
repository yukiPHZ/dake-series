# DAKE_BUILD_RULE

PyInstallerと `build.bat` の共通ルールです。

## 基本方針

- Windows向けexeを `dist` に出力する。
- 1アプリ1exeを基本にする。
- exe名はREADMEの `DAKE_META.exe_name` と一致させる。
- build成果物はRelease配布用であり、Gitの正本にしない。

## build.bat基本形

```bat
@echo off
chcp 65001 > nul
cd /d "%~dp0"

rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
del /q *.spec 2>nul

pyinstaller ^
--onefile ^
--noconsole ^
--clean ^
--icon=..\..\02_assets\dake_icon.ico ^
--name DakeApp_Name ^
main.py

pause
```

## 必須に近い指定

- `--onefile`: 配布しやすい単一exeにする。
- `--noconsole`: GUIアプリではコンソールを出さない。
- `--clean`: 古いbuild状態の影響を減らす。
- `--icon=..\..\02_assets\dake_icon.ico`: 共通アイコンを使う。
- `--name`: Releaseで扱うexe名を固定する。

## build前整理

build前に削除または整理するもの:

- `build/`
- `dist/`
- `*.spec`

注意:

- 削除はアプリフォルダ内だけで行う。
- 別アプリや上位フォルダを巻き込まない。
- PowerShellやcmdの文字コードで日本語が壊れないようにする。

## hidden-importとcollect-data

ライブラリによってはPyInstallerが自動検出できない。

必要に応じて追加する例:

```bat
--hidden-import=tkinterdnd2 ^
--collect-all=tkinterdnd2 ^
--collect-all qrcode ^
--collect-submodules qrcode ^
--collect-data qrcode ^
```

ルール:

- build失敗や起動時ImportErrorが出た場合だけ追加する。
- 追加理由をREADMEや作業メモに短く残す。
- 不要なcollect-allを増やしすぎない。

## 依存関係

- `requirements.txt` がある場合はbuild前に確認する。
- 特殊ライブラリはREADMEに注意点を書く。
- CodexランタイムのPythonを使う場合でも、通常ユーザー環境でのbuild可否を意識する。

## exe名

- 日本語exe名は避け、英数字とアンダースコア中心にする。
- 表示名はアプリ内とREADMEで日本語にしてよい。
- `DAKE_META.exe_name` は実際の配布exe名と一致させる。

## build後確認

- `dist/<exe_name>.exe` が存在する。
- exeが起動する。
- 共通アイコンが表示される。
- 画面が文字化けしていない。
- 起動直後スクリーンショットが撮れる。
