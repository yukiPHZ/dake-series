# DAKE_BUILD_RULE

PyInstallerと `build.bat` の共通ルールです。

## 基本方針

配布要件の正本は [DAKE_COMMON_SPEC.md](DAKE_COMMON_SPEC.md) の「Windows exe共通必須規則」。本書はそのビルド・検証手順を扱う。

- Windows向けexeを `dist` に出力する。
- Windows配布は1アプリ1exeの単独動作を必須にする。高速化のための `--onedir` 化は禁止。条件を満たせない場合は出荷保留とし、原因・代案を報告する。
- exe名はREADMEの `DAKE_META.exe_name` と一致させる。
- build成果物はRelease配布用であり、Git管理対象にしない。

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

## 必須指定

- `--onefile`: PyInstallerの正式配布方式。単独動作を検証し、隣接runtimeを前提にしない。
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

次の確認は、次回の許可されたビルド・出荷時に行う。文書更新だけの工程では実行しない。

- `dist/<exe_name>.exe` が存在する。
- exeが起動する。
- 共通アイコンが表示される。
- 画面が文字化けしていない。
- 起動直後スクリーンショットが撮れる。
- 空のフォルダへexeだけをコピーして起動し、主要機能まで確認する。開発PC以外の、Python等の追加環境構築がないWindowsでも確認する。
- `_internal`、Python/DLL/ライセンス/アイコン等の隣接フォルダ・ファイルなしで成立する。ZIP配布では展開後のexeだけでも同じ確認を行う。
- 必要な実行時リソースはexe内に収める。追加依存が必要なら同梱方法を検討し、利用者の手動配置を解決策にしない。
- WindowsメタデータのCompanyNameはPEAKHEADZ、個人制作者・著作権者はYukihiko Kikuta。画面フッターの文字列をファイルプロパティへ機械的に流用しない。
- 出荷判定と実測証跡は `DAKE_RELEASE_FLOW.md` のWindows単独exeゲートへ記録する。
