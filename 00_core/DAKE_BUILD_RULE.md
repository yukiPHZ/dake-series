# DAKE_BUILD_RULE

PyInstallerと `build.bat` の共通ルールです。

build方式、依存、PyInstaller、DLL / runtime、packaging、Windowsメタデータを変更する時だけ参照します。通常の文言修正・販売情報変更では毎回読みません。

## 基本方針

配布要件の正本は [DAKE_COMMON_SPEC.md](DAKE_COMMON_SPEC.md) の「Windows exe共通必須規則」。本書はそのビルド・検証手順を扱う。

- Windows向けexeを `dist` に出力する。
- Windows配布は1アプリ1exeの単独動作を基本にする。高速化のためだけの `--onedir` 化は禁止。条件を満たせない場合は原因・代案を報告する。
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
- 生成された配布exeで主要機能が成立する。
- UIの公開表示が変わった場合だけ、必要に応じてスクリーンショットを更新する。
- 空フォルダ、ZIP展開後、非開発Windowsでの重複確認は、PyInstaller方式、Python依存、DLL / runtime、packaging構成を変更した場合だけ行う。
- `_internal`、Python/DLL/アイコン等の必須隣接runtimeなしで成立する。ライセンス上必要な通知資料の同梱は実行時依存と区別する。
- 必要な実行時リソースはexe内に収める。追加依存が必要なら同梱方法を検討し、利用者の手動配置を解決策にしない。
- Windowsメタデータは `tools/generate_version_info.py` 等でbuild時に自動付与する。ProductName=対象アプリ名、FileDescription=対象アプリの日本語表示名、CompanyName=`PEAKHEADZ`、LegalCopyright=`© 2026 Yukihiko Kikuta`、Authors=`Yukihiko Kikuta`、FileVersion / ProductVersion、OriginalFilenameを基本とする。
- メタデータを人間の毎回手入力・毎回目視ゲートにしない。生成方法を変更した時に確認する。
- SHIP判定は `DAKE_RELEASE_FLOW.md` を参照する。
