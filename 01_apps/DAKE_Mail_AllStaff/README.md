# Dake全社員メール起動

全社員宛メールの新規作成画面を、宛先とCCが入った状態で開くDAKEシリーズの単機能Windowsデスクトップアプリです。

送信は自動実行されません。件名、本文、履歴、複数パターン切替などの機能もありません。

## 使い方

1. `Dake_AllStaff_Mail.exe` を起動します。
2. `メールを開く` をクリックします。
3. 既定のメーラーで、TOとCCが入力済みの新規メール作成画面が開きます。
4. 内容を確認して、送信はメーラー上で手動で行います。

## 設定ファイル

初回起動時に、アプリと同じ階層へ `Dake_AllStaff_Mail_config.json` が自動生成されます。

```json
{
  "to": "all@example.co.jp",
  "cc": "example@example.co.jp"
}
```

実運用前に、`to` と `cc` を実際の宛先へ差し替えてください。複数宛先はカンマ区切りで指定できます。

`cc` が空の場合は、CCなしでメーラーを起動します。`to` が空の場合は、メーラーを起動しません。

設定ファイルはユーザー環境依存のため、Git管理対象外です。

## build方法

Python と PyInstaller が使える環境で、アプリフォルダ内の `build.bat` を実行します。

```bat
build.bat
```

出力されるexe名は `Dake_AllStaff_Mail.exe` です。共通アイコンとして `..\..\02_assets\dake_icon.ico` を参照します。

`build.bat` は、PATH上の `pyinstaller` を優先して使用します。見つからない場合は、利用可能な Python から `python -m PyInstaller` を実行します。

## DAKE共通仕様レビュー（2026-05-06）

最新のDAKE共通仕様に合わせて、以下を確認・修正しました。

- フォントは `BIZ UDPGothic` を最優先、`Yu Gothic UI`、`Meiryo` の順でフォールバックします。
- ヘッダーはアプリ名の重複を避け、機能タイトルと短い説明文のみを表示します。
- フッターは `シンプルそれDAKEシリーズ ｜ 止まらない、迷わない、すぐ終わる。` を1行固定で表示します。
- フッターリンクは `戸建買取査定` と `Instagram` のみクリック可能です。通常時は補助文字色、ホバー時のみアクセント色にします。
- UI文言は `APP_NAME`、`WINDOW_TITLE`、`COPYRIGHT`、`UI_TEXT` に集約しています。
- `text=` への日本語直書きがないことを確認しました。
- 共通アイコン `..\..\02_assets\dake_icon.ico` をTkinter起動時とPyInstallerビルド時に参照します。
- `version_info.txt` により、Windowsのファイル説明と製品名を `全社員メール` に設定しています。

### ビルド・起動確認

2026-05-06 に `build.bat` でビルドし、`dist\Dake_AllStaff_Mail.exe` の生成を確認しました。

生成されたexeは短時間起動確認を行い、即終了しないことと、`Dake_AllStaff_Mail_config.json` が初回自動生成されることを確認しました。

## DAKEシリーズ

このアプリは「全社員メールを開くだけ」のDAKEシリーズ単機能アプリです。

送信、自動化、本文テンプレート、件名テンプレート、履歴保存、宛先パターン切替、アドレス帳機能は実装していません。
