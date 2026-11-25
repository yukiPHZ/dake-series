# Dake声ログ

PCのそばで交わされた声を、発話時だけ記録し、文字にして、一日ごとにまとめて残すローカルファーストの業務記録アプリです。

## できること

- アプリ起動中にマイクを監視
- 発話を検知した区間だけ自動録音
- pre-rollで話し始めの欠けを抑制
- 日付別にWAVと文字起こしテキストを保存
- daily_transcript.mdを自動生成
- Ollamaが準備済みならdaily_summary.mdを作成
- トレイ常駐、一時停止、再開、完全終了

## 保存先

ユーザー向け記録は既定で次に保存します。

```text
Documents/Dake声ログ/YYYY-MM-DD/
```

アプリ内部状態は次に保存します。

```text
%LOCALAPPDATA%/DakeVoiceLog/
```

## 必要なもの

- Windows
- マイク
- Pythonで実行する場合は `requirements.txt` のライブラリ
- 日次要約を使う場合のみOllama

## ビルド

```bat
build.bat
```

ビルド後のexe名は `dist\DakeVoiceLog.exe` です。DAKE共通アイコンを使用します。

## 注意

- 初期版では音声と文字起こし本文をクラウドへ自動送信しません。
- 録音を行う場所のルールや必要な同意・許可は、利用環境に応じて確認してください。
- モデル未導入やOllama未設定でも、録音済み音声は削除しません。

## DAKE_META

```json
{
  "app_key": "dake_voice_log",
  "display_name": "Dake声ログ",
  "launcher_title": "声ログ",
  "launcher_description": "声を、あとから振り返れる記録に。",
  "site_title": "Dake声ログ",
  "site_description": "PCのそばで交わされた声を発話時だけ記録し、文字起こしと日次ログにまとめるローカルファーストの業務記録アプリです。",
  "update_summary": "常時待機と音声検知による録音、文字起こし、日次統合、任意のOllama要約に対応した初版。",
  "folder_name": "DAKE_Voice_Log",
  "exe_name": "DakeVoiceLog.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "draft",
  "show_in_launcher": false,
  "show_on_site": false,
  "app_type": "market",
  "completion_goal": "formal_release",
  "payment_status": "unset"
}
```

