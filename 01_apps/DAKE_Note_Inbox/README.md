---
DAKE_META:
  id: DAKE_Note_Inbox
  name: note素材受信箱
  version: 0.1.1
  status: active
  app_type: desktop
  exe_name: DakeNote_Inbox.exe
  show_in_launcher: false
  show_on_site: false
  canonical_readme: true
---

# note素材受信箱

Slackからnote素材を受け取り、`PEAKHEADZ_ROOT\INBOX`へMarkdownで保存する最小アプリです。

このアプリは補助脳ではありません。検索、通知、承認、遠隔操作、記事生成は行いません。Slackから原文を受け取り、PEAKHEADZ_ROOTへ置き、Obsidianで読むための導線だけを持ちます。

## RELEASE_BODY

DAKE_Note_Inbox v0.1.1 は、Slackから素材を差分取得し、Markdownとして `PEAKHEADZ_ROOT\INBOX` に保存する初期版のUX改善版です。

- Slack Bot Token / Channel ID を設定して差分同期
- `last_ts` を `data/note_inbox_config.json` に保存
- Slack本文は改変せずMarkdownへ保存
- Obsidian、INBOX、NOTES、ARTICLESを開くボタン
- Obsidianは実行ファイルパスを指定
- 起動時は最大化
- 最小化からタスクトレイ常駐、ダブルクリックで復帰
- 手動同期は完了ダイアログを表示し、自動同期は状態表示だけを更新
- 軽量な星背景

Ollama、記事候補、Codex素材生成、Embedding、Semantic Search、通知、承認、遠隔操作、Wake機能は含みません。

## 目的

正式パイプラインは次の通りです。

```text
Slack
↓
DAKE_Note_Inbox
↓
PEAKHEADZ_ROOT
↓
Obsidian
```

v0.1では、Slackから素材を受信して `INBOX` に保存するところまでを成立させます。

## 使い方

1. `DakeNote_Inbox.exe` または `python main.py` で起動します。
2. Slack Bot Token、Slack Channel ID、PEAKHEADZ_ROOT、Obsidian実行ファイル、同期間隔を入力します。
3. `設定保存` を押します。
4. `今すぐ同期` を押すと、Slackの差分がMarkdownで保存されます。
5. `Obsidianを開く`、`INBOXを開く`、`NOTESを開く`、`ARTICLESを開く` から確認先を開けます。

設定は `data/note_inbox_config.json` に保存されます。このファイルはGit管理しません。

Obsidian実行ファイルには、VaultではなくObsidian本体のexeを指定します。

```text
C:\Users\yukiz\AppData\Local\Programs\Obsidian\Obsidian.exe
```

## タスクトレイ

最小化するとタスクトレイへ収納されます。

タスクトレイメニュー:

- 開く
- 今すぐ同期
- Obsidianを開く
- 終了

トレイアイコンをダブルクリックすると画面へ復帰します。

## 同期ダイアログ

手動同期では完了またはエラーのダイアログを表示します。

自動同期ではダイアログを表示せず、画面上の状態表示だけを更新します。

## Slack設定方法

Slack AppでBot Tokenを作成し、対象チャンネルへBotを参加させます。

必要な基本権限は次の通りです。

- `channels:history` または対象チャンネル種別に対応する履歴取得権限
- 対象チャンネルへのBot参加

Channel IDはSlackのチャンネル詳細から取得します。TokenやWebhookの実値はREADME、ログ、Git管理ファイルへ書かないでください。

## 保存形式

保存先は `PEAKHEADZ_ROOT\INBOX` です。

```markdown
---
source: slack
channel_id: ""
timestamp: ""
status: raw
---

# Slack原文

本文
```

## ビルド

```powershell
cd C:\Users\yukiz\devlop\DAKE_series\01_apps\DAKE_Note_Inbox
.\build.bat
```

出力名は `DakeNote_Inbox.exe` です。

## 実装しないもの

- Ollama
- 記事候補
- article_candidate / article_type / article_hint
- Codex素材生成
- Embedding
- Semantic Search
- 通知
- 熾火
- 熱検索
- OIKAWA巡回
- 承認
- 遠隔操作
- Wake機能
