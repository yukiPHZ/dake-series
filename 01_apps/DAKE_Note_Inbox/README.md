---
DAKE_META:
  id: DAKE_Note_Inbox
  name: DAKE_Note_Inbox
  display_name: note素材受信箱
  version: 0.2.0
  status: active
  app_type: desktop
  exe_name: DakeNote_Inbox.exe
  show_in_launcher: false
  show_on_site: false
  canonical_readme: true
---

# DAKE_Note_Inbox

表示名「note素材受信箱」。Slackからnote素材を受け取り、`PEAKHEADZ_ROOT\INBOX`へMarkdownで保存する最小アプリです。

このアプリは補助脳ではありません。検索、通知、承認、遠隔操作、記事生成は行いません。Slackから原文を受け取り、PEAKHEADZ_ROOTへ置き、Obsidianで読むための導線だけを持ちます。

## RELEASE_BODY

DAKE_Note_Inbox v0.2.0 は、Slackから素材を差分取得して `PEAKHEADZ_ROOT\INBOX` に保存し、Ollamaで軽い札付けを行って `PEAKHEADZ_ROOT\NOTES` に素材Markdownを作る版です。

- Slack Bot Token / Channel ID を設定して差分同期
- `last_ts` を `data/note_inbox_config.json` に保存
- Slack本文は改変せずMarkdownへ保存
- Obsidian、INBOX、NOTES、ARTICLESを開くボタン
- Obsidianは実行ファイルパスを指定
- Obsidian.exeの参照ボタン
- 未設定時のObsidian自動探索
- 設定は `%APPDATA%\DAKE_Note_Inbox\note_inbox_config.json` に保存
- Ollama使用ON/OFF
- Ollamaモデル名設定
- INBOX raw MarkdownからNOTES material Markdownを生成
- Obsidian内部リンクと記事化メモを付与
- 起動時は最大化
- 最小化からタスクトレイ常駐、ダブルクリックで復帰
- 手動同期は完了ダイアログを表示し、自動同期は状態表示だけを更新
- 軽量なコネクティングドッツ背景

記事候補、記事本文生成、Codex素材生成、Embedding、Semantic Search、通知、承認、遠隔操作、Wake機能は含みません。

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
3. Ollama使用ON/OFFとOllamaモデル名を設定します。
4. `設定保存` を押します。
5. `今すぐ同期` を押すと、Slackの差分がMarkdownで保存されます。
6. `札付けする` を押すと、INBOXの未処理MarkdownからNOTESへ素材Markdownが作られます。
7. `Obsidianを開く`、`INBOXを開く`、`NOTESを開く`、`ARTICLESを開く` から確認先を開けます。

設定はユーザー領域に保存されます。このファイルはGit管理しません。

```text
%APPDATA%\DAKE_Note_Inbox\note_inbox_config.json
```

Obsidian実行ファイルには、VaultではなくObsidian本体のexeを指定します。

```text
C:\Users\yukiz\AppData\Local\Programs\Obsidian\Obsidian.exe
```

未設定時は標準インストール先とスタートメニューのObsidianショートカットを自動探索します。起動時は `PEAKHEADZ_ROOT` をVaultとして開くことを試み、失敗時はObsidian URIでの起動へフォールバックします。

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

## Ollama札付け

Ollamaは記事を書く係ではありません。Slack原文にタグ、Obsidianリンク、短い記事化メモを付ける係です。

接続先:

```text
http://127.0.0.1:11434/api/generate
```

初期モデル:

```text
qwen2.5:7b
```

Ollamaが未起動、またはモデル応答に失敗した場合もアプリは落ちません。fallback札付けで `ollama_status: fallback` の素材Markdownを保存します。Ollama使用をOFFにした場合は `ollama_status: disabled` として保存します。

raw Markdownは改変しません。material Markdownは `PEAKHEADZ_ROOT\NOTES` に新規保存されます。

付与する項目:

- `tags`
- `links`
- `article_hint`
- `ollama_status`

付与しない項目:

- score
- ranking
- article_candidate
- article_type

## material保存形式

```markdown
---
source: "slack"
status: material
original_path: ""
slack_ts: ""
tags:
  - "note素材"
links:
  - "[[在る]]"
article_hint: "この断片は、note記事の素材として後から読み返せる。"
ollama_status: "fallback"
---

# 原文

Slack本文

---

# 札付け

## タグ

#note素材 #在る

## Obsidianリンク

[[在る]]

## 記事化メモ

この断片は、note記事の素材として後から読み返せる。
```

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

- 記事候補
- article_candidate / article_type
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
