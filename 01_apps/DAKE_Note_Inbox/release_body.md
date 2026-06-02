DAKE_Note_Inbox v0.2.0

Slackからnote素材を受け取り、PEAKHEADZ_ROOTのINBOXへMarkdown保存し、Ollamaで軽い札付けを行ってNOTESへ素材Markdownを作る版です。

- Slack Bot Token / Channel ID による差分同期
- last_ts管理
- 設定保存先をAPPDATAへ変更
- Slack原文のMarkdown保存
- Ollama使用ON/OFF
- Ollamaモデル名設定
- INBOX raw MarkdownからNOTES material Markdownを生成
- tags / links / article_hint / ollama_status の最小札付け
- 起動時最大化
- Obsidian、INBOX、NOTES、ARTICLESを開く導線
- Obsidian実行ファイル指定と参照ボタン
- Obsidian標準パス自動探索とURI fallback
- タスクトレイ常駐とダブルクリック復帰
- 自動同期の完了ダイアログ抑制
- 軽量なコネクティングドッツ背景

Ollamaは記事を書く係ではなく札付け係です。記事候補、記事本文生成、Codex素材生成、Embedding、Semantic Search、通知、承認、遠隔操作、Wake機能は含みません。
