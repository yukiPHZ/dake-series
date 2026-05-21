# BRAINZ

記憶庫 / 取り込み母艦。

補助脳BRAINZ は、ローカルの会話・仕様・メモ・README・Codex結果を忘れず保存し、OIKAWAへ渡すための記憶庫・取り込み母艦です。

BRAINZはChatGPTやCodexの代替AIではありません。ローカルに置いた記憶を読み込み、取り込み、indexし、過去の思考や作業結果へ再接続するための保存層・橋渡し係です。

## QPSC正本ルート: PEAKHEADZ_ROOT

BRAINZは、今後のQPSC正本ルートとして `C:\Users\yukiz\Documents\PEAKHEADZ_ROOT` を扱います。

- 未設定時の推奨保存先は `PEAKHEADZ_ROOT` です
- 既存configに `memory_folder` がある場合は、その保存先を勝手に変更しません
- 旧 `brainz_memory` はlegacy fallbackとして扱い、削除もrenameもしません
- BRAINZの `PEAKHEADZ_ROOTを作成` 導線で標準フォルダ構成を作れます
- 自動移行は行わず、移行する場合もcopy候補として扱います

標準フォルダ構成:

```text
PEAKHEADZ_ROOT/
  00_inbox/
  10_slack/
    brainz-inbox/
    brainz-note/
    brainz-codex/
    brainz-aru/
    brainz-reaction/
  20_chatgpt/
  30_codex/
  40_borinef/
    note/
      published/
      reactions/
      source_fragments/
  90_system/
    logs/
    config/
```

Obsidianでは `PEAKHEADZ_ROOT` を保管庫として開きます。Obsidianは本体ではなくMarkdown観測UIです。Obsidianを消しても、正本Markdownはローカルに残ります。

## 思想

補助脳が提案し続ける。  
菊田が選び続け、決め続ける。  
補助脳が決定を反映し続ける。

BRAINZの役割は、忘れない、探せる、関連づく、再投入できる、ChatGPT / Codex / Claude / Gemini の橋渡しをすることです。

## QPSCでの役割

QPSC（Quiet Personal Cognitive System）では、BRAINZを「見るアプリ」ではなく、深層記憶・取り込み・index・設定を担う記憶庫として扱います。

BRAINZ側で管理するもの:

- ペースト投稿取り込み
- Slack取り込み
- ChatGPT export取り込み
- Codex報告取り込み
- 今後増える接続端末設定
- 接続設定ページ

検索・原本表示・熱検索・熾火・熱提案・通知表示は、前面UIであるOIKAWAへ寄せていきます。BRAINZのUIは検索閲覧を主目的にせず、脳をイメージした抽象的で微細な動きが漂う程度に整理していく方針です。

起動時には `data/config/qpsc_brainz_status.json` に `brainz_awake`、`started_at`、`last_heartbeat_at`、`status_message` を記録します。起動中は一定間隔で heartbeat を更新し、OIKAWAやDAKE_Wake_Brainzが状態確認に使える導線にします。

## QPSC v0.1 正本

BRAINZはQPSC v0.1で、前面の読む場所ではなく、記憶庫 / 取り込み母艦として固定します。

- 記憶フォルダ、取り込み、index、接続/設定、awake/heartbeatを担います
- ペースト投稿、Slack、ChatGPT export、Codex報告を取り込みます
- 取り込み時に `data/config/qpsc_notifications.json` へ通知イベントを保存します
- 起動時とheartbeatで `data/config/qpsc_brainz_status.json` を更新します
- OIKAWAを開く導線を持ちます
- 検索、原本表示、熾火、ORBITはOIKAWA側を正面にします

v0.1では、BRAINZが判断するのではなく、原本と状態を静かに残します。OIKAWAがそれを前面へ戻し、最後に菊田が選びます。

## QPSC v0.2 現在地

BRAINZはv0.2でも、読む前面ではなく記憶庫 / 取り込み母艦です。検索、原本表示、熾火、ORBIT、巡回、側に在る表示はOIKAWA側へ寄せます。

- ChatGPT exportは `zip`、展開済みフォルダ、`conversations.json`、`conversations-000.json` 形式に対応します
- 巨大会話はchunk分割し、Ollama embeddingのcontext超過を避けます
- Slack通知文はログではなく、記憶へ戻る入口として整流します
- Codex報告は要約ではなく、正本Markdown / TXTとして保存します
- BRAINZは取り込み、index、通知保存、awake/heartbeat、設定、OIKAWA起動導線に集中します
- OpenClaw GatewayとSlack通知本格化はまだ先です

## QPSC Phase 2: 取り込み通知

Phase 2では、BRAINZで発生した取り込みをOIKAWA側の通知・提案欄へ静かに渡す導線を追加しました。

- BRAINZは取り込み完了時に `data/config/qpsc_notifications.json` へ通知イベントを保存します
- 対象はペースト投稿、Slack、Codex報告、ChatGPT exportの取り込みです
- 通知イベントには `source`、`title`、`message`、`status`、`kind`、`related_path` を保持します
- 原本主義を維持し、`related_path` がある通知はOIKAWAから元のMarkdownや取り込み元へ戻れる設計です

Phase 3以降、原本表示と読む導線はOIKAWA側へ寄せます。BRAINZは保存・取り込み・indexを優先し、前面の閲覧体験を増やしすぎない方針です。

Phase 4以降、熾火表示と熱検索入口もOIKAWA側で扱います。BRAINZは通知イベントと原本への `related_path` を残し、候補表示や読み戻しはOIKAWAに寄せます。

Phase 5以降、今日の整理表示もOIKAWA側で扱います。BRAINZは取り込み通知と原本パスを残し、ORBITの集計・表示・読み戻しはOIKAWAに寄せます。

Phase 14以降、OllamaはOIKAWA側の熾火補助として扱います。BRAINZは原本と通知を保存し、AIによる熱補助や表示はOIKAWAに寄せます。

Phase 15以降、巡回 / 再訪表示もOIKAWA側で扱います。BRAINZは原本と通知を保存し、何度も戻っている記憶の表示やORBITへの反映はOIKAWAに寄せます。

Phase 16以降、「静かに浮く」表示もOIKAWA側で扱います。BRAINZは原本と通知を保存し、側に在る記憶の生成・表示・ORBIT反映はOIKAWAに寄せます。

Phase 19以降、時間を跨いだ熱と静かだった記憶の浮上もOIKAWA側で扱います。BRAINZは原本を保持し、OIKAWAが巡回ログから長期間後の再訪、深夜巡回、巡回増加を静かに表示します。

## QPSC Phase 8: 静かな母艦UI

Phase 8では、BRAINZを「見るアプリ」ではなく、記憶庫・取り込み母艦として静かに扱うトップ画面へ整理しました。

- BRAINZは記憶の保存、取り込み、index、設定、状態確認を担います
- OIKAWAが検索、原本表示、熾火、ORBITを担います
- BRAINZはOIKAWAを開く導線を持ち、読む場所へ静かに移れます
- 検索や原本確認はBRAINZ内では補助表示とし、主役化しません
- 取り込み、接続/設定、最終取り込み状態、awake/heartbeatをBRAINZ側の主情報として扱います

### OIKAWA起動導線

BRAINZからOIKAWAを開く時は、まず `DAKE_Brainz_OIKAWA/dist/DakeBrainz_OIKAWA.exe` を優先します。

完全一致で見つからない場合も、`dist` 内の `.exe` を探索します。1つだけある場合はそれを使い、複数ある場合は `OIKAWA` / `Oikawa` / `Brainz` を含むexeを優先します。

exeがない開発中の状態では、`DAKE_Brainz_OIKAWA/main.py` を `python main.py` として起動します。buildされていない場合でも、開発起動できる導線です。

exeと `main.py` のどちらも見つからない場合は、削除や強制終了はせず、OIKAWAのbuildまたは配置確認を促す静かな状態表示だけにします。BRAINZ Logには探索した候補パスを短く残します。

## QPSC Phase 11: ChatGPT export取り込み導線

Phase 11では、BRAINZの取り込み入口にChatGPT export専用カードを追加しました。

- `zip`、展開済みフォルダ、`conversations.json` を選んで取り込めます
- 新形式の `conversations-000.json` / `conversations-001.json` 形式も、ファイル名順に結合して取り込めます
- `chat.html` は現時点では取り込み対象にせず、JSON正本を優先します
- 各メッセージは、chunk分割やembeddingより先に `PEAKHEADZ_ROOT/20_chatgpt/` へMarkdown原本として保存します
- 巨大会話は小さめのchunkに分割し、Ollama embeddingのcontext超過を避けます
- 取り込み処理は既存の `core/chatgpt_importer.py` を使い、重複した取り込みロジックは持ちません
- 取り込み中と結果はBRAINZ画面内に短く表示します
- 取り込み成功時は `data/config/qpsc_notifications.json` に `chatgpt_export` 通知を保存します
- 通知の `related_path` はBRAINZ側の取り込みログを指し、OIKAWAから取り込み元と結果を確認できます
- 取り込んだ記憶はOIKAWAで検索・通知・原本表示・熾火・ORBITの導線から戻れます

ChatGPT exportの基本手順:

1. ChatGPTからexportを取得します
2. BRAINZで `zip`、`フォルダ`、`conversations.json`、または `conversations-000.json` を選びます
3. 取り込み完了後、OIKAWAの通知や検索から記憶へ戻ります

対応形式:

- `conversations.json`
- `conversations-000.json` 形式の分割JSON
- ChatGPT export zip
- 展開済みChatGPT exportフォルダ

## QPSC UI表示方針

QPSCのBRAINZ / OIKAWAは黒基調の静かなUIを基本にします。

- 背景は黒〜ダークグレーを優先します
- 日本語本文は `Yu Gothic UI` / `Meiryo` を優先します
- 小さすぎる日本語や太字の多用は避けます
- 原本、通知、取り込み状態は読みやすさを優先します

## QPSC Phase 12: Slack通知文の整流

Phase 12では、Slack取り込み通知をログではなく、記憶へ戻る入口として整流しました。

- Slack投稿はBRAINZへ保存し、OIKAWAの通知・熾火・ORBITへ流れます
- 通知文は命令せず、短く、戻れる入口として扱います
- `handoff_codex`、`note:`、`search:`、`import:` は軽く分類して通知タイトルへ反映します
- 原本主義を維持し、通知の `related_path` から保存済みMarkdownへ戻れます

## QPSC Phase 13: Codex報告の正本入口化

Phase 13では、Codex結果・Codex報告の通知を、正本へ戻る入口として整流しました。

- Codex報告は要約せず、原本Markdown / TXTを正本として保存します
- 通知文は短く、命令せず、OIKAWAから戻れる入口として扱います
- commit、push、修正ファイル、確認が必要な報告は軽く分類してmessageへ反映します
- OIKAWAではCodex報告通知、熾火、ORBITから原本へ戻れます

## QPSC Phase 18: Slack静かな通知導線

Phase 18では、QPSC内部で静かに浮いた記憶をSlackへ小さく置くための設定を追加しました。

- 設定は `data/config/brainz_config.json` の `slack_notify_enabled`、`slack_webhook_url`、`slack_notify_max_per_day`、`slack_notify_quiet_hours` で扱います
- Webhook URLはローカルconfigに保存し、Git管理しません
- 通知は命令ではなく、OIKAWAへ戻るための再接続として扱います
- 大量通知は避け、初期値では1日最大3件に抑えます
- OpenClaw本格導入前の軽い外部導線です

## QPSC Phase 20: BRAINZ母艦UIと保存地形

Phase 20では、BRAINZを検索や原本表示の前面から外し、Slack / ChatGPT export / CodexログをローカルPCへ保存する母艦UIへ寄せます。

- Slackは投入口です。`#brainz-inbox` は巡回される熱、`#brainz-aru` は育てる熱として扱います
- BRAINZは新しい `PEAKHEADZ_ROOT/` 地形を作成できます
- Slackチャンネル別保存設定を持ち、保存先、source、OIKAWA通知、heat対象を設定できます
- Codexログ自動監視入口として `%USERPROFILE%\.codex\sessions\` の存在確認を置きます
- 自由ペースト投稿は削除せず、緊急保存として残します

新しい保存地形:

```text
PEAKHEADZ_ROOT/
  00_inbox/
  10_slack/
    brainz-inbox/
    brainz-note/
    brainz-codex/
    brainz-aru/
    brainz-reaction/
  20_chatgpt/
  30_codex/
  40_borinef/
    note/
      published/
      reactions/
      source_fragments/
  90_system/
    logs/
    config/
```

既存の `slack/`、`aru/`、`codex_reports/`、`chatgpt/`、`BORINEF/` は移行候補です。削除せず、OIKAWAやOpenClawが巡回しやすい地形へ少しずつ寄せます。

OpenClawはまだ本格導入しません。現時点では、BRAINZが保存母艦、OIKAWAが巡回 / 熾火 / 正本ニュース、Obsidianが観測UIという役割をREADME上で固定します。

## QPSC Phase 20: BORINEF note循環

Phase 20では、SlackへnoteタイトルとURLを投げるだけで、公開済noteをBRAINZへ戻す循環を追加しました。

- `https://note.com/` を含むSlack投稿を検出します
- URL直前行、または先頭行をタイトルとして扱います
- 保存先は `PEAKHEADZ_ROOT/40_borinef/note/published/YYYY/` です
- Markdownは `status: published`、`platform: note`、`url` を持つpublished正本として保存します
- 同じnote URLが既に保存済みなら、再保存せず静かにskipします
- 保存後は `borinef_note` 通知を `qpsc_notifications.json` に残し、OIKAWAの巡回・熾火・熱検索対象に自然に混ざります
- reactions保存は今後 `PEAKHEADZ_ROOT/40_borinef/note/reactions/` に寄せる予定です

## Phase 1でできること

- ローカル記憶フォルダの指定
- `.txt` / `.md` / `.json` の読み込み
- SQLite `brainz.db` への保存
- SQLite FTS5 による全文検索
- うろ覚え検索
- Ollama接続チェック
- CUDA / GPUチェック表示
- 検索結果一覧
- 結果詳細表示
- ChatGPTに貼る用まとめ生成
- Codexに貼る用指示素材生成
- index更新ログ保存
- QPSC起床状態ファイルの記録

## Phase 1でやらないこと

- ChatGPT API接続
- OpenAI API接続
- OpenAI APIキー入力欄
- 自動でChatGPT / Codexへ投稿
- Claude / Gemini自動ログイン
- ブラウザ操作自動化
- 人格AI化
- 自動判断
- クラウド同期
- 自動削除
- 勝手なファイル移動

## OpenAI APIについて

OpenAI APIは使いません。APIキー入力欄もありません。

基本方針は完全ローカル優先です。OllamaがあればローカルembeddingによるSemantic Searchに使います。Ollamaがなくても SQLite FTS5 による全文検索とうろ覚え検索は動きます。

## ローカル記憶フォルダ

ユーザーが自由に指定できます。初期案は次のような構成です。

```text
brainz_memory/
├ chatgpt/
├ codex/
├ claude/
├ gemini/
├ ideas/
├ specs/
├ README/
├ logs/
└ thoughts/
```

Phase 1の読み込み対象は `.txt` / `.md` / `.json` です。原本ファイルは編集しません。削除もしません。BRAINZは内容を読み取り、SQLiteに索引として保存します。

## 検索

全文検索は SQLite FTS5 を使います。うろ覚え検索は、全文検索に加えて `LIKE` と検索語の分割で拾います。

例:

- 静かな青
- 補助脳
- Codexに投げたやつ
- DAKEのGitルール
- quiet workflow

## Ollama / GPU

起動時に可能な範囲で `nvidia-smi`、CUDA相当のGPU検出、Ollamaのローカル接続を確認します。

表示例:

- CUDA ONLINE
- GPU DETECTED
- OLLAMA LOCAL READY
- SQLITE READY

未検出でもアプリは落ちません。

## ChatGPT / Codex / Claude / Gemini との関係

BRAINZは自動送信しません。検索結果から、ChatGPTに貼る用まとめとCodexに貼る用素材をMarkdownとして生成するだけです。Claude / Gemini についてもPhase 1では自動連携しません。

## 将来対応候補

- `.html`
- ChatGPT export zip
- `conversations.json` parser
- Claudeコピー
- Codex結果ログ
- PDF OCR
- 音声Whisper
- screenshots
- browser logs
- FAISS検索
- FAISS GPU対応
- 関連タグ自動生成
- 選択履歴 / 決定履歴
- ChatGPT投入用プロンプト生成強化
- Codex指示書生成強化
- Git差分連携
- 処理完了通知

## ビルド

```bat
build.bat
```

PyInstallerで `dist/DakeBrainz_Search.exe` を作成します。exeアイコンは `../../02_assets/dake_icon.ico` を参照します。PEAKHEADZロゴはアプリ内表示用で、`assets/peakheadz_logo.png` が存在する場合のみ表示します。



## Phase 2: ChatGPT export取り込み

補助脳BRAINZ v0.2.0 では、ChatGPTの公式Export zipまたは展開済みフォルダから `conversations.json` を検出して取り込めます。

- ChatGPT公式Export zipに対応
- 展開済みフォルダに対応
- `conversations.json` をローカルで解析
- 会話メッセージを `PEAKHEADZ_ROOT/20_chatgpt/` にMarkdown原本として先に保存
- 会話タイトル・発言者・本文をindex化
- 取り込んだ会話は `source_type=chatgpt_export` として検索可能
- 同じexportを再取り込みしても重複を避ける
- 元zip・元フォルダ・元JSONは編集しない
- OpenAI APIは使わない
- APIキー入力欄は追加しない

BRAINZはChatGPTの代替ではありません。ChatGPTの会話履歴を、ローカル記憶として再接続するための検索補助脳です。

## Phase 3: Codex結果ログ取り込み

補助脳BRAINZ v0.3.0 では、Codexの完了報告・修正結果・commit / push結果を、貼り付けまたは `.txt` / `.md` ファイルから取り込めます。

- Codexの完了報告を貼り付けで取り込み
- `.txt` / `.md` ファイルから取り込み
- commit hash / 修正ファイル / 確認結果 / push結果を可能な範囲で抽出
- `source_type=codex_result` として検索可能
- ChatGPT用handoff / Codex用handoffに実装履歴を反映
- 同じCodex結果の再取り込み時は重複を避ける
- Codexを自動操作しない
- 外部送信しない
- OpenAI APIは使わない

BRAINZは、Codexが実際に何を作成・修正・確認・commit / pushしたかをローカル記憶として保存します。ChatGPTで仕様検討し、Codexが実装し、BRAINZが覚えて橋渡しするための履歴層です。

## Phase 5: Semantic Search

補助脳BRAINZ v0.4.0 では、Ollama embeddingを使ったSemantic Searchに対応しました。

- Ollama embeddings APIに対応
- 推奨モデルは `nomic-embed-text`
- 完全ローカル処理
- OpenAI API未使用
- embedding unavailableでもFTS検索は継続
- 意味的に近い記録をRelated Memoryとして表示
- Semantic Search ON / OFF
- `.txt` / `.md` / `.json`、ChatGPT export、Codex結果ログのchunkにembeddingを保存

Ollama未起動、モデル未導入、embedding API失敗、GPU未検出のいずれでもアプリは落ちません。Semantic Searchが使えない場合はFTS only modeとして従来検索を使います。

## Phase 6: Related Timeline / Memory Flow

補助脳BRAINZ v0.5.0 では、検索結果を選択した時にRelated Timeline / Memory Flowを表示します。

- 記憶の前後関係を表示
- 同一conversation / semantic類似 / 時系列近接を利用
- title類似 / 同一source_type / 関連commitを補助的に利用
- 古い記憶から新しい記憶へ辿る表示
- timeline itemクリックでPreviewへ移動
- Semantic Search OFFでもtitle / date / source_type中心で生成
- 完全ローカル処理
- OpenAI API未使用

BRAINZは単語を探すだけの検索ツールではなく、過去の記録を意味と流れで再接続する補助脳です。

## Phase 7: PEAKHEADZロゴ表示

補助脳BRAINZは、`assets/peakheadz_logo.png` がある場合にヘッダーへ小さくPEAKHEADZロゴを表示します。

- ロゴは左上ヘッダー付近に小さく表示
- 縦横比を維持
- ロゴがない場合も通常起動
- BRAINZはPEAKHEADZロゴをアプリアイコンに使用

## Phase 8: アイコン・フォント・検索中UI改善

補助脳BRAINZは、PEAKHEADZロゴをexeアイコンにも使用します。

- `assets/peakheadz_logo.ico` をビルドアイコンに使用
- フォント可読性を改善
- Preview / Log / Search Results の文字サイズと行間を調整
- 検索中ステータスとSearchボタンのSearching表示を追加

## Phase 9: Watch Folder / Auto Index

補助脳BRAINZは、指定したWatch Folderを静かに監視し、新規/更新された `.txt` / `.md` / `.json` を自動indexできます。

- Watch Folderを設定可能
- Auto Index ON/OFFに対応
- polling方式で5〜10秒ごとにローカル確認
- modified time / hashを利用して新規・更新ファイルを検出
- `build/`、`dist/`、`__pycache__/`、`.git/`、`.venv/`、`node_modules/`、`data/logs/`、`data/exports/` は監視対象から除外
- Semantic Searchが使える場合は新規indexにもembedding生成を試行
- Ollama / embedding unavailableでもFTS検索と通常indexは継続
- 完全ローカル処理
- OpenAI API未使用

## UI再調整: フォントと読みやすさ

補助脳BRAINZは、長時間読める記憶OSとしてフォントと余白を再調整しました。

- フォント可読性を再調整
- 日本語太字の潰れを避けるため、サイズと余白で階層化
- Logは等幅フォント、本文は読みやすい日本語フォントへ調整
- Search Results / Preview / BRAINZ Log の行間と文字色を整理

## Phase 10: Notification system

補助脳BRAINZは、処理完了や新規記憶検出をアプリ内の小さな通知で静かに知らせます。

- Enable Notifications ON/OFFに対応
- Auto Index通知
- ChatGPT / Codex import通知
- Semantic更新通知
- handoff生成通知
- Notification queueで順番表示
- 通知はBRAINZ Logにも保存
- Windows toastは任意扱いで、Phase 10ではアプリ内通知を優先
- 完全ローカル処理

## Phase 11: Remote Queue

補助脳BRAINZは、remote_queueフォルダに置かれた `.md` / `.txt` / `.json` を検出し、処理予約として受け取れます。

- Remote Queue Folderを設定可能
- Enable Remote Queue ON/OFFに対応
- import / search / handoff_chatgpt / handoff_codex / note task対応
- note taskは `source_type=remote_queue_note` として記憶DBへ取り込み
- search / handoff taskは検索欄へqueryをセットし、候補として扱う
- 処理済みは `processed/`、失敗は `failed/` へ移動
- Queue履歴をSQLiteへ保存
- Remote Queue通知に対応
- 外部通信しない
- スマホ連携は同期フォルダ経由を想定
- ChatGPT / Codexへ自動送信しない

## Phase 15: Codex Report Auto Import

補助脳BRAINZは、`codex_reports/` に置かれた `.md` / `.txt` のCodex報告を自動検出し、`source_type=codex_report_auto` として記憶できます。

- Codex Report Auto Import
- `codex_reports/` に置くだけで自動記憶
- Markdown正本をそのまま保持
- 要約しない
- 自動圧縮しない
- commit hash / changed files / build結果 / push結果を可能な範囲で抽出
- 処理済みは `processed/`、失敗は `failed/` へ移動
- Semantic Search / Related Memory / Memory Flow と連携
- 通知ON時はCodex報告取り込みを通知
- OpenAI API未使用

記憶フォルダ例：

```text
brainz_memory/
├ codex_reports/
│  ├ processed/
│  └ failed/
```

## Phase 16: Slack Inbox Import

補助脳BRAINZ v1.0.0 では、Slackの専用チャンネルをスマホ共通の投入口として使い、投稿されたメモをローカルMarkdown正本として保存・indexできます。

- Slack Inbox Import
- iPhone / Android からSlackチャンネルへ投稿した内容を取り込み
- Slack API `conversations.history` をpolling
- Poll Interval は5〜15秒の範囲で設定
- `last_ts` を保存し、新規メッセージだけ取得
- 取得した本文は `brainz_memory/slack/` にMarkdown正本として保存
- Slack履歴自体を正本にせず、BRAINZ側のMarkdownを正本として扱う
- `source_type=slack` として検索可能
- 保存後にAuto Index / Semantic Search / Related Memory / Memory Flowへ連携
- Slack tokenは `data/config/brainz_config.json` に保存し、Git commitしない
- private channel推奨。Botを対象チャンネルへ招待してください
- 外部送信、OpenAI API、自動ChatGPT/Codex送信は行いません

必要なSlack scope例:

- `channels:history`
- `groups:history`
- `channels:read`
- `groups:read`

記憶フォルダ例:

```text
brainz_memory/
├ slack/
├ codex_reports/
├ chatgpt/
└ remote_queue/
```

## Phase 17: Slack Task Parser

補助脳BRAINZ v1.1.0 では、Slack投稿の先頭prefixをtaskとして検出し、Remote Queueへ変換できます。

- Slack Task Parser
- `search:` / `note:` / `handoff_chatgpt:` / `handoff_codex:` / `import:` に対応
- `type:` 形式の簡易taskにも対応
- Slack投稿のraw Markdownは `slack` として保持
- 検出したtaskは `slack_task` としても検索可能
- Slack taskは `remote_queue/processed/` または `remote_queue/failed/` へ振り分け
- note taskは既存Remote Queueと同じく `remote_queue_note` として記憶化
- search / handoff taskは検索欄へqueryを反映し、候補として扱う
- handoff taskでもChatGPT / Codexへ自動送信しない
- Slack ts + task hashで重複処理を避ける

Slack投稿例:

```text
search: quiet workflow
```

```text
note:
補助脳は判断しない。
選択を支える。
```

```text
handoff_chatgpt:
補助脳BRAINZ
Codex Report Auto Import
```

## 熾火 / Slack正本保存

補助脳BRAINZは、検索結果を管理するためだけのUIではなく、忘れていた記憶へ戻るためのUIとして育てています。

- 熾火は「何を探したいかわからない」状態から記憶へ戻る巡回モードです
- Slackはスマホからの投入口として扱い、取得後はローカルMarkdownを正本として保存します
- Slack本文 `message.text` / `ts` / `user` / `channel` / permalinkを正本優先で保持します
- attachments / unfurl は補助情報として保存し、本文の代替にはしません
- `last_ts` 以降のbacklog syncにより、PCスリープ中の投稿も起床後に差分取り込みします
- `#aru` / `#embers` / `#thought` / `#brainz` などの断片も、整理しすぎず熱源として保持します
- Codex完了報告やSlack投稿は要約せず、Markdown正本をBRAINZ内で読めるようにします
- AI解析は現段階では必須ではなく、heat tags / reignition score は将来拡張用のメタデータです
- BRAINZは「忘れを肯定する検索」として、単語一致だけでなく巡り直しを支えます

常駐・タスクトレイ構成は今後の拡張候補です。現段階では既存のwatch / pollingをバックグラウンドで動かし、UIを開いて巡回・確認します。

## Aru Inbox

Aru Inbox は、既存の Slack Inbox とは別の「在る」専用取り込み導線です。

- 既存Slack Inboxは実務 / handoff / note / bookmark / Remote Queue 用として維持します
- Aru Inboxは `#aru` などの在る断片、熱、巡り、未完、火照り、散文素材、創作素材を受け取ります
- Aru Slack Token / Aru Channel ID / Enable Aru Inbox をSlack Inboxとは別に保存します
- `slack_last_ts` とは別に `aru_last_ts` を保持し、PCスリープ中の投稿も起床後に差分取得します
- 投稿本文そのものを `brainz_memory/aru/` にMarkdown正本として保存します
- `source_type=aru` として検索 / Semantic Search / Related Memory / Memory Flow / 熾火モードの対象になります
- Aru側では `note:` や `handoff:` などのprefixは不要です
- prefixが含まれていてもRemote Queue taskとして処理せず、本文正本として保持します
- attachments / unfurl は補助情報として保存し、本文の代替にはしません
- 将来的に記事 / 小説 / 散文素材へ戻すための素材箱として使います

## DAKE_META

```json
{
  "app_key": "DAKE_Brainz_Search",
  "display_name": "補助脳BRAINZ",
  "launcher_title": "補助脳BRAINZ",
  "launcher_description": "ローカルの会話・仕様・メモを忘れず探すための記憶検索アプリです。",
  "site_title": "補助脳BRAINZ",
  "site_description": "ChatGPT、Codex、Claude、Geminiのやり取りをローカル記憶として再接続する検索補助脳です。",
  "update_summary": "QPSC v0.1として記憶庫・取り込み母艦の現在地を正本化しました。",
  "folder_name": "DAKE_Brainz_Search",
  "exe_name": "DakeBrainz_Search.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "draft",
  "show_in_launcher": false,
  "show_on_site": false
}
```

## RELEASE_BODY

```markdown
# 補助脳BRAINZ v1.3.0

ローカルの会話・仕様・Codex実装結果・Slack断片を、記憶へ戻るための正本として保持する補助脳です。

## 追加

- Aru Inbox
- `source_type=aru`
- Aru専用 `last_ts`
- `brainz_memory/aru/` 正本保存
- 熾火モードでのAru優先巡回

## 継続

- 熾火モード
- Slack正本保存の強化
- Slack backlog sync
- `source_type=slack`
- embers_index メタデータ
- BRAINZ内Markdown読書導線
- Slack Inbox Import
- Slack Task Parser
- Codex Report Auto Import
- Notification system
- Semantic Search
- Related Memory
- Memory Flow
- handoff生成
- Ollama / CUDA / GPU状態チェック
```

## 使い方

初回利用者向けの短い運用ガイドとして、`使い方.md` を追加しました。

最初に何をすればいいか、どう運用すると便利かは [使い方.md](使い方.md) を参照してください。

Remote Queue の具体例も [使い方.md](使い方.md) に追記しています。
