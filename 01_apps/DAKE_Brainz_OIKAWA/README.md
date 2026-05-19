# OIKAWA

OIKAWAは、BRAINZ記憶フォルダを静かに巡回し、熱の痕跡・再出現ワード・関連断片を見つけて、OIKAWA提案MarkdownとしてBRAINZへ戻す観測装置です。

チャットAIではありません。通知AIとして判断はしません。呼ばれた時だけ記憶層を覗き、機械的な巡回結果とBRAINZ状態を静かに返します。

## OIKAWA思想

OIKAWAはBRAINZ本体とは別アプリです。

BRAINZ本体は保存・検索・閲覧の軽さを保ち、OIKAWAは重い巡回や将来的なローカルLLM処理を担います。OIKAWAの出力は再びBRAINZの記憶として保存され、後のOIKAWAがまた読む循環を作ります。

## QPSCでの役割

QPSC（Quiet Personal Cognitive System）では、OIKAWAを検索・原本表示・熾火・熱提案・通知表示の前面UIとして育てます。Phase 1では完全統合せず、BRAINZが書き出す `data/config/qpsc_brainz_status.json` を読み、小さな通知/提案エリアに起床状態を表示します。

OIKAWAの前面表示対象:

- 通常検索
- 検索結果からの原本表示
- 熱検索
- 熾火
- 熱提案
- BRAINZ状態通知
- Slack取り込み後の通知/提案文

通知欄は主役にせず、BRAINZが起きているか、記憶庫が応答しているか、今日の提案があるかを控えめに伝える補助表示として扱います。

```text
菊田
↓
Slack / ChatGPT / Codex / 在る
↓
BRAINZへ保存
↓
OIKAWAが巡回
↓
熱を再発見
↓
OIKAWA提案Markdown生成
↓
BRAINZへ保存
```

## QPSC Phase 2: BRAINZ取り込み通知

Phase 2では、BRAINZが保存する `data/config/qpsc_notifications.json` を読み、OIKAWAのQPSC通知欄に未読通知を最大3件まで表示します。

- BRAINZは取り込みイベントを保存します
- OIKAWAは通知・提案として静かに表示します
- 小さな既読操作で `status` を `read` に変更できます
- `related_path` がある通知は原本へ戻れます
- 原本主義を維持し、全文プレビューではなく導線だけを置きます

## QPSC Phase 3: 原本プレビュー

Phase 3では、OIKAWAに原本プレビューを追加しました。通知の `related_path` や記憶検索の結果から、BRAINZに保存されたMarkdown / TXT原本へ戻れます。

- OIKAWAはBRAINZの記憶を直接編集しません
- 原本主義を維持し、Markdownはテキストとして表示します
- BRAINZは保存とindexを担い、OIKAWAは読む / 戻る前面UIを担います
- OIKAWA上では、記憶を検索する、原本へ戻る、熾火を見る、BRAINZからの通知を受ける役割を短く表示します

## QPSC Phase 4: 熾火・熱検索

Phase 4では、OIKAWAに熾火候補と熱検索入口を追加しました。熾火はAI判断ではなく、BRAINZの通知や最近の取り込みから原本へ戻るための静かな候補表示です。

- 熾火候補は最大3件だけ表示します
- 熱検索は通常検索に未読通知、最近の通知、通知タイトル/本文/source、`related_path` の補正を加えます
- 原本主義を維持し、候補や熱検索結果は `related_path` から原本プレビューへ戻します
- 将来Ollamaで熱判定を補助する可能性がありますが、Phase 4ではルールベースで扱います

## 使い方

1. `python main.py` で起動します。
2. 記憶フォルダが見つからない場合は、画面のボタンからBRAINZ記憶フォルダを選択します。
3. 右下の「巡回する」を押します。
4. 浮上した痕跡と関連断片がカードとして表示されます。
5. 生成されたMarkdownはBRAINZ記憶フォルダ内に保存されます。

初期探索候補は次の順です。

```text
1. 既存BRAINZアプリの data/config/brainz_config.json
2. C:\Users\yukiz\Documents\brainz_memory
3. OIKAWAの oikawa_config.json
```

`oikawa_config.json` はローカル設定ファイルです。Git管理しません。

## 保存先

OIKAWAの出力は次の場所へ保存されます。

```text
brainz_memory\OIKAWA\suggestions\
```

ファイル名は次の形式です。

```text
YYYYMMDD_HHMMSS_oikawa.md
```

## 初期版の観測

初期版はローカルLLMやOpenAI APIを使いません。`.md` と `.txt` を横断スキャンし、ルールベースで熱語を抽出します。

スコアは本文出現、見出し出現、ファイル名出現、直近更新、複数ファイル再出現、近距離共起をもとに計算します。

## 注意事項

- クラウド送信は行いません。
- OpenAI API接続は行いません。
- Slack通知や常時通知は行いません。
- 常時ローカルLLM推論は行いません。
- 5MBを超えるファイルは初期版ではスキップします。
- `build/`、`dist/`、`*.spec`、`*_config.json` はGit管理しません。

## ビルド

```bat
build.bat
```

共通DAKEアイコン `..\..\02_assets\dake_icon.ico` を使い、次の実行ファイルを生成します。

```text
dist\DakeBrainz_OIKAWA.exe
```

## DAKE_META

```json
{
  "app_key": "DAKE_Brainz_OIKAWA",
  "display_name": "OIKAWA",
  "launcher_title": "OIKAWA",
  "launcher_description": "BRAINZの記憶層を巡回し、熱の痕跡を静かに浮かび上がらせます。",
  "site_title": "OIKAWA",
  "site_description": "BRAINZ記憶フォルダを横断し、再出現する熱語と関連断片を観測する補助脳アプリです。",
  "update_summary": "初期実装。Markdown横断スキャン、熱語抽出、OIKAWA提案Markdown保存に対応。",
  "folder_name": "DAKE_Brainz_OIKAWA",
  "exe_name": "DakeBrainz_OIKAWA.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "experimental",
  "show_in_launcher": true,
  "show_on_site": false
}
```

## RELEASE_BODY

```text
OIKAWA 初期版
BRAINZ記憶フォルダを巡回
熱の痕跡と関連断片を観測
提案MarkdownをBRAINZへ保存
Windows向けexe
```
