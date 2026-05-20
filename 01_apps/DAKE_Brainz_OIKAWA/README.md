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

## QPSC v0.1 正本

OIKAWAはQPSC v0.1で、BRAINZに保存された記憶を探して、戻って、読む前面UIとして固定します。

- BRAINZ状態と取り込み通知を表示します
- 通知、検索結果、熾火候補、ORBIT候補から原本プレビューへ戻れます
- 熾火候補はAI判断ではなく、戻るきっかけとして扱います
- 熱検索入口はルールベースで、未読通知・最近通知・関連パスを優先します
- ORBITは今日の整理を短く表示し、命令ではなく次に触れられる候補を置きます
- 原本へ戻った履歴は `data/config/qpsc_recent_returns.json` に保存しますが、Git管理しません

OIKAWAはBRAINZの記憶を直接編集しません。原本主義を維持し、判断は菊田に残します。

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

## QPSC Phase 5: 今日の整理 / ORBIT

Phase 5では、OIKAWAに「今日の整理 / ORBIT」を追加しました。ORBITは判断代行ではなく、BRAINZ通知、熾火候補、検索語、起床状態から今日の流れを静かに整理する表示です。

- 今日取り込まれた記憶、未読通知、熾火候補、BRAINZ awake状態を短く表示します
- 今日の流れは最大5行、次の候補は最大3件に抑えます
- 次の候補は命令ではなく、戻れる入口として扱います
- 整理結果は `data/config/qpsc_orbit_today.json` に保存しますが、Git管理しません
- 将来Ollamaで要約補助する可能性がありますが、Phase 5ではルールベースで扱います

## QPSC Phase 6: OIKAWA UI整流

Phase 6では、原本プレビューを中央メインに移し、OIKAWAをBRAINZの記憶を探して、戻って、読む場所として整理しました。

- 左カラムはタイトル、補助脳BRAINZ、状態、記憶フォルダ、検索入口、検索結果を置きます
- 中央メインは原本プレビューを最優先し、Markdown / TXTを読む領域として扱います
- 右カラムはQPSC通知、熾火、今日の整理 / ORBIT、次の候補をサブ情報としてまとめます
- フッターはQPSC表記に変更し、BRAINZにある記憶をOIKAWAが呼び戻す文脈を表示します

### タイトルバーについて

Tkinter標準タイトルバーの黒化はOS依存です。OIKAWAではWindows環境で可能な範囲だけ標準タイトルバーのdark化を試み、失敗してもアプリ動作を優先します。独自タイトルバー化はまだ行いません。将来、独自タイトルバー化すればより一体感のある黒い枠にできる可能性はありますが、ドラッグ移動、最小化、閉じる、リサイズ実装が必要です。

## QPSC Patch: OIKAWAウインドウ外観をBRAINZ寄せ

OIKAWAのウインドウ外観をBRAINZと並べても違和感が少ない黒基調へ寄せました。

- 標準タイトルバーdark化はWindows環境で可能な範囲だけ適用します
- 失敗しても起動と操作を優先します
- 独自タイトルバー化は行わず、安定性を優先します

## QPSC Phase 7: 通知整流 / 未読循環

Phase 7では、通知をログではなく循環として扱うため、OIKAWA側の通知表示を整流しました。

- 通知は未読、熾火関連、今日の通知、最近通知、古い既読通知の順に扱います
- readかつ古い通知、またはreadかつ原本導線のない通知は沈降し、初期表示から外します
- 熾火を通知より少し優先し、右カラム上部で静かに表示します
- ORBITには今日静かになった通知数、熾火候補数、最近戻った原本数を反映します
- 原本へ戻った履歴は `data/config/qpsc_recent_returns.json` に保存しますが、Git管理しません

## QPSC Phase 8: BRAINZからOIKAWAへ

Phase 8では、BRAINZ側にOIKAWAを開く導線を追加しました。BRAINZは取り込み母艦として状態と設定に集中し、検索・原本表示・熾火・ORBITはOIKAWA側で受け持ちます。

## QPSC Phase 10: QPSC状態確認

Phase 10では、OIKAWAに小さなQPSC状態確認を追加しました。実使用前にBRAINZ / OIKAWA連携の現在地を軽く確認できます。

- BRAINZ awake、通知ファイル、原本プレビュー、ORBIT、最近戻った原本数を短く表示します
- 状態確認は詳細ログではなく、安心確認として扱います
- JSONが未作成でも落とさず、まだ記録がない状態として静かに表示します
- `related_path` がある通知から、最低1件の原本へ戻れるかを確認します

## QPSC Phase 11: ChatGPT export取り込み後

Phase 11では、BRAINZ側でChatGPT exportの取り込み入口を簡素化しました。

- OIKAWAは取り込み後の通知を `chatgpt_export` として受け取ります
- ChatGPT exportの取り込み通知は未読通知、熾火、ORBITの候補に反映されます
- `related_path` がある場合は、BRAINZ側の取り込みログや原本導線へ戻れます
- 取り込まれた記憶は通常検索・熱検索から探し、原本プレビューで確認します

## QPSC UI表示方針

QPSCのOIKAWA / BRAINZは黒基調の静かなUIを基本にします。

- 背景は黒〜ダークグレーを優先します
- 日本語本文は `Yu Gothic UI` / `Meiryo` を優先します
- 小さすぎる日本語や太字の多用は避けます
- 原本プレビュー、通知、熾火、ORBITは読みやすさを優先します

## QPSC Phase 12: Slack通知の整流

Phase 12では、Slack由来の通知をログではなく、記憶へ戻るための静かな入口として表示します。

- Slack通知はタイトルと短いmessageを中心に表示します
- 古いSlack通知も、OIKAWA側では静かな文言へ寄せて表示します
- Slack通知は熾火とORBITにも流れ、今日の流れでは自然な言葉で扱います
- `related_path` がある通知は、保存済みMarkdownへ戻れます

## QPSC Phase 13: Codex報告の正本入口化

Phase 13では、Codex由来の通知を正本へ戻る入口として表示します。

- 古いCodex通知も、OIKAWA側では静かな正本入口の文言へ寄せます
- Codex報告は通知、熾火、ORBITの候補から原本プレビューへ戻れます
- `related_path` があるCodex報告は、熾火候補として少し浮上しやすくします
- OIKAWAはCodex報告を編集せず、読む場所として扱います

## QPSC Phase 14: Ollama熱補助

Phase 14では、熾火候補に対してOllamaの弱い熱補助を追加しました。

- AIは判断代行ではなく、熾火候補の補助として扱います
- 「熱を読む」操作で1件だけOllamaに短い断片を渡します
- 原本全文は送らず、title / message / source / related_path / 原本先頭だけを読みます
- Ollama未起動でもQPSCは動き、熱補助は眠っている状態として表示します
- 結果は `data/config/qpsc_heat_hints.json` に保存しますが、Git管理しません

## QPSC Phase 15: 巡回 / 再訪

Phase 15では、OIKAWAで原本へ戻った履歴を巡回ログとして保存し、何度も戻っている記憶を小さな再訪候補として表示します。

- 巡回ログは `data/config/qpsc_revisit_log.json` に保存しますが、Git管理しません
- 巡回候補は「おすすめ」ではなく、戻ってしまう記憶への静かな入口として扱います
- スコアは最近開いた、何度か開いた、深夜に戻った、熾火候補だった、熱補助がある、を弱く加点します
- ORBITには今日の巡回数、最近戻った記憶数、側に残っている記憶数を反映します
- AIは候補を決めず、Ollama熱補助がある場合だけ少し加点します

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
  "update_summary": "QPSC v0.1として検索・原本表示・通知・熾火・ORBITの前面UIを正本化しました。",
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
