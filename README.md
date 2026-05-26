# DAKE Series

Simple tools for real work.

Download:
https://github.com/yukiPHZ/dake-series/releases

---

Lightweight desktop tools designed to reduce small but real work friction.

Built around the idea:

「現場で止まらない道具」

---

## Principles

- one app = one job
- fast to launch
- easy to understand
- local-first
- less is better

---

## Structure

```text
00_core/
01_apps/
02_assets/
03_docs/
```

---

## QPSC v0.1 現在地

QPSC（Quiet Personal Cognitive System）は、記憶を保存し、静かに戻れる入口を置くための構成です。判断は代行せず、最後に選ぶのは菊田です。

- BRAINZ = 記憶庫 / 取り込み母艦
- OIKAWA = 検索 / 原本表示 / 通知 / 熾火 / ORBIT
- DAKE_Wake_Brainz = 起床 / 状態確認レイヤー
- Ollama = 将来のローカル読解補助
- OpenClaw = 将来のGateway / エージェント補助候補
- DAKE_Approve_Brainz = frozen

```text
Slack / Paste / ChatGPT Export / Codex Report
↓
BRAINZ = 記憶庫・取り込み母艦
↓
qpsc_notifications / qpsc_status
↓
OIKAWA = 通知・検索・原本表示・熾火・ORBIT
↓
菊田が選ぶ
```

### v0.1でできること

- BRAINZ起動状態を記録できる
- BRAINZ heartbeatを更新できる
- 取り込み通知を保存できる
- OIKAWAで通知を見られる
- OIKAWAで原本プレビューできる
- 熾火候補を表示できる
- 熱検索入口がある
- ORBITで今日の整理が見られる
- BRAINZからOIKAWAを開ける
- BRAINZは母艦UIへ整理済み

### v0.1の固定方針

QPSC v0.1では、BRAINZが保存・取り込み・index・設定・状態を担い、OIKAWAが検索・原本表示・通知・熾火・ORBITを担います。DAKE_Wake_Brainzは起床と状態確認に集中し、DAKE_Approve_BrainzはQPSC本線から外して凍結します。

承認導線はCodex公式モバイル承認 / 遠隔操作導線を優先します。OllamaとOpenClawはまだ本線へ入れず、将来の補助候補として扱います。

## QPSC正本ルート: PEAKHEADZ_ROOT

QPSCの正本ルートは、今後 `C:\Users\yukiz\Documents\PEAKHEADZ_ROOT` へ寄せます。

- 既存の `brainz_memory` はlegacy扱いです
- 既存configが `brainz_memory` や別フォルダを指している場合は勝手に変更しません
- 初回作成時は `PEAKHEADZ_ROOT` にPhase 20の標準フォルダ構成を作れます
- 移行は安全優先です。rename / delete は行わず、必要な場合もcopy候補として扱います
- Obsidianでは `PEAKHEADZ_ROOT` を保管庫として開きます

Obsidianは本体ではなくMarkdown観測UIです。Obsidianを消しても、正本Markdown、README、release_body、Git上の記録は残ります。

## QPSC v0.3 方向整理 / Phase 20

QPSC v0.3では、BRAINZを「見るアプリ」から外し、ローカルPCへ正本を受け止める母艦として固定します。

- Slack = 超低摩擦の投入口
- BRAINZ = 保存母艦。取り込み、設定、状態、OIKAWA起動導線を担います
- OIKAWA = 巡回、熾火、正本ニュース、今日の戻りを扱う前面端末
- OpenClaw = 将来の巡回生物。README巡回、Git差分確認、正本ニュース生成を静かに担う候補です
- Codex = 呼び出す職人。実装、修正、Git、README更新を高密度に行います
- Obsidian = ローカルMarkdown世界の観測UI。graph、backlink、preview、searchのビュー層です

正本はMarkdown、README、release_body、BRAINZ、Gitに残します。Obsidianを消しても正本は残り、OIKAWAはOpenClawの便りを受け取る顔として育てます。

## QPSC v0.2 現在地

QPSC v0.2では、取り込みから原本へ戻る導線に加えて、OIKAWA側で「熱」「巡回」「側に在る」を扱える現在地まで固定します。QPSCは通知アプリでもタスク管理アプリでもなく、忘れていいための記憶庫です。判断は代行せず、AIは熱の気配を小さく添えるだけにします。

- BRAINZ = 記憶庫 / 取り込み母艦 / 状態 / 設定
- OIKAWA = 原本表示 / 通知 / 熾火 / ORBIT / 巡回 / 側に在る
- Ollama = 熱補助のみ。判断代行しない
- DAKE_Wake_Brainz = 起床 / 状態確認レイヤー
- OpenClaw = 将来のGateway候補。まだ本線未導入
- Slack通知本格化 = まだ先。まずOIKAWA内部循環を優先

### v0.2で追加されたもの

- QPSC状態確認
- ChatGPT export新形式対応
- 巨大会話chunk分割
- Slack通知文の整流
- Codex報告の正本入口化
- OIKAWAの黒基調 / 日本語可読性改善
- OIKAWAウインドウ外観BRAINZ寄せ
- Ollama熱補助
- 巡回 / 再訪
- 側に在る

### v0.2でできること

- ChatGPT exportをzip / フォルダ / 分割JSONで取り込める
- 巨大会話をchunk分割して扱える
- Slack投稿を記憶入口として通知できる
- Codex報告を正本として保存し、OIKAWAから戻れる
- OIKAWAで原本を中央に表示できる
- 熾火候補を表示できる
- Ollamaで熱の気配を小さく添えられる
- 何度も戻る記憶を巡回として扱える
- 側に在る記憶を静かに浮かべられる
- QPSC状態確認ができる

### v0.2思想メモ

OIKAWAは、必要な記憶が静かに側へ戻ってくる場所です。QPSCは「これをやるべき」と言わず、原本へ戻れる入口だけを置きます。

### 今後候補

- Phase 18: 実使用レビュー
- Phase 19: OIKAWA表示密度調整
- Phase 20: BRAINZ接続 / 設定ページ
- Phase 21: Slack通知本格化
- Phase 22: OpenClaw Gateway検証
- Phase 23: Wake / スマホ起動連携整理

---

© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta

## QPSC通常運用 / Phase 22

通常運用では、QPSCを次の流れで扱います。

```text
Slackへ投げる
↓
BRAINZが受け止める
↓
PEAKHEADZ_ROOTへMarkdown正本保存
↓
OIKAWAで再会する
↓
Obsidianで観測する
```

- Slack = 投入口
- BRAINZ = 保存母艦 / 設定 / 監視
- PEAKHEADZ_ROOT = 正本保管庫
- OIKAWA = 熾火 / 巡回 / 側に在る / 熱補助
- Obsidian = Markdown観測UI
- Codex = 実装職人
- OpenClaw = 将来の巡回生物

`PEAKHEADZ_ROOT/slack/` はlegacy保存先です。削除やrenameはせず、新規Slack保存は `PEAKHEADZ_ROOT/10_slack/*` へ寄せます。

## QPSC Phase 23 UI最終整流

- BRAINZは1カラムの母艦UIとして、保存・設定・監視に集中する。
- OIKAWAは全文プレビューではなく、熾火・ORBIT・巡回・側に在るを並べる巡回端末として扱う。
- PEAKHEADZ_ROOTが正本保管庫で、ObsidianはMarkdown観測UIとして開く。
- 旧 `PEAKHEADZ_ROOT/slack/` はlegacyであり、新規保存先は `10_slack/*` に寄せる。

## QPSC Phase 24 通常運用UX

- Slackへ投げるだけで、BRAINZが `PEAKHEADZ_ROOT` へMarkdown正本として保存する流れを前面に出した。
- BRAINZはSlack接続と5チャンネル保存設定を、通常運用者が読めるカードUIへ寄せる。
- OIKAWAは正本ニュース端末として、全文表示よりも熾火・巡回・側に在る・原本を開く導線を優先する。
- 旧 `slack/` はlegacyとして残し、新規保存は `10_slack/*` に寄せる。

## QPSC OIKAWA最小巡回UI

- OIKAWAは読む場所ではなく、正本へ戻るきっかけを置く端末として整理する。
- 正本ニュース、熾火、巡回、側に在るを前面に置き、詳細ログや数字の羅列は主役にしない。
- 原本はOIKAWA内で抱え込まず、ローカルMarkdownを直接開く。

## QPSC読み取り補助

Slackへ投げた正本は、BRAINZでMarkdown保存後に軽く読まれ、タグ・熱・BORINEF候補がfrontmatterへ補われます。

- 分類は管理ではなく、戻るための弱い補助です。
- OIKAWAはそのメタを熾火、巡回、側に在る、正本ニュースへ静かに混ぜます。
- Obsidianでは `PEAKHEADZ_ROOT` のMarkdown正本をそのまま観測できます。

## QPSC v0.4 現在地

QPSCは、熱を生成するシステムではありません。

```text
QPSC = 熱へ戻る導線
```

普通の管理思想は、忘れるな、継続しろ、止まるな、TODOを消化しろ、になりがちです。QPSCは違います。

```text
忘れていい。
止まっていい。
離れていい。

でも、戻れるようにはしておく。
```

### OIKAWAは再接続要求である

OIKAWAは常駐OSではなく、通知アプリでも、TODO管理でも、AIダッシュボードでもありません。

OIKAWAは、補助脳BRAINZへ接続するための生体アクセスキーです。

```text
OIKAWAを開く = 再接続要求
```

熱がある時、OIKAWAはいりません。すでに流れているからです。熱が冷めた時、止まった時、何をしていたか分からなくなった時に、OIKAWAを開きます。

OIKAWAは「頑張れ」と言いません。ただ、まだ残っている線を返します。

### BRAINZは戻れるように残す

BRAINZは記録そのものではありません。熱が戻れるように、Slack断片、Codexログ、ChatGPT export、note、README、Git、実務メモ、BORINEF断片を正本として残す母艦です。

QPSCの救いは、忘れてもいい、でも戻れる、という構造にあります。
