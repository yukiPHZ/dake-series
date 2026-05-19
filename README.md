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

---

© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta
