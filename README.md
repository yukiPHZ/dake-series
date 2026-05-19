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

## QPSC Phase 1

QPSC（Quiet Personal Cognitive System）は、BRAINZ / OIKAWA / DAKE_Wake_Brainz を静かに分担させるための構成方針です。

- BRAINZ = 深層記憶・取り込み・index・設定
- OIKAWA = 検索・原本表示・熾火・熱提案・通知
- DAKE_Wake_Brainz = 起床・遠隔起動・状態確認レイヤー
- Ollama = ローカル読解補助
- OpenClaw = ローカルエージェント / Gateway候補

Phase 1では完全統合せず、BRAINZが起床状態をローカル状態ファイルへ記録し、OIKAWAが小さな通知欄でその状態を表示します。BRAINZは「見るアプリ」ではなく記憶庫・取り込み母艦へ寄せ、OIKAWAを検索・原本表示・熱検索・熾火・熱提案・通知表示の前面UIとして育てます。

## QPSC Frozen Apps

DAKE_Approve_Brainz はQPSC本線から外し、凍結扱いにします。承認導線はCodex公式モバイル承認 / 遠隔操作導線を優先し、QPSCはBRAINZ / OIKAWA / DAKE_Wake_Brainz / Ollama / OpenClaw補助に集中します。

---

© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta
