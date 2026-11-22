# Dakeメール下書き: Codex作業入口

このdirectoryの作業では、repoルートの `AGENTS.md` に加えて次を読みます。

1. `ORIGINAL.md` — 製品固有の正本
2. `STATE.md` — 現在地、次、既知事項
3. `../../../peakheadz-standards/common/AI_WORKFLOW.md`
4. 変更内容に関係する `../../../peakheadz-standards/common/` と `../../../peakheadz-standards/app/` の仕様

共通仕様や共通Pythonツールをこのdirectoryへコピーしません。`AI_CONTEXT.md` は `AI_CONTEXT_SOURCES.json` から共通generatorで生成し、直接編集しません。

通常の入口は `AI_CONTEXT.md` です。実装判断で要約だけでは足りない時に、記載されたSource Referencesの正本へ戻ります。

```powershell
python ../../../peakheadz-standards/tools/generate_ai_context.py .
python ../../../peakheadz-standards/tools/quality_check.py . --profile app
```

build、配布ZIP、private R2を扱う時は、repoルートのDAKE build / release仕様も確認します。Stripe、Payment Link、GitHub Release、Store販売導線は、明示された別作業なしに変更しません。
