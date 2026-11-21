# Dake確認印: Codex作業入口

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

build、release、配布、販売へ進む時はrepoルートのDAKE仕様を追加で確認します。今回の新正本主義実証はsource・正本・必要な検査までで、SHIP / DISTRIBUTION / COMMERCEを含みません。
