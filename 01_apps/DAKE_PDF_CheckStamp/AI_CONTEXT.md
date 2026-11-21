<!--
AUTO-GENERATED FILE
DO NOT EDIT MANUALLY
Content hash: a1819f0220fbdaf5
-->

# AI Context: Dake確認印

- Profile: `app`
- Standards version: `1.0.0`

## Purpose

PDFプレビュー上で位置を選び、名字と日付入りの確認印を1か所押せるWindows向けアプリです。

## Must Keep

- PDFを読み込み、preview上で位置を選び、名字と日付入りの確認印を1か所へ反映して別名保存する流れを維持する。
- previewと保存PDFで、対象ページ・位置・名字・日付・サイズを一致させる。
- 元PDFを上書きせず、保存先が重なる場合は別名を選ぶ。
- 古いpreview処理結果を、後から現在のPDFやページへ表示しない。
- UI threadをPDF描画で止めず、終了時にworkerとtimerを停止する。

## Important Rules

- 単機能の確認印アプリを、多機能なPDF編集・電子署名アプリへ変えない。
- 電子署名、電子契約、本人性証明を行う機能であると表現しない。
- 確認印の位置指定と保存結果の対応を崩さない。
- 既存のPDF処理、preview、保存、drag and dropの挙動を、文章・ブランド修正のついでに変更しない。

## Current State

- 既存のPDF読込、preview、確認印位置指定、保存処理を維持する。
- 新正本主義の実証対象として `common + app + 製品固有` の参照構造を適用済み。
- 利用者向け旧クレジットを確定表記へ揃え、エラー見出しを穏やかな案内へ変更済み。
- SHIP / DISTRIBUTION / COMMERCEは今回の対象外。

## Next

- 今回のPhase 5完了後は停止する。次の製品やWebへ自動展開しない。

## Known Issues

- build済みexeへの反映、実機GUI、タスクバー表示は今回未確認。source変更の出荷時に別途確認する。
- 既存依存のPyMuPDFを含む配布・ライセンス状態は今回変更せず、再評価もしない。

## Quality Check Command

```powershell
python ../../../peakheadz-standards/tools/quality_check.py . --profile app
```

## Source References

- `AGENTS.md`
- `ORIGINAL.md`
- `STATE.md`
- `../../00_core/DAKE_COMMON_SPEC.md`
- `../../../peakheadz-standards/common/AI_WORKFLOW.md`
- `../../../peakheadz-standards/common/UX_WRITING.md`
- `../../../peakheadz-standards/common/BRANDING.md`
- `../../../peakheadz-standards/app/QUALITY.md`
- `../../../peakheadz-standards/app/UI.md`
- `../../../peakheadz-standards/app/WINDOWS_APP.md`
