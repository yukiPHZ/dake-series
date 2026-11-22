<!--
AUTO-GENERATED FILE
DO NOT EDIT MANUALLY
Content hash: 8e2b9abbb4f5fd7d
-->

# AI Context: Dakeメール下書き

- Profile: `app`
- Standards version: `1.0.0`

## Purpose

CSV名簿からOutlookの個別下書きを作る、送信しない実務補助ツールです。

## Must Keep

- CSVを読み込み、会社名・名前・メールアドレスに使う列を利用者が選択できる。
- 件名・本文テンプレートへ、会社名・名前・メールアドレスを差し込む。
- Outlook Classicの個別メールを下書きとして保存し、自動送信しない。
- Outlook側で生成された既定署名HTMLを可能な限り維持し、その前へ本文を挿入する。
- ローカルファイルを複数添付でき、存在しない添付は生成前に停止する。
- 5 / 10 / 20 / 50 / 100 / 全件の生成件数制限を維持する。
- 生成前に先頭1件の宛先・件名・本文をプレビューする。
- 空メール、不正形式、重複、作成成功、エラーをCSVレポートへ記録する。
- 利用者がOutlook上で宛先・本文・添付を確認してから、自分の判断で送信する流れを維持する。

## Important Rules

- Outlook下書き作成補助を、自動送信・SMTP送信・営業メール配信ツールへ変えない。
- `CreateItem`、既定署名の読込、本文挿入、`Save` による下書き保存の流れを、ブランド修正のついでに変更しない。
- CSV列選択、テンプレート差込、添付、件数制限、プレビュー、レポートの既存挙動を崩さない。
- CSV、テンプレート、添付、メール内容、利用履歴を外部送信しない。
- Outlook Classicを必要とする現在の前提を、未検証の別方式へ置き換えない。

## Current State

- CSV名簿からOutlook Classicの個別下書きを作る既存機能を維持する。
- Stripe決済後にprivate R2の `desktop/DAKE_Mail_Draft/current.zip` を配布する本番導線が稼働している。
- 新正本主義の `common + app + 製品固有` 参照構造を適用する。
- 利用者向けクレジットは `© Yukihiko Kikuta / PEAKHEADZ` に統一する。
- 2026-11-22に既存build方式で販売exeを再buildし、metadata、共通icon、launch-check、実GUIを確認した。
- 新しい3ファイル構成の販売ZIPを検証済み。SHA256は `06DDA686FEFDB2C3950DF781ECC3323BDDDF677357F11A8DD547AD3E736DC8F5`。

## Next

- private R2の `current.zip` を検証済み販売ZIPに揃え、人間がStore購入履歴から再ダウンロードし、ZIP展開とexe起動を確認する。
- 再ダウンロード確認前は、旧Payment Link、GitHub Release、無料Download導線を変更しない。

## Known Issues

- Outlook下書き生成の実機確認には、初期設定済みのWindows版Outlook Classicが必要。
- exeはコード署名していないため、Windowsの確認画面が表示される場合がある。
- GitHub Releaseの既存無料配布は、人間の再ダウンロード確認が終わるまで維持する。

## Quality Check Command

```powershell
python ../../../peakheadz-standards/tools/quality_check.py . --profile app
```

## Source References

- `AGENTS.md`
- `ORIGINAL.md`
- `STATE.md`
- `../../00_core/DAKE_COMMON_SPEC.md`
- `../../00_core/DAKE_BUILD_RULE.md`
- `../../00_core/DAKE_RELEASE_FLOW.md`
- `../../../peakheadz-standards/common/AI_WORKFLOW.md`
- `../../../peakheadz-standards/common/UX_WRITING.md`
- `../../../peakheadz-standards/common/BRANDING.md`
- `../../../peakheadz-standards/app/QUALITY.md`
- `../../../peakheadz-standards/app/RELEASE.md`
- `../../../peakheadz-standards/app/UI.md`
- `../../../peakheadz-standards/app/WINDOWS_APP.md`
