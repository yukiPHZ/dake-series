<!--
AUTO-GENERATED FILE
DO NOT EDIT MANUALLY
Content hash: d8e26a7046fe9bdc
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
- 2026-11-22に既存onefile build方式で販売exeを再buildし、metadata、launch-check、実GUIを確認した。
- 共通icoをonefile内部へ同梱し、Explorer、GUI titlebar、Windows taskbarの実行中アイコンを実機確認した。
- 新しい3ファイル構成の販売ZIPを検証済み。SHA256は `EEB66EFFD271B8CA847FF598153F633143097B7152BA89FA10A1AF03EB3CE94B`。
- DAKE Storeの現行Stripe Payment Linkを正式販売経路とし、決済後はprivate R2の `desktop/DAKE_Mail_Draft/current.zip` を即時配布する。
- 2026-11-22にStripe本番決済、Store redirect、private R2 download、ZIP download、exe起動を人間が確認した。
- 旧形式Payment Linkは無効化済み。GitHub Releaseの配布exe assetを削除し、公開無料Downloadを終了した。

## Next

- 現行のStripe Product、Price、private-delivery Payment Link、private R2 `current.zip` を維持する。
- BOOTHはDAKE Storeと独立した販売先として維持する。

## Known Issues

- Outlook下書き生成の実機確認には、初期設定済みのWindows版Outlook Classicが必要。
- exeはコード署名していないため、Windowsの確認画面が表示される場合がある。
- Chromeのdownload scan表示は許容し、今回は追加対策を行わない。

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
