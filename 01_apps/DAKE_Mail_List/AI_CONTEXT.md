<!--
AUTO-GENERATED FILE
DO NOT EDIT MANUALLY
Content hash: cae8d6a2882db696
-->

# AI Context: Dakeメールリスト

- Profile: `app`
- Standards version: `1.0.0`

## Purpose

Outlookで選択したメールを直接、または保存済みの `.msg` メールをドラッグ＆ドロップし、会社名・お名前・メールアドレスだけをCSVにする小さな実務ツールです。

## Must Keep

- Outlookで選択した1通または複数メールを、アプリへ直接ドラッグ＆ドロップして処理できる。
- 保存済みの1通または複数 `.msg` も同じドロップ領域で処理できる。
- 出力CSVはUTF-8 with BOMで、`会社名`、`お名前`、`メールアドレス` の3列だけとする。
- 1件の解析失敗で、ほかのメールの処理を止めない。
- 処理はローカルで完結し、メール内容や利用履歴を外部送信しない。

Outlook直接ドラッグ＆ドロップ、複数メール、CSV生成、3項目抽出は会社実機で確認済みです。

## Important Rules

- Outlookから直接メールを落とす操作を、保存操作や別ボタン操作へ置き換えない。
- CSV列を追加・変更しない。
- メールアドレスのlocal-partから氏名を推測しない。
- Outlook COM取得ボタン、監視、履歴、名簿、CRM、AI/API、クラウド機能を追加しない。
- 会社実機確認済みのWindows native OLE入力層は、必要な不具合修正がない限り変更しない。

## Inputs and Outputs

- 対応入力: Outlookで選択したメールのWindows OLE仮想 `.msg`、保存済みの `.msg` ファイル。どちらも複数同時ドロップに対応する。
- 対応出力: UTF-8 with BOMのCSV。列は `会社名`、`お名前`、`メールアドレス` の3列だけとする。
- 非対応: `.msg` 以外のファイル。対象外ファイルは処理せず、アプリを終了させない。
- Outlook仮想メールはWindows native OLE `IDropTarget` で `FileGroupDescriptorW` と、各アイテムの `lIndex` を指定した `FileContents` を取得する。
- 保存済みファイルは同じnative OLE入力層で `CF_HDROP` として受け取る。tkinterdnd2との二重登録は行わない。

## Non-goals

- メール送信、受信監視、自動返信
- Outlook COMで選択メールを取得する別ボタン
- メール一覧、名簿、履歴、CRM、電話番号抽出
- AI/API、クラウド解析、外部保存
- CSV編集画面、設定画面、学習機能

## Quality Check Command

```powershell
python ../../../peakheadz-standards/tools/quality_check.py . --profile app
```

## Source References

- `ORIGINAL.md`
- `../../AGENTS.md`
- `../../00_core/DAKE_COMMON_SPEC.md`
- `../../../peakheadz-standards/common/AI_WORKFLOW.md`
- `../../../peakheadz-standards/common/UX_WRITING.md`
- `../../../peakheadz-standards/common/BRANDING.md`
- `../../../peakheadz-standards/app/QUALITY.md`
- `../../../peakheadz-standards/app/UI.md`
- `../../../peakheadz-standards/app/WINDOWS_APP.md`
- `../../../peakheadz-standards/app/RELEASE.md`
