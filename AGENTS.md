# DAKE Series: Codex作業規則

## 作業前の必読

1. 対象アプリの `ORIGINAL.md`。
2. `00_core/DAKE_COMMON_SPEC.md`（シリーズ共通必須規則の正本）。
3. `00_core/DAKE_BUILD_RULE.md`。
4. `00_core/DAKE_UI_TEXT_RULE.md`。
5. `00_core/DAKE_RELEASE_FLOW.md`。
6. 対象作業に関係するその他の共通仕様、対象README・派生ビュー。

ORIGINALがない既存対象はREADMEを暫定参照し、正本作成の必要性を確認する。個別仕様はORIGINAL、シリーズ必須規則は00_coreで管理する。共通必須規則と個別正本が矛盾したら、旧仕様を無断採用せず、共通規則に合わせて個別文書との整合を取る。

## 新規開発・改修・再ビルド・再出荷

- Windows exeは単独配布・単独起動必須。PyInstallerは原則 `--onefile`。起動高速化のための `--onedir` 化は不可。満たせなければ出荷を保留し、原因・代案を報告する。
- 変更対象アプリの実装を調べ、共通仕様との差分（配布依存・フッター・UI_TEXT等）を修正対象に含める。共通仕様の全文を個別ORIGINALへ複製しない。
- フッター正式文字列は `00_core/DAKE_COMMON_SPEC.md`、定義方法は `00_core/DAKE_UI_TEXT_RULE.md` に従う。共通アイコン・既存リンク・キャッチコピー・配色を勝手に変えない。
- 出荷前に `00_core/DAKE_RELEASE_FLOW.md` と `00_core/DAKE_REVIEW_CHECKLIST.md` のWindows単独exeゲートを実施する。未検証はPASSにしない。
- 過去Releaseや計測値は履歴として保持する。新規則制定だけで既存成果物の適合済み・差し替え済みと扱わない。
- 調査・文書作業だけなら、不要なビルド・GUI起動・成果物生成・出荷を行わない。Web等の非exe成果物にWindows配布規則を強制しない。

## Gitと作業範囲

- 正式ルートは `C:\Users\yukiz\devlop\DAKE_series`。作業前後にbranch・mainとの差分・`git status` を確認する。
- 未commit変更・未追跡ファイルを勝手に戻す、削除する、stashする、commitへ混ぜることは禁止。
- 指定範囲だけ編集し、対象ファイルのみstageする。`git add .` を安易に使わない。
- commit / push / PR / main merge / Release / Productionは別の権限。依頼に含まれない公開や他repo編集をしない。
