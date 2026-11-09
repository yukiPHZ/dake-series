# ChatGPT / Codex Workflow

## 役割

- ChatGPT: 相談、設計、判断整理。
- Codex: 依頼範囲の調査、実装、検証、Git作業。
- 人間: 最終判断、外部公開、販売・決済に関する承認。
- `ORIGINAL.md`: 対象アプリ・商品の真の正本。
- Git: 変更履歴と直前Releaseへ戻るための履歴。

チャットは作業机であり正本ではありません。長大な棚卸しレポートを通常工程にしません。

## 作業開始

通常は次の3つだけを読みます。

1. repoルートの `AGENTS.md`
2. `00_core/DAKE_COMMON_SPEC.md`
3. 対象アプリの `ORIGINAL.md`

対象作業が発生した時だけ、AGENTSのルーティングに従ってbuild、UI、Git、画像、BOOTH、Store、Stripe等の仕様を追加で読みます。

README、DAKE_META、release_body、booth_product、Store、generated JSONは派生ビューです。作業開始時にすべて読まず、ORIGINALの変更が影響するものだけ確認・更新します。

## 実行

1. 対象repo、対象アプリ、依頼レイヤーを確認する。
2. 既存の未commit変更・未追跡ファイルを保護する。
3. 依頼範囲だけを最小差分で変更する。
4. 変更内容に応じた確認だけを行う。
5. 対象差分とfinal `git status` を確認する。

`git add .` を安易に使わず、既存変更を戻す・削除する・stashする・混在commitしません。

## レイヤー

- SHIP: build、生成exeの主要動作、必要なGit作業、GitHub Release。
- DISTRIBUTION: BOOTH、dakeapp.com、store.dakeapp.com。
- COMMERCE: Stripe、Payment Link、`payment_status`、checkout。

SHIP依頼からDISTRIBUTION / COMMERCEへ進みません。販売情報だけの変更でアプリを再buildしません。Stripe変更でBOOTHやスクリーンショットを再確認しません。

既存のStore / StripeツールはDISTRIBUTION / COMMERCE支援として残します。ファイル名に `release` が含まれていてもSHIP判定ツールとして扱いません。

## 外部操作

commit、push、GitHub Release、BOOTH公開、Store公開、Stripe live操作、Cloudflare反映、SNS投稿、価格変更、送信、削除は別の操作です。ユーザーの依頼に含まれない操作を実行しません。

公開結果は実際に確認できた状態だけを報告し、未確認を完了扱いしません。

## 一時作業場所

正式なrepoは `C:\Users\yukiz\devlop\DAKE_series` です。一時worktreeが必要な場合だけ `C:\Users\yukiz\devlop\_worktrees\` を使い、未commit・未push・Human Review待ちの必要作業を勝手に削除しません。

通常の小修正に一時worktree、専用バックアップ、追加レポートを必須化しません。

## 完了報告

次だけを簡潔に報告します。

1. 主要変更ファイル
2. 統合・降格したルール
3. SHIP / DISTRIBUTION / COMMERCEの状態
4. 実施テスト
5. 触らなかったもの
6. 残る明確な課題
7. final `git status`
