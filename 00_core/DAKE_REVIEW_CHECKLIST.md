# DAKE_REVIEW_CHECKLIST

通常作業で毎回読む必須文書ではありません。完了前に、変更したレイヤーの項目だけを使います。

## 共通

- [ ] 対象アプリの判断は `ORIGINAL.md` を正本として行った。
- [ ] 変更が影響する派生ビューだけを更新した。
- [ ] 依頼外のレイヤー、既存ユーザー変更、未追跡ファイルへ触れていない。
- [ ] 未確認をPASSにしていない。

## SHIP

- [ ] buildが成功した。
- [ ] 生成された配布exeが起動した。
- [ ] 主機能が成立した。
- [ ] データ破壊等の重大リスクがない。
- [ ] 対象差分だけをcommit / pushした（依頼された場合）。
- [ ] GitHub Releaseが完了した（依頼された場合）。
- [ ] packaging・依存・runtime変更時だけ、開発環境依存を除いた確認を追加した。

BOOTH、画像、Store、Stripe、Cloudflare、SNSはSHIPチェックへ入れません。

## UI変更時だけ

- [ ] 表示崩れ・文字化けがない。
- [ ] 基本フッターは `戸建買取査定 ｜ by Yukihiko Kikuta`。
- [ ] `戸建買取査定` だけがDAKE恒久入口へのリンク。
- [ ] `by Yukihiko Kikuta` は非リンク。
- [ ] Instagram、長いコピーライト、Vibe-Coded、不要な外部リンクを新仕様へ入れていない。
- [ ] 公開表示へ影響する場合だけスクリーンショットを更新した。

## DISTRIBUTION作業時だけ

- [ ] 対象のBOOTH / dakeapp.com / Storeだけを確認した。
- [ ] `DISTRIBUTION = COMPLETE / PARTIAL / PENDING` を独立して記録した。
- [ ] generatedファイルを手編集していない。
- [ ] 公開操作は明示依頼の範囲内。

## COMMERCE作業時だけ

- [ ] `COMMERCE = READY / PENDING / NOT_REQUIRED` を独立して記録した。
- [ ] Stripe Secret、APIキー、Webhook Secretをrepoや生成物へ入れていない。
- [ ] live操作・決済変更は明示依頼の範囲内。

## 完了

- [ ] 変更内容に応じた最小テストを実施した。
- [ ] SHIP / DISTRIBUTION / COMMERCEを混同せず報告した。
- [ ] final `git status` を確認した。
