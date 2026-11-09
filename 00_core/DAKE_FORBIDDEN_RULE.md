# DAKE_FORBIDDEN_RULE

詳細な禁止事項を確認する必要がある時だけ参照します。通常作業の共通禁止は `DAKE_COMMON_SPEC.md` に集約します。

## 増やさない

- 1アプリへ複数目的を入れない。
- 便利そうという理由だけで機能、設定、モード、画面を増やさない。
- 新しい状態管理、DB、manifest、CI、監査、多段承認、バックアップ、ログ基盤を理由なく追加しない。
- About画面やアプリ内テレメトリを追加しない。

## 壊さない

- 元ファイルを不用意に上書きしない。
- 文字化け、起動不能、主機能不成立、データ破壊リスクを放置しない。
- 既存のユーザー変更、未追跡ファイル、過去Release、販売設定を勝手に削除・改変しない。

## 混ぜない

- SHIPへBOOTH、Store、Stripe、Cloudflare、SNSの完了条件を混ぜない。
- DISTRIBUTION / COMMERCEを依頼なしに実行しない。
- README、DAKE_META、release_body、booth_product、Store、generated JSONを正本にしない。
- Stripe Secret、APIキー、Webhook Secret、個人情報、ローカルパスをrepoへ入れない。

## UI

- Instagram、長いコピーライト、Vibe-Coded、不要な外部リンクを新しい共通フッターへ戻さない。
- `by Yukihiko Kikuta` をリンクにしない。
- しまりす不動産を査定・買取業務の主体として表現しない。
