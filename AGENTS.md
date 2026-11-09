# DAKE Series: Codex作業入口

## 通常作業で最初に読むもの

原則、次の3つだけを読む。

1. この `AGENTS.md`
2. `00_core/DAKE_COMMON_SPEC.md`
3. 対象アプリの `ORIGINAL.md`

`ORIGINAL.md` がないLegacyアプリは、次回そのアプリを修正・出荷する時だけREADMEを暫定参照し、新仕様へ段階移行する。README、DAKE_META、release_body、booth_product、Store、generated JSONは派生ビューであり、作業開始時に全件確認しない。変更が影響するビューだけを確認・更新する。

## 条件付き参照

- build方式、依存、PyInstaller、Windowsメタデータを変える: `DAKE_BUILD_RULE.md`
- SHIP、GitHub Release、出荷判定: `DAKE_RELEASE_FLOW.md`
- UI文言・共通フッターそのものを変える: `DAKE_UI_TEXT_RULE.md` / `DAKE_UI_GUIDE.md`
- Git除外・成果物の扱いを変える: `DAKE_GIT_RULE.md`
- アイコンを変える: `DAKE_ICON_RULE.md`
- UI変更で画像更新が必要: `DAKE_SCREENSHOT_RULE.md`
- BOOTH作業: `DAKE_BOOTH_REGISTER_SPEC.md`
- Store作業: `DAKE_STORE_OPERATION_RULE.md` / `DAKE_STORE_GENERATED_SPEC.md`
- Stripe・決済作業: `tools/commerce/` と対象Commerce仕様

関連しない仕様を「念のため」一括読込・一括更新しない。

## 作業境界

- 依頼されたレイヤーだけを扱う。SHIP依頼からBOOTH、Store、Stripe、Cloudflare、SNS、販売文章、価格、画像生成へ自動的に進まない。
- 公開、送信、決済作成・変更、BOOTH公開、価格変更、外部登録、削除は、人間の明示依頼なしに実行しない。
- 未commit変更、未追跡ファイル、過去Release、既存設定を勝手に戻す・削除・stash・混在commitしない。
- 新機能、全面リファクタリング、新しい状態管理・DB・CI・監査・承認・バックアップ・ログ基盤を理由なく追加しない。削除・統合・単純化を優先する。
- `commit`、`push`、GitHub Release、販売・公開は別権限。依頼範囲を越えない。

## 完了時

変更内容に応じた最小テストを行い、対象差分とfinal `git status` を確認する。未確認をPASSにせず、SHIP / DISTRIBUTION / COMMERCEを混ぜずに報告する。
