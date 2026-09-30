# DakePDF俯瞰名前変更 v1.0.2 正式出荷記録

## Human Review PASS（2026-09-30）

利用者本人がWindows実機でrc2-viewportを合成PDF **800件**で確認し、PASSと報告した。
対象source: caaa213fa70c0141aada4c9bf9ca073c2e780d51。

- 初回順次表示、800件の末尾まで到達
- 小→特大、特大→他サイズの連続切替
- 追加スクロールなしで表示復元、スクロール中の空白化なし、UIフリーズなし
- ページ数・容量表示、名前入力保持
- 末尾付近のrename/Undo、再読み込み後の全件維持

400件を超える今回の実機回帰結果として記録する。Codexが独立実施した99自動テストとは区別する。
過去のFAILと84/99 tests PASSの履歴は維持。Human Review待ちのHOLDは解除。
正式版へのコード差分は候補タイトル接尾辞の除去のみ。rename_core等の動作は変更しない。
価格500円、BOOTH既存商品8798555、Stripe既存Product/Price/Payment Link、stripe_readyを維持する。
Issue #46はmergeだけでcloseせず、配布先・サイト・Cloudflareの実反映まで追跡する。

## 出荷工程

正式版タイトルへの変更後、pytest: **99 passed in 119.91s**。候補からの機能変更なし。

PR #47と最新mainを確認: mergeable CLEAN、チェック未登録、対象はOverviewRename・調査記録・依頼された作業場所運用規則。無関係なアプリの変更なし。
正式build・ZIP・公開反映の結果はIssue #46へ追記する。全工程完了前に正式出荷完了としない。
