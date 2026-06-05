# Phase 5-5 Stripe Payment Link Pilot

## 目的

Advanced Timer用のStripe Payment Linkを`ORIGINAL.md`へ記録し、Store generated JSONとStore本番表示へ反映できることを確認する。

## 対象商品

- `DAKE_Time_AdvancedTimer`
- generated id: `time_advanced_timer`

## ORIGINAL.md 更新内容

- `Store表示用情報` に `Stripe Payment Link` を追加。
- `Store表示用情報` に `Store販売状態: stripe_ready` を追加。
- 既存のBOOTH URL、GitHub Release URL、価格、説明文は維持。

## generated JSON 更新結果

- `time_advanced_timer.stripe_payment_link` に `https://buy.stripe.com/5kQdR9eIh8gfgl25vz0gw00` が入った。
- `time_advanced_timer.payment_status` は `stripe_ready` になった。
- `video_shorts_cut.payment_status` は `preparing` のまま維持した。
- `source_policy` / `do_not_edit` は維持する。

## payment_status 件数

- stripe_ready: 1
- booth_only: 51
- preparing: 1
- Stripe Payment Linkあり: 1
- items: 53

## Stripe Secretを扱っていないこと

このPhaseではStripe Payment Linkのみを扱う。

Stripe Secret Key、APIキー、Webhook Secret、Checkout API、Pages Functions、R2、download_url発行は扱わない。

## 次Phase提案

1. 本番StoreでAdvanced Timerの `Stripeで購入` ボタンを確認する。
2. Stripe Payment Linkを追加する次の商品候補を絞る。
3. Payment Linkを `ORIGINAL.md` へ戻す運用を定着させる。
