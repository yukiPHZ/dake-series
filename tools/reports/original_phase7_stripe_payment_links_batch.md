# Phase 7 Stripe Payment Links Batch

## 目的

Advanced Timerに続き、PDF結合、PDF圧縮、画像リサイズ、SHIMARISU PackへStripe Payment Linkを追加し、DAKE StoreでStripe購入導線を表示できる状態にする。

このPhaseでは、Stripe API、Secret Key、Webhook、Pages Functions、R2、購入後自動配布は扱わない。

## 追加対象

| product | source ORIGINAL | Stripe Payment Link |
|---|---|---|
| DakePDF結合 | `01_apps/DAKE_PDF_Merge/ORIGINAL.md` | `https://buy.stripe.com/9B6fZh0Rr2VV6Ksf690gw01` |
| DakePDF圧縮 | `01_apps/DAKE_PDF_Compress/ORIGINAL.md` | `https://buy.stripe.com/aFa6oHeIh3ZZ6Kse250gw02` |
| Dake画像リサイズ | `01_apps/DAKE_Image_Resize/ORIGINAL.md` | `https://buy.stripe.com/5kQ5kD7fP8gf7Owf690gw03` |
| しまりすくん 実務判断Pack | `C:/Users/yukiz/devlop/SHIMARISU/ORIGINAL.md` | `https://buy.stripe.com/4gM9AT57H7cbc4M3nr0gw04` |

## 更新したORIGINAL.md

- `01_apps/DAKE_PDF_Merge/ORIGINAL.md`
- `01_apps/DAKE_PDF_Compress/ORIGINAL.md`
- `01_apps/DAKE_Image_Resize/ORIGINAL.md`
- `C:/Users/yukiz/devlop/SHIMARISU/ORIGINAL.md`

各ORIGINALのStore表示用情報に以下を追加した。

```text
- Stripe Payment Link: <各商品のPayment Link>
- Store販売状態: stripe_ready
```

既存のBOOTH URL、GitHub Release URL、価格、説明文、Pack情報は削除していない。

## generated JSON 更新結果

`tools/store/generate_store_products.py` で `tools/generated/store_products.generated.json` を再生成した。

検証結果:

- items: 53
- stripe_ready: 5
- booth_only: 47
- preparing: 1
- Stripe Payment Linkあり: 5
- `source_policy` 維持
- `do_not_edit: true` 維持

## payment_status 件数

| payment_status | count |
|---|---:|
| stripe_ready | 5 |
| booth_only | 47 |
| preparing | 1 |

## SHIMARISU Pack反映

SHIMARISU PackのStripe Payment Linkは `SHIMARISU/ORIGINAL.md` に追加する。

`SHIMARISU/ORIGINAL.md` はSHIMARISU Packの真の正本であり、DAKE_series側ではstageしない。

## Stripe Secretを扱っていないこと

このPhaseでは公開済みのStripe Payment Link URLのみを扱う。

以下は扱わない。

- Stripe Secret Key
- Stripe APIキー
- Webhook Secret
- Stripe APIでの商品登録
- Payment Link自動作成
- Pages Functions
- Cloudflare R2
- 購入後自動配布

## 次Phase提案

- Store本番で5商品のStripe購入導線を確認する
- Payment Linkの遷移先と購入後戻り先を商品ごとに確認する
- 次にStripe化する商品候補を小さく選ぶ
- 将来のStripe API利用はSecretの管理場所を決めてから別Phaseで行う