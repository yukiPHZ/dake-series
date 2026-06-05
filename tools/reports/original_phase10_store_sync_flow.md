# Phase 10 DAKE Store Sync Flow

## 目的

DAKE Storeの商品情報を、各商品の `ORIGINAL.md` から生成される `store_products.generated.json` 経由で `dake-store-site` へ迷わず同期できるようにする。

Storeは正本ではない。正本は各商品の `ORIGINAL.md` であり、`store_products.generated.json` はStore表示用の生成物である。

## 作成したスクリプト

- `tools/store/sync_store_to_site.py`
- `tools/store/sync_store_to_site.bat`

Pythonスクリプトを主とし、batはWindowsから短く呼び出すための補助ファイルとする。

## 同期対象

- 生成元: `C:\Users\yukiz\devlop\DAKE_series\tools\generated\store_products.generated.json`
- 同期先: `C:\Users\yukiz\devlop\dake-store-site\public\assets\data\store_products.generated.json`

`--store-site` でStore repoのパスを指定できる。

```powershell
python tools\store\sync_store_to_site.py --store-site C:\Users\yukiz\devlop\dake-store-site
```

## 同期フロー

1. `tools/store/generate_store_products.py` を実行してStore generated JSONを再生成する。
2. JSONとして読み込み、必須メタ情報を検証する。
3. `generated_at` を除いた意味的な差分を確認する。
4. 意味的な差分がない場合は、タイムスタンプだけの差分を避けるため既存JSONを復元する。
5. `dake-store-site/public/assets/data/store_products.generated.json` へ同期する。
6. 同期後JSONを再検証する。
7. 必要な次のgit操作を表示する。

## 検証項目

- items件数
- type別件数
- payment_status別件数
- stripe_payment_linkあり件数
- booth_urlあり件数
- preparing件数
- `source_policy` の存在
- `do_not_edit: true`
- `shimarisu_pack` の存在
- JSONとして有効であること

## 実行結果

Phase 10初回実行では、generated JSONの意味的な差分はなかった。

そのため、DAKE_series側の `tools/generated/store_products.generated.json` はタイムスタンプだけの差分を避けるため既存内容を維持した。

Store側JSONも既存内容と一致しており、同期による差分は発生しなかった。

## payment_status件数

初回実行時の確認結果。

- items: 53
- app: 50
- pack: 2
- shimarisu_pack: 1
- stripe_ready: 5
- booth_only: 47
- preparing: 1
- stripe_payment_linkあり: 5
- booth_urlあり: 52

## 今後のDAKE出荷時の使い方

1. 商品の `ORIGINAL.md` を更新する。
2. DAKE_seriesで同期スクリプトを実行する。

```powershell
cd C:\Users\yukiz\devlop\DAKE_series
python tools\store\sync_store_to_site.py
```

3. 表示された件数と差分を確認する。
4. DAKE_series側に差分がある場合のみ、生成JSONをcommit / pushする。
5. dake-store-site側に差分がある場合のみ、同期JSONをcommit / pushする。
6. Cloudflare Pages反映後、本番の `/assets/data/store_products.generated.json` と代表商品ページを確認する。

## 注意点

- 商品情報、価格、Stripe URL、BOOTH URLはStore側で手編集しない。
- `store_products.generated.json` は正本ではなく、手編集禁止の生成物である。
- 同期スクリプトは自動commit / pushを行わない。
- 複数repoをまたぐため、人間確認を挟んでからgit操作する。
- `generated_at` だけの差分はcommit対象にしない。
- Store UI、Stripe API、Webhook、R2はこの同期フローの対象外である。

## 今回やらなかったこと

- DAKE出荷定義の正式更新
- プロジェクト情報源の更新
- Stripe API自動登録
- Stripe Payment Link自動作成
- Stripe Secretの利用
- Checkout API
- Webhook
- R2
- Store UI変更
- 商品情報の手編集
- 価格変更
- 全商品Stripe化
