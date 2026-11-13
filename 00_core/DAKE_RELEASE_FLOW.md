# DAKE_RELEASE_FLOW

SHIP / DISTRIBUTION / COMMERCEを別々に扱うための最小フローです。共通原則は `DAKE_COMMON_SPEC.md` を正本とします。

## 1. SHIP

```text
対象ORIGINALを確認
↓
変更を実装
↓
build
↓
生成された配布exeで起動・主要機能確認
↓
対象差分だけcommit / push
↓
GitHub Releaseへ成果物を公開
↓
SHIP = COMPLETE
```

GitHub Releaseが完了すればSHIPは完了です。DISTRIBUTION / COMMERCEが未完了でもSHIPをPENDINGへ戻しません。

### SHIP停止条件

原則、次だけです。

1. build失敗
2. exeが起動しない
3. 主機能が成立しない
4. データ破壊等の重大リスクが確認された

BOOTH未登録、スクリーンショット未更新、Store未同期、Stripe未確認、Cloudflare未確認、SNS未投稿、販売説明文未更新はSHIP停止条件ではありません。

### SHIPの確認量

- 通常の小修正: build、生成exeの主要動作、GitHub Release。
- packaging・PyInstaller・依存・DLL・runtime変更: 開発環境依存を除いたWindows確認を追加。
- UI変更: 必要な表示確認。公開画像へ影響する場合だけスクリーンショット更新。
- 文書・仕様だけの変更: build、GUI、Releaseを実施したと扱わない。

通常は生成された配布exeを一度確認します。空フォルダ、ZIP展開後、別PCで同じ操作を重ねるのはpackaging関連変更時だけです。

GitHub Releaseの過去版は原則保持します。patch / minor / majorは `DAKE_COMMON_SPEC.md` の簡潔な定義に従います。

## 2. DISTRIBUTION

対象:

- BOOTH
- dakeapp.com
- store.dakeapp.com

状態:

- `DISTRIBUTION = COMPLETE`
- `DISTRIBUTION = PARTIAL`
- `DISTRIBUTION = PENDING`

販売・配布先への掲載を依頼された時だけ、影響する派生ビューと対象サービスを確認します。BOOTH、Store、Cloudflareの未完了はSHIPへ影響させません。

### BOOTH

```text
ORIGINAL
↓
booth_ready生成・確認
↓
BOOTH_PREP = READY
↓
Codex STOP
↓
DakeBOOTHアシスト + 人間確認
↓
人間がBOOTH登録・公開
```

CodexはBOOTH Web UIを操作しません。登録後のURLは、ユーザーがURLを明示して記録を依頼した時だけ `tools/record_booth_registration.py` でORIGINALへ記録します。

### dakeapp.com

ORIGINAL由来の掲載用派生データを `tools/sync_dakeapp_apps_json.py` で生成・同期します。BOOTH URLがなくても、GitHub Release等の掲載情報で成立する商品は独立して掲載準備できます。自動commit / push / deployは行いません。

### store.dakeapp.com

ORIGINAL由来の `store_products.generated.json` を `tools/store/sync_store_to_site.py` で生成・同期・検証します。自動commit / push / deployは行いません。BOOTHとStripeは商品ごとの購入導線であり、Store掲載そのものの完了条件ではありません。

### 販売準備の完了報告

通常の販売準備を依頼された場合、Codexは各導線の素材・派生データを準備して次のように独立報告し、外部登録・公開へ進まず終了します。

```text
BOOTH_PREP = READY または PENDING
DAKEAPP = READY または PENDING
STORE = READY または PENDING
```

BOOTH登録後のURL記録や再同期は別依頼です。Stripe状態はCOMMERCEとして別に報告します。

## 3. COMMERCE

対象:

- Stripe
- Payment Link
- `payment_status`
- checkout確認

状態:

- `COMMERCE = READY`
- `COMMERCE = PENDING`
- `COMMERCE = NOT_REQUIRED`

無料アプリや決済不要対象では `NOT_REQUIRED` が正常です。Stripe live操作、決済作成・変更は人間の明示依頼と個別確認なしに実行しません。

## 4. 正本と派生ビュー

判断は対象の `ORIGINAL.md` で行います。README、DAKE_META、release_body、booth_product、Store、generated JSONは、今回の変更が影響するものだけ更新します。

GitHub Release作業ならrelease bodyと成果物を確認します。BOOTH作業でない時にbooth_productやサムネイルを更新せず、Store作業でない時にgenerated JSONやCloudflareを確認しません。

## 5. 既存ツールの責務

- `tools/release_capture.py`: SHIPの起動・主要機能・GitHub Release証跡。画像・BOOTH・サイト・動画は任意情報で、SHIP停止条件にしない。
- `tools/release_product_pipeline.py`: Legacy互換のDISTRIBUTION / COMMERCE状態確認。通常販売フローでは自動advanceせず、各レイヤーの専用ツールを使う。
- `tools/release_product.py` / `tools/finalize_product_release.py` / `tools/commerce/`: COMMERCE専用。
- `tools/store/`: DISTRIBUTION / COMMERCE専用。
- `tools/make_booth_ready.py`: BOOTH登録素材の生成専用。SHIP完了を判定しない。
- `tools/check_booth_ready.py`: 人間がDakeBOOTHアシストで登録するための素材だけを確認する。
- `tools/record_booth_registration.py`: 人間が登録済みのBOOTH URLをORIGINALへ記録するだけ。登録ツールではない。

既存の `release` / `formal_release` というファイル名・互換フィールドは直ちに全改名しません。これらはLegacy名であり、新しいSHIP定義を上書きしません。

## 6. Codexの境界

依頼されたレイヤーだけを実施します。例えばSHIP依頼なら、必要な権限がある範囲でSHIPを完了し、

```text
SHIP = COMPLETE
DISTRIBUTION = PENDING
COMMERCE = NOT_REQUIRED または PENDING
```

のように分けて報告します。完全性や念のためを理由に他レイヤーへ進みません。
