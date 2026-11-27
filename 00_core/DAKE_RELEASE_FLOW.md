# DAKE_RELEASE_FLOW

SHIP / DISTRIBUTION / COMMERCEを別々に扱うための最小フローです。共通原則は `DAKE_COMMON_SPEC.md` を正本とします。

## 0. 現行の正式出荷・配布方式（確定正本）

- **無料配布枠**：`DakePDF結合`（`DAKE_PDF_Merge`）のみ GitHub Release に公開アセット（`.exe`）を残す（Store に同機能の応援購入版 500円 も併設）。
- **有料アプリ・Pack**：DAKE Store（Stripe Payment Link → Cloudflare Pages Functions `/download/` の Stripe Checkout Session 検証 → プライベート Cloudflare R2 `dake-desktop-releases` の `desktop/<id>/current.zip` 自動配信）＋ BOOTH 並行販売（`booth_ready/` 生成を維持）。
- **GitHub Releases**：有料化済みアプリ・Pack は公開アセット（`.exe` / `.zip`）を置かず、リリースタグ・リリース本文（履歴テキスト）のみ残す。

## 1. SHIP

```text
対象ORIGINALを確認
↓
変更を実装
↓
build（単独exe・共通アイコン・VersionInfo確認）
↓
生成された配布exeで起動・主要機能確認
↓
対象差分だけcommit / push
↓
GitHub Release作成（無料枠のDakePDF結合のみ公開アセット付与。有料アプリ・Packはタグとリリース本文のみ作成し、配布ZIPはR2とbooth_readyへ回す）
↓
SHIP = COMPLETE
```

GitHub Release（タグ・履歴確定、無料枠は公開アセット含む）が完了すればSHIPは完了です。DISTRIBUTION / COMMERCEが未完了でもSHIPをPENDINGへ戻しません。

### 実機確認用Prerelease

正式出荷前に別PCで確認が必要な場合は、`DAKE_App_v1.0.1-test.1` のようにtest版と分かるタグでGitHub Prereleaseを作成できます。

- exeは通常commitへ含めず、Prerelease Assetとして公開する。
- 対象commit、Asset名、SHA256、`isPrerelease=true` を読み戻す。
- 公開Assetを `C:\Users\yukiz\devlop\` 配下へ再ダウンロードし、build版とSHA256を照合する。
- 実機確認用Prereleaseの作成だけでは正式出荷または `SHIP = COMPLETE` としない。

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

### store.dakeapp.com（R2自動配信＋カタログ同期）

有料アプリおよび Pack の正式配信経路は、Stripe Payment Link → Cloudflare Pages Functions（`/download/`・Stripe Checkout Session 検証）→ プライベート Cloudflare R2 バケット `dake-desktop-releases`（`desktop/<id>/current.zip`）の自動配信です。

1. 配布 ZIP をプライベート R2 バケット `dake-desktop-releases` の `desktop/<id>/current.zip` へ配置（または上書き更新）する。
2. 新規商品の場合は `dake-store-site/functions/_lib/private-download.mjs` の `DESKTOP_DELIVERIES` ホワイトリストへ登録する。
3. ORIGINAL由来の `store_products.generated.json` を `tools/store/sync_store_to_site.py` で生成・同期・検証する（`dake-store-site` および `dakeapp-site` 両方へ同期）。自動commit / push / deployは行いません。

BOOTHはStoreとは独立した並行販売経路として維持します（`booth_ready/` 生成を継続）。

### アプリ新規追加・ブラッシュアップ時の掲載面更新チェックリスト（正式工程）

アプリの新規追加、または既存アプリのブラッシュアップ（機能改善・UI/フッター刷新・バージョン更新）を実施した際は、以下の掲載面・同梱物を漏れなく確認・更新します。

- [ ] **1. `dakeapp.com`（`dakeapp-site`）**
  - 詳細ページ `public/apps/<slug>/index.html`（概要・FEATURES・CTA・UPDATE 欄のバージョン／更新内容）
  - スクリーンショット `public/assets/images/apps/<slug>/screenshot-01.webp`（UI変更時は差し替え）
  - 一覧・検索・更新情報ページ（`public/apps/index.html`、`public/update/index.html`、`tools/sync_dakeapp_apps_json.py` による `public/assets/data/apps.json` 同期、`public/assets/data/store_products.generated.json` 同期）
  - `sitemap.xml` / `robots.txt` 再生成（新規ページ追加時は `node scripts/generate-sitemap.js`）
- [ ] **2. `store.dakeapp.com`（`dake-store-site`）**
  - プライベート R2 バケット `dake-desktop-releases` の `desktop/<id>/current.zip` 配置／差し替え
  - `functions/_lib/private-download.mjs`（`DESKTOP_DELIVERIES`）の R2 配信ホワイトリスト確認・追加
  - `tools/store/sync_store_to_site.py` による `store_products.generated.json` 再生成・同期・検証
  - Stripe Payment Link（`dake_item_id`・`purchase_delivery_method=stripe_session_private_r2`・リダイレクト先 `/download/?session_id={CHECKOUT_SESSION_ID}`）整合確認
- [ ] **3. Pack 同梱物更新**
  - 対象アプリが `DAKE_Pack_Document` / `DAKE_Pack_Mail` / `SHIMARISU_Pack` 等の収録アプリに含まれるか確認
  - 含まれる場合は各 Pack ZIP を再生成（`tools/make_pack_ready.py` / `SHIMARISU/tools/build_shimarisu_pack.py`）し、Pack 用 R2 `desktop/<pack_id>/current.zip` および BOOTH 側 Pack 配布物も更新
- [ ] **4. BOOTH 更新（並行販売）**
  - `tools/make_booth_ready.py` / `tools/check_booth_ready.py` で `booth_ready/`（ZIP・サムネイル・`booth_product.txt`）を生成・検証
  - 人間が DakeBOOTHアシスト等で BOOTH 商品を登録・ファイル差し替え・更新し、必要時に `tools/record_booth_registration.py` で URL を ORIGINAL へ記録
- [ ] **5. `peakheadz-project-index` 更新**
  - 新規追加・公開状態変更・アーカイブ／廃止がある場合は `peakheadz-project-index` の一覧・状態記録を整合させる

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
