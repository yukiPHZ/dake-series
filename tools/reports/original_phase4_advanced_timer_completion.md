# ORIGINAL Phase4 Completion Report - DAKE_Time_AdvancedTimer

## 目的

`DAKE_Time_AdvancedTimer` を対象にした `ORIGINAL.md` 試験導入の結果を記録する。

このレポートは、DAKEの正本主義エンジンが最初の既存アプリで機能したことを残し、次の `store.dakeapp.com` 接続を単なるECサイト構築ではなく、以下の流れの延長として進めるための記録である。

```text
ORIGINAL.md
↓
派生ビュー
↓
自社販売ビュー
```

## 対象アプリ

- app: `DAKE_Time_AdvancedTimer`
- folder: `01_apps/DAKE_Time_AdvancedTimer`
- title: Dakeアドバンスドタイマー
- status: available
- app_type: market
- completion_goal: formal_release

## 実施したPhase

- Phase 1: ORIGINAL.mdルール策定
- Phase 2: Codexルール改定
- Phase 3: 対象アプリへのORIGINAL.md導入
- Phase 4-1: 派生ビュー差分レビュー
- Phase 4-2: 正本への情報回収
- Phase 4-3: 派生ビューへの最小反映

## Phase 1：ORIGINAL.mdルール策定

`00_core/DAKE_ORIGINAL_RULE.md` と `00_core/DAKE_ORIGINAL_TEMPLATE_APP.md` により、DAKEシリーズでは `ORIGINAL.md` を真の正本として扱う方針を定義した。

定義された関係:

```text
ORIGINAL.md
↓
README.md
DAKE_META
release_body.md
booth_product.txt
Store表示
```

README、DAKE_META、release_body、booth_product、Store表示は、正本ではなく派生ビューとして扱う。

## Phase 2：Codexルール改定

00_core内の共通ルールを改定し、Codex作業時の優先順位を `README.md` ではなく `ORIGINAL.md` に移した。

作業前確認順は以下へ寄せた。

```text
1. ORIGINAL.md
2. README.md
3. DAKE_META
4. release_body.md
5. booth_product.txt
6. 関連仕様ファイル
```

既存アプリで `ORIGINAL.md` が未導入の場合は、過渡期ルールとしてREADMEを暫定参照してよい。

## Phase 3：対象アプリへのORIGINAL.md導入

`01_apps/DAKE_Time_AdvancedTimer/ORIGINAL.md` を新規作成し、既存のREADME、DAKE_META、release_body、booth_product、booth_ready内テキストから情報を集約した。

集約した主な情報:

- 基本情報
- 目的
- 対象ユーザー
- 解決する困りごと
- 主な機能
- 使い方の要点
- README生成用情報
- DAKE_META生成用情報
- release_body生成用情報
- booth_product生成用情報
- Store表示用情報
- 価格・販売方針
- 配布・ダウンロード方針
- 免責・注意事項
- 同梱ファイル方針
- スクリーンショット・画像方針
- Codex作業時の注意
- 派生物一覧

このPhaseでは、READMEやbooth_productなどの派生ビューは更新しなかった。

## Phase 4-1：派生ビュー差分レビュー

`tools/reports/original_phase4_advanced_timer_review.md` を作成し、`ORIGINAL.md` と既存派生ビューの差分をレビューした。

比較した派生ビュー:

- `README.md`
- README内 `DAKE_META`
- `release_body.md`
- `booth_product.txt`
- `booth_ready/booth_product.txt`
- `booth_ready/README.txt`
- `booth_ready/注意事項.txt`

確認結果:

- README、DAKE_META、release_bodyには大きな矛盾なし。
- booth_product系には、BOOTH本文の締め文が派生ビュー側にだけ残っていた。
- 注意事項の粒度に差はあるが、重要なWindows警告の注意はORIGINAL.md側に保持済み。

## Phase 4-2：正本への情報回収

Phase 4-1で見つかった、派生ビュー側にだけ残っていた文言を `ORIGINAL.md` へ戻した。

正本へ戻した情報:

```text
実務の流れを、少し静かにするための道具です。
```

この文言は単なるBOOTH登録用の一文ではなく、DAKEアプリの思想、公開説明、Store表示にも使える文言として扱った。

追加した場所:

- `公開用説明の元情報`
- `booth_product生成用情報`
- `Store表示用情報`

## Phase 4-3：派生ビューへの最小反映

Phase 4-2で補強した `ORIGINAL.md` を真の正本として、BOOTH登録用ビューへ最小反映した。

再反映した派生ビュー:

- `01_apps/DAKE_Time_AdvancedTimer/booth_product.txt`
- `01_apps/DAKE_Time_AdvancedTimer/booth_ready/booth_product.txt`

反映内容:

```text
実務の流れを、少し静かにするための道具です。
```

既存では同じ意味の文が2行に分かれていたため、`ORIGINAL.md` と同じ一文として検出できる形に揃えた。

価格、BOOTH URL、GitHub Release URL、作品ファイル名、画像ファイル名、タグ、注意事項は維持した。

## 実証できたこと

1. `ORIGINAL.md` を真の正本として作成できた。
2. 既存README / release_body / booth_product等から情報を集約できた。
3. 派生ビュー側にだけ残っていた重要情報を発見できた。
4. その情報を `ORIGINAL.md` へ戻せた。
5. `ORIGINAL.md` 由来の情報として派生ビューへ再反映できた。
6. 新しい正本を作らずに、派生ビューを更新できた。

今回、以下の流れが成立した。

```text
派生ビュー側にだけ残っていた文言を発見
↓
ORIGINAL.mdへ戻す
↓
ORIGINAL.md由来の情報としてbooth_product系へ再反映
```

つまり、最小範囲ではあるが、以下の実証に成功した。

```text
ORIGINAL.md
↓
booth_product.txt
↓
booth_ready/booth_product.txt
```

## 発見したこと

- READMEはGitHub公開用ビューとして短くてよく、すべての説明を載せる必要はない。
- `ORIGINAL.md` はREADMEより広い情報を保持する場所として機能する。
- booth_productは登録用ビューとして実運用情報を持ちやすく、正本へ戻すべき文言が残りやすい。
- BOOTH用の短い締め文は、Store表示にも使える可能性があるため、`ORIGINAL.md` へ戻す価値がある。
- DAKE_METAとrelease_bodyは、今回の対象では大きな差分がなく、無理に更新する必要はなかった。

## 正本へ戻した情報

```text
実務の流れを、少し静かにするための道具です。
```

この文言は、BOOTH登録用ビュー側だけに残っていたが、DAKEアプリの思想・公開説明・Store表示にも使える文言として `ORIGINAL.md` へ戻した。

## ORIGINAL.mdから再反映した派生ビュー

- `01_apps/DAKE_Time_AdvancedTimer/booth_product.txt`
- `01_apps/DAKE_Time_AdvancedTimer/booth_ready/booth_product.txt`

両方ともBOOTH登録用ビューとして整合している。

## 触らなかったもの

- `README.md`
- README内 `DAKE_META`
- `release_body.md`
- `booth_ready/README.txt`
- `booth_ready/注意事項.txt`
- `main.py`
- `build.bat`
- `assets/`
- `dist/`
- Store関連
- Stripe関連
- BOOTH Web
- GitHub Release

## Store接続前の判断

`DAKE_Time_AdvancedTimer` における `ORIGINAL.md` 試験導入は成功とする。

次Phaseでは、`store.dakeapp.com` を自社販売ビューとして検討してよい。

ただし、Store専用の商品正本は作らない。

Storeは `ORIGINAL.md` 由来の情報を読む、またはビルド時に生成される generated データを読む。

## 次に進むPhase

次Phase候補:

1. `DAKE_Time_AdvancedTimer` を対象に、Store表示用データの最小生成を検討する。
2. `ORIGINAL.md` からREADME / DAKE_META / release_body / booth_productを生成・検証する小さなスクリプト設計に進む。
3. Storeへ接続する前に、Store用のgeneratedデータ形式を定義する。

推奨は、まずStore専用正本を作らずに済む generated データ形式の定義から始めること。

## 残課題

- Storeのダウンロード導線は未確定。
- Storeサポート方針は未確定。
- Store用画像の最終方針は未確定。
- README / release_body / DAKE_META の自動生成は未実装。
- 全アプリへの `ORIGINAL.md` 横展開は未実施。
- 生成スクリプトは未作成。

## 結論

`DAKE_Time_AdvancedTimer` での `ORIGINAL.md` 試験導入は成功した。

今回の実証により、DAKEの情報管理は次の形へ進められる。

```text
ORIGINAL.md
↓
派生ビュー
↓
自社販売ビュー
```

今後のStore構築は、ECサイト単体の構築ではなく、`ORIGINAL.md` 正本主義の派生ビューとして進める。
