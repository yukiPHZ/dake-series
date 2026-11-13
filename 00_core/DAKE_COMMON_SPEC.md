# DAKE_COMMON_SPEC

DAKEシリーズの最低限の共通原則を定める正本です。個別アプリの固有仕様は各 `ORIGINAL.md` に置き、ここへ複製しません。

## DAKEの基本思想

- 単機能、最短距離、迷わない、止まらない、軽い。
- 必要以上に機能・設定・画面・工程を増やさない。
- 元データを不用意に壊さず、失敗時に戻りやすくする。
- AIは補助し、人間が最終判断する。

## 通常の読込

通常のアプリ開発・修正で最初に読むものは原則3つだけです。

1. repoルートの `AGENTS.md`
2. この `DAKE_COMMON_SPEC.md`
3. 対象アプリの `ORIGINAL.md`

build、UI、Git、アイコン、スクリーンショット、BOOTH、Store、Stripe等の詳細仕様は、その作業が発生する時だけ参照します。README、DAKE_META、release_body、booth_product、Store、generated JSONを作業前に全件照合しません。

## 正本と派生ビュー

- 個別アプリ・商品の真の正本は `ORIGINAL.md`。
- README、DAKE_META、release_body、booth_product、Store表示、generated JSONは派生ビュー。
- 判断はORIGINALで行い、変更が影響する派生ビューだけ更新する。
- `ORIGINAL.md` がないLegacyアプリは、次回そのアプリを修正・出荷する時に新仕様へ段階移行する。全アプリ一括移行はしない。
- generatedファイルは手編集しない。

## ローカル作業場所

- DAKEシリーズと関連開発の作業ファイル、フォルダ、一時成果物、テスト資材、バックアップ、ビルド作業場所は原則 `C:\Users\yukiz\devlop\` 配下に置く。
- Codex判断でDesktop、Documents、Downloads、`C:\temp`、`C:\work`、ユーザーフォルダ直下等へ永続的な開発用ファイルやフォルダを作らない。
- 一時処理は可能な限り対象プロジェクト配下に限定する。OSやライブラリが内部利用する標準TEMP領域は除く。
- 既存の `devlop` 外ファイルは、明示指示なしに移動、削除、整理しない。

## SHIP / DISTRIBUTION / COMMERCE

### SHIP

アプリ本体を完成し、GitHub Releaseで公開する工程です。

```text
ORIGINAL確認
↓
build
↓
生成された配布exeで起動・主要機能確認
↓
必要なGit commit / push
↓
GitHub Release
↓
SHIP = COMPLETE
```

GitHub Releaseが完了した時点でSHIPは完了です。BOOTH、dakeapp.com、Store、Stripe、Cloudflare、SNSが未完了でも、SHIPをPENDINGへ戻しません。

SHIPを止める条件は原則次の4つだけです。

1. build失敗
2. exeが起動しない
3. 主機能が成立しない
4. データ破壊等の重大リスクが確認された

### DISTRIBUTION

販売・配布先への掲載です。BOOTH、dakeapp.com、store.dakeapp.comを扱います。

状態は `COMPLETE` / `PARTIAL` / `PENDING` とし、SHIPと独立して記録します。

### COMMERCE

決済関連です。Stripe、Payment Link、`payment_status`、checkout確認を扱います。

状態は `READY` / `PENDING` / `NOT_REQUIRED` とします。無料アプリや決済不要対象の `NOT_REQUIRED` は正常状態です。

「正式出荷」は完了判定の総称として使いません。既存の `formal_release` 等の機械値はLegacy互換名として残せますが、SHIPをDISTRIBUTIONやCOMMERCEの完了で再判定しません。

## 変更内容に応じた確認

- 通常の小修正: build、生成された配布exeの主要動作、GitHub Release。
- build / dependency / PyInstaller / DLL / runtime / packaging変更: 通常確認に加え、開発環境依存を除いた確認。
- UI変更: 必要な表示確認。公開表示が変わる場合だけスクリーンショット更新。
- 販売情報変更: 対象のBOOTH / dakeapp.com / Storeだけを確認。
- Stripe変更: Commerceだけを確認。
- 変更されていないレイヤーを毎回再確認しない。

通常出荷では、生成された配布exeで主要機能が成立することを確認します。空フォルダ、ZIP展開後、非開発PCで同じ確認を重ねるのは、packaging・依存・runtime構成が変わった場合に限ります。

## Windows exe

- 利用者へPythonや追加ライブラリの手動導入を要求しない。
- Windows配布は原則 `--onefile` とし、必須の隣接runtimeへ依存しない。
- 技術的に成立しない場合は自動的に方式を変えず、原因と代案を報告する。
- Web等の非exe成果物へこの規則を強制しない。

## UI

新規アプリと次回更新対象の基本フッターは次の2要素だけです。

```text
戸建買取査定 ｜ by Yukihiko Kikuta
```

- `戸建買取査定`: リンク。
- `by Yukihiko Kikuta`: 非リンク。
- Instagram、長いコピーライト、しまりす不動産の長いブランド表記、不要な外部リンクを新仕様へ入れない。
- AI開発手法や開発ツール名を公開クレジットへ使わない。
- フッターは主機能より目立たせず、About画面を新設しない。
- 既存全アプリを一括変更せず、次回更新時に移行する。

戸建買取査定は各exeへ実ページURLを固定せず、DAKE側の恒久入口を使います。

```text
https://dakeapp.com/go/kaitori?app=<app_id>
```

`app_id` は英小文字・数字・アンダースコアの安定した識別子とし、Web遷移後の流入元識別だけに使います。

## ブランドと営業主体

- しまりす不動産: Yukihiko Kikuta個人の情報発信・制作活動。
- DAKE: Yukihiko Kikutaが制作する実務ツール群。
- 戸建買取査定: 株式会社さくら都市における業務。同社の基準・手続きに基づいて対応。

「しまりす不動産が戸建を買い取る」と読める表現を使いません。査定・買取業務の主体を株式会社さくら都市以外へ誤認させません。

## privacy

DAKEはlocal-firstです。アプリ本体から次を外部送信しません。

- 使用回数、処理回数、操作履歴
- 開いたファイル名、ファイルパス
- PDF・画像・入力内容
- PC情報、個人情報

営業計測は、利用者が戸建買取査定リンクをクリックしてWebへ遷移した後だけを対象とします。アプリ内テレメトリを追加しません。

## Windowsファイルプロパティ

build時に可能な限り自動付与します。

- ProductName: 対象アプリ名
- FileDescription: 対象アプリの日本語表示名
- CompanyName: `PEAKHEADZ`
- LegalCopyright: `© 2026 Yukihiko Kikuta`
- Authors: `Yukihiko Kikuta`
- FileVersion / ProductVersion: `x.x.x`
- OriginalFilename: 対象exe名

これは毎回の手入力・目視ゲートではありません。build方式やメタデータ生成を変更した時に確認します。

## Gitとバージョン

- `build/`、`dist/`、`*.spec`、`*.exe`、`*_config.json`、`__pycache__/`、`*.pyc` は通常commitしない。
- exeは通常GitHub Release成果物として扱う。
- 別PCでの確認が必要なexeは、通常commitへ含めず、実機確認用GitHub PrereleaseのAssetとしてダウンロード可能にする。実機確認用Releaseは正式出荷・SHIP完了と区別する。
- 既存のユーザー変更・未追跡ファイルを戻す、削除する、混在commitすることは禁止。
- patch = バグ修正・小変更、minor = 後方互換の機能追加・比較的大きな改善、major = 互換性を壊す変更。
- 過去のGitHub Releaseは原則保持し、問題時は直前Releaseへ戻す。専用バックアップ工程を追加しない。

## 作業範囲と外部操作

Codexは依頼されたレイヤーの外へ自動的に進みません。SHIP依頼からBOOTH、Store、Stripe、Cloudflare、SNS、販売文章、価格、画像生成へ進まないでください。

公開、送信、決済作成・変更、Stripe live操作、BOOTH公開、価格変更、外部登録、削除は、人間の明示依頼なしに実行しません。

新しい状態管理、DB、manifest、CI、自動監査、多段承認、バックアップ、ログ基盤を理由なく追加しません。既存構造の削除・統合・単純化を優先します。
