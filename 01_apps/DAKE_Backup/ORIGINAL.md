# ORIGINAL.md

## 正本宣言

このファイルは `DAKE_Backup` の真の正本です。

README.md、DAKE_META、release_body.md、booth_product.txt、Store表示などは、このファイルから派生するビューです。

## 基本情報

- app_id: DAKE_Backup
- title: Dakeバックアップ
- short_title: バックアップ
- category: バックアップ / 保全
- status: available
- version: 1.0.0
- price: 500円
- distribution: GitHub ReleaseとBOOTHで配布する。
- target_platform: Windows

## 目的

ローカル正本フォルダを、指定した避難先フォルダへ一方向にコピー保存する。

## 対象ユーザー

- ローカル正本を別フォルダや別ドライブへ静かに残したい人
- 削除を伝播させないバックアップがほしい人
- Google Drive、USBメモリ、外付けSSDなどを避難先として使いたい人

## 解決する困りごと

- 同期アプリで削除まで伝播するのが怖い
- 正本と避難先の差分を手で確認するのが面倒
- 上書き前の旧ファイルを退避しておきたい

## 主な機能

- 正本フォルダを避難先へコピー保存
- 削除は伝播しない一方向バックアップ
- 差分確認と退避保存に対応
- Windows向けexe

## 使い方の要点

- 正本フォルダを選ぶ。
- 避難先フォルダを選ぶ。
- 差分を見る。
- 内容を確認して、残すを実行する。

## 公開用説明の元情報

ローカル正本フォルダを、指定した避難先フォルダへ一方向にコピー保存するWindows向けアプリです。

正本フォルダを、指定した避難先へ静かに残します。

実務の流れを、少し静かにするための道具です。

## README生成用情報

- 概要: ローカル正本フォルダを、指定した避難先フォルダへ一方向にコピー保存するWindows向けアプリです。
- 使い方: 正本フォルダを選ぶ。
- 必要なもの: Windows環境。
- 注意: 同期アプリではない。
- ビルド: `build.bat` を実行し、`dist/DakeBackup.exe` を生成する。

## DAKE_META生成用情報

既存README内の `DAKE_META` ブロックを元にした機械利用ビュー情報です。
単独の `DAKE_META` ファイルは既存ファイルに存在しません。


```json
{
  "app_key": "DAKE_Backup",
  "display_name": "Dakeバックアップ",
  "launcher_title": "バックアップ",
  "launcher_description": "正本フォルダを、指定した避難先へ静かに残します。",
  "site_title": "Dakeバックアップ",
  "site_description": "ローカル正本フォルダを、指定した避難先フォルダへ一方向にコピー保存するWindows向けアプリです。",
  "update_summary": "ローカル正本を削除伝播させずに残すバックアップアプリを追加しました。",
  "folder_name": "DAKE_Backup",
  "exe_name": "DakeBackup.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Backup_v1.0.0",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## release_body生成用情報

- 正本フォルダを避難先へコピー保存
- 削除は伝播しない一方向バックアップ
- 差分確認と退避保存に対応
- Windows向けexe

## booth_product生成用情報

- 商品名: Dakeバックアップ
- 価格案: 500円
- 商品紹介文: 正本フォルダを、指定した避難先へ静かに残します。
- 補足紹介文:
  - 正本フォルダを避難先へコピー保存
  - 削除は伝播しない一方向バックアップ
  - 差分確認と退避保存に対応
  - Windows向けexe
  - 実務の流れを、少し静かにするための道具です。
- タグ:
  - Windows
  - 実務
  - ツール
  - 仕事効率化
  - 軽量
  - シンプル
- 商品画像: assets/booth_thumbnail.jpg
- 補助画像: assets/screenshot.jpg
- 作品ファイル: booth_ready/DakeBackup.zip
- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Backup_v1.0.0
- BOOTH URL: https://peakheadz.booth.pm/items/8447578

## Store表示用情報

- 商品名: Dakeバックアップ
- キャッチ: 正本フォルダを、指定した避難先へ静かに残します。
- キャッチ補足: 実務の流れを、少し静かにするための道具です。
- 説明: ローカル正本フォルダを、指定した避難先フォルダへ一方向にコピー保存するWindows向けアプリです。
- 価格: 500円
- 画像: assets/booth_thumbnail.jpg / assets/screenshot.webp
- ダウンロード導線: 未確定
- サポート方針: 既存ファイルに記載なし

Storeは未構築のため、Store専用の商品正本は作りません。
- Stripe Payment Link: https://buy.stripe.com/eVq00j43D9kjc4Me250gw06
- Store雋ｩ螢ｲ迥ｶ諷・ stripe_ready

## 価格・販売方針

- BOOTH価格案: 500円
- BOOTH URL: https://peakheadz.booth.pm/items/8447578
- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Backup_v1.0.0
- Store販売: 未確定

## 配布・ダウンロード方針

- GitHub Releaseで `DakeBackup.exe` を配布する。
- BOOTHでは `booth_ready/DakeBackup.zip` を作品ファイルとして使う。
- dakeapp.com掲載対象です。
- Store配布導線は未確定です。

## 免責・注意事項

BOOTH ready内の注意事項、または既存READMEの注意事項を元にします。

- Windows向けアプリです。
- ご利用は自己責任でお願いいたします。
- 大切なファイルは事前にバックアップを推奨します。
- 本ソフトウェアの無断転載・再配布を禁止します。

## 同梱ファイル方針

- exe: DakeBackup.exe
- README.txt: booth_ready/README.txt (既存)
- 注意事項.txt: booth_ready/注意事項.txt (既存)
- 配布zip: booth_ready/DakeBackup.zip
- 入れないもの: ソースコード、build/、dist/、*.spec、__pycache__/、個人設定ファイル

## スクリーンショット・画像方針

- assets/screenshot.webp: assets/screenshot.webp
- assets/screenshot.jpg: assets/screenshot.jpg
- assets/booth_thumbnail.jpg: assets/booth_thumbnail.jpg
- Store用画像: 未確定。既存画像を元に派生する想定。

## 今後の改善予定

既存README、release_body.md、booth_product.txtには今後の改善予定の記載なし。

現時点では未設定です。

## Codex作業時の注意

- 触ってよい: ORIGINAL.mdの更新、派生ビューとの整合確認。
- 触らない: main.py、build.bat、assets、dist、booth_readyの内容をこのPhaseで変更しない。
- 外部公開しない: 未確定のStore URLや未確認の販売導線を確定情報として書かない。
- 自動操作しない: BOOTH更新、GitHub Release更新、Store構築、Stripe実装はこのPhaseでは行わない。

## 派生物一覧

- README.md: GitHub公開用ビュー。既存。
- DAKE_META: README.md内のJSONブロックとして存在。単独ファイルはなし。
- release_body.md: GitHub Release用ビュー。既存。
- booth_product.txt: BOOTH登録用ビュー。アプリ直下は 既存、`booth_ready/` は 既存。
- booth_ready/README.txt: 配布zip同梱用ビュー。既存。
- booth_ready/注意事項.txt: 配布zip同梱用ビュー。既存。
- Store: 未構築。将来 `ORIGINAL.md` 由来の情報から生成する。
