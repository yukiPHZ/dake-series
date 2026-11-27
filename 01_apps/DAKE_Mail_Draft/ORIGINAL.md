# ORIGINAL.md

## 正本宣言

このファイルは `DAKE_Mail_Draft` の真の正本です。

README.md、DAKE_META、release_body.md、booth_product.txt、Store表示などは、このファイルから派生するビューです。

## 基本情報

- app_id: DAKE_Mail_Draft
- title: Dakeメール下書き
- short_title: Dakeメール下書き
- category: メール
- status: available
- version: 1.0.0
- price: 500円
- distribution: DAKE Store（Stripe private delivery） / BOOTH / dakeapp.com
- target_platform: Windows

## 目的

CSV名簿からOutlookの個別下書きを作る、送信しない実務補助ツールです。

## 守るべき挙動

- CSVを読み込み、会社名・名前・メールアドレスに使う列を利用者が選択できる。
- 件名・本文テンプレートへ、会社名・名前・メールアドレスを差し込む。
- Outlook Classicの個別メールを下書きとして保存し、自動送信しない。
- Outlook側で生成された既定署名HTMLを可能な限り維持し、その前へ本文を挿入する。
- ローカルファイルを複数添付でき、存在しない添付は生成前に停止する。
- 5 / 10 / 20 / 50 / 100 / 全件の生成件数制限を維持する。
- 生成前に先頭1件の宛先・件名・本文をプレビューする。
- 空メール、不正形式、重複、作成成功、エラーをCSVレポートへ記録する。
- 利用者がOutlook上で宛先・本文・添付を確認してから、自分の判断で送信する流れを維持する。

## 変更してはいけない核

- Outlook下書き作成補助を、自動送信・SMTP送信・営業メール配信ツールへ変えない。
- `CreateItem`、既定署名の読込、本文挿入、`Save` による下書き保存の流れを、ブランド修正のついでに変更しない。
- CSV列選択、テンプレート差込、添付、件数制限、プレビュー、レポートの既存挙動を崩さない。
- CSV、テンプレート、添付、メール内容、利用履歴を外部送信しない。
- Outlook Classicを必要とする現在の前提を、未検証の別方式へ置き換えない。

## 対象ユーザー

既存ファイルに明示なし

## 解決する困りごと

- CSV差し込み、添付、署名保持、生成レポートに対応。送信機能はありません。
- CSV名簿からOutlookの個別下書きを作る、送信しない実務補助ツールです。

## 主な機能

- Dakeメール下書き
- CSV名簿からOutlook下書きを作成
- 件名・本文差し込みと添付に対応
- 送信機能なし
- Windows向けexe

## 使い方の要点

既存ファイルに記載なし

## CLI連携・外部ツール

- 使用する外部ツール: Outlook, pywin32
- 同梱有無: 既存ファイルに記載なし
- PATH依存: 既存ファイルに記載なし
- エラー時の扱い: 既存ファイルに記載なし
- Codex作業時の注意: CLIや外部ツール情報を追加する場合は、README等の派生ビューだけでなくこのORIGINALへ戻す。

## 対応形式・非対応形式

- 対応入力: 既存ファイル内の関連語: CSV, メール
- 対応出力: 既存ファイルに記載なし
- 非対応: 既存ファイルに記載なし
- 注意: 形式を推測で追加しない。既存READMEまたは実装確認後に更新する。

## やらないこと / 非ゴール

- New Outlook / Web Outlook では動作しない可能性があります。
- 迷惑メール、無断大量送信、法令・規約違反用途での利用は禁止です。
- 存在しないパスがある場合、生成前に警告します。
- ### Outlookが起動しない
- Windows版 Microsoft Outlook Classic がインストールされ、初期設定済みか確認してください。New Outlook / Web Outlook では動作しない場合があります。

## 設定・ログ・保存方針

- SMTP送信、メールアカウント保存、パスワード保存はしません。
- Windows版 Microsoft Outlook Classic がインストールされ、初期設定済みか確認してください。New Outlook / Web Outlook では動作しない場合があります。
- Outlook Classic 側で既定署名が設定されているか確認してください。このアプリは下書きを一度表示し、Outlook側に署名を生成させてから本文を挿入します。
- 1. Outlook Classicを起動し、既定署名が必要なら設定します。

## 非破壊・上書き禁止方針

- 7. 作成されたメールはOutlook上で内容確認後、不要なら削除してください。
- 大切なファイルは事前にバックアップを推奨します。
- 大切なファイルは事前にバックアップを推奨します

## 公開用説明の元情報

- display_name: Dakeメール下書き
- site_title: Dakeメール下書き
- site_description: CSV名簿からOutlookの個別下書きを作る、送信しない実務補助ツールです。
- launcher_description: CSVからOutlookの個別下書きだけ作ります。
- update_summary: CSV差し込み、添付、署名保持、生成レポートに対応。送信機能はありません。
- BOOTH紹介文: CSV名簿から、Outlookの個別下書きを作ります。

・Outlook下書き作成のみ
・自動送信なし
・件名、本文、添付ファイルに対応
・送信前にOutlookで確認
・Windows向け

送信を急がず、
下書きだけを静かに整える道具です。

## README生成用情報

- 概要: CSV名簿からOutlookの個別下書きを作る、送信しない実務補助ツールです。
- 使い方: 既存ファイルに記載なし
- 必要なもの: Windows環境
- 注意: READMEへ出す内容は公開可能な情報に限定する。
- ビルド: 既存READMEまたはbuild.batを参照。

## DAKE_META生成用情報

```json
{
  "app_key": "DAKE_Mail_Draft",
  "display_name": "Dakeメール下書き",
  "launcher_title": "メール下書き作成",
  "launcher_description": "CSVからOutlookの個別下書きだけ作ります。",
  "site_title": "Dakeメール下書き",
  "site_description": "CSV名簿からOutlookの個別下書きを作る、送信しない実務補助ツールです。",
  "update_summary": "CSV差し込み、添付、署名保持、生成レポートに対応。送信機能はありません。",
  "folder_name": "DAKE_Mail_Draft",
  "exe_name": "DakeMail_Draft.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Mail_Draft_v1.0.0",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true,
  "demo_video_path": "release_artifacts/demo.mp4",
  "demo_video_url": "",
  "social_release_path": "release_artifacts/social_release.json"
}
```

## release_body生成用情報

- Dakeメール下書き
- CSV名簿からOutlook下書きを作成
- 件名・本文差し込みと添付に対応
- 送信機能なし
- Windows向けexe

## booth_product生成用情報

- 商品名: Dakeメール下書き
- 価格案: 500円
- 商品紹介文: CSV名簿から、Outlookの個別下書きを作ります。

・Outlook下書き作成のみ
・自動送信なし
・件名、本文、添付ファイルに対応
・送信前にOutlookで確認
・Windows向け

送信を急がず、
下書きだけを静かに整える道具です。
- タグ: メール
Outlook
Windows
実務
ツール
仕事効率化
軽量
シンプル
- 商品画像: assets/booth_thumbnail.jpg
- 補助画像: assets/screenshot.jpg
- 作品ファイル: booth_ready/DakeMail_Draft.zip
- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Mail_Draft_v1.0.0（履歴ページ。無料配布assetなし）
- BOOTH URL: https://peakheadz.booth.pm/items/8448001

## Store表示用情報

- 商品名: Dakeメール下書き
- キャッチ: CSV名簿からOutlookの個別下書きを作る、送信しない実務補助ツールです。
- 説明: CSV名簿から、Outlookの個別下書きを作ります。

・Outlook下書き作成のみ
・自動送信なし
・件名、本文、添付ファイルに対応
・送信前にOutlookで確認
・Windows向け

送信を急がず、
下書きだけを静かに整える道具です。
- 価格: 500円
- 画像: assets/booth_thumbnail.jpg（存在: あり）
- ダウンロード導線: Stripe Checkout Session検証後にprivate R2から即時配布
- サポート方針: 既存ファイルに記載なし
- Stripe Payment Link: https://buy.stripe.com/14A6oH57H9kj9WEe250gw0Z
- Store販売状態: stripe_ready
## 価格・販売方針

- BOOTH価格案: 500円
- BOOTH URL: https://peakheadz.booth.pm/items/8448001
- GitHub Release: 履歴ページのみ。無料配布assetは2026-11-22に終了
- Store販売: Stripe Payment Link（300円）

## 配布・ダウンロード方針

- GitHub Release: 履歴ページのみ。公開無料Downloadは終了
- BOOTH: https://peakheadz.booth.pm/items/8448001
- BOOTH配布zip: booth_ready/DakeMail_Draft.zip
- Store配布導線: Stripe Checkout Session検証後にprivate R2から即時配布
- purchase_delivery_method: stripe_session_private_r2
- purchase_delivery_ready: yes
- delivery_object_key: desktop/DAKE_Mail_Draft/current.zip
- 正式販売経路: DAKE StoreのStripe private-delivery Payment Link
- GitHub無料配布: 2026-11-22終了（Release履歴とtagは維持、配布exe assetを削除）
- 本番確認: Stripe実決済、Store redirect、private R2 download、ZIP download、exe起動を人間が確認済み

## 免責・注意事項

- 【注意事項】
- Windows向けアプリです
- ご利用は自己責任でお願いいたします
- 大切なファイルは事前にバックアップを推奨します
- 本ソフトウェアの無断転載・再配布を禁止します
- 環境によっては起動時にWindowsの警告が表示される場合があります
- https://peakheadz.com

## 同梱ファイル方針

- exe: DakeMail_Draft.exe
- README.txt: booth_ready/README.txt（存在: あり）
- 注意事項.txt: booth_ready/注意事項.txt（存在: あり）
- 入れないもの: build/、dist/、*.spec、設定ファイル、個人データ、ソース一式は正式配布zipへ混ぜない。
- 利用者向けCopyright: © Yukihiko Kikuta / PEAKHEADZ

## 販売版build確認

- 確認日: 2026-11-22
- build方式: 既存 `build.bat` のPyInstaller `--onefile` / `--noconsole`
- exe: `dist/DakeMail_Draft.exe`
- exe size: 15,691,033 bytes
- exe SHA256: `1F605224F4E654D8FF6470502E6F344A3745A9E8BEE7E989C07E634ACF2C068D`
- Windows metadata: CompanyName=`PEAKHEADZ`、FileVersion / ProductVersion=`1.0.0`、OriginalFilename=`DakeMail_Draft.exe`、LegalCopyright=`© Yukihiko Kikuta / PEAKHEADZ`
- icon: 共通icoをonefile内部へ同梱し、sourceはrepo正本、frozenはPyInstaller展開先から解決する
- Explorer exe icon: 共通DAKEアイコンとpixel一致
- GUI titlebar icon: 実行中exeで共通DAKEアイコンを確認
- Windows taskbar running icon: 実行中exeで共通DAKEアイコンを確認
- launch-check: PASS
- GUI credit: `© Yukihiko Kikuta / PEAKHEADZ`
- 販売ZIP: `booth_ready/DakeMail_Draft.zip`
- ZIP size: 15,057,205 bytes
- ZIP SHA256: `EEB66EFFD271B8CA847FF598153F633143097B7152BA89FA10A1AF03EB3CE94B`
- ZIP contents: `DakeMail_Draft.exe` / `README.txt` / `注意事項.txt`

## スクリーンショット・画像方針

- assets/screenshot.webp: あり
- assets/screenshot.jpg: あり
- assets/booth_thumbnail.jpg: あり
- Store用画像: 未確定。既存画像を元に派生する想定。

## 今後の改善予定

現時点では未設定です。

## Codex作業時の注意

- 触ってよい: ORIGINAL.md、および明示された派生ビュー。
- 触らない: アプリ本体、README、release_body、booth_product、booth_ready、assets、distは今回の横展開では変更しない。
- 外部公開しない: 未確定のStore URLや未確認の販売導線を確定情報として書かない。
- 自動操作しない: BOOTH公開、GitHub Release更新、Store公開はこのファイル作成では行わない。

## 派生物一覧

- README.md: GitHub公開用ビュー（存在: あり）
- DAKE_META: README内の機械利用ビュー（存在: あり）
- release_body.md: GitHub Release用ビュー（存在: あり）
- booth_product.txt: BOOTH登録用ビュー（存在: あり）
- booth_ready/booth_product.txt: BOOTH登録時の実使用ビュー（存在: あり）
- Store: 自社販売ビュー。Store専用正本は作らず、このORIGINAL由来の情報を使う。

## 参照した既存ファイル

- README.md: あり
- README内DAKE_META: あり
- release_body.md: あり
- booth_product.txt: あり
- booth_ready/booth_product.txt: あり
- booth_ready/README.txt: あり
- booth_ready/注意事項.txt: あり
- assets/screenshot.webp: あり
- assets/booth_thumbnail.jpg: あり

## Commerce

- payment_status: stripe_ready