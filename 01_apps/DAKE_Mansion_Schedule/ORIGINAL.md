# ORIGINAL.md

## 正本宣言

このファイルは `DAKE_Mansion_Schedule` の真の正本です。

README.md、DAKE_META、release_body.md、booth_product.txt、Store表示などは、このファイルから派生するビューです。

## 基本情報

- app_id: dake_mansion_schedule
- title: マンション工程表
- short_title: マンション工程表
- category: 作業補助
- status: available
- version: 1.0.0（Release URLより）
- price: 500円
- distribution: GitHub Release / BOOTH / dakeapp.com（Storeは未確定）
- target_platform: Windows

## 目的

リフォーム許可申請用の工程表を、A3横PDFで作成します。

## 対象ユーザー

既存ファイルに明示なし

## 解決する困りごと

- マンション管理会社提出用の工程表作成を正式公開
- リフォーム許可申請用の工程表を、A3横PDFで作成します。

## 主な機能

- マンション管理会社へ提出するための、リフォーム工事工程表を作成します。
- 管理会社提出用のA3横工程表PDFを作成
- 引渡し日から着工日・完工日を自動計算
- 土日を含む45日間の工期内で、合計28営業日の初期工程を自動入力
- 残り平日は予備日として工程間に分散
- チェックした工事行だけをPDFに上詰めで出力
- 自由工事欄を追加し、任意の工事項目をPDF出力可能
- 各工事項目の開始日・終了日を手動修正可能
- 土日も表示し、工事バーは平日のみに表示するガントチャート風PDF出力
- PDFタイトルを「リフォーム工事工程表」に変更
- PDFファイル名に出力日時を付け、同名上書きを回避
- PDF作成後に保存先フォルダを自動で開く
- 会社支店名、担当者名、連絡先を次回起動時に自動入力
- PDF提出物からDAKEのコピーライトやブランド表記を除外

## 使い方の要点

既存ファイルに記載なし

## CLI連携・外部ツール

- 使用する外部ツール: 該当なし / 既存ファイルに記載なし
- 同梱有無: 既存ファイルに記載なし
- PATH依存: 既存ファイルに記載なし
- エラー時の扱い: 既存ファイルに記載なし
- Codex作業時の注意: CLIや外部ツール情報を追加する場合は、README等の派生ビューだけでなくこのORIGINALへ戻す。

## 対応形式・非対応形式

- 対応入力: 既存ファイル内の関連語: PDF
- 対応出力: 既存ファイルに記載なし
- 非対応: 既存ファイルに記載なし
- 注意: 形式を推測で追加しない。既存READMEまたは実装確認後に更新する。

## やらないこと / 非ゴール

- 本ソフトウェアの無断転載・再配布を禁止します。
- 本ソフトウェアの無断転載・再配布を禁止します

## 設定・ログ・保存方針

- PDFは出力日時付きのファイル名で保存し、同名ファイルがある場合は連番を付けて上書きを避けます。
- PDF作成後、保存先フォルダを自動で開きます。
- 7. 「保存先を選ぶ」でPDFの保存先を指定します。
- 10. PDF作成後、保存先フォルダが自動で開きます。
- PDF作成後に保存先フォルダを自動で開く

## 非破壊・上書き禁止方針

- PDFは出力日時付きのファイル名で保存し、同名ファイルがある場合は連番を付けて上書きを避けます。
- PDFファイル名に出力日時を付け、同名上書きを回避
- 大切なファイルは事前にバックアップを推奨します。
- 大切なファイルは事前にバックアップを推奨します

## 公開用説明の元情報

- display_name: マンション工程表
- site_title: マンション工程表
- site_description: リフォーム許可申請用の工程表を、A3横PDFで作成します。
- launcher_description: 管理会社提出用のA3横工程表を作成します。
- update_summary: マンション管理会社提出用の工程表作成を正式公開
- BOOTH紹介文: マンション管理会社へ提出するための、リフォーム工事工程表を作成します。
工事進捗管理ではなく、管理会社提出用の工程表作成アプリです。

・土日を含む45日間の工期内で、合計28営業日の初期工程を自動入力
・土日は表示し、工事バーは平日のみに表示
・チェックした工事だけをPDF出力
・自由工事欄に任意の工事を追加可能
・PDFは出力日時付きファイル名で保存

実務の流れを、
少し静かにするための道具です。

## README生成用情報

- 概要: リフォーム許可申請用の工程表を、A3横PDFで作成します。
- 使い方: 既存ファイルに記載なし
- 必要なもの: Windows環境
- 注意: READMEへ出す内容は公開可能な情報に限定する。
- ビルド: 既存READMEまたはbuild.batを参照。

## DAKE_META生成用情報

```json
{
  "app_key": "dake_mansion_schedule",
  "display_name": "マンション工程表",
  "launcher_title": "マンション工程表",
  "launcher_description": "管理会社提出用のA3横工程表を作成します。",
  "site_title": "マンション工程表",
  "site_description": "リフォーム許可申請用の工程表を、A3横PDFで作成します。",
  "update_summary": "マンション管理会社提出用の工程表作成を正式公開",
  "folder_name": "DAKE_Mansion_Schedule",
  "exe_name": "DakeMansion_Schedule.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Mansion_Schedule_v1.0.0",
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

- マンション管理会社へ提出するための、リフォーム工事工程表を作成します。
- 管理会社提出用のA3横工程表PDFを作成
- 引渡し日から着工日・完工日を自動計算
- 土日を含む45日間の工期内で、合計28営業日の初期工程を自動入力
- 残り平日は予備日として工程間に分散
- チェックした工事行だけをPDFに上詰めで出力
- 自由工事欄を追加し、任意の工事項目をPDF出力可能
- 各工事項目の開始日・終了日を手動修正可能
- 土日も表示し、工事バーは平日のみに表示するガントチャート風PDF出力
- PDFタイトルを「リフォーム工事工程表」に変更
- PDFファイル名に出力日時を付け、同名上書きを回避
- PDF作成後に保存先フォルダを自動で開く
- 会社支店名、担当者名、連絡先を次回起動時に自動入力
- PDF提出物からDAKEのコピーライトやブランド表記を除外

## booth_product生成用情報

- 商品名: マンション工程表
- 価格案: 500円
- 商品紹介文: マンション管理会社へ提出するための、リフォーム工事工程表を作成します。
工事進捗管理ではなく、管理会社提出用の工程表作成アプリです。

・土日を含む45日間の工期内で、合計28営業日の初期工程を自動入力
・土日は表示し、工事バーは平日のみに表示
・チェックした工事だけをPDF出力
・自由工事欄に任意の工事を追加可能
・PDFは出力日時付きファイル名で保存

実務の流れを、
少し静かにするための道具です。
- タグ: PDF
Windows
実務
ツール
仕事効率化
軽量
シンプル
- 商品画像: assets/booth_thumbnail.jpg
- 補助画像: assets/screenshot.jpg
- 作品ファイル: booth_ready/DakeMansion_Schedule.zip
- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Mansion_Schedule_v1.0.0
- BOOTH URL: https://peakheadz.booth.pm/items/8448169

## Store表示用情報

- 商品名: マンション工程表
- キャッチ: リフォーム許可申請用の工程表を、A3横PDFで作成します。
- 説明: マンション管理会社へ提出するための、リフォーム工事工程表を作成します。
工事進捗管理ではなく、管理会社提出用の工程表作成アプリです。

・土日を含む45日間の工期内で、合計28営業日の初期工程を自動入力
・土日は表示し、工事バーは平日のみに表示
・チェックした工事だけをPDF出力
・自由工事欄に任意の工事を追加可能
・PDFは出力日時付きファイル名で保存

実務の流れを、
少し静かにするための道具です。
- 価格: 500円
- 画像: assets/booth_thumbnail.jpg（存在: あり）
- ダウンロード導線: 未確定
- サポート方針: 既存ファイルに記載なし
- Stripe Payment Link: https://buy.stripe.com/bJefZh57H7cb7OwcY10gw0r
- Store雋ｩ螢ｲ迥ｶ諷・ stripe_ready

## 価格・販売方針

- BOOTH価格案: 500円
- BOOTH URL: https://peakheadz.booth.pm/items/8448169
- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Mansion_Schedule_v1.0.0
- Store販売: 未確定

## 配布・ダウンロード方針

- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Mansion_Schedule_v1.0.0
- BOOTH: https://peakheadz.booth.pm/items/8448169
- BOOTH配布zip: booth_ready/DakeMansion_Schedule.zip
- Store配布導線: 未確定

## 免責・注意事項

- 【注意事項】
- Windows向けアプリです
- ご利用は自己責任でお願いいたします
- 大切なファイルは事前にバックアップを推奨します
- 本ソフトウェアの無断転載・再配布を禁止します
- 環境によっては起動時にWindowsの警告が表示される場合があります
- https://peakheadz.com

## 同梱ファイル方針

- exe: DakeMansion_Schedule.exe
- README.txt: booth_ready/README.txt（存在: あり）
- 注意事項.txt: booth_ready/注意事項.txt（存在: あり）
- 入れないもの: build/、dist/、*.spec、設定ファイル、個人データ、ソース一式は正式配布zipへ混ぜない。

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
