# ORIGINAL.md

## 正本宣言

このファイルは `DAKE_Work_Calendar` の真の正本です。

README.md、DAKE_META、release_body.md、booth_product.txt、Store表示などは、このファイルから派生するビューです。

## 基本情報

- app_id: dake_work_calendar
- title: Dake工程カレンダー
- short_title: Dake工程カレンダー
- category: 作業補助
- status: available
- version: 1.0.0（Release URLより）
- price: 500円
- distribution: GitHub Release / BOOTH / dakeapp.com（Storeは未確定）
- target_platform: Windows

## 目的

開始日から終了日までの日付枠を並べ、工程確認用のA4縦PDFを作成できるWindows向けアプリです。

## 対象ユーザー

既存ファイルに明示なし

## 解決する困りごと

- READMEメタ情報とRelease本文を整備。スクリーンショットをassets/screenshot.webpに作成。
- 開始日から終了日までの日付枠を並べ、工程確認用のA4縦PDFを作成できるWindows向けアプリです。

## 主な機能

- 工程カレンダーPDF作成アプリ
- 指定期間の日付枠に対応
- A4縦PDF出力
- Windows向けexe

## 使い方の要点

既存ファイルに記載なし

## CLI連携・外部ツール

- 使用する外部ツール: 該当なし / 既存ファイルに記載なし
- 同梱有無: 既存ファイルに記載なし
- PATH依存: 既存ファイルに記載なし
- エラー時の扱い: 既存ファイルに記載なし
- Codex作業時の注意: CLIや外部ツール情報を追加する場合は、README等の派生ビューだけでなくこのORIGINALへ戻す。

## 対応形式・非対応形式

- 対応入力: 既存ファイル内の関連語: PDF, メール
- 対応出力: 既存ファイルに記載なし
- 非対応: 既存ファイルに記載なし
- 注意: 形式を推測で追加しない。既存READMEまたは実装確認後に更新する。

## やらないこと / 非ゴール

- 本ソフトウェアの無断転載・再配布を禁止します。
- 本ソフトウェアの無断転載・再配布を禁止します

## 設定・ログ・保存方針

- 4. 保存先フォルダを選択します。
- 6. 保存完了ダイアログのOK後、保存フォルダが開きます。
- 保存先
- 担当者名、電話番号、支店名、前回保存先は `dake_work_calendar_config.json` に保存され、次回起動時に再利用されます。
- 初期実装では、工程の自動計算、ガントチャート化、業者別管理、案件データベース化、複数現場管理、クラウド保存、メール送信、印刷直接実行、Excel出力には対応していません。

## 非破壊・上書き禁止方針

- PDF調整：日曜始まり、タイトル簡略化、ページ表記削除、終了日「完工」表示、印刷安全余白を適用
- 大切なファイルは事前にバックアップを推奨します。
- 大切なファイルは事前にバックアップを推奨します

## 公開用説明の元情報

- display_name: Dake工程カレンダー
- site_title: Dake工程カレンダー
- site_description: 開始日から終了日までの日付枠を並べ、工程確認用のA4縦PDFを作成できるWindows向けアプリです。
- launcher_description: 指定期間の日付枠をA4縦PDFに並べます。
- update_summary: READMEメタ情報とRelease本文を整備。スクリーンショットをassets/screenshot.webpに作成。
- BOOTH紹介文: 指定期間の日付枠をA4縦PDFに並べます。

・工程カレンダーPDF作成アプリ
・指定期間の日付枠に対応
・A4縦PDF出力
・Windows向けexe

実務の流れを、
少し静かにするための道具です。

## README生成用情報

- 概要: 開始日から終了日までの日付枠を並べ、工程確認用のA4縦PDFを作成できるWindows向けアプリです。
- 使い方: 既存ファイルに記載なし
- 必要なもの: Windows環境
- 注意: READMEへ出す内容は公開可能な情報に限定する。
- ビルド: 既存READMEまたはbuild.batを参照。

## DAKE_META生成用情報

```json
{
  "app_key": "dake_work_calendar",
  "display_name": "Dake工程カレンダー",
  "launcher_title": "工程カレンダー",
  "launcher_description": "指定期間の日付枠をA4縦PDFに並べます。",
  "site_title": "Dake工程カレンダー",
  "site_description": "開始日から終了日までの日付枠を並べ、工程確認用のA4縦PDFを作成できるWindows向けアプリです。",
  "update_summary": "READMEメタ情報とRelease本文を整備。スクリーンショットをassets/screenshot.webpに作成。",
  "folder_name": "DAKE_Work_Calendar",
  "exe_name": "DakeWork_Calendar.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Work_Calendar_v1.0.0",
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

- 工程カレンダーPDF作成アプリ
- 指定期間の日付枠に対応
- A4縦PDF出力
- Windows向けexe

## booth_product生成用情報

- 商品名: Dake工程カレンダー
- 価格案: 500円
- 商品紹介文: 指定期間の日付枠をA4縦PDFに並べます。

・工程カレンダーPDF作成アプリ
・指定期間の日付枠に対応
・A4縦PDF出力
・Windows向けexe

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
- 作品ファイル: booth_ready/DakeWork_Calendar.zip
- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Work_Calendar_v1.0.0
- BOOTH URL: https://peakheadz.booth.pm/items/8448282

## Store表示用情報

- 商品名: Dake工程カレンダー
- キャッチ: 開始日から終了日までの日付枠を並べ、工程確認用のA4縦PDFを作成できるWindows向けアプリです。
- 説明: 指定期間の日付枠をA4縦PDFに並べます。

・工程カレンダーPDF作成アプリ
・指定期間の日付枠に対応
・A4縦PDF出力
・Windows向けexe

実務の流れを、
少し静かにするための道具です。
- 価格: 500円
- 画像: assets/booth_thumbnail.jpg（存在: あり）
- ダウンロード導線: 未確定
- サポート方針: 既存ファイルに記載なし
- Stripe Payment Link: https://buy.stripe.com/eVq7sLdEd9kjb0I1fj0gw0J
- Store雋ｩ螢ｲ迥ｶ諷・ stripe_ready

## 価格・販売方針

- BOOTH価格案: 500円
- BOOTH URL: https://peakheadz.booth.pm/items/8448282
- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Work_Calendar_v1.0.0
- Store販売: 未確定

## 配布・ダウンロード方針

- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Work_Calendar_v1.0.0
- BOOTH: https://peakheadz.booth.pm/items/8448282
- BOOTH配布zip: booth_ready/DakeWork_Calendar.zip
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

- exe: DakeWork_Calendar.exe
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
