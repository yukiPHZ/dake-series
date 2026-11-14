# ORIGINAL.md

## 正本宣言

このファイルは `DAKE_PDF_Insert` の真の正本です。

README.md、DAKE_META、release_body.md、booth_product.txt、Store表示などは、このファイルから派生するビューです。

## 基本情報

- app_id: `dake_pdf_insert`
- title: DakePDF差し込み
- internal_app_name: DakePDF差し込み
- display_name: PDF差し込み
- short_title: PDF差し込み
- folder_name: DAKE_PDF_Insert
- exe_name: DakePDF_Insert.exe
- category: PDF
- status: draft
- version: 0.1.0
- payment_status: 未設定
- price: 未設定
- distribution: 未設定
- target_platform: Windows

## 一言説明

元PDFの好きなページ間に、別のPDFを1ファイル差し込んで保存するWindows向け単機能アプリです。

## 目的

PDF編集ソフトを開かずに、「このPDFの、このページとこのページの間に、別PDFを入れたい」という実務上の小さな作業を最短で終わらせる。

## 対象ユーザー

- WindowsでPDF差し込み作業を短く終わらせたい人
- PDF編集ソフトを開くほどではない小さな修正をしたい人
- DAKEシリーズの単機能ツールとして迷わず使いたい人

## 解決する困りごと

元PDFのページ順を壊さず、指定した1か所へ別PDFの全ページを差し込んだ新しいPDFを作る。

## 基本機能

- 元PDFを1ファイル指定する
- 差し込むPDFを1ファイル指定する
- 元PDFのサムネイルを見ながら差し込み位置を1か所選ぶ
- 先頭、各ページの後、末尾を選べる
- 差し込みPDFの全ページを指定位置へ入れる
- 新しいPDFとして保存する
- 保存完了後に完了ダイアログを表示し、保存フォルダを開く

## 入力

- 元PDF: 1ファイル
- 差し込むPDF: 1ファイル

ファイル選択とドラッグ＆ドロップに対応する。

## 出力

- 差し込み済みの新しいPDF
- 元PDFと差し込みPDFは変更しない

## 初版の機能範囲

```text
元PDF
+
差し込むPDF
+
差し込み位置1か所
↓
保存
```

差し込み位置は、内部的には「差し込み前に元PDFから何ページ入れるか」で管理する。

- 先頭: 0
- 1ページ目の後: 1
- 2ページ目の後: 2
- 最終ページの後: 元PDFページ数

## 初版で実装しない機能

- 複数の差し込みPDF
- 複数箇所への一括差し込み
- 差し込みPDF内のページ選択
- PDFページの並べ替え
- ページ削除
- PDF編集
- OCR
- AIによるファイル名判定
- クラウド保存
- 履歴管理
- 複雑な設定画面
- アニメーション
- ドラッグマウスオーバー時の演出

## UI方針

- 既存DAKE PDF系アプリの白 / ライトグレー基調を使う
- DAKE共通ヘッダー、フッター、共通アイコンを使う
- UI文言は `APP_NAME`、`WINDOW_TITLE`、`COPYRIGHT`、`UI_TEXT` に集約する
- 元PDFと差し込みPDFを明確に分ける
- 元PDFページをサムネイル表示する
- ページ間の「＋」とページカードの「この後へ」で差し込み位置を選ぶ
- 現在の位置を `選択中：4ページ目の後` のように明示する
- 完成後ページ数を表示する
- サムネイルサイズ変更はPDF再レンダリングせず、キャッシュ画像をリサイズする
- PDF解析、サムネイル生成、保存処理はUIスレッドで行わない

## 出力ファイル名ルール

初期出力名:

```text
元ファイル名_差し込み済み.pdf
```

同名ファイルが存在する場合は、`_2`、`_3` のように連番を付け、意図しない上書きを避ける。

## PDF処理方針

- ページ結合と書き出しは `pypdf` を使う
- ページ数取得とサムネイル生成は `PyMuPDF` を使う
- サムネイル表示は `Pillow` と `ImageTk` を使う
- ドラッグ＆ドロップは `tkinterdnd2` を使い、未導入環境ではファイル選択で動作する
- 暗号化・パスワード保護PDFへの完全対応は行わない

## 非破壊・上書き禁止方針

- 元PDFを変更しない
- 差し込みPDFを変更しない
- 出力PDFは新規ファイルとして保存する
- 同名出力は連番で回避する
- 元PDFと差し込みPDFに同一ファイルが指定された場合は、誤操作防止のためエラー表示する

## ステータス方針

最低限、以下を区別する。

- 未選択
- 読み込み中
- 準備完了
- 差し込み中
- 保存中
- 保存完了
- エラー

処理中の補助表示は `処理中.`、`処理中..`、`処理中...` を使用する。

## GitHub情報欄

- release_url: 未設定
- repository: https://github.com/yukiPHZ/dake-series
- release_tag: 未設定

## BOOTH情報欄

- BOOTH URL: 未設定
- 価格: 未設定
- 商品画像: 未設定
- booth_product.txt: 派生ビューとして作成。公開情報は未設定のまま扱う。

## Store情報欄

- Store URL: 未設定
- Stripe Payment Link: 未設定
- payment_status: 未設定

## README生成用情報

- 概要: 元PDFの好きなページ間に、別のPDFを1ファイル差し込んで保存するWindows向け単機能アプリです。
- 使い方: 元PDFを選び、差し込むPDFを選び、元PDFサムネイルから差し込み位置を選び、保存する。
- 必要なもの: Windows環境
- 注意: 初回起動時にWindowsの確認画面が表示される場合があります
- ビルド: `build.bat` を実行する

## DAKE_META生成用情報

```json
{
  "app_key": "dake_pdf_insert",
  "display_name": "DakePDF差し込み",
  "launcher_title": "PDF差し込み",
  "launcher_description": "元PDFの好きなページ間に、別のPDFを差し込みます。",
  "site_title": "DakePDF差し込み",
  "site_description": "元PDFの好きなページ間に、別のPDFを1ファイル差し込んで保存するWindows向け単機能アプリです。",
  "update_summary": "新規作成。元PDF、差し込みPDF、差し込み位置1か所を指定して保存できる初版。",
  "folder_name": "DAKE_PDF_Insert",
  "exe_name": "DakePDF_Insert.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "draft",
  "show_in_launcher": false,
  "show_on_site": false,
  "app_type": "market",
  "completion_goal": "formal_release",
  "payment_status": "unset"
}
```

## release_body生成用情報

- 元PDFの任意のページ間に別PDFを差し込み
- サムネイルで差し込み位置を選択
- 先頭 / 各ページの後 / 末尾に対応
- 新しいPDFとして保存
- Windows向けexe

## booth_product生成用情報

- 商品名: DakePDF差し込み
- 価格案: 未設定
- 商品紹介文: 元PDFの好きなページ間に、別のPDFを差し込んで保存します。
- タグ:
  - PDF
  - Windows
  - 実務
  - ツール
  - 仕事効率化
  - 軽量
  - シンプル
- 商品画像: 未設定
- 補助画像: 未設定
- 作品ファイル: 未設定
- GitHub Release: 未設定
- BOOTH URL: 未設定

## 配布・ダウンロード方針

- GitHub Release: 未設定
- BOOTH: 未設定
- dakeapp.com: 未設定
- Launcher: 未設定
- Store: 未設定

## 免責・注意事項

```text
【注意事項】

・Windows向けアプリです
・ご利用は自己責任でお願いいたします
・大切なファイルは事前にバックアップを推奨します
・本ソフトウェアの無断転載・再配布を禁止します
・環境によっては起動時にWindowsの警告が表示される場合があります

PEAKHEADZ
https://peakheadz.com
```

## 同梱ファイル方針

- exe: DakePDF_Insert.exe
- README.txt: 正式出荷時に作成
- 注意事項.txt: 正式出荷時に作成
- 入れないもの: ソースコード、build、dist、spec、個人設定ファイル、APIキー、個人情報

## スクリーンショット・画像方針

- assets/screenshot.webp: 未設定
- assets/screenshot.jpg: 未設定
- assets/booth_thumbnail.jpg: 未設定
- Store用画像: 未設定

## 今後の改善予定

- 正式出荷時にスクリーンショット、BOOTH素材、Release情報を整える
- `ORIGINAL.md` からREADME / release_body / booth_productを生成・同期する仕組みに寄せる

## Codex作業時の注意

- 触ってよい: ORIGINAL.md、main.py、README.md、release_body.md、booth_product.txt、requirements.txt、build.bat
- 触らない: 既存PDF系アプリ、既存出荷済みアプリ、Store/Stripe実装
- 外部公開しない: 個人情報、APIキー、未公開メモ
- 自動操作しない: BOOTH公開、Store決済、GitHub Release更新

## 派生物一覧

- README.md: GitHub公開用ビュー
- DAKE_META: README内の機械利用ビュー
- release_body.md: GitHub Release用ビュー
- booth_product.txt: BOOTH登録用ビュー
- Store: 未構築。Store専用正本は作らず、ORIGINAL由来のgeneratedを読む方針

## 参照した既存アプリ・実装

- `DAKE_PDF_Merge_Mini`: PDF保存、完了ダイアログ、保存フォルダを開く処理
- `DAKE_PDF_SplitSelect`: サムネイル表示、非同期ロード、queue / after
- `DAKE_PDF_Reorder`: PDFページ数取得、サムネイル生成、ステータス、フッター、ドラッグ＆ドロップ
- `DAKE_PDF_Merge`: DAKE共通フッター、処理中ステータス、保存処理
