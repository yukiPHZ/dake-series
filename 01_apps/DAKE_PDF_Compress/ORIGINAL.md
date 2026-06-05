# ORIGINAL.md

## 正本宣言

このファイルは `DAKE_PDF_Compress` の真の正本です。

README.md、DAKE_META、release_body.md、booth_product.txt、Store表示などは、このファイルから派生するビューです。

## 基本情報

- app_id: dake_pdf_compress
- title: DakePDF圧縮
- short_title: PDF圧縮
- category: PDF / 圧縮
- status: available
- version: 1.0.0
- price: 500円
- distribution: GitHub ReleaseとBOOTHで配布する。SHIMARISU連携CLI対象。
- target_platform: Windows

## 目的

PDFを追加して、元PDFを上書きせずにファイルサイズを軽くした別名PDFとして保存する。

## 対象ユーザー

- メール添付や共有前にPDFを軽くしたい人
- 元PDFを壊さず、別名保存で圧縮結果を確認したい人
- SHIMARISUからPDF軽量化を呼び出したい人

## 解決する困りごと

- PDFの容量が大きく、送付や共有に時間がかかる
- 圧縮時に元PDFを上書きしてしまう不安がある
- 圧縮効果がない場合に、どのファイルを使えばよいか迷う

## 主な機能

- PDF圧縮アプリ
- v2のしっかり圧縮
- 別名保存・自動連番対応
- SHIMARISU連携CLI対応
- Windows向けexe

## 使い方の要点

- PDFをドラッグ＆ドロップする。
- ファイル名、元サイズ、保存予定ファイル名を確認する。
- 圧縮して保存する。
- 完了後、保存先フォルダを確認する。

## 公開用説明の元情報

PDFを1つ追加し、ファイルサイズを軽くした別名PDFとして保存できるWindows向けアプリです。

PDFを追加してファイルサイズを軽くします。

実務の流れを、少し静かにするための道具です。

## README生成用情報

- 概要: PDFを1つ追加し、ファイルサイズを軽くした別名PDFとして保存できるWindows向けアプリです。
- 使い方: PDFをドラッグ＆ドロップする。
- 必要なもの: Windows環境。
- 注意: 元PDFは上書きしない。
- ビルド: `build.bat` を実行し、`dist/DakePDF_Compress.exe` を生成する。

## DAKE_META生成用情報

既存README内の `DAKE_META` ブロックを元にした機械利用ビュー情報です。
単独の `DAKE_META` ファイルは既存ファイルに存在しません。


```json
{
  "app_key": "dake_pdf_compress",
  "display_name": "DakePDF圧縮",
  "launcher_title": "PDF圧縮",
  "launcher_description": "PDFを追加してファイルサイズを軽くします。",
  "site_title": "DakePDF圧縮",
  "site_description": "PDFを1つ追加し、ファイルサイズを軽くした別名PDFとして保存できるWindows向けアプリです。",
  "update_summary": "READMEメタ情報とRelease本文を整備。スクリーンショットをassets/screenshot.webpに作成。",
  "folder_name": "DAKE_PDF_Compress",
  "exe_name": "DakePDF_Compress.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_PDF_Compress_v1.0.0",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## release_body生成用情報

- PDF圧縮アプリ
- v2のしっかり圧縮
- 別名保存・自動連番対応
- SHIMARISU連携CLI対応
- Windows向けexe

## booth_product生成用情報

- 商品名: DakePDF圧縮
- 価格案: 500円
- 商品紹介文: PDFを追加してファイルサイズを軽くします。
- 補足紹介文:
  - PDF圧縮アプリ
  - v2のしっかり圧縮
  - 別名保存・自動連番対応
  - SHIMARISU連携CLI対応
  - Windows向けexe
  - 実務の流れを、少し静かにするための道具です。
- タグ:
  - PDF
  - Windows
  - 実務
  - ツール
  - 仕事効率化
  - 軽量
  - シンプル
- 商品画像: assets/booth_thumbnail.jpg
- 補助画像: assets/screenshot.jpg
- 作品ファイル: booth_ready/DakePDF_Compress.zip
- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_PDF_Compress_v1.0.0
- BOOTH URL: https://peakheadz.booth.pm/items/8448178

## Store表示用情報

- 商品名: DakePDF圧縮
- キャッチ: PDFを追加してファイルサイズを軽くします。
- キャッチ補足: 実務の流れを、少し静かにするための道具です。
- 説明: PDFを1つ追加し、ファイルサイズを軽くした別名PDFとして保存できるWindows向けアプリです。
- 価格: 500円
- 画像: assets/booth_thumbnail.jpg / assets/screenshot.webp
- ダウンロード導線: 未確定
- サポート方針: 既存ファイルに記載なし
- Stripe Payment Link: https://buy.stripe.com/aFa6oHeIh3ZZ6Kse250gw02
- Store販売状態: stripe_ready

Storeは未構築のため、Store専用の商品正本は作りません。

## 価格・販売方針

- BOOTH価格案: 500円
- BOOTH URL: https://peakheadz.booth.pm/items/8448178
- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_PDF_Compress_v1.0.0
- Store販売: 未確定

## 配布・ダウンロード方針

- GitHub Releaseで `DakePDF_Compress.exe` を配布する。
- BOOTHでは `booth_ready/DakePDF_Compress.zip` を作品ファイルとして使う。
- dakeapp.com掲載対象です。
- Store配布導線は未確定です。

## 免責・注意事項

BOOTH ready内の注意事項、または既存READMEの注意事項を元にします。

- Windows向けアプリです。
- ご利用は自己責任でお願いいたします。
- 大切なファイルは事前にバックアップを推奨します。
- 本ソフトウェアの無断転載・再配布を禁止します。

## 同梱ファイル方針

- exe: DakePDF_Compress.exe
- README.txt: booth_ready/README.txt (既存)
- 注意事項.txt: booth_ready/注意事項.txt (既存)
- 配布zip: booth_ready/DakePDF_Compress.zip
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
