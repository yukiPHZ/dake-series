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
- PDFの性質に応じた自動圧縮（Adaptive Compression）
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
- PDFの性質に応じた自動圧縮（Adaptive Compression）
- 別名保存・自動連番対応
- SHIMARISU連携CLI対応
- Windows向けexe

## booth_product生成用情報

- 商品名: DakePDF圧縮
- 価格案: 500円
- 商品紹介文: PDFを追加してファイルサイズを軽くします。
- 補足紹介文:
  - PDF圧縮アプリ
  - PDFの性質に応じた自動圧縮（Adaptive Compression）
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

- 今回の開発範囲: Adaptive Compression、GUI操作品質、依存固定、build、ベンチマーク、回帰、commit/push。
- 対象外: SHIMARISU、共通core、販売assets、booth_ready。他アプリの差分を巻き込まない。人手確認用distは再ビルドする。
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

## Adaptive Compression / 操作品質（2026-09-12実装・正式出荷前）

- 操作は「PDFを追加 → 圧縮して保存」のみ。DPI・品質・モード選択UIを設けない。
- ページ数、画像数と頁内被覆率、テキスト量、画像色種、元容量と容量/頁を解析する。サンプル解析は最大12頁。全頁高解像度レンダリングをしない。
- Aロスレス最適化を基準とし、B写真バランス、C文書/スキャン・二値FAX、D任意Ghostscriptを内部比較する。Bは220→150dpi/JPEG82、Cカラー文書は320→300dpi/JPEG92。
- 検証済み依存はWindows11 / CPython3.12.14 x64 / PyMuPDF==1.28.2。rewrite_imagesを実行時にも確認し、不在を成功扱いしない。無条件の最新版追随はしない。
- 二値FAXは解像度維持、文字認識による置換なし。検証用PDFは300dpi・全3頁で元のレンダリングと完全一致。今回のfixtureではAがより小さく最終採用。
- グレー画像の非可逆再圧縮は、検証で描画破損が出たため無効。グレー文書はロスレス、二値はFAXも候補にする。写真を含む場合もグレー画像を書き換えない。
- 全候補で開けること、非0サイズ、頁数・寸法・回転・全頁テキスト指紋を照合。先頭/中央/末尾＋危険頁最大2頁の中解像度画素差と局所インク消失を検査。不合格候補は採用しない。
- 品質を通過し元より小さい候補を、容量・画質・時間の得点で選ぶ。僅差はロスレス優先。Aだけで90%以上削減かつ64KiB/頁未満なら追加の非可逆処理を省く。
- Ghostscriptは検出時のみ候補。存在するだけでは採用しない。技術事情はdebug logへ記録し、通常UIへ表示しない。
- 元PDFを変更しない。完成済み一時ファイルをatomicかつ既存名を上書きしない形で公開する。*_compressed.pdf、*_compressed_2.pdfの連番を維持する。
- 5%未満は「このPDFは、すでにかなり軽いようです。」と短く伝える。効果なしは保存せず、CLIは既存のexit1を維持する。
- PDF追加直後は「PDFを確認中...」。別プロセスで検証し「圧縮できます」へ移る。常駐worker1個、待機要求は最新1件、世代IDで古い結果を破棄する。圧縮要求は破棄しない。
- 圧縮中は解析・画像最適化・PDF整理・仕上がり確認・保存の実フェーズを優先。処理中だけ不定進捗バーを動かし、停止中は空にする。8秒以後だけ共通英語フレーズを補助欄へ表示する。
- 完了は「元サイズ → 保存後サイズ」「xx.x%軽くなりました」を主情報にする。OK後に保存先フォルダを開く。ライブラリ例外は画面へ出さない。
- SHIMARISU契約維持: --from-shimarisu --inputs、成功exit0/stdout=出力パス（UTF-8、複数時1行1件）、失敗exit1/stderr=短いエラー、Tracebackなし。
- 初期860x740/min760x720、BIZ UDPGothic優先、共通フッター・アイコンとPyInstaller onefileを維持する。
- ベンチマーク概要: テキスト・写真・二値・カラー文書・細線・圧縮済・60頁・100頁を試験。合成風景は約94.9%削減、文書系fixtureは多くを画素差ゼロのAで採用。圧縮済は出力しない。元PDFのハッシュを全件確認。
- 大幅削減は未圧縮/重複画像を含むfixtureでの結果。実務品質・iLovePDFに対する競争力の最終判断は人手Gate。外部送信や自動比較はしない。
- 実測値、閾値、旧ロジック比較、試験範囲と制限は [開発用検証記録](dev/ADAPTIVE_BENCHMARK.md) に保存する。

## バージョンの整理

正本versionと既存Release tagは1.0.0。過去の「v2」は圧縮ロジック更新の説明で、正式SemVerの2.0.0ではなかった。今回の実装は未出荷であり、タグやGitHub Releaseを追加・更新しない。互換CLIを維持した機能改善として、人手Gate後の正式出荷は1.1.0を提案する。
