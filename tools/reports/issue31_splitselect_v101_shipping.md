# Issue #31 - DakePDF分割Select v1.0.1 正式出荷

## 判定

**CLOSED（2026-09-27 JST）。** Issue #44 / PR #45 のPDFium移行がChatGPTレビューPASSを受け、PR #45をmergeした。PyMuPDF再配布ライセンスのHOLD理由は、製品・build・配布物から当該依存を除去したことで解消した。

## 正式ソースと配布物

- 正式ローカルrepo: `C:\Users\yukiz\devlop\DAKE_series`
- PR #45 merge commit: `af1cfa52983be402216c113994fee503f227cf90`
- v1.0.1正本・派生ビュー・生成データ commit: `ac107ca35c606f6d780216924f4f5c8ac18dab27`
- onedir EXE: `01_apps/DAKE_PDF_SplitSelect/dist/DakePDF_Split_Select/DakePDF_Split_Select.exe`
- 正式ZIP: `01_apps/DAKE_PDF_SplitSelect/booth_ready/DakePDF_Split_Select.zip`
- ZIP: 17,368,580 bytes, SHA-256 `3C11FC56510603BC61CA3918E9F78DC6F83F8A383D722CDCE8419743D4BF2268`
- EXE SHA-256: `69E6FD20091E9A3778C7260461A1E85659BA49BD3D7CB7FFBF29BA70A8A66DF4`
- EXE FileVersion / ProductVersion: 1.0.1
- 旧v1.0.0 Releaseは変更せず保持。

## Build・ライセンス・動作

- 正式ルートのclean Python 3.12 venvで依存導入・`pip check`・`build.bat`を実施。
- 正式ZIPは1,093ファイル（runtime 1,091 + README.txt / 注意事項.txt）。展開後のruntime 1,091ファイルはonedir元ファイルとSHA-256一致。
- ZIP内の`pdfium.dll`は1件、`fitz` / `PyMuPDF` / `PyMuPDFb` / Pillow / PILに該当するpayloadは0件。PyInstaller TOCも旧renderer payload 0件。
- pypdfium2 5.13.0の実インストールwheel `License-File` 19件を、同梱19件とファイル単位のSHA-256で照合。`THIRD_PARTY_NOTICES.txt`も同梱。
- `unittest discover -s tests -v`: 10件PASS。3/30/100/300ページ、visible-first、LRU、F5、Ctrl+O、保存・CLIを含む。
- 展開ZIPのEXEでCLI 9ケースPASS。実Tk GUIで合成3ページPDFを開き、サムネイル・ページ選択・保存・F5・通常終了を確認。保存PDFは選択ページのみ、入力PDFのSHAは不変。
- physical Explorer DnD、DPI 125/150/200%、実スキャンPDFは未確認事項として残す。今回の正式出荷blockerには戻さない。

## 公開先

- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_PDF_SplitSelect_v1.0.1
  - 公開アセットは17,368,580 bytes。GitHubのasset digestはローカルZIP SHA-256と一致。
- BOOTH既存商品: https://peakheadz.booth.pm/items/8448213
  - 価格500円、公開中、カートに入れる操作が表示される。
  - v1.0.1 ZIPをアップロードし、旧ファイルID `9055823` の提供チェックを外した。保存後の管理画面を再読込し、新ファイルID `9754428` のみがeffective downloadableであることを確認。旧アップロードは削除していない。
  - BOOTHからの再ダウンロードSHAは独立検証できていない。アップロード元は上記正式ZIPで、管理画面の提供ID切替と公開ページのv1.0.1説明・価格を確認した。
- dakeapp.com: https://dakeapp.com/apps/pdf-split-select/ （HTTP 200、v1.0.1 Release導線）
- Store: https://store.dakeapp.com/product/?id=dake_pdf_splitselect （HTTP 200、v1.0.1 / 500円 / Stripeで購入 / BOOTHで見るを実画面確認）
- 本番のdakeapp.com / Store生成JSON: 各HTTP 200、56商品を維持、SplitSelect `version=1.0.1`、`payment_status=stripe_ready`、既存BOOTH URLとStripe Payment Linkを維持。`source_policy`と`do_not_edit=true`を確認。
- dakeapp-site commit: `7030a14`。dake-store-site commit: `1115f2d`。両方mainへpushし、Cloudflare本番URLの反映を確認。

## 運用メモ

- `ORIGINAL.md`を正本とし、README / DAKE_META / BOOTH説明 / generated JSON / Storeは派生ビューとして同期した。
- 既存Stripe Product / Price / Payment Link、既存BOOTH商品URLと価格は変更していない。
- 古いローカルonefile EXEは公開v1.0.0アセットとSHA一致を確認したうえで正式distから除き、v1.0.1 onedir EXEを共通出荷処理が選択する状態にした。公開v1.0.0 Releaseは保持。
- 一時worktreeは正式ルートへの同期・成果物照合後に、他作業のworktreeと混同せず整理する。
