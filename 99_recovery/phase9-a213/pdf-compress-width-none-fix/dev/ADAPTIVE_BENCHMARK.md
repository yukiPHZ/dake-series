# Adaptive Compression 検証記録（正式出荷前）

実施日: 2026-09-12。Windows 11 26200 / CPython 3.12.14 x64 / PyMuPDF 1.28.2 / tkinterdnd2 0.4.3 / PyInstaller 6.20.0。
GitHub Release、BOOTH、Web、Storeへの公開は実施しない。元PDFは全件SHA-256一致。

## ベンチマーク

時間はソース版の解析・候補プロセス起動・品質比較・保存を含む実測。旧ロジックは変更前main.pyを同じPyMuPDF 1.28.2で動かした内蔵処理で、旧配布exeの性能値ではない。
他の開発作業が同じPCで並行していたため、時間は参考値。値は中央値ではなく記録した最終実行の1回分。

|PDF|頁|元bytes|採用bytes|削減率|秒|採用候補|最大視覚MAE|旧bytes|旧秒|旧品質|
|---|---:|---:|---:|---:|---:|---|---:|---:|---:|---|
|A_text.pdf|3|35,354|3,322|90.60%|0.51|A_lossless|0.000000|3,359|0.01|PASS|
|B_NASA_photo.pdf|3|338,674|337,831|0.25%|1.38|A_lossless|0.000000|337,825|0.34|PASS|
|B_photo_like.pdf|3|6,484,948|330,987|94.90%|6.00|B_balanced|0.005477|330,987|0.67|PASS|
|C_bitonal.pdf|3|3,265,880|32,889|98.99%|1.22|A_lossless|0.000000|32,901|0.06|PASS|
|D_color_scan.pdf|3|26,111,805|473,552|98.19%|4.11|A_lossless|0.000000|466,808|1.22|PASS|
|E_fine_drawing.pdf|3|8,707,048|251,433|97.11%|1.20|A_lossless|0.000000|462,988|1.38|棄却相当|
|F_precompressed.pdf|3|330,987|330,987|0.00%|1.39|出力なし|0.000000|330,987|0.45|PASS|
|G_60_pages.pdf|60|65,309,272|42,404|99.94%|6.58|A_lossless|0.000000|42,632|1.16|PASS|
|H_100_pages.pdf|100|26,160,050|491,252|98.12%|2.05|A_lossless|0.000000|484,653|46.19|PASS|

- A: 架空契約文。日本語、5/7pt英数字、ベクター細線、印影を含む。
- B_photo_like: 固定乱数の連続階調・テクスチャを持つ合成風景。実写真ではない。B_NASA_photoはNASA公開の地球写真を3頁に配置した補完ケース。元のJPEGが既に小さく、無理な圧縮を避ける例。
- C: 300dpiの真の1-bitスキャン。D: 300dpiのカラー文書スキャン。E: 300dpiグレーの図面・細線・印影。
- F: Bで予め圧縮済み。サイズが減らないためファイルを作らず、CLIは従来どおりexit 1。
- G/H: 60/100頁。画像は同じ架空資料を再使用し、頁番号は各頁固有。重複排除に有利なストレステストであり、一般の実文書の圧縮率を代表しない。
- 高い削減率は未圧縮のテスト画像・繰返しコンテンツによる。顧客文書や実務PDF全体に同じ圧縮率を保証しない。

## 品質と選択

- 解析: 全頁のテキスト・ページ寸法・画像メタデータ、最大12頁の画像配置と低解像度サムネイル。画像数・頁内被覆率・画像のある頁比率・色種・容量/頁を記録する。被覆率は重なりを上限1に丸めたサンプル推定。
- A: garbage=4 / clean / deflate / object stream。B: 写真向け220→150dpi, JPEG82。C_document: 文書向け320→300dpi, JPEG92。C_bitonal_fax: 二値のみFAX、解像度変更なし。
- グレーJPEG再圧縮は不採用。Eで描画の大幅な変化が再現したため、最終実装ではgray=Falseを明示し、グレーのみの文書にはその候補を生成しない。
- FAXは候補として採用。300dpi・全3頁のレンダリングで元と完全一致し、細線・印影の拡大PNGも確認。今回はAのほうが小さく、最終選択はA。bitonal=Trueだけで済ませず、PdfImageRewriterOptionsでFAXのみを指定する。
- D: Ghostscript検出時だけ比較に追加。検出=優先採用にしない。注釈・フォーム・リンク・添付のある文書にはpdfwriteを使用しない。このPCでは未検出。実GSの圧縮品質は未測定。未検出と起動失敗からAへの継続は試験済み。
- 全候補を開き直し、全頁のページ数・回転・MediaBox/CropBox・テキスト指紋・注釈等の件数・目次・添付件数を照合。先頭・中央・末尾＋最大2危険頁を96dpi（長辺最大1400px）で比較。
- 画素差の最大チャネルMAE、差32超の画素率、濃い画素の消失率、64px区画ごとの局所消失率を使用。numpy/scipy/OpenCV不要。
- 閾値: 写真MAE .025 / 変化率 .12 / ink_loss .12 / 局所 .45。文書 .008 / .045 / .035 / 局所 .20。A/FAXは最初の3項目 .00001。
- 上記閾値はfixtureと文字/印影/細線を意図的に消した不良対照で確認した工学的な初期値。OCRや全頁の知覚的保証ではない。未抽出頁の微小欠落、読順や署名の意味的有効性まで証明しない。人手確認を必須とする。
- 品質検証を通過し元より小さいものだけを得点化。score=削減率(0..1) − MAE×4 − ink_loss×0.5 − min(秒/120,1)×.03 + ロスレス優遇.015。
- Aで90%以上軽くなり64KiB/頁未満なら、追加の非可逆候補を省略する。100頁fixtureで約45秒→最終約2秒。単独の初期再測定では約1.2秒。
- 各候補60秒、候補生成全体120秒の予算。ネイティブ処理は別プロセスで期限を監視する。GUI検証30秒、圧縮全体160秒のwatchdog。

## 回帰

- python -m py_compile main.py adaptive.py: PASS。
- tests/test_regression.py: 8試験PASS（ページ/文字/寸法/順序の破壊、ラスタ印影・細線の欠落、元保護・連番、品質優先、rewrite_images欠落、GS起動失敗、期限、UI_TEXT）。
- CLI: ソース/exe各8ケースPASS（help、1件、複数、日本語・空白パス、存在しないPDF、非PDF、破損、入力なし、効果なし。日本語パスは1件/複数ケースに含む）。stdoutはパスのみ、エラーは1行・exit1・Tracebackなし。DEBUG有効でもstdoutへログを混入しない。
- SHIMARISUのsubprocess.runがUTF-8で読むことをread-onlyで確認。圧縮exe側でUTF-8を明示し、SHIMARISUには変更なし。
- 実Tkの統合試験: 起動860x740/min760x720、DnD callback、確認中→準備完了、最新1件優先、古いエラー破棄、選択ダイアログ経由、圧縮・再圧縮・連番、完了、フォルダOpen呼出、クリア、エラー、狭幅/広幅フッター、worker終了がPASS。
- GUI: 1個の常駐PDF workerと、候補ごとの一時worker最大1個。Tkは親プロセスだけ。ウインドウを閉じる操作が圧縮中なら、保存処理を完了してから終了する。
- 実exeのネイティブ操作: 通常起動、共通アイコン、ファイル選択、圧縮中phase、94.9%完了ダイアログ、OK後Explorerで保存先を開くことを目視確認。
- Explorerからの物理DnD、再起動後cold起動、複数標準PDFビューアで全fixtureの閲覧は人手Gate。DnD登録とイベント経由は自動試験済み。

## 起動

tests/startup_probe.pyでexe起動直前から、Tkウインドウがmappedになりevent loopを実行した時刻まで測定。測定時だけ環境変数でreadyファイルを出し50ms後に閉じる。通常操作に追加UIなし。
旧比較用は変更前main.pyに同じ計測hookのみを加え、同じPython/PyInstaller/依存でonefile再ビルド。元の配布exeそのものの測定ではない。旧配布exeは通常起動・見た目を確認済み。

- 旧ロジック再ビルド: 初回 2.401秒、warm 5回 1.759, 1.623, 1.513, 1.564, 1.615秒、warm中央値 1.615秒。

- 新版: 初回 1.914秒、warm 5回 1.374, 1.415, 1.431, 1.498, 1.459秒、warm中央値 1.431秒。

これは「初回プロセス起動」とwarmであり、OSファイルキャッシュを消したcold測定ではない。ユーザーの作業中PCを再起動・キャッシュ消去していないため、真のcold startは未測定。onefile維持。

## iLovePDF

比較結果は未提供。自動アップロード・外部送信は一切していない。手動の「推奨圧縮」結果を受領した後で以下へ追記する。

|テストPDF|元bytes|DAKE bytes/削減率|iLovePDF bytes/削減率|文字可読性|画像品質|図面品質|処理秒|
|---|---:|---|---|---|---|---|---|
|各fixture|上表|上表|未提供|人手比較待ち|人手比較待ち|人手比較待ち|手動結果待ち|

## 再現

```powershell
python -m pip install -r requirements.txt
python tests/benchmark.py --fixtures tests/output/fixtures --generate
python tests/benchmark.py --fixtures tests/output/fixtures --results tests/output/benchmark.json
python tests/test_regression.py -v
python tests/gui_regression.py tests/output/fixtures tests/output/gui.json
build.bat --no-pause
python tests/startup_probe.py dist/DakePDF_Compress.exe tests/output/startup.json
```

NASA補完写真は [NASA PIA18033](https://science.nasa.gov/photojournal/earth/)（NASAの衛星写真合成、2012）を開発用に取得。生成器はネット接続しない。NASA素材はGitには含めない。

API根拠: [rewrite_images・保存](https://pymupdf.readthedocs.io/en/latest/document.html)、[プロセス分離](https://pymupdf.readthedocs.io/en/latest/recipes-multiprocessing.html)。

## Human Gate

1. 最終exeで起動・物理DnD・確認中の即応・フェーズ表示・完了の強弱を確認する。
2. human_gate/fixturesとoutputsを標準ビューアで比較。5pt文字、日本語、0.15pt細線、印影、頁順を確認する。
3. 個人情報を外部へ送らず手元の実務PDFで確認。サンプル外の文字/図面/印影が崩れないこと。
4. 再起動後cold 3回とwarm 5回を測る。iLovePDF比較を希望する場合のみ、テスト専用PDFを人手で扱う。
5. Gate通過後の正式version案は1.1.0。今回は正本version1.0.0と既存Release tagを維持。

## Human Gate再現バグ回帰（2026-09-20）

根本原因は、`page.get_drawings()`がfill-only drawingに対して返す`width=None`を、数値として`0 < width < 0.6`で比較していたこと。実10頁PDFで修正前の`TypeError: '<' not supported between instances of 'int' and 'NoneType'`を再現した。

- `is_thin_stroke()`で、boolを除く有限のint/floatだけを評価する。fill-only、widthキーなし、文字列、NaN、Infinityは細線とみなさない。0.3は保護、1.0は非保護。
- 3頁fixtureで、1頁目fill-only（type=f / width=None）、2頁目0.3pt、3頁目1.0ptを生成。`analyze()`は完走し、`protected_pages == [1]`を確認。
- workerからGUIへ内部detailとcontextを維持。開発ログは`stage`、`page_index`、`candidate`、`exception_type`、`exception_message`を保存する。通常GUIは短い日本語のまま。
- 品質閾値、候補計画、Adaptive Compressionの得点、UI文言は変更していない。

|実PDF|頁|元bytes|出力bytes|削減率|実exe秒|profile kind|candidate plan|採用|視覚検証|
|---|---:|---:|---:|---:|---:|---|---|---|---|
|dake-pdf-sample-photo-heavy-a4-10-pages|10|3,200,000|1,668,484|47.86%|6.485|photo|A_lossless, B_balanced|B_balanced|PASS、MAE最大0.000832、changed/ink loss 0|
|結合済み640頁|640|176,840,968|2,765,846|98.44%|4.766|photo|A_lossless, B_balanced|A_lossless|PASS、MAE/changed/ink loss 0|

両方とも出力を再度開き、ページ数10/10、640/640を確認した。640頁はAだけで十分に小さくなったため、既存ルールどおりBを省略。Popplerで先頭・中間・最終の代表頁を描画し、640頁は元と出力のPNG SHA-256も一致した。

最終回帰は、`py_compile`、単体11件、既存8カテゴリbenchmark、ソースCLI 8件、実exe CLI 8件、実Tk GUI、DnD callback/登録、日本語パス、効果なし、図面・細線をPASS。実Explorerから別ウインドウへの物理ドラッグは自動操作APIがウインドウ外座標を拒否するため未実施で、DnDは登録とcallback経路で確認した。正式出荷、Release、BOOTH、Store更新は実施していない。
