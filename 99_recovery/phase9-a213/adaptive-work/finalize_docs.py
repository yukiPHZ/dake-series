from pathlib import Path
import json
import hashlib
import platform
import sys

ROOT=Path(__file__).parent
APP=ROOT/'01_apps/DAKE_PDF_Compress'
E=ROOT/'evidence'
docs=APP/'dev';docs.mkdir(exist_ok=True)
rows=json.loads((E/'benchmark.json').read_text(encoding='utf-8'))
for name in ['benchmark.json','cli.json','gui.json','startup-legacy.json','startup-new.json','calibration.json']:
    if (E/name).exists():
        text=(E/name).read_text(encoding='utf-8').replace(str(ROOT).replace('\\','\\\\'),'LOCAL_WORK')
        (docs/name).write_text(text,encoding='utf-8')

lines=['# Adaptive Compression 検証記録（正式出荷前）','',
       '実施日: 2026-09-12。Windows 11 26200 / CPython 3.12.14 x64 / PyMuPDF 1.28.2 / tkinterdnd2 0.4.3 / PyInstaller 6.20.0。',
       'GitHub Release、BOOTH、Web、Storeへの公開は実施しない。元PDFは全件SHA-256一致。','',
       '## ベンチマーク','',
       '時間はソース版の解析・候補プロセス起動・品質比較・保存を含む実測。旧ロジックは変更前main.pyを同じPyMuPDF 1.28.2で動かした内蔵処理で、旧配布exeの性能値ではない。',
       '他の開発作業が同じPCで並行していたため、時間は参考値。値は中央値ではなく記録した最終実行の1回分。','',
       '|PDF|頁|元bytes|採用bytes|削減率|秒|採用候補|最大視覚MAE|旧bytes|旧秒|旧品質|',
       '|---|---:|---:|---:|---:|---:|---|---:|---:|---:|---|']
for r in rows:
    chosen=r['selected'];size=chosen['size'] if chosen else r['original']
    rate=(1-size/r['original'])*100
    lines.append(f"|{r['file']}|{r['profile']['pages']}|{r['original']:,}|{size:,}|{rate:.2f}%|{r['seconds']:.2f}|{chosen['name'] if chosen else '出力なし'}|{chosen['worst']['mae'] if chosen else 0:.6f}|{r['legacy']['size']:,}|{r['legacy']['seconds']:.2f}|{'棄却相当' if 'rejected' in r['legacy'] else 'PASS'}|")
lines += ['',
    '- A: 架空契約文。日本語、5/7pt英数字、ベクター細線、印影を含む。',
    '- B_photo_like: 固定乱数の連続階調・テクスチャを持つ合成風景。実写真ではない。B_NASA_photoはNASA公開の地球写真を3頁に配置した補完ケース。元のJPEGが既に小さく、無理な圧縮を避ける例。',
    '- C: 300dpiの真の1-bitスキャン。D: 300dpiのカラー文書スキャン。E: 300dpiグレーの図面・細線・印影。',
    '- F: Bで予め圧縮済み。サイズが減らないためファイルを作らず、CLIは従来どおりexit 1。',
    '- G/H: 60/100頁。画像は同じ架空資料を再使用し、頁番号は各頁固有。重複排除に有利なストレステストであり、一般の実文書の圧縮率を代表しない。',
    '- 高い削減率は未圧縮のテスト画像・繰返しコンテンツによる。顧客文書や実務PDF全体に同じ圧縮率を保証しない。','',
    '## 品質と選択','',
    '- 解析: 全頁のテキスト・ページ寸法・画像メタデータ、最大12頁の画像配置と低解像度サムネイル。画像数・頁内被覆率・画像のある頁比率・色種・容量/頁を記録する。被覆率は重なりを上限1に丸めたサンプル推定。',
    '- A: garbage=4 / clean / deflate / object stream。B: 写真向け220→150dpi, JPEG82。C_document: 文書向け320→300dpi, JPEG92。C_bitonal_fax: 二値のみFAX、解像度変更なし。',
    '- グレーJPEG再圧縮は不採用。Eで描画の大幅な変化が再現したため、最終実装ではgray=Falseを明示し、グレーのみの文書にはその候補を生成しない。',
    '- FAXは候補として採用。300dpi・全3頁のレンダリングで元と完全一致し、細線・印影の拡大PNGも確認。今回はAのほうが小さく、最終選択はA。bitonal=Trueだけで済ませず、PdfImageRewriterOptionsでFAXのみを指定する。',
    '- D: Ghostscript検出時だけ比較に追加。検出=優先採用にしない。注釈・フォーム・リンク・添付のある文書にはpdfwriteを使用しない。このPCでは未検出。実GSの圧縮品質は未測定。未検出と起動失敗からAへの継続は試験済み。',
    '- 全候補を開き直し、全頁のページ数・回転・MediaBox/CropBox・テキスト指紋・注釈等の件数・目次・添付件数を照合。先頭・中央・末尾＋最大2危険頁を96dpi（長辺最大1400px）で比較。',
    '- 画素差の最大チャネルMAE、差32超の画素率、濃い画素の消失率、64px区画ごとの局所消失率を使用。numpy/scipy/OpenCV不要。',
    '- 閾値: 写真MAE .025 / 変化率 .12 / ink_loss .12 / 局所 .45。文書 .008 / .045 / .035 / 局所 .20。A/FAXは最初の3項目 .00001。',
    '- 上記閾値はfixtureと文字/印影/細線を意図的に消した不良対照で確認した工学的な初期値。OCRや全頁の知覚的保証ではない。未抽出頁の微小欠落、読順や署名の意味的有効性まで証明しない。人手確認を必須とする。',
    '- 品質検証を通過し元より小さいものだけを得点化。score=削減率(0..1) − MAE×4 − ink_loss×0.5 − min(秒/120,1)×.03 + ロスレス優遇.015。',
    '- Aで90%以上軽くなり64KiB/頁未満なら、追加の非可逆候補を省略する。100頁fixtureで約45秒→最終約2秒。単独の初期再測定では約1.2秒。',
    '- 各候補60秒、候補生成全体120秒の予算。ネイティブ処理は別プロセスで期限を監視する。GUI検証30秒、圧縮全体160秒のwatchdog。','',
    '## 回帰','',
    '- python -m py_compile main.py adaptive.py: PASS。',
    '- tests/test_regression.py: 8試験PASS（ページ/文字/寸法/順序の破壊、ラスタ印影・細線の欠落、元保護・連番、品質優先、rewrite_images欠落、GS起動失敗、期限、UI_TEXT）。',
    '- CLI: ソース/exe各8ケースPASS（help、1件、複数、日本語・空白パス、存在しないPDF、非PDF、破損、入力なし、効果なし。日本語パスは1件/複数ケースに含む）。stdoutはパスのみ、エラーは1行・exit1・Tracebackなし。DEBUG有効でもstdoutへログを混入しない。',
    '- SHIMARISUのsubprocess.runがUTF-8で読むことをread-onlyで確認。圧縮exe側でUTF-8を明示し、SHIMARISUには変更なし。',
    '- 実Tkの統合試験: 起動860x740/min760x720、DnD callback、確認中→準備完了、最新1件優先、古いエラー破棄、選択ダイアログ経由、圧縮・再圧縮・連番、完了、フォルダOpen呼出、クリア、エラー、狭幅/広幅フッター、worker終了がPASS。',
    '- GUI: 1個の常駐PDF workerと、候補ごとの一時worker最大1個。Tkは親プロセスだけ。ウインドウを閉じる操作が圧縮中なら、保存処理を完了してから終了する。',
    '- 実exeのネイティブ操作: 通常起動、共通アイコン、ファイル選択、圧縮中phase、94.9%完了ダイアログ、OK後Explorerで保存先を開くことを目視確認。',
    '- Explorerからの物理DnD、再起動後cold起動、複数標準PDFビューアで全fixtureの閲覧は人手Gate。DnD登録とイベント経由は自動試験済み。','',
    '## 起動','',
    'tests/startup_probe.pyでexe起動直前から、Tkウインドウがmappedになりevent loopを実行した時刻まで測定。測定時だけ環境変数でreadyファイルを出し50ms後に閉じる。通常操作に追加UIなし。',
    '旧比較用は変更前main.pyに同じ計測hookのみを加え、同じPython/PyInstaller/依存でonefile再ビルド。元の配布exeそのものの測定ではない。旧配布exeは通常起動・見た目を確認済み。']
for name,label in [('startup-legacy.json','旧ロジック再ビルド'),('startup-new.json','新版')]:
    if (E/name).exists():
        x=json.loads((E/name).read_text())
        lines += ['',f"- {label}: 初回 {x['runs'][0]['seconds']:.3f}秒、warm 5回 {', '.join(f'{r["seconds"]:.3f}' for r in x['runs'][1:])}秒、warm中央値 {x['warm_median']:.3f}秒。"]
lines += ['', 'これは「初回プロセス起動」とwarmであり、OSファイルキャッシュを消したcold測定ではない。ユーザーの作業中PCを再起動・キャッシュ消去していないため、真のcold startは未測定。onefile維持。','',
    '## iLovePDF','',
    '比較結果は未提供。自動アップロード・外部送信は一切していない。手動の「推奨圧縮」結果を受領した後で以下へ追記する。','',
    '|テストPDF|元bytes|DAKE bytes/削減率|iLovePDF bytes/削減率|文字可読性|画像品質|図面品質|処理秒|',
    '|---|---:|---|---|---|---|---|---|',
    '|各fixture|上表|上表|未提供|人手比較待ち|人手比較待ち|人手比較待ち|手動結果待ち|','',
    '## 再現','',
    '```powershell',
    'python -m pip install -r requirements.txt',
    'python tests/benchmark.py --fixtures tests/output/fixtures --generate',
    'python tests/benchmark.py --fixtures tests/output/fixtures --results tests/output/benchmark.json',
    'python tests/test_regression.py -v',
    'python tests/gui_regression.py tests/output/fixtures tests/output/gui.json',
    'build.bat --no-pause',
    'python tests/startup_probe.py dist/DakePDF_Compress.exe tests/output/startup.json',
    '```','',
    'NASA補完写真は [NASA PIA18033](https://science.nasa.gov/photojournal/earth/)（NASAの衛星写真合成、2012）を開発用に取得。生成器はネット接続しない。NASA素材はGitには含めない。','',
    'API根拠: [rewrite_images・保存](https://pymupdf.readthedocs.io/en/latest/document.html)、[プロセス分離](https://pymupdf.readthedocs.io/en/latest/recipes-multiprocessing.html)。','',
    '## Human Gate','',
    '1. 最終exeで起動・物理DnD・確認中の即応・フェーズ表示・完了の強弱を確認する。',
    '2. human_gate/fixturesとoutputsを標準ビューアで比較。5pt文字、日本語、0.15pt細線、印影、頁順を確認する。',
    '3. 個人情報を外部へ送らず手元の実務PDFで確認。サンプル外の文字/図面/印影が崩れないこと。',
    '4. 再起動後cold 3回とwarm 5回を測る。iLovePDF比較を希望する場合のみ、テスト専用PDFを人手で扱う。',
    '5. Gate通過後の正式version案は1.1.0。今回は正本version1.0.0と既存Release tagを維持。','']
(docs/'ADAPTIVE_BENCHMARK.md').write_text('\n'.join(lines),encoding='utf-8')

p=APP/'ORIGINAL.md';s=p.read_text(encoding='utf-8')
s=s.replace('- v2のしっかり圧縮','- PDFの性質に応じた自動圧縮（Adaptive Compression）')
s=s.replace('- 触ってよい: ORIGINAL.mdの更新、派生ビューとの整合確認。','- 今回の開発範囲: Adaptive Compression、GUI操作品質、依存固定、build、ベンチマーク、回帰、commit/push。')
s=s.replace('- 触らない: main.py、build.bat、assets、dist、booth_readyの内容をこのPhaseで変更しない。','- 対象外: SHIMARISU、共通core、販売assets、booth_ready。他アプリの差分を巻き込まない。人手確認用distは再ビルドする。')
s += '''
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
'''
p.write_text(s,encoding='utf-8')
p=APP/'README.md';s=p.read_text(encoding='utf-8')
a=s.index('## v2 しっかり圧縮');b=s.index('## ビルド方法',a)
s=s[:a]+'''## 自動で、読みやすく軽く

- PDFの内容に合わせて、画質と容量のバランスがよい結果を自動で選びます。設定は不要です。
- 追加すると「PDFを確認中...」、準備ができると「圧縮できます」と表示します。
- 処理中は現在の作業を表示し、完了するとサイズの変化と削減率が分かります。
- 元PDFは変更せず、同じフォルダへ別名保存します。十分に軽いPDFは保存しない場合があります。
- 今回の改善は正式出荷前です。人手確認後に公開します。

''' +s[b:]
s=s.replace('PyMuPDF', 'PyMuPDF')
s=s.replace('事前に依存ライブラリをインストールします。','Python 3.10以上を使用し、検証済みの依存をインストールします（Windows / Python 3.12.14で検証、PyMuPDF 1.28.2に固定）。')
s=s.replace('共通アイコン `..\\..\\02_assets\\dake_icon.ico` が存在する場合は自動で使用します。','共通アイコン `..\\..\\02_assets\\dake_icon.ico` を埋め込みます。onefileを維持しています。')
s=s.replace('- 圧縮処理は別スレッドで実行し、処理中ステータスとプログレス表示でUIが固まって見えないようにしています。','- PDFの確認と圧縮は別プロセスで実行し、現在の作業を表示します。')
s=s.replace('- 初期実装はPDF 1件のみ対応です。','- GUIはPDF1件、CLIは複数件に対応します。')
s=s.replace('- v2のしっかり圧縮','- 内容に合わせた自動圧縮')
s=s.replace('## 起動確認結果','## 開発検証\n\n2026-09-12: Adaptive Compression、非同期追加、完了表示を改善。検証記録と制限は [開発用レポート](dev/ADAPTIVE_BENCHMARK.md) を参照してください。ログはアプリフォルダの `logs/` に保存します。\n\n## 過去の確認履歴（v2は当時の説明名）')
p.write_text(s,encoding='utf-8')
(APP/'.gitignore').write_text('build/\ndist/\n*.spec\n*_config.json\n__pycache__/\n*.pyc\nlogs/\nhuman_gate/\ntests/output/\n',encoding='utf-8')
print('Documentation generated')
