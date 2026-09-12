# SplitSelect Performance / Refresh Report

検証日: 2026-09-12 (JST)。対象: `01_apps/DAKE_PDF_SplitSelect`。
この報告のPASSは記載した条件・ケースに限る。未確認をPASSとは扱わない。

## 作業範囲

- `main` を fast-forward で `ec63604` へ同期し、cleanを確認して専用worktreeを作成。
- branch: `codex/pdf-split-select-fast-refresh`。
- 最初に `ORIGINAL.md` を読み、共通仕様・ビルド・Git・UI_TEXT・CLI・操作品質・latest-job規約を確認。
- コード変更より先に正式契約を更新し、`1cdb3f7` (`Define SplitSelect performance and refresh contract`) をcommit。
- UIレイアウト、フッター、配布名、機能の範囲は変更しない。merge / Release / BOOTH / Store / 正式出荷は行わない。

## 変更前と問題点

- GUI起動前にfitz、Pillow、pypdfをimportしていた。
- onefileの展開待ちが毎回発生。buildで毎回pip自体を更新していた。
- 可視範囲の要求に加え、24ページ単位のafterで最終ページまで予約していた。
- 全ページ生成完了がready条件。大量ページでも全画像をキャッシュし続けた。
- generationは旧イベントの表示だけを防止し、workerの古い仕事・開いたPDFは残り得た。
- 大きくレンダーした画像をPillowで縮小していた。

## 変更後

- UIを先に表示。fitzは固定レンダーワーカー内のPDFロード時、pypdfは保存・CLI時のみimport。
- `pdf_backend.py` は標準ライブラリだけで起動し、1ワーカーがfitz.Documentを所有。
- 可視ページを先頭に、近接ページを前後8ページまで要求。スクロールで不要な待機要求を置換する。
- 可視ページの画像がCanvasへ描画された時点でready。全ページ生成フラグ・batch enqueueは削除。
- 152 x 214の枠に収まる倍率で直接RGBレンダー。PPM bytesをqueueで渡し、Tk画像はmain threadで生成。
- 画像キャッシュは上限96のLRU。可視画像を優先して残し、evict画像は再要求できる。選択・範囲は別管理。
- イベントqueueは32件上限、UIは1回8件までを20ms周期で処理。描画afterは16msでまとめる。
- refresh: generation更新、cancel設定、pending破棄、workerへのclose制御要求、UI初期化。既に実行中の1ページだけは完了を許容。
- closeはworker内のみ。新PDFのopenは旧PDFのclose後。旧generation / 古いviewport requestの画像はUIに採用しない。
- worker例外・アプリ通常終了でもclose。保存中のrefresh・PDF置換・終了は抑止し、保存キャンセルには拡張しない。
- 手動保存先を保持、自動保存先・選択・Shiftアンカー・入力・エラー・完了状態は破棄。出力済みファイルは削除しない。
- `Ctrl+O` / `F5` を標準bindで追加。重い演出・スプラッシュなし。

## ビルド環境とcollect-all監査

Windows 11 build 26200、AMD Ryzen 7 5800X (8 cores / 16 logical processors)、Python 3.12.4。
既存のビルド環境を確認し、アップグレードせず以下を固定した。

| package | verified version |
| --- | --- |
| PyMuPDF | 1.24.10 |
| Pillow | 12.3.0 |
| pypdf | 6.10.2 |
| pyinstaller | 6.19.0 |
| pyinstaller-hooks-contrib | 2026.4 |
| tkinterdnd2 | 0.4.3 |

PyMuPDFbは既存環境で1.24.10。完全な新規venvからの再ビルドは未確認。
Pillowはアプリから直接importしなくなったが、既存環境・依存の互換性を優先し今回は削除しない。

`build.bat` はpipの自己更新を廃止。onedirをPRの検証用ビルド方式として採用。
出力は `dist/DakePDF_Split_Select/DakePDF_Split_Select.exe` と `_internal/`。
共通アイコンを同梱し、開発時・packaged実行時とも共通画像を参照する。
既存の配布パイプラインがexe単体を前提としていないか、正式出荷前に別途確認が必要。

fitzのcollect-allを外した試作onedirでも起動・PDFロード・サムネイル・merged/single保存は成功した。
ただし物理DnDの検証が未完了なので、削除条件を満たしたとは判定しない。
**最終build.batは `--collect-all=fitz` と `--collect-all=tkinterdnd2` を両方維持する。**
サイズだけを根拠に削除していない。以下の正式起動計測は両方を維持した最終ビルドを使用。

## 起動計測

基準A: `ec63604` の未改変main.py、onefile/windowed、従来のfitz / tkinterdnd2 collect-all付き。
基準exeは計測用外部フォルダへ作成し、version resourceのみ省略（動作コードは同一）。
比較B: 最終候補のbuild.batによるonedir。

Python `perf_counter()` をPopen直前に開始。5ms間隔のEnumWindowsで、対応するexeのプロセスが所有する
可視 `TkTopLevel` ウィンドウを検出するまでを測定した。onefileは展開後の子プロセスもexeフルパスで識別。
MainWindowHandle相当のウィンドウ出現計測であり、全ピクセル描画・PDF準備完了の時間ではない。
各10回をA/B交互に実行、試行間0.5秒。p90はnearest-rank（昇順9番目）。
再起動・OSファイルキャッシュ消去はしていない。通常デスクトップ上の測定で、他アプリも起動している。
各試行後は計測対象プロセスだけを終了した。通常終了のハンドル解放は別テストで確認。

| trial | onefile seconds | onedir seconds |
| --- | ---: | ---: |
| 1 | 2.163035 | 0.916020 |
| 2 | 2.392026 | 0.302286 |
| 3 | 2.007172 | 0.267962 |
| 4 | 1.964627 | 0.301686 |
| 5 | 2.039999 | 0.369651 |
| 6 | 2.053975 | 0.251758 |
| 7 | 2.521928 | 0.427564 |
| 8 | 2.054703 | 0.274571 |
| 9 | 2.169806 | 0.348697 |
| 10 | 1.979095 | 0.240397 |
| median | 2.054339 | 0.301986 |
| p90 | 2.392026 | 0.427564 |
| min | 1.964627 | 0.240397 |
| max | 2.521928 | 0.916020 |

中央値改善率: **85.30%**。中央値1.5秒以内・50%以上短縮は、この環境・測定条件ではPASS。
計測器の初回はWin32プロセスパス取得APIの誤りで検出失敗したため破棄し、修正後に全10回を計測し直した。
最終設定前の予備計測（2.151419秒 / 0.327147秒）は上表の統計に混ぜていない。

## PDF / キャッシュ試験

`tests/test_lifecycle.py` をWindowsの実Tk UIで実行。5テスト、11.272秒、PASS。
PDFは一時生成したA4・ページ番号文字・矩形の文書。実務スキャンPDF・極端に重いベクターPDFの結果ではない。
時間は `_load_pdf_async()` 呼出から可視画像描画済みの `document_ready` まで。
同一プロセス内の測定でPDFライブラリはfixture生成によりロード済み。コールド初回PDF計測ではない。

| pages | ready seconds | initial request count | visible pages | generated at ready | result |
| --- | ---: | ---: | ---: | ---: | --- |
| 3 | 0.0906 | 3 | 3 | 3 | PASS |
| 30 | 0.1536 | 16 | 8 | 16 | PASS |
| 100 | 0.0819 | 16 | 8 | 16 | PASS |
| 300 | 0.0982 | 16 | 8 | 16 | PASS |

初期request数はworker.requestへ渡した配列の実測値。ワーカーが同時に消費するpending queueの瞬間値とは区別する。
可視ページが要求配列の先頭にあることをassert。300ページ全部を予約していない。
軽いfixtureでは近接ページもready前に生成終了したが、ready条件には使っていない。

各PDFで上から下へ40段階スクロールし、上へ戻る試験を実行。
キャッシュ最大値は3 / 30 / 96 / 96。累計レンダー数は3 / 30 / 104 / 316。
100・300ページでは先頭画像のevictをassertし、戻った時に再生成・表示されることを確認。
選択ページ・range集合の不変、Shift範囲選択と解除、選択数・状態・ボタン更新もPASS。
高速に上下を切り替えても最後の可視範囲がreadyに収束し、無限再生成は観測しなかった。
画像数とqueue上限は確認したが、長時間のプロセスRSS推移・全種類PDFでのメモリ安定性は未計測。

## リフレッシュ / ハンドル

生成途中を確実に作るため、ソース試験ではget_pixmapに80msのテスト用遅延を挿入した。製品コードに遅延はない。

| case | confirmation | result |
| --- | --- | --- |
| A: 300ページ生成途中refresh | UIリセット呼出100ms未満、cancel設定、キャッシュ空、close後レンダー数増加なし | PASS |
| B: refresh直後に3ページ追加 | 旧pendingを破棄、新3画像のみ表示、readyまで0.3823秒（テスト遅延込み） | PASS |
| 旧イベント | 旧generationのpdf_loadedを注入しても無視。旧画像復活なし | PASS |
| C: 元PDF rename / move | A/B後とも一時PDFをrename→別フォルダmove→復元できる | PASS |
| 手動保存先 | refresh後も保持、rangeエラー・Shiftアンカー・選択・完了状態は破棄 | PASS |
| 例外 | レンダー失敗の詳細をUI表示、worker.close確認、refresh後の再ロード成功 | PASS |
| 読込失敗 | 存在しないPDFでエラー表示とclose確認、refreshで未選択へ戻る | PASS |
| アプリ終了 | stop→worker thread終了→closedをassert、保存中の終了・置換・refreshは抑止 | PASS |

最終onedirの実操作でも300ページロード→下方へドラッグ/ホイール→F5→未選択→3ページ再ロードを確認。
F5後、exeを起動したまま元300ページPDFをWindows上でrename / moveでき、SHA-256も不変。
packagedでの「レンダー途中の瞬間」に対する自動タイミング制御は行っていない（ソースのA/Bで検証）。
明示的なdelete試験は未実施。出力済みファイルは削除していない。

## 保存 / CLI

ソースUI試験では `[1]`、`[3]`、`[5]`、Shiftの1〜5、`1-3,5,8-10` 相当のページ集合を検証。
各集合でmerged/single両方を保存し、ページ数・順番・PAGE文字列・single各1ページをassert。
ダイアログの後でopen_directoryが呼ばれることはソース試験でmock検証。
packaged試作の実UIでは `1-3,5,8-10` のmerged7ページ、single7ファイルを保存し、pypdfで内容と順番を再確認。
完了ダイアログ表示、mergedのOK後に保存先Explorerが開くことも実機確認した。

最終onedir CLIで以下9ケースPASS（成功0 / 失敗1）。

- inputs + pages + output PDF + silent、7ページの内容・順序。
- 複数inputsの先頭PDFのみ使用、outputフォルダ。
- output省略 + silent、元PDFフォルダへの既定名出力。
- 不正範囲 `1--3`。
- 総ページ数超過。
- inputsなし。
- 存在しないPDF。
- pages引数なし。
- output引数なし。

失敗6ケースすべてstderrを確認。この環境のリダイレクト出力は既存通りCP932。
検証スクリプト側でUTF-8と決めつけた表示が一度失敗したため、raw bytes確認に直して全ケースを再実行した。
CLIの既存エラー文言・exit code・silent解釈は変更していない。

### SHA-256

ソース試験4文書、packaged試験4文書とも前後一致。packaged用入力の実測値:

```text
3   e2edd4897338f9cabf73ad2f9dda303b96480cc78d680477d7e22d87b16681a9
30  e77174b5ae3c4ff80c671d9a8eb7fc3ae9d575171332276ae13dcbb1362a1778
100 1b4dfbad2544cfe60cf182bf0d29e89f5798e06da521abd751cfc3455a651462
300 d8b9af315a1ac345d6ee4407c03683fdb3bbd651b843cde1a845b911af0d49c0
```

## Packaged / DPI

- 最終onedir: 起動10回、共通アイコン表示、300/3ページのロード・画像、スクロールバーdrag・wheel、F5、Ctrl+O、再PDF追加、通常終了PASS。
- 共通 `02_assets/dake_icon.ico` と同梱 `_internal/dake_icon.ico` のbytes一致PASS。
- onedir試作: サムネイルクリック・選択解除・range・merged/single・完了ダイアログ・保存フォルダopen PASS。
- 最終onedir: CLIによる保存・エラー9ケースPASS。最終collect-all復元後のGUI保存一式の再実行は未確認。
- ソースのDrop受取ハンドラーはPASS。packagedのExplorer→アプリ間の物理DnDは未確認。実機操作ツールがwindow bounds外へのdragを拒否するため、成功扱いにしない。
- 実機100%: WindowDPI=96、GetScaleFactorForMonitor=100、初期・ロード・選択・スクロール画面を確認PASS。
- 125% / 150% / 200%: 未確認。OS設定は変更していない。Tk scalingだけの擬似検証で代用していない。

## 静的確認

- UTF-8でAST parse成功。定数群外の日本語UI文字列0、UI_TEXT重複キー0、未定義参照0、不要import0。
- CLI日本語をUI_TEXTへ移動。未使用の重複UIキー・resource_path・全ページ生成カウンタ・batch enqueueを削除。
- workerからTk呼出なし。Tk PhotoImage / widget / afterの操作はmain threadだけ。
- generationとviewport requestを分離。旧viewportのレンダー例外でDocumentが閉じた場合は、同じsession全体のエラーとして通知して読み込み待ちに残さない。
- pendingは最新可視＋近接要求だけ、eventsは最大32、cacheは最大96。全ページ予約用afterなし。
- `git diff --check` PASS。対象外アプリ・共通ルールファイルは変更していない。
- build / dist / spec / version_info / pycache / config / exeは既存gitignoreで除外。fixture PDF・一時計測生成物はrepo外に配置しcommit対象外。

## ライセンス確認メモ

現行レンダラーはPyMuPDF 1.24.10 / PyMuPDFb 1.24.10。インストール済みmetadataは `GNU AFFERO GPL 3.0`。
ArtifexはPyMuPDFをAGPLまたは商用ライセンスで提供していることを明示している。
参照: [ArtifexのPyMuPDF告知](https://artifex.com/blog/pymupdf-acquired-by-artifex)、
[ライセンス案内](https://artifex.com/licensing)、[AGPL本文](https://artifex.com/licensing/gnu-agpl-v3)。

DAKE正本にはソースを配布zipへ含めない方針、無断再配布禁止の表記がある。
契約済み商用ライセンスの有無・対応ソース提供方法・配布条件との整合は今回確認できていない。
**正式出荷前の別Phaseでライセンス確認が必要。法的に問題なしとは判定しない。**
今回レンダラー置換は行っていない。

## 残る確認 / 出荷条件

- packagedの物理DnDと、最終collect-all復元版でのGUI保存一式の再実行。
- 125 / 150 / 200%実DPI、実務スキャン・重いベクターPDF、長時間RSS・厳密なコールド起動。
- クリーンvenvでの固定依存ビルド。単体exe前提の配布・ランチャー・zip工程とonedirの整合。
- packaged生成途中の厳密なA/Bタイミング試験、delete可能性の明示試験。
- ライセンス・現在の配布条件の追加確認。
- PRの実機体感確認が完了するまでmerge・正式出荷しない。

## Phase 1.1: PR #29 出荷前ハードニング

検証日: 2026-09-12 (JST)。上記はPhase 1の履歴として保持する。
clean venv、Pillow、最終GUI保存、delete、配布ラインに関する最新の判定は本節を優先する。
起動10回の計測値はPhase 1の成果物の値であり、今回のPillow除去成果物を再計測した値ではない。

### 対象とclean build

- 同じbranch `codex/pdf-split-select-fast-refresh`、同じPR #29を使用。新branch / PRなし。
- 開始時: `git status --short` 空、PR OPEN、head `4e74e219ad3159e8315f3aa565800778ad5f1f14`。
- branchを `git pull --ff-only` で確認（already up to date）。ORIGINALと本レポートを再読。
- `origin/main` は調査用にfetchのみ。`c9e0a37f29436fbfbfc1d06e8fbbf685e07b25d8`。
  mainとの比較で、調査したtools / Launcherの実装差分はなし（生成JSONのみ差分あり）。mainのmergeはしていない。
- 新規 `splitselect-phase11-venv` をrepo外に作成。Python 3.12.4、`include-system-site-packages = false`。
- 更新したrequirementsだけをpip installし、そのvenvを `PYTHON_EXE` に指定して現行 `build.bat` を実行。PASS。
- `build.bat` / `main.py` / `pdf_backend.py` はPhase 1.1では変更していない。
  onedir、両collect-all、共通アイコン指定をそのまま使用した。
- venvのsys.pathにグローバルsite-packagesなし。`find_spec('PIL') is None`、`pip check` PASS。
- 最終候補: `dist/DakePDF_Split_Select/DakePDF_Split_Select.exe` と同じフォルダの `_internal/`。
- 最終exe SHA-256: `30724f0297427f77bad3beac82126459b59200f763f0bc876ffd3e7f69172c7b`。

固定された直接依存: PyMuPDF 1.24.10 / pypdf 6.10.2 / pyinstaller 6.19.0 /
pyinstaller-hooks-contrib 2026.4 / tkinterdnd2 0.4.3。
今回解決された間接依存: PyMuPDFb 1.24.10 / altgraph 0.17.5 / packaging 26.3 /
pefile 2024.8.26 / pywin32-ctypes 0.2.3 / setuptools 84.0.0。pipはvenv初期の24.0のまま。
間接依存まですべてをlockした構成ではない。今回は固定した製品依存の更新を行っていない。

### Pillow判定: REMOVE

- 製品 `main.py` / `pdf_backend.py` のPillow / PIL importは0件（テキスト検索・AST確認）。
- PyMuPDF 1.24.10の必須依存はPyMuPDFb 1.24.10。Pillow importは未使用の `pil_save()` 内。
- tkinterdnd2 0.4.3にPillow必須依存なし。pypdfではimage/full extraのみで、今回そのextraを使用しない。
- PyInstaller 6.19.0は今回の正しいICOファイルをそのまま扱う。Pillowはアイコン形式変換のオプション経路。
- 除去前のAnalysis-00.tocにはPIL.Image / PIL.ImageTk等が取り込まれていた。
  clean build後のAnalysis-00.tocでは `PIL|Pillow` 0件。
- Pillow未インストールでビルド、GUIロード・サムネイル・両保存、CLI、回帰テストが成功した。
  本アプリの検証対象経路では不要と判定し、requirementsの1行のみ削除。
- 共通 `tools/make_booth_ready.py` は販促画像生成にPillowを使用する。
  これは別の出荷ツール環境の依存であり、本アプリの実行・ビルド依存へ戻す根拠ではない。

### 最終候補GUI / 保存

新規fixtureはrepo外の `splitselect-phase11-evidence` に生成し、旧試作版の出力を流用していない。
以下はすべて今回のclean buildのexeで再実行した結果。

| 項目 | 結果と確認範囲 |
| --- | --- |
| exe起動 / import | PASS。例外ダイアログなし、初期UIを表示 |
| icon | PASS。タイトルバー表示、共通ICOと同梱ICOのbytes一致 |
| PDF追加 / thumbnail | PASS。30ページをPDF追加からロードし、先頭8画像を表示 |
| click | PASS。P.1 ONで1ページ・選択中・ボタン有効、再クリックでOFF |
| Shift | ソース回帰PASS。最終packagedでの物理Shift+クリックは未確認 |
| range | PASS。`1-3,5,8-10` 単独で選択数7、保存結果も7ページ |
| merged | PASS。クリック選択の1ページと、範囲単独の7ページをそれぞれ保存 |
| single | PASS。1,2,3,5,8,9,10の7ファイル、各1ページ |
| 完了dialog / OK | PASS。mergedの1/7ページ、singleの7ページで完了表示とOK操作 |
| Explorer | PASS。OK後、試験保存先のExplorerウィンドウが開いたことを確認 |
| F5 | PASS。ファイル未選択、0ページ、range空、選択0、画像空、未選択案内、完了表示消去 |
| Ctrl+O / 再PDF | PASS。F5後のファイルダイアログ、300ページと別3ページのロード |
| 通常終了 | PASS。閉じるボタン後に対象ウィンドウ消滅、強制終了はしていない |

サムネイル選択が存在すればrangeより優先する既存仕様は変更していない。
クリックで選んだP.1とrangeが併存した最初の保存は1ページ。
P.1を再クリックで解除してrange単独とし、7ページ保存を再試験した。

merged / singleはpypdfでページ数・順番・`PHASE11 PAGE NNN`内容をassert。
さらにPyMuPDFで各出力ページと対応元ページを同じ既定倍率でレンダーし、全pixels一致をassert。
全入力SHA-256は処理前後一致（PASS）:

```text
3   d786bf814eb876a390677e244bf076a290232bf4467aad04dcc6b64252a00c98
30  388f01394beb24f153c729ba702b8a25d7254a67914b62517128c117e2c7d0b7
100 751cb7f1577450992f1ec65ec27429aefde3678b3881bc2a0dcadbbbb8f0ae93
300 296d3b4cc572dcbac7105420fe3476fb7f7e27e545100ebacb2cbc5b294e451b
```

### Refresh / delete / CLI / 静的回帰

- 最終packaged: 300ページ専用コピーをロード → F5 → Ctrl+O → 3ページロード、PASS。
  3ページの画像・ページ数・準備完了を確認。その後も旧画像・旧完了表示の復活なし。
- 最初の画面取得時点で300ページの可視画像は準備完了だった。
  **packaged生成途中の厳密タイミングは未確認**。通常F5のPASSで代用しない。
- ソース制御試験では生成途中refresh / 新3ページ / 旧task停止 / generation無視 / closeがPASS。
  今回の新3ページreadyは0.3855秒（テスト用80ms遅延込み）。製品コードへsleep追加なし。
- refresh後、最終exeを起動したまま専用コピーをWindows上で
  `disposable-handle-300.pdf` → rename → `moved/` へmove → delete。すべて成功（PASS）。
  元fixture・出力PDFは削除していない。Windowsファイル操作を妨げる旧ハンドルがないことを実証。
  packagedの内部taskカウンタは観測していないため、内部停止そのものはソース試験の結果と区別する。
- clean venvの `tests/test_lifecycle.py`: **5 tests PASS / 10.707秒**。
  3/30/100/300ページ初期要求3/16/16/16、cache最大3/30/96/96、evict後復帰と選択保持PASS。
- 最終exeの `--from-shimarisu` CLIを9ケース再実行: 成功3件exit 0、エラー6件exit 1とstderr、PASS。
  inputs先頭採用、pages、output PDF/フォルダ/省略、silent、不正範囲、超過、欠落引数を確認。
- `python -m compileall -q main.py pdf_backend.py tests`: PASS。
- UTF-8 AST、定数群外の日本語UI文字列0、UI_TEXT重複0・未定義参照0、PIL import0: PASS。
- `git diff --check`: PASS。製品コード・非同期構造・UIレイアウトにPhase 1.1の変更なし。

### Onedir配布ライン調査: B（出荷ツールに対応が必要）

正式出荷ツールは実行していない。以下はコード・ローカル商品データ・GitHub Release APIを読み取った結果。
Web/Storeの本番反映や購入・配信は試していない。外部の未発見スクリプトまで互換とは断定しない。

| 対象 | 判定 | 根拠 / 次工程 |
| --- | --- | --- |
| onedir ZIP | compatible | repo外のみでフォルダ全体をZIP化・展開。1084ファイル全SHA一致、ZIP CRC正常、展開先exeのCLI保存exit 0 / 3ページ。30,227,806 bytes。正式booth_readyは未生成 |
| BOOTH生成 | change required | `tools/make_booth_ready.py:447` create_zipはexe/README/注意事項のみ同梱、455 find_exeはdist直下のみ。現行では入れ子exeを見つけず、exeを渡しても_internalを落とす |
| GitHub Release | change required | 現行v1.0.0の実assetは `DakePDF_Split_Select.exe` 単体（47,779,048 bytes）。次回は全onedir ZIPへ変更が必要。repo内py/ps1/bat/yml/yaml検索でRelease upload実装は未発見、従来アップロード工程自体はnot checked |
| Launcher（現行distレイアウト） | change required | `AppMeta.standard_exe_path` は `app_folder/dist/exe_name`。新しい `dist/DakePDF_Split_Select/exe` は自動検出されない。下記の配置調整または既存手動指定で本体変更を回避可能 |
| dakeapp商品リンク | compatible | ローカル `dakeapp-site/public/apps/pdf-split-select/index.html:40` はRelease tagページへリンク。exe直リンクではない。`tools/sync_dakeapp_apps_json.py:72` 以降もrelease_urlをそのままデータ化 |
| Store商品表示 | compatible | ローカルstore.js:110 purchaseActionはStripe/BOOTH URLを使用、exeパス非依存。該当商品JSONはdownload_url/type null、GitHubはtag URL。決済後の実ファイル配信はnot checked |
| formal shipping scripts | change required | `check_exe_launch.py:79`、`release_capture.py:153`、`check_booth_ready.py:401` もdist直下exe前提。検出・capture・検査を新配置対応にする必要あり |
| Pack ZIP | compatible | `make_pack_ready.py:453` は個別ZIPをそのまま `apps/<folder>/` に入れる。修正済み個別ZIPを供給すれば内容の再解釈なし。正式Pack生成は未実行 |

Launcher詳細（`01_apps/DAKE_Launcher/main.py`）:

- 98行目: 標準exeはアプリフォルダ直下ではなく `dist/exe_name`。
- 663行目 `resolve_exe_path`: `custom_exe_paths` が最優先。見つからなければ標準パスだけを確認。
- 675行目 `choose_custom_exe`: 既存のファイル選択・パス保存機能あり。
- 708行目 `launch_app`: 選択したexeを、その親フォルダをcwdとしてPopen。
  Release URLはブラウザを開くだけで、Releaseから直接exeを起動・展開する仕組みではない。
- 方法1: 既存の手動exe指定で新しい入れ子のexeを選ぶ。`_internal`を隣に保持する。
- 方法2: 出荷/開発配置の別工程でonedirの**中身全体**を `<app_folder>/dist/` に置き、
  `dist/exe` と `dist/_internal/` を並べる。これなら標準パスのLauncher変更は不要。
  ただしBOOTH ZIPのexeだけ同梱する処理は、どちらの方法でも修正が必要。
- Launcher実UIでの起動は未実行。上記は実コードのパス解決・Popen契約の確認結果。
  今回Launcher、共有出荷スクリプト、DAKE_METAのexe_nameは変更していない。

正式出荷の別Phaseで必要な最小対応:

1. onefile / onedir両方を検出し、onedirの場合はexeと_internalの同階層関係を維持してZIPへ再帰同梱する。
2. 検出・起動検査・captureを同じ規則へ合わせ、配布READMEに「ZIPを展開しフォルダ全体を保持」を明記する。
3. Launcherは手動指定か上記配置調整を採用し、実際のLauncher起動で検証する。
4. Releaseはexe単体ではなく全onedir ZIPを出荷単位にする。BOOTH/Store配信元も同じ検証済みZIPへ揃える。
5. 両collect-allにより_internalに第三者ライブラリのpyファイル5件が含まれる（fitz 3件、tkinterdnd2 2件）。
   本アプリのmain.pyを同梱しているという意味ではない。依存物を無差別削除せず、配布方針・ライセンス/noticeを整理する。

### Human Review / formal shipping blocker

- physical Explorer → app DnD: 未確認。ソースのDropイベント試験PASSとは区別。
- 最終packagedのShift+クリック選択/解除: 未確認。操作APIはclickの修飾キー/keydown保持を提供せず、
  一瞬のShiftキー送信を代用しない。ソースの範囲選択/解除回帰はPASS。
- packaged生成途中のF5厳密タイミング: 未確認。ソース制御試験PASS、packaged通常F5 PASS。
- 実DPI 125 / 150 / 200%: 未確認。OS設定は変更せず、Phase 1の100%確認を維持。
- Launcher実UI、Web/Storeの本番配信、外部Releaseアップロード工程: 未確認。
- 実務スキャン/重いベクター、長時間RSS、厳密コールド起動: Phase 1から継続未確認。
- PyMuPDF 1.24.10 / PyMuPDFb 1.24.10、AGPL-3.0 / Artifex商用ライセンスの確認メモは維持。
  契約状況・対応ソース提供・DAKEの再配布条件の整合は未解決。**merge blocker: NO / formal shipping blocker: YES**。
  renderer replacement: 今回しない。法的に問題なしとは判定しない。
- 配布ラインBの修正とHuman Review受入も正式出荷前に必要。現時点で出荷可能とは扱わない。

### Phase 1.1 Git確認

- 変更対象は `requirements.txt` と本レポートのみ。対象外変更なし。
- build/dist/spec/exe/config/venv/fixture/ローカルZIP/検証scriptはstage・commitしない。
- 同じbranchへcommit/push、PR #29の本文更新のみ。main merge、Release、BOOTH、Store、Cloudflare更新なし。
- 提出時の最終 `git status --short` が空であること、およびPR OPENとhead一致をcommit/push後に確認し、PR本文・最終報告へ記録する。
