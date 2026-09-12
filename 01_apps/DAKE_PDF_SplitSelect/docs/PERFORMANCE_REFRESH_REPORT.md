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
