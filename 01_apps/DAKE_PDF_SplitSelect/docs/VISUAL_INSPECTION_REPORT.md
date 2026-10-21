# SplitSelect Phase 2: Visual Page Inspection

確認日: 2026-10-22 / Windows 11 build 26200 / Python 3.12.4 / system DPI 96。

## 判定と作業範囲

実装・ソース上のTk/PDFium回帰はPASS。最終onefileの起動・CLI・payload監査は実施。
**最終EXEのGUI操作一式は未確認であり、出荷可能とは判定しない。** Windows操作接続のアプリ承認がタイムアウトしたため、物理操作をPASSへ読み替えていない。レビュー用Draft PRで停止する。

- ブランチ: `codex/splitselect-visual-inspection`。
- 基点: 作業開始時の `origin/main` = `13427e0c73ad706ad4b2063c12835f3eeedc877f`。ユーザー提示の過去SHAではなくfetch後のmainを使用。
- 正本 `C:\Users\yukiz\devlop\DAKE_series` は既存のtracked変更6件と他案件のuntrackedがあるため、切替・stash・stageを行っていない。
- 専用worktree: `C:\Users\yukiz\devlop\_worktrees\splitselect-visual-inspection`。
- ORIGINALを最初に更新。過去の出荷レポート、正式配布物、STADIO安全退避データは変更していない。
- 正式v1.0.2 EXEのSHA-256は作業後も `5228b583f490b41c9f304f331d59fc5afa2588852ebb930b04644132b617b6e0`。
- バージョンは更新していない。候補のFileVersionも既存の1.0.2であり、正式Releaseの差し替えではない。

## 実装と操作

「一覧で探す。大きく見て判断する。一覧で選ぶ。」を維持する。

| 操作 | 挙動 |
| --- | --- |
| 小 / 標準 / 大 / 特大 | 108×152 / 152×214 / 228×321 / 304×428のPDFium画像。カードは各画像に横24・縦50を加算 |
| 左クリック | 従来の選択・解除 |
| Shift＋左クリック | 従来のアンカー範囲選択・解除 |
| Ctrl＋左クリック / 右クリック | スクロール座標で特定した対象ページのプレビュー。選択は変更しない |
| 背景右クリック | 何もしない |
| 一覧ホイール / scrollbar | 一覧スクロール。既存ttk scrollbarを維持 |
| プレビューホイール | 初期fit=100%、50〜300%、カーソル位置を基準に拡縮 |
| プレビュー左ドラッグ | 画像範囲内でpan |
| Esc / × | プレビューだけ閉じる |
| F5 / PDF切替 / アプリ終了 | プレビューを閉じ、旧要求を無効化してworker上でPDFをclose |
| Ctrl+O | 従来のPDF追加。プレビューからも使用可能 |

プレビューは非モーダルToplevelを1つだけ再利用する。選択ボタン・チェック・別の選択操作は設けない。
初期900×700、画面内かつ最大1400×1000。最大寸法でも300%の実画像を描ける12MP予算とし、巨大ウインドウで倍率表示だけが300%になる状態を避ける。

OverviewRenameからは位置保持、fit基準、カーソル基準ズーム、最新要求の採用という設計を参考にした。
同アプリの独立worker・Pillow画像処理・複数文書モデルはコピーせず、Selectの既存所有関係へ適合させた。

## Workerとキャッシュ

- PDFium文書の所有者は既存の `PdfRenderWorker` 1本のみ。thumbnail / previewは同じworkerで直列処理する。
- document generation、thumbnail request IDと要求寸法、preview request IDと対象ページを区別する。旧世代・旧要求の結果はUIに採用しない。
- previewはopen/close barrierの後、待機thumbnailの前へ投入する。実行中の1ページは中断せず、その後は最新要求を採用する。
- ズーム・resizeは100ms debounce。preview結果queueは最新1件のみ、通常イベントqueueは既存の最大32件。Tk更新はmain threadの20ms pollから行う。
- `page_preview.py` は最初に開くまでimportしない。PDFiumとpypdfも従来のlazy importを維持。新規依存なし、requirements変更なし。
- 一覧は可視ページ＋前後8ページのみ要求する。サイズ変更時には不要な待機thumbnailをキャンセルし、同じサイズの再クリックではキャンセルしない。
- 全サイズ共通で最大96画像、推定画素メモリ64MiB。現サイズのLRUと退避した解像度キャッシュを同じ予算で管理し、再利用可能な解像度は再生成しない。
- evictionは画像だけに適用する。選択・Shiftアンカー・範囲入力・保存先は別の既存状態として保持する。
- サイズ変更の連続操作では同じ閲覧ページを保持し、列数の変化でアンカーが累積して先頭へずれないようにする。
- preview画像は一覧キャッシュに入れず、表示1枚・最新結果1枚・実行中1要求に限定する。通常ウインドウのfit293,664画素、300%2,639,112画素を確認。

## 自動回帰

実行: `.venv\Scripts\python.exe -m unittest discover -s tests -v`

**25 tests / 49.255秒 / PASS / skipなし。** source Tk integrationとpackaged CLI/payload検査を含む。

| 項目 | 結果と範囲 |
| --- | --- |
| 4サイズ・連続切替・同サイズ再クリック | PASS（実Tk/PDFium） |
| 閲覧位置・選択・範囲入力・Shiftアンカー | PASS（100/300ページ、切替前後の内部値と表示状態） |
| 末尾到達・evict後の再表示 | PASS（100/300ページ） |
| 右クリック / Ctrl＋左クリック / 背景 | PASS（Tkイベント注入。物理入力ではない） |
| 左クリック選択・解除 / Shift範囲 | PASS（Tkイベント注入＋既存selection回帰） |
| 閲覧専用・ウインドウ再利用 | PASS（選択集合・範囲・アンカー・保存先・PDFパス不変） |
| 50〜300%・pan・resize・Esc・WM_DELETE_WINDOW | PASS（実Tk/PDFium。最大ウインドウで実画像寸法3倍とカーソル基準点も検証） |
| preview優先順位・worker所有スレッド | PASS（待機thumbnail24件よりpreviewが先、全renderのthread ID一致） |
| preview中 / 描画途中F5・即別PDF | PASS（制御Eventを使ったソース側試験。製品にsleep追加なし） |
| PDF切替・旧イベント・終了・PDF close | PASS（generation不一致結果を無視、worker終了、保存中保護維持） |
| preview失敗後の一覧利用 | PASS（previewエラーのみ表示、一覧PDFを不用意に閉じない） |
| rename / move / delete | PASS（実際にworkerで開きpreviewしたテストコピーをF5後に操作） |
| merged / single | PASS（source GUI経路。ダイアログとExplorer呼出しはmockで確認） |
| 保存内容 | PASS（1 / 3 / 5 / Shift範囲 / 1-3,5,8-10、ページ数・順序・内容、single各1ページ） |
| CLI / 元PDF SHA-256 | PASS（sourceと最終EXE、成功0・失敗1、既存ファイル保護・入力不変） |
| 日本語widget直書き | PASS（ASTでtext/title/messageへの日本語リテラルなし） |
| compileall / git diff --check | PASS |

多重Tkを作るテストのafter_cancel所有元を修正し、破棄したTkオブジェクトはテスト/UI threadで回収する。これはテスト終了手順の修正であり、workerからTkを操作する実装を追加したものではない。

## 表示・メモリ実測

生成した番号付きベクターPDFを使用。実契約書・スキャン大判図面の性能を代表するものではない。

| ページ数 | 標準サイズ初期可視ready（秒） | 初期要求数 | 可視数 |
| --- | ---: | ---: | ---: |
| 3 | 0.1727 | 3 | 3 |
| 30 | 0.0674 | 16 | 8 |
| 100 | 0.0680 | 16 | 8 |
| 300 | 0.0779 | 16 | 8 |

300ページを300件予約せず、ready時にも全ページ生成を待っていない。優先順位テストで直接投入した24件は初期queue数と混同しない。

| ページ数 | preview初回（秒） | 300%反映（秒、debounce込み） | Working Set bytes | Private bytes |
| --- | ---: | ---: | ---: | ---: |
| 3 | 0.127887 | 0.137789 | 127131648 | 101756928 |
| 30 | 0.152103 | 0.132950 | 128835584 | 102920192 |
| 100 | 0.137425 | 0.136047 | 128589824 | 102801408 |
| 300 | 0.150640 | 0.140920 | 130908160 | 104955904 |

このテストプロセスのPeak Working Set採取値は160,546,816 bytes（153.109MiB、最大ウインドウ試験を含む）。
サイズ切替の採取値は100ページで0.0261〜0.1196秒、300ページで0.0436〜0.1169秒。近接要求まで完了した時点の値であり、初回1枚表示だけの時間ではない。
キャッシュは全サイズ合計96画像以内・推定64MiB以内を各スクロールでassert。サイズ切替ログ採取点の最大は23,168,584 bytes。これはプロセス全体の上限ではない。

## Clean onefile候補

新規 `.venv` に固定requirementsだけをinstallし、対象 `build.bat` そのものでclean buildした。
Python 3.12.4、pypdfium2 5.13.0、pypdf 6.10.2、PyInstaller 6.19.0、hooks-contrib 2026.4、tkinterdnd2 0.4.3。
解決された間接依存はaltgraph 0.17.5、packaging 26.3、pefile 2024.8.26、pywin32-ctypes 0.2.3、setuptools 84.2.1。Pillow/PyMuPDFはinstallしていない。

- 形式: `--onefile --noconsole` を維持。EXE名・正式版番号は変更なし。
- 最終候補: 17,684,734 bytes / SHA-256 `e04a1a257625d24a524d6f86353ae20d70159b5a4b29de98844382340ffe39e9`。
- 共通アイコンはPE指定と埋め込みicoの一致を確認。表示そのものの目視は未確認。
- Windows metadata: CompanyName=`PEAKHEADZ`、LegalCopyright=`Copyright Yukihiko Kikuta`、FileDescription=`DakePDF分割Select`、FileVersion=`1.0.2`。
- 共通generatorには呼出し単位の任意overrideを追加し、SplitSelectのbuildだけが使用する。他アプリの既定出力が変わらないことをテスト済み。
- PDFium DLL 1件、tkdnd DLL、python312.dll、page_preview/pdf_backend/pypdfium2モジュールをarchive内で確認。
- PyMuPDF / fitz / Pillow payload 0。wheel由来19ライセンス文書を配布用dist・EXE内部の双方でSHA照合。noticesと共通icoも照合。
- 法的適合の新しい断定はしない。依存追加はなく、既存PDFium配布条件を維持する。

## 起動10回

測定定義: 過去v1.0.2と同じstdlib benchmark。Popenから同じEXEパスに属するvisible `TkTopLevel` の検出まで（5ms poll）。各回WM_CLOSEで通常終了しexit 0、0.5秒間隔。OSキャッシュの強制消去や他PCでの測定ではない。

| 試行 | 空フォルダのEXE単体（秒） | DesktopのEXE単体（秒） |
| --- | ---: | ---: |
| 1 | 1.128622 | 1.121849 |
| 2 | 1.250720 | 1.308062 |
| 3 | 1.308578 | 1.295594 |
| 4 | 1.311537 | 1.324016 |
| 5 | 1.373749 | 1.225689 |
| 6 | 1.287372 | 1.096913 |
| 7 | 1.243710 | 1.309819 |
| 8 | 1.071479 | 1.370683 |
| 9 | 1.315558 | 1.307966 |
| 10 | 1.253949 | 1.276049 |
| median | **1.270660** | **1.301780** |
| p90（nearest rank） | **1.315558** | **1.324016** |
| 最速 | 1.071479 | 1.096913 |
| 最遅 | 1.373749 | 1.370683 |

両フォルダともEXE以外の隣接ファイルなし、20回すべてvisible windowを検出し通常終了exit 0。
過去正式v1.0.2のmedian1.437372秒 / p90 1.553547秒に対し、今回の空フォルダmedianは参考比較で約11.6%短い。
過去onedirのmedian約0.253秒よりは遅いが、今回onefileを維持し、明確な起動退行は観測していない。
同時A/B・cold boot統制ではないので、性能向上の厳密な因果は主張しない。試作候補の数値を最終候補の値として流用していない。

## Human Review / 未確認

- **最終EXEのGUI一式**: PDF追加、thumbnail、4サイズ、プレビュー、高解像度PDFium runtime、クリック/Shift/Ctrl、range、merged/single保存、完了ダイアログ、OK後Explorer、F5、再追加、Ctrl+O、通常終了。source integrationのPASSをpackaged GUIのPASSにしていない。
- 最終EXEでサムネイル／preview描画途中を狙ったF5と、実際の物理DnD（Explorer→app）。sourceイベント注入とは別ゲート。
- 契約書・重要事項説明書・細線図面・スキャンPDFの目視判読、フォント・余白・フッター非欠け、実マウスのzoom/pan感触。
- DPI 125/150/200%、複数モニター、非開発PC。system DPI 96の読取りだけで画面全体を目視PASSにはしていない。OS設定は変更していない。
- 最大1400×1000というメモリ保護上限の実用性もユーザーレビュー対象。
- 正式配布ZIPの新規作成・解凍後確認、正式公開は今回未実施。過去v1.0.2のZIPを再検証済みとは扱わない。

## 証跡・変更ファイル・Git

ローカル証跡: `C:\Users\yukiz\devlop\DAKE_release_evidence\splitselect-phase2-20261022`。
`final-tests.txt`、`build.txt`、`startup*.json/txt`、`desktop-cli-tests.txt`、生成fixture、EXE単体コピーを保持する。prototype/pre-window-cap計測は途中候補であり最終値とは分離。
Desktopの `DakePDF_SplitSelect_Phase2_Test_20261022` はEXE単体テスト／Human Review用で、正式配布物ではない。

変更対象は次の13ファイルのみ。requirements・19ライセンス文書・他アプリは変更なし。

- `ORIGINAL.md`, `README.md`, `build.bat`, `main.py`, `pdf_backend.py`, `page_preview.py`
- `tests/test_lifecycle.py`, `tests/test_license_payload.py`, `tests/test_packaged_cli.py`
- `tests/test_visual_inspection.py`, `tests/test_visual_contract.py`, `docs/VISUAL_INSPECTION_REPORT.md`
- `../../tools/generate_version_info.py`（既定互換の任意metadata overrideのみ）

対象ファイルだけをstageする。build/dist/spec/venv/cache/テストPDF/EXE/evidenceはcommitしない。
実装commit `ba26afb53932f66f74dcdea4fdab581659325a3e` を指定ブランチへpushし、専用worktreeの `git status --porcelain=v1` が空であることを確認した。
PR [#56](https://github.com/yukiPHZ/dake-series/pull/56) はDraft / OPEN、base=`main`、mergedAt=nullを確認。最終EXEのGUI未確認ゲートを本文にも明記している。
この結果を追記する文書のみのcommit後もclean・PR head一致を再確認する。正本側の既存dirty状態は維持。
main merge、Release、BOOTH、Stripe、Store、dakeapp.com、Cloudflare、本番配布物の更新は行わない。
