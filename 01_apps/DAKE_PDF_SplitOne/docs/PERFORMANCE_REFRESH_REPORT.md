# SplitOne Performance / Refresh Report

検証日: 2026-09-20 (JST)。対象: `01_apps/DAKE_PDF_SplitOne`。
この報告のPASSは記載した条件・経路に限る。未確認項目をPASSとは扱わない。

## Scope / baseline

- branch: `codex/pdf-split-one-fast-refresh`
- base: `ddf5d14` (`origin/main`、作業中に進んだmainへrebase済み)
- 正式仕様を実装より先に `ORIGINAL.md` へ反映した。
- SplitOneだけを変更し、Release / BOOTH / Store / Cloudflare / `booth_ready` / 共通出荷ツールは変更していない。
- SplitSelect PR #29のlazy import、generation cancellation、onedirの知見を参照したが、thumbnail、fitz、LRU cacheなどの構造は持ち込んでいない。

## Before

- GUI起動前に `from pypdf import PdfReader, PdfWriter` を実行していた。
- buildはonefileで、毎回展開を伴った。
- `busy` の間はリフレッシュが無効だった。
- worker eventにgenerationがなく、中止後の旧progress / done / errorを識別できなかった。
- `元ファイル名_split` へ直接書き、既存 `pNNN.pdf` を削除して再生成していた。
- 中止tokenとjob固有一時フォルダがなかった。

## After

- GUI importと初期画面表示ではpypdfを読み込まない。CLI実行時、または分割worker開始後に遅延importする。
- 各jobにgenerationと `threading.Event` cancel tokenを持たせた。
- loaded / progress / done / errorの全イベントへgenerationを付け、現在generation以外はUIで破棄する。
- リフレッシュ / F5は処理中も有効。cancelを送って即時に未選択へ戻り、次PDFを受け付ける。
- 処理中のボタン文言はUI_TEXTの `中止`。保存先変更と新規投入は処理中だけ無効。
- ページ取得前、書き込み前、書き込み後 / 次ページ前にcancelを確認する。
- 保存先内の `.<stem>_split_work_<random>` へ生成し、全ページ成功後だけ正式フォルダへrenameする。
- 正式名は `sample_split`、`sample_split_2`、`sample_split_3` の順で空きを選び、既存成果物を変更しない。
- cancel / error / close cancelではwork folderを削除する。
- Ctrl+O、F5、WM_CLOSE時cancelと最大3秒の非blocking cleanup待ちを追加した。
- 最終候補は `dist/DakePDF_Split_One/DakePDF_Split_One.exe` と `_internal/` のonedir。

## Clean venv / dependencies

repo外の新規venv、`include-system-site-packages = false` で検証した。

- Windows 11 build 26200
- Python 3.12.4
- pip 24.0（upgradeしていない）
- pypdf 6.10.2
- tkinterdnd2 0.4.3
- PyInstaller 6.19.0
- pyinstaller-hooks-contrib 2026.7（間接依存）
- altgraph 0.17.5 / packaging 26.3 / pefile 2024.8.26 / pywin32-ctypes 0.2.3 / setuptools 84.0.0（間接依存）
- `pip check`: PASS (`No broken requirements found.`)
- clean venvから最終onedir build: PASS
- `build.bat` はpipをupgradeしない。空白を含む `PYTHON_EXE` を引用して実行できることも確認した。

## Startup measurement

計測Aは最新main `ddf5d14` の未改変SplitOne（eager pypdf import / onefile）。
計測Bは本branchの最終SplitOne（lazy pypdf import / onedir）。両方とも同じclean venvと固定依存でbuildした。

`perf_counter()` をprocess start直前に開始し、5ms間隔のEnumWindowsで、そのexe pathのプロセスが所有する可視 `TkTopLevel` の出現までを測定した。onefileは展開後の子プロセスもexe pathで識別した。A/Bを交互に各10回、試行間0.5秒。各試行後は対象exeのプロセスだけを終了した。p90はnearest-rank（昇順9番目）。再起動やOS file cache消去は行っていない。

| trial | onefile before (s) | onedir after (s) |
| --- | ---: | ---: |
| 1 | 1.444664 | 0.238439 |
| 2 | 1.185959 | 0.251235 |
| 3 | 1.170729 | 0.277515 |
| 4 | 1.036039 | 0.223529 |
| 5 | 1.045474 | 0.300464 |
| 6 | 1.078353 | 0.222619 |
| 7 | 1.196634 | 0.302592 |
| 8 | 1.328570 | 0.294855 |
| 9 | 1.207916 | 0.292124 |
| 10 | 1.130985 | 0.299543 |
| median | **1.178344** | **0.284820** |
| p90 | **1.328570** | **0.300464** |
| min | 1.036039 | 0.222619 |
| max | 1.444664 | 0.302592 |

median改善率: **75.83%**。目安のmedian 0.5秒前後以下、現行比50%以上短縮を、この環境と測定条件では満たした。スプラッシュや遅延表示は追加していない。

参考として最終コード同士のonefile / onedir比較も別runで実施し、median 1.107885秒 / 0.247456秒、77.66%短縮だった。正式なbefore / after値には混ぜていない。

## Config startup impact

`AppConfig.last_save_folder` の通常ローカル既存pathに対する `Path.exists()` を1,000回 x 5回測定した。median runは31.346ms、1 callあたり約0.031ms。通常ローカルpathでは今回の起動時間に対する問題を認めず、ネットワーク切断path専用の非同期config機構は追加していない。

## Split / SHA-256

`tests/test_lifecycle.py` の製品service経路と、最終onedirのpackaged CLI経路で3 / 30 / 100 / 300ページを確認した。fixtureはページごとにMediaBox幅を1ずつ変え、追加ライブラリなしで順序を検証した。

| pages | source service | final packaged | output count | each 1 page | order | source SHA-256 unchanged |
| ---: | --- | --- | ---: | --- | --- | --- |
| 3 | PASS | PASS (CLI) | 3 | PASS | PASS | PASS |
| 30 | PASS | PASS (CLI) | 30 | PASS | PASS | PASS |
| 100 | PASS | PASS (CLI) | 100 | PASS | PASS | PASS |
| 300 | PASS | PASS (CLI) | 300 | PASS | PASS | PASS |

最終packaged fixture SHA-256:

```text
3   1a38b68b8d83c1b8640bb31ba17e4b059e95321ee3494651bdb4495ff80d1902
30  17287c39d372981749f0f188780d124eb619bae159d0a8574c2e71334d55b5dc
100 769bf414fd9822e015372055916527dfd69db1ba881a45e8eb4cd3b8e1d2aaa2
300 69b9ee742d97142f159e2e3d946faa4506689c087038725c550f49b0075752b3
```

packagedの表はGUI操作ではなく、同一exeの `--from-shimarisu` 経路でpypdf同梱、ページ出力、順序、SHA不変を確認した結果。GUI serviceのlifecycleはソース回帰で確認した。

## Non-destructive repeat

- 1回目: `sample_split`
- 2回目: `sample_split_2`
- 1回目の全 `pNNN.pdf` SHA-256を保存し、2回目実行後も全件一致: PASS
- 元PDF SHA-256不変: PASS
- `_clear_previous_outputs` / `OUTPUT_PATTERN`: 削除済み、検索0件
- 正式outputへの直接書き込み: なし

## Refresh / cancel / immediate replacement

製品コードにsleepは入れていない。テスト側だけで `PdfWriter.write` をgateし、300ページ処理途中を確実に作った。

| case | confirmation | result |
| --- | --- | --- |
| mid-split cancel | cancel token設定、UI即idle、error扱いなし | PASS (source lifecycle) |
| immediate replacement | 300ページcancel直後に3ページを開始し、新jobが完了 | PASS (source lifecycle) |
| old progress | 旧generationの299/300 progressを注入しても表示なし | PASS |
| old done | 旧generationのdoneを注入してもcompletionなし | PASS |
| old error | 旧generationのerrorを注入してもerror表示なし | PASS |
| old dialog / Explorer | completion callbackが新jobの1回だけ | PASS (mocked UI boundary) |
| output isolation | 旧300ページの正式folderなし、新3ページだけ3出力 | PASS |
| manual save folder | refreshは `save_folder` を変更しない | PASS (controller design / test setup) |

packagedの実キーF5、F5直後の別PDF、旧dialog / Explorer抑止はnative Computer Useがこのホストで無効だったため未確認。ソース検証をpackaged実機PASSへ読み替えていない。

## Temp cleanup

| path | result |
| --- | --- |
| success | PASS、work temp 0 |
| cancel | PASS、work temp 0、正式な半端output 0 |
| error | PASS、work temp 0、正式な半端output 0 |
| close cancel | PASS、controller cancel-all後にworker終了、work temp 0 |

closeのpackaged GUI実操作は未確認。WM_CLOSEは全cancel tokenを設定し、50ms `after` pollingでworker終了を待ち、最大3秒でUIを長時間blockしない実装。

## CLI compatibility

最終onedirで以下を確認した。

- explicit `--inputs` + `--output` + `--silent`: exit 0
- output省略のtimestamp folder: exit 0
- 複数inputsは先頭だけ使用: exit 0
- inputsなし / missing / non-PDF / unknown option: すべてexit 1
- GUI、dialog、ExplorerはCLIで起動しない
- 3 / 30 / 100 / 300ページのpackaged出力数、各1ページ、順序、入力SHA不変: PASS

明示 `--output` 内に既存 `元stem_pNNN.pdf` がある場合、現行どおり同名ファイルを上書きしexit 0となることを実測した。GUIの非破壊化を理由にCLI互換を変更していない。DAKE CLI共通仕様との将来的な整合は別課題として扱う。

## Packaged candidate / icon

- final path: `dist/DakePDF_Split_One/DakePDF_Split_One.exe`
- `_internal/`: 存在
- final exe SHA-256: `1da514767bd3cefca95e5d809800b7c0b0e2efd84d46cadda6fbdb50a259d045`
- build output: `Copying icon to EXE` を確認
- generated specのicon: `../../02_assets/dake_icon.ico`
- common ICO SHA-256: `761adcba5e348027a705855cfa88cf19cd855cb2d3aeb0eff5282e544614a8d6`
- visible main window startup 10回: PASS

common iconの目視、click PDF追加、physical DnD、progress表示、completion dialog、Explorer open、repeat GUI、F5、processing中F5、F5直後別PDF、Ctrl+O、normal closeは、native Windows Computer Use APIが利用不能だったため未確認。Human Reviewへ残す。

## Physical DnD / packaged F5

- sourceのDnD受取ハンドラーは既存のTkinterDnD経路を維持し、処理中だけ入力を無視する。
- PyInstaller analysisにtkinterdnd2とpypdfが含まれること、最終onedir起動、packaged CLI処理は確認した。
- Explorerから最終onedirへのphysical DnD: **未確認**。
- packagedで300ページ処理中F5 → 即idle → 3ページ追加: **未確認**。
- 未確認理由: このCodexホストのComputer Useはbrowserのみで、`computer/getApp/listWindows` native APIが公開されていなかった。

## License metadata

installed metadataの記録であり、法的判断ではない。

- pypdf 6.10.2: `License-Expression: BSD-3-Clause`
- tkinterdnd2 0.4.3: classifier `License :: OSI Approved :: MIT License`
- PyInstaller 6.19.0: license field `GPLv2-or-later with a special exception which allows to use PyInstaller to build and distribute non-free programs (including commercial ones)`、classifier GPLv2
- pyinstaller-hooks-contrib 2026.7: classifiers Apache Software License / GPLv2
- PyMuPDF / fitz: 使用していない。SplitSelectのPyMuPDF blockerをOneへ引き継がない。

Oneで現在確認したformal shipping blockerは、onedir folder一式を扱う共通出荷ツール / Launcher / BOOTH / Release工程の整備。SelectとOneの両方をmergeした後に共通側で一度だけ対応する方針とし、このPRでは変更していない。

## Static verification

- `compileall`: PASS
- UTF-8 read / AST parse: PASS
- UI_TEXT外の日本語UI文字列: 0
- UI_TEXT未使用 / 未定義 / 重複key: 0 / 0 / 0
- unused imports: 0 (AST name use check)
- GUI通常import直後の `sys.modules['pypdf']`: false
- dead `_clear_previous_outputs` / `OUTPUT_PATTERN`: 0件
- worker event generation: loaded / progress / done / errorすべてに付与
- stale completion / Explorer: generation checkより後だけ実行
- threadからTk直接操作: なし。workerはserviceとqueue notifierだけを呼ぶ
- temp cleanup: success / cancel / error / close test PASS
- `git diff origin/main...HEAD --check`: PASS
- build / dist / spec / version_info / config / pycache / PDF fixture / venv: ignoreまたはrepo外、commit対象0

最終rebase後の `tests/test_lifecycle.py`: 6 tests、1.465秒、PASS。

## Git status

- branchは最新 `origin/main` (`ddf5d14`) の上に2実装commitをrebase済み。
- report commit後の `git status --short` は空であることを確認してからpushする。
- commit対象は `01_apps/DAKE_PDF_SplitOne` のみ。
- mainへmergeしていない。

## Human Review

次の項目は未確認であり、merge / 正式出荷前にWindows実機で確認が必要。

- 最終onedirのcommon icon目視
- 中央panel clickからPDF追加
- Explorerからのphysical DnD
- GUIで3 / 30 / 100 / 300ページ分割、progress表示
- completion dialogのOK後にExplorerが開くこと
- GUI repeat splitで既存成果物が変化しないこと
- F5通常refresh
- 300ページ処理中F5 → 即idle → 直後に別3ページPDF
- 上記後に旧progress / done / error / dialog / Explorerが出ないこと
- Ctrl+O
- 分割中closeとnormal close
- 125% / 150% / 200%実DPI
- 実務スキャン / encrypted / malformed / 極端に重いPDF
- onedir対応後の共通shipping tool / Launcher / BOOTH / Release工程
