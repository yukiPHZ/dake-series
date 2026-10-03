# DAKE STADIO 技術立て直し実行記録

更新日: 2026-09-30。対象: **0.4.0-eval.1 / Windows x64**。仕様の正本は [ORIGINAL.md](ORIGINAL.md) の「2026-09-30 技術立て直し・統合制作の正本仕様」。本書は設計判断、実行順序、実装状況、再開手順を残す派生記録。検証の最終判定と配布物の識別は [TEST_REPORT.md](TEST_REPORT.md) を参照する。

**最終ZIP展開EXEで、新規設定と既存設定の安全な複製の両方を使い、通し制作・正常終了・オフライン再起動・再編集・出力を確認した。** 以下の合格は今回の各検査が確認した範囲を指す。ソース実行の合格や旧版の合格を、最終配布版・全利用条件の完成根拠にしない。最終結果は TEST_REPORT と対応する evidence を照合する。

## 保全と初期監査

- `artifacts/technical-recovery-20260930/` に着手時の src / desktop / scripts / tests / assets とルート資料を保全した。以前からの未コミット変更、旧0.3.0配布フォルダ・ZIPは維持する。
- 0.3.0 EXE SHA256: `ebff99d49ecdc2b2f4c53d224ec2f89c39342a717286410ef64202ea8d530014`。app.asar SHA256: `990fb93832a58899fd8e6163367dcdc8d1afd3d276de620064ce25860a6d847d`。配布主要コードと着手時コードは一致し、古いEXEの取り違えを原因とする根拠はなかった。
- 設定の解決順は `STADIO_USER_DATA` → EXE横に `.stadio-portable` がある場合の `user-data` → `%APPDATA%\DAKE STADIO`。実利用は通常のAPPDATAで、取得済みGoogle Fonts 9種と復旧データがあった。元設定を削除せず、専用領域へ安全に複製して検査する。
- 検査用profile、素材、保存先は `test-output/` 以下。ユーザー作品や設定のコピーを配布へ入れない。作業中の旧EXEに残る未保存作品を終了・移動して検査を成立させない。
- 見本文字入力の TypeError、新規文字の Yu Gothic 固定、未取得Google書体の代替表示、Windowsの太字見本の不足を再現した。Windows 301 faces の列挙と既存Googleキャッシュの登録は修正前にも成功しており、「すべての取得が故障」とは断定しない。詳細と前後比較は [フォント監査](docs/FONT_RELIABILITY_AUDIT.md)。

## 設計と実装順序

既存の Electron / Fabric / Paper.js / worker / 原子的保存を維持した。単一Engineを捨てて全面移植するのではなく、文書セッションとプレビューの確定境界を追加した。新規依存の追加なし。UI文言は `src/ui-text.json` に統合する。

| 段階 | 依存関係と到達条件 | 現在の状態 |
| --- | --- | --- |
| 0 保全・再現 | 配布識別、元設定保護、フォント失敗段階の切り分け | 今回の修正前EXEで再現。条件と未再現範囲を監査へ記録 |
| 1 フォント | 実バイナリ登録 → 見本・現文字 → 新規文字 → 文書依存と再開 | 実装、fresh/copied最終EXE、再起動・オフライン・画素比較を確認 |
| 2 セッション | 保存先・dirty・Undo・素材・非同期結果を文書単位に分離 | 実装。複数作品、遅延保存の帰属、復旧保護を確認 |
| 3 色 | 共通プレビュー境界 → 確定1 Undo・取消 → 履歴・登録色 | 実装。実レイヤーへの途中反映、確定、取消、Undoを確認 |
| 4 編集可能な部品 | v3検証 → 子文書 → 親プレビュー → 確定・保存 → リンク能力 | 実装。子編集、親1 Undo、再読込み・別プロセス・参照切れを確認 |
| 5 補正・マスク | 原画像 / recipe / 2種mask → worker → Undo・保存・出力 | 実装。画素と原画像保持、実ポインタ、保存復元を確認 |
| 6 寸法・出力 | 作品操作を分離 → 実エンコード → previewと保存物の同一性 | 実装。色との結合、PDF倍率・物理寸法、実測容量を確認 |
| 7 配布検証 | 最終ZIP全内容照合 → fresh/copied EXE → 指定制作 → 再起動 → 再編集・出力 | 確認済。最終実行コード・画面・出力結果は TEST_REPORT へ集約 |

## フォントの責務を分離

`desktop/fonts.cjs` と `desktop/font-binary.cjs` は取得・キャッシュ・実バイナリ・許諾・cmapを扱う。`src/production-ui.js` は実行時登録、作品依存、適用対象と次の文字の既定書式を管理する。`src/font-coverage.js` は収録文字をローカルで確認する。

- Windowsも実バイナリから FontFace を読み込み、指定faceの太さ・斜体を見本へ渡す。未取得Google書体は正しい見本であるように代替表示せず、取得、見本、適用を分ける。
- 新しい文字の作成は既定書体の準備完了を待つ。既存文字への適用と同じdescriptorを引き継ぎ、作品に必要なGoogleバイト・許諾を保存する。
- タブ切替で取得ライブラリを消さず、各作品の実バイナリ版を有効にする。一作品内の同じfamily/style/重なるweightへ異なる版を混ぜる場合は競合として拒否し、黙って別字形に置き換えない。
- 登録失敗や取得中の対象変更は成功通知にしない。収録されていない文字は代替表示の警告を出す。Windowsバイトは作品へ埋め込まないため、別PCへの書体同一性は保証しない。

## タブ・部品・保存の整合

`src/workspace.js` が文書ごとのシーン、履歴、素材表、view、既定書体、保存metadataを退避復元する。既存UIが参照するEngineは一つを維持し、非同期結果を見分けるrevisionを巻き戻さない。保存完了は開始したセッションにだけ戻す。同じ保存パスの通常文書を開いた場合は既存タブへ移動し、未保存内容を二重タブで上書きしない。

`src/engine-components.js` の部品は Fabric Group の見た目に加え、**編集可能な子文書を正本として保持**する。Group内の見た目は読込み時に子文書から再構成する。埋込みは独立コピー、リンクは最終snapshotと出所を持つ。

- 子タブでは文字・ベクター・画像・レイヤーを編集できる。子タブ内の配置先プレビューを更新し、親の確定内容は「配置先へ反映」まで維持する。反映は親の1 Undo、取消は親を変えない。
- プレビューは単一の描画処理と最新要求の待ち行列にまとめる。古い結果を後から採用しない。親で配置が消えた場合は子の編集内容を独立した未保存文書として残す。
- 子タブから元ファイルは保存しない。元作品を通常の文書として明示して開いたタブでの Ctrl+S は独立して許可する。この明示許可がない部品取込みだけでは元作品への書込み能力を与えない。
- `desktop/components.cjs` は選択されたファイルの読取り能力、canonical path、hashを管理する。外部リンクは自動更新せず、明示更新する。再起動後は元ファイルを選び直して読取りを再承認する。参照切れでも保存済みの編集可能snapshotは開ける。
- 新規保存はv3。旧v1/v2を読めるが、新版保存後のv3を旧EXEで開くことはできない。原子的保存と `.bak` を維持する。比較・試用には別名保存を勧める。
- 文書IDの祖先検査と出所pathの検査で自己参照・循環を拒否する。部品の深さ8、文書64、全体12,000オブジェクト等の上限を検証する。タブは16、workspace復旧は総量128 MiB以下。これらは快適性の保証値ではない。

## 復旧・終了のレビューで直した問題

最初の統合レビューで、古いworkspace復旧を選ぶと現在の未保存タブが置き換わる問題を再現した。また、復旧案内を後回しにして新作を編集すると旧復旧ファイルが新しいautosaveや正常終了で失われる経路を確認した。次の境界へ変更した。

1. `desktop/recovery-manager.cjs` は起動時に残ったworkspaceと旧単一文書の復旧候補を `pending-recovery/` へ退避する。初回autosaveも退避完了を待ち、未読候補を上書きしない。
2. 復旧は現在のタブへ追加し、セッションIDと子の親IDを付け直す。上限を超える場合は現在作品を保持してエラーにする。復旧したパスから書込み権限を勝手に復元しない。
3. 復旧した全workspaceを新しい自動復旧へ保存できてから、読み終えたpending候補だけを消す。保存失敗時は候補を保持する。
4. 正常終了で現在の作業用復旧を整理しても、未読pending候補は残す。候補の破棄は一覧から明示確認したものに限る。旧単一文書の候補にも到達できる。
5. 終了前に未確定プレビューと子の編集を解決し、終了確認後に新しい編集が起きた場合は閉じずに復旧へ保存する。

同じ元 .dake を部品に配置した後、通常文書タブで明示保存できることも回帰確認した。UI上の同一タブ判定におけるhardlink/junction別名の全条件は未検証であり、nativeの元資産書込み保護と同一視しない。

## 色・画像・寸法・出力

- `src/color-picker.js`: 色面、色相、HEX、既定カラーセット、使用履歴12色、登録24色。編集中の色は作品へ即反映し、確定1 Undo、Esc取消。取消した色を使用履歴へ残さない。登録色の管理は作品とは別のローカル設定。
- `src/engine-imaging.js` / `src/imaging-ui.js`: 原画像と補正recipeを分離し、明るさ・コントラスト・露出・彩度・自然彩度・モノクロ・5点カーブを再編集する。補正ダイアログは光、色、カーブへ分け、同じ長い設定スクロールを増やさない。
- 透明度マスクと補正範囲マスクを独立保持。画像／透明度／補正範囲を明示して選び、筆で隠す・戻す、硬さ・筆径・ぼかし、矩形選択からの作成、解除を扱う。既存の図形クリップとは別データ。補正やmaskの取消、Undo、保存復元で原画を保持する。
- 作品のキャンバス寸法変更、作品全体の比例拡縮、レイヤー変形、トリミング、出力時リサイズを区別する。余白だけを増やした時は既存レイヤーの寸法を変えない。
- `src/output-ui.js` は実際のエンコード結果、寸法、形式、品質、透過、**実測**容量を表示する。最新のpreview tokenを保存に使用し、古い要求の結果で書き出さない。JPEG/PDFは白へ合成する。
- PDFの倍率変更で画素寸法と申告寸法がずれる経路を修正。画素密度と物理寸法を分離する。現方式はRGB画像1枚の1ページであり、ベクター/文字保持、CMYK/ICC、PDF/X、TrimBox/BleedBox、トンボは未対応。

## 今回確認した証拠

件数は状態・画素・境界の確認数であり、完成機能数ではない。自動操作には実Electronのポインタ/キー入力、rendererの診断API、fixture挿入が混在する。workspace系および画像UIの一部はネイティブファイルダイアログの返答だけを試験用パスへ置換している。Windowsの保存画面を人がすべて操作したという主張はしない。

| 検証 | 今回確認した内容 | 根拠 |
| --- | --- | --- |
| 修正前フォント | fresh/copied両環境の見本入力例外、新規文字書体、未取得見本 | `evidence/font-reliability-baseline/results.json` |
| 修正後フォント | fresh/copied各33確認。Windows Regular/Bold、Google実取得、見本/適用/新規文字、別プロセス・オフライン、独立登録FontFaceとの字形画素一致と代替書体との差 | `evidence/font-reliability-source-results.json`、`docs/FONT_RELIABILITY_AUDIT.md` |
| 部品の制作と再保存 | 23確認。子live、親1 Undo、Undo/Redo、元書込みなし、保存前後PNG一致、リンク切れ・再選択・循環防御 | `test-output/workspace/run-1790761624765/results.json` |
| 部品の別プロセス再開 | 9確認。cached子の編集、リンク再承認、更新Undo、元SHA不変 | `test-output/workspace/restart-1790761511940/results.json` |
| 保存権限と復旧レビュー | 6確認。通常元タブのCtrl+S、部品のみの無断保存拒否、同path再開、現在未保存タブを保った復旧 | `test-output/integrity-review-1790762286123/results.json` |
| 未読復旧の保全 | 7確認。新autosave、候補復旧、active整理後も旧候補保持、明示破棄 | `test-output/pending-recovery-1790762286091/results.json` |
| 部品・復旧のnative単体 | 4テスト。型/深さ/循環、read能力・hash、pending移行・選択破棄・再起動 | `tests/components.test.cjs`、`tests/recovery-manager.test.cjs` |
| 非破壊補正とmask | 27確認。旧4filter移行、原画保持、cancel、最新要求、透明度/部分補正、筆、保存復元画素一致 | `test-output/imaging-v4/report.json` |
| 画像UI | 8確認。補正ダイアログ、取消、実ポインタmask、矩形選択、native保存、例外なし | `test-output/imaging-ui-v4/report.json` |
| 色・出力結合 | 10確認。色live/確定/Undo/Esc、実測容量、PDF半分/2倍、キャンバス余白、v3保存と旧bak | `evidence/output-color-source-results.json` |
| 既存エンジン回帰 | 167確認。v1/v2、パス、図形演算、旧filter/clip、文字アウトライン、画素保存、入力保護等 | `evidence/engine-check-results.json` |
| 配布EXEの即時部品反映 | fresh10 / copied10、計20回。全回プレビュー未表示のまま文字入力・実クリックで反映し、親の内容一致。例外0、元fixture不変 | `scripts/check-component-rapid.cjs`、`test-output/component-rapid-1790763325475/report.json` |
| 最終ZIPと指定通し制作 | **確認済**。新規33/全設定コピー34、別PID・native offline・保存/再編集/実出力が一致。最終実行コードはASAR 81a7d817c20502d2d192f3b94243ba6255286265306c83b7a2d5047cf775d732 | `TEST_REPORT.md`、`evidence/evaluation-build.json`、`evidence/*-packaged-results.json` |

配布通し試験の1回で子の反映待ちがtimeoutした（`test-output/creation-flow-1790762728589/copied/report.json`）。当時はクリック受信と処理状態の記録が不足しており、原因未確定。通常手順の再試験に加え、preview完了を一切待たない独立回帰を行った。ASAR `7752d70a087033d26c9b52b6f6d2ea017b7ad9e064a3b8e2ad1036c219ef3d19` のEXEで全20回のクリック受信と内容一致を確認し、最長約2.4秒で親へ戻った。検査ウィンドウだけを非表示にし背景描画抑制を解除、プロファイルは隔離した。製品コードを変更して解消したという結果ではなく、元の単発条件は未再現として残す。

オフラインは再起動した試験プロセスでrenderer通信とnativeフォントHTTPを禁止して確認した範囲。PC全体のネットワーク設定は変更していない。FontFace失敗注入は失敗表示を試す検査であり、現実のOS障害を再現したとはしない。旧0.3.0/0.2.0の過去の合格一覧を今回の再検証として加算しない。

## 未達・制約

- 最終ZIP展開EXEでの通し制作は新規/全設定コピーで完了。最終文書を含む再圧縮・全ファイル同一性は evidence/evaluation-archive-results.json、確認範囲は TEST_REPORT.md に記録する。評価版の到達点であり、以下の未達を残す。
- 同一Windows PCでの検査。別PC、すべてのWindows/Google書体、日本語IME各方式、高DPI・狭い画面、実ユーザーの全既存作品は網羅していない。今回再現しなかった元ユーザーの障害条件を解消済みとは断定しない。
- 大量レイヤー、大きな写真、深い部品、長時間利用の応答性は継続課題。入力上限まで快適という保証はなく、Undoは再起動を越えて保存しない。
- Windows書体の別PC同一表示、同一文書内の同名異版書体混在、取得ライブラリの高度な整理、混在書式・縦書き・高度組版は未対応。Google取得済み書体は許諾とバイトを作品へ保持する。
- maskは画像レイヤーと矩形選択を対象とする。自由形選択、自動被写体抽出、筆圧、色別RGBカーブは未対応。カーブは5点の区分線形補間。
- 外部リンクの自動監視・自動上書きは行わない。再起動後の更新は再選択が必要。hardlink/junction別名のタブ重複防止は未検証。
- 印刷用色管理とベクターPDF、PDF編集読込み、PSD/AI/RAW、複数アートボードは未対応。実印刷・印刷所入稿適性は未検証。
- 未署名のローカル評価版。公開、push、販売、課金、作品送信、OS関連付けの変更は行っていない。物理電源断・ディスク故障・別ファイルシステム上の全異常は未検証。

## 再開手順

1. 本書、ORIGINALの最新仕様、TEST_REPORT、`evidence/evaluation-build.json` を読み、検証したASARと次に触るコードの識別を照合する。`artifacts/technical-recovery-20260930/` と既存成果物を変更しない。
2. ユーザーの通常APPDATAを検査起動に使わない。`test-output/` の専用profileまたはその安全コピーを使う。旧版の未保存作品を検査のために閉じない。通常の試用では旧版で作品を保存して終了してから新版を起動し、同時起動による旧プロセスへの転送を避ける。
3. 依存が揃っていれば `npm run build`。コード変更時は `npm test`、`npm run test:engine` と変更範囲に対応する検証を実施する。フォントは `npm run test:fonts-ui`、画像は `npm run test:imaging`、色/出力は `npm run test:output-ui`。既存profile試験は `STADIO_QA_PROFILE_SOURCE` にコピー元を指定する。
4. workspace系は `node_modules\electron\dist\electron.exe scripts/check-workspace.cjs`、`scripts/check-integrity-review.cjs`、`scripts/check-pending-recovery.cjs` をそれぞれ同様に実行する。`ELECTRON_RUN_AS_NODE` があればこの試験プロセスの環境から外す。別プロセス再開は `STADIO_RESTART_FIXTURE` に検査で生成した `linked-card.dake` を指定して `scripts/check-workspace-restart.cjs` を実行する。各検査は自分のfixtureだけを更新する。
5. ソースの修正範囲が通ったら `npm run package:evaluation`。新しい日時別フォルダ・ZIPへ出力し、旧配布を上書きしない。フォント/色の各検証スクリプトには第1引数で展開EXEを指定できる。指定の通し制作とfresh/copied、終了後再開は最終配布物で実行し、ソース検査と区別する。
6. 今回追加のREADME・本記録・TEST_REPORT・BUILD_IDを配布文書へ反映し、再圧縮後はZIPと展開全件のhashを確定する。実行コード変更があれば対象のEXE検証をやり直す。改善・未達・短い試用手順とともにローカル評価ZIPを提出する。

検証用profile、ビルド、dist、artifacts、生成bundleはGit除外を維持する。元作品・設定やそのコピーを誤って追跡・配布しない。次段階は未達を小さく切り、実際の制作で現れた不具合を再現してから直す。
