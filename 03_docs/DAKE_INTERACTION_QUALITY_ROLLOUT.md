# DAKE操作品質基準 v1 横展開調査

## 調査概要

- 初回調査日: 2026-08-27
- Sランク横展開完了・Aランク再評価日: 2026-08-28
- 対象: `01_apps/` 配下の全68アプリ
- 確認範囲: 各アプリの `main.py` と、必要な関連Pythonファイル（合計157ファイル）
- 方法: 起動、blocking、描画、cache、thread、完了体験の静的コード調査
- Phase 5までの実装対象: 代表基準アプリ `DAKE_PDF_Merge` とSランク5アプリ

本調査は横展開候補を選ぶための一次監査であり、各アプリのbuildや実機性能計測は行っていない。検出語の件数だけでは判定せず、呼び出し経路と処理粒度を確認した。例えばlaunch-checkと通常起動が別々の `tk.Tk()` を持つ場合は、同一起動内の二重生成とは数えていない。

調査時のworktreeには対象外の既存変更が存在した。特に `DAKE_PDF_Extract` は変更中、`DAKE_PDF_Insert` は未追跡の状態を読んだため、この2件の判定は調査時点のスナップショットに対する暫定評価である。本工程ではいずれも変更していない。

## DakePDF結合で確認した成功状態

### 起動

- 本番rootを1回生成し、フォント候補判定もそのrootを使う。
- `pypdf` と `fitz` は安全なloader経由で必要時に読み込む。
- Window icon適用を初期化経路で重複させない。
- GUIを構築してから、必要なバックグラウンド処理を開始する。

### PDF追加

- サムネイルjobは `Queue` に積み、固定3 workerで処理する。
- workerからTk Widgetを触らず、結果queueを `after` で反映する。
- PDF単位でページ数とサムネイルをcacheし、並び替えでは再取得しない。
- PyMuPDFでページ数と先頭ページ画像を同じopen内で取得し、失敗時だけpypdfへfallbackする。

### 並び替え

- `self.files` を正本として画面順と結合順を一致させる。
- 並び替え時はカードをdestroy/recreateせず、既存Widgetを再配置して番号だけ更新する。
- 生成済みサムネイルを再利用する。
- ドラッグ中だけ一覧端で自動スクロールし、終了時にtimerを確実に解除する。

### 結合と終了

- 結合はworker thread、UI反映はqueue / `after` で行う。
- 取得済みページ数を利用してページ単位に近い進捗を表示する。
- ページ処理ループ内でもキャンセルを確認する。
- 完了、エラー、キャンセルを分け、保存結果と保存先への導線を明確にする。

これらは単発の高速化ではなく、「起動が速い、触れば即反応する、余計な待ちがない、思った通り動く、終わったことが分かる」という一連の体験を成立させた要因である。

## 分類結果

| 分類 | 件数 | 判断 |
|---|---:|---|
| S | 5 | PDF結合と同じく、固定worker、最新job制御、Widget再利用、cacheが明確に効く |
| A | 38 | 改善余地はあるが、現行も非同期化済み、利用件数が限定的、または個別設計の確認が必要 |
| B | 20 | 現状軽量、入力規模が小さい、または今回の手法を加える利益が小さい |
| 対象外 | 5 | ゲームloop、Flaskサービス、Win32ネイティブUIなど、今回のTk業務アプリ型と構造が異なる |
| 合計 | 68 | `01_apps/` の全アプリ |

Sは「問題が深刻」という順位ではなく、「同じ解決が効き、改善効果を確認しやすい」候補である。

## Sランク上位5アプリ

### 1. DAKE_Image_BatchPDF

- 問題候補: 画像1件の読込完了ごとに一覧canvasを全消去・全再描画するため、まとめて追加すると処理量が二次的に増える。選択・移動・削除でも全行を描き直す。
- 改善候補: 追加行だけ生成し、選択は旧行と新行だけ更新する。並び替えは既存canvas itemの再配置を優先し、連続追加中の全体layout更新をまとめる。
- 期待効果: 10〜30画像で追加中の引っかかりと操作後のちらつきが減る。
- 危険度: 低〜中。PDF生成順と内部リスト順の一致を回帰確認する。

### 2. DAKE_PDF_Reorder

- 問題候補: ドラッグ中に挿入位置が変わるたび、全ページのcanvas itemを削除・再生成する。ページ数に比例してドラッグが重くなる。
- 改善候補: ドラッグ中は挿入markerと自動スクロールだけ更新し、既存ページitemを移動する。必要な全体整列はdrop確定時に1回へ集約する。
- 期待効果: 大量ページPDFでもドラッグ追従が滑らかになる。
- 危険度: 中。選択ページ、ページ番号、保存順、scroll位置を一体で確認する。

### 3. DAKE_PDF_Viewer

- 問題候補: PDF追加ごとにsummary thread、表示要求ごとにrender threadを新規生成する。tokenで古い結果は捨てるが、速いページ移動やzoomで古い処理自体は残る。fit表示倍率計算ではUI thread上でPDFをopenする経路がある。
- 改善候補: summary用固定workerと、最新要求優先のrender queueを分ける。ページ寸法・summary・描画結果を必要範囲でcacheし、fit計算から同期openを外す。
- 期待効果: 連続操作時のthread増加とCPU競合を抑え、ページ送り・zoomの応答を安定させる。
- 危険度: 中。古いrender結果の破棄、終了時worker停止、印刷・検索とのfitz同時利用を確認する。

### 4. DAKE_PDF_Merge_Mini

- 問題候補: 移動・削除などでカード領域をdestroyし、全カードを再生成する。上限5件のため影響は限定的だが、PDF結合と同じ改善を小さく適用できる。
- 改善候補: ファイルごとのカード参照を保持し、既存カードの再配置と番号更新だけで順序変更する。サムネイルは再利用する。
- 期待効果: 操作直後のちらつきが減り、PDF結合製品間の操作品質が揃う。
- 危険度: 低。5件上限、削除後の参照解放、結合順だけを重点確認する。

### 5. DAKE_PDF_CheckStamp

- 問題候補: 位置・大きさなどの連続プレビュー要求ごとに新規render threadを作る。tokenで表示競合は防ぐが、古い描画処理は並行して残り得る。
- 改善候補: debounce済み入力を最新1jobへ集約し、単一render workerまたは置換可能queueで処理する。ページ寸法と原画像をcacheする。
- 期待効果: 連続調整時のCPU負荷とプレビュー遅延が安定する。
- 危険度: 低〜中。最新の押印位置と保存結果の一致を確認する。

## Sランク横展開完了総括（Phase 5）

初回調査で選んだSランク5件は、2026-08-28までにすべて個別実装・build・exe起動・出力回帰を完了した。以下の数値は各Phaseで同じ操作を繰り返した計測、または変更前後コードを同じTk計測器へ通した結果である。件数上限を持つアプリでは、製品仕様内の実PDFテストと、表示modelだけのストレス計測を分けた。

| アプリ | 変更前問題 | 採用原則 | 変更後結果・実測 | 固有実装 | 展開可能な原則 |
|---|---|---|---|---|---|
| DAKE_Image_BatchPDF | 画像追加完了、選択、上下移動、削除、resizeで一覧Canvasを全消去・全生成 | 差分更新、生成済み資産再利用、内部配列を正本化 | 30画像の逐次追加で生成対象行は累計465行相当から30行へ。選択は全30行再描画から変更前後2行のstyle更新、上下移動は全30行から対象2行の再配置へ。3 / 10 / 30画像のPDF出力順を確認 | 1画像を複数Canvas itemで構成する `ImageCanvasRow` を各itemへ保持 | 一覧itemをdata itemへ結び、追加・選択・移動・削除ごとに影響範囲だけ更新する |
| DAKE_PDF_Reorder | drag targetが変わるたび全page itemを削除・再生成し、drag中にも順序を確定 | 差分更新、drag previewとdrop確定の分離 | drag中のCanvas全削除0、page item新規生成0、thumbnail再生成0。3 / 10 / 30 / 100ページでmarker、drop後再配置、保存順を確認 | Canvas item IDをページmodelへ保持し、drag中は1本のmarkerだけ `coords` 更新 | 高頻度Motionでは同じtargetを省略し、drop時に1回だけ内部順を確定する |
| DAKE_PDF_Viewer | summary・render要求ごとにthreadを生成し、tokenで表示を捨てても古いrender自体は継続 | Latest Job、固定worker、generation、UI thread分離 | render要求ごとの新規thread生成を廃止し、render workerは固定1、summary workerは固定2。連続ページ移動50回、zoom 30回、resize 100回で待機要求を最新1件へ置換し、古い結果のUI反映0、最終ページ一致を確認 | page / zoom / rotation / view modeをsnapshot化したrender request | 新要求で旧結果が不要になる表示処理は、FIFO完遂ではなく最新要求を優先する |
| DAKE_PDF_CheckStamp | preview要求ごとにthreadを生成。stamp位置変更でも背景再構築経路があった | Latest Jobと差分更新の併用 | preview 50回はthread / render / 最大同時renderが50 / 50 / 50から0 / 1 / 1へ。page変更30回は30 renderから2 render、古い結果反映0。stamp位置100回は背景更新100から0、overlay更新100。resize 100回は背景・stamp更新各100から各1 | PDF背景renderとstamp overlayを分離し、generation付きpreview requestを固定1 workerへ渡す | 最新要求優先に加え、背景不変ならoverlayだけ更新する |
| DAKE_PDF_Merge_Mini | 並び替え・削除で全カードWidgetを破棄・再生成。起動時にfitz / pypdfをimport | 差分更新、遅延読込、内部配列を正本化 | 並び替えのWidget生成/破棄は3件15/15、5件25/25、合成30件150/150から全件0/0。削除は全カード破棄から対象カード5 Widgetだけへ。import中央値295.99msから63.59ms。3件と仕様上限5件で出力順一致 | 最大5件の単純なTk Frameカードを保持し、影響開始index以降だけgrid再配置 | 小規模アプリでも不要な全再生成を残さず、重いimportは安定性を確認して利用時へ送る |

### 代表基準アプリを含む6アプリ

| アプリ | 基準として確定したこと | 主な確認値 |
|---|---|---|
| DAKE_PDF_Merge | root 1回、遅延import、固定3 thumbnail worker、PDF単位cache、カード再利用、ページ単位進捗・cancel、自動scroll | import約0.30秒→0.08〜0.09秒、onefile GUI約2.17秒→中央値約1.47秒。10 PDF / 30ページを1.793秒で順序どおり結合。30 PDFでもworker=3 |
| DAKE_Image_BatchPDF | Canvas rowをdata itemへ保持し、追加・選択・移動・削除・resizeを差分更新 | 30画像追加の行生成465相当→30、選択・上下移動は影響2行のみ |
| DAKE_PDF_Reorder | drag中はpreviewだけ、dropで1回確定 | 100ページでもdrag中の全Canvas delete / item生成 / thumbnail生成は0 |
| DAKE_PDF_Viewer | 最新表示だけを固定workerでrender | page 50、zoom 30、resize 100の連続操作で古い結果反映0 |
| DAKE_PDF_CheckStamp | Latest Jobと背景/overlay差分更新を組み合わせる | preview 50→実render 1、stamp 100→背景render 0、resize 100→最終1回 |
| DAKE_PDF_Merge_Mini | 小規模でもカードを捨てず、起動前importを減らす | import中央値約78%削減。並び替えWidget再生成0。onefile操作可能ウインドウ1.990秒 |

Merge Miniは最大5 PDFが製品仕様である。10 / 30 PDFの実投入は仕様どおりPDF I/O前に拒否し、3 / 5 PDFで結合結果を検証した。10 / 30件については上限を変えずにカードmodelだけを合成し、並び替え・削除・resizeの再生成回数を測った。

## DAKE操作品質パターン A〜E

| Pattern | 意味 | 実証アプリ | 判定 |
|---|---|---|---|
| A: 差分更新 | 変わった表示だけを更新する | PDF Merge、Image BatchPDF、PDF Reorder、PDF CheckStamp、PDF Merge Mini | 5アプリで実証。Widget、Canvas row、page item、overlayで実装は異なる |
| B: Latest Job | 古くなった待機jobを置換し、古い結果を表示しない | PDF Viewer、PDF CheckStamp | 2アプリで実証。正本には原則を記載済み、module化は3例目まで保留 |
| C: 生成済み資産の再利用 | thumbnail、metadata、PhotoImage、Canvas item、Widgetを軽い操作で作り直さない | 代表PDF Mergeを含む6アプリ | 6アプリで実証。cache有無より「再取得しない責務」を共通化する |
| D: UI thread分離 | 重い処理はworker、結果はqueue / afterでUIへ返す | 代表PDF Mergeを含む6アプリ | 6アプリで実証。worker数・queue形式はアプリ特性ごとに決定する |
| E: 内部状態を正本化 | 内部配列 / state = 画面 = 出力結果 | 代表PDF Mergeを含む6アプリ | 6アプリで出力順・ページ順・stamp位置を回帰確認 |

## 全アプリ調査一覧

「共通化候補」は候補となる設計要素を示すだけで、この工程でmodule化する意味ではない。

| アプリ名 | 分類 | 問題候補 | 改善候補 | 期待効果 | 共通化候補 | 注意事項 |
|---|:---:|---|---|---|---|---|
| DAKE_App_Dashboard | A | 全体scanはworker化済みだが、個別操作ごとにthreadを作る経路と複数の外部起動がある | 同時実行guard、scan差分cache、status統一を確認 | 大規模repoで安定 | UI queue、status | 外部ツール起動を直列化しすぎない |
| DAKE_App_Doko | A | scan後にカードを全再生成する | 結果差分更新、exe情報cache | 再scan表示が軽い | scanner、カード差分更新 | 2つのrootは通常起動とlaunch-checkの排他的経路で二重生成ではない |
| DAKE_Approve_Brainz | B | 単発worker中心で入力規模が小さい | 現状維持、長時間処理だけ計測 | 変更利益は小さい | status | self-checkのsleepは通常操作と分けて評価 |
| DAKE_Backup | A | backup対象走査と複数file処理 | 走査結果cache、進捗粒度とcancel確認 | 大量fileで安心 | file job queue | 非破壊・復元性を性能より優先 |
| DAKE_BGM_Loop | A | audio/subprocess状態管理が個別threadに分散 | 単一再生controller、終了処理整理 | 再生切替が安定 | status lifecycle | 音切れとUI速度を別に測る |
| DAKE_BOOTH_Assist | A | 大きな業務flowと外部処理を持つ | stage別job管理、結果cache | 長い処理でも状況が明確 | UI queue、status | 公開・販売系の外部状態は自動変更しない |
| DAKE_Brainz_OIKAWA | A | 複数panelを全再生成し、用途別threadが多い | panel単位差分更新、同種jobの排他 | 検索・scan後の描画が安定 | UI queue、status | 大規模なため個別計測後に限定改修 |
| DAKE_Brainz_Search | A | index/search/import等のthreadと全結果再描画が多い | job世代管理、結果virtualize、scan cache | 大量記憶で安定 | job controller、cache | requests/外部連携を含み、S手法の一括移植は危険 |
| DAKE_Column_Memo | B | 軽量メモUI | debounce保存を維持 | 現状で十分 | status timer | 全再生成は小規模固定UIのみ |
| DAKE_Document_Cover | A | reportlabを起動時import | 安全なら生成時lazy load | 初回windowが早い | lazy loader | PyInstaller hidden import確認必須 |
| DAKE_FAX_Cover | A | reportlabを起動時import | 安全なら生成時lazy load | 初回windowが早い | lazy loader | PDF生成失敗の表示を維持 |
| DAKE_Folder_List | A | directory走査と一覧生成 | scan worker結果のchunk反映、cancel | 深いfolderでも停止しにくい | file scanner、UI queue | 権限errorを握り潰さない |
| DAKE_Game_Alien_Road | 対象外 | real-time game loop | ゲーム用frame計測で別監査 | 業務アプリ基準の誤適用を防ぐ | なし | Tkでも描画モデルが異なる |
| DAKE_Game_Diver_Catch | 対象外 | real-time game loop | ゲーム用frame計測で別監査 | 同上 | なし | 操作latencyはgame基準で扱う |
| DAKE_Game_ShimarisuRealEstate | 対象外 | pygame loop | pygame profilerで別監査 | 同上 | なし | Tk worker patternを移植しない |
| DAKE_Git_Memo | A | repo/file走査と画面更新が同期寄り | scan差分化、root生成経路整理 | repo増加時に軽い | scanner、single-root | launch-check経路と通常起動を混同しない |
| DAKE_HolidayJinja_Post | B | 小規模画像と固定入力 | 現状維持、PIL初期化だけ計測 | 変更利益は小さい | icon/font | 画像品質を落とさない |
| DAKE_Image_BatchPDF | S | 追加1件ごと・操作ごとに一覧を全再描画 | 行の増分追加、Widget/item再利用 | 10〜30画像で明確 | list model、UI queue | PDF出力順を回帰確認 |
| DAKE_Image_HEICtoJPG | A | batch変換とPIL起動import | 固定job queue、必要時load検討 | 大量変換で安定 | worker queue、lazy loader | HEIC pluginの初期化順を確認 |
| DAKE_Image_iPhoneToPC | A | 取込・転送・previewの複数job | 同種job排他、metadata cache | 大量写真で安定 | worker queue、cache | device切断時の終了処理 |
| DAKE_Image_PasteA4 | A | capture・window一覧・PDF化が別thread | 連続capture guard、preview差分更新 | 連続操作が安定 | UI queue、status | A4配置精度を優先 |
| DAKE_Image_Receiver | A | server threadと画像処理、短いpoll sleep | server lifecycleと画像queueを明確化 | 受信連続時に安定 | worker queue、status | network受信をUI終了時に閉じる |
| DAKE_Image_Resize | A | batch処理は良好だがPIL起動importと一部一覧再構築 | lazy load可否、差分行更新を計測後判断 | 初回と大量追加を改善 | lazy loader、row renderer | 現行queue/afterを維持 |
| DAKE_Image_ToPDF | A | 変換workerはあるがPILを起動時import | lazy loadと複数入力job上限を検討 | window表示を早く | lazy loader | 変換error単位の表示を維持 |
| DAKE_Launcher | B | 固定件数のアプリ一覧 | 現状維持、一覧増加時のみ差分化 | 変更利益は小さい | icon/font | 起動導線の確実性を優先 |
| DAKE_Mail_Address_Format | B | 文字列整形中心 | 現状維持 | 即時処理のまま | status | clipboard例外だけ確認 |
| DAKE_Mail_AllStaff | B | 小規模入力と単発処理 | 現状維持 | 変更利益は小さい | UI queue | 宛先安全性を優先 |
| DAKE_Mail_Draft | B | 単発連携と短い待機処理 | 待機がUI threadならafterへ寄せる程度 | わずかな応答改善 | status timer | 文面・宛先を変えない |
| DAKE_Mail_Kikuta | B | 軽量固定UI | 現状維持 | 変更利益は小さい | icon/font | なし |
| DAKE_Mail_List | B | 軽量一覧と単発worker | 現状維持 | 変更利益は小さい | UI queue | address取扱いを優先 |
| DAKE_Maji_Memo | 対象外 | ctypesによるWin32ネイティブUI | message loop・timer基準で別監査 | 誤ったTk手法を避ける | なし | `after` / Tk Widget基準は適用不能 |
| DAKE_Mansion_Schedule | B | 計算・固定表示中心 | 現状維持 | 変更利益は小さい | icon/font | 日付正確性を優先 |
| DAKE_Music_Otooku | A | pygame/audioと複数workerを併用 | 再生controller、job排他、cacheを個別設計 | 音源切替が安定 | status lifecycle | audio engineを止めない |
| DAKE_Note_Inbox | A | 複数folder走査と外部open | scan cache、差分更新、job排他 | note増加時に安定 | scanner、UI queue | ファイル更新競合に注意 |
| DAKE_PDF_CheckStamp | S | 連続preview要求ごとにrender thread | 最新1jobのrender queue、page cache | 調整操作が滑らか | latest-job renderer | 保存座標との一致を確認 |
| DAKE_PDF_Compress | A | library lazy loadとworkerは既に良好、一覧更新に小さな余地 | 大容量1件のcancel粒度と進捗を計測 | 長い圧縮の安心感 | lazy loader、UI queue | 圧縮品質を速度より優先 |
| DAKE_PDF_Crop | A | fitz/PIL起動importとpreview処理 | lazy load可否、最新preview job | 初回と連続調整が安定 | lazy loader、renderer | crop座標の回帰確認 |
| DAKE_PDF_Extract | A | 調査時版はworker/queue化済みだが重いimportと大量page時の余地 | 実機計測後にlazy load・page cache判断 | 起動と大量pageで改善余地 | lazy loader、UI queue | 既存変更中のため暫定評価、sleepはlaunch-check系 |
| DAKE_PDF_Insert | A | thumbnail領域の全再構築とresize job重複の可能性 | Widget再利用、最新resize job、page cache | 挿入位置変更が軽い | renderer、cache | 未追跡コードを読んだ暫定評価 |
| DAKE_PDF_LookHere | A | fitz起動import、同期PDF操作の可能性 | lazy loadと重い保存処理のworker確認 | 初回と保存時の応答 | lazy loader | 単一PDFなので過剰設計しない |
| DAKE_PDF_Marker | A | fitz起動import、preview/saveと外部openが分散 | 最新preview job、page cache、外部open整理 | 連続編集が安定 | renderer、cache | marker位置の正確性を優先 |
| DAKE_PDF_Merge | B | 今回の代表改善を実装済み | 品質基準のreferenceとして維持 | 基準点を保つ | worker queue、cache | 3 workerを全アプリへ強制しない |
| DAKE_PDF_Merge_Mini | S | 操作ごとにカードを全destroy/recreate | 既存カード再配置と番号更新 | ちらつき低減 | card model | 最大5件の制約を維持 |
| DAKE_PDF_Rename | B | 単純処理をworker化済み | 現状維持 | 変更利益は小さい | UI queue | rename衝突を優先確認 |
| DAKE_PDF_Reorder | S | drag中に全page canvasを反復再描画 | marker差分更新、既存item移動 | 大量pageで明確 | reorder model | 保存順・選択・scrollを同時回帰 |
| DAKE_PDF_SplitOne | A | pypdf起動import、大容量PDFの処理粒度 | lazy load、page単位cancel/progress | 大容量で安心 | lazy loader、progress | 出力欠落を起こさない |
| DAKE_PDF_SplitSelect | A | queue/workerは良好だがfitz/PIL/pypdfを起動時import | lazy loadとthumbnail cacheの実測 | 初回windowを早く | lazy loader、cache | PriorityQueue設計を崩さない |
| DAKE_PDF_ToImages | A | fitz起動importと大量page変換 | lazy load、page単位progress/cancel | 大量pageで明確 | lazy loader、progress | 画像memory上限に注意 |
| DAKE_PDF_Viewer | S | 入力・render要求ごとにthread、fit計算で同期open | 固定summary worker、最新render queue、cache | page送りが安定 | worker queue、renderer | fitz同時利用と終了処理 |
| DAKE_Price_Apportionment | B | 計算量が小さい | 現状維持 | 即時計算のまま | icon/font | 丸め規則を優先 |
| DAKE_Price_FixedTax | B | 計算量が小さい | 現状維持 | 即時計算のまま | icon/font | 税計算規則を優先 |
| DAKE_QPSC_Dashboard | A | scanと外部起動、一覧再生成 | scan cache、差分表示、job排他 | 大量projectで安定 | scanner、UI queue | 外部process状態を明示 |
| DAKE_Reform_Progress | A | 複数panelの再構築とfile操作 | 変更範囲だけ更新、file cache | 案件増加時に軽い | row renderer、cache | 業務data整合性を優先 |
| DAKE_Screen_WebP | A | capture/listenerの複数threadとPIL起動import | lifecycle整理、latest capture job | 連続captureが安定 | worker queue、lazy loader | hotkey解除を確実にする |
| DAKE_Screenshot_Print | A | screenshot/PIL処理とprint job | 処理中即時表示、lazy load可否 | 起動と印刷待ちを改善 | lazy loader、status | OS print dialogの待機と分ける |
| DAKE_Sticky_Memo | B | 固定小規模UI | 現状維持 | 変更利益は小さい | status timer | destroy/recreateは固定部品のみ |
| DAKE_Time_AdvancedTimer | B | timer中心で軽量 | timer driftだけ別確認 | 現状で十分 | status timer | mainloopをblockしないことを維持 |
| DAKE_TwoPerson_Memo | A | 複数workerとpanel再構築 | 保存job排他、差分panel更新 | 連続編集が安定 | UI queue、row renderer | 同時編集・保存競合に注意 |
| DAKE_Video_Shorts_Cut | A | ffmpeg/probe/serverを含む長時間処理 | stage job controller、cancel粒度、進捗整合 | 長い動画でもUI維持 | progress、status | 動画処理は専用設計、worker数固定を安易に適用しない |
| DAKE_Wake_Brainz | 対象外 | FlaskサービスでTk GUIを持たない | request latencyとserver lifecycleを別監査 | 適切な指標で評価 | なし | pingはHTTP request側のblockingとして扱う |
| DAKE_Web_Dashboard | A | 複数file scanと大きな一覧 | scan cache、差分更新、chunk反映 | site数増加時に安定 | scanner、UI queue | build/deployは本調査対象外 |
| DAKE_Web_Index | A | scan workerはあるが一覧更新と外部処理に余地 | 差分更新、同時実行guard | 再scanが軽い | scanner、UI queue | URL/fileの正本関係を守る |
| DAKE_Work_Calendar | B | 小規模計算と単発worker | 現状維持 | 変更利益は小さい | UI queue | calendar正確性を優先 |
| DAKE_Year_Age | B | 計算量が小さい | 現状維持 | 即時計算のまま | icon/font | なし |
| DAKE_Year_Notice | B | 起動時の固定計算だけ | 現状維持 | 変更利益は小さい | icon/font | single-rootで軽量 |
| DAKE_Yesterday_Task_Memo | B | 固定小規模UI | 現状維持、保存debounce確認 | 変更利益は小さい | status timer | 全再生成は固定部品のみ |
| DAKE_Yukiz_KadouChu | A | 29 Pythonファイル、動画・画像・多数の操作別threadを持つ | pipeline単位の排他、job管理、cacheを個別監査 | 長い制作flowが安定 | job controller、status | 高複雑度。小さな横展開を先に検証してから触る |
| DAKE_YukizBlog_Post | A | file走査と投稿準備worker | scan cache、同時実行guard | 記事増加時に安定 | scanner、status | 投稿・公開は明示操作を維持 |

## 共通化候補

### 先に共通仕様へ揃えやすいもの

| 候補 | 確認した重複 | 次の判断 |
|---|---|---|
| フォント検出 | 多くのTkアプリが同じ候補探索を個別実装 | 既存rootを受け取る関数形を2〜3アプリで揃え、build影響がなければmodule候補 |
| Window icon | icon path探索と `iconbitmap` の例外処理が反復 | 配布形態ごとのpath差異を整理後、小さなhelper候補 |
| UI queue pump | worker → queue → `after` が多数アプリに存在 | event型、終了条件、例外通知の差が小さい3アプリで検証 |
| status animation | ドット進行、reset timer、終了時cancelが反復 | 見た目ではなくtimer lifecycleだけを共通候補にする |

### Sランク改修後の最終判断

| 候補 | 実コードで確認した重複 | 判定 |
|---|---|---|
| フォント検出 | 6アプリ中4アプリが既存rootからfont family候補を探索する小関数を個別保持 | 小さな共通helper候補。ただし候補fontとfallback差を揃えてから。今回module化しない |
| Window icon | 6アプリでicon path解決と `iconbitmap` の例外処理が反復 | 最有力の小helper候補。onefileの `_MEIPASS` とsource実行path差を統一テスト後に判断 |
| UI queue pump | 6アプリすべてにworker結果queueと `after` pollがある | event payload、poll間隔、終了条件、dialog責務が異なるため現時点では各アプリ保持 |
| worker停止lifecycle | ViewerとCheckStampがcondition、pending最新1件、generation、stop flag、timer cancelを持つ | `Latest Job Renderer` パターンとして文書化。2例ではmodule化せず、3例目で同じlifecycleなら再検討 |
| debounce timer helper | Viewer / CheckStampのresize、Viewerのzoom、各アプリのlayoutで `after_cancel` + `after` が反復 | APIを小さくできる可能性あり。例外時・終了時cancelまで同一化できる3例を待つ |
| 遅延loader | PDF Merge / Merge Miniでfitz・pypdfの利用時importが有効 | import対象、dependency error、PyInstaller hookが異なるため関数本体は各アプリへ小さく持つ |
| file/page cache | PDFページ数、thumbnail、previewで反復 | key、無効化、memory上限、document lifecycleが異なるため共通module化しない |
| list/card再配置 | BatchPDF、Reorder、Merge Miniで有効 | Canvas row、複数Canvas item、Tk FrameでAPIが一致しないため共通module化しない |
| PDF renderer / request payload | ViewerとCheckStampで最新jobが有効 | page pair、zoom、rotation、stamp状態、document lifecycleが異なるため共通module化しない |

共通化の入口は「同じ名前のhelperを作ること」ではなく、2〜3アプリで同じ責務・同じ失敗条件・同じ終了処理が確認できることとする。共通moduleで理解が難しくなる場合は、品質基準と小さなローカル実装だけを共有する。

## Aランク38件の再評価と次候補上位10件

2026-08-28の `origin/main` でAランク38件・Python 126ファイルを再走査した。評価軸は、ユーザー体感改善幅、起動速度、UI停止リスク、大量データ頻度、改善安全性、代表利用頻度、他アプリへの学習価値である。検出件数だけでなく呼び出し経路を確認した。`DAKE_PDF_Insert` は初回調査時の未追跡スナップショットにはPython実装があったが、現在の `origin/main` には追跡Pythonがないため上位候補から外した。

| 順位 | アプリ | 現在コードで確認した候補 | 最初の改善単位 | 期待効果 / 安全性 |
|---:|---|---|---|---|
| 1 | DAKE_PDF_Marker | PDF open、page render、zoom / resize renderがUI thread。page描画でpixmap生成後にCanvas全更新 | render requestをsnapshot化し、固定1 worker + Latest Jobへ。marker overlayは既存差分更新を維持 | ページ移動・zoomの体感改善大。保存workerは既存のため描画だけを分離できる |
| 2 | DAKE_PDF_LookHere | PDF open・page render・保存が同期経路。resize debounce後もUI threadで再render | open / render / saveを段階的にworker化し、まずrenderを最新要求優先へ | UI停止リスクが明確。単一PDF構造で責務を分けやすいが保存座標回帰が必須 |
| 3 | DAKE_App_Doko | scanはworker済みだが結果反映でカード一覧を再構築。exe情報を再走査 | scan結果をkeyで比較し、変化カードだけ追加・更新・削除 | Sランク差分更新を安全に再利用でき、一覧アプリへの学習価値が高い |
| 4 | DAKE_App_Dashboard | 68アプリscanはworker済み。filter / searchごとにTreeview全行をdelete / insert | record IDを正本にし、filter差分または再利用可能な行更新を計測 | 代表利用頻度と体感効果が高い。3,236行の大アプリなので表示層だけへ限定する |
| 5 | DAKE_PDF_SplitSelect | fitz / PIL / pypdfを起動時import。thumbnailは固定PriorityQueue workerだがCanvas全再描画経路あり | 起動import実測後の遅延化、viewport内item保持と選択差分更新 | 大量ページ効果大。既存worker優先度とcacheを壊さない限定改修が可能 |
| 6 | DAKE_Image_iPhoneToPC | 取込・metadata・previewに3種thread。大量写真時の同種job競合余地 | job種別ごとのinflight guard、metadata / thumbnail再利用、終了lifecycle監査 | 大量入力頻度が高く効果大。device切断を含むため安全性は中 |
| 7 | DAKE_Image_Resize | batch worker / queueは既に良好。PIL起動importと一覧再構築に限定余地 | 起動計測とrow差分更新だけを小さく適用 | 危険が小さくImage BatchPDFのrow原則を再検証できる |
| 8 | DAKE_Image_PasteA4 | capture、window列挙、PDF化に別job。preview Canvas全更新と短いsleep経路あり | 連続capture guard、背景と配置itemの差分更新、終了timer整理 | 連続操作の追従改善。A4配置精度を固定したまま表示層を限定可能 |
| 9 | DAKE_PDF_ToImages | fitz起動import、大量page変換を単一workerで処理 | 遅延import、page単位progress / cancel、画像保持上限を実測 | 大量ページで完了品質が上がる。出力画質を触らずjob粒度だけ改善可能 |
| 10 | DAKE_PDF_Crop | page変更ごとにload threadを生成し、load IDで古い結果だけ破棄。preview Canvasは全更新 | 固定1 latest preview workerを検証し、selectionは既存tag差分更新を維持 | Viewer / CheckStampに続くLatest Job 3例目候補。crop座標回帰が必要 |

上位外でも、`DAKE_PDF_Extract`、`DAKE_Brainz_Search`、`DAKE_Yukiz_KadouChu` は改善余地が大きい。ただしコード規模・外部状態・既存workerが複雑で、「効果が大きく既存ロジックへの危険が小さい」という今回の順序では後段とした。`DAKE_Document_Cover` と `DAKE_FAX_Cover` は現在コードでreportlabの利用時import経路が確認でき、初回調査時より優先度を下げた。

## Phase 5で実施しないこと

- Aランクアプリのコード変更
- 共通module作成
- version、GitHub Release、BOOTH、Store、dakeapp.comの更新
- `DAKE_PDF_Merge_Mini` 専用branchのmain反映（本Phaseでは検証結果を報告してから判断）

今後も各候補の着手時に、正本・最新コード・変更前実測を確認してから順位と手法を再確定する。
