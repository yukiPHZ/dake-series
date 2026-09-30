# v1.0.2 rc2-viewport — Human Review FAILへの修正（2026-09-30）

PR #47 / Issue #46。基点31a31f3513d12e6c0bb7b94be01e6a4323f61c73。
前回84 tests PASSは履歴として維持するが、実機400件の正常表示を検証できていなかった。
今回の状態は「Computer Use待ちだけ」ではなく「Human Review FAIL後の修正候補・再確認待ち」。
Ready化、main merge、tag、Release、BOOTH、サイト、Store、決済、Cloudflare更新は禁止・未実施。

## 再現報告と確定した原因

利用者実機: PDF400/400・失敗0でも小→特大で空白、途中以降が見えない、追加スクロールで復旧する場合がある。
小9列・特大4列となる1920px条件を追加した。これは26件疑いと同一原因とは断定しない。

同じWindows/Python3.12/Tk・1920×1080・400基準画像で旧commitと比較:

- 旧サイズ変更は毎回400画像を同期変換し、PhotoImageを400個作成。ハンドラ2.063〜4.907秒、最大イベント間隔4.913秒。
- 旧thumbnail結果は全件入力欄更新と変更待ち全走査を繰り返す。400結果のUI受領に23.227秒。
- 特大の単一Frameは60,300px。縮小後もactual heightが60,300pxに残る測定を取得。要求高さ・実高さ・scrollregionが不整合となる経路を確認。
- native Windowsクリッピングそのものの画素再現・厳密な上限値は未確定。Tk座標だけで利用者の空白症状の完全再現としない。

## 修正

全PDFのモデルと表示部品を分離。論理scrollregionに対して表示Frameをviewport＋先読みに限定し、全件分の巨大Frameを廃止。
可視＋上下1行だけ部品を再利用し、画像適用を8ms予算でUIスレッドに分割する。afterは各1予約、連続サイズ変更は最新要求へ集約。
ファイルIDと画面内相対位置で表示を維持。全件再scan/renderはしない。
全件の入力変数・snapshot・ページ数はモデルに保持。フォーカス中のEntryを別ファイルへ再利用せず、カーソル・選択・横位置を保持。
画面外の変更待ちもrename/Undoへ渡す。rename_core.pyは変更なし。
変更待ち集合を使い、thumbnail結果から全件control更新を切り離す。
基準画像はPDFごとに1個のlossless RGB圧縮データ、PhotoImageは表示部品分だけ。clear時に変数traceも解除し世代のモデル・画像を解放する。
PDFium mutexとpreview Latest Jobの責務・全件thumbnail処理は維持する。

## 検証

- pytest最終: **99 passed / 114.43秒**（前回84件を維持・仮想化に合わせ観測点を更新）。
- 最終点検で再利用時のTcl callback蓄積も防止。bindは部品ごとに1回、現在モデル参照のみ差し替え、100回再利用してcallback数一定を検証。旧wheel試験の架空Frame/手動scrollregionはConfigureとの競合で1回失敗したため、実モデル・実表示部品を使うwheel/refresh試験へ修正し全件再実行した。
- 400件×4サイズ×先頭/中間/末尾×読み込み前後×900/1180/1920幅×100/125/150%相当 = 216組み合わせ。
  ファイルID、Tk画像の実ピクセル色、可視矩形の交差、入力状態、メタ情報、末尾399を検証。全grid検査だけではない。
- 400件実合成PDF: 完了前表示、読み込み途中のサイズ/別フォルダ切替、ページ数・容量を独立期待値と照合、
  wheelのみ先頭/中間/末尾、末尾rename/Undo、reload400件維持、サイズ変更による再renderなし。
- 編集中Entryの同一性、日本語入力文字列、カーソル・選択、画面外変更待ちを保持。
- 境界0/1/23/24/25/26/27/32/47/48/49/100/300/400の合成PDF試験成功。
  異常PDF後の正常PDF、scan失敗明示、refresh/reload、preview/wheel、mutex回帰を維持。1000モデルの追加batch境界も検証。
- header/toolbar/status/footer: 900/1180/1920×100/125/150%相当の9条件自動幾何検査成功。

### 比較計測（同一400基準画像、15回切替）

|項目|旧RC|修正候補|
|---|---:|---:|
|サイズ変更ハンドラ|2,063〜4,907ms|0.04〜0.08ms|
|イベント間隔最大|4,913ms|109ms|
|可視範囲更新完了最大|約4,955ms|439ms|
|1切替の画像更新|400|16〜31（同サイズ0）|
|最大実Frame高|60,300px|2,695px|
|プロセスメモリ最大|826MiB|318MiB|
|400結果UI受領|23.227秒|0.119秒|

比較probeは両版へ同じ非圧縮基準画像を注入する保守的条件。worker圧縮後の実PDFメモリとは分ける。
実PDF400件: 最初の可視画像0.755秒（50/400）、全処理4.820秒。
読み込み中イベント間隔最大87ms、サイズ変更中最大115ms、変更表示完了最大400ms。
15回反復時メモリ78〜93.5MiB、最後78.6MiB。継続的な0.5秒超停止はこのheartbeat測定ではなし。
先行試行では表示収束547msもあった（操作停止とは異なる）。環境負荷に依存し全環境の上限保証ではない。
生の測定値はアプリ evidence/viewport-before.json、viewport-after.json、real-pdf-400.json、synthetic-boundaries.json。

## 未確認・Human Review手順

以上はソースTk自動試験であり配布exeのnative画面PASSではない。
Computer Useは Trusted RPC service is not configured: sky で接続不能。native画素、タスクバーアイコン、実IME変換中、OS DPI変更、GUI資源数は未確認。
固定distのrc2-viewport exeで400件を初回選択し、小の末尾→特大、特大中間→小、連続切替、
追加スクロールなしの復元、wheel/スクロールバー、入力中移動、末尾rename/Undo/reloadを利用者が再確認する。
公開画像は更新していない。Human Reviewは勝手にPASSへ変更しない。

## 作業場所

worktreeをGit管理のまま devlop/_worktrees/dake-series-overview-v1.0.2 へ移動した。
通常repoは別件STADIOのdirty branchのためソースをコピーせず保持。
確認用exeは通常repoの対象アプリdistへ固定配置し、旧公開1.0.1を退避、ZIP展開exeとhash照合する。
build-info.txtへsource commit、build日時、exe/ZIP hash、実ビルド場所を記録する。
この未merge worktreeと400件fixtureは再確認のため保持する。移動容量は削減に数えない。
旧Documents側のoverview-rename cloneは別アプリLookHereの未commit変更を持ち、overview-reload-v1.0.1はそのGit管理に紐づくため今回削除しない。他案件フォルダも触らない。
