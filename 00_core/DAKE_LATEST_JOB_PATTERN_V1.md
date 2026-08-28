# DAKE Latest Job Pattern v1

## 位置づけ

この文書は、高頻度に更新され、新しい要求によって古い要求の結果が不要になる処理の設計パターンを定める。

`DAKE_INTERACTION_QUALITY_V1.md` は「何を守るか」を示す品質基準であり、本書はその中の最新要求優先原則を「どう判断し、どう実装するか」に分解した設計資料である。個別アプリの `ORIGINAL.md` が製品固有の正本であることは変わらない。

## 1. 基本思想

すべての要求を最後まで処理することを目的にしない。

新しい要求によって古くなった仕事は、可能な限り実行しない。重要なのは、今ユーザーが必要としている結果である。

```text
新しい要求が来る
↓
古い結果の価値がなくなる
↓
古い待機jobは実行しない
↓
古い完了結果は表示しない
```

これは処理結果を省略するための仕組みではなく、表示の鮮度を守りながら不要な仕事を減らす仕組みである。

## 2. 適用対象

次のように、新しい要求が来ると古い結果に価値がなくなる処理を候補とする。

- PDF preview
- ページrender
- zoom render
- resize render
- 画像preview
- 検索候補preview
- 入力に追従する検証結果や表示用変換

適用前に、次の条件を確認する。

- 結果の目的が、現在状態の表示またはpreviewである。
- 古い結果を表示しなくても、ユーザーのデータや履歴が失われない。
- 新しい要求が古い要求を完全に置き換える。
- request時点の状態をsnapshotとして表現できる。

## 3. 適用禁止

次の処理には原則として適用しない。

- 保存
- 送信
- 公開
- 決済
- ファイル書き込み
- 履歴記録
- ユーザーが明示的に実行した不可逆処理
- すべての要求を処理する必要があるbatch処理

古い仕事を捨ててよい責務と、捨ててはいけない責務を分離する。表示更新と保存処理が同じqueueやgenerationを共有している場合は、先に責務を分ける。

## 4. 基本構造

```text
UI request
↓
generation更新
↓
request時点の状態をsnapshot
↓
最新requestをpendingへ置換
↓
固定workerが処理
↓
結果をUI threadへ返す
↓
generation確認
↓
最新結果だけUIへ反映
```

pending requestは原則として最新1件だけ保持する。進行中jobがある間に複数の要求が来ても、待機するのは最後の要求だけでよい。

## 5. 実装原則

1. pending requestは原則最新1件とする。
2. worker数は必要最小限とする。順序より鮮度が重要な単一previewでは、固定1 workerを第一候補とする。
3. 古い待機jobは開始しない。
4. 進行中jobは安全に強制停止できなければ完了してよい。
5. 進行中jobの古い結果はUIへ反映しない。
6. request発行時点の状態をsnapshotし、workerが後から変化するTk変数やWidget状態を読まない。
7. workerからTk Widgetを直接操作しない。
8. 結果はqueue、`after`、または同等の仕組みでUI threadへ戻す。
9. 終了時にworker、pending request、polling timer、debounce timerを確実に停止する。
10. debounceは必要な場合だけ使用し、Latest Jobの代用と考えない。

## 6. generation ID

generation IDは、結果が現在も有効かを判定するために使う。

```text
request generation = 20

処理完了時:
current generation == 20
→ UI反映可

current generation == 24
→ 古い結果なので破棄
```

generationは要求を受け付けた時点で更新する。worker開始時だけ更新すると、待機中の古い結果が現在の結果と誤認される可能性がある。

結果には、最低でもgenerationと処理結果または例外情報を含める。PDF切替など別の入力単位をまたぐ場合は、generationに加えて入力識別子を検証してもよい。

## 7. request snapshot

workerへ渡すrequestには、表示結果を決める情報を明示的に含める。

候補例:

- 入力ファイルまたはdocument識別子
- ページ番号
- zoom
- rotation
- canvasサイズ
- 表示モード
- previewに必要な設定値
- generation

requestは可能な限り不変の値として扱う。workerが処理途中にUI状態を読み直す設計は、表示と保存状態のずれや再現しにくい競合を生むため避ける。

## 8. pending requestとworker

FIFO queueへすべてのrequestを積むだけでは、古い仕事が順番待ちとして残る。Latest Jobでは、次のどちらかを選ぶ。

- 共有するpending slotを最新requestで置換する。
- Queueを使う場合は、未処理の古いrequestを取り除いて最新1件だけ残す。

workerは進行中jobを完了した後、その時点の最新pending requestを取得する。新しいpendingがなければ待機する。

worker数を増やすと古いjobも並行してCPUとメモリを使うため、単一previewでは速さにつながらないことがある。複数workerは、結果が独立し、同時実行が実際に有効で、鮮度管理を複雑にしない場合だけ採用する。

## 9. UI threadへの結果返却

workerは、画像データ、座標、metadata、例外などTkに依存しない結果を返す。

UI thread側で次を行う。

- generationの最終確認
- `PhotoImage` などTk依存objectの生成
- CanvasやWidgetの更新
- 画像参照の保持
- statusやエラー表示

worker完了時にgenerationが一致していても、UI queueから取り出すまでに新しい要求が来る場合がある。UIへ反映する直前にもgenerationを確認する。

## 10. debounceとの違い

debounceは、短時間に連続した要求そのものをまとめる。

Latest Jobは、要求を受け付けたうえで、古くなった待機jobと結果を捨てる。

必要なら併用できる。

```text
resize
↓
140ms debounce
↓
Latest Job request
```

debounceだけでは、すでに開始した古いjobの結果競合を防げない。Latest Jobだけでは、resize eventごとの軽いrequest生成やtimer更新が不要になるとは限らない。どちらが必要かを計測して決める。

## 11. 差分更新との違い

差分更新は、変わった部分だけ更新する。

Latest Jobは、もう不要な仕事を実行しない。

同一アプリ内で併用できる。

```text
CheckStamp:
stamp移動
→ stamp overlayだけ差分更新

ページ変更
→ 最新ページの背景renderだけ実行
```

背景が変わらない操作でrender request自体を発行しない方が、Latest Jobで後から捨てるよりよい。まず不要な要求を作らず、それでも発生する高頻度の重い要求へLatest Jobを適用する。

## 12. 終了と入力切替

終了時は次を行う。

- 新規request受付を停止する。
- generationを進め、進行中結果を無効化する。
- pending requestを破棄する。
- workerへ停止を通知する。
- polling・debounce timerを解除する。
- 安全な範囲でworker終了を待つ。
- 終了後にUI更新を予約しない。

別PDFへの切替も小さなライフサイクル終了として扱う。旧PDFのrequestと結果を無効化し、documentを再利用する場合は所有workerとclose時点を明確にする。

## 13. 例外処理

例外結果もgenerationと関連付ける。古いjobの失敗を、現在表示中の新しい要求のエラーとして表示しない。

現在generationのjobが失敗した場合は、UI threadで既存のエラー表示方針に従う。workerを例外で終了させず、次の最新requestを処理できる状態へ戻す。

## 14. 計測と成功判定

高頻度操作を一定回数行い、次を確認する。

- UI操作回数
- request発行数
- 実際に開始した重いjob数
- 開始前に置換したjob数
- 古い完了結果の破棄数
- 古い結果のUI反映数
- 最大同時worker数
- 最終表示状態
- 終了後の残留workerとtimer

成功判定は、操作回数と重い処理回数を同じにすることではない。

```text
操作回数 != 重い処理回数
古い結果のUI反映数 = 0
最終表示 = 現在の内部状態
```

を確認する。

## 15. 実証アプリ

### DAKE_PDF_Viewer

ページ移動、zoom、resizeなど、PDF背景自体が変わる描画要求に適用した。固定1 render worker、最新1 pending request、generation照合により、連続操作後に古いページが遅れて表示されることを防いだ。

Viewer固有なのは、ページ・zoom・rotation・表示領域などのrequest payloadと、閲覧中の描画状態である。

### DAKE_PDF_CheckStamp

PDF背景のページ変更やpreview更新にLatest Jobを適用し、stamp位置変更はCanvas overlayの差分更新とした。背景が変わらない操作ではrender request自体を発行しない構造を組み合わせた。

CheckStamp固有なのは、stampの名字・日付・位置・サイズを保存状態と一致させる責務、および背景とstamp overlayの分離である。

### DAKE_PDF_Marker

PDF open、ページ変更、zoom、resize後の描画を固定1 workerへ渡し、最新pending requestとgenerationで鮮度を管理した。workerが表示中documentを所有・再利用し、Canvasの背景itemとmarker itemはUI側で保持して差分更新した。

Marker固有なのは、marker一覧のPDF座標を保存結果の正本とすること、ドラッグ中に同じmarker itemを動かすこと、documentの所有と再利用である。

### 3アプリで共通した責務

- 最新requestだけを待機対象として保持する。
- 固定1 workerで重い描画を直列化する。
- generationで結果の鮮度を確認する。
- 古い待機jobを開始せず、古い完了結果をUIへ反映しない。
- Tk更新をUI threadへ戻す。
- 終了時にworker、pending request、timerを停止する。

### アプリごとに異なる責務

- request payload
- render関数
- result形式と画像変換
- PDF documentの所有とclose lifecycle
- 差分更新するoverlayの種類
- debounce対象と時間
- 保存状態との整合条件
- エラー表示と再試行条件

## 16. module化判断

現時点では設計パターンだけを共通化し、共通moduleは作らない。

3アプリで基本原則は実証できたが、request payload、PDF lifecycle、result形式、debounce対象、shutdown責務に差がある。これらを汎用APIへ押し込むと、各アプリで短く読める処理がcallbackと抽象状態へ分散し、理解が難しくなる可能性がある。

今後さらに適用例を確認し、次をすべて満たす場合だけmodule化を再検討する。

- 同じworker lifecycleを自然に共有できる。
- generationとpending管理以外の分岐が少ない。
- result deliveryとshutdownのAPIを小さく保てる。
- PyInstallerや依存ライブラリ固有の処理を漏らさない。
- 共通moduleを使った方が各アプリのコードが短く、理解しやすくなる。

同じ思想と同じコードを混同しない。共通module化しない場合でも、本書を設計・レビュー・計測の共通判断基準として使用する。

## 17. レビュー用チェックリスト

- [ ] 新要求で古い結果が不要になる責務か。
- [ ] 保存・書き込み・履歴など、捨ててはいけない処理と分離されているか。
- [ ] pending requestは最新1件に制限されているか。
- [ ] worker数を必要最小限にしているか。
- [ ] request時点の状態をsnapshotしているか。
- [ ] workerがTk WidgetやTk変数を直接触っていないか。
- [ ] UI反映直前にgenerationを確認しているか。
- [ ] 古い例外を現在のエラーとして表示しないか。
- [ ] 背景が変わらない操作は差分更新で済ませているか。
- [ ] polling・debounce・workerを終了時に停止するか。
- [ ] 操作回数と実render数、古い結果破棄数を計測したか。
- [ ] 最終表示が内部状態と一致するか。
