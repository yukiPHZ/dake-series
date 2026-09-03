# ORIGINAL.md

## 正本宣言

このファイルは `DAKE_PDF_OverviewRename` の真の正本です。

README.md、DAKE_META、release_body.md、booth_product.txt、Store表示、実装指示書は、このファイルから派生するビューです。矛盾がある場合は本ファイルを優先します。

## 基本情報

- app_id: `dake_pdf_overview_rename`
- title: `DakePDF俯瞰名前変更`
- short_title: `PDF俯瞰名前変更`
- category: `PDF`
- status: `available`
- version: `1.0.1`
- app_type: `market`
- completion_goal: `formal_release`
- price: `500円`
- distribution: GitHub Release / BOOTH / dakeapp.com / Store
- target_platform: Windows
- folder_name: `DAKE_PDF_OverviewRename`
- exe_name: `DakePDF_OverviewRename.exe`
- booth_url: https://peakheadz.booth.pm/items/8798555
- store_url: https://store.dakeapp.com/product/?id=dake_pdf_overview_rename
- stripe_payment_link: https://buy.stripe.com/28E28r8jTfIHc4M6zD0gw0R
- payment_status: stripe_ready
## 目的

フォルダ内のPDFを1件ずつ開かなくても、1ページ目のサムネイルを俯瞰しながら、その場でPDFごとのファイル名を変更できるようにする。

内容確認のためだけに繰り返している次の往復をなくす。

```text
PDFを開く
↓
何の書類か確認する
↓
閉じる
↓
エクスプローラーへ戻る
↓
名前を変更する
```

本アプリでは次の最短経路に変える。

```text
フォルダを選ぶ
↓
PDFをサムネイルで俯瞰する
↓
見ながら名前を入力する
↓
変更分だけ一括反映する
```

## 対象ユーザー

- 不動産実務で大量のPDF資料を整理する人
- スキャンやダウンロード後の連番ファイル名を整理する人
- 契約書、重要事項説明書、登記、管理規約、図面などを扱う人
- PDFを開く回数そのものを減らしたい人

## 解決する困りごと

- 中身の分からないPDF名が大量に並ぶ
- 名前を付けるためだけにPDFビューアを何度も開閉する
- PDFビューアとエクスプローラーを往復する
- 変更対象を取り違える
- 一括リネーム途中の衝突や失敗でファイル名が崩れる

## 既存アプリとの役割分離

`DAKE_PDF_Rename` は、複数PDFのファイル名の前後へ同じ任意テキストを一括追加するアプリです。

本アプリは、PDFの内容をサムネイルで確認し、PDFごとに異なる意味のある名前を付けるアプリです。

目的と操作が異なるため、統合・置換しません。

## v1の主機能

- 選択フォルダ直下のPDF一覧
- 1ページ目サムネイル
- サムネイルカードによる俯瞰表示
- 各PDFの新しいファイル名入力
- `.pdf` 拡張子の固定表示
- 変更待ちカードの薄青表示
- 変更待ち件数の表示
- 表示サイズ `小 / 標準 / 大`
- サムネイルクリックによる1ページ目の大プレビュー
- 変更分だけの一括安全リネーム
- 直前の一括変更を1回だけ元に戻す

## 使い方の要点

1. `フォルダを選ぶ` を押す。
2. フォルダ直下のPDFカードが先に表示され、サムネイルが順次入る。
3. サムネイルを見ながらカード下の名前欄を編集する。
4. `名前変更を反映 n` を押す。
5. 事前検証と確認後に、変更待ちのPDFだけ名前を変更する。

## UI絶対条件

### ヘッダー

- 機能タイトルと機能説明を横並びにする。
- 機能タイトル: `PDFを見ながら名前を変える`
- 機能説明: `フォルダ内のPDFをサムネイルで一覧表示し、その場で名前を変更します。`
- ヘッダーに `シンプルそれDAKEシリーズ` を表示しない。
- タイトルと説明を意図なく縦積みにしない。
- ブランド用の帯、ロゴ、余計な上段を追加しない。
- Windows表示倍率100%、125%、150%で重なり・折返し・欠けを確認する。

### ブランド表記

`シンプルそれDAKEシリーズ` はフッターにのみ表示します。

### ツールバー

- `フォルダを選ぶ`
- `リフレッシュ`（現在のフォルダ選択を解除し、起動直後の状態へ戻す）
- `再読み込み`（現在選択中の同一フォルダを再スキャンする）
- 選択フォルダのパス
- `表示サイズ 小 / 標準 / 大`
- `変更を元に戻す`
- `名前変更を反映 n`

主ボタンは `名前変更を反映 n` とし、変更待ちが0件の場合は無効にします。

### カード

各カードには最低限、次を表示します。

- 1ページ目サムネイル
- ページ数
- 元ファイル名
- 新ファイル名入力欄
- 編集できない `.pdf`

長いファイル名でカード幅を広げません。入力欄は横スクロールを許可します。

変更待ちカードだけ背景または枠を薄青にし、未変更カードは静かな白背景とします。

### ステータス

同時に必要な情報を1行へ整理します。

例:

```text
48件 ｜ サムネイル 18 / 48 ｜ 3件の変更待ち
```

入力開始によってサムネイル進捗が消えないようにします。処理中、完了、エラーは明確に分けます。

### フッター

必須表示:

- `シンプルそれDAKEシリーズ / 止まらない、迷わない、すぐ終わる。`
- `戸建買取査定` リンク
- `Instagram` リンク
- `© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta`

広幅時は左右2ブロック、狭幅時は中央寄せ2段構成を許可します。思想タイトルとキャッチコピーは途中改行しません。

### 共通アイコン

DAKE共通アイコン `02_assets/dake_icon.ico` を使用します。

- 開発時のウインドウアイコン
- exeアイコン
- Windowsタスクバー
- PyInstallerの `--icon`
- onefile実行時の `--add-data`

で同じアイコンを確認します。

## 一覧・表示動作

- 対象は選択フォルダ直下の `.pdf` のみです。
- サブフォルダは再帰しません。
- 初期順序はファイル名の昇順です。
- PDFカード枠とファイル名を先に表示し、全サムネイル完成を待ちません。
- 見えている範囲を優先してサムネイルを生成します。
- 1ファイルのプレビュー失敗で一覧全体を失敗にしません。
- 暗号化PDF、破損PDF、0ページPDFなどは、そのカードだけ `プレビューできません` とし、可能なら名前入力は維持します。
- 表示サイズを切り替えても、入力中の名前、変更待ち状態、スクロール位置を不用意に失いません。
- 表示サイズ切替だけで、全PDFをPDFiumから再レンダリングし直しません。キャッシュ画像から派生するか、旧世代の未開始ジョブを破棄します。
- リフレッシュでは同じフォルダを再読込せず、破棄確認後に古いジョブ、カード、Undo履歴、大プレビューを破棄してフォルダ未選択の初期状態へ戻します。
- 再読み込みはフォルダ選択中だけ有効にし、フォルダ選択ダイアログを開かず、現在選択中の同一フォルダをファイル名昇順で再スキャンします。
- 再読み込みでは現在のカード、入力中の名前、変更待ち状態、Undo履歴、大プレビュー、古いscan/render/preview jobと旧generation結果を破棄します。選択フォルダと表示サイズは維持し、スクロール位置は先頭へ戻します。
- 再読み込みを連打しても未開始scan/render jobを蓄積せず、名前変更またはUndoの処理中は再読み込みできません。

## 大プレビュー

- サムネイルクリックで1ページ目を大きく表示します。
- 外部PDFビューアを主経路にしません。
- 新しいカードを選択した後に、古いプレビュー結果を表示しません。
- 大プレビューには Latest Job / generation 管理を適用できます。
- v1では全文閲覧・ページ送り・本文検索を実装しません。

## 名前入力

- 入力欄には拡張子を含まないファイル名本体だけを入れます。
- `.pdf` は固定表示し、編集対象にしません。
- Tabで次の入力欄へ自然に移動できるようにします。
- 入力しただけでは実ファイルを変更しません。
- 末尾へ `.pdf` を入力した場合は、二重拡張子にせず、入力不要であることを分かる形で扱います。
- 入力値の末尾空白や末尾ピリオドを黙って削除して成功扱いにしません。

## 未反映変更の保護

変更待ちが1件以上ある状態で次を行う場合は、破棄確認を出します。

- 別フォルダを選ぶ
- リフレッシュする
- 再読み込みする
- アプリを閉じる

表示サイズの変更では確認を出さず、入力を保持します。

## 名前変更安全仕様

### 反映前

変更対象全件を先に検証し、1件でも問題があればファイル操作を開始しません。

検証対象:

- 空欄
- Windows禁止文字 `< > : " / \\ | ? *`
- 制御文字
- 末尾ピリオド
- 末尾空白
- Windows予約名
- `CON.txt` や `LPT1.memo` のように予約名へ拡張相当文字列を付けた名前
- 大文字小文字を無視した重複
- 変更対象外ファイルとの衝突
- ファイル名要素の長さ
- 元ファイルの消失
- 読み込み後に元ファイルが外部で差し替え・変更されていないか

大文字小文字だけの変更は、衝突として誤拒否せず、安全に処理します。

### 反映中

- 一時名を挟む2段階リネームを使います。
- `A.pdf → B.pdf`、`B.pdf → A.pdf` の入れ替えに対応します。
- 3件以上の循環変更にも対応します。
- 同名ファイルを上書きしません。
- 処理途中で失敗した場合は、可能な範囲で全件を元名へロールバックします。
- 一時名は対象フォルダ内で衝突しないことを保証します。
- PDFレンダラーが対象ファイルを開いている状態でリネームを開始しません。未開始の古いレンダージョブを止め、進行中のファイルハンドルが閉じたことを確認してから処理します。
- リネームとUndoはUIスレッドで長時間実行しません。
- 書き込み開始後は中途半端なキャンセルを許可せず、完了またはロールバックまで進めます。

### 反映後

- カードのパスと元ファイル名を差分更新します。
- サムネイルを全部作り直しません。
- その場のカード配置を不用意に並び替えません。再度フォルダを選択した場合は、初期順序の規則で一覧を作り直します。
- 直前の一括変更だけUndo可能にします。

## Undo仕様

- 直前の成功した一括変更だけ1回戻せます。
- Undo前にも、元の名前へ戻せるか全件を事前検証します。
- 外部で同名ファイルが作られていた場合は上書きせず、Undoを開始しません。
- Undoも2段階リネームとロールバックを使用します。
- Undo成功後はUndo履歴を消します。
- 新しい一括変更を行った場合は、以前のUndo履歴を置き換えます。
- フォルダ変更、リフレッシュまたは再読み込みでUndo履歴を破棄します。

## 応答性・ジョブ管理

DAKE原則:

```text
処理は止まってもいい
UIは止めるな
応答なしは敗北
```

実装条件:

- 起動直後に重いPDFライブラリ処理を同期実行しません。
- フォルダ読み込み後、ユーザーへ0.1〜0.2秒程度で状態変化を返します。
- PDFレンダリングをUIスレッドで行いません。
- workerからTk WidgetやStringVarを直接操作しません。
- worker → queue → after → UI thread の責務を守ります。
- worker数を入力件数に比例させません。
- フォルダ変更・リフレッシュ・再読み込み・表示サイズ変更で古くなった未開始ジョブを蓄積しません。
- 古いgenerationの結果をUIへ反映しません。
- 終了時にworker、executor、pending job、after、timerを停止します。
- `bind_all` は使わず、このアプリのroot配下だけでマウスホイールを一覧Canvasへルーティングします。Canvas背景、カード、サムネイル、元ファイル名、名前入力欄の上で縦スクロールでき、大プレビューでは本一覧を動かしません。

合理的なv1基準:

- 48件: 実利用の中心。画面・入力・反映まで快適であること。
- 100件: 通常利用の上限目安。UIが固まらず作業できること。
- 300件: ストレス試験。時間がかかっても暴走・無限増加・操作不能を起こさないこと。
- 500件対応はv1の出荷条件にしません。

## 設定・ログ・保存方針

- 設定画面: 作らない
- 設定保存: v1では原則なし
- 通常ログ: 作らない
- Undo履歴: メモリ上の直前1回のみ
- PDF内容: 変更しない
- ネットワーク通信: しない
- 一時ファイル: リネーム処理中だけ対象フォルダ内に作り、正常終了またはロールバック後に残さない

永続トランザクションジャーナルはv1では実装しません。処理開始前検証、短時間の2段階リネーム、処理内ロールバックを採用します。

## 対応形式・非対応形式

- 対応入力: 選択フォルダ直下のPDF
- 対応出力: 同じPDFの新しいファイル名
- PDF本文: 変更しない
- サブフォルダ: 非対応
- PDF以外: 一覧・変更対象外

## v1でやらないこと

- AI自動命名
- OCR自動命名
- PDF本文検索
- ページ送りを含むPDFビューア化
- フォルダ自動分類
- サブフォルダ再帰
- PDF編集
- PDF結合・分割
- ファイル移動
- クラウド同期
- OpenAI API
- タグ管理
- 複雑な一括命名ルール
- 永続トランザクションジャーナル

## 技術方針

- Python
- Tkinter
- Pillow
- pypdfium2 / PDFiumを第一候補とする
- 固定worker、queue、after、generation管理
- PyInstaller onefile / noconsole
- インターネット接続不要

PDFレンダリングライブラリの採用・配布時は、現行ライセンスとPyInstaller同梱方法を実装時に確認します。必要な第三者ライセンス文書は正式配布物へ含めます。

## UI_TEXT

- 日本語UI文言はすべて `UI_TEXT` で一元管理します。
- `text=""` へ日本語を直接書きません。
- ボタン、見出し、説明、状態、ダイアログ、エラー、フッターを含みます。
- `APP_NAME`、`WINDOW_TITLE`、`COPYRIGHT` をファイル上部へまとめます。
- PythonファイルはUTF-8で扱います。

## 共通デザイン

- ベース背景: `#F6F7F9`
- カード背景: `#FFFFFF`
- 本文色: `#1E2430`
- 補助文字色: `#667085`
- 境界線: `#E6EAF0`
- アクセント色: `#2F6FED`
- アクセントホバー: `#2458BF`
- 変更待ち背景: `#EAF2FF`
- 変更待ち枠: `#7AA7FF`
- 完了色: `#12B76A`
- フォント: BIZ UDPGothic優先、Yu Gothic UI / Meiryoフォールバック

## テスト・完成判定

### 自動テスト

最低限、純粋な名前変更ロジックをUIから分離してテストします。

- 通常の複数変更
- 2件入れ替え
- 3件循環変更
- 大文字小文字だけの変更
- 日本語・全角・括弧を含む名前
- 空欄・禁止文字・制御文字
- 末尾ピリオド・末尾空白
- 予約名 `CON`、`CON.txt`、`LPT1.memo`
- 大文字小文字を無視した重複
- 変更対象外ファイルとの衝突
- 長すぎる名前
- 元ファイル消失
- 読み込み後の外部変更
- 処理途中の失敗を模擬したロールバック
- Undo成功
- Undo衝突時の開始前中止

### UI・操作確認

- ヘッダーがタイトル＋説明の横並び
- ヘッダーにブランド表記なし
- ブランド表記はフッターのみ
- フッターリンクとコピーライト
- 共通アイコンのウインドウ・exe・タスクバー表示
- 変更カードだけ薄青
- Tab移動
- 0件時の主ボタン無効
- 変更破棄確認
- サイズ切替で入力とスクロールを保持
- サムネイル失敗をカード単位で分離
- 100%、125%、150%表示倍率
- 初期サイズと最小サイズ

### データ件数

Codex環境では、機密情報を含まない合成PDFで1件、48件、100件、300件を確認します。

菊田の実フォルダ48件による最終確認は、人間の実機受入試験として別に行います。実フォルダへアクセスできない環境で「実データ確認済み」と報告しません。

### Windows build

Windows環境で可能な場合:

1. `build.bat` 実行
2. `dist/DakePDF_OverviewRename.exe` の生成確認
3. exe起動確認
4. 文字化け確認
5. ウインドウアイコン確認
6. タスクバーアイコン確認
7. 基本操作smoke test

実行できない項目は、未確認として明記します。

## 公開用説明の元情報

### 概要

PDFを開いて、閉じて、名前を変える。その往復をなくします。

フォルダ内のPDFを1ページ目のサムネイルで俯瞰しながら、PDFごとに新しい名前を入力し、変更分だけまとめて反映するWindows向けアプリです。

### 短い説明

PDFを開かず、見ながら名前を変える。

## DAKE_META生成用情報

```json
{
  "app_key": "dake_pdf_overview_rename",
  "display_name": "DakePDF俯瞰名前変更",
  "launcher_title": "PDF俯瞰名前変更",
  "launcher_description": "PDFを開かず、サムネイルを見ながら名前を変更します。",
  "site_title": "DakePDF俯瞰名前変更",
  "site_description": "フォルダ内のPDFをサムネイルで俯瞰しながら、PDFごとに名前を変更できるWindows向けアプリです。",
  "update_summary": "v1.0.1。同一フォルダを再スキャンする再読み込み機能を追加。",
  "folder_name": "DAKE_PDF_OverviewRename",
  "exe_name": "DakePDF_OverviewRename.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_PDF_OverviewRename_v1.0.0",
  "app_type": "market",
  "completion_goal": "formal_release",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## booth_product生成用情報

- 商品名: DakePDF俯瞰名前変更
- 価格案: 500円
- 商品紹介文: PDFを開いて、閉じて、名前を変える。その往復をなくします。フォルダ内のPDFを1ページ目のサムネイルで俯瞰しながら、変更分だけを安全にまとめて反映するWindows向けアプリです。
- タグ: PDF
Windows
実務
仕事効率化
軽量
シンプル
- 商品画像: assets/booth_thumbnail.jpg
- 補助画像: assets/screenshot.jpg
- 作品ファイル: booth_ready/DakePDF_OverviewRename.zip
- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_PDF_OverviewRename_v1.0.0
- BOOTH URL: https://peakheadz.booth.pm/items/8798555

## Store表示用情報

- 商品名: DakePDF俯瞰名前変更
- キャッチ: PDFを開かず、1ページ目を見ながら名前を変えます。
- 説明: フォルダ内のPDFを1ページ目のサムネイルで俯瞰しながら、PDFごとに新しい名前を入力し、変更分だけまとめて反映するWindows向けアプリです。
- 価格: 500円
- 画像: assets/booth_thumbnail.jpg
- ダウンロード導線: https://peakheadz.booth.pm/items/8798555
- サポート方針: GitHub ReleaseとBOOTHで配布します。
- Stripe Payment Link: https://buy.stripe.com/28E28r8jTfIHc4M6zD0gw0R
- Store販売状態: stripe_ready
- Store URL: https://store.dakeapp.com/product/?id=dake_pdf_overview_rename

## 価格・販売方針

- 正式販売価格: 500円
- BOOTH、Store、決済導線の表示価格を500円へ統一します。
- Store販売: Stripe Payment Linkで販売（BOOTH導線も併記）

## 配布・ダウンロード方針

GitHub Release、BOOTH、dakeapp.com、Storeを正式配布先とし、DAKE正式出荷ラインに従います。

- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_PDF_OverviewRename_v1.0.0
- BOOTH URL: https://peakheadz.booth.pm/items/8798555
- BOOTH配布zip: booth_ready/DakePDF_OverviewRename.zip
- Store URL: https://store.dakeapp.com/product/?id=dake_pdf_overview_rename
- Store決済導線: https://buy.stripe.com/28E28r8jTfIHc4M6zD0gw0R
- Store補助導線: https://peakheadz.booth.pm/items/8798555

## 免責・注意事項

- Windows向けアプリです。
- ファイル名を直接変更します。
- 大切なファイルは事前にバックアップを推奨します。
- 処理前に変更内容を確認してください。
- ご利用は自己責任でお願いいたします。
- 本ソフトウェアの無断転載・再配布を禁止します。

## 同梱ファイル方針

正式配布zip:

- `DakePDF_OverviewRename.exe`
- `README.txt`
- `注意事項.txt`
- 必要な第三者ライセンス文書

build、dist、spec、設定ファイル、個人データ、ソース一式を配布zipへ混ぜません。

## 現在地

- 正本仕様: 確定。Phase 2でも本ファイルを真の正本とする
- Phase 1技術受入: PASS。名前変更、Undo、wheel routing、リフレッシュ初期化、PDFium排他制御を維持
- Phase 1人間受入: Issue #17の開始条件として、名前変更、Undo、マウスホイール、リフレッシュを含む主要操作はPASS済み
- Phase 2自動テスト: 2026-09-02に41件成功。名前変更、Undo、wheel、リフレッシュ、PDFium mutex、出荷派生ビュー、公開画像を含む
- Phase 2合成PDF試験: 2026-09-02に1件、48件、100件、300件でPDFium描画、名前変更、Undo、内容ハッシュ維持、一時ファイル残留なし、worker停止を再確認
- Windows build: Windows 11 / Python 3.12.4 / PyInstaller 6.19.0でonefileクリーンビルド成功。`dist/DakePDF_OverviewRename.exe` の起動、48件読込、ホイール移動、リフレッシュ破棄確認、exeアイコン資源、ウインドウアイコンを確認
- 表示倍率: 100%、125%、150%を1180px / 900px幅で自動確認し、ヘッダー横並び、ツールバー、フッターの収まりは全条件PASS
- タスクバーアイコン: 直接目視は未確認。確認済みとは扱わない
- 公開画像: 一時作成した `C:\Users\Public\Documents\DAKE_synthetic_release_20260902_48` の無機密合成PDF 48件だけを使い、実アプリ画面から `assets/screenshot.webp` / `assets/screenshot.jpg` / `assets/booth_thumbnail.jpg` を作成。一時フォルダはキャプチャ後に削除済み
- 第三者ライセンス: ビルド環境の pypdfium2 5.13.0 / PDFium 153.0.7999.0（origin: pdfium-binaries）のwheelに記録されたLicense-File全19件を原文のまま `third_party_licenses/pypdfium2-5.13.0/` と配布物へ収録。コピー元とのSHA-256集合一致を確認
- 派生ビュー: `README.md`、`DAKE_META`、`release_body.md`、`booth_product.txt`、`booth_ready/` を本正本から整備。価格、GitHub Release URL、BOOTH URLを正式出荷値へ統一
- 次期パッチ版: `version: 1.0.1`、`price: 500円`、`status: available`。コードと配布ビューを準備し、公開はPR受入後の別工程
- 公開対象: `show_in_launcher: true`、`show_on_site: true`
- Phase 3正式出荷: Issue #19に従い、実在する公開URLを作成後に本正本へ順次記録する
- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_PDF_OverviewRename_v1.0.0（v1.0.0、公開・HTTP 200・配布zip添付を確認）
- BOOTH: https://peakheadz.booth.pm/items/8798555（500円、購入可能、商品画像2点、配布zip設定を確認）
- Stripe: Product `prod_VBbAdnHXBZVPjN`、Price `price_1UBDyVHrsJubFuDOfPQGhhSM`、Payment Link `plink_1UBE1sHrsJubFuDOhzt7u8BQ` を本番モードで作成。公開Checkoutで商品名、500円、購入ボタン、メール導線を確認し、実購入は未実施
- dakeapp.com: https://dakeapp.com/apps/pdf-overview-rename/（商品名、説明、画像、GitHub Release、BOOTH、Store導線を本番表示で確認）
- Store: https://store.dakeapp.com/product/?id=dake_pdf_overview_rename（500円、`Stripe対応`、Stripe購入導線、BOOTH補助導線、商品画像を本番表示で確認）
- Cloudflare Pages: dakeapp-site / dake-store-site ともmainマージ後の本番デプロイ成功、および上記カスタムドメインの公開表示を確認
- v1.0.1開発: 同一フォルダを再スキャンする `再読み込み` を追加。現行v1.0.0のGitHub Release、BOOTH、dakeapp.com、Store、Stripe、Cloudflareは本PRでは更新しない

## Codex作業時の注意

- 作業ブランチ: `codex/dake-pdf-overview-rename-formal-release`
- 対象: `01_apps/DAKE_PDF_OverviewRename/`
- 旧ブランチ `feature/dake-pdf-overview-rename` とDraft PR #11のコードは暫定参考であり、採用済み実装ではありません。
- 旧コードを自動でcherry-pickしません。
- 使える発想を参照してもよいですが、本正本と現行 `00_core` を優先してCodex自身で実装します。
- Issue #19の出荷順序と安全ゲートに従い、main、GitHub Release、BOOTH、dakeapp.com、Store、決済、Cloudflareを正式出荷します。
- GitHub Release、BOOTH、Stripe等のURLは作成前に推測せず、公開後の実URLだけを本正本へ書き戻します。
- 名前変更、Undo、PDFium排他、wheel、refreshの既存挙動を不要に変更しません。
