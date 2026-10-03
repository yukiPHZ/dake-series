# DAKE STADIO 0.4.0-eval.1 検証記録

検証日: 2026-09-30。Windows x64、同一PC上のローカル評価版。正本は ORIGINAL.md、設計と再開手順は TECHNICAL_RECOVERY.md。0.3.0の旧記録は artifacts/technical-recovery-20260930 に保全し、今回の合格に流用していない。

## 到達点

フォントを選んで文字とベクターのロゴを保存 → .dakeを名刺へ編集可能な部品として配置 → 子タブの文字変更を親へライブ表示し反映 → 写真を非破壊補正し筆マスクで合成 → 保存して正常終了 → 別PIDでオフライン再起動 → 部品・画像を再編集 → PNG/JPEG/PDFを実保存、という動線を新規設定・既存設定の安全な全コピーで検証した。

各自動制作はアプリのポインタ/キー、実native IPC、エンジン状態/画素の検査を組み合わせたもの。位置のfixture設定、診断API、ダイアログ回答のテスト用置換も使う。人が全操作したという意味ではない。別途、Windowsの通常画面で文字入力・フォント見本/適用・作品保存・PNG出力のダイアログを操作した。

全機能・全素材・あらゆる端末の完成保証ではない。未検証と制約を末尾に記す。

## 配布の識別

- バージョン: 0.4.0-eval.1、Windows version 0.4.0.1、Electron 44.4.5。
- 最終実行コード ASAR SHA256: 81a7d817c20502d2d192f3b94243ba6255286265306c83b7a2d5047cf775d732
- EXE SHA256: 78923186630e97ca249cda2a26a4233129569230cf4188cc761b0fd4ba937fe0
- ZIPの最終SHAは同梱ではなくZIP横の .sha256 と evidence/evaluation-build.json に記録（自己参照を避ける）。BUILD_ID.json にpackaging入力ごとのhashを記録。
- ドキュメント更新後に再圧縮・再展開し、全ファイルのhashを照合する。結果は evidence/evaluation-archive-results.json。文書の更新だけでは実行コードを差し替えない。

## フォント不具合の再現と解消

未変更の0.3.0配布EXEで先に再現した。ASARは 990fb93832a58899fd8e6163367dcdc8d1afd3d276de620064ce25860a6d847d。配布と当時の主要ソースは一致し、「古いEXEを使っていた」と決めつける根拠はなかった。

| 段階 | 修正前に確認したこと | 今回の確認 |
|---|---|---|
| 一覧・取得 | Windows 301 face列挙、元9種Googleキャッシュの存在。すべてが取得不能とは再現しなかった | Windowsバイナリ読込とGoogle公開書体の実取得。破損/登録失敗は適用成功にしない |
| 見本 | 見本文字変更で単一DOM要素をfor-ofするTypeError。未取得Googleは既定書体の文字を見本のように表示 | 読込済み実書体だけ見本表示。編集中の文字の変更に追従。未取得は未取得と表示 |
| 新規文字 | Windows/Googleを使った後も新規文字がYu Gothicへ戻る | 書体descriptorをライブラリと新規文字書式で保持、既存/新規/別作品へ適用 |
| 保存・再起動 | 古い合格報告を継承しない | 実保存→別プロセス→nativeフォントHTTPとrenderer通信を止めて再適用 |
| 字形 | 名前/成功通知だけでは不足 | Georgia Regular/Bold、Zen Maru Gothicなど各環境6描画を独立FontFaceバイナリ参照と画素比較。完全一致、汎用代替書体とは不一致 |

今回の最終EXE結果は evidence/font-reliability-packaged-results.json。元症状のうち再現できた経路を修正したのであり、原作者の過去の全不具合条件を同定したとはしていない。登録エラーを注入した検査は失敗時表示の確認で、現実のOS障害の再現とは別。

## 今回の検証と結果

| 対象 | 結果と証拠 |
|---|---|
| 最終EXEのフォント | 新規/全設定コピーとも成功。実字形、キャッシュ、新規文字、他作品、オフライン再起動、失敗表示、例外0。evidence/font-reliability-packaged-results.json |
| 最終EXEの指定制作動線 | 新規33/コピー34確認。保存前後PNG完全一致、4起動すべて正常終了・別PID。evidence/creation-flow-packaged-results.json |
| 最終EXEの色・出力 | 10確認。色プレビュー/履歴/登録/一回Undo/Esc取消、PDF半分/2倍の解像度、実測容量、作品寸法変更、旧v2→v3とbak。evidence/output-color-packaged-results.json |
| native単体 | 36テスト合格。fonts、storage、components、font identity、recovery等。test-output/native-unit-v4.log |
| 既存エンジン | 167確認、v1/v2、パス、図形演算、旧filter/clip、ブラシ、文字アウトライン、Undo、保存復元。evidence/engine-check-results.json |
| 画像の個別検証 | 27確認。原画不変、補正解除、旧filterと画素一致、筆隠す/戻す、フェザー、矩形選択、部分補正、複製、保存復元、破損マスク拒否。test-output/imaging-v4/report.json |
| 画像UI | 8確認。補正ダイアログ、取消、曲線ページ表示、実ポインタ筆/矩形、保存（実曲線変更は通し制作で検証）。test-output/imaging-ui-v4/report.json |
| 部品/再開 | 23 + 別プロセス9確認。編集可能性、子live、親Undo、リンク切れcache、再選択、元資産保護。test-output/workspace/run-1790761624765/results.json と test-output/workspace/restart-1790761511940/results.json |
| データ保護 | 6 + 7確認。通常元タブ保存と部品保存権限の区別、同一pathの二重タブ防止、未読復旧保護、現在タブを残した追加復旧。test-output/integrity-review-1790762286123/results.json、test-output/pending-recovery-1790762286091/results.json |
| 部品の高速操作 | 候補1の実行コードでfresh10/copied10成功。初期preview未表示のまま実クリックを受信し反映、最長約2.4秒、例外0。test-output/component-rapid-1790763325475/report.json |
| Windows実画面 | 候補1で新規ロゴ、Georgia見本と適用、通常保存ダイアログ、PNG出力、正常終了。evidence/manual-eval-v4-results.json。最後の変更は版表示とフォント一覧余白・通信文言のみ。最終コードでも起動表記、実Windows読込ダイアログ、Georgia字形復元、文字をDAKE Studio +へ再編集・Ctrl+S保存を確認 |

件数は検査した状態や境界の数で、機能の完成度の点数ではない。古いevidence/*の0.3.0記録は履歴資料であり、この表に明記した今回の根拠を使う。

## 保存結果と出力構造

通し制作の元logoと写真はhash不変。子文書は文字とベクターを保ったまま親へ保存される。補正recipeと2種類のmaskは原画と別。PNG/JPEGは1146×720 px、PDFは約97.03×60.96 mm（塗り足し込み）でRGB画像1枚の1ページ。例: PNG 353342 B、JPEG 62244 B、PDF 351122 B。UIの実測容量と実際に保存したbyte数の一致を確認。

Windows手操作のlogoはv3/Georgia/本文DAKE Studio、PNGは1000×1000 px・33263 B、UI表示32.5 KiBと一致。作品と出力は test-output/manual-eval/ に残す。通し制作の写真はWindows同梱画像を使った試験用であり、作品を配布ZIPには入れない。

## 保全と失敗の扱い

元APPDATAは削除・初期化していない。profile-seed-20260930へ104ファイル/42932380 Bを複製し、各検査はさらにそのコピーを使用。使用中Chromiumのleveldb/LOCKのみEBUSYで除外し、manifestに明記。既存アプリは未保存作品を開いていたため閉じず、旧配布フォルダ/ZIPも保持。seedと元フォントdescriptorの不変を照合した。使用中の元アプリ自身による設定更新まで禁止・不変保証したという意味ではない。

初期の検証ハーネスに非表示の文字欄をクリックしていた問題があり、本文を実確認する検査へ修正して再実行した。別の1回は子反映待ちでtimeout、当時のクリック受信/状態証拠が不足して原因未確定。通常再試験とpreviewを待たない20回の独立検査では再現しなかった。製品コードを修正して解消したと称さず、未再現として残す。

## 未達・適用限界

- 同一PCでの確認。全Windows/Google書体、全既存.dake、日本語IMEの各方式、別PC、高DPI・狭い画面、長時間・大量レイヤーは網羅していない。
- 2048×1536画像の今回測定: 筆6回19〜21 ms、Undo 3 ms、export約2005 ms。非表示ウィンドウの最大frame gap1007 ms。背景/待機の影響を切り分けておらず、前面UIの停止時間とも全操作の応答性合格値とも扱わない。
- Windows書体バイトは埋め込まない。別PCには同じ書体が必要。同一文書内の同名・同style・重なるweightの異なる既知SHAは拒否し、レイヤー別aliasでの混在は未対応。
- Undoは再起動をまたがない。リンクは自動更新せず、更新時は明示操作。再起動後は元ファイル再選択が必要。hardlink/junctionによる同一作品別名のUI重複防止は未検証。
- マスクは画像レイヤー、選択からの作成は矩形。自由形選択、自動抽出、筆圧は未対応。カーブは5点補間、色別RGBカーブではない。
- PDFはRGBラスター。CMYK/ICC、PDF/X、ベクター/文字保持、TrimBox/BleedBox、トンボは未対応。実印刷・入稿適性は未検証。SVGの文字は閲覧側フォントに依存。
- 物理電源断/ディスク故障の復旧を保証しない。未署名評価版。外部公開/push/販売/課金/作品送信/OS関連付けの変更はしていない。

次回は使用感に基づく再現条件を専用profileで追加し、性能・編集境界・出力互換の順に検証範囲を拡張する。手順は TECHNICAL_RECOVERY.md の再開節。
