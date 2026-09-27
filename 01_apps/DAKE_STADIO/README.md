# DAKE STADIO 0.2.0

画像の画素とベクターパスを、一つの作品で編集するWindows日本語制作アプリです。バナー、チラシ、ロゴ、名刺を中心に、文字・画像・図形を配置し、保存後も編集を続けられます。専門的なAdobe製品の全機能と互換性を再現する版ではありません。

## 起動

`dist/DAKE_STADIO-0.2.0-win32-x64.zip` を展開し、`DAKE_STADIO-win32-x64/DAKE_STADIO.exe` を起動します。runtime、resources、locales、ライセンスを含むフォルダ全体を保持してください。exeだけのコピーでは動きません。利用者はNode.jsやnpmを別途入れる必要はありません。

Windows x64の未署名ローカル評価版です。配布版の確認結果は `TEST_REPORT.md`、機械的な証拠は `evidence/` を参照してください。別PC・SmartScreenの表示は未検証です。

## 最初の制作

1. 初期画面でバナー・チラシ・ロゴ・名刺を選び、名前・寸法・単位・dpiを指定します。開始後は案内が閉じ、中央のキャンバスを広く使えます。
2. PNG / JPEG / WebP / SVGを読み込むか、画面へファイルをドロップします。画像を選んで明るさなどを調整し、ブラシ・消しゴムで作品内の画素を編集できます。素材ファイル自体は書き換えません。
3. 左のツールで図形・文字・ペンのパスを追加します。文字を選ぶと右側の上部に書体、大きさ、太さ、スタイル、字間、行間、揃えが現れます。位置・幅・高さは作品のpx / mm、文字サイズはpxです。
4. 右のレイヤーをCtrl / Shiftで複数選択し、整列、グループ、順序変更、マスク、パス演算を使います。名前はダブルクリック、順序はドラッグで変更できます。
5. 「保存」で編集可能な `.dake` を残します。あとで「開く」または `.dake` のドロップで再編集できます。「書き出し」は完成画像等の出力です。

ホイールで表示移動、Ctrl＋ホイールでズーム。Tabで右パネルを隠し、境界のドラッグで幅を変えられます。表示の変更だけでは未保存になりません。ガイド・スナップ・塗り足し等の作品設定は .dake に残ります。

## フォント

文字設定のフォント一覧で、このPCの書体とGoogle Fontsを切り替えます。同梱カタログは1936書体、日本語対応67書体。検索・日本語絞り込み・太さ・スタイルを選んで「取得」を押すと、選んだフォントと許諾文を配布元からダウンロードします。この明示操作だけがフォント取得の通信を行い、作品・画像・入力文章を送信しません。

取得したGoogle Fontsはライセンス・同一フォントデータとともに .dake へ埋め込み、次回は通信せず使えます。Windowsのインストール済み書体は無条件に埋め込みません。他PCに同じ書体がない場合は表示が変わる可能性があります。使った書体で文字をアウトライン化する機能もあり、元の編集可能な文字は隠れたレイヤーとして残します。

フォントデータは単体base64 48 MiB、合計96 MiB、作品全体128 MiBまでです。上限を超える取得・適用は作品の変更前に拒否します。Google Fontsは書体ごとに許諾が異なります。埋込みフォントと許諾文を切り離さないでください。

## マスクとパス

画像等と最上面の図形・パスを複数選択してマスクを作ります。元の内容を保持し、有効・無効の切替と解除ができます。現在のマスクは図形・パスで見える範囲を決める方式で、筆で濃淡を描くマスクやぼかし境界は未対応です。

パスの結合・型抜き・交差・分割・複合化、ノードの追加・削除・滑らかさ・開閉、文字のアウトライン化、レイヤーの画像化を使えます。グループ内を直接編集する操作は未対応で、必要に応じて解除して編集します。

## 書き出しと印刷寸法

| 形式 | 内容と使い分け |
| --- | --- |
| .dake | 編集可能な作品。画像、文字、パス、レイヤー、取得フォント、配置設定を保存 |
| PNG / JPEG / WebP | 合成した画像。透明を使う場合はPNG / WebPを選択 |
| SVG | 対応する図形・パス・文字はベクター、画像は埋込み。生テキストのフォントデータは埋め込まないため、他環境で形を揃えるなら同書体の導入かアウトライン化が必要 |
| PDF | 指定dpiから実寸を計算した1ページのRGBラスターPDF。透明部分は白。文字・図形も画素化し、ガイドは出力しない |

印刷向けのキャンバスは塗り足しを含む全体です。名刺は300dpi・全体97×61mm、内側の仕上がり91×55mm、塗り足し3mm、セーフ4mm。チラシは全体216×303mm、仕上がりA4（210×297mm）です。整数画素への丸めにより、300dpiでは最大約0.043mmの寸法差があります。PDFはその整数画素数とdpiを正確に反映します。塗り足しとセーフは設定で調整できます。

PDFはベクターPDF、CMYK、ICC、PDF/Xではなく、トンボとTrimBox / BleedBoxも付与しません。商用印刷所の受理や実印刷は未検証です。入稿先の条件と照合してください。

## 保存の保護と制限

上書き前の作品は同じ場所の `.dake.bak` へ一世代残します。一時ファイルへの書込み完了後に本体を置換し、失敗時は未保存のまま保ちます。編集後の復旧候補は `%APPDATA%\DAKE STADIO\recovery.json` とその `.bak` に保存し、次回起動時に復旧を選べます。通常保存・正常終了後の不要な候補は消します。自動保存前の直近操作や物理的な電源断を必ず救えるという保証ではありません。

作品は各辺8192px・32MP、画像出力は各辺16384px・64MP、文書は最大2000オブジェクト・128 MiBまで。履歴は最大50状態で概算96 MiBに応じて整理し、最低1回の取り消しを残すため厳密なメモリ上限ではありません。開き直した作品へ前回の取り消し履歴は引き継ぎません。

選択範囲、筆圧、縦書き、文字ボックス内の混在書式、段落ごとの前後余白を含む高度な組版、複数アートボード、PSD / AI / PDFの編集読込み、RAW現像、CMYK / ICCは未対応です。32MPなどの大作品では同期描画と画像処理に負荷が残ります。実測と未検証事項を `TEST_REPORT.md` に記載しています。

## 作例と検証

配布の `examples/` にバナー・チラシ・ロゴ・名刺の編集可能な作品と書き出しを同梱する構成です。ソース側の検証済み作例は `artifacts/production/`、制作・再起動の記録は `evidence/production-results.json` にあります。作例の写真風素材は生成した抽象画像、人物・連絡先・会場は架空です。

実装済み、確認済み、未完了は `TEST_REPORT.md` で区別します。制作をすべて代替したとの主張ではありません。

## 開発

正規ディレクトリは `C:\Users\yukiz\devlop\DAKE_series\01_apps\DAKE_STADIO`。仕様の正本は `ORIGINAL.md`、日本語UI_TEXTは単一 `src/ui-text.json`（ネイティブ文言はdesktop配下）です。旧Documents配下は0.1.0保護用で逆同期しません。

Node.js 22.12以降 / npmの開発環境で実行します。依存の取得にはネットワークを使います。

```powershell
npm ci
npm start
npm test
npm run test:engine
npm run test:app
npm run test:restart
node scripts/run-production-check.cjs
npm run package
```

制作検査は実Google Fonts取得を含みます。試験用profile、TEMP、キャッシュ、成果物はアプリ内の `test-output/` を使用します。PDF独立検査にはpypdf / PillowとPopplerが必要で、QAスクリプトはCodex bundled runtimeを参照します。固定依存はElectron44.4.5、Fabric7.4.0、Paper0.12.18、opentype.js2.0.0、esbuild0.25.12です。

## ライセンスと公開状態

第三者告知は `assets/THIRD_PARTY_NOTICES.txt`、配布では `THIRD_PARTY_NOTICES.txt`。Electronの `LICENSE` / `LICENSES.chromium.html` も保持してください。一次資料は ORIGINAL.md に記載しています。GitHub Release・BOOTH・Storeへの公開、販売、課金はしていません。

© 2026 しまりす不動産 / Vibe-Coded by Yukihiko Kikuta

<!-- DAKE_META_START -->
```json
{
  "app_key": "dake_stadio",
  "display_name": "DAKE STADIO",
  "launcher_title": "DAKE STADIO",
  "launcher_description": "画像とベクターを一つの作品で編集する日本語制作アプリ。",
  "site_title": "DAKE STADIO",
  "site_description": "Windows向け・ローカルファーストの画像編集とベクター制作の統合アプリ。",
  "update_summary": "0.2.0 日本語文字、Google Fonts、マスク、パス演算、ガイド、実寸RGB PDFを追加。検証範囲は TEST_REPORT.md。",
  "folder_name": "DAKE_STADIO",
  "exe_name": "DAKE_STADIO.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.png",
  "status": "draft",
  "app_type": "market",
  "completion_goal": "local_ready",
  "show_in_launcher": false,
  "show_on_site": false
}
```
<!-- DAKE_META_END -->

<!-- RELEASE_BODY_START -->
DAKE STADIO 0.2.0は、画像とベクターを一つの作品で編集するWindows向けローカル評価版です。
バナー・チラシ・ロゴ・名刺の制作に向け、日本語の文字設定、Google Fontsの明示取得と作品への埋込み、マスク、パス集合演算、文字アウトライン、mm配置・ガイド、実寸RGB PDFを追加しました。
編集可能な作品は .dake へ保存し、PNG / JPEG / WebP / SVG / PDFへ書き出せます。ZIPを展開し、フォルダ内の DAKE_STADIO.exe を起動してください。Node.jsの別途導入は不要です。
検証した操作、性能、未検証事項は TEST_REPORT.md に記載しています。未署名のローカル評価用で、公開・販売は行っていません。CMYK / ICC / PDF/X、PSD / AI互換、実プリンターと別PCは未対応または未検証です。
<!-- RELEASE_BODY_END -->
