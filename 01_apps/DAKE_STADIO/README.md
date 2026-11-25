# DAKE STADIO

画像とベクターを、一つの作品で編集する Windows 日本語アプリです。正式名称は **DAKE STADIO** です。

**0.1.0 ローカル評価版。** 画像の画素編集、図形・文字・パス制作、レイヤー、履歴、`.dake` 保存・再読込み、PNG / JPEG / WebP / SVG 書き出しを実装しています。検証済みと未確認の区別は `TEST_REPORT.md`、設計判断は `ORIGINAL.md` を参照してください。

## 起動

配布物 `dist/DAKE_STADIO-0.1.0-win32-x64.zip` を展開し、`DAKE_STADIO-win32-x64` 内の `DAKE_STADIO.exe` を起動します。開発領域からは `dist/DAKE_STADIO-win32-x64/DAKE_STADIO.exe` が起動対象です。ZIPから展開したEXEの起動、日本語UI、サンプル、画像worker、ローカル復旧保存、PNG書き出し、正常終了を同じWindows PCで検証しました。

runtime、resources、locales、ライセンスを含むフォルダ全体を保管してください。exe だけでは動きません。利用時の Node.js、npm、Adobe 製品の導入は不要です。作品編集にログインやクラウド接続は不要です。

未署名の評価版です。確認したのは開発に使用したWindows PCです。SmartScreen の表示と別 PC での起動は未検証です。

## 最初の制作

すぐに試す場合は **「サンプルで試す」** を押します。画像・編集可能な文字・曲線が一つの作品に入っています。サンプル画像はアプリ内で生成した抽象画像です。

1. **「新しく作る」** で名前、幅、高さ、背景を決め、**「作り始める」** を押します。
2. **「画像・SVGを追加」** で素材を読み込みます。画像レイヤーを選んで **「ブラシ」** または **「消しゴム」** で実画素を編集します。画像未選択でブラシを使うと透明なペイントレイヤーを作ります。
3. 画像選択時の **「画像の調整」** で明るさ・コントラスト・彩度・白黒を調整します。**「調整を戻す」** で調整前へ戻せます。
4. 四角形・楕円・三角形・直線・文字を追加します。文字は右のテキスト欄またはダブルクリックで編集します。色・線幅・不透明度・大きさも右で調整できます。
5. **「ペン」** はクリックで角、ドラッグで曲線を作ります。Enter または **「パスを確定」** で終了し、**「閉じて確定」** で閉じた図形にします。**「パス編集」** で端点と制御点を動かします。
6. **「選択」** で移動・変形。Shift で複数選択し、グループ化できます。レイヤー欄で表示、ロック、前後順を変えます。**「切り抜き」** では残す作品範囲をドラッグします。
7. **「保存」** で `.dake` を保存します。次回は **「作品を開く」** で続きを作ります。
8. **「書き出す」** で PNG / JPEG / WebP / SVG と解像度を選び、保存先を指定します。再編集用の `.dake` も残してください。

## よく使う操作

| 操作 | キー・マウス |
| --- | --- |
| 保存 / 別名保存 | Ctrl＋S / Ctrl＋Shift＋S |
| 開く / 新規 | Ctrl＋O / Ctrl＋N |
| 取り消し / やり直し | Ctrl＋Z / Ctrl＋Shift＋Z（Ctrl＋Y も可） |
| 複製 / 削除 | Ctrl＋D / Delete |
| グループ / 解除 | Ctrl＋G / Ctrl＋Shift＋G |
| 選択 / パス編集 / ペン | V / A / P |
| ブラシ / 消しゴム / 手のひら / 切り抜き | B / E / H / C |
| 四角形 / 楕円 / 文字 | R / O / T |
| パス確定 / 中止 | Enter / Esc |
| 1px / 10px 移動 | 矢印 / Shift＋矢印 |
| 表示移動 / 拡大縮小 | ホイール / Ctrl＋ホイール |
| 全体を表示 | Ctrl＋0 |

文字入力中は編集用ショートカットを抑制します。画像編集も取り消せます。履歴は最大50状態で、概算96 MiBを目安に古いものを減らし、最低1回の取り消しを残します。厳密なメモリ上限ではありません。作品を開き直すと新しい履歴を始めます。

## 保存と復旧

- 元画像を変更せず作品内のコピーを編集します。元画像への書き出し上書きも拒否します。
- `.dake` に画像・パス・文字・レイヤー・調整を保存します。画像は埋め込みなのでリンク切れになりません。
- 上書き保存は前回分を隣の `<作品名>.dake.bak` に保持します。戻す場合は `.bak` を別の `.dake` 名へコピーして開いてください。
- 未保存の変更は `%APPDATA%\DAKE STADIO\recovery.json` と `.bak` に自動保存します。異常終了後の起動時は **「復旧する」** を選び、別名保存してください。
- 通常保存成功時に古い復旧候補を消します。正常終了でも保存または未保存内容の破棄を確認して復旧候補を消します。
- 自動保存前の直近の操作は復旧できない場合があります。通常保存も併用してください。

破損ファイル・不正な外部参照は拒否して現在作品を保持します。保存は一時ファイルとバックアップを使い、失敗を成功表示にしません。確認した異常ケースは検証報告にあります。

PNG / JPEG / WebP は合成画像です。PNG / WebP は透明背景対応、JPEG は白で埋めます。SVG は対応図形・パス・文字・埋め込み画像を保持する交換形式です。全編集属性や履歴を保管する形式ではありません。SVG の文字は受け取り先のインストール済みフォントに依存します。

## 上限と制約

- 作品は各辺8192 px・3200万画素以内。画像書き出しは各辺16384 px・6400万画素以内。ファイルは128 MiB以内、文書読込みは2000オブジェクト以内。上限内でも大きい作品は負荷が増えます。
- RGB 制作向けです。CMYK / ICC、PSD / AI / PDF の編集読込み、RAW現像、Adobe固有機能との完全互換は未対応。
- 選択範囲、レイヤーマスク、パスの結合・型抜き、筆圧、縦書き、複数ページは未実装。
- 巨大画像・多数レイヤー・長時間利用、さまざまなGPU / DPI、別PCの文字表示は網羅検証していません。

日常制作を代替する統合基盤を実装した評価版です。Photoshop / Illustrator の専門用途をすべて代替できるという版ではありません。

## 開発・ビルド

Electron 44.4.5、Fabric.js 7.4.0、esbuild 0.25.12を固定。Windows x64 の Node.js / npm 環境で、このディレクトリを開きます。依存取得時はネットワークが必要です。

```powershell
npm ci
npm run build
npm start
```

保存・ネイティブ境界テストと配布物の作成:

```powershell
npm test
npm run package
```

`npm start` は起動前にビルドします。実Electronの編集・統合試験と手動確認は検証報告を参照してください。配布結果は `dist/` に生成します。

QAで利用者の復旧データを分離する場合:

```powershell
$env:STADIO_USER_DATA = Join-Path $PWD 'evidence\qa-user-data'
npm start
```

## ライセンス・公開状態

追加告知はソース側の `assets/THIRD_PARTY_NOTICES.txt`、配布フォルダの `THIRD_PARTY_NOTICES.txt` です。Electron runtime同梱の `LICENSE` / `LICENSES.chromium.html` も削除しないでください。依存の一次資料は ORIGINAL.md にあります。

ローカル評価用です。GitHub Release・BOOTH・Storeの正式公開や販売は行っていません。

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
  "update_summary": "0.1.0 ローカル評価版。実装・検証状況は TEST_REPORT.md を参照。",
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
DAKE STADIO 0.1.0は画像とベクターを同じ作品で編集するWindows向けローカル評価版です。
画像画素・パス・文字・レイヤー編集、.dake保存、画像・SVG書き出しを実装しています。
検証済み・未確認・未完了はTEST_REPORT.mdで区別して案内します。
<!-- RELEASE_BODY_END -->


