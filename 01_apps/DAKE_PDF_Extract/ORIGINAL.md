# ORIGINAL.md

## 正本宣言

このファイルは `DAKE_PDF_Extract` の真の正本です。

README.md、DAKE_META、release_body.md、booth_product.txt などの派生ファイルは、このファイルの内容から整合させます。

## 基本情報

- app_id: dake_pdf_extract
- title: DakePDF抽出
- short_title: PDF抽出
- category: PDF
- status: available
- version: 1.0.0
- price: 500円
- distribution: GitHub Release / BOOTH / dakeapp.com / store.dakeapp.com
- target_platform: Windows
- payment_status: booth_only
- booth_url: https://peakheadz.booth.pm/items/8746426
- store_url: https://store.dakeapp.com/product/?id=dake_pdf_extract
- stripe_payment_link: 未設定
- exe_name: DakePDF_Extract.exe
- folder_name: DAKE_PDF_Extract
- app_user_model_id: ShimarisuFudosan.DAKE.PDFExtract
- release_tag: DAKE_PDF_Extract_v1.0.0

## 目的

1つのPDFを開いたまま、必要なページを何度も連続して抽出するための単機能アプリです。

中心になる流れは、ページを選ぶ、抽出して保存する、選択が自動解除される、次のページを選ぶ、また抽出する、というループです。

## 想定利用

1. PDFを1つ追加します。
2. サムネイルを見ながら1ページまたは複数ページを選択します。
3. `抽出して保存` または Enter で保存します。
4. 保存後も元PDF、スクロール位置、保存先、抽出方法を維持します。
5. 抽出したページだけ自動で選択解除され、抽出済み表示になります。
6. そのまま次のページを選んで連続抽出します。

## 主要機能

- `PDFを追加` または空状態中央の選択領域からPDFを1ファイル読み込み
- ドラッグ＆ドロップでPDFを1ファイル読み込み
- 全ページのサムネイル一覧表示
- ファイル名、総ページ数、保存先表示
- 読み込み済みPDFをリフレッシュして最新状態へ再読み込み
- 通常クリックで選択 / 再クリックで選択解除
- Shift+クリックで最後にクリックしたページから範囲選択
- Enterで現在の抽出方法により保存
- Escで選択解除
- Ctrl+マウスホイール、－ボタン、フラットなスライダー、＋ボタンでサムネイルサイズ変更
- 選択ページを1つのPDFにまとめて保存
- 選択ページを1ページずつ保存
- 保存成功後も元PDFを閉じない
- 保存成功後もスクロール位置を戻さない
- 保存成功後に選択を自動解除
- 抽出済みページを少し薄く表示し、小さなチェックを表示
- 抽出済みページも再選択可能
- 保存成功後に保存先フォルダを表示。既存の同一フォルダwindowがある場合は前面化

## 抽出後の挙動

保存成功後は、読み込んだPDF、現在のスクロール位置、保存先、抽出方法を維持します。

完了ダイアログは表示しません。保存成功後は画面内ステータスで通知し、保存先フォルダを表示します。同じ保存先フォルダが既にExplorerで開いている場合は新しいwindowを増やさず、既存windowを前面化します。フォルダ表示に失敗しても抽出結果は維持します。

例外として、ファイル選択と保存先選択にはOS標準ダイアログを使用します。保存エラーや読み込みエラーは、連続作業を止めにくいように画面内ステータスへ具体的に表示します。

## 保存先

初期保存先は、読み込んだ元PDFと同じフォルダです。

ユーザーが保存先を変更した場合は手動指定を優先し、PDFの再読み込み、抽出後、同じ起動中に別PDFを開いた場合もその保存先を維持します。

## 出力命名規則

元PDF名と抽出ページ番号から確実に生成します。

- 1ページ: `元PDF名_p001.pdf`
- 連続ページをまとめて抽出: `元PDF名_p002-004.pdf`
- 非連続ページをまとめて抽出: `元PDF名_p002_004_007.pdf`
- 1ページずつ抽出: `元PDF名_p002.pdf`、`元PDF名_p004.pdf`、`元PDF名_p007.pdf`

同名ファイルが存在する場合は上書きせず、`_2`、`_3` のように連番を付けます。

ページ番号の桁数は総ページ数に応じて最低3桁です。

## 共通アイコン

DAKEシリーズ共通アイコン `../../02_assets/dake_icon.ico` を使用します。

アプリ固有アイコンは作成しません。

ソース実行時は `__file__` 基準で共通アイコンを探します。PyInstaller版では同梱アイコンとDAKE_series配下の共通アイコンを候補にします。見つからない場合も起動不能にはしません。

## ヘッダー

通常幅では、機能タイトル `PDFからページを抽出する` と機能説明を同一行に表示し、その右側へ `PDFを追加`、`リフレッシュ`、`保存先を変更` を配置します。

シリーズ名 `シンプルそれDAKEシリーズ` はヘッダーへ表示せず、フッターだけで使用します。

## フッター

フッター文言は以下で固定します。

```text
シンプルそれDAKEシリーズ
止まらない、迷わない、すぐ終わる。
戸建買取査定
Instagram
© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta
```

区切り文字は ` ｜ ` です。

幅に余裕がある場合は左右2ブロック、幅が足りない場合は中央寄せ2段に切り替えます。

`戸建買取査定` と `Instagram` だけをリンクにし、リンクURLは既存DAKEアプリと同じものを使用します。

## 正式出荷状況

初回正式バージョンは `1.0.0`、tagは `DAKE_PDF_Extract_v1.0.0` です。

- 正式価格: 500円
- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_PDF_Extract_v1.0.0
- BOOTH登録: 500円で正式公開済み
- Store公開: https://store.dakeapp.com/product/?id=dake_pdf_extract
- Cloudflare反映: dakeapp.com / store.dakeapp.com 本番反映済み
- dakeapp.com掲載: https://dakeapp.com/apps/pdf-extract/

価格は人間の決定により500円で確定しています。Stripe Payment Linkは作成せず、BOOTH正式公開済みのため `payment_status` は `booth_only` です。

## README生成用情報

- 概要: 1つのPDFを開いたまま、必要なページを何度でも連続して抽出するWindows向けDAKEアプリです。
- 使い方: PDFを追加し、サムネイルでページを選択して抽出します。
- 必要なもの: Windows環境
- 注意: 元PDFは変更せず、同名出力は連番で保存します。
- ビルド: `build.bat`

## DAKE_META生成用情報

```json
{
  "app_key": "dake_pdf_extract",
  "display_name": "DakePDF抽出",
  "launcher_title": "PDF抽出",
  "launcher_description": "1つのPDFから、必要なページを何度でも連続して抽出します。",
  "site_title": "DakePDF抽出",
  "site_description": "PDFを開いたまま、選択、抽出、自動選択解除、次の選択を繰り返せるWindows向けアプリです。",
  "update_summary": "DakePDF抽出 v1.0.0を正式公開しました。",
  "folder_name": "DAKE_PDF_Extract",
  "exe_name": "DakePDF_Extract.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_PDF_Extract_v1.0.0",
  "booth_url": "https://peakheadz.booth.pm/items/8746426",
  "store_url": "https://store.dakeapp.com/product/?id=dake_pdf_extract",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "version": "1.0.0",
  "show_in_launcher": true,
  "show_on_site": true,
  "app_type": "market",
  "completion_goal": "formal_shipping",
  "payment_status": "booth_only",
  "demo_video_path": "release_artifacts/demo.mp4",
  "demo_video_url": "",
  "social_release_path": "release_artifacts/social_release.json"
}
```

## release_body生成用情報

- 1つのPDFから必要なページを何度でも連続抽出
- 選択ページを1つのPDFにまとめて保存
- 選択ページを1ページずつ別ファイルで保存
- サムネイル選択、サイズ変更、リフレッシュに対応
- 抽出後の自動選択解除、抽出済み表示、保存先フォルダ表示

## booth_product生成用情報

- 商品名: DakePDF抽出
- 価格案: 500円
- 商品紹介文: 1つのPDFを開いたまま、必要なページを何度でも連続して抽出するWindows向けアプリです。

・1ページ単位または複数ページをまとめて抽出
・選択ページを1ページずつ別ファイルで保存
・サムネイル選択とサムネイルサイズ変更
・元PDF更新後のリフレッシュ
・抽出後の自動選択解除と抽出済み表示
・保存成功後に保存先フォルダを表示

- タグ: PDF
Windows
実務
ツール
仕事効率化
軽量
シンプル
- 商品画像: assets/booth_thumbnail.jpg
- 補助画像: assets/screenshot.jpg
- 作品ファイル: booth_ready/DakePDF_Extract.zip
- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_PDF_Extract_v1.0.0
- BOOTH URL: https://peakheadz.booth.pm/items/8746426

## Store表示用情報

- 商品名: DakePDF抽出
- キャッチ: 1つのPDFから、必要なページを何度でも連続して抽出します。
- 説明: PDFを開いたまま、選択、抽出、自動選択解除、次の選択を繰り返せるWindows向けアプリです。
- 価格: 500円
- 画像: assets/booth_thumbnail.jpg
- ダウンロード導線: https://peakheadz.booth.pm/items/8746426
- サポート方針: GitHub ReleaseとBOOTHで配布します。
- Stripe Payment Link: 未設定
- Store販売状態: booth_only
- Store URL: https://store.dakeapp.com/product/?id=dake_pdf_extract

## 価格・販売方針

- BOOTH価格案: 500円
- BOOTH URL: https://peakheadz.booth.pm/items/8746426
- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_PDF_Extract_v1.0.0
- Store販売: BOOTH導線で販売

## 配布・ダウンロード方針

- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_PDF_Extract_v1.0.0
- BOOTH: https://peakheadz.booth.pm/items/8746426
- BOOTH配布zip: booth_ready/DakePDF_Extract.zip
- Store配布導線: https://peakheadz.booth.pm/items/8746426

## 免責・注意事項

- Windows向けアプリです。
- ご利用は自己責任でお願いいたします。
- 大切なファイルは事前にバックアップを推奨します。
- 本ソフトウェアの無断転載・再配布を禁止します。
- 環境によっては起動時にWindowsの警告が表示される場合があります。

## 同梱ファイル方針

- exe: DakePDF_Extract.exe
- README.txt: booth_ready/README.txt
- 注意事項.txt: booth_ready/注意事項.txt
- 入れないもの: build/、dist/、*.spec、設定ファイル、個人データ、ソース一式

## 派生ファイル方針

- README.md: GitHub閲覧用
- DAKE_META: 機械利用の補助JSON
- release_body.md: GitHub Release本文
- booth_product.txt: BOOTH正式商品情報
- build.bat: PyInstallerビルド

未公開URLは派生ファイルにも記載しません。
