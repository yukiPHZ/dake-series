# DakePDF抽出

DakePDF抽出は、1つのPDFを開いたまま、必要なページを何度でも連続して抽出するWindows向けDAKEアプリです。

正式価格は500円です。

## 特徴

- PDFを1ファイルだけ読み込み
- `PDFを追加`、空状態中央の選択領域、ドラッグ＆ドロップに対応
- 全ページのサムネイル一覧表示
- 読み込み済みPDFをリフレッシュして最新状態へ再読み込み
- クリックで選択、再クリックで解除
- Shift+クリックで範囲選択
- Enterで抽出、Escで選択解除
- 選択ページを1つのPDFにまとめて保存
- 選択ページを1ページずつ保存
- 保存成功後もPDFとスクロール位置を維持
- 保存成功後は選択を自動解除し、保存先フォルダを表示
- 抽出済みページは薄い表示と小さなチェックで表示
- サムネイルサイズをボタン、スライダー、Ctrl+マウスホイールで変更

## 使い方

1. `PDFを追加` または空状態中央の選択領域からPDFを選ぶか、PDFを画面へドラッグ＆ドロップします。
2. サムネイルをクリックして抽出したいページを選びます。
3. 必要に応じて抽出方法を選びます。
4. `抽出して保存` または Enter で保存します。
5. 選択が自動解除され、保存先フォルダが表示されるので、続けて次のページを選べます。同じフォルダが既に開いている場合はそのwindowを前面に出します。

元PDFを外部アプリで更新した場合は、`リフレッシュ` で同じPDFを再読み込みできます。保存先とサムネイルサイズは維持し、選択状態と抽出済み表示はリセットします。

## 保存先とファイル名

初期保存先は最初に読み込んだ元PDFと同じフォルダです。保存先は `保存先を変更` から変更でき、手動で変更した保存先はその起動中の以後のPDFでも優先されます。

出力ファイル名は元PDF名とページ番号から生成します。

- 1ページ: `元PDF名_p001.pdf`
- 連続ページ: `元PDF名_p002-004.pdf`
- 非連続ページ: `元PDF名_p002_004_007.pdf`
- 1ページずつ: `元PDF名_p002.pdf` など

同名ファイルがある場合は上書きせず、`_2`、`_3` のように連番を付けます。

## ビルド

同じフォルダで以下を実行します。

```bat
build.bat
```

生成物は `dist\DakePDF_Extract.exe` です。

## 開発確認

```powershell
python -m py_compile main.py
python main.py --launch-check
python main.py --process-check
python main.py --self-check
```

## DAKE_META

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

## RELEASE_BODY

- 1つのPDFから必要なページを何度でも連続抽出
- 選択ページを1つのPDFにまとめて保存
- 選択ページを1ページずつ別ファイルで保存
- サムネイル選択、サイズ変更、リフレッシュに対応
- 抽出後の自動選択解除、抽出済み表示、保存先フォルダ表示

## 出荷状況

初回正式バージョンは `1.0.0`、tagは `DAKE_PDF_Extract_v1.0.0` です。[GitHub Release](https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_PDF_Extract_v1.0.0)、[BOOTH商品](https://peakheadz.booth.pm/items/8746426)、[dakeapp.com](https://dakeapp.com/apps/pdf-extract/)、[DAKE Store](https://store.dakeapp.com/product/?id=dake_pdf_extract)を正式公開しています。BOOTH価格は500円です。
