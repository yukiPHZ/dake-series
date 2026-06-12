# DakePDFトリミング

PDFをドラッグ指定した範囲でトリミングして保存する単機能アプリです。

名刺スキャンPDFなど、余白を切りたい用途向けです。自動判定ではなく、ユーザーが残したい範囲を手動で指定することで安定性を優先しています。

## 使い方

1. `PDFを選ぶ`、またはPDFを画面にドラッグ＆ドロップします。
2. 複数ページPDFの場合は、`前へ` / `次へ`でトリミングしたいページを選びます。
3. プレビュー上で、残したい範囲をドラッグします。
4. `この範囲で保存`を押します。
5. 元PDFと同じフォルダに`元ファイル名_pページ番号_crop.pdf`が保存され、保存フォルダが開きます。

`リフレッシュ`を押すと、現在のPDFを読み直し、1ページ目・範囲未選択の状態に戻します。

## 仕様

- PDFのみを対象にしています。
- 選択した1ページだけをトリミングし、1ページPDFとして保存します。
- 他ページは保存しません。
- 既存ファイルがある場合は、上書きせず連番で保存します。
- 回転が含まれるPDFは、PDFの状態によって意図通りの範囲にならない場合があります。

## 注意

PDFの状態により、意図通り表示・保存されない場合があります。保存後に必ず内容を確認してください。

## ビルド

```bat
build.bat
```

生成物は`dist\DakePDF_Crop.exe`です。

## DAKE_META

```json
{
  "app_key": "dake_pdf_crop",
  "display_name": "DakePDFトリミング",
  "launcher_title": "PDFトリミング",
  "launcher_description": "PDFをドラッグ指定した範囲でトリミングします。",
  "site_title": "DakePDFトリミング",
  "site_description": "PDFプレビュー上で残したい範囲をドラッグし、トリミング済みPDFとして保存できるWindows向けアプリです。",
  "update_summary": "READMEメタ情報とRelease本文を整備。スクリーンショットをassets/screenshot.webpに作成。",
  "folder_name": "DAKE_PDF_Crop",
  "exe_name": "DakePDF_Crop.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_PDF_Crop_v1.0.0",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true,
  "demo_video_path": "release_artifacts/demo.mp4",
  "demo_video_url": "",
  "social_release_path": "release_artifacts/social_release.json"
}
```

## RELEASE_BODY

- PDFトリミングアプリ
- ドラッグ範囲指定に対応
- 別名PDFとして保存
- Windows向けexe
