# DakePDFここ見て

PDFに「ここを見てほしい」箇所を伝えるための、DAKEシリーズ用 Windows デスクトップアプリです。

## できること

- PDFを開く
- 赤い丸を付ける
- 赤い矢印を付ける
- 最後に付けた丸または矢印を1つ戻す
- 元PDFとは別名で保存する

## できないこと

- 文字入力、コメント、付箋、スタンプ
- マーカー、フリーハンド、色選択、線幅選択
- ページ編集、削除、並べ替え
- OCR、検索、PDF結合、PDF分割

## 使い方

1. `PDFを開く` でPDFを選択します。
2. `○ 丸` または `→ 矢印` を選びます。
3. PDF上をドラッグして、確認箇所に丸または矢印を付けます。
4. 間違えた場合は `戻す` で最後の1操作を戻します。
5. `保存` で `元ファイル名_ここ見て.pdf` として別名保存します。

## ビルド方法

事前に必要なライブラリをインストールします。

```bat
pip install pymupdf pyinstaller
```

その後、アプリフォルダで以下を実行します。

```bat
build.bat
```

`dist/DakePDF_LookHere.exe` が作成されます。

## DAKEシリーズ共通思想

単機能、軽量、爆速、迷わない。多機能化せず、現場で仕事が止まらないことを優先します。

## 2026-05-06 DAKE共通仕様レビュー

- UI_TEXTの文字化けを修正し、日本語UI文言を上部の `UI_TEXT` に集約しました。
- ヘッダーを「機能タイトル + 短い機能説明」に整理し、画面内のアプリ名重複表示を避けました。
- フッターをDAKE共通仕様に合わせ、狭幅時は中央寄せ2段構成へ切り替えるようにしました。
- 共通アイコン `../../02_assets/dake_icon.ico` の参照を維持しました。
- `build.bat` による exe 再ビルドに成功しました。
- `dist/DakePDF_LookHere.exe` は短時間起動確認で正常に起動しました。

## DAKE_META

```json
{
  "app_key": "dake_pdf_lookhere",
  "display_name": "DakePDFここ見て",
  "launcher_title": "PDFここ見て",
  "launcher_description": "PDFに赤い丸と矢印を付けます。",
  "site_title": "DakePDFここ見て",
  "site_description": "PDFに赤い丸や矢印を付けて、確認してほしい箇所を伝えられるWindows向けアプリです。",
  "update_summary": "READMEメタ情報とRelease本文を整備。スクリーンショットをassets/screenshot.webpに作成。",
  "folder_name": "DAKE_PDF_LookHere",
  "exe_name": "DakePDF_LookHere.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## RELEASE_BODY

- PDF注目箇所マークアプリ
- 赤い丸・矢印に対応
- 元PDFとは別名で保存
- Windows向けexe
