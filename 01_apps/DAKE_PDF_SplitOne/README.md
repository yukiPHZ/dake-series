# DakePDF分割One

`DakePDF分割One` は、PDFを1ファイル入れると全ページを1枚ずつのPDFへ自動分割する、DAKEシリーズ向けの単機能アプリです。判断項目や実行ボタンを置かず、ドラッグ＆ドロップを主導線にしています。

## アプリ概要

- PDF 1ファイルだけを受け付けます
- 投入したら確認なしで自動分割を始めます
- 保存先に `元ファイル名_split` フォルダを作成します。既にある場合は `_2`、`_3` のように別名で作成します
- 各ページを `p001.pdf` `p002.pdf` のように保存します
- 分割中でも `中止` または F5 ですぐ初期状態へ戻せます

## 使い方

1. アプリを起動します
2. PDFをドラッグ＆ドロップします
3. 完了ダイアログの OK を押します
4. 保存フォルダが自動で開きます

補助操作として、中央エリアをクリックしてPDFを選ぶこともできます。保存先は上部の `保存先を選ぶ` から変更できます。
`Ctrl+O` でもPDF選択を開けます。分割中の `中止` または F5 は現在の処理を取り消し、次のPDFを受け付ける初期状態へ戻します。

## 保存先

- 初期値は `Downloads`
- 保存先は設定ファイル `dake_pdf_split_one_config.json` に保持されます
- 出力先フォルダ名は最初が `元ファイル名_split`、同名がある場合は `元ファイル名_split_2`、`元ファイル名_split_3` です

例:

- `sample.pdf`
- `Downloads\sample_split\p001.pdf`
- `Downloads\sample_split\p002.pdf`
- `Downloads\sample_split\p003.pdf`

同じ `sample.pdf` をもう一度分割した場合は、既存の `sample_split` を変更せず `sample_split_2` を作成します。

## 出力形式

- 出力はページごとの個別PDFです
- ファイル名は3桁ゼロ埋めです
- ページごとに別フォルダは作りません

## 注意事項

- PDF以外は追加できません
- 複数ファイルの同時投入はできません
- 既存の出力フォルダや `p001.pdf` などの成果物は上書き・削除しません
- 分割途中のファイルは一時フォルダへ保存し、成功した場合だけ正式な出力フォルダとして確定します
- 大きなPDFでもUIが止まらないよう、分割処理は別スレッドで実行します

## ビルド方法

```bat
python -m pip install -r requirements.txt
build.bat
```

ビルドが成功すると `dist\DakePDF_Split_One\DakePDF_Split_One.exe` と `_internal` フォルダが生成されます。exe単体ではなくフォルダ一式で動作します。

正式配布ZIPには、runtimeフォルダ一式、README.txt、注意事項.txt、THIRD_PARTY_NOTICES.txt、third_party_licensesを同梱します。

`python` や `py` がPATHにない場合は、先に `set PYTHON_EXE=C:\Path\To\python.exe` を指定してから `build.bat` を実行できます。

## Shimarisu CLI

`--from-shimarisu` がある場合だけ、GUIを表示せずにCLIモードでPDFをページ単位に分割します。通常起動では従来どおりGUIを表示します。

```bat
dist\DakePDF_Split_One\DakePDF_Split_One.exe --from-shimarisu --inputs "C:\path\sample.pdf"
dist\DakePDF_Split_One\DakePDF_Split_One.exe --from-shimarisu --inputs "C:\path\sample.pdf" --output "C:\path\out" --silent
```

- `--inputs`: PDFファイルを1件以上指定します。複数指定時は先頭の1件だけ使用します
- `--output`: 出力先フォルダです。未指定時は元PDFと同じフォルダに `split_YYYYMMDD_HHMMSS` を作成します
- `--from-shimarisu`: CLIモード起動フラグです
- `--silent`: しまりすくん側から渡せる任意フラグです
- 正常時は exit code 0、エラー時は exit code 1 を返します

出力ファイル名は `sample_p001.pdf`、`sample_p002.pdf` の形式です。

## DAKE_META

```json
{
  "app_key": "dake_pdf_splitone",
  "display_name": "DakePDF分割One",
  "launcher_title": "PDF分割One",
  "launcher_description": "PDFを全ページ1枚ずつに自動分割します。",
  "site_title": "DakePDF分割One",
  "site_description": "PDFを1ファイル追加すると、全ページを1枚ずつのPDFへ自動分割できるWindows向けアプリです。",
  "update_summary": "READMEメタ情報とRelease本文を整備。スクリーンショットをassets/screenshot.webpに作成。",
  "folder_name": "DAKE_PDF_SplitOne",
  "exe_name": "DakePDF_Split_One.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_PDF_SplitOne_v1.0.1",
  "version": "1.0.1",
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

- PDF全ページ分割アプリ
- 1ページずつ自動保存
- ドラッグ＆ドロップ対応
- Windows向けexe
- F5で処理を中断して初期状態に戻し、Ctrl+Oで次のPDFを開けます
