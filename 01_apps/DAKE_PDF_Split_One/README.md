# DakePDF分割One

`DakePDF分割One` は、PDFを1ファイル入れると全ページを1枚ずつのPDFへ自動分割する、DAKEシリーズ向けの単機能アプリです。判断項目や実行ボタンを置かず、ドラッグ＆ドロップを主導線にしています。

## アプリ概要

- PDF 1ファイルだけを受け付けます
- 投入したら確認なしで自動分割を始めます
- 保存先に `元ファイル名_split` フォルダを作成します
- 各ページを `p001.pdf` `p002.pdf` のように保存します

## 使い方

1. アプリを起動します
2. PDFをドラッグ＆ドロップします
3. 完了ダイアログの OK を押します
4. 保存フォルダが自動で開きます

補助操作として、中央エリアをクリックしてPDFを選ぶこともできます。保存先は上部の `保存先を選ぶ` から変更できます。

## 保存先

- 初期値は `Downloads`
- 保存先は設定ファイル `dake_pdf_split_one_config.json` に保持されます
- 出力先フォルダ名は `元ファイル名_split` です

例:

- `sample.pdf`
- `Downloads\sample_split\p001.pdf`
- `Downloads\sample_split\p002.pdf`
- `Downloads\sample_split\p003.pdf`

## 出力形式

- 出力はページごとの個別PDFです
- ファイル名は3桁ゼロ埋めです
- ページごとに別フォルダは作りません

## 注意事項

- PDF以外は追加できません
- 複数ファイルの同時投入はできません
- 既存の `p001.pdf` など同名ファイルは上書きされます
- 大きなPDFでもUIが止まらないよう、分割処理は別スレッドで実行します

## ビルド方法

```bat
python -m pip install -r requirements.txt
build.bat
```

ビルドが成功すると `dist\DakePDF_Split_One.exe` が生成されます。

`python` や `py` がPATHにない場合は、先に `set PYTHON_EXE=C:\Path\To\python.exe` を指定してから `build.bat` を実行できます。
