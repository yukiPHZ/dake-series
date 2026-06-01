# DAKE_BOOTH_Assist v2 読み取り点検 2026-06-02

## 原因

BOOTH_Assist はファイル検出には成功していたが、読み取り対象と parser 仕様が Factory v2 とずれていた。

- root の `booth_product.txt` は Factory の確認用テキストで、`商品名`、`概要`、`できること` の section 型。
- BOOTH 登録用の価格、タグ、GitHub Release は主に `booth_ready/booth_product.txt` にある。
- 旧 parser は `TITLE=...` や `# 商品名` 形式を中心に見ており、section 型や ready 正本優先に対応しきれていなかった。

## 読み取り仕様

BOOTH_Assist は BOOTH 登録補助アプリなので、読み取りは次の順にする。

1. `app_dir/booth_ready/booth_product.txt`
2. `app_dir/booth_product.txt`
3. `README.md` の `DAKE_META`

項目ごとの fallback:

- 商品名: booth product、なければ `DAKE_META.display_name`、`launcher_title`、`site_title`
- 価格: booth product
- タグ: booth product
- 説明文: booth product、なければ `DAKE_META.update_summary`
- GitHub Release URL: `DAKE_META.release_url` 優先、なければ booth product 内の GitHub Release URL

## 対応フォーマット

- `TITLE=...` / `PRICE=...` / `DESCRIPTION=...` / `TAGS=...`
- `# 商品名` のような Markdown 見出し型
- `商品名` の次行に値を書く Factory v2 section 型
- `# 価格案`、`# 商品紹介文`、`# タグ`、`# GitHub Release`

## 指定5件の結果

| アプリ | 読み取り元 | 商品名 | 価格 | 未設定 |
|---|---|---|---|---|
| DAKE_PDF_Compress | booth_ready/booth_product.txt | DakePDF圧縮 | 500円 | 0 |
| DAKE_Work_Calendar | booth_ready/booth_product.txt | Dake工程カレンダー | 500円 | 0 |
| DAKE_Mail_Draft | booth_ready/booth_product.txt | Dakeメール下書き | 300円 | 0 |
| DAKE_Column_Memo | booth_ready/booth_product.txt | Dakeずっとメモ | 300円 | 0 |
| DAKE_App_Doko | booth_ready/booth_product.txt | アプリどこ | 300円 | 0 |

対象5件の未設定件数:

| 項目 | 未設定件数 |
|---|---:|
| 商品名 | 0 |
| 価格 | 0 |
| タグ | 0 |
| 説明文 | 0 |
| GitHub Release URL | 0 |

## 全 product app 集計

対象: product source がある 55 app

| 項目 | 未設定件数 |
|---|---:|
| 商品名 | 0 |
| 価格 | 0 |
| タグ | 7 |
| 説明文 | 0 |
| GitHub Release URL | 9 |

タグまたは GitHub Release URL が未設定の app:

- DAKE_Approve_Brainz: GitHub Release URL
- DAKE_BGM_Loop: GitHub Release URL
- DAKE_Brainz_OIKAWA: GitHub Release URL
- DAKE_Brainz_Search: GitHub Release URL
- DAKE_HolidayJinja_Post: タグ、GitHub Release URL
- DAKE_Image_HEICtoJPG: タグ
- DAKE_Image_iPhoneToPC: タグ
- DAKE_Image_PasteA4: タグ
- DAKE_Image_Receiver: タグ
- DAKE_Image_Resize: タグ
- DAKE_Music_Otooku: GitHub Release URL
- DAKE_Screenshot_Print: タグ
- DAKE_Wake_Brainz: GitHub Release URL
- DAKE_YukizBlog_Post: GitHub Release URL
- DAKE_Yukiz_KadouChu: GitHub Release URL

## 検証

- `python -m py_compile 01_apps/DAKE_BOOTH_Assist/main.py`: OK
- `python 01_apps/DAKE_BOOTH_Assist/main.py --launch-check`: OK
- PyInstaller rebuild: OK
- `dist/DakeBOOTH_Assist.exe --launch-check`: exit 0

Computer Use の Windows helper が起動時に落ちたため、実画面スクリーンショットの自動確認は未実施。読み取りロジックと exe の launch-check で確認した。
