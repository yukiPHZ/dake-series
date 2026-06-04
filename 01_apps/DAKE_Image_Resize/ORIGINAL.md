# ORIGINAL.md

## 正本宣言

このファイルは `DAKE_Image_Resize` の真の正本です。

README.md、DAKE_META、release_body.md、booth_product.txt、Store表示などは、このファイルから派生するビューです。

## 基本情報

- app_id: dake_image_resize
- title: Dake画像リサイズ
- short_title: 画像リサイズ
- category: 画像 / リサイズ
- status: available
- version: 1.0.0
- price: 500円
- distribution: GitHub ReleaseとBOOTHで配布する。しまりすくん連携CLIあり。
- target_platform: Windows

## 目的

スマホで撮影した画像を、長辺1600px上限のJPEGに整えて軽量化する。

## 対象ユーザー

- スマホ写真を軽くして共有・保管したい人
- 画像の比率を保ったまま、決まった上限サイズへ整えたい人
- 元画像を上書きせずに、軽量版だけを作りたい人

## 解決する困りごと

- スマホ写真の容量が大きく、送付や保存に扱いづらい
- 画像ごとにリサイズ設定を考えるのが面倒
- 元画像を誤って上書きしたくない

## 主な機能

- 画像リサイズアプリ
- 長辺1600px上限で軽量化
- JPEG出力に対応
- Windows向けexe

## 使い方の要点

- 画像ファイルを追加する。
- 必要に応じて保存先を選ぶ。
- 長辺1600px上限のJPEGとして出力する。
- 出力フォルダ内の軽量化画像を確認する。

## 公開用説明の元情報

スマホ写真などの画像を長辺1600px上限のJPEGへ軽量化できるWindows向けアプリです。

スマホ写真を長辺1600px上限のJPEGに整えます。

実務の流れを、少し静かにするための道具です。

## README生成用情報

- 概要: スマホ写真などの画像を長辺1600px上限のJPEGへ軽量化できるWindows向けアプリです。
- 使い方: 画像ファイルを追加する。
- 必要なもの: Windows環境。
- 注意: 出力はJPEG固定。
- ビルド: `build.bat` を実行し、`dist/Dake_Image_Resize.exe` を生成する。

## DAKE_META生成用情報

既存README内の `DAKE_META` ブロックを元にした機械利用ビュー情報です。
単独の `DAKE_META` ファイルは既存ファイルに存在しません。


```json
{
  "app_key": "dake_image_resize",
  "display_name": "Dake画像リサイズ",
  "launcher_title": "画像リサイズ",
  "launcher_description": "スマホ写真を長辺1600px上限のJPEGに整えます。",
  "site_title": "Dake画像リサイズ",
  "site_description": "スマホ写真などの画像を長辺1600px上限のJPEGへ軽量化できるWindows向けアプリです。",
  "update_summary": "READMEメタ情報とRelease本文を整備。スクリーンショットをassets/screenshot.webpに作成。",
  "folder_name": "DAKE_Image_Resize",
  "exe_name": "Dake_Image_Resize.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Image_Resize_v1.0.0",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## release_body生成用情報

- 画像リサイズアプリ
- 長辺1600px上限で軽量化
- JPEG出力に対応
- Windows向けexe

## booth_product生成用情報

- 商品名: Dake画像リサイズ
- 価格案: 500円
- 商品紹介文: スマホ写真を長辺1600px上限のJPEGに整えます。
- 補足紹介文:
  - 画像リサイズアプリ
  - 長辺1600px上限で軽量化
  - JPEG出力に対応
  - Windows向けexe
  - 実務の流れを、少し静かにするための道具です。
- タグ:
  - 画像
  - Windows
  - 実務
  - ツール
  - 仕事効率化
  - 軽量
  - シンプル
- 商品画像: assets/booth_thumbnail.jpg
- 補助画像: assets/screenshot.jpg
- 作品ファイル: booth_ready/Dake_Image_Resize.zip
- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Image_Resize_v1.0.0
- BOOTH URL: https://peakheadz.booth.pm/items/8397622

## Store表示用情報

- 商品名: Dake画像リサイズ
- キャッチ: スマホ写真を長辺1600px上限のJPEGに整えます。
- キャッチ補足: 実務の流れを、少し静かにするための道具です。
- 説明: スマホ写真などの画像を長辺1600px上限のJPEGへ軽量化できるWindows向けアプリです。
- 価格: 500円
- 画像: assets/booth_thumbnail.jpg / assets/screenshot.webp
- ダウンロード導線: 未確定
- サポート方針: 既存ファイルに記載なし

Storeは未構築のため、Store専用の商品正本は作りません。

## 価格・販売方針

- BOOTH価格案: 500円
- BOOTH URL: https://peakheadz.booth.pm/items/8397622
- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Image_Resize_v1.0.0
- Store販売: 未確定

## 配布・ダウンロード方針

- GitHub Releaseで `Dake_Image_Resize.exe` を配布する。
- BOOTHでは `booth_ready/Dake_Image_Resize.zip` を作品ファイルとして使う。
- dakeapp.com掲載対象です。
- Store配布導線は未確定です。

## 免責・注意事項

BOOTH ready内の注意事項、または既存READMEの注意事項を元にします。

- 出力はJPEG固定。
- 長辺は1600pxを上限にし、比率を維持する。
- 元画像は上書きしない。
- HEIC / HEIF は初期版では非対応。

## 同梱ファイル方針

- exe: Dake_Image_Resize.exe
- README.txt: booth_ready/README.txt (既存)
- 注意事項.txt: booth_ready/注意事項.txt (既存)
- 配布zip: booth_ready/Dake_Image_Resize.zip
- 入れないもの: ソースコード、build/、dist/、*.spec、__pycache__/、個人設定ファイル

## スクリーンショット・画像方針

- assets/screenshot.webp: assets/screenshot.webp
- assets/screenshot.jpg: assets/screenshot.jpg
- assets/booth_thumbnail.jpg: assets/booth_thumbnail.jpg
- Store用画像: 未確定。既存画像を元に派生する想定。

## 今後の改善予定

既存README、release_body.md、booth_product.txtには今後の改善予定の記載なし。

現時点では未設定です。

## Codex作業時の注意

- 触ってよい: ORIGINAL.mdの更新、派生ビューとの整合確認。
- 触らない: main.py、build.bat、assets、dist、booth_readyの内容をこのPhaseで変更しない。
- 外部公開しない: 未確定のStore URLや未確認の販売導線を確定情報として書かない。
- 自動操作しない: BOOTH更新、GitHub Release更新、Store構築、Stripe実装はこのPhaseでは行わない。

## 派生物一覧

- README.md: GitHub公開用ビュー。既存。
- DAKE_META: README.md内のJSONブロックとして存在。単独ファイルはなし。
- release_body.md: GitHub Release用ビュー。既存。
- booth_product.txt: BOOTH登録用ビュー。アプリ直下は 既存、`booth_ready/` は 既存。
- booth_ready/README.txt: 配布zip同梱用ビュー。既存。
- booth_ready/注意事項.txt: 配布zip同梱用ビュー。既存。
- Store: 未構築。将来 `ORIGINAL.md` 由来の情報から生成する。
