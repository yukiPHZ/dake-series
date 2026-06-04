# ORIGINAL.md

## 正本宣言

このファイルは `DAKE_Folder_List` の真の正本です。

README.md、DAKE_META、release_body.md、booth_product.txt、Store表示などは、このファイルから派生するビューです。

## 基本情報

- app_id: dake_folder_list
- title: Dakeフォルダ一覧
- short_title: フォルダ一覧
- category: ファイル / 一覧
- status: available
- version: 1.0.0
- price: 300円
- distribution: GitHub ReleaseとBOOTHで配布する。
- target_platform: Windows

## 目的

指定フォルダを基準に、フォルダ構成とファイル一覧を見える化する。

## 対象ユーザー

- フォルダ構成やファイル一覧をすばやく把握したい人
- 一覧をコピーして共有・記録したい人
- フォルダ内容を編集せず、見るだけの安全な確認ツールがほしい人

## 解決する困りごと

- フォルダ内のファイル構成を手で書き出すのが面倒
- ファイル名、拡張子、サイズ、更新日時をまとめて見たい
- 一覧作成ツールに編集・削除機能があると不安

## 主な機能

- フォルダ一覧作成アプリ
- フォルダ構成とファイル情報を表示
- コピー・txt保存に対応
- Windows向けexe

## 使い方の要点

- フォルダを選ぶ。
- 表示された構成とファイル情報を確認する。
- 必要に応じて一覧をコピーまたはtxt保存する。
- フォルダ内容を更新したら再読み込みする。

## 公開用説明の元情報

指定フォルダの構成、ファイル名、拡張子、サイズ、更新日時を一覧化できるWindows向けアプリです。

フォルダ構成とファイル一覧を表示・コピー・保存します。

実務の流れを、少し静かにするための道具です。

## README生成用情報

- 概要: 指定フォルダの構成、ファイル名、拡張子、サイズ、更新日時を一覧化できるWindows向けアプリです。
- 使い方: フォルダを選ぶ。
- 必要なもの: Windows環境。
- 注意: ファイルの編集、削除、移動、リネームは行わない。
- ビルド: `build.bat` を実行し、`dist/DakeFolder_List.exe` を生成する。

## DAKE_META生成用情報

既存README内の `DAKE_META` ブロックを元にした機械利用ビュー情報です。
単独の `DAKE_META` ファイルは既存ファイルに存在しません。


```json
{
  "app_key": "dake_folder_list",
  "display_name": "Dakeフォルダ一覧",
  "launcher_title": "フォルダ一覧",
  "launcher_description": "フォルダ構成とファイル一覧を表示・コピー・保存します。",
  "site_title": "Dakeフォルダ一覧",
  "site_description": "指定フォルダの構成、ファイル名、拡張子、サイズ、更新日時を一覧化できるWindows向けアプリです。",
  "update_summary": "READMEメタ情報とRelease本文を整備。スクリーンショットをassets/screenshot.webpに作成。",
  "folder_name": "DAKE_Folder_List",
  "exe_name": "DakeFolder_List.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Folder_List_v1.0.0",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## release_body生成用情報

- フォルダ一覧作成アプリ
- フォルダ構成とファイル情報を表示
- コピー・txt保存に対応
- Windows向けexe

## booth_product生成用情報

- 商品名: Dakeフォルダ一覧
- 価格案: 300円
- 商品紹介文: フォルダ構成とファイル一覧を表示・コピー・保存します。
- 補足紹介文:
  - フォルダ一覧作成アプリ
  - フォルダ構成とファイル情報を表示
  - コピー・txt保存に対応
  - Windows向けexe
  - 実務の流れを、少し静かにするための道具です。
- タグ:
  - Windows
  - 実務
  - ツール
  - 仕事効率化
  - 軽量
  - シンプル
- 商品画像: assets/booth_thumbnail.jpg
- 補助画像: assets/screenshot.jpg
- 作品ファイル: booth_ready/DakeFolder_List.zip
- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Folder_List_v1.0.0
- BOOTH URL: https://peakheadz.booth.pm/items/8447591

## Store表示用情報

- 商品名: Dakeフォルダ一覧
- キャッチ: フォルダ構成とファイル一覧を表示・コピー・保存します。
- キャッチ補足: 実務の流れを、少し静かにするための道具です。
- 説明: 指定フォルダの構成、ファイル名、拡張子、サイズ、更新日時を一覧化できるWindows向けアプリです。
- 価格: 300円
- 画像: assets/booth_thumbnail.jpg / assets/screenshot.webp
- ダウンロード導線: 未確定
- サポート方針: 既存ファイルに記載なし

Storeは未構築のため、Store専用の商品正本は作りません。

## 価格・販売方針

- BOOTH価格案: 300円
- BOOTH URL: https://peakheadz.booth.pm/items/8447591
- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Folder_List_v1.0.0
- Store販売: 未確定

## 配布・ダウンロード方針

- GitHub Releaseで `DakeFolder_List.exe` を配布する。
- BOOTHでは `booth_ready/DakeFolder_List.zip` を作品ファイルとして使う。
- dakeapp.com掲載対象です。
- Store配布導線は未確定です。

## 免責・注意事項

BOOTH ready内の注意事項、または既存READMEの注意事項を元にします。

- ファイルの編集、削除、移動、リネームは行わない。
- ファイル内容の読み取り、検索、分類は行わない。
- アクセスできないフォルダはスキップし、集計にスキップ件数を表示する。

## 同梱ファイル方針

- exe: DakeFolder_List.exe
- README.txt: booth_ready/README.txt (既存)
- 注意事項.txt: booth_ready/注意事項.txt (既存)
- 配布zip: booth_ready/DakeFolder_List.zip
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
