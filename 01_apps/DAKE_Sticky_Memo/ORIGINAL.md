# ORIGINAL.md

## 正本宣言

このファイルは `DAKE_Sticky_Memo` の真の正本です。

README.md、DAKE_META、release_body.md、booth_product.txt、Store表示などは、このファイルから派生するビューです。

## 基本情報

- app_id: DAKE_Sticky_Memo
- title: 付箋メモ
- short_title: 付箋メモ
- category: メモ / 付箋
- status: available
- version: 1.0.0
- price: 300円
- distribution: GitHub ReleaseとBOOTHで配布する。
- target_platform: Windows

## 目的

アプリウインドウ内に付箋メモを作成し、動かし、不要になったら捨てられるようにする。

## 対象ユーザー

- 思いついたことを一時的に画面内へ置きたい人
- 保存や分類を考えず、その場のメモだけ使いたい人
- 軽い付箋メモをWindows上で使いたい人

## 解決する困りごと

- 一時メモに保存・分類・検索まであると重い
- 小さなメモを作業画面内で自由に動かしたい
- 終わったメモをすぐ捨てたい

## 主な機能

- 付箋を追加して、その場でメモできます。
- 付箋をドラッグして自由に並べ替えできます。
- 不要な付箋を個別削除、または全消去できます。
- Windows向けexeです。

## 使い方の要点

- 付箋を追加する。
- 付箋にテキストを書く。
- 必要な位置へドラッグする。
- 不要になったら個別削除、または全消去する。

## 公開用説明の元情報

アプリ内に付箋を作成し、自由に動かし、不要になったら消せるWindows向け軽量メモアプリです。

付箋を書いて、動かして、捨てられる軽量メモです。

実務の流れを、少し静かにするための道具です。

## README生成用情報

- 概要: アプリ内に付箋を作成し、自由に動かし、不要になったら消せるWindows向け軽量メモアプリです。
- 使い方: 付箋を追加する。
- 必要なもの: Windows環境。
- 注意: 初期版では保存しない。
- ビルド: `build.bat` を実行し、`dist/DakeSticky_Memo.exe` を生成する。

## DAKE_META生成用情報

既存README内の `DAKE_META` ブロックを元にした機械利用ビュー情報です。
単独の `DAKE_META` ファイルは既存ファイルに存在しません。


```json
{
  "app_key": "DAKE_Sticky_Memo",
  "display_name": "付箋メモ",
  "launcher_title": "付箋メモ",
  "launcher_description": "付箋を書いて、動かして、捨てられる軽量メモです。",
  "site_title": "Dake付箋メモ",
  "site_description": "アプリ内に付箋を作成し、自由に動かし、不要になったら消せるWindows向け軽量メモアプリです。",
  "update_summary": "フッター表示とUI文言をDAKE共通仕様に合わせて調整しました。",
  "folder_name": "DAKE_Sticky_Memo",
  "exe_name": "DakeSticky_Memo.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Sticky_Memo_v1.0.0",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## release_body生成用情報

- 付箋を追加して、その場でメモできます。
- 付箋をドラッグして自由に並べ替えできます。
- 不要な付箋を個別削除、または全消去できます。
- Windows向けexeです。

## booth_product生成用情報

- 商品名: Dake付箋メモ
- 価格案: 300円
- 商品紹介文: 付箋を書いて、動かして、捨てられる軽量メモです。
- 補足紹介文:
  - 付箋を追加して、その場でメモできます。
  - 付箋をドラッグして自由に並べ替えできます。
  - 不要な付箋を個別削除、または全消去できます。
  - Windows向けexeです。
  - 実務の流れを、少し静かにするための道具です。
- タグ:
  - メモ
  - Windows
  - 実務
  - ツール
  - 仕事効率化
  - 軽量
  - シンプル
- 商品画像: assets/booth_thumbnail.jpg
- 補助画像: assets/screenshot.jpg
- 作品ファイル: booth_ready/DakeSticky_Memo.zip
- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Sticky_Memo_v1.0.0
- BOOTH URL: https://peakheadz.booth.pm/items/8448263

## Store表示用情報

- 商品名: 付箋メモ
- キャッチ: 付箋を書いて、動かして、捨てられる軽量メモです。
- キャッチ補足: 実務の流れを、少し静かにするための道具です。
- 説明: アプリ内に付箋を作成し、自由に動かし、不要になったら消せるWindows向け軽量メモアプリです。
- 価格: 300円
- 画像: assets/booth_thumbnail.jpg / assets/screenshot.webp
- ダウンロード導線: 未確定
- サポート方針: 既存ファイルに記載なし

Storeは未構築のため、Store専用の商品正本は作りません。
- Stripe Payment Link: https://buy.stripe.com/9B65kD2Zzaon9WE0bf0gw0c
- Store雋ｩ螢ｲ迥ｶ諷・ stripe_ready

## 価格・販売方針

- BOOTH価格案: 300円
- BOOTH URL: https://peakheadz.booth.pm/items/8448263
- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Sticky_Memo_v1.0.0
- Store販売: 未確定

## 配布・ダウンロード方針

- GitHub Releaseで `DakeSticky_Memo.exe` を配布する。
- BOOTHでは `booth_ready/DakeSticky_Memo.zip` を作品ファイルとして使う。
- dakeapp.com掲載対象です。
- Store配布導線は未確定です。

## 免責・注意事項

BOOTH ready内の注意事項、または既存READMEの注意事項を元にします。

- 初期版では保存しない。
- 色変更、検索、タグ、画像添付、エクスポート、分類は行わない。
- 付箋位置はウインドウ外へ出ないよう制限する。

## 同梱ファイル方針

- exe: DakeSticky_Memo.exe
- README.txt: booth_ready/README.txt (既存)
- 注意事項.txt: booth_ready/注意事項.txt (既存)
- 配布zip: booth_ready/DakeSticky_Memo.zip
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
