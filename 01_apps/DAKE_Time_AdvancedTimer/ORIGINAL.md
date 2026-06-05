# ORIGINAL.md

## 正本宣言

このファイルは `DAKE_Time_AdvancedTimer` の真の正本です。

README.md、DAKE_META、release_body.md、booth_product.txt、Store表示などは、このファイルから派生するビューです。

## 基本情報

- app_id: time_advanced_timer
- title: Dakeアドバンスドタイマー
- short_title: アドバンスドタイマー
- category: 時間 / タイマー
- status: available
- version: 1.0.0
- price: 300円
- distribution: GitHub Release / BOOTH / dakeapp.com
- target_platform: Windows

## 目的

集中時間と休憩時間を、迷わずすぐに始めるための静かなタイマーです。

20年前の自分との約束を、いま起動できる形にしたDAKEアプリです。

## 対象ユーザー

- 25分集中、5分休憩、15分休憩をすぐに使いたい人
- 作業中に大きな設定画面や複雑な操作を挟みたくない人
- Windows上で静かに動く軽いタイマーを使いたい人

## 解決する困りごと

- 集中時間と休憩時間を始めるまでに迷う
- タイマー設定やモード選択が多く、作業前に疲れる
- 作業中の画面に強い演出や過剰な通知を置きたくない

## 主な機能

- 25分集中
- 5分休憩
- 15分休憩
- 1〜180分のカスタム時間
- 開始、一時停止、リセット
- 終了後の控えめな完了表示と「もう一回」

## 使い方の要点

1. 25分集中、5分休憩、15分休憩、またはカスタム時間を選ぶ。
2. 開始する。
3. 必要に応じて一時停止、リセット、もう一回を使う。

## 公開用説明の元情報

集中時間と休憩時間をすぐに始められる静かなタイマー。

25分集中、5分休憩、15分休憩、カスタム時間に対応したWindows向けタイマーです。

実務の流れを、少し静かにするための道具です。

## README生成用情報

- 概要: 集中時間と休憩時間をすぐに始めるための静かなタイマー。
- 使い方: 25分集中、5分休憩、15分休憩、またはカスタム時間を選び、開始する。
- 必要なもの: Windows環境。
- 注意: 既存READMEに詳細な注意事項の記載なし。
- ビルド: 同じフォルダ内で `build.bat` を実行し、`dist/DakeAdvanced_Timer.exe` を生成する。

## DAKE_META生成用情報

既存README内の `DAKE_META` ブロックを元にした機械利用ビュー情報です。
単独の `DAKE_META` ファイルは既存ファイルに存在しません。

```json
{
  "app_key": "time_advanced_timer",
  "display_name": "Dakeアドバンスドタイマー",
  "launcher_title": "Dakeアドバンスドタイマー",
  "launcher_description": "集中時間と休憩時間をすぐに始められる静かなタイマー。",
  "site_title": "Dakeアドバンスドタイマー",
  "site_description": "25分集中、5分休憩、15分休憩、カスタム時間に対応したWindows向けタイマーです。",
  "update_summary": "集中・休憩・カスタム時間に対応したタイマーを正式出荷準備しました。",
  "folder_name": "DAKE_Time_AdvancedTimer",
  "exe_name": "DakeAdvanced_Timer.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Time_AdvancedTimer_v1.0.0",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true,
  "app_type": "market",
  "completion_goal": "formal_release"
}
```

## release_body生成用情報

- 集中時間と休憩時間をすぐ始めるタイマー
- 25分 / 5分 / 15分 / カスタム時間に対応
- 静かなDAKE UI
- Windows向けexe

## booth_product生成用情報

- 商品名: Dakeアドバンスドタイマー
- 価格案: 300円
- 商品紹介文: 集中時間と休憩時間をすぐに始められる静かなタイマー。
- 補足紹介文:
  - 集中時間と休憩時間をすぐ始めるタイマー
  - 25分 / 5分 / 15分 / カスタム時間に対応
  - 静かなDAKE UI
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
- 作品ファイル: booth_ready/DakeAdvanced_Timer.zip
- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Time_AdvancedTimer_v1.0.0
- BOOTH URL: https://peakheadz.booth.pm/items/8449210

## Store表示用情報

- 商品名: Dakeアドバンスドタイマー
- キャッチ: 集中時間と休憩時間をすぐに始められる静かなタイマー。
- キャッチ補足: 実務の流れを、少し静かにするための道具です。
- 説明: 25分集中、5分休憩、15分休憩、カスタム時間に対応したWindows向けタイマーです。
- 価格: 300円
- 画像: assets/booth_thumbnail.jpg / assets/screenshot.webp
- ダウンロード導線: 未確定
- サポート方針: 既存ファイルに記載なし
- Stripe Payment Link: https://buy.stripe.com/5kQdR9eIh8gfgl25vz0gw00
- Store販売状態: stripe_ready

Storeは未構築のため、Store専用の商品正本は作りません。


## 購入後導線・サポート方針

- MVPではStripe購入後の自動ダウンロード発行は未実装。
- GitHub Release / BOOTH導線を案内候補とする。
- 将来的にStore購入後URLまたはR2連携を検討する。
- ダウンロード不備、不具合報告には可能な範囲で対応する。
- 個別環境での完全な動作保証はしない。
## 価格・販売方針

- BOOTH価格案: 300円
- BOOTH URL: https://peakheadz.booth.pm/items/8449210
- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Time_AdvancedTimer_v1.0.0
- Store販売: 未確定

## 配布・ダウンロード方針

- GitHub Releaseで `DakeAdvanced_Timer.exe` を配布する。
- BOOTHでは `booth_ready/DakeAdvanced_Timer.zip` を作品ファイルとして使う。
- dakeapp.com掲載対象です。
- Store配布導線は未確定です。

## 免責・注意事項

BOOTH ready内の注意事項を元にします。

- Windows向けアプリです。
- ご利用は自己責任でお願いいたします。
- 大切なファイルは事前にバックアップを推奨します。
- 本ソフトウェアの無断転載・再配布を禁止します。
- 環境によっては起動時にWindowsの警告が表示される場合があります。

## 同梱ファイル方針

- exe: DakeAdvanced_Timer.exe
- README.txt: booth_ready/README.txt
- 注意事項.txt: booth_ready/注意事項.txt
- 配布zip: booth_ready/DakeAdvanced_Timer.zip
- 入れないもの: ソースコード、build/、dist/、*.spec、__pycache__/、個人設定ファイル

## スクリーンショット・画像方針

- assets/screenshot.webp: 正式出荷スクリーンショット。実際に動く感を伝える。
- assets/screenshot.jpg: BOOTH補助画像。
- assets/booth_thumbnail.jpg: BOOTH一覧用の商品画像。
- booth_ready/booth_thumbnail.jpg: BOOTH登録時の実使用画像。
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
- booth_product.txt: BOOTH登録用ビュー。アプリ直下と `booth_ready/` に既存。
- Store: 未構築。将来 `ORIGINAL.md` 由来の情報から生成する。
