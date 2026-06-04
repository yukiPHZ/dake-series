# ORIGINAL.md

## 正本宣言

このファイルは `DAKE_Mail_List` の真の正本です。

README.md、DAKE_META、release_body.md、booth_product.txt、Store表示などは、このファイルから派生するビューです。

## 基本情報

- app_id: dake_mail_list
- title: Dakeメールリスト
- short_title: Dakeメールリスト
- category: メール
- status: available
- version: 1.0.0（Release URLより）
- price: 300円
- distribution: GitHub Release / BOOTH / dakeapp.com（Storeは未確定）
- target_platform: Windows

## 目的

Outlookの.msgメールから、会社名・お名前・メールアドレスだけをCSVにする小さな実務ツールです。

## 対象ユーザー

既存ファイルに明示なし

## 解決する困りごと

- Outlook .msg メールから3列CSVを作成します。
- Outlookの.msgメールから、会社名・お名前・メールアドレスだけをCSVにする小さな実務ツールです。

## 主な機能

- Dakeメールリスト
- Outlookの.msgメールをCSV化
- 会社名・お名前・メールアドレスだけ抽出
- ローカル完結のWindows向けexe

## 使い方の要点

既存ファイルに記載なし

## CLI連携・外部ツール

- 使用する外部ツール: Outlook
- 同梱有無: 既存ファイルに記載なし
- PATH依存: 既存ファイルに記載なし
- エラー時の扱い: 既存ファイルに記載なし
- Codex作業時の注意: CLIや外部ツール情報を追加する場合は、README等の派生ビューだけでなくこのORIGINALへ戻す。

## 対応形式・非対応形式

- 対応入力: 既存ファイル内の関連語: .msg, CSV, メール
- 対応出力: 既存ファイルに記載なし
- 非対応: 既存ファイルに記載なし
- 注意: 形式を推測で追加しない。既存READMEまたは実装確認後に更新する。

## やらないこと / 非ゴール

- メール送信なし
- 本ソフトウェアの無断転載・再配布を禁止します。
- 本ソフトウェアの無断転載・再配布を禁止します

## 設定・ログ・保存方針

- Outlookから保存した `.msg` メールをドロップすると、会社名・お名前・メールアドレスだけをCSVにする小さなWindows向けアプリです。
- 1. Outlookからメールを `.msg` 形式で保存します。
- 3. デスクトップ、または選択した保存先に `mail_list_YYYYMMDD_HHMMSS.csv` を保存します。
- CSVはUTF-8 with BOMで保存します。列は次の3列だけです。
- "launcher_description": "Outlookから保存したメールをCSVリストにします。",

## 非破壊・上書き禁止方針

- 大切なファイルは事前にバックアップを推奨します。
- 大切なファイルは事前にバックアップを推奨します

## 公開用説明の元情報

- display_name: Dakeメールリスト
- site_title: Dakeメールリスト
- site_description: Outlookの.msgメールから、会社名・お名前・メールアドレスだけをCSVにする小さな実務ツールです。
- launcher_description: Outlookから保存したメールをCSVリストにします。
- update_summary: Outlook .msg メールから3列CSVを作成します。
- BOOTH紹介文: Outlookの.msgメールから、
会社名・お名前・メールアドレスをCSVに整理します。

・.msg からCSV整理
・メール送信なし
・会社名、お名前、メールアドレスだけ抽出
・ローカル完結
・Windows向け

集めたメールを、
実務で扱いやすい一覧にします。

## README生成用情報

- 概要: Outlookの.msgメールから、会社名・お名前・メールアドレスだけをCSVにする小さな実務ツールです。
- 使い方: 既存ファイルに記載なし
- 必要なもの: Windows環境
- 注意: READMEへ出す内容は公開可能な情報に限定する。
- ビルド: 既存READMEまたはbuild.batを参照。

## DAKE_META生成用情報

```json
{
  "app_key": "dake_mail_list",
  "display_name": "Dakeメールリスト",
  "launcher_title": "メールリスト作成",
  "launcher_description": "Outlookから保存したメールをCSVリストにします。",
  "site_title": "Dakeメールリスト",
  "site_description": "Outlookの.msgメールから、会社名・お名前・メールアドレスだけをCSVにする小さな実務ツールです。",
  "update_summary": "Outlook .msg メールから3列CSVを作成します。",
  "folder_name": "DAKE_Mail_List",
  "exe_name": "DakeMail_List.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Mail_List_v1.0.0",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## release_body生成用情報

- Dakeメールリスト
- Outlookの.msgメールをCSV化
- 会社名・お名前・メールアドレスだけ抽出
- ローカル完結のWindows向けexe

## booth_product生成用情報

- 商品名: Dakeメールリスト
- 価格案: 300円
- 商品紹介文: Outlookの.msgメールから、
会社名・お名前・メールアドレスをCSVに整理します。

・.msg からCSV整理
・メール送信なし
・会社名、お名前、メールアドレスだけ抽出
・ローカル完結
・Windows向け

集めたメールを、
実務で扱いやすい一覧にします。
- タグ: メール
Outlook
CSV
Windows
実務
ツール
仕事効率化
軽量
シンプル
- 商品画像: assets/booth_thumbnail.jpg
- 補助画像: assets/screenshot.jpg
- 作品ファイル: booth_ready/DakeMail_List.zip
- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Mail_List_v1.0.0
- BOOTH URL: https://peakheadz.booth.pm/items/8448014

## Store表示用情報

- 商品名: Dakeメールリスト
- キャッチ: Outlookの.msgメールから、会社名・お名前・メールアドレスだけをCSVにする小さな実務ツールです。
- 説明: Outlookの.msgメールから、
会社名・お名前・メールアドレスをCSVに整理します。

・.msg からCSV整理
・メール送信なし
・会社名、お名前、メールアドレスだけ抽出
・ローカル完結
・Windows向け

集めたメールを、
実務で扱いやすい一覧にします。
- 価格: 300円
- 画像: assets/booth_thumbnail.jpg（存在: あり）
- ダウンロード導線: 未確定
- サポート方針: 既存ファイルに記載なし

## 価格・販売方針

- BOOTH価格案: 300円
- BOOTH URL: https://peakheadz.booth.pm/items/8448014
- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Mail_List_v1.0.0
- Store販売: 未確定

## 配布・ダウンロード方針

- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Mail_List_v1.0.0
- BOOTH: https://peakheadz.booth.pm/items/8448014
- BOOTH配布zip: booth_ready/DakeMail_List.zip
- Store配布導線: 未確定

## 免責・注意事項

- 【注意事項】
- Windows向けアプリです
- ご利用は自己責任でお願いいたします
- 大切なファイルは事前にバックアップを推奨します
- 本ソフトウェアの無断転載・再配布を禁止します
- 環境によっては起動時にWindowsの警告が表示される場合があります
- https://peakheadz.com

## 同梱ファイル方針

- exe: DakeMail_List.exe
- README.txt: booth_ready/README.txt（存在: あり）
- 注意事項.txt: booth_ready/注意事項.txt（存在: あり）
- 入れないもの: build/、dist/、*.spec、設定ファイル、個人データ、ソース一式は正式配布zipへ混ぜない。

## スクリーンショット・画像方針

- assets/screenshot.webp: あり
- assets/screenshot.jpg: あり
- assets/booth_thumbnail.jpg: あり
- Store用画像: 未確定。既存画像を元に派生する想定。

## 今後の改善予定

現時点では未設定です。

## Codex作業時の注意

- 触ってよい: ORIGINAL.md、および明示された派生ビュー。
- 触らない: アプリ本体、README、release_body、booth_product、booth_ready、assets、distは今回の横展開では変更しない。
- 外部公開しない: 未確定のStore URLや未確認の販売導線を確定情報として書かない。
- 自動操作しない: BOOTH公開、GitHub Release更新、Store公開はこのファイル作成では行わない。

## 派生物一覧

- README.md: GitHub公開用ビュー（存在: あり）
- DAKE_META: README内の機械利用ビュー（存在: あり）
- release_body.md: GitHub Release用ビュー（存在: あり）
- booth_product.txt: BOOTH登録用ビュー（存在: あり）
- booth_ready/booth_product.txt: BOOTH登録時の実使用ビュー（存在: あり）
- Store: 自社販売ビュー。Store専用正本は作らず、このORIGINAL由来の情報を使う。

## 参照した既存ファイル

- README.md: あり
- README内DAKE_META: あり
- release_body.md: あり
- booth_product.txt: あり
- booth_ready/booth_product.txt: あり
- booth_ready/README.txt: あり
- booth_ready/注意事項.txt: あり
- assets/screenshot.webp: あり
- assets/booth_thumbnail.jpg: あり
