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
- version: 1.0.1
- price: 500円
- distribution: GitHub Release / BOOTH / dakeapp.com（Storeは未確定）
- target_platform: Windows

## 目的

Outlookで選択したメールを直接、または保存済みの `.msg` メールをドラッグ＆ドロップし、会社名・お名前・メールアドレスだけをCSVにする小さな実務ツールです。

## 守るべき挙動

- Outlookで選択した1通または複数メールを、アプリへ直接ドラッグ＆ドロップして処理できる。
- 保存済みの1通または複数 `.msg` も同じドロップ領域で処理できる。
- 出力CSVはUTF-8 with BOMで、`会社名`、`お名前`、`メールアドレス` の3列だけとする。
- 1件の解析失敗で、ほかのメールの処理を止めない。
- 処理はローカルで完結し、メール内容や利用履歴を外部送信しない。

Outlook直接ドラッグ＆ドロップ、複数メール、CSV生成、3項目抽出は会社実機で確認済みです。

## 変更してはいけない核

- Outlookから直接メールを落とす操作を、保存操作や別ボタン操作へ置き換えない。
- CSV列を追加・変更しない。
- メールアドレスのlocal-partから氏名を推測しない。
- Outlook COM取得ボタン、監視、履歴、名簿、CRM、AI/API、クラウド機能を追加しない。
- 会社実機確認済みのWindows native OLE入力層は、必要な不具合修正がない限り変更しない。

## 対象ユーザー

Outlookで受信したメールから、会社名・お名前・メールアドレスの一覧を短い操作で作成したいWindows利用者。

## 解決する困りごと

- Outlookメールを一度保存することなく、直接ドラッグして3列CSVを作成します。
- 複数メールをまとめて処理し、手作業での転記を減らします。

## 主な機能

- Outlookで選択した1通または複数メールの直接ドラッグ＆ドロップ
- 保存済みの複数 `.msg` ファイルのドラッグ＆ドロップ
- 会社名・お名前・メールアドレスだけ抽出
- ローカル完結のWindows向けexe

## 使い方の要点

1. Outlookで1通または複数のメールを選択します。保存済みの `.msg` ファイルも使用できます。
2. Dakeメールリストの「メールをここへドロップ」へ直接ドラッグ＆ドロップします。
3. デスクトップ、または選択した保存先に `mail_list_YYYYMMDD_HHMMSS.csv` を保存します。

## CLI連携・外部ツール

- 使用する外部ツール: Outlook（直接ドラッグ元として利用）
- PATH依存: なし
- 外部AI/API/クラウド通信: なし

## 対応形式・非対応形式

- 対応入力: Outlookで選択したメールのWindows OLE仮想 `.msg`、保存済みの `.msg` ファイル。どちらも複数同時ドロップに対応する。
- 対応出力: UTF-8 with BOMのCSV。列は `会社名`、`お名前`、`メールアドレス` の3列だけとする。
- 非対応: `.msg` 以外のファイル。対象外ファイルは処理せず、アプリを終了させない。
- Outlook仮想メールはWindows native OLE `IDropTarget` で `FileGroupDescriptorW` と、各アイテムの `lIndex` を指定した `FileContents` を取得する。
- 保存済みファイルは同じnative OLE入力層で `CF_HDROP` として受け取る。tkinterdnd2との二重登録は行わない。

## やらないこと / 非ゴール

- メール送信、受信監視、自動返信
- Outlook COMで選択メールを取得する別ボタン
- メール一覧、名簿、履歴、CRM、電話番号抽出
- AI/API、クラウド解析、外部保存
- CSV編集画面、設定画面、学習機能

## 設定・ログ・保存方針

- 主操作は、Outlookで1通または複数メールを選び、Dakeメールリストへ直接ドラッグ＆ドロップすることです。
- 保存済みの `.msg` も補助入力として利用できます。
- デスクトップ、または選択した保存先に `mail_list_YYYYMMDD_HHMMSS.csv` を保存します。
- CSVはUTF-8 with BOMで保存し、列は `会社名`、`お名前`、`メールアドレス` の3列だけです。
- Outlook仮想メールは標準ランタイムTEMP内の一時 `.msg` として既存解析へ渡し、処理後に削除する。永続キャッシュ・履歴は作らない。
- 実機確認用の最小診断は `%LOCALAPPDATA%\PEAKHEADZ\DakeMailList\logs\ole_dnd.log` に起動時上書き保存し、本文・アドレス・氏名・会社名・件名全文を記録しない。
- CSV保存先設定は `%LOCALAPPDATA%\PEAKHEADZ\DakeMailList\config.json` に `save_folder` だけを保存する。exe横の旧 `Dake_Mail_List_config.json` は新configが未作成の場合だけ一度読み取り、新しい保存先へ移行する。旧ファイルは削除しない。
- お名前はMSGの送信者表示名、Fromヘッダー、現在の送信者側の署名の順に取得する。メールアドレスのlocal-partから氏名を推測せず、根拠がなければ空欄にする。
- 会社名と署名由来のお名前は返信・転送履歴より前にある現在送信者側の署名を優先する。同一行に会社名、部署・役職、氏名がある場合は、根拠のある範囲だけを保守的に分離する。
- HTML本文は改行、段落、表セルの境界を保持して署名を解析し、Outlook等の引用コンテナ内は現在送信者の署名候補に含めない。
- native OLE `Drop()` 内では仮想 `.msg` の実体化とthread-safe queueへの投入だけを行い、Tkinter UI更新と解析worker開始はTkのpoll処理へ戻してから行う。
- `FileContents` は各アイテムの `lIndex` を指定し、`TYMED_HGLOBAL | TYMED_ISTREAM | TYMED_ISTORAGE` をまとめて要求してOutlook側に返却mediumを選択させる。返された実mediumが `IStorage` の場合は、pywin32の `IStorage.CopyTo` で一時 `.msg` のcompound storageへ保存する。

## 非破壊・上書き禁止方針

- 大切なファイルは事前にバックアップを推奨します。
- 大切なファイルは事前にバックアップを推奨します

## 公開用説明の元情報

- display_name: Dakeメールリスト
- site_title: Dakeメールリスト
- site_description: Outlookで選択した1通または複数メールを直接ドロップし、会社名・お名前・メールアドレスだけをCSVにする小さな実務ツールです。保存済みの.msgにも対応します。
- launcher_description: Outlookのメールを直接ドロップしてCSVリストにします。
- update_summary: Outlookメールの直接ドラッグ＆ドロップと複数メール処理に対応しました。
- BOOTH紹介文: Outlookで選択したメールを直接ドロップして、
会社名・お名前・メールアドレスをCSVに整理します。

・Outlookから直接ドラッグ
・複数メールをまとめてCSV整理
・保存済み.msgにも対応
・メール送信なし
・会社名、お名前、メールアドレスだけ抽出
・ローカル完結
・Windows向け

集めたメールを、
実務で扱いやすい一覧にします。

## README生成用情報

- 概要: Outlookで選択した1通または複数メールを直接ドロップし、会社名・お名前・メールアドレスだけをCSVにする小さな実務ツールです。保存済みの.msgにも対応します。
- 使い方: Outlookでメールを1通または複数選択し、アプリへ直接ドラッグ＆ドロップする。保存済み.msgも同じ操作で処理できる。
- 必要なもの: Windows環境
- 注意: READMEへ出す内容は公開可能な情報に限定する。
- ビルド: 既存READMEまたはbuild.batを参照。

## DAKE_META生成用情報

```json
{
  "app_key": "dake_mail_list",
  "display_name": "Dakeメールリスト",
  "version": "1.0.1",
  "launcher_title": "メールリスト作成",
  "launcher_description": "Outlookのメールを直接ドロップしてCSVリストにします。",
  "site_title": "Dakeメールリスト",
  "site_description": "Outlookで選択した1通または複数メールを直接ドロップし、会社名・お名前・メールアドレスだけをCSVにする小さな実務ツールです。保存済みの.msgにも対応します。",
  "update_summary": "Outlookメールの直接ドラッグ＆ドロップと複数メール処理に対応しました。",
  "folder_name": "DAKE_Mail_List",
  "exe_name": "DakeMail_List.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Mail_List_v1.0.1",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true,
  "demo_video_path": "release_artifacts/demo.mp4",
  "demo_video_url": "",
  "social_release_path": "release_artifacts/social_release.json"
}
```

## release_body生成用情報

- Dakeメールリスト
- Outlookで選択した1通または複数メールを直接ドラッグ＆ドロップ
- 保存済みの複数.msgにも対応
- 現在の送信者側の署名を優先して会社名とお名前を抽出
- 会社名・お名前・メールアドレスだけ抽出
- UTF-8 BOM、3列固定のCSV
- ローカル完結のWindows向けexe

## booth_product生成用情報

- 商品名: Dakeメールリスト
- 価格案: 500円
- 商品紹介文: Outlookで選択したメールを直接ドロップして、
会社名・お名前・メールアドレスをCSVに整理します。

・Outlookから直接ドラッグ
・複数メールをまとめてCSV整理
・保存済み.msgにも対応
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
- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Mail_List_v1.0.1
- BOOTH URL: https://peakheadz.booth.pm/items/8448014

## Store表示用情報

- 商品名: Dakeメールリスト
- キャッチ: Outlookで選択したメールを直接ドロップし、会社名・お名前・メールアドレスだけをCSVにする小さな実務ツールです。
- 説明: Outlookで選択したメールを直接ドロップして、
会社名・お名前・メールアドレスをCSVに整理します。

・Outlookから直接ドラッグ
・複数メールをまとめてCSV整理
・保存済み.msgにも対応
・メール送信なし
・会社名、お名前、メールアドレスだけ抽出
・ローカル完結
・Windows向け

集めたメールを、
実務で扱いやすい一覧にします。
- 価格: 500円
- 画像: assets/booth_thumbnail.jpg（存在: あり）
- ダウンロード導線: 未確定
- サポート方針: 既存ファイルに記載なし
- Stripe Payment Link: https://buy.stripe.com/dRm6oHdEdcwv1q80bf0gw16
- Store雋ｩ螢ｲ迥ｶ諷・ stripe_ready

## 価格・販売方針

- BOOTH価格案: 500円
- BOOTH URL: https://peakheadz.booth.pm/items/8448014
- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Mail_List_v1.0.1
- Store販売: 未確定

## 配布・ダウンロード方針

- GitHub Release: https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Mail_List_v1.0.1
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

## 作業時の注意

- 製品固有の正本はこのORIGINALとし、README等の派生ビューはここから整合させる。
- PEAKHEADZ Standardの共通仕様・toolは製品配下へコピーしない。
- `AI_CONTEXT.md` は共有generatorから生成し、直接編集しない。
- 外部公開しない: 未確定のStore URLや未確認の販売導線を確定情報として書かない。
- BOOTH、Store、Stripe等は明示された作業範囲だけを扱う。

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
- payment_status: stripe_ready
