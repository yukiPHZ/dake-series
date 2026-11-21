# Dake名前フォルダ中身見ながら変更 — ORIGINAL

## 正本宣言
このファイルが製品の正本。README、DAKE_META、release_body、BOOTH、Storeは派生ビュー。

## 基本情報
- app_id: DAKE_Folder_OverviewRename
- version: 1.0.0
- target_platform: Windows 10/11 デスクトップ
- status: draft / Human Review HOLD
- price: 未決定
- distribution: 候補exeのみ。正式公開はHuman Review合格後。

## 目的・対象ユーザー
新人営業がデスクトップに作った「新しいフォルダー」などを、中の資料を見て分かりやすく改名する。既存DakePDF俯瞰名前変更とは別アプリ。

## 操作と機能
1. デスクトップなど親フォルダを選択するか、親フォルダ専用領域へドロップ。
2. 直下の子フォルダを一覧表示。直接対象を入れる場合は複数フォルダ専用ドロップ領域を使用。
3. フォルダを選ぶと全対象ファイルをスクロールで閲覧。PDF先頭ページとページ数、画像、元名、容量、子フォルダを表示。
4. 変更予定名を入力。「次のフォルダへ」で連続整理。入力・チェックはモデルへ保持。
5. 「変更内容を確認」で旧名・新名を確認してから実行。直前の成功分はUndo可能。

## 対応形式
PDF（先頭ページ）、JPG/JPEG、PNG、WebP、BMP、TIFF。Word・Excel等は種類表示。破損・暗号化・非対応・読取不可も表示から消さない。子フォルダは存在を表示し、内部を走査しない。

## 安全設計
フォルダ名のみを同じ親の中で変更する。内部ファイルの改名、変更、削除、移動、複製、上書きはしない。AI・API・自動命名なし。
シンボリックリンク・ジャンクションを含むreparse pointは祖先も含め拒否。無制限再帰なし。
名前の不正文字、予約名、末尾空白・ピリオド、UTF-16長、同名・大文字小文字衝突、重複対象、親子同時指定を検証。
入替と大文字小文字だけの改名は初期版で拒否し、一時名を使わない。
実行直前にID・親ID・直下内容のスナップショットを再検証。Windowsハンドルで対象と祖先の削除共有を禁止して固定し、SetFileInformationByHandleで上書き禁止rename。
途中失敗は即停止。成功分・失敗分・未実行を区別し、成功済み分だけUndo候補にする。自動rollbackしない。
Undoも同一性・直下内容・衝突を検証。再帰的内容監視はしない。外部アプリによる子孫ファイルの編集を禁止・巻き戻しはしない。

## 履歴と復旧
実行時にLOCALAPPDATA/DAKE_Folder_OverviewRename/historyへUUID別JSONLを追記しflush/fsync。対象フォルダ内には書き込まない。
intentを先に永続化し、実行後successまたはfailureを記録。書込失敗前は改名しない。改名後ログ失敗は成功した改名と区別して表示しUndoを保持。
Undoはセッション内の直前一括変更の成功分。再起動後は自動再実行しない。
強制終了等でintentだけ残った場合はsource/destinationと記録identityを確認し、どちらに実体があるか確定するまで再実行しない。新旧両方が存在する場合は手動確認する。ログはローカルパスを含むため公開しない。

## 応答性
読み取り専用scan/renderは各固定1 worker、pending最新1要求、結果queue各1。世代が変わった結果は破棄。
表示は可視範囲を優先し最大32カードを再利用。画像LRUは64枚、1枚最大300x220 RGB。Tk画像は表示カード分のみ。
ファイル一覧メタデータは全件保持。ディスク走査、PDF解析、画像縮小、名前変更検証・実行はUIスレッド外。
0.5秒以上のUI無反応は改善対象。実測とHuman Reviewは別々に記録する。

## ブランド
白・ライトグレー、BIZ UDPGothic優先、共通dake_icon.ico、文言はUI_TEXT。共通フッター。アニメーションなし。

## 非ゴール
ファイル編集、再帰検索、AI判定、API連携、整理先への移動、クラウド同期、再起動後の自動復旧。

## 開発・検証
合成データのみ。0/1/10/30/100フォルダ、空、混在、連番、破損・暗号化PDF、日本語・長名、数百ファイル、切替、入力保持、部分失敗、Undo、外部変更、drop、候補exeを検証する。
正式公開とHuman Reviewは別ゲート。現時点でHuman Reviewは未実施。

## 配布・販売
GitHub Release・BOOTH・dakeapp.com・Store・Cloudflare反映はHOLD。価格、BOOTH URL、Release URLは未確定。
合格後に正式出荷ラインに従う。Store生成物は手編集しない。
同梱候補: 単一exe、説明、注意事項、依存ライセンス。exe/build/distはGitへ入れない。

## README生成用情報
親選択→内容確認→入力→確認→改名。python main.pyで起動、python build.pyでonefile候補作成。

## DAKE_META生成用情報

```json
{
  "app_key": "dake_folder_overviewrename",
  "display_name": "Dake名前フォルダ中身見ながら変更",
  "launcher_title": "中身を見てフォルダ名変更",
  "launcher_description": "PDF・画像を見ながらフォルダ名をまとめて変更。",
  "site_title": "Dake名前フォルダ中身見ながら変更",
  "site_description": "フォルダを一つずつ開かず、中身を見て名前を整理するWindowsアプリ。",
  "update_summary": "v1.0.0 初期候補。Human Review待ち。",
  "folder_name": "DAKE_Folder_OverviewRename",
  "exe_name": "DakeFolder_OverviewRename.exe",
  "version": "1.0.0",
  "release_url": "",
  "app_type": "market",
  "completion_goal": "formal_release",
  "screenshot_path": "assets/screenshot.webp",
  "status": "draft",
  "show_in_launcher": false,
  "show_on_site": false
}
```

## release_body生成用情報
- PDF・画像の中身を見ながらフォルダ名を入力。
- フォルダの一括名前変更と検証付きUndo。
- フォルダ内ファイルは閲覧のみ。
- Windows向け単一exe。Human Review待ち・公開HOLD。

## BOOTH・Store表示用情報
商品名は表示名と同じ。説明は概要から派生。価格未定、掲載HOLD、payment_statusはpreparing想定（外部反映なし）。

## 検証記録
候補検証を以下に記録。正式公開の完了記録ではない。

### 2026-10-16 候補検証結果
- 安全性・PDF自動試験: 41 passed。テスト依存reportlabからPython将来互換のDeprecationWarning 1件。
- 合成対象: 0/1/10/30/100フォルダ、空、連番、混在、破損・暗号化PDF、日本語長名、共有ロック、権限エラー注入、途中失敗、Undo、外部置換、ジャンクション。
- 30フォルダの改名→Undoで全ファイルSHA-256の集合が一致。実データの破壊的試験なし。
- Windows no-replace競合試験: 検証直後に作った衝突先を上書きせず拒否。
- GUI: 30フォルダ、406項目。末尾到達、別フォルダ切替、入力・チェック保持を確認。callback errorなし。
- UI heartbeat最大間隔: 0.1119秒（この合成試験の計測値。全環境保証ではない）。表示カード最大3枚、実装上限32。
- ドロップ: tkinterdnd2組込み、2領域の登録、Tclの複数パス分割と直接読込み経路を検証。エクスプローラー実ドラッグの最終受入は未実施。
- Windows候補exe実画面: 起動、アイコン、日本語、共通フッター、親選択ダイアログ、30フォルダ読込み、画像、正常PDF先頭ページ・2ページ表示、スクロール、ペイン幅変更を確認。
- 候補: dist/20261016-213620/DakeFolder_OverviewRename.exe（onefile / noconsole / clean、FileVersion 1.0.0）。
- 候補SHA-256: `2ebb180db8904b9a325f19399b86141fae97f975d4e2a5c08f54a8a53a1176d2`
- Human Review: 未実施。日本語IME、実エクスプローラー単一親・複数直接drop、10件以上の連続整理、候補exeで改名とUndoの操作感、高DPI・狭い画面を確認する。
- GitHub Release / BOOTH / dakeapp.com / Store / Cloudflare: 未実施・HOLD。
- 旧候補distは履歴として保持。レビュー対象は上記候補のみ。
