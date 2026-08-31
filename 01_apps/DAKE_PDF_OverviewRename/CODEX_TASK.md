# CODEX TASK — DakePDF俯瞰名前変更 Phase 1

## 0. この文書の位置づけ

この文書は Codex 実装指示書です。

ChatGPT は設計・判断・指示書作成を担当し、Codex は実装・修正・調査・commit・push・レポートを担当します。

真の正本はこのファイルではなく、必ず次です。

`01_apps/DAKE_PDF_OverviewRename/ORIGINAL.md`

この `CODEX_TASK.md` と `ORIGINAL.md` が矛盾した場合は、`ORIGINAL.md` を優先してください。

---

## 1. 対象

- repo: `yukiPHZ/dake-series`
- branch: `feature/dake-pdf-overview-rename`
- target folder: `01_apps/DAKE_PDF_OverviewRename/`
- Draft PR: `#11`
- tracking issue: `#12`

作業はこの feature branch 上で行ってください。
main へ merge しないでください。

---

## 2. 作業前に必ず読む

次の順で確認してください。

1. `01_apps/DAKE_PDF_OverviewRename/ORIGINAL.md`
2. `00_core/CHATGPT_CODEX_WORKFLOW.md`
3. `00_core/DAKE_COMMON_SPEC.md`
4. `00_core/DAKE_FORBIDDEN_RULE.md`
5. `00_core/DAKE_GIT_RULE.md`
6. `00_core/DAKE_ICON_RULE.md`
7. `00_core/DAKE_INTERACTION_QUALITY_V1.md`
8. `00_core/DAKE_LATEST_JOB_PATTERN_V1.md`
9. `00_core/DAKE_BUILD_RULE.md`
10. 必要に応じて既存PDFアプリ

参考実装として見るもの：

- `01_apps/DAKE_PDF_Viewer/main.py`
  - PDF表示
  - worker / queue / generation
  - Latest Job
  - 終了処理
- `01_apps/DAKE_PDF_Merge/main.py`
  - ヘッダー横並び
  - フッター構造
  - UI_TEXT
  - 共通アイコン
- `01_apps/DAKE_PDF_Merge/build.bat`
  - PyInstaller
  - `dake_icon.ico`
  - version info

作業前に `git status` と `git log --oneline -5` も確認してください。

---

## 3. 現在の main.py の扱い

現在 feature branch にある `main.py` は、ChatGPT 側で先に置かれた暫定・参考実装です。

完成品として信用しないでください。

Codex 自身で、ORIGINAL.md と 00_core の現行仕様に照らして監査し、必要なら大きく修正・再構成してください。

ただし、使える部分をわざわざ捨ててゼロから書き直す必要もありません。

判断基準は「既存コードを守ること」ではなく、

- ORIGINAL に合っているか
- UI が止まらないか
- 実務で迷わないか
- 安全にファイル名を変更できるか

です。

---

## 4. アプリの目的

フォルダ内のPDFを1件ずつ開かなくても、1ページ目のサムネイルを俯瞰しながら、その場でファイル名を変更できるWindows向け単機能アプリを作ります。

従来：

```text
PDFを開く
↓
書類を確認
↓
閉じる
↓
Explorerへ戻る
↓
名前変更
↓
次へ
```

本アプリ：

```text
フォルダを選ぶ
↓
PDFをサムネイルで俯瞰
↓
見ながら名前を入力
↓
変更分だけ一括反映
```

主役は「PDF一覧表」ではなく、書類を机に並べたようなサムネイルカードの俯瞰です。

---

## 5. UI絶対条件

ここは変更しないでください。

### ヘッダー

機能タイトルと機能説明を横並びにします。

- タイトル: `PDFを見ながら名前を変える`
- 説明: `フォルダ内のPDFをサムネイルで一覧表示し、その場で名前を変更します。`

禁止：

- ヘッダーに `シンプルそれDAKEシリーズ` を表示する
- タイトルと説明を意図なく縦積みにする
- ブランドロゴや余計な帯を追加する

### ブランド表記

`シンプルそれDAKEシリーズ` はフッターのみです。

### フッター

必須：

- `シンプルそれDAKEシリーズ / 止まらない、迷わない、すぐ終わる。`
- `戸建買取査定` リンク
- `Instagram` リンク
- `© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta`

広幅時は左右2ブロック。
狭幅時は中央寄せ2段を許可します。
思想タイトル＆キャッチコピーの途中改行は禁止です。

### アイコン

DAKE共通：

`02_assets/dake_icon.ico`

を使用します。

- 開発時 window icon
- exe icon
- PyInstaller `--icon`
- onefile時の `--add-data`
- Windowsタスクバー反映

を確認してください。

### UI文言

日本語UI文言はすべて `UI_TEXT` で管理してください。

`text=""` に日本語を直接書かないでください。

ダイアログ、エラー、状態表示も対象です。

PythonファイルはUTF-8前提で扱ってください。
文字コード確認質問は不要です。

---

## 6. 画面の主構造

概念は次です。

```text
┌──────────────────────────────────────────────────────────────┐
│ PDFを見ながら名前を変える  フォルダ内のPDFをサムネイルで… │
│                                                              │
│ [フォルダを選ぶ] [リフレッシュ]  C:\...                    │
│ 表示サイズ  小  標準  大    [Undo] [名前変更を反映 n]       │
├──────────────────────────────────────────────────────────────┤
│                                                              │
│  [PDFカード] [PDFカード] [PDFカード] [PDFカード]             │
│  [PDFカード] [PDFカード] [PDFカード] [PDFカード]             │
│                                                              │
│                       ↓ scroll                               │
│                                                              │
├──────────────────────────────────────────────────────────────┤
│ サムネイル生成中 18 / 48                                     │
├──────────────────────────────────────────────────────────────┤
│ DAKEフッター                                                  │
└──────────────────────────────────────────────────────────────┘
```

カードは最低限：

- 1ページ目サムネイル
- ページ数
- 元ファイル名
- 新ファイル名入力欄
- `.pdf` 固定表記

を持ちます。

変更されたカードは薄青で識別します。

---

## 7. 必須機能

### フォルダ読み込み

- フォルダ直下の `.pdf` のみ
- サブフォルダ再帰しない
- PDFカード枠とファイル名を先に出す
- 全サムネイル完成までUIを待たせない

### サムネイル

- 1ページ目
- pypdfium2 / PDFium 第一候補
- PillowでTk表示用へ変換
- サイズ切替: 小 / 標準 / 大
- 大量PDFでもUI threadでrenderしない

### 大プレビュー

- サムネイルクリックで1ページ目を大きく表示
- 外部PDFビューア起動を主経路にしない
- 古いpreview結果が新しい選択へ遅れて表示されないよう、必要なら Latest Job / generation を適用

### 名前入力

- `.pdf` は編集不可
- Tabで次の入力欄へ自然に移れること
- 入力しただけでは実ファイルを変更しない
- 変更待ち件数を表示

### 変更反映

`名前変更を反映` で初めて実ファイルを書き換えます。

反映前に全件を事前検証します。

検証対象：

- 空欄
- 同名
- Windows禁止文字 `< > : " / \\ | ? *`
- 末尾のドット
- 末尾の空白
- Windows予約名
- 既存ファイルとの衝突
- アクセス不能

### 2段階リネーム

A.pdf → B.pdf
B.pdf → A.pdf

のような入れ替えを安全に処理できるよう、一時名を挟む2段階処理を行ってください。

一時名は対象フォルダ内で衝突しないことを保証してください。

### ロールバック

処理途中で失敗した場合、可能な範囲で元名へ戻してください。

「途中まで成功したのに画面は成功扱い」は禁止です。

### Undo

直前の一括名前変更だけ1回戻せるようにします。

- Undoは保存/書込系なので Latest Job のように捨ててはいけない
- Undo可能性がなくなったらボタン状態を更新
- Undo自体の失敗も明確に表示

---

## 8. 応答性

DAKE原則：

- 処理は止まってもいい
- UIは止めるな
- 0.5秒以上の無反応を避ける
- 見えている範囲を優先

目標：

- 50件：余裕
- 100件：普通に使える
- 300件：時間がかかってもUIは操作可能
- 500件：メモリを無制限に膨らませない

PDFium呼び出しのスレッド安全性を確認し、危険ならrender呼び出しは直列化してください。

workerからTk Widgetを直接触らないでください。
Tk更新はmain threadへ戻してください。

終了時には worker / executor / timer / after / pending job を適切に止め、アプリ終了後にUI更新しないようにしてください。

---

## 9. エラー分離

1ファイルのプレビュー失敗で、フォルダ全体を失敗にしないでください。

カードだけ：

`プレビューできません`

等の状態にして、ファイル名変更機能は可能なら維持します。

暗号化PDF、壊れたPDF、0ページ等を想定してください。

---

## 10. v1でやらないこと

追加しないでください。

- AI自動命名
- OCR自動命名
- PDF本文検索
- フォルダ自動分類
- サブフォルダ再帰
- PDF編集
- PDF結合
- PDF分割
- ファイル移動
- クラウド同期
- OpenAI API
- 複雑な一括命名ルール
- タグ管理

「便利そうだから」は機能追加理由になりません。

---

## 11. 既存 DAKE_PDF_Rename との分離

`DAKE_PDF_Rename` は、PDF名の前後に任意テキストを一括付加する既存アプリです。

今回のアプリは、PDF内容を見てファイルごとに意味のある名前を付けるアプリです。

既存 `DAKE_PDF_Rename` を置き換えたり統合したりしないでください。

---

## 12. build

想定exe：

`DakePDF_OverviewRename.exe`

PyInstaller onefile / noconsole。

共通アイコン：

`../../02_assets/dake_icon.ico`

build / dist / spec / version_info 等は既存DAKEルールに従い、通常commitへ含めないでください。

Windows build可能な環境なら：

1. build.bat実行
2. exe生成確認
3. exe起動確認
4. window icon確認
5. タスクバーicon確認
6. 基本操作smoke test

まで行ってください。

Windows build不可の環境なら、不可と明記し、成功したことにしないでください。

---

## 13. テスト

最低限、次を確認してください。

### 静的

- Python構文
- import
- UI_TEXT直書き日本語レビュー
- 禁止生成物がgit対象になっていない

### リネーム

一時ディレクトリを使い：

- 1件 rename
- 複数 rename
- A⇄B swap
- A→B / B→C chain
- 重複名拒否
- 既存非対象ファイル衝突拒否
- 禁止文字拒否
- 予約名拒否
- ロールバック確認
- Undo確認

を可能な範囲で自動テストしてください。

### PDF

テストPDFを複数作れるなら：

- 1件
- 20件
- 50件以上
- 壊れたPDF混在

を確認。

UIがrender完了待ちで固まらないことを見てください。

### UI

確認項目：

- ヘッダーが横並び
- ヘッダーにDAKEブランドなし
- フッターにブランド・リンク・copyright
- カードが俯瞰として自然
- 小 / 標準 / 大で崩れない
- 変更待ち薄青
- スクロール自然
- 初期ウインドウサイズが不自然に巨大/狭小でない
- 共通アイコン

---

## 14. Git運用

作業 branch：

`feature/dake-pdf-overview-rename`

mainへmergeしない。

`git add .` を安易に使わず、対象ファイルを明示してcommitしてください。

無関係な既存変更を巻き込まないでください。

commit / push は実施してよいです。

Draft PR #11を更新対象として扱ってください。

Store / BOOTH / Release / dakeapp.com / Cloudflare は今回触らないでください。

---

## 15. 今回は出荷工程ではない

今回は Phase 1 実装・検証です。

次はまだ行わない：

- screenshot.webp正式作成
- booth_thumbnail.jpg
- README公開ビュー最終化
- release_body
- booth_product
- GitHub Release
- BOOTH公開
- dakeapp.com掲載
- Store同期
- Stripe
- Cloudflare確認

Phase 1が実機で良好になってから別工程で行います。

---

## 16. 完了条件

次を満たして Phase 1 完了候補です。

- ORIGINAL.mdを確認した
- 現行 main.py を監査/修正した
- ヘッダー横並び
- ヘッダーにブランド表記なし
- ブランドはフッターのみ
- フッターリンク正常
- 共通dake_icon適用
- UI_TEXT完全管理
- サムネイル俯瞰が成立
- UI threadで重いrenderをしていない
- 一括安全リネームが成立
- swapが成立
- Undoが成立
- 失敗時ロールバック設計が成立
- preview失敗が全体を巻き込まない
- テスト結果を記録した
- feature branchへcommit / pushした
- Draft PR #11はmergeしていない
- final git statusを確認した

---

## 17. Codex完了報告フォーマット

最後に簡潔に以下を報告してください。

```text
Phase 1 実装結果

ORIGINAL.md確認: 済 / 未
変更ファイル:
- ...

実装:
- ...

テスト:
- 構文:
- rename:
- swap:
- undo:
- PDF thumbnail:
- UI:
- Windows exe build:

残課題:
- ...

commit:
- <sha> <message>

push:
- 済 / 未

Draft PR:
- #11 維持 / 状態

final git status:
- ...
```

不明・未検証は未検証と明記してください。
推測で「完了」にしないでください。
