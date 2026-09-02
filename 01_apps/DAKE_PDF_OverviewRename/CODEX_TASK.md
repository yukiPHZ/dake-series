# CODEX TASK — DakePDF俯瞰名前変更 Phase 1

## 役割

このファイルは一時的な実装指示書です。真の正本は `ORIGINAL.md` です。重複説明や矛盾がある場合は、必ず `ORIGINAL.md` を優先してください。

## 対象

- repo: `yukiPHZ/dake-series`
- branch: `codex/dake-pdf-overview-rename`
- target: `01_apps/DAKE_PDF_OverviewRename/`
- tracking issue: `#12`

mainへmergeしないでください。

旧ブランチ `feature/dake-pdf-overview-rename` とDraft PR #11は、ChatGPTが先行作成した未受入の暫定実装です。自動でcherry-pickせず、必要な場合のみ参考として読んでください。

## 作業前に読む

1. `01_apps/DAKE_PDF_OverviewRename/ORIGINAL.md`
2. `00_core/CHATGPT_CODEX_WORKFLOW.md`
3. `00_core/DAKE_COMMON_SPEC.md`
4. `00_core/DAKE_FORBIDDEN_RULE.md`
5. `00_core/DAKE_GIT_RULE.md`
6. `00_core/DAKE_ICON_RULE.md`
7. `00_core/DAKE_INTERACTION_QUALITY_V1.md`
8. `00_core/DAKE_LATEST_JOB_PATTERN_V1.md`
9. `00_core/DAKE_BUILD_RULE.md`
10. 必要に応じて `DAKE_PDF_Viewer` と `DAKE_PDF_Merge`

開始前に `git status` と `git log --oneline -5` を確認してください。

## Phase 1で行うこと

### 1. 実装

ORIGINAL.mdに従ってWindows向けアプリを実装してください。

想定ファイル:

```text
01_apps/DAKE_PDF_OverviewRename/
├─ ORIGINAL.md
├─ CODEX_TASK.md
├─ main.py
├─ rename_core.py        # 必要なら。純粋ロジックをテスト可能にする
├─ build.bat
├─ requirements.txt
├─ .gitignore
└─ tests/
   └─ test_rename_core.py
```

ファイル分割は目的ではありません。安全な名前変更ロジックをGUIなしでテストできる最小構成にしてください。

### 2. UI絶対条件

- ヘッダーは機能タイトル＋機能説明の横並び
- ヘッダーに `シンプルそれDAKEシリーズ` を出さない
- ブランド表記はフッターのみ
- フッターにブランド、キャッチコピー、戸建買取査定、Instagram、copyright
- 共通 `02_assets/dake_icon.ico` をウインドウ・exe・タスクバーへ適用
- 日本語UI文言はすべて `UI_TEXT` 管理
- 表ではなくサムネイルカードの俯瞰を主役にする
- 100%、125%、150%表示倍率でヘッダー・ツールバー・フッターを確認

### 3. 応答性

- PDFカード枠と名前を先に表示
- PDF renderはUI threadで行わない
- workerからTkを触らない
- 古いgenerationの結果を表示しない
- フォルダ変更・リフレッシュ・サイズ変更で旧ジョブを蓄積しない
- 表示サイズ変更だけで全PDFをPDFium再レンダリングしない
- 大プレビューはLatest Job方式を候補にする
- リネーム開始時はPDFレンダラーの対象ファイルハンドルが閉じるまで同期する
- リネームとUndoはUI threadを長時間塞がない
- 終了時にworker、executor、after、timerを停止

### 4. 安全な名前変更

ORIGINAL.mdの全検証条件を実装してください。特に次を落とさないでください。

- 大文字小文字だけの変更
- `CON.txt`、`LPT1.memo` 等の予約名派生
- 2件入れ替え
- 3件循環
- 変更対象外ファイルとの衝突
- 読み込み後の元ファイル変更
- 2段階リネーム
- 途中失敗時ロールバック
- Undo前の全件事前検証
- 上書き禁止

入力値の末尾空白を `strip()` で黙って消し、成功扱いにしないでください。

### 5. 未反映変更の保護

変更待ちがある状態で、別フォルダ選択、リフレッシュ、終了を行う場合は破棄確認を出してください。

表示サイズ変更では入力、変更待ち、スクロール位置を保持してください。

## テスト

### 自動テスト

ORIGINAL.mdのテスト一覧を満たしてください。最低限:

- normal / swap / 3-cycle / case-only
- Japanese names
- invalid characters / control characters / trailing dot-space
- reserved names including suffix variants
- case-insensitive duplicate
- collision with non-target
- too-long name
- missing or externally changed source
- injected failure rollback
- Undo success and Undo collision abort

テストがOS依存の場合は、Windows限定テストと非Windowsでも実行できるテストを分けてください。

### 合成PDF試験

機密情報を含まない合成PDFで、1件、48件、100件を通常試験、300件をストレス試験してください。

実フォルダへアクセスできない環境で、菊田の実データ48件を確認済みと報告しないでください。

### 静的確認

- Python構文
- import
- UI_TEXT外の日本語UI直書き
- Git除外対象混入
- 個人パス・機密情報混入
- sourceとORIGINALの不一致

### Windows build

Windows環境で可能なら:

1. `build.bat`
2. `dist/DakePDF_OverviewRename.exe` 生成
3. exe起動
4. 文字化け
5. window icon
6. taskbar icon
7. 基本操作smoke test

できなかった項目は未確認と明記し、成功扱いにしないでください。

## Git

- 仕様変更と実装は可能ならcommitを分ける
- `git add .` を安易に使わない
- build、dist、spec、exe、ローカル設定をcommitしない
- commit / pushは `codex/dake-pdf-overview-rename` のみ
- mainへmergeしない
- Release、BOOTH、dakeapp.com、Store、Cloudflareへ公開しない

## 完了報告

次を報告してください。

- 読んだ正本・共通仕様
- 変更ファイル
- 実装概要
- 自動テスト結果
- 1 / 48 / 100 / 300件の結果
- main thread、worker、queue、generation、キャッシュ方針
- Windows build / exe / icon確認結果
- 未確認事項と残課題
- commit SHA
- push先ブランチ
- final git status

実装完了後も、人間の実データ48件受入と画面確認が終わるまで正式出荷へ進めないでください。
