# Current Status Review — DakePDF俯瞰名前変更

Date: 2026-09-02

## 結論

企画は採用を維持する。

「PDFを1件ずつ開かず、1ページ目サムネイルを俯瞰しながら、PDFごとに名前を付けて変更分だけ反映する」という中心価値は明確で、既存 `DAKE_PDF_Rename` とも役割が分離できている。

一方、旧開発ブランチはChatGPTがCodex工程より先に暫定コードを置いたため、正式な実装開始点としては採用しない。クリーンなCodex用ブランチからやり直す。

## 正しい現在地

- 正本仕様: `01_apps/DAKE_PDF_OverviewRename/ORIGINAL.md`
- Codex指示: `01_apps/DAKE_PDF_OverviewRename/CODEX_TASK.md`
- 現作業ブランチ: `codex/dake-pdf-overview-rename`
- Codex実装: 未着手
- 自動テスト: 未実施
- Windows exe build: 未実施
- 菊田実データ48件確認: 未実施
- 正式出荷: 未実施

## 旧ブランチの扱い

- 旧ブランチ: `feature/dake-pdf-overview-rename`
- 旧Draft PR: #11
- 状態: 暫定参考。未受入。自動cherry-pick禁止。

旧ブランチには `main.py`、`build.bat` 等があるが、実行記録、テストコード、CI、Windows build結果、受入記録がない。したがって、旧ブランチ上の性能・動作については確認済みと扱わない。

## セルフレビューで修正した点

### 1. ChatGPT / Codexの役割分離

ChatGPTは設計、判断、正本、Codex指示書を担当する。実装、テスト、commit、pushはCodexへ戻す。

### 2. 指示書の重複削減

旧 `CODEX_TASK.md` は正本と内容が重複しすぎていた。新指示書は作業手順と完了条件へ絞り、仕様はORIGINALへ一本化した。

### 3. statusモデル修正

旧正本の `status: development` はDAKE共通statusモデル外だった。新正本では `status: draft`、`app_type: market`、`completion_goal: formal_release` とした。

### 4. 非現実的な大量件数要求の整理

v1の中心は実フォルダ相当の48件。100件を通常利用目安、300件をストレス試験とし、500件対応は出荷条件から外した。

### 5. 未反映入力の保護

別フォルダ、リフレッシュ、終了時に、変更待ち入力を無警告で失わない仕様を追加した。

### 6. レンダリングとリネームの競合防止

PDFレンダラーが対象PDFを開いたままリネームしないよう、旧レンダージョブ停止とファイルハンドル解放を確認してから書き込みを始める条件を追加した。

### 7. 表示サイズ切替の軽量化

小・標準・大の切替だけで全PDFをPDFium再レンダリングしない。入力、変更待ち、スクロール位置も保持する。

### 8. リネーム検証の強化

- case-only rename
- `CON.txt` 等の予約名派生
- 3件循環
- 元ファイルの外部変更
- 長すぎる名前
- Undo衝突
- 故障注入によるロールバック

を明示的なテスト対象にした。

### 9. 実データと合成データの区別

Codex環境の合成PDF試験と、菊田の実フォルダ48件受入を分けた。実フォルダへアクセスできない環境では実データ確認済みと報告しない。

## 次工程

Codexは `codex/dake-pdf-overview-rename` ブランチで、ORIGINAL.mdを正本としてPhase 1を実装・テスト・build確認し、commit / pushする。

人間側はCodex完了後に、実フォルダ48件で次を確認する。

- 俯瞰しやすさ
- ヘッダー横並び
- フッターとリンク
- ウインドウ・タスクバーの共通アイコン
- 名前入力の速さ
- スクロール感
- 実際の一括変更とUndo

この受入が終わるまでRelease、BOOTH、dakeapp.com、Store、Cloudflareへ進めない。
