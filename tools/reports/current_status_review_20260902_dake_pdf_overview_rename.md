# Current Status Review — DakePDF俯瞰名前変更

Date: 2026-09-02

## 結論

Phase 1の技術受入と人間受入を引き継ぎ、Phase 2の正式出荷準備を完了した。

`01_apps/DAKE_PDF_OverviewRename/ORIGINAL.md` を真の正本として、公開用派生ビュー、実アプリのスクリーンショット、BOOTH用資材、第三者ライセンス、Windows onefile配布候補を整えた。状態は公開前の `draft` のままとし、価格・公開URLは決めていない。

GitHub Release、BOOTH、dakeapp.com、Store、Stripe、Cloudflareへの公開、およびmainへのmergeは行っていない。

## 正しい現在地

- 正本仕様: `01_apps/DAKE_PDF_OverviewRename/ORIGINAL.md`
- 作業ブランチ: `codex/dake-pdf-overview-rename-release-prep`
- 開始SHA: `55aa87c5302792ba55ece1df2d58ad8e704be608`
- status: `draft`
- app_type: `market`
- completion_goal: `formal_release`
- price: 未定
- distribution: 未公開
- Phase 1技術受入: PASS
- Phase 1人間受入: Issue #17の開始条件として、名前変更、Undo、マウスホイール、リフレッシュを含む主要操作PASS
- Phase 2正式出荷準備: ローカル資材とDraft PR準備まで完了。公開待ち

## ORIGINAL由来の派生ビュー

- `README.md`: GitHub公開用の概要、使い方、安全仕様、開発・ビルド、第三者ライセンス、DAKE_META、RELEASE_BODY
- `DAKE_META`: `draft` / `market` / `formal_release` を維持。`release_url` は空、launcher/site表示はfalse
- `release_body.md`: READMEのRELEASE_BODYと完全一致する4項目
- `booth_product.txt`: 正本の概要・注意事項を反映。価格、GitHub Release URL、BOOTH URLは空欄
- `booth_ready/booth_product.txt`: ルートのBOOTH登録用ビューと完全一致

## 公開画像

- キャプチャ時の入力元: `C:\Users\Public\Documents\DAKE_synthetic_release_20260902_48`（キャプチャ後に削除済み）
- データ: `tests/synthetic_pdf_trial.py` と同じ生成関数による無機密合成PDF 48件
- 実データ・個人情報: 不使用
- `assets/screenshot.webp`: 実行中のビルド済みexeのアプリウインドウ、1182 x 812、WebP
- `assets/screenshot.jpg`: 同一キャプチャのJPEG派生
- `assets/booth_thumbnail.jpg`: DAKE共通生成方式による1200 x 1200 JPEG
- 表示内容: タイトルと説明の横並び、48/48のサムネイル進捗、カードと名前入力欄、変更待ち1件、フッター

## 第三者ライセンス

実際のビルド環境を照会した。

- pypdfium2: 5.13.0
- PDFium: 153.0.7999.0
- PDFium origin: `pdfium-binaries`
- PDFium flags: なし
- wheel metadata: `License: BSD-3-Clause, Apache-2.0, dependency licenses`

`pypdfium2-5.13.0.dist-info/METADATA` の `License-File` 全19件を、内容を変更せず次へ収録した。

- `01_apps/DAKE_PDF_OverviewRename/third_party_licenses/pypdfium2-5.13.0/`
- `01_apps/DAKE_PDF_OverviewRename/booth_ready/third_party_licenses/pypdfium2-5.13.0/`
- ローカル配布zip内の `third_party_licenses/pypdfium2-5.13.0/`

コピー元とGit管理側19件のファイル名・SHA-256集合が一致することを確認した。根拠としてpypdfium2公式リポジトリとPDFium公式LICENSEも確認した。要約は `THIRD_PARTY_NOTICES.txt`、条件の原文は各同梱文書を参照する。

## 自動テスト

実行:

```text
python -m pytest -q
```

結果:

```text
41 passed in 1.09s
```

含まれる主な回帰:

- 通常変更、入れ替え、循環、case-only、日本語名、衝突・予約名・外部変更、ロールバック
- Undo成功と衝突時の開始前中止
- root配下のwheel routing、大プレビュー除外、実Tkでの48カードwheel
- リフレッシュの破棄拒否と完全初期化
- thumbnail worker / preview worker間のPDFium critical section最大同時実行数1
- README / DAKE_META / release_body / booth_product / wheelライセンス集合の整合

## 合成PDF再試験

実行:

```text
python tests/synthetic_pdf_trial.py 1 48 100 300
```

| 件数 | 生成 | scan | render | rename | undo | render失敗 | ハッシュ維持 | 一時ファイル | worker停止 |
| ---: | ---: | ---: | ---: | ---: | ---: | ---: | --- | ---: | --- |
| 1 | 0.078s | 0.001s | 0.079s | 0.004s | 0.002s | 0 | OK | 0 | OK |
| 48 | 0.106s | 0.012s | 0.112s | 0.124s | 0.128s | 0 | OK | 0 | OK |
| 100 | 0.216s | 0.027s | 0.349s | 0.393s | 0.411s | 0 | OK | 0 | OK |
| 300 | 0.757s | 0.122s | 0.912s | 0.892s | 0.788s | 0 | OK | 0 | OK |

## Windows build / smoke

- OS: Windows 11 `10.0.26200`
- Python: 3.12.4
- PyInstaller: 6.19.0
- `build.bat`: PASS
- onefile / noconsole: PASS
- 生成物: `dist/DakePDF_OverviewRename.exe`
- exe起動: PASS
- 合成PDF 48件読込・サムネイル48/48: PASS
- 名前入力・変更待ち表示: PASS
- マウスホイールによる一覧移動: PASS
- リフレッシュ時の未反映入力破棄確認: PASS
- リフレッシュ完全初期化: pytest実Tk統合試験でPASS
- exeアイコン資源: 2件検出
- ウインドウ: 検出
- ウインドウアイコンハンドル: 非0
- 共通アイコンファイル: 検出
- タスクバーアイコン直接目視: 未確認

表示倍率100%、125%、150%について1180px / 900px幅の全6条件で、ヘッダー横並び、ヘッダー・ツールバー・フッターの収まりを確認しPASSした。

## booth_ready / 配布候補

ローカル `booth_ready/DakePDF_OverviewRename.zip` を作成した。zipはGit除外対象で公開していない。

収録23ファイル:

- `DakePDF_OverviewRename.exe`
- `README.txt`
- `注意事項.txt`
- `THIRD_PARTY_NOTICES.txt`
- `third_party_licenses/` 配下19件

画像・登録文面はzip外の `booth_ready/` にも揃えている。

## 未確認・未実施

- タスクバーアイコンの直接目視
- 価格決定
- GitHub Release URL / BOOTH URL / Stripe導線
- GitHub Release公開
- BOOTH公開
- dakeapp.com / Store / Cloudflare反映
- mainへのmerge

上記は未確認または明示的な非実施であり、正式出荷完了とは扱わない。
