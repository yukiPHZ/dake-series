# DakePDF俯瞰名前変更 v1.0.2 調査・出荷記録

更新日: 2026-09-29 JST

2026-09-30追記: この記録の後、実機400件でHuman Review FAIL。現在は単なる接続待ちではなく表示不具合の修正候補・利用者再確認待ち。旧84 testsの履歴は以下に維持し、原因・修正・99 testsと400件結果は [追加記録](dake_pdf_overview_rename_v102_viewport.md) を参照。正式出荷HOLD、PR #47 Draftを維持する。

## 現在の判定

実装・自動試験・Windows候補ビルドまでは完了した。正式公開は **HOLD** とする。

理由は、CodexのWindows native Computer Use接続にアプリ一覧が提供されず、最終配布ZIPを展開したexeについて、初回フォルダ選択後の順次表示、27件目以降と末尾へのスクロール、タスクバーアイコンを実画面で直接確認できていないためである。ソース上の実Tk試験やWin32ハンドル検査を、この出荷ゲートの代用にはしない。

## 調査対象

- 公開版: `DAKE_PDF_OverviewRename_v1.0.1`
  - tag commit: `d9e79825ea0a5c3c99eb4fb67eeab26ff37d6557`
  - GitHub配布ZIP: 34,140,032 bytes
  - ZIP SHA-256: `2b84ba6cdb14f33e5b3ecb0aa0558c24a3fa9b430a9cd3d30c1eabb3de1b7541`
  - 展開exe: 34,755,329 bytes
  - exe SHA-256: `649f3470a5ea460b594332704cd6e21ad1d89a8114770c1716f37d597a3b5c91`
- 作業基点: `origin/main` `e0fa3efa232265db7053a08ae847c0b98f4d2a65`
- 作業ブランチ: `codex/dake-pdf-overview-rename-v1.0.2`
- 会社で使用中のexe: 入手できていないため、版と症状の再現有無は未確認

## 26件前後の疑い

現行v1.0.1と作業前mainに固定26件上限は見つからなかった。`CARD_BATCH_SIZE=24` はカード生成単位であり、全件数の上限ではない。v1.0.0にあった同一列数時のlayout early returnはv1.0.1の `_laid_out_count` 管理で既に修正されている。

作業前mainを0/1/23/24/25/26/27/32/47/48/49/100/300件で試験し、対象PDF集合、scan結果、カード集合、レンダー結果、rename/Undo後の内容ハッシュを照合した。1～300件で欠落・重複・後続停止は再現しなかった。0件はrename対象なしとして扱う試験へ修正した。

一方、scan中の `FileSnapshot.capture()` が `OSError` になったPDFを黙って除外する経路があった。これは固定上限ではないが、「フォルダにはあるのに一覧に出ない」症状を作り得るため、部分一覧を完成扱いせずscan全体を明示的なエラーにするよう変更した。

PDFの内容エラーはカードを消さない。空、破損、0ページ、暗号化PDFを24～27番目に置いた32件試験で、全32カードを保持し、24～27を失敗表示、28～32を正常処理することを確認した。

## v1.0.2変更

- タイトルへ `v1.0.2` を控えめに表示。
- 各カードへ `24ページ ｜ 3.8 MB` 形式のページ数・実ファイル容量を表示。
- 初期状態は `ページ数確認中… ｜ 286 KB`、取得不能時は `ページ数不明 ｜ 286 KB` とし、0ページを偽らない。
- 容量は1024進で、Bは整数、KBは整数丸め、MB/GBは小数1桁。
- ページ数取得後にレンダーだけ失敗した場合も、既知のページ数と容量を残す。
- ステータスを `PDF 48件 ｜ サムネイル処理 32/48（失敗2）｜ 変更待ち3件` へ整理。処理数は成功＋失敗で、未処理を完了扱いしない。
- 緑色・太字・非モーダルのrename/Undo成功表示は維持し、サムネイル更新では消さない。
- `rename_core.py`、two-phase rename、Undo、refresh/reload、wheel、PDFium mutex、Latest Jobの意味は変更していない。

## 試験

- `pytest 01_apps/DAKE_PDF_OverviewRename/tests -q`: **84 passed**
- 境界合成PDF: **0/1/23/24/25/26/27/32/47/48/49/100/300 全条件PASS**
- 300件性能比較:
  - 公開v1.0.1: scan 0.082秒、render 1.048秒、render増分working set 12.7 MB
  - v1.0.2候補: scan 0.086秒、render 0.959秒、render増分working set 12.2 MB
- 6レイアウト条件（100/125/150% × 幅1180/900相当）: toolbar/header/status/footerの欠け・重なりなし
- Windows候補exe:
  - FileVersion/ProductVersion 1.0.2
  - exe icon resource 2件
  - 起動ウインドウを検出し、window icon handle非zero

## 候補配布物（未公開）

- exe: 34,757,879 bytes
- exe SHA-256: `225dfdf498f9db46effa34603679f43d85551fe9a3c9c9691a1d159e804dd09d`
- ZIP: 34,157,382 bytes
- ZIP SHA-256: `bc72b427f03243c65c26def77226c5af2984a93d19dc6f552d565c5aed876aa6`
- ZIP内exeとビルド元exeのSHA-256一致
- README、注意事項、THIRD_PARTY_NOTICES、pypdfium2/PDFiumライセンス19件を同梱

これらはPR前の候補ビルドであり、main merge後の正式ビルドではない。公開物のSHAとして確定しない。

## 公開HOLD中の項目

- 最終ZIP展開exeでの初回48件順次表示
- 27件目以降と最終カードへのスクロール、ページ数・容量の直接目視
- 最終カードのrename/Undo、reload後の全件維持
- 最終exeで100/300件の末尾到達
- タスクバーアイコンの直接目視
- 新UI公開画像
- main merge、tag、GitHub Release、BOOTH、dakeapp.com、Store、Cloudflare、Issueの正式出荷完了化

このHOLDを解消してmain merge後に正式ビルドし直し、正式ZIPの実画面確認が完了するまで「正式出荷完了」としない。
