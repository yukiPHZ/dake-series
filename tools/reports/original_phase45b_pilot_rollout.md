# ORIGINAL Phase4-5B Pilot Rollout

## 目的

Advanced Timer で確認した `ORIGINAL.md` 正本運用を、A優先ORIGINAL対象のうちジャンルの異なる5アプリへ試験横展開した。

今回は全件展開ではなく、PDF系・画像系・ファイル整理系・バックアップ系・メモ系で、同じテンプレートが破綻しないかを確認することを目的とした。

## 選定したアプリ

| folder | title | reason |
|---|---|---|
| DAKE_PDF_Compress | DakePDF圧縮 | PDF系。BOOTH登録済み、GitHub Releaseあり、SHIMARISU CLI連携情報もあり、複雑寄りの実務アプリとして検証できるため。 |
| DAKE_Image_Resize | Dake画像リサイズ | 画像系。BOOTH登録済み、GitHub Releaseあり、対応形式・非対応形式を持つ画像処理アプリとして検証できるため。 |
| DAKE_Folder_List | Dakeフォルダ一覧 | ファイル整理系。フォルダ構造の表示・コピー・保存と、非破壊方針を持つアプリとして検証できるため。 |
| DAKE_Backup | Dakeバックアップ | バックアップ系。保存方針・ログ・設定・安全側の思想が強く、販売説明以外の正本情報を整理する検証に向くため。 |
| DAKE_Sticky_Memo | Dake付箋メモ | メモ系。保存しない・分類しない・書いたら終わるという、やらないことが明確な軽量アプリとして検証できるため。 |

## 作成した ORIGINAL.md

| folder | source files used | missing info | memo |
|---|---|---|---|
| DAKE_PDF_Compress | README.md, README内DAKE_META, release_body.md, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/ | Store URL、Storeダウンロード導線、Storeサポート方針は未確定。 | SHIMARISU CLI連携・Ghostscript利用・元PDF非破壊方針を正本側へ整理できた。 |
| DAKE_Image_Resize | README.md, README内DAKE_META, release_body.md, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/ | Store URL、Storeダウンロード導線、Storeサポート方針は未確定。booth_product側にはGitHub Release URL欄がなく、README内DAKE_METAを参照した。 | 対応形式とHEIC/HEIF非対応を、StoreやREADMEへ派生しやすい形で整理できた。 |
| DAKE_Folder_List | README.md, README内DAKE_META, release_body.md, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/ | Store URL、Storeダウンロード導線、Storeサポート方針は未確定。booth_product側にはGitHub Release URL欄がなく、README内DAKE_METAを参照した。 | ファイル編集・削除・移動・リネーム・内容読取をしない非破壊方針を正本側へ残せた。 |
| DAKE_Backup | README.md, README内DAKE_META, release_body.md, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/ | Store URL、Storeダウンロード導線、Storeサポート方針は未確定。 | 「消さない」「一方向バックアップ」「削除同期しない」など、事故防止に関わる正本情報を整理できた。 |
| DAKE_Sticky_Memo | README.md, README内DAKE_META, release_body.md, booth_ready/booth_product.txt, booth_ready/README.txt, booth_ready/注意事項.txt, assets/ | Store URL、Storeダウンロード導線、Storeサポート方針は未確定。booth_product側にはGitHub Release URL欄がなく、README内DAKE_METAを参照した。 | 保存・検索・タグ・分類・エクスポートをしない軽量思想を正本側へ残せた。 |

## 共通してうまくいった点

- `README.md`、README内`DAKE_META`、`release_body.md`、`booth_ready/booth_product.txt` の組み合わせで、アプリ説明・Release情報・BOOTH情報をおおむね集約できた。
- BOOTH URL、GitHub Release URL、zip名、画像パスは既存ファイルから取得でき、推測で埋めずに整理できた。
- Store関連は未確定のまま残せるため、Store構築前でも `ORIGINAL.md` を導入できる。
- PDF、画像、ファイル整理、バックアップ、メモの各ジャンルで、Advanced Timer方式は大きく破綻しなかった。
- 既存の派生ビューやアプリ本体を変更せず、真の正本だけを追加できた。

## アプリごとに詰まった点

- DAKE_PDF_Compress は、CLI連携・Ghostscript・圧縮方式・元ファイル非破壊など、通常テンプレートより運用情報が多い。
- DAKE_Image_Resize は、対応形式・非対応形式を明示できる欄がテンプレート上にあると扱いやすい。
- DAKE_Folder_List は、「やらないこと」や非破壊方針が価値の一部なので、機能一覧とは別に残す欄がほしい。
- DAKE_Backup は、設定・ログ・保存先・削除しない方針など、事故防止系の注意を独立して扱えるとよい。
- DAKE_Sticky_Memo は、保存しない・分類しないなどの非ゴールが重要で、軽量アプリほど「やらないこと」欄が効く。
- 一部の `booth_ready/booth_product.txt` には GitHub Release URL 欄がなく、README内DAKE_METAの `release_url` を参照した。

## テンプレート改善候補

- 任意セクションとして `CLI連携` を追加できると、SHIMARISU連携アプリを扱いやすい。
- 任意セクションとして `対応形式・非対応形式` を追加できると、画像・PDF・ファイル系で説明の重複を減らせる。
- 任意セクションとして `やらないこと / 非ゴール` を追加できると、DAKEらしい単機能思想を保持しやすい。
- 任意セクションとして `設定・ログ・保存方針` を追加できると、バックアップ系・管理系アプリで事故防止情報を整理しやすい。
- 任意セクションとして `非破壊・上書き禁止方針` を追加できると、ファイル操作系アプリの安全思想を明確にできる。

## 次Phaseへの提案

Phase 4-5C では、A優先対象の残り全件へ一気に進む前に、テンプレートへ任意セクションを追加するか、または10〜15件単位のバッチ展開にするのが安全。

特にPDF/画像/ファイル操作/メモ/管理ツールでは、基本テンプレートに加えて「やらないこと」「対応形式」「CLI連携」「保存方針」を必要に応じて入れる運用がよい。

## 全件展開してよいかの判断

Advanced Timer方式は、今回選定した5ジャンルでも大きく破綻しなかったため、A優先対象への横展開は可能と判断する。

ただし、49件を無確認で一括生成するより、次はテンプレート改善を反映したうえで、10〜15件単位のバッチ展開を推奨する。

全件展開時も、READMEやbooth_productなどの派生ビューは更新せず、まず `ORIGINAL.md` の作成だけに限定するのが安全。
