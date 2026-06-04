# ORIGINAL Phase4 Review - DAKE_Time_AdvancedTimer

## 目的

`DAKE_Time_AdvancedTimer` に試験導入した `ORIGINAL.md` を真の正本として、既存の派生ビューである README / DAKE_META / release_body / booth_product / booth_ready内テキストとの整合性を確認する。

今回はレビューのみ行い、派生ビューは更新しない。

## 参照した正本

- `01_apps/DAKE_Time_AdvancedTimer/ORIGINAL.md`

## 比較した派生ビュー

- `01_apps/DAKE_Time_AdvancedTimer/README.md`
- `01_apps/DAKE_Time_AdvancedTimer/release_body.md`
- `01_apps/DAKE_Time_AdvancedTimer/booth_product.txt`
- `01_apps/DAKE_Time_AdvancedTimer/booth_ready/booth_product.txt`
- `01_apps/DAKE_Time_AdvancedTimer/booth_ready/README.txt`
- `01_apps/DAKE_Time_AdvancedTimer/booth_ready/注意事項.txt`

DAKE_META は単独ファイルではなく、README内の `DAKE_META` JSONブロックとして確認した。

## README.md 差分

| 項目 | レビュー結果 | 判断 |
| --- | --- | --- |
| タイトル | READMEは `アドバンスドタイマー`、ORIGINALは `Dakeアドバンスドタイマー` / `アドバンスドタイマー` を保持。短縮タイトルとして整合。 | 矛盾なし |
| 概要 | READMEの概要はORIGINALの目的・公開用説明と一致。 | 矛盾なし |
| 対象ユーザー | READMEには対象ユーザー欄なし。ORIGINALには対象ユーザーを追加済み。 | ORIGINALにだけある重要情報 |
| 使い方 | READMEには詳細な使い方欄なし。ORIGINALには使い方の要点を追加済み。 | ORIGINALにだけある重要情報 |
| 機能一覧 | READMEとORIGINALで一致。 | 矛盾なし |
| 注意事項 | READMEには詳細な注意事項なし。ORIGINALでは「既存READMEに詳細な注意事項の記載なし」と整理し、BOOTH ready由来の注意事項を保持。 | 次回更新候補 |
| ビルド方法 | READMEとORIGINALで一致。 | 矛盾なし |
| Release URL | READMEのDAKE_METAとORIGINALで一致。 | 矛盾なし |

READMEが古いというより、READMEはGitHub公開用ビューとして短く、ORIGINALが対象ユーザー・困りごと・Store用情報まで広く持つ状態になっている。

## DAKE_META 差分

README内の `DAKE_META` とORIGINAL内の `DAKE_META生成用情報` は、以下の項目で一致している。

- `app_key`
- `display_name`
- `launcher_title`
- `launcher_description`
- `site_title`
- `site_description`
- `update_summary`
- `folder_name`
- `exe_name`
- `release_url`
- `screenshot_path`
- `status`
- `show_in_launcher`
- `show_on_site`
- `app_type`
- `completion_goal`

差分なし。

## release_body.md 差分

`release_body.md` とORIGINALの `release_body生成用情報` は一致している。

- 集中時間と休憩時間をすぐ始めるタイマー
- 25分 / 5分 / 15分 / カスタム時間に対応
- 静かなDAKE UI
- Windows向けexe

価格、BOOTH URL、Store情報などは混ざっていない。

GitHub Release用ビューとして適切。

## booth_product.txt 差分

アプリ直下の `booth_product.txt` と `booth_ready/booth_product.txt` は同内容。

| 項目 | レビュー結果 | 判断 |
| --- | --- | --- |
| 商品名 | ORIGINALと一致。 | 矛盾なし |
| 価格案 | ORIGINALと一致。 | 矛盾なし |
| 商品紹介文 | 主文と機能箇条書きはORIGINALと一致。 | 矛盾なし |
| 締め文 | `実務の流れを、少し静かにするための道具です。` はbooth_product側にのみ存在。 | ORIGINAL.mdへ戻すべき情報 |
| タグ | ORIGINALと一致。 | 矛盾なし |
| 商品画像 | ORIGINALと一致。 | 矛盾なし |
| 補助画像 | ORIGINALと一致。 | 矛盾なし |
| 作品ファイル | ORIGINALと一致。 | 矛盾なし |
| GitHub Release URL | ORIGINALと一致。 | 矛盾なし |
| BOOTH URL | ORIGINALと一致。 | 矛盾なし |
| 注意事項 | booth_product側は4項目。ORIGINALと `booth_ready/注意事項.txt` にはWindows警告の注意もある。 | 次回更新候補 |

booth_productはBOOTH登録用ビューとして成立している。
ただし、BOOTH本文を完全にORIGINALから再生成できる状態にするなら、締め文をORIGINALの `booth_product生成用情報` に戻すのがよい。

## booth_ready README / 注意事項 差分

### booth_ready/README.txt

`booth_ready/README.txt` は配布zip同梱用の短い説明として、ORIGINALの公開用説明・release_body生成用情報と整合している。

README.txtにだけある情報:

- `PEAKHEADZ`
- `https://peakheadz.com`
- `Vibe-Coded by Yukihiko Kikuta`

これは配布物の同梱表示として置かれている情報で、現時点では派生ビュー側だけでよい。
将来、全配布物の署名・クレジットをORIGINALから生成する方針にする場合は、ORIGINALの免責・同梱方針へ戻してもよい。

### booth_ready/注意事項.txt

ORIGINALの免責・注意事項と一致している。

booth_product.txtよりも1項目多く、以下を含む。

- 環境によっては起動時にWindowsの警告が表示される場合があります

この注意はORIGINALに戻っているため、正本情報としては保持済み。

## ORIGINAL.mdへ戻すべき情報

- BOOTH商品紹介文の締め文:
  - `実務の流れを、少し静かにするための道具です。`

この文を今後もBOOTH本文として使うなら、ORIGINAL.mdの `booth_product生成用情報` に明示しておくと再生成時に消えにくい。

## 派生ビュー側だけでよい情報

- `booth_ready/README.txt` の同梱表示:
  - `PEAKHEADZ`
  - `https://peakheadz.com`
  - `Vibe-Coded by Yukihiko Kikuta`
- `booth_ready/booth_product.txt` はBOOTH登録時の実使用ビュー。
- アプリ直下 `booth_product.txt` は開発・確認用のBOOTH登録ビュー。
- `release_body.md` はGitHub Releaseへ貼るための短いビュー。

## 次Phaseで更新するべきファイル

優先度順:

1. `01_apps/DAKE_Time_AdvancedTimer/ORIGINAL.md`
   - BOOTH締め文を `booth_product生成用情報` に戻すか判断する。
2. `01_apps/DAKE_Time_AdvancedTimer/booth_product.txt`
   - ORIGINAL.mdへ締め文を戻した後、必要なら再生成または整合確認する。
3. `01_apps/DAKE_Time_AdvancedTimer/booth_ready/booth_product.txt`
   - アプリ直下booth_productと同内容を維持する。

README、DAKE_META、release_bodyは現時点で必須更新なし。

## 更新しない方がよいファイル

- `main.py`
- `build.bat`
- `assets/`
- `dist/`
- `build/`
- `*.spec`
- `booth_ready/DakeAdvanced_Timer.zip`

今回のPhase 4-1では、派生ビューも更新しない。

## 判断メモ

- `ORIGINAL.md` は、READMEより広い情報を保持している。
- READMEはGitHub公開用ビューとして短くてよく、対象ユーザーやStore未確定項目まで載せる必要はない。
- DAKE_META、release_body、booth_productの主要値は概ねORIGINALと整合している。
- 差分の中心は、BOOTH本文の締め文と注意事項の粒度。
- 次Phaseでは、派生ビューを直す前に「締め文をORIGINALに戻すか」を決めるときれい。
