# ORIGINAL Phase4-5A App Inventory

## 目的

Advanced Timerでの `ORIGINAL.md` 試験導入成功を受け、全DAKEアプリへ横展開する前の対象棚卸しを行う。

今回は `ORIGINAL.md` を一括作成せず、Store接続前にどのアプリを優先して正本化するべきかを整理する。

## 調査対象

- 主対象: `01_apps/` 直下のアプリフォルダ
- 参考対象: `04_packs/` 直下のPackフォルダ

参照した主なルール・レポート:

- `00_core/DAKE_ORIGINAL_RULE.md`
- `00_core/DAKE_ORIGINAL_TEMPLATE_APP.md`
- `00_core/CHATGPT_CODEX_WORKFLOW.md`
- `00_core/DAKE_COMMON_SPEC.md`
- `tools/reports/original_phase4_advanced_timer_completion.md`

## 集計

- アプリ総数: 65
- ORIGINAL.md あり: 1
- ORIGINAL.md なし: 64
- BOOTH登録済み: 49
- market系: 4
- available: 50
- 優先ORIGINAL対象: 50
- 通常ORIGINAL対象: 1
- 保留・確認対象: 14
- 対象外候補: 2

## A：優先ORIGINAL対象

| folder | title | status | app_type | booth | release | original | priority | memo |
|---|---|---|---|---|---|---|---|---|
| DAKE_App_Doko | アプリどこ | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_Backup | Dakeバックアップ | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_BOOTH_Assist | BOOTHアシスト | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_Column_Memo | ずっとメモ | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_Document_Cover | Dake書類送付状 | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_FAX_Cover | DakeFAX送付状 | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_Folder_List | Dakeフォルダ一覧 | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_Game_Alien_Road | DakeAlien Road | available | market | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, market, formal_release, available, Releaseあり |
| DAKE_Game_Diver_Catch | Dake潜って捕る | available | market | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, market, formal_release, available, Releaseあり |
| DAKE_Git_Memo | DakeGitメモ | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_Image_HEICtoJPG | HEIC→JPG変換 | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_Image_iPhoneToPC | Dake画像iPhoneToPC | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_Image_PasteA4 | 貼る | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_Image_Receiver | DakeImage_Receiver | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_Image_Resize | Dake画像リサイズ | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_Image_ToPDF | DakeImageToPDF | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_Launcher | Dakeランチャー | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_Mail_Address_Format | Dakeメールアドレス整形 | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_Mail_AllStaff | Dake全社員メール起動 | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_Mail_Draft | Dakeメール下書き | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_Mail_Kikuta | Dake菊田メール | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_Mail_List | Dakeメールリスト | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_Maji_Memo | マジでメモ | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_Mansion_Schedule | マンション工程表 | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_PDF_CheckStamp | Dake確認印 | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_PDF_Compress | DakePDF圧縮 | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_PDF_Crop | DakePDFトリミング | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_PDF_LookHere | DakePDFここ見て | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_PDF_Marker | DakePDFマーカー | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_PDF_Merge | DakePDF結合 | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_PDF_Merge_Mini | DakePDF結合mini | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_PDF_Rename | DakePDFファイル名整理 | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_PDF_Reorder | DakePDFページ並べ替え | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_PDF_SplitOne | DakePDF分割One | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_PDF_SplitSelect | DakePDF分割Select | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_PDF_ToImages | DakePDFto画像 | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_PDF_Viewer | DakePDF見る | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_Price_Apportionment | Dake価格按分 | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_Price_FixedTax | Dake固都税計算 | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_Reform_Progress | リフォーム進捗管理 | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_Screen_WebP | DakeScreen_WebP | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_Screenshot_Print | Dakeスクショ印刷 | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_Sticky_Memo | 付箋メモ | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_Time_AdvancedTimer | Dakeアドバンスドタイマー | available | market | url | yes | yes | done | ORIGINAL導入済み; BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, market, formal_release, available, Releaseあり |
| DAKE_TwoPerson_Memo | Dake二人メモ | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_Video_Shorts_Cut | Dakeショート切り出し | available | market | ready | yes | no | high | zipあり; Store掲載候補; booth_readyあり, market, formal_release, available, Releaseあり |
| DAKE_Work_Calendar | Dake工程カレンダー | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_Year_Age | Dake築年数 | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_Year_Notice | Dake今年の注意点 | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |
| DAKE_Yesterday_Task_Memo | Dake昨日タスクメモ | available | unknown | url | yes | no | high | BOOTH URLあり; zipあり; Store掲載候補; BOOTH登録済み, available, Releaseあり |

## B：通常ORIGINAL対象

| folder | title | status | app_type | booth | release | original | priority | memo |
|---|---|---|---|---|---|---|---|---|
| DAKE_Game_ShimarisuRealEstate | Dakeしまりす不動産 | prototype | unknown | - | - | no | medium | screenshotなし; thumbnailなし; 今後確認候補 |

## C：保留・確認対象

| folder | title | reason | memo |
|---|---|---|---|
| DAKE_App_Dashboard | DAKE Dashboard | status=internal, app_type=qpcs, completion_goal=system_ready | screenshotなし; thumbnailなし |
| DAKE_Approve_Brainz | 承認Brainz | status=frozen | zipあり; screenshotなし; thumbnailなし |
| DAKE_BGM_Loop | Dake BGM Loop | status=frozen | zipあり; screenshotなし; thumbnailなし |
| DAKE_Brainz_OIKAWA | OIKAWA | status=frozen, app_type=qpcs, completion_goal=reference_ready | zipあり; screenshotなし; thumbnailなし |
| DAKE_Brainz_Search | 補助脳BRAINZ | status=frozen, app_type=qpcs, completion_goal=reference_ready | zipあり; screenshotなし; thumbnailなし |
| DAKE_HolidayJinja_Post | holiday-jinja 投稿DAKE | status=internal | zipあり |
| DAKE_Music_Otooku | 音を置く | status=frozen | zipあり; screenshotなし; thumbnailなし |
| DAKE_Note_Inbox | DAKE_Note_Inbox | DAKE_METAなし/解析不可 | screenshotなし; thumbnailなし; missing |
| DAKE_QPSC_Dashboard | QPCS Dashboard | status=internal, app_type=qpcs, completion_goal=system_ready | screenshotなし; thumbnailなし |
| DAKE_Wake_Brainz | DAKE_Wake_Brainz | status=draft, app_type=qpcs, completion_goal=system_ready | zipあり; screenshotなし; thumbnailなし |
| DAKE_Web_Dashboard | DAKE Web Dashboard | status=frozen, app_type=qpcs, completion_goal=system_ready | screenshotなし; thumbnailなし |
| DAKE_Web_Index | DAKE Web Index | status=internal | screenshotなし; thumbnailなし |
| DAKE_Yukiz_KadouChu | Dakeユキズ稼働中 | status=draft | zipあり; screenshotなし; thumbnailなし |
| DAKE_YukizBlog_Post | YUKIZ BLOG 投稿DAKE | status=internal | zipあり |

## D：対象外候補

| folder | reason | memo |
|---|---|---|
| DAKE_Pack_Document | アプリ本体ではなくPackフォルダ | Pack用ORIGINAL設計は別Phaseで扱う |
| DAKE_Pack_Memo | アプリ本体ではなくPackフォルダ | Pack用ORIGINAL設計は別Phaseで扱う |

## 次Phaseへの提案

1. A分類のうち `original=no` のアプリから、BOOTH登録済みまたはStore掲載候補を優先して `ORIGINAL.md` を作成する。
2. Advanced Timerと同じ手順で、まず3〜5本だけ小さく横展開してテンプレートの不足を確認する。
3. Store用 generated データ形式を先に定義し、Store専用の商品正本を作らない方針を維持する。
4. C分類はStore候補に入れず、status / app_type / completion_goal の見直しや凍結理由の整理を別Phaseで行う。
5. `04_packs/` はアプリ本体とは別のPack商品として、Pack用ORIGINALテンプレートを検討する。

## 注意点

- 今回は調査レポートのみで、`ORIGINAL.md` は新規作成していない。
- README、DAKE_META、release_body、booth_product、assets、dist、booth_ready内ファイルは変更していない。
- A分類はStore掲載候補になり得るが、即Store掲載を意味しない。公開前に各アプリの `ORIGINAL.md` 作成と派生ビュー整合確認が必要。
- C分類のうちRelease URLや過去素材を持つものも、statusやapp_typeが非market系なら通常Store候補から外す。
- BOOTH URLは `booth_ready/booth_product.txt` を優先し、なければアプリ直下 `booth_product.txt` を参照した。
- `market系` は、現時点で `DAKE_META.app_type: market` が明示されている件数。未設定でも `status: available` のアプリはA分類に含めた。

## 生成メモ

このレポートは `01_apps/` のREADME内 `DAKE_META` とBOOTH関連ファイルの有無を機械的に棚卸しし、最終判断をPhase 4-5A用に整理したもの。
