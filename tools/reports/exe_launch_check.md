# DAKE exe Launch Check

## Summary

- total apps: 65
- checked apps: 50
- OK: 14
- OK_GUI: 36
- TIMEOUT: 0
- ERROR: 0
- NO_EXE: 0
- SKIPPED_STATUS: 15
- launch-check unsupported among checked exe apps: 36

## Problems

- none

## Checked

| app | display_name | status | exe | launch-check | result | exit code | elapsed | note |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| DAKE_App_Doko | アプリどこ | available | 01_apps\DAKE_App_Doko\dist\DakeApp_Doko.exe | True | OK | 0 | 1.45 |  |
| DAKE_Backup | Dakeバックアップ | available | 01_apps\DAKE_Backup\dist\DakeBackup.exe | True | OK | 0 | 1.38 | LAUNCH CHECK OK |
| DAKE_BOOTH_Assist | BOOTHアシスト | available | 01_apps\DAKE_BOOTH_Assist\dist\DakeBOOTH_Assist.exe | True | OK | 0 | 2.42 | 初期DOMに安全なfile inputがないため手動案内: 商品画像 初期DOMに安全なfile inputがないため手動案内: zipファイル DakeBOOTHアシスト launch-check OK apps_root=C:\Users\yukiz\devlop\DAKE_series\01_apps apps=68 product_apps=61 pack_products=3 shimarisu_pack=ready dake_backup_product=booth_ready/booth_product.txt dake_backup_fields=title=True,price=True,description=True,tags=True dake_backup_tag_count=6 cdp_url=http://127.0.0.1:9222 edit_url_hint=https://manage.booth.pm/items/数字/edit chrome_path=C:\Program Files\Google\Chrome\Application\chrome.exe chrome_profile=C:\Users\yukiz\AppData\Local\DakeBOOTH_Assist\chrome_profile fixtures=missing_product, missing_ready, multiple_zip, missing_playwright_python, ready_product_lookup, ready_product_priority, factory_v2_product, readme_fallback, open_edit_url_only, bad_page_detection, tag_split, file_input_detection, manual_file_upload, booth_url_save, pack_product |
| DAKE_Column_Memo | ずっとメモ | available | 01_apps\DAKE_Column_Memo\dist\DakeColumn_Memo.exe | False | OK_GUI | terminated | 4.03 | GUI process started and was stopped after smoke timeout |
| DAKE_Document_Cover | Dake書類送付状 | available | 01_apps\DAKE_Document_Cover\dist\DakeDocument_Cover.exe | False | OK_GUI | terminated | 4.03 | GUI process started and was stopped after smoke timeout |
| DAKE_FAX_Cover | DakeFAX送付状 | available | 01_apps\DAKE_FAX_Cover\dist\DakeFAX_Cover.exe | False | OK_GUI | terminated | 4.02 | GUI process started and was stopped after smoke timeout |
| DAKE_Folder_List | Dakeフォルダ一覧 | available | 01_apps\DAKE_Folder_List\dist\DakeFolder_List.exe | False | OK_GUI | terminated | 4.02 | GUI process started and was stopped after smoke timeout |
| DAKE_Game_Alien_Road | DakeAlien Road | available | 01_apps\DAKE_Game_Alien_Road\dist\DakeGame_AlienRoad.exe | True | OK | 0 | 1.38 | launch-check ok |
| DAKE_Game_Diver_Catch | Dake潜って捕る | available | 01_apps\DAKE_Game_Diver_Catch\dist\DakeGame_Diver_Catch.exe | True | OK | 0 | 1.22 | LAUNCH CHECK OK |
| DAKE_Git_Memo | DakeGitメモ | available | 01_apps\DAKE_Git_Memo\dist\DakeGit_Memo.exe | True | OK | 0 | 2.12 |  |
| DAKE_Image_HEICtoJPG | HEIC→JPG変換 | available | 01_apps\DAKE_Image_HEICtoJPG\dist\DakeHEIC_JPG.exe | False | OK_GUI | terminated | 4.02 | GUI process started and was stopped after smoke timeout |
| DAKE_Image_iPhoneToPC | Dake画像iPhoneToPC | available | 01_apps\DAKE_Image_iPhoneToPC\dist\DakeImage_iPhoneToPC.exe | False | OK_GUI | terminated | 4.02 | GUI process started and was stopped after smoke timeout |
| DAKE_Image_PasteA4 | 貼る | available | 01_apps\DAKE_Image_PasteA4\dist\DakeImage_PasteA4.exe | False | OK_GUI | terminated | 4.03 | GUI process started and was stopped after smoke timeout |
| DAKE_Image_Receiver | DakeImage_Receiver | available | 01_apps\DAKE_Image_Receiver\dist\DakeImage_Receiver.exe | False | OK_GUI | terminated | 4.00 | GUI process started and was stopped after smoke timeout |
| DAKE_Image_Resize | Dake画像リサイズ | available | 01_apps\DAKE_Image_Resize\dist\Dake_Image_Resize.exe | False | OK_GUI | terminated | 4.05 | GUI process started and was stopped after smoke timeout |
| DAKE_Image_ToPDF | DakeImageToPDF | available | 01_apps\DAKE_Image_ToPDF\dist\DakeImageToPDF.exe | False | OK_GUI | terminated | 4.02 | GUI process started and was stopped after smoke timeout |
| DAKE_Launcher | Dakeランチャー | available | 01_apps\DAKE_Launcher\dist\Dake_Launcher.exe | False | OK_GUI | terminated | 4.02 | GUI process started and was stopped after smoke timeout |
| DAKE_Mail_Address_Format | Dakeメールアドレス整形 | available | 01_apps\DAKE_Mail_Address_Format\dist\DakeMail_Address_Format.exe | True | OK | 0 | 1.34 | DAKE_Mail_Address_Format launch-check OK |
| DAKE_Mail_AllStaff | Dake全社員メール起動 | available | 01_apps\DAKE_Mail_AllStaff\dist\Dake_AllStaff_Mail.exe | True | OK | 0 | 1.50 | DAKE_Mail_AllStaff launch-check OK |
| DAKE_Mail_Draft | Dakeメール下書き | available | 01_apps\DAKE_Mail_Draft\dist\DakeMail_Draft.exe | True | OK | 0 | 1.44 |  |
| DAKE_Mail_Kikuta | Dake菊田メール | available | 01_apps\DAKE_Mail_Kikuta\dist\DakeKikuta_Mail.exe | False | OK_GUI | terminated | 4.00 | GUI process started and was stopped after smoke timeout |
| DAKE_Mail_List | Dakeメールリスト | available | 01_apps\DAKE_Mail_List\dist\DakeMail_List.exe | False | OK_GUI | terminated | 4.02 | GUI process started and was stopped after smoke timeout |
| DAKE_Maji_Memo | マジでメモ | available | 01_apps\DAKE_Maji_Memo\dist\DakeMaji_Memo.exe | False | OK_GUI | terminated | 4.02 | GUI process started and was stopped after smoke timeout |
| DAKE_Mansion_Schedule | マンション工程表 | available | 01_apps\DAKE_Mansion_Schedule\dist\DakeMansion_Schedule.exe | True | OK | 0 | 2.23 | Dakeマンション工程表 launch-check OK pdf=C:\Users\yukiz\devlop\DAKE_series\01_apps\DAKE_Mansion_Schedule\dist\tmp9zh9wqlb\mansion_schedule.pdf a3_landscape_one_page=OK calendar_days=45 start=2026-06-01 finish=2026-07-15 workdays=33 allocated_workdays=28 reserve_workdays=5 weekend_work_bars=none selectable_rows=OK free_rows=OK checkbox_display=OK timestamped_output_filename=OK config_save_restore=OK open_output_folder=OK pdf_submission_footer_removed=OK |
| DAKE_PDF_CheckStamp | Dake確認印 | available | 01_apps\DAKE_PDF_CheckStamp\dist\Dake_Check_Stamp.exe | False | OK_GUI | terminated | 4.03 | GUI process started and was stopped after smoke timeout |
| DAKE_PDF_Compress | DakePDF圧縮 | available | 01_apps\DAKE_PDF_Compress\dist\DakePDF_Compress.exe | True | OK | 0 | 1.88 |  |
| DAKE_PDF_Crop | DakePDFトリミング | available | 01_apps\DAKE_PDF_Crop\dist\DakePDF_Crop.exe | False | OK_GUI | terminated | 4.00 | GUI process started and was stopped after smoke timeout |
| DAKE_PDF_LookHere | DakePDFここ見て | available | 01_apps\DAKE_PDF_LookHere\dist\DakePDF_LookHere.exe | False | OK_GUI | terminated | 4.02 | GUI process started and was stopped after smoke timeout |
| DAKE_PDF_Marker | DakePDFマーカー | available | 01_apps\DAKE_PDF_Marker\dist\DakePDF_Marker.exe | False | OK_GUI | terminated | 4.02 | GUI process started and was stopped after smoke timeout |
| DAKE_PDF_Merge | DakePDF結合 | available | 01_apps\DAKE_PDF_Merge\dist\DakePDF_Merge.exe | False | OK_GUI | terminated | 4.02 | GUI process started and was stopped after smoke timeout |
| DAKE_PDF_Merge_Mini | DakePDF結合mini | available | 01_apps\DAKE_PDF_Merge_Mini\dist\DakePDF_Merge_Mini.exe | False | OK_GUI | terminated | 4.02 | GUI process started and was stopped after smoke timeout |
| DAKE_PDF_Rename | DakePDFファイル名整理 | available | 01_apps\DAKE_PDF_Rename\dist\DakePDF_Rename.exe | False | OK_GUI | terminated | 4.02 | GUI process started and was stopped after smoke timeout |
| DAKE_PDF_Reorder | DakePDFページ並べ替え | available | 01_apps\DAKE_PDF_Reorder\dist\DakePDF_Reorder.exe | False | OK_GUI | terminated | 4.02 | GUI process started and was stopped after smoke timeout |
| DAKE_PDF_SplitOne | DakePDF分割One | available | 01_apps\DAKE_PDF_SplitOne\dist\DakePDF_Split_One.exe | False | OK_GUI | terminated | 4.02 | GUI process started and was stopped after smoke timeout |
| DAKE_PDF_SplitSelect | DakePDF分割Select | available | 01_apps\DAKE_PDF_SplitSelect\dist\DakePDF_Split_Select.exe | False | OK_GUI | terminated | 4.02 | GUI process started and was stopped after smoke timeout |
| DAKE_PDF_ToImages | DakePDFto画像 | available | 01_apps\DAKE_PDF_ToImages\dist\DakePDF_to_Images.exe | False | OK_GUI | terminated | 4.02 | GUI process started and was stopped after smoke timeout |
| DAKE_PDF_Viewer | DakePDF見る | available | 01_apps\DAKE_PDF_Viewer\dist\DakePDF_Viewer.exe | False | OK_GUI | terminated | 4.02 | GUI process started and was stopped after smoke timeout |
| DAKE_Price_Apportionment | Dake価格按分 | available | 01_apps\DAKE_Price_Apportionment\dist\DakePrice_Apportionment.exe | False | OK_GUI | terminated | 4.02 | GUI process started and was stopped after smoke timeout |
| DAKE_Price_FixedTax | Dake固都税計算 | available | 01_apps\DAKE_Price_FixedTax\dist\DakeFixedTax.exe | False | OK_GUI | terminated | 4.02 | GUI process started and was stopped after smoke timeout |
| DAKE_Reform_Progress | リフォーム進捗管理 | available | 01_apps\DAKE_Reform_Progress\dist\DakeReform_Progress.exe | True | OK | 0 | 2.23 | Dakeリフォーム進捗管理 launch-check OK pdf=C:\Users\yukiz\devlop\DAKE_series\01_apps\DAKE_Reform_Progress\dist\tmpf86an2qc\リフォーム進捗カレンダー_テストリフォーム_20260529-20260713.pdf a4_portrait_one_page=OK calendar_days=46 start=2026/05/29 finish=2026/07/13 weeks=8 completion_cell=2026-07-13 完工 weekend_background_and_no_work_items=OK calendar_drag_move_and_resize_logic=OK row_reorder_logic=OK reschedule_done_items_fixed=OK config_save_restore=OK project_save_restore=OK pdf_multi_day_bands=OK reference_day_45=OK max_56_day_limit=OK open_output_folder=OK |
| DAKE_Screen_WebP | DakeScreen_WebP | available | 01_apps\DAKE_Screen_WebP\dist\DakeScreen_WebP.exe | False | OK_GUI | terminated | 4.02 | GUI process started and was stopped after smoke timeout |
| DAKE_Screenshot_Print | Dakeスクショ印刷 | available | 01_apps\DAKE_Screenshot_Print\dist\DakeScreenshot_Print.exe | False | OK_GUI | terminated | 4.02 | GUI process started and was stopped after smoke timeout |
| DAKE_Sticky_Memo | 付箋メモ | available | 01_apps\DAKE_Sticky_Memo\dist\DakeSticky_Memo.exe | False | OK_GUI | terminated | 4.02 | GUI process started and was stopped after smoke timeout |
| DAKE_Time_AdvancedTimer | Dakeアドバンスドタイマー | available | 01_apps\DAKE_Time_AdvancedTimer\dist\DakeAdvanced_Timer.exe | True | OK | 0 | 1.48 | launch-check ok |
| DAKE_TwoPerson_Memo | Dake二人メモ | available | 01_apps\DAKE_TwoPerson_Memo\dist\DakeTwoPerson_Memo.exe | False | OK_GUI | terminated | 4.00 | GUI process started and was stopped after smoke timeout |
| DAKE_Video_Shorts_Cut | ショート切り出し | available | 01_apps\DAKE_Video_Shorts_Cut\dist\DakeVideo_Shorts_Cut.exe | True | OK | 0 | 3.28 | Dakeショート切り出し launch-check OK ffmpeg=available ffprobe=available segments=OK transfer_html=OK transcript_format=OK review=OK qr_dependency=available |
| DAKE_Work_Calendar | Dake工程カレンダー | available | 01_apps\DAKE_Work_Calendar\dist\DakeWork_Calendar.exe | False | OK_GUI | terminated | 4.03 | GUI process started and was stopped after smoke timeout |
| DAKE_Year_Age | Dake築年数 | available | 01_apps\DAKE_Year_Age\dist\DakeYear_Age.exe | False | OK_GUI | terminated | 4.02 | GUI process started and was stopped after smoke timeout |
| DAKE_Year_Notice | Dake今年の注意点 | available | 01_apps\DAKE_Year_Notice\dist\DakeYear_Notice.exe | False | OK_GUI | terminated | 4.02 | GUI process started and was stopped after smoke timeout |
| DAKE_Yesterday_Task_Memo | Dake昨日タスクメモ | available | 01_apps\DAKE_Yesterday_Task_Memo\dist\DakeYesterday_Task_Memo.exe | False | OK_GUI | terminated | 4.02 | GUI process started and was stopped after smoke timeout |

## Skipped By Status

| app | display_name | status | note |
| --- | --- | --- | --- |
| DAKE_App_Dashboard | DAKE Dashboard | internal | status=internal |
| DAKE_Approve_Brainz | 承認Brainz | frozen | status=frozen |
| DAKE_BGM_Loop | Dake BGM Loop | frozen | status=frozen |
| DAKE_Brainz_OIKAWA | OIKAWA | frozen | status=frozen |
| DAKE_Brainz_Search | 補助脳BRAINZ | frozen | status=frozen |
| DAKE_Game_ShimarisuRealEstate | Dakeしまりす不動産 | prototype | status=prototype |
| DAKE_HolidayJinja_Post | holiday-jinja 投稿DAKE | internal | status=internal |
| DAKE_Music_Otooku | 音を置く | frozen | status=frozen |
| DAKE_Note_Inbox | DAKE_Note_Inbox | unknown | status=unknown |
| DAKE_QPSC_Dashboard | QPCS Dashboard | internal | status=internal |
| DAKE_Wake_Brainz | DAKE_Wake_Brainz | draft | status=draft |
| DAKE_Web_Dashboard | DAKE Web Dashboard | frozen | status=frozen |
| DAKE_Web_Index | DAKE Web Index | internal | status=internal |
| DAKE_Yukiz_KadouChu | Dakeユキズ稼働中 | draft | status=draft |
| DAKE_YukizBlog_Post | YUKIZ BLOG 投稿DAKE | internal | status=internal |
