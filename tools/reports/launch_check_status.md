# DAKE launch-check Status

Generated: 2026-05-31 11:06:10

## Summary

- checked: 57
- available checked: 48
- supported: 10
- unsupported: 38
- timeout: 0
- needs_check: 0
- excluded: 9

## Status Counts

- available: 48
- frozen: 3
- draft: 3
- experimental: 1
- private: 0
- internal: 2
- unknown: 0

## Notes

- Inventory only. Do not add launch-check to every app in this pass.
- DAKE_BOOTH_Assist had a previous timeout record, but current source and dist exe checks returned successfully.
- status: available is the priority group for formal shipping checks; other statuses are listed separately.

## Available Apps

| app | status | classification | result | output excerpt |
| --- | --- | --- | --- | --- |
| DAKE_App_Doko | available | supported | exit 0 |  |
| DAKE_Backup | available | supported | exit 0 | LAUNCH CHECK OK |
| DAKE_BOOTH_Assist | available | supported | exit 0 | ����DOM�Ɉ��S��file input���Ȃ����ߎ蓮�ē�: ���i�摜<br>����DOM�Ɉ��S��file input���Ȃ����ߎ蓮�ē�: zip�t�@�C��<br>DakeBOOTH�A�V�X�g launch-check OK<br>apps_root=C:\Users\yukiz\devlop\DAKE_series\01_ap |
| DAKE_Column_Memo | available | unsupported | not implemented |  |
| DAKE_Document_Cover | available | unsupported | not implemented |  |
| DAKE_FAX_Cover | available | unsupported | not implemented |  |
| DAKE_Folder_List | available | unsupported | not implemented |  |
| DAKE_Git_Memo | available | supported | exit 0 |  |
| DAKE_HolidayJinja_Post | available | unsupported | not implemented |  |
| DAKE_Image_HEICtoJPG | available | unsupported | not implemented |  |
| DAKE_Image_iPhoneToPC | available | unsupported | not implemented |  |
| DAKE_Image_PasteA4 | available | unsupported | not implemented |  |
| DAKE_Image_Receiver | available | unsupported | not implemented |  |
| DAKE_Image_Resize | available | unsupported | not implemented |  |
| DAKE_Image_ToPDF | available | unsupported | not implemented |  |
| DAKE_Launcher | available | unsupported | not implemented |  |
| DAKE_Mail_Address_Format | available | supported | exit 0 | DAKE_Mail_Address_Format launch-check OK |
| DAKE_Mail_AllStaff | available | supported | exit 0 | DAKE_Mail_AllStaff launch-check OK |
| DAKE_Mail_Draft | available | supported | exit 0 |  |
| DAKE_Mail_Kikuta | available | unsupported | not implemented |  |
| DAKE_Mail_List | available | unsupported | not implemented |  |
| DAKE_Maji_Memo | available | unsupported | not implemented |  |
| DAKE_Mansion_Schedule | available | supported | exit 0 | Dake�}���V�����H���\ launch-check OK<br>pdf=C:\Users\yukiz\devlop\DAKE_series\01_apps\DAKE_Mansion_Schedule\tmpt421415a\mansion_schedule.pdf<br>a3_landscape_one_page=OK<br>calendar_days=45  |
| DAKE_PDF_CheckStamp | available | unsupported | not implemented |  |
| DAKE_PDF_Compress | available | supported | exit 0 |  |
| DAKE_PDF_Crop | available | unsupported | not implemented |  |
| DAKE_PDF_LookHere | available | unsupported | not implemented |  |
| DAKE_PDF_Marker | available | unsupported | not implemented |  |
| DAKE_PDF_Merge | available | unsupported | not implemented |  |
| DAKE_PDF_Merge_Mini | available | unsupported | not implemented |  |
| DAKE_PDF_Rename | available | unsupported | not implemented |  |
| DAKE_PDF_Reorder | available | unsupported | not implemented |  |
| DAKE_PDF_SplitOne | available | unsupported | not implemented |  |
| DAKE_PDF_SplitSelect | available | unsupported | not implemented |  |
| DAKE_PDF_ToImages | available | unsupported | not implemented |  |
| DAKE_PDF_Viewer | available | unsupported | not implemented |  |
| DAKE_Price_Apportionment | available | unsupported | not implemented |  |
| DAKE_Price_FixedTax | available | unsupported | not implemented |  |
| DAKE_Reform_Progress | available | supported | exit 0 | Dake���t�H�[���i���Ǘ� launch-check OK<br>pdf=C:\Users\yukiz\devlop\DAKE_series\01_apps\DAKE_Reform_Progress\tmptfkf6u1o\���t�H�[���i���J�����_�[_�e�X�g���t�H�[��_20260529-20260713.pdf |
| DAKE_Screen_WebP | available | unsupported | not implemented |  |
| DAKE_Screenshot_Print | available | unsupported | not implemented |  |
| DAKE_Sticky_Memo | available | unsupported | not implemented |  |
| DAKE_TwoPerson_Memo | available | unsupported | not implemented |  |
| DAKE_Work_Calendar | available | unsupported | not implemented |  |
| DAKE_Year_Age | available | unsupported | not implemented |  |
| DAKE_Year_Notice | available | unsupported | not implemented |  |
| DAKE_Yesterday_Task_Memo | available | unsupported | not implemented |  |
| DAKE_YukizBlog_Post | available | unsupported | not implemented |  |

## Non-shipping Apps

| app | status | classification | result | output excerpt |
| --- | --- | --- | --- | --- |
| DAKE_App_Dashboard | internal | excluded | non-shipping; source launch-check exit 0 | LAUNCH CHECK OK |
| DAKE_Approve_Brainz | frozen | excluded | non-shipping; source launch-check exit 0 | LAUNCH CHECK OK |
| DAKE_BGM_Loop | frozen | excluded | not implemented |  |
| DAKE_Brainz_OIKAWA | experimental | excluded | non-shipping; source launch-check exit 0 | LAUNCH CHECK OK |
| DAKE_Brainz_Search | draft | excluded | non-shipping; source launch-check exit 0 | LAUNCH CHECK OK |
| DAKE_Music_Otooku | frozen | excluded | non-shipping; source launch-check exit 0 |  |
| DAKE_Wake_Brainz | draft | excluded | non-shipping; source launch-check exit 0 | LAUNCH CHECK OK<br>config_exists=False<br>web_port=8766 |
| DAKE_Web_Dashboard | internal | excluded | non-shipping; source launch-check exit 0 | LAUNCH CHECK OK: sites=31 api_review=7 dirty_sites=1 git_errors=5 |
| DAKE_Yukiz_KadouChu | draft | excluded | non-shipping; source launch-check exit 0 | {<br>  "app": "Dakeユキズ稼働中",<br>  "version": "0.1.0",<br>  "exe_name": "DakeYukiz_KadouChu.exe",<br>  "app_root": "C:\\Users\\yukiz\\devlop\\DAKE_series\\01_apps\\DAKE_Yukiz_KadouChu",<br>  "cli": |
