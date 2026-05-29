# Dakeマンション工程表

マンション管理会社へリフォーム許可申請を行うときに提出する、A3横の工事工程表PDFを作成するDAKEシリーズのWindows向けアプリです。

## アプリ説明

- 引渡し日から、着工日と45平日の完工日を自動計算します。
- 固定の工事項目に、提出用として自然に見える初期日程を自動入力します。
- 各工程の開始日・終了日は手動で修正できます。
- 出力はA3横1枚のPDFです。

## 使い方

1. 現場名、会社支店名、担当者名、連絡先を入力します。
2. 引渡し日を `YYYY-MM-DD` 形式で入力します。
3. 「日程を自動入力」を押します。
4. 必要に応じて各工程の開始日・終了日を修正します。
5. 「保存先を選ぶ」でPDFの保存先を指定します。
6. 「PDFを作成」を押します。

## 注意事項

本工程表は、マンション管理会社への提出・申請補助を目的とした参考工程表です。実際の施工工程・管理規約・申請条件に応じて、施工会社・管理会社へ確認してください。

## DAKE_META

```json
{
  "app_key": "dake_mansion_schedule",
  "display_name": "マンション工程表",
  "launcher_title": "マンション工程表",
  "launcher_description": "管理会社提出用のA3横工程表を作成します。",
  "site_title": "マンション工程表",
  "site_description": "リフォーム許可申請用の工程表を、A3横PDFで作成します。",
  "update_summary": "管理会社提出用の工程表作成に対応",
  "folder_name": "DAKE_Mansion_Schedule",
  "exe_name": "DakeMansion_Schedule.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## RELEASE_BODY

- 管理会社提出用のA3横工程表PDFを作成
- 引渡し日から着工日・完工日を自動計算
- 45平日ベースの初期工程を自動入力
- 各工事項目の開始日・終了日を手動修正可能
- 土日を除いた平日ベースのガントチャート風PDF出力
