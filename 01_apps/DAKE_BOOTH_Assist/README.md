# DakeBOOTHアシスト

## 概要

DakeBOOTHアシストは、DAKEアプリのBOOTH登録作業を補助するアプリです。

`booth_product.txt` と `booth_ready/` を読み取り、商品名・価格・説明文・タグ・登録素材を確認しながら、BOOTH管理画面への入力作業をPlaywrightで補助します。

## できること

- `booth_product.txt` 読み取り
- `booth_ready` 確認
- 商品情報のコピー
- BOOTH管理画面を開く
- Playwrightによる入力補助
- 画像/zipアップロード補助

## やらないこと

- 自動公開
- 自動販売開始
- ログイン情報保存
- BOOTH内部API操作

## 初回セットアップ

Playwright利用には初回のみ以下が必要です。

```bat
python -m pip install -r requirements.txt
python -m playwright install chromium
```

## 注意事項

BOOTHの画面仕様変更により、自動入力できない場合があります。
その場合はコピー補助として使用してください。
最終公開前には必ず内容を人間が確認してください。

ログイン済みセッションを利用するため、Playwrightのプロファイルは `playwright_profile/` に保存される場合があります。アプリ側でログインIDやパスワードは保存しません。

## DAKE_META

```json
{
  "app_key": "dake_booth_assist",
  "display_name": "BOOTHアシスト",
  "launcher_title": "BOOTHアシスト",
  "launcher_description": "BOOTH登録に必要な商品情報と素材を確認し、入力作業を補助します。",
  "site_title": "DakeBOOTHアシスト",
  "site_description": "DAKEアプリのBOOTH登録作業を、商品情報の確認とPlaywright入力補助で止まらず進めるためのアプリです。",
  "update_summary": "BOOTH登録補助アプリを追加",
  "folder_name": "DAKE_BOOTH_Assist",
  "exe_name": "DakeBOOTH_Assist.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## RELEASE_BODY

- DakeBOOTHアシスト
- BOOTH登録用の商品情報と素材を確認
- Playwrightで入力作業を補助
- 最終公開は人間確認
- Windows向けexe
