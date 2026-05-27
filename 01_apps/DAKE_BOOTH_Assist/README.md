# DakeBOOTHアシスト

## 概要

DakeBOOTHアシストは、DAKEアプリのBOOTH登録作業を補助するアプリです。

`booth_product.txt` と `booth_ready/` を読み取り、商品名・価格・説明文・タグ・登録素材を確認しながら、ログイン済みChrome上のBOOTH商品登録画面へ入力補助します。

ログインやreCAPTCHA、パスキー認証は人間がChrome上で行います。DakeBOOTHアシストはログイン後の入力補助だけを行い、公開ボタンは押しません。

## できること

- `booth_product.txt` 読み取り
- `booth_ready` 確認
- 商品情報のコピー
- ログイン済みChromeの起動補助
- Chrome DevTools Protocol接続による入力補助
- 画像/zipアップロード補助

## やらないこと

- 自動公開
- 自動販売開始
- ログイン情報保存
- reCAPTCHAやパスキー認証の突破
- BOOTH内部API操作
- 決済情報、売上情報、個人情報の取得

## 初回セットアップ

Playwright Pythonの利用には初回のみ以下が必要です。

```bat
python -m pip install -r requirements.txt
```

専用Chromiumのインストールは不要です。普段のChromeとは別プロファイルのChromeを、remote debugging port `9222` 付きで起動して接続します。

## 使い方

1. DakeBOOTHアシストを起動
2. 対象アプリを選択
3. 「ログイン済みChromeを起動」を押す
4. Chrome上でBOOTHへログインする
5. BOOTH商品管理画面または商品登録画面を開く
6. DakeBOOTHアシストに戻る
7. 「Chrome接続で入力補助」を押す
8. 入力内容を確認する
9. 公開ボタンは人間が最終判断する

Chromeが自動検出できない場合は、以下を手動で実行してください。

```bat
chrome.exe --remote-debugging-port=9222 --user-data-dir="%LOCALAPPDATA%\DakeBOOTH_Assist\chrome_profile" https://manage.booth.pm/items
```

## 注意事項

BOOTHの画面仕様変更により、自動入力できない場合があります。
その場合はコピー補助として使用してください。
最終公開前には必ず内容を人間が確認してください。

Chromeプロファイルは `%LOCALAPPDATA%\DakeBOOTH_Assist\chrome_profile` に作成されます。ログイン状態の保存はChrome側の通常機能に任せ、アプリ側でログインID、パスワード、cookieを読み取りません。

## DAKE_META

```json
{
  "app_key": "dake_booth_assist",
  "display_name": "BOOTHアシスト",
  "launcher_title": "BOOTHアシスト",
  "launcher_description": "BOOTH登録に必要な商品情報と素材を確認し、ログイン済みChromeへの入力作業を補助します。",
  "site_title": "DakeBOOTHアシスト",
  "site_description": "DAKEアプリのBOOTH登録作業を、商品情報の確認とログイン済みChromeへの入力補助で止まらず進めるためのアプリです。",
  "update_summary": "ログイン済みChrome接続によるBOOTH登録補助へ変更",
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
- ログイン済みChromeへ接続して入力作業を補助
- 最終公開は人間確認
- Windows向けexe
