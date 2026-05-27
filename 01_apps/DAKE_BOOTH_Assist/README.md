# DakeBOOTHアシスト

## 概要

DakeBOOTHアシストは、DAKEアプリのBOOTH登録作業を補助するアプリです。

`booth_product.txt` と `booth_ready/` を読み取り、商品名・価格・説明文・タグ・登録素材を確認しながら、ログイン済みChromeで既に開いているBOOTHの商品登録または編集画面へ入力補助します。

ログインやreCAPTCHA、パスキー認証、商品登録画面を開く操作は人間がChrome上で行います。DakeBOOTHアシストは開いている編集画面への入力補助だけを行い、公開ボタンは押しません。

DakeBOOTHアシストは、現在開いているBOOTH編集ページへ入力補助します。新しいタブは開きません。

## できること

- `booth_product.txt` 読み取り
- `booth_ready` 確認
- 商品情報のコピー
- ログイン済みChromeの起動補助
- Chrome DevTools Protocol接続による入力補助
- タグを1件ずつ入力してEnterで確定する補助
- 画像/zipの手動設定補助

## やらないこと

- 自動公開
- 自動販売開始
- ログイン情報保存
- reCAPTCHAやパスキー認証の突破
- 商品登録URLへの自動遷移
- 商品管理画面から登録ボタンの自動クリック
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
4. ChromeでBOOTHへログインする
5. BOOTH上で「商品登録」または既存商品の「編集」を人間が開く
6. URLが `https://manage.booth.pm/items/数字/edit` のようになっていることを確認する
7. DakeBOOTHアシストに戻る
8. 「Chrome接続で入力補助」を押す
9. 入力内容を確認する
10. 公開判断は人間が行う

Chromeが自動検出できない場合は、以下を手動で実行してください。

```bat
chrome.exe --remote-debugging-port=9222 --user-data-dir="%LOCALAPPDATA%\DakeBOOTH_Assist\chrome_profile" https://manage.booth.pm/items
```

## 注意事項

BOOTHの画面仕様変更により、自動入力できない場合があります。
その場合はコピー補助として使用してください。
最終公開前には必ず内容を人間が確認してください。

画像とzipは現時点では手動設定を基本とします。
DakeBOOTHアシストはパスコピーとbooth_readyフォルダ表示で補助します。
BOOTH画面が変わるため、無理に自動アップロードしません。
カテゴリと代理購入サービスは、現時点では人間確認を基本とします。
最終公開・保存判断は人間が行います。

Chromeプロファイルは `%LOCALAPPDATA%\DakeBOOTH_Assist\chrome_profile` に作成されます。ログイン状態の保存はChrome側の通常機能に任せ、アプリ側でログインID、パスワード、cookieを読み取りません。

## DAKE_META

```json
{
  "app_key": "dake_booth_assist",
  "display_name": "BOOTHアシスト",
  "launcher_title": "BOOTHアシスト",
  "launcher_description": "BOOTH登録に必要な商品情報と素材を確認し、開いているBOOTH編集画面への入力作業を補助します。",
  "site_title": "DakeBOOTHアシスト",
  "site_description": "DAKEアプリのBOOTH登録作業を、商品情報の確認と開いているBOOTH編集画面への入力補助で止まらず進めるためのアプリです。",
  "update_summary": "開いているBOOTH編集画面への入力補助に変更",
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
- 開いているBOOTH編集画面へ接続して入力作業を補助
- 最終公開は人間確認
- Windows向けexe

## v2候補

- NEW / UPDATE 自動判定
- BOOTH URL未記入検出
- 公開ボタンは押さない方針を維持
- 画像1枚目が booth_thumbnail.jpg か確認
- zip選択前に中身を表示
- 入力後スクリーンショット保存
- BOOTH公開後のURLを booth_product.txt に戻す導線
- frozen / draft / private / experimental を候補から除外
