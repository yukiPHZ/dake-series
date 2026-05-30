# メール整形

貼り付けたメールアドレスらしい文字列だけを抽出し、重複を除いてカンマ区切りに整える単機能DAKEアプリです。

## 使い方

1. 上の入力欄に宛先候補をまとめて貼り付けます。
2. 「整形する」を押します。
3. 結果欄のカンマ区切り文字列を「コピー」でコピーします。

## 仕様

- カンマ、改行、セミコロン、名前付き表記が混在していてもメールアドレス形式らしい文字列だけを抽出します。
- 重複判定は小文字比較で行い、貼り付け順を維持します。
- 出力はカンマ区切りで、前後空白を除いた抽出時の表記を使います。
- メールアドレスが0件の場合は警告を表示します。
- 自動送信、Outlook操作、mailto起動、CSV取込、Excel操作、アドレス帳、保存機能はありません。

## 例

```text
田中 tanaka@example.co.jp, suzuki@example.co.jp
sato@example.co.jp; yamada@example.co.jp
```

```text
tanaka@example.co.jp,suzuki@example.co.jp,sato@example.co.jp,yamada@example.co.jp
```

## 検証

```bat
python -m py_compile main.py
python main.py --launch-check
build.bat
dist\DakeMail_Address_Format.exe --launch-check
```

## DAKE_META

```json
{
  "app_key": "DAKE_Mail_Address_Format",
  "display_name": "メール整形",
  "launcher_title": "メール整形",
  "launcher_description": "貼り付けたメールアドレスを、カンマ区切りに整えます。",
  "site_title": "メール整形",
  "site_description": "名前付き表記、改行、セミコロン混在のメールアドレスをまとめて整形します。",
  "update_summary": "メールアドレスの貼り付け整形に対応",
  "folder_name": "DAKE_Mail_Address_Format",
  "exe_name": "DakeMail_Address_Format.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## RELEASE_BODY

メール整形アプリ
貼り付けた宛先を自動抽出
カンマ区切りでコピー
Windows向けexe
