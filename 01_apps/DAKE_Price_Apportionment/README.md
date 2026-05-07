# Dake価格按分

不動産売買価格を、土地・建物の評価額比率で参考按分する Windows デスクトップアプリです。  
売買契約書や確定申告資料へ転記するための参考値を、すばやく確認できるようにしています。

## 使い方

1. 売買価格（税込）、土地評価額、建物評価額を入力します。
2. 必要に応じて消費税率を確認し、`消費税を計算する` のオン・オフを選びます。
3. `計算する` を押すと、参考計算結果が表示されます。
4. `結果をコピー` で参考計算結果のみをタブ区切りでコピーできます。
5. `印刷用を開く` で参考計算結果のみを白黒HTMLで表示できます。

## 免責事項

本ツールは、入力された評価額に基づく機械的な按分計算を行うものです。  
税務上の適正な区分を保証するものではありません。  
最終的な判断は税理士等の専門家にご確認ください。

## ビルド方法

`build.bat` を実行すると、PyInstaller で `dist/DakePrice_Apportionment.exe` を生成します。

## DAKEシリーズ

シンプルそれDAKEシリーズ / 止まらない、迷わない、すぐ終わる。

## 共通仕様確認

- 2026-05-06: DAKE共通仕様に基づき、フォント、ヘッダー、フッター、UI_TEXT、共通アイコン、build.bat、.gitignore を確認しました。
- 2026-05-06: フッターリンクは通常時を補助文字色、ホバー時のみアクセント色に調整しました。
- 2026-05-06: コピー機能を参考計算結果のタブ区切りコピーに統合し、印刷HTMLを参考計算結果のみへ簡素化しました。

## DAKE_META

```json
{
  "app_key": "dake_price_apportionment",
  "display_name": "Dake価格按分",
  "launcher_title": "価格按分",
  "launcher_description": "売買価格を土地・建物評価額比率で参考按分します。",
  "site_title": "Dake価格按分",
  "site_description": "不動産売買価格を土地・建物の評価額比率で参考按分し、転記しやすい結果を表示するWindows向けアプリです。",
  "update_summary": "READMEメタ情報とRelease本文を整備。スクリーンショットをassets/screenshot.webpに作成。",
  "folder_name": "DAKE_Price_Apportionment",
  "exe_name": "DakePrice_Apportionment.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## RELEASE_BODY

- 不動産価格按分アプリ
- 土地・建物評価額比率で計算
- 結果コピー・印刷用表示に対応
- Windows向けexe
