# DakeImage_Receiver

スマホから画像受け取るDAKE は、PCに表示したQRコードをスマホで読み取り、画像をこのPCへ送るための入口アプリです。

## 使い方

1. アプリを起動
2. QRコードをスマホで読む
3. 画像を選んで送信
4. PC側の保存フォルダを確認

## 注意

- 同じWi-Fi内で使う
- HEICは変換しない
- 画像を受け取るだけのアプリ

## ビルド方法

`build.bat` を実行すると、`dist/DakeImage_Receiver.exe` を生成します。

## アイコン

DAKE共通アイコン `..\..\02_assets\dake_icon.ico` を使用します。アプリ個別アイコンは作成しません。
