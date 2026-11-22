# Dakeメール下書き State

## Current State

- CSV名簿からOutlook Classicの個別下書きを作る既存機能を維持する。
- Stripe決済後にprivate R2の `desktop/DAKE_Mail_Draft/current.zip` を配布する本番導線が稼働している。
- 新正本主義の `common + app + 製品固有` 参照構造を適用する。
- 利用者向けクレジットは `© Yukihiko Kikuta / PEAKHEADZ` に統一する。
- 2026-11-22に既存onefile build方式で販売exeを再buildし、metadata、launch-check、実GUIを確認した。
- 共通icoをonefile内部へ同梱し、Explorer、GUI titlebar、Windows taskbarの実行中アイコンを実機確認した。
- 新しい3ファイル構成の販売ZIPを検証済み。SHA256は `EEB66EFFD271B8CA847FF598153F633143097B7152BA89FA10A1AF03EB3CE94B`。

## Next

- private R2の `current.zip` を検証済み販売ZIPに揃え、人間がStore購入履歴から再ダウンロードし、ZIP展開とexe起動を確認する。
- 再ダウンロード確認前は、旧Payment Link、GitHub Release、無料Download導線を変更しない。

## Known Issues

- Outlook下書き生成の実機確認には、初期設定済みのWindows版Outlook Classicが必要。
- exeはコード署名していないため、Windowsの確認画面が表示される場合がある。
- GitHub Releaseの既存無料配布は、人間の再ダウンロード確認が終わるまで維持する。
