# Dakeメール下書き State

## Current State

- CSV名簿からOutlook Classicの個別下書きを作る既存機能を維持する。
- Stripe決済後にprivate R2の `desktop/DAKE_Mail_Draft/current.zip` を配布する本番導線が稼働している。
- 新正本主義の `common + app + 製品固有` 参照構造を適用する。
- 利用者向けクレジットは `© Yukihiko Kikuta / PEAKHEADZ` に統一する。
- 2026-11-22に既存onefile build方式で販売exeを再buildし、metadata、launch-check、実GUIを確認した。
- 共通icoをonefile内部へ同梱し、Explorer、GUI titlebar、Windows taskbarの実行中アイコンを実機確認した。
- 新しい3ファイル構成の販売ZIPを検証済み。SHA256は `EEB66EFFD271B8CA847FF598153F633143097B7152BA89FA10A1AF03EB3CE94B`。
- DAKE Storeの現行Stripe Payment Linkを正式販売経路とし、決済後はprivate R2の `desktop/DAKE_Mail_Draft/current.zip` を即時配布する。
- 2026-11-22にStripe本番決済、Store redirect、private R2 download、ZIP download、exe起動を人間が確認した。
- 旧形式Payment Linkは無効化済み。GitHub Releaseの配布exe assetを削除し、公開無料Downloadを終了した。

## Next

- 現行のStripe Product、Price、private-delivery Payment Link、private R2 `current.zip` を維持する。
- BOOTHはDAKE Storeと独立した販売先として維持する。

## Known Issues

- Outlook下書き生成の実機確認には、初期設定済みのWindows版Outlook Classicが必要。
- exeはコード署名していないため、Windowsの確認画面が表示される場合がある。
- Chromeのdownload scan表示は許容し、今回は追加対策を行わない。
