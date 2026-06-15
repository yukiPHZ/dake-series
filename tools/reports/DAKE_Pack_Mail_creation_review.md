# DAKE Mail Pack Creation Review

## 目的

Phase 17B-1 creates `DAKE_Pack_Mail` as a real Pack product and places it on the product release pipeline from the beginning.

This phase stops at BOOTH registration readiness. BOOTH publication, Stripe API, Stripe Secret, Store generation, Store sync, and Store production changes were not performed.

## 商品提案

- pack_id: `DAKE_Pack_Mail`
- 商品名: DAKE メール準備パック
- 価格: 780円
- currency: JPY
- category: メール実務Pack
- pack_type: mail_pack
- target_platform: Windows
- catch: 集める、整える、下書きにする。送信前までを静かにまとめるPack。

## 収録アプリ

| folder | role | source zip | size | sha256 |
|---|---|---|---:|---|
| DAKE_Mail_List | Outlookの.msgメールから会社名・氏名・メールアドレスをCSV化する | `DakeMail_List.zip` | 32925063 | `1012c378bb28d62b1c7fec9d50f5de966c606a9ff44a74b60e20854bfebf8bc1` |
| DAKE_Mail_Address_Format | 名前付き表記、改行、セミコロン等が混在したメールアドレスを抽出・整形する | `DakeMail_Address_Format.zip` | 10106704 | `3fe241b58b95de37ae4b9d08b9fb6096e485fbf72d28d9926cb43120267d53de` |
| DAKE_Mail_Draft | CSV名簿からOutlookの個別下書きを作成する | `DakeMail_Draft.zip` | 13902128 | `d6d691d890eb8d5b3fa903765f756a6e66d666400a228eebdb3135174a55724e` |

All three source apps are `status=available` and have BOOTH URL, GitHub Release URL, and formal `booth_ready/*.zip` files.

## 除外アプリ

- `DAKE_Mail_AllStaff`: 組織固有の用途を含むため、汎用販売Packには収録しない。
- `DAKE_Mail_Kikuta`: 利用者固有の用途を含むため、汎用販売Packには収録しない。

## 正本

Created:

- `04_packs/DAKE_Pack_Mail/ORIGINAL.md`

Initial state:

- `status: preparing`
- `payment_status: preparing`
- `booth_url: 未設定`
- `stripe_payment_link: 未設定`

## Pack ZIP

- ZIP name: `DAKE_Pack_Mail.zip`
- ZIP path: `04_packs/DAKE_Pack_Mail/pack_ready/DAKE_Pack_Mail.zip`
- ZIP size: 56950902 bytes
- ZIP SHA256: `dfc972b91529161bbf688fbe4fb5bf91b5e27956afe058486a0b5d79ab293ad4`

ZIP contents:

```txt
README.txt
注意事項.txt
apps/DAKE_Mail_List/DakeMail_List.zip
apps/DAKE_Mail_Address_Format/DakeMail_Address_Format.zip
apps/DAKE_Mail_Draft/DakeMail_Draft.zip
```

The Pack ZIP is ignored by Git via `**/pack_ready/*.zip`.

## Manifest

Created:

- `04_packs/DAKE_Pack_Mail/pack_manifest.json`

Manifest records Pack id, title, version, price, currency, included app ZIP names, source ZIP sizes, source ZIP SHA256 values, Pack ZIP path, Pack ZIP size, Pack ZIP SHA256, and generated timestamp.

Hash consistency was checked against the generated ZIP and the three source app ZIPs.

## README / 注意事項

Created:

- `04_packs/DAKE_Pack_Mail/README.md`
- `04_packs/DAKE_Pack_Mail/pack_ready/README.txt`
- `04_packs/DAKE_Pack_Mail/pack_ready/注意事項.txt`

The notices explicitly state:

- Windows target
- Dakeメール下書き requires Microsoft Outlook Classic
- New Outlook / Web Outlook may not work
- Mail is not automatically sent
- Draft content must be checked by a human before sending
- Spam, unsolicited sending, legal violations, and terms violations are prohibited

## BOOTH商品情報

Created:

- `04_packs/DAKE_Pack_Mail/booth_product.txt`
- `04_packs/DAKE_Pack_Mail/pack_ready/booth_product.txt`

Tags:

- メール
- Outlook
- CSV
- Windows
- 実務
- 仕事効率化
- 下書き
- ツール

## BOOTH画像

Created:

- `04_packs/DAKE_Pack_Mail/assets/booth_thumbnail.jpg`
- `04_packs/DAKE_Pack_Mail/pack_ready/booth_thumbnail.jpg`

Image size: 1200 x 1200.

The image shows:

- DAKE Pack
- メール準備パック
- 集める / 整える / 下書きにする
- Outlook / CSV / Windows
- 自動送信ではなく人間確認が必要であること

## Manual Delivery

The Pack source includes the same manual delivery policy used by existing Packs:

- `purchase_delivery_method: manual_email_private_download`
- `purchase_delivery_ready: yes`
- `delivery_sla: 決済確認後、次営業日以内`
- `delivery_rule: 00_core/DAKE_PACK_MANUAL_DELIVERY_RULE.md`

Stripe sale is not enabled in this phase.

## Pipeline Stage

Command:

```powershell
python tools\release_product_pipeline.py DAKE_Pack_Mail
```

Observed:

- current_stage: `BOOTH_REGISTRATION_PENDING`
- pack_zip_ready: true
- booth_assets_ready: true
- booth_url: missing
- stripe_payment_link: missing
- next_action: register product on BOOTH and record the product URL

Pipeline status was saved to:

- `tools/reports/release_artifacts/DAKE_Pack_Mail/pipeline_status.json`
- `tools/reports/release_artifacts/DAKE_Pack_Mail/pipeline_status.md`

## Existing Product Regression

Observed stages:

- `DAKE_Pack_Document`: `RELEASE_COMPLETE`
- `DAKE_Pack_Memo`: `RELEASE_COMPLETE`
- `dake_pdf_viewer`: `LEGACY_COMPLETE`
- `video_shorts_cut`: `PREPARING_BLOCKED`

Store generated data was not regenerated. Existing Store counts remain:

- total: 53
- stripe_ready: 52
- booth_only: 0
- preparing: 1

## Security Review

Checked:

- Pack ZIP contains only README, 注意事項, and three app ZIP files.
- No source code, build directory, dist directory, spec file, venv, `.env`, Secret, API key, personal email list, real CSV, Outlook account data, logs, or temporary files were intentionally included.
- The source app ZIPs contain release artifacts only.
- Pack ZIP path is repo-relative in manifest and source.
- Pack ZIP is Git ignored.
- No Stripe API call was made.
- No Stripe Secret was used.
- No BOOTH API call was made.
- Store generated JSON and dake-store-site were not changed.

## Operator Handoff

Created:

- `tools/reports/release_artifacts/DAKE_Pack_Mail/booth_registration_handoff.md`

After BOOTH registration, record the URL with:

```powershell
python tools\record_booth_registration.py DAKE_Pack_Mail --booth-url <BOOTH商品URL>
python tools\record_booth_registration.py DAKE_Pack_Mail --booth-url <BOOTH商品URL> --apply --confirm-product-id DAKE_Pack_Mail --confirmation-text "RECORD BOOTH REGISTRATION DAKE_Pack_Mail"
```

## Conclusion

READY_FOR_BOOTH_REGISTRATION
