# Stripe Payment Link Pilot10 Test Result

## Summary

- mode: test
- count: 10
- errors: 0
- live mode used: no
- Secret saved: no

## Items

| id | title | product_id | price_id | payment_link_id | payment_link_url |
| --- | --- | --- | --- | --- | --- |
| dake_pdf_viewer | DakePDF見る | prod_UhNzuV6tkw3Fdr | price_1ThzDDHrsJubFuDOvMOrv0Wz | plink_1ThzDEHrsJubFuDOnxXouGon | https://buy.stripe.com/test_5kQdR9eIh8gfgl25vz0gw00 |
| dake_pdf_reorder | DakePDFページ並べ替え | prod_UhNz5SmMSvinql | price_1ThzDEHrsJubFuDOg51ihuVx | plink_1ThzDFHrsJubFuDOujTBnxC1 | https://buy.stripe.com/test_9B6fZh0Rr2VV6Ksf690gw01 |
| dake_pdf_splitone | DakePDF分割One | prod_UhNzZhwhGDhh8Z | price_1ThzDGHrsJubFuDOgksKSrS5 | plink_1ThzDGHrsJubFuDOtbqv8z4G | https://buy.stripe.com/test_aFa6oHeIh3ZZ6Kse250gw02 |
| dake_image_heictojpg | HEIC→JPG変換 | prod_UhNzXTGmw6S9Gj | price_1ThzDHHrsJubFuDOiIiCwmbX | plink_1ThzDHHrsJubFuDOwmPWFI7p | https://buy.stripe.com/test_5kQ5kD7fP8gf7Owf690gw03 |
| dake_image_topdf | DakeImageToPDF | prod_UhNzqXx9aXHzJB | price_1ThzDIHrsJubFuDOAzOOIyml | plink_1ThzDIHrsJubFuDO8fZojEi6 | https://buy.stripe.com/test_4gM9AT57H7cbc4M3nr0gw04 |
| DAKE_Sticky_Memo | 付箋メモ | prod_UhNz12E7iuJJR8 | price_1ThzDJHrsJubFuDOrc8zRagJ | plink_1ThzDJHrsJubFuDOuwJ8zWJA | https://buy.stripe.com/test_8x200j6bL1RRecU0bf0gw05 |
| DAKE_Mail_Draft | Dakeメール下書き | prod_UhNzSrpik7P0sW | price_1ThzDKHrsJubFuDOsYGzJ9py | plink_1ThzDKHrsJubFuDORuyh7OMO | https://buy.stripe.com/test_eVq00j43D9kjc4Me250gw06 |
| DAKE_Backup | Dakeバックアップ | prod_UhNzvtWiLxQ0zV | price_1ThzDLHrsJubFuDO1FkV06O0 | plink_1ThzDLHrsJubFuDOdNEYd0sB | https://buy.stripe.com/test_aFacN56bL3ZZ3yg7DH0gw07 |
| dake_folder_list | Dakeフォルダ一覧 | prod_UhNzBRhhGpczlU | price_1ThzDMHrsJubFuDOnrYnrAra | plink_1ThzDMHrsJubFuDO7eNaLWve | https://buy.stripe.com/test_6oU6oH8jT3ZZd8Q1fj0gw08 |
| dake_year_age | Dake築年数 | prod_UhNzfigeHWfstC | price_1ThzDNHrsJubFuDOYTGXJsS8 | plink_1ThzDNHrsJubFuDOfVvbTaLF | https://buy.stripe.com/test_fZu3cvbw57cbc4Mf690gw09 |

## Errors

| error |
| --- |
| - |

## Safety Notes

- Stripe Secret Key was read only from the `STRIPE_SECRET_KEY` environment variable.
- The secret value is not written to this file.
- Live mode keys are rejected.
- `ORIGINAL.md`, generated JSON, dake-store-site, and Store production were not updated.
