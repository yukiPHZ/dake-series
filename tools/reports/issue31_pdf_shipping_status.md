# Issue #31 PDF Shipping Status

Date: 2026-09-23 (JST)

## Scope and decision

Four PDF applications were handled independently. `ORIGINAL.md` remains the source of truth. No new BOOTH listing or Stripe Product/Price/Payment Link was created, and no existing v1.0.0 Release was changed. `DAKE_PDF_Extract`, `DAKE_Image_BatchPDF`, and soredake.com are outside this issue.

| App | v1.0.1 publication | Final gate | Decision |
|---|---|---|---|
| DAKE_PDF_OverviewRename | Release, existing BOOTH item, Store, dakeapp.com verified | 57 tests and packaged synthetic/GUI checks passed | CLOSED |
| DAKE_PDF_Merge | Existing Release is public; Store and dakeapp.com point to v1.0.1 | Explorer-to-app physical DnD, held-edge distant drag/autoscroll, final packaged review unverified | HOLD: do not replace BOOTH ZIP yet |
| DAKE_PDF_SplitSelect | Improved code and shared onedir tooling are on main; no v1.0.1 Release created | PyMuPDF redistribution basis unresolved; packaged physical GUI checks outstanding | HOLD: no v1.0.1 external distribution |
| DAKE_PDF_SplitOne | Release, existing BOOTH item, Store, dakeapp.com verified | Extracted ZIP CLI/GUI passed; Explorer-to-app physical DnD and normal-close final check unverified | HOLD for final CLOSED declaration; publication already live |

## Shared onedir shipping

PR #32 merged as `25cdb0a2ed6cce6a4d0e6dafebdda67aac267afd`. The onefile/onedir-aware shipping utilities now discover nested EXEs, recursively package runtime siblings, verify ZIP entries by SHA-256, and support Launcher/release-capture discovery. Focused shipping tests passed (3/3). An older pipeline test requires ignored Pack ZIP files absent from isolated worktrees; it is not evidence that these four PDF packages failed.

## OverviewRename

- Implementation PR #25: `652a45e38328b17c15d653b0a528fc7fb810cc91`; release preparation PR #33: `d9e79825ea0a5c3c99eb4fb67eeab26ff37d6557`; publication record PR #34: `42f2a73d1538d28c3b084182713698abfdde2b25`.
- [GitHub Release v1.0.1](https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_PDF_OverviewRename_v1.0.1): ZIP 34,140,032 bytes, SHA-256 `2b84ba6cdb14f33e5b3ecb0aa0558c24a3fa9b430a9cd3d30c1eabb3de1b7541`. Local EXE SHA-256 `649f3470a5ea460b594332704cd6e21ad1d89a8114770c1716f37d597a3b5c91`.
- [Existing BOOTH item](https://peakheadz.booth.pm/items/8798555): 500 JPY; the effective downloadable is new ZIP ID `9730893`. The prior ZIP remains stored but is not selected. The new images are first in display order; older images were retained.
- [dakeapp.com detail](https://dakeapp.com/apps/pdf-overview-rename/) shows the v1.0.1 Release link; [Store detail](https://store.dakeapp.com/product/?id=dake_pdf_overview_rename) shows v1.0.1, Stripe and BOOTH links. The existing Stripe link was retained.
- Tests: 57 passed; synthetic 1/48/100/300-page checks, source hashes, ZIP roundtrip, packaged GUI and screenshot review passed. No real document was used.

## Merge

- Acceptance PR #26: `ec63604c29f4a20decc238a2a05a47d5b01bc31b`; release-preparation PR #28: `1b14c3fa71ff57dd666c4aeb56508062424d9505`.
- [Existing public v1.0.1 Release](https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_PDF_Merge_v1.0.1) contains `DakePDF_Merge.exe`, 51,906,304 bytes, SHA-256 `ee5a78072fd719df2cebb18fe7b8e4ea0aba196f2425a120ef2cf24454fd90b7`. A fresh public download matched both. The EXE left in the original repo's ignored `dist` has a different SHA; it must not be used as the formal asset.
- [Existing BOOTH item](https://peakheadz.booth.pm/items/8448196) remains 500 JPY and was not altered while the physical DnD gate is open. [Store](https://store.dakeapp.com/product/?id=dake_pdf_merge) shows v1.0.1, the existing Stripe link and BOOTH link. [dakeapp.com](https://dakeapp.com/apps/pdf-merge/) points Download to v1.0.1, though its update date remains old.
- The ignored `booth_ready/DakePDF_Merge.zip` is **not** the formal v1.0.1 package: its embedded EXE SHA-256 is `537d7fcad484024b7bf74872cc4f36e09370c89be86cc62ca315f512a88ec89f`, unlike the public Release asset. It must not be uploaded. A separate unpublished replacement candidate was built from the downloaded, SHA-matched Release EXE and existing README/notice; ZIP SHA-256 `906327af209961cc88764c2b5077f1a7fc8a444f4927dddff82a1519f3502947`, 51,060,039 bytes, with all entries SHA-verified. It remains outside Git and BOOTH pending the physical gate.
- The downloaded formal EXE opens a Windows GUI. The browser-only Computer Use surface in this task cannot bind native Windows windows, so physical Explorer DnD and held-edge drag could not be asserted. Thirty synthetic one-page PDFs were prepared outside Git for the narrowly scoped final check. Do not claim CLOSED or replace the BOOTH ZIP until that check passes.

## SplitSelect

- Performance PR #29: `4a354dabc88368326a49de9ebaf57a970b29a5d5`. The shared onedir support is available, but no v1.0.1 Release was created.
- `requirements.txt` pins `PyMuPDF==1.24.10`; `build.bat` collects `fitz`; `docs/PERFORMANCE_REFRESH_REPORT.md` records AGPL-3.0 metadata and unresolved distribution conditions. No repository evidence was found for an AGPL-compliant DAKE distribution decision, a commercial redistribution agreement, or a replacement renderer with completed license audit. [Artifex's licensing guidance](https://artifex.com/licensing) describes AGPL and commercial paths. This report makes no legal clearance decision.
- Keep the new version's external Release/BOOTH distribution on HOLD. Existing products or old Releases were not modified or retracted. Physical Explorer DnD, packaged Shift+click/F5, actual DPI 125/150/200%, and Launcher live UI are also not claimed as passed.

## SplitOne

- Implementation PR #30: `a23b6d1dee886e45ec519b08162e10983db2c004`; formal release PR #35: `31e3df032f9ed9f3195c163c8de5bd0d5355013e`; generated Store data PR #36: `6d19da346ffbb59aedd12f6ff3df6958d360c20e`.
- [GitHub Release v1.0.1](https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_PDF_SplitOne_v1.0.1): full onedir ZIP, 35,700,281 bytes, SHA-256 `dd65ccad033fdaf1bf9d8075c0d21b503d37aa6f1324ffaa90e4a4acbb2499b0`. The BOOTH upload used the same local ZIP; the Release asset digest matches it. EXE SHA-256 `78d6a820eb5f491eb685e35824575ff388edd8ce40ce928025a2648bb691e15c`.
- ZIP has 1,140 SHA-verified entries including the runtime folder, `_internal`, README, notice, and third-party notices/licenses. Extracted ZIP CLI split a synthetic three-page PDF into three one-page outputs, preserved the source hash, and opened a visible GUI. App and shipping tests passed (9/9). No PyMuPDF payload was found in this ZIP.
- [Existing BOOTH item](https://peakheadz.booth.pm/items/8448207): 500 JPY; effective downloadable is new ZIP ID `9731063`, old ZIP deselected. Public description includes the F5/Ctrl+O change and v1.0.1 Release link.
- [Store detail](https://store.dakeapp.com/product/?id=dake_pdf_splitone) shows v1.0.1, 500 JPY, existing Stripe link `https://buy.stripe.com/28E14ndEd7cb0m47DH0gw0z`, and BOOTH link. [dakeapp.com detail](https://dakeapp.com/apps/pdf-split-one/) shows the v1.0.1 Download link and update. The Store sync preserved 56 items, `source_policy`, `do_not_edit`, and the payment-status distribution (stripe_ready 54, booth_only 1, preparing 1). The site test suite passed 17/17 after correcting a stale fixed page-count assertion.
- A final physical Explorer-to-app drop and normal-close check remains unverified. Publication was not rolled back; this verification gap blocks a CLOSED declaration.

## Remaining human-only checks

Native Windows UI access is unavailable to the connected Computer Use tool in this task. These are the only requested human actions; authentication, payment, and publication do not require user intervention.

1. Merge: open the SHA-verified v1.0.1 EXE, drag prepared synthetic PDFs from Explorer into the app, then hold a card at the list's lower edge long enough to confirm auto-scroll and correct reorder. Report DnD and edge-drag PASS/FAIL separately.
2. SplitOne: open the packaged onedir EXE, drop a synthetic PDF from Explorer, verify a one-page output, then close normally. Report DnD and close PASS/FAIL.

If Merge passes, replace only its existing BOOTH downloadable with a verified v1.0.1 ZIP, keep 500 JPY and existing Stripe configuration, then confirm the live page. If SplitOne passes, close its remaining verification gate. SplitSelect remains HOLD independently until documented license clearance and remaining GUI checks.
