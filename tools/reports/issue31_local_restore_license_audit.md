# Issue #31 local restoration and license audit

Date: 2026-09-24 (JST)

## Canonical local workspace

- Canonical root: `C:/Users/yukiz/devlop`.
- DAKE formal repo: `C:/Users/yukiz/devlop/DAKE_series`.
- PR #42 (`Define canonical local workspace and cleanup policy`) merged as `c7a8504bdc7615bcba634d1761107fca9200a202` and fast-forwarded into the formal repo.
- The formal `dake-store-site` and `dakeapp-site` repos were also fast-forwarded to their published main branches. Worktrees used for this issue were checked for clean status and main ancestry before removal; see `issue31_worktree_cleanup.json`.

## Formal artifacts

| App | Formal artifact | Verification |
|---|---|---|
| Merge | `01_apps/DAKE_PDF_Merge/booth_ready/DakePDF_Merge.zip` | 51,060,039 bytes; ZIP SHA-256 `906327af209961cc88764c2b5077f1a7fc8a444f4927dddff82a1519f3502947`; embedded EXE SHA-256 `ee5a78072fd719df2cebb18fe7b8e4ea0aba196f2425a120ef2cf24454fd90b7`, matching the published Release and formal `dist` EXE |
| OverviewRename | `01_apps/DAKE_PDF_OverviewRename/booth_ready/DakePDF_OverviewRename.zip` and `dist` | ZIP SHA-256 `2b84ba6cdb14f33e5b3ecb0aa0558c24a3fa9b430a9cd3d30c1eabb3de1b7541`; source and formal runtime files matched by SHA-256 |
| SplitOne | `01_apps/DAKE_PDF_SplitOne/booth_ready/DakePDF_Split_One.zip` and onedir `dist` | ZIP SHA-256 `dd65ccad033fdaf1bf9d8075c0d21b503d37aa6f1324ffaa90e4a4acbb2499b0`; 1,135 runtime files copied and hash-verified; obsolete top-level onefile EXE removed after its known SHA matched |

The temporary `issue31-merge-artifact` contained only the verified duplicate ZIP/EXE and 30 synthetic PDFs. It was removed after formal ZIP/EXE verification (102,979,273 bytes). The 13 clean, merged worktrees listed in `issue31_worktree_cleanup.json` were removed with `git worktree remove` and pruned (1,033,989,706 bytes measured before removal). Unique Merge test screenshot and JSON summaries were preserved in `issue31_merge_evidence/`; synthetic PDF inputs and duplicate binaries were not retained.

The LookHere quality worktree and an older temporary clone with LookHere modifications were retained because they are outside Issue #31 and may contain separate user work. Other repositories under `devlop/_worktrees` were not touched.

## BOOTH publication and duplicate product

- Official Merge item: [#8448196](https://peakheadz.booth.pm/items/8448196). `ORIGINAL.md`, both product text views, DAKE generated JSON, Store generated JSON, dakeapp.com generated JSON, and Issue #31 consistently reference it.
- On 2026-09-24 the user reported replacing the official item's downloadable with the v1.0.1 ZIP. The public page returned HTTP 200, showed `DakePDF結合`, 500 JPY, and an add-to-cart control. Seller UI access failed after the replacement; the effective downloadable ID and remote hash could not be independently read. The old ID `9647349` was observed before the user's update only.
- [#8386233](https://peakheadz.booth.pm/items/8386233) is a distinct public 500 JPY product of the same name. Its public copy is older and lacks the new small-screen/save-safety description. This is a duplicate-product review item, not the canonical BOOTH URL. Purchase history and customer impact should be reviewed before any action; neither product was removed or unpublished.
- Merge status: **CLOSED under the user's acceptance and confirmation of the BOOTH ZIP replacement**, with independent post-save downloadable selection retained as a verification note. OverviewRename and SplitOne remain CLOSED; SplitSelect remains on license HOLD. Issue #31 may stay open for the HOLD.

## PyMuPDF cross-app audit

The local PyInstaller CArchive and embedded PYZ tables were inspected without executing the binaries. Raw paths and payload lists are in `issue31_pymupdf_payload_audit.jsonl`.

| App | Runtime source | Observed bundle | Redistribution evidence in repo |
|---|---|---|---|
| Merge | `main.py` lazily imports `fitz` for PDF page preview | Published v1.0.1 EXE contains `fitz` / `pymupdf` modules and `_mupdf.pyd` / `mupdfcpp64.dll`; embedded version constant `1.24.10` | No commercial agreement or documented AGPL compliance path found |
| SplitSelect | `pdf_backend.py` imports `fitz` to open and render preview pages | Local EXE contains `PyMuPDF-1.24.10.dist-info`, PyMuPDF native binaries; metadata declares `GNU AFFERO GPL 3.0` and PyMuPDFb 1.24.10 | `docs/PERFORMANCE_REFRESH_REPORT.md` explicitly leaves the distribution basis unresolved |
| SplitOne / OverviewRename | No PyMuPDF runtime dependency found for this release | Inspected formal EXEs show no PyMuPDF payload | Not applicable to PyMuPDF for these artifacts |
| Other PDF apps | CheckStamp, Compress, Crop, Extract, Insert, LookHere, Marker, Merge_Mini, Reorder, ToImages, Viewer have PyMuPDF payloads in local EXEs | See raw audit for individual CArchive evidence | No repo-wide commercial license or AGPL compliance decision found |

[PyMuPDF's licensing documentation](https://pymupdf.readthedocs.io/en/latest/about.html#license-and-copyright) states that PyMuPDF and MuPDF are offered under AGPL and commercial licensing. The observed technical facts do **not** provide a legal basis for clearing Merge while holding only SplitSelect. The different publication states reflect the existing user release decision, not an audited license distinction. A documented license route and legal review are needed for SplitSelect and the already distributed PyMuPDF applications. No existing product was withdrawn or altered by this audit.

## Remaining verification

- Read the post-save BOOTH file selection for official item #8448196 when the seller UI is available; confirm only the v1.0.1 ZIP is effective. This is a publication evidence gap, not a reopened acceptance gate.
- Review purchase history for duplicate item #8386233 before choosing how to handle that listing.
- Resolve PyMuPDF redistribution basis before releasing SplitSelect v1.0.1. Separately review other distributed PyMuPDF artifacts using the same standard.
