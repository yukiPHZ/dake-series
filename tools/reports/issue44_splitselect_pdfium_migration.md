# Issue #44 SplitSelect PDFium Migration Audit

Date: 2026-09-24 (JST). Branch: `codex/splitselect-pdfium-license-unblock`.
This is an unpublished Draft PR candidate, not a formal release decision.

## Scope and source of truth

`01_apps/DAKE_PDF_SplitSelect/ORIGINAL.md` remains the source of truth. Only this app's
renderer, pinned requirements, build, tests, notices and technical documentation changed.
No Release, BOOTH, Store, dakeapp.com or Cloudflare publication was performed.

| Boundary | Before | After |
|---|---|---|
| PDF open / thumbnail | `fitz.open`, `Page.get_pixmap` | worker-owned `pypdfium2.PdfDocument`, `Page.render` |
| Save / CLI | `pypdf` | unchanged `pypdf` |
| UI image contract | PPM bytes to Tk `PhotoImage` | unchanged PPM P6 bytes |
| Product dependencies | PyMuPDF 1.24.10 + PyMuPDFb | pypdfium2 5.13.0; no Pillow |
| Build collection | `--collect-all=fitz` | normal pypdfium2 hook; DnD collect retained |

`Page.render()` requests PDFium three-channel BGR with `rev_byteorder=True`. The PPM
payload copies each row's first `width * 3` RGB bytes, excluding stride padding.
Bitmap and page close in `finally`; document closes on refresh/replacement/stop/error.
The single worker owns product PDFium operations; no extra mutex was introduced.

## Verification

- Python 3.12.4 clean venv; `pip check` passed. Neither PyMuPDF nor Pillow is installed.
- Full test discovery: **10 tests passed in 12.558 s**. Includes 3/30/100/300 pages,
  visible-first requests, +/-8 prefetch, cache max 96/re-render, click/Shift/range,
  merged and single save, F5, Ctrl+O, cancellation/generation, replacement/handle
  release, invalid/protected PDFs, shutdown and 9 CLI cases. Fixtures are generated
  with `pypdf`, not PyMuPDF. Portrait, landscape, rotated, vector and image-like
  RGB fixtures have nonempty/nonuniform PPM payloads.
- Formal `build.bat` onedir succeeded. `dist` has 1,091 files (39,711,118 bytes), one
  `pdfium.dll`, notices and all 19 license files. Product source `fitz` imports:
  **0**; `Analysis-00.toc` PyMuPDF/fitz/MuPDF entries: **0**; onedir forbidden
  filenames: **0**. No Pillow/PIL payload found.
- Ten warmed onedir process launches to first window handle: 0.283, 0.249, 0.253,
  0.240, 0.255, 0.247, 0.244, 0.254, 0.257, 0.314 s. Median **0.253 s**,
  p90 **0.283 s**. Ten normal `WM_CLOSE` exits. Cold-boot timing is not asserted.
- Unpublished ZIP: 17,775,190 bytes, SHA-256
  `a271e9a186c30a5bf0e7db8ef37342f36e4c3f9172b3850588f162317341bea4`.
  All 1,091 extracted files matched the onedir source by SHA-256. Extracted EXE
  passed 9 CLI cases. Its GUI loaded a synthetic three-page PDF via Ctrl+O, rendered
  colored thumbnails in page order, selected page 1, saved a one-page PDF containing
  `PAGE 001`, cleared with F5 and exited normally. Source PDF hash was unchanged.
- EXE SHA-256: `c050664341455f05f70b4f9c8acb14ca266abcedb0ec54f883b725a10897ba1c`.
  ZIP and EXE are ignored local candidates; neither was committed or published.

Physical Explorer drag-and-drop, DPI variants and real-world scanned PDFs were not
directly verified. They remain review/field-check notes rather than fabricated PASSes.

## Installed-wheel third-party license audit

Installed `pypdfium2-5.13.0.dist-info/METADATA` lists 19 `License-File` entries.
`data/windows_x64/BUILD_LICENSES` was flattened to `BUILD_LICENSES` in the shipped
layout. `tests/test_license_payload.py` checks the exact path set and SHA-256 of each
file in both source and onedir against the actual installed wheel. The complete
filenames are in `THIRD_PARTY_NOTICES.txt`. Installed-wheel SHA-256 manifest:

| file | SHA-256 |
|---|---|
| BUILD_LICENSES/abseil.txt | f54fff0b905df5b3464527c652a30e903b172d6dcab4d89b5e6f105d5e4a4603 |
| BUILD_LICENSES/agg23.txt | c110d3ea2ad77467ce0dcff7d3337e6c8be8049a5103f4b9bd5fd911a77972e5 |
| BUILD_LICENSES/fast_float.txt | bf1b57355feca8fce77ee95f48002f8d4789fb71b30ec7599c06cda4901fbb2b |
| BUILD_LICENSES/freetype.txt | f4b133e25df1f86ad3ffea453aa0e613f0474f34778dbbb3e437e7b2724937d8 |
| BUILD_LICENSES/icu.txt | 93679f4389d53b6835d89843f251844fb9bc455b35bab036d3c8e7abe497a47a |
| BUILD_LICENSES/lcms.txt | 7312b68c5b25e9bf2b828706fb4e29588f22705112f411fd42e1f7d84c3d139a |
| BUILD_LICENSES/libjpeg_turbo.ijg | db16a04128171879c60708d171b88d97345a2dd20f9bfc173680a4497c73f704 |
| BUILD_LICENSES/libjpeg_turbo.md | be2b2b5ab168bce87bc3e31f2a5c5adba4b7f6e9e51d618e958d1d46972ebd95 |
| BUILD_LICENSES/libopenjpeg.txt | c5ab0890a737c2dfa7ba675036554f6d17741d98629b0c2a145354d00617e6b2 |
| BUILD_LICENSES/libpng.txt | 452390433ba0f88aa3e2b122c647741b72a0c117cd6ed7a329b49785aecb5511 |
| BUILD_LICENSES/libtiff.txt | 92b72ba97e6c2749c2a94bc0ef646b47080217f1e772a482b33cf5a5f98a6506 |
| BUILD_LICENSES/llvm-libc.txt | 3b6226c32e168c83b891d8d6f0d3c29c2116dc3ef93dc93c307b54f279ecf383 |
| BUILD_LICENSES/pdfium-binaries.txt | 8854f4388f1ca13b3ad9baa42e95f5546b4c0b17109c159256d3eca7be39b09b |
| BUILD_LICENSES/pdfium.txt | 961eacd9633fff6d051db7208b755e9210e30efac7adec3e6a6d52798f0ccf0e |
| BUILD_LICENSES/simdutf.txt | c172a0ba936ff31230febb5dad869e25cb7c1a07480c7a381be8cf011bb52719 |
| BUILD_LICENSES/zlib.txt | 33fd641c9f3b0e0be64bc78fea9e94807674cdd70c48477599226cb8956565fe |
| LICENSES/Apache-2.0.txt | 3ddf9be5c28fe27dad143a5dc76eea25222ad1dd68934a047064e56ed2fa40c5 |
| LICENSES/BSD-3-Clause.txt | ad9a9e823df025f42389c1812eae28019f657d1ed7b3a4ebfd5010b0736a0da4 |
| LICENSES/CC-BY-4.0.txt | f1b1748301cd4274f46733c218d510929661b4990855de5fc645dacd69c4371c |

Eight files differ by SHA from OverviewRename's prior same-version license copy;
that copy was **not** used as the source. The installed wheel reports pypdfium2
5.13.0 and PDFium 153.0.7999.0. The notices point users to the unmodified texts.

## Merge cross-check and remaining gate

`DAKE_PDF_Merge/requirements.txt` still names PyMuPDF. Its `main.py` lazily imports
`fitz` for previews; the published Merge v1.0.1 EXE audit found the native payload.
SplitSelect's PDFium/PPM pattern can inform a separate Merge migration, but UI
integration and release require a separate Issue/PR. No Merge source or listing
changed here.

The old SplitSelect PyMuPDF redistribution blocker is technically addressed in this
candidate, subject to ChatGPT review and remaining formal shipment checks. Issue
#31's HOLD and all external publication remain unchanged until review.
