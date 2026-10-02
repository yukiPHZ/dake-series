# PDFium onefile distribution review

Date: 2026-10-03 (JST). Local, unpublished ChatGPT-review candidate.
Base: `origin/main` at `32848f5881defc507cd2b910d085454434ae0060`.
Branch: `codex/splitselect-pdfium-onefile-distribution-20261003`.

## Decision

Restore `--onefile --noconsole` as SplitSelect's formal build configuration.
The final EXE starts and renders PDFs after an EXE-only copy to both a fresh empty
folder and a fresh Desktop folder. Neither contains `_internal`, adjacent DLLs,
license folders, or any other files. User-prioritized EXE portability is achieved.

This is **not** a startup-speed improvement over PDFium onedir: measured onefile
median is 1.349418 s, versus historical onedir 0.253 s. Its median remains below
the earlier 1.5 s goal; this tradeoff is acceptable for the requested portable UX.
ChatGPT review and the remaining Human Review checks still precede shipment.

No renderer, selection, refresh, save, CLI, or UI-layout changes were made.
`main.py`, `pdf_backend.py`, and dependency versions are unchanged.

## Clean build

New Python 3.12.4 venv; system site packages disabled. Installed only the existing
`requirements.txt`, with no package upgrade:

| Dependency | Version |
| --- | --- |
| pypdfium2 | 5.13.0 |
| PDFium wheel binary | 153.0.7999.0 |
| pypdf | 6.10.2 |
| PyInstaller | 6.19.0 |
| pyinstaller-hooks-contrib | 2026.4 |
| tkinterdnd2 | 0.4.3 |

`pip check`: PASS. `find_spec('fitz')` / `find_spec('PIL')`: both `None`.
The original onedir `build.bat` was built first in this environment and moved to
an external evidence directory. The modified formal `build.bat` then produced
the final clean onefile EXE. No old onedir runtime remains beside that EXE.

- Output: `dist/DakePDF_Split_Select.exe`.
- Size: 17,673,614 bytes.
- SHA-256: `181916b7f78ca9f3cd0711ab62db459ff3aaca65aa52dc139f6aaf729c97ab63`.
- Flags: `--clean --onefile --windowed --noconsole`.
- PDFium: standard installed hook; exactly one `pdfium.dll` in the EXE archive.
- DnD: retain `--collect-all=tkinterdnd2`; embedded native tkdnd library verified.
- Common icon: unchanged `../../02_assets/dake_icon.ico`, embedded bytes match.
- Notice and all 19 wheel license files: embedded and also copied beside the build
  output for readable ZIP distribution. Runtime does not read adjacent notices.

## Startup benchmark

Windows, `Popen` immediately before launching -> first visible `TkTopLevel` owned
by a process whose full executable path matches the test EXE. This includes the
onefile child process and extraction delay. Read-only Win32 window/process queries,
5 ms polling, no splash. All 20 launches closed normally with `WM_CLOSE`, exit 0.
Ten launches per build; onefile measured from the EXE-only empty-folder copy.
Same machine/session, onefile group then onedir group, 0.5 s between trials.
This is a repeated/warmed launch test, **not** a reboot/cold-cache benchmark.
p90 uses nearest rank (9th of 10 sorted results); no outliers were discarded.

| Trial | Final onefile (s) | Clean baseline onedir (s) |
| --- | ---: | ---: |
| 1 | 1.439526 | 0.859372 |
| 2 | 1.369983 | 0.258575 |
| 3 | 1.413612 | 0.271644 |
| 4 | 1.353324 | 0.246955 |
| 5 | 1.246732 | 0.264398 |
| 6 | 1.379509 | 0.283924 |
| 7 | 1.220825 | 0.268144 |
| 8 | 1.345511 | 0.264308 |
| 9 | 1.327551 | 0.253689 |
| 10 | 1.287148 | 0.245268 |
| Median | **1.349418** | **0.264353** |
| p90 | **1.413612** | **0.283924** |
| Fastest | 1.220825 | 0.245268 |
| Slowest | 1.439526 | 0.859372 |

Historical onedir from the existing PDFium report: median **0.253 s**, p90
**0.283 s**; these are previous measurements, not this run. Relative to those
historical values, onefile median adds **1.096418 s (+433.4%)** and p90 adds
**1.130612 s (+399.5%)**. The cost is accepted for EXE-only portability, not
presented as a speedup. Against today's clean onedir, the median difference is
1.085065 s. First-launch behavior on other machines/security software is unknown.

## Final packaged EXE-only regression

The copied EXE has the identical final hash above. The following UI tests were
performed on that copy, not a prototype or source-only substitute.

| Check | Result / evidence |
| --- | --- |
| Empty-folder EXE only | PASS: GUI launch, all page sizes, saves, F5, Ctrl+O; 10 launches |
| Desktop-folder EXE only | PASS: GUI launch, button PDF addition, 3-page render, F5, normal exit |
| Adjacent files | PASS: both copy folders still contain only `DakePDF_Split_Select.exe` |
| Python/DnD runtime | PASS: actual Desktop child process loaded `python312.dll` and `libtkdnd2.9.4.dll` from `%TEMP%/_MEI...`, not beside EXE |
| PDFium | PASS: real colored/text page thumbnails rendered, archive has one DLL |
| 3 pages | PASS: all three thumbnails and ready status; click -> 1 selected -> merged save |
| 30 pages | PASS: initial thumbnails, wheel scroll, scrollbar drag to pages 25-30 |
| 100 pages | PASS: initial thumbnails, scrollbar drag to pages 93-100 |
| 300 pages | PASS: initial thumbnails, scrollbar drag to pages 293-300 |
| Click selection | PASS: selected card, count, action buttons, status |
| Range `1-3,5,8-10` | PASS: count 7, merged and single GUI saves |
| Merged | PASS: 7 pages in order 1,2,3,5,8,9,10; extracted PAGE labels match |
| Single | PASS: 7 files; each 1 page; same PAGE labels |
| Completion dialog | PASS: merged and single count/message; OK closes dialog |
| Explorer | PASS: OK opens fixture output folder, files visible |
| F5 | PASS: filename/count/range/selection/thumbnails/completion/status reset |
| 300 -> F5 -> 3 pages | PASS: new 3-page thumbnails, no old images/status observed |
| Ctrl+O | PASS: PDF picker opens; subsequent files load |
| Common icon | PASS: title-bar icon visible, embedded icon hash matches common asset |
| Normal exit | PASS: GUI Alt+F4; benchmark normal close exit 0 |
| Physical Shift+click | NOT CHECKED: input API cannot hold a modifier during mouse click; source regression PASS |
| Physical Explorer -> app DnD | NOT CHECKED: cross-window drag unavailable; DLL payload/load and source event-handler tests PASS, not physical-DnD PASS |

The packaged 300-page render was too fast to reliably capture an in-flight task
before F5. Normal packaged refresh PASS; exact in-flight cancellation is covered
by the source controlled test (test-only delay, no product sleep), and remains
a packaged Human Review timing check.

### Handle release and non-destruction

- PASS: after the final EXE's 300-page F5, that fixture was renamed, moved into a
  different directory, hashed, and moved back successfully while the app remained open.
- PASS: Desktop EXE loaded a dedicated 3-page test copy; after F5, Windows deletion
  succeeded while the app remained open. Original fixtures were not deleted.
- PASS: all four input fixtures' SHA-256 values matched their before-test values
  after GUI output. Saving left originals unchanged.

### CLI

`test_packaged_cli.py`: all nine cases PASS for the final build and separately for
the empty-folder and Desktop EXE-only copies. `--from-shimarisu`, `--inputs`,
`--pages`, `--output`, `--silent`: success exit 0; invalid/missing input exit 1;
output file/directory/default naming and non-overwrite behavior checked;
source SHA-256 unchanged. No GUI opens for CLI. Console attachment/visible stderr
display of the noconsole EXE was not separately measured.

## Source lifecycle regression (supporting, not packaged timing)

All 11 unittest tests PASS in the final clean build environment (27.470 s):
7 lifecycle, 3 license/archive, and the nine-case packaged CLI test.
Synthetic fixtures contain numbered text and color; orientation/image cases also
PASS. Source DnD event handler, Shift range selection/clear, protected/invalid PDF,
save guard, cancellation, document close, and worker shutdown PASS.

| Pages | Source ready (s) | Initial request size | Final cache size | Total renders after traversal/return |
| --- | ---: | ---: | ---: | ---: |
| 3 | 0.0873 | 3 | 3 | 3 |
| 30 | 0.0847 | 16 | 30 | 30 |
| 100 | 0.1392 | 16 | 96 | 104 |
| 300 | 0.0667 | 16 | 96 | 316 |

These are source test timings with warmed backend/fixtures, not packaged PDF-open
benchmarks. Visible-first +/-8 and bounded 96-page LRU remain unchanged. More than
96 renders during traversal is expected; eviction permits later regeneration.
Controlled old-task cancellation -> new 3-page readiness: 0.3927 s; observed
pending maximum 15, old PDF closed. No product delay was added.

## ZIP / distribution and licenses

Local review ZIP: `DakePDF_Split_Select_PDFium_onefile_REVIEW.zip`, 17,085,566 bytes.
SHA-256: `f353b3f7d367a94e046f853e0ed293f9d6b2f348bfb2d6f76985a1045aaec6af`.
23 members: EXE, README.txt, 注意事項.txt, THIRD_PARTY_NOTICES.txt, 19 license files.
No `_internal` directory or loose DLLs. ZIP member list, CRC and every member's
SHA-256 verified through the existing shared ZIP writer/verifier.

`tools/shipping_artifacts.py` previously included only the EXE for onefile. It now
also includes explicitly named sibling notices/license trees, without collecting
unrelated dist files, and rejects symlinks. Its five tests PASS, including existing
onedir handling and missing-license rejection. `tools/make_booth_ready.py` uses
this writer; `tools/check_booth_ready.py` uses its runtime verifier.
Launcher `AppMeta.standard_exe_path` already prefers `dist/<exe_name>` and falls
back to onedir, so this onefile path requires no Launcher change (code inspection,
not a new Launcher executable test).

Existing `booth_ready` assets and external Release/BOOTH/Store/site were **not**
regenerated or published. They still describe the previous onedir payload and must
be regenerated after approval; this local ZIP is a review candidate, not replacement
shipping inventory. README, ORIGINAL and local release-body source now agree with
the intended onefile configuration. EXE name/version/public URLs unchanged.

The installed wheel's 19 `License-File` entries match repository, readable dist and
embedded EXE documents by SHA-256. pypdfium2 texts include Apache-2.0 / BSD-3-Clause /
CC-BY-4.0, plus bundled PDFium/build dependency notices. PyMuPDF/fitz/Pillow are absent
from product requirements, source renderer imports and packaged Analysis/archive.
No legal-clearance claim is made; review the full bundled notices before shipping.
Removing adjacent documents for runtime testing does not authorize omission of
license documents when redistributing the product.

## Static / Git / remaining checks

- PASS: app `compileall`, `pip check`, 11 app tests, 5 shipping tests.
- PASS: `git diff --check`; no Japanese literals directly in widget `text=` found.
- Product source untouched; no new worker-to-Tk access or lifecycle change introduced.
- Isolated worktree; unrelated primary-checkout STADIO edits not copied or staged.
- build/dist/spec/venv/fixtures/EXE/ZIP/benchmark artifacts excluded from commit.
- No push, public PR, main merge, Release/BOOTH/Store/Cloudflare/site update.
- Human Review remaining: physical Explorer DnD, physical packaged Shift+click,
  precise packaged refresh-during-render timing, DPI 125/150/200%, real-world scanned
  and unusual PDFs, other-machine first-launch/security behavior, ChatGPT review.
- Existing OS display settings unchanged; no new DPI certification claimed.
- Final local Git status is recorded in the task completion, not inferred here.
