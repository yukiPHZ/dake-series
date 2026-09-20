"""Local adaptive PDF compression. No UI, network, or optional imaging dependencies."""
from __future__ import annotations

import hashlib
import multiprocessing as mp
import os
from pathlib import Path
import subprocess
import tempfile
import time

VERIFIED_PYMUPDF = "1.28.2"
ANALYSIS_SAMPLE_LIMIT = 12
VISUAL_DPI = 96
# Calibrated with tests/benchmark.py, including deliberately damaged controls.
QUALITY_LIMITS = {
    "photo": (0.025, 0.12, 0.12),
    "document": (0.008, 0.045, 0.035),
    "lossless": (0.00001, 0.00001, 0.00001),
}
CANDIDATE_TIMEOUT = 60.0
TOTAL_CANDIDATE_BUDGET = 120.0


class AdaptiveStageError(RuntimeError):
    """Preserve internal stage context without exposing it in the normal UI."""

    def __init__(self, stage, exc, *, page_index=None, candidate=None):
        super().__init__(f"{type(exc).__name__}: {exc}")
        self.stage = stage
        self.page_index = page_index
        self.candidate = candidate
        self.exception_type = type(exc).__name__
        self.exception_message = str(exc)


def is_thin_stroke(drawing):
    """Return true only for a finite numeric stroke width below 0.6 pt."""
    width = drawing.get("width")
    return (
        isinstance(width, (int, float))
        and not isinstance(width, bool)
        and 0 < width < 0.6
    )


def library():
    import pymupdf as pdf
    pdf.TOOLS.mupdf_display_errors(False)
    pdf.TOOLS.mupdf_display_warnings(False)
    if not callable(getattr(pdf.Document, "rewrite_images", None)):
        raise RuntimeError("rewrite_images unavailable; reinstall verified dependencies")
    return pdf


def sample_pages(count, limit):
    return sorted({round(i * (count - 1) / (min(count, limit) - 1))
                   for i in range(min(count, limit))}) if count > 1 else [0]


def text_digest(page):
    # Ignore extraction whitespace; keep every character and its order per page.
    text = "".join(page.get_text(sort=True).split())
    return hashlib.sha256(text.encode("utf-8")).hexdigest(), len(text)


def geometry(page):
    return tuple(page.mediabox) + tuple(page.cropbox) + (page.rotation,)


def render(page):
    pdf = library()
    scale = min(VISUAL_DPI / 72, 1400 / max(page.rect.width, page.rect.height))
    return page.get_pixmap(matrix=pdf.Matrix(scale, scale), colorspace=pdf.csRGB, alpha=False)


def visual_difference(before, after):
    if (before.width, before.height, before.n) != (after.width, after.height, after.n):
        raise ValueError("render dimensions differ")
    a, b = before.samples, after.samples
    if a == b:
        return {"mae": 0.0, "changed": 0.0, "ink_loss": 0.0, "tile_ink_loss": 0.0}
    total = changed = ink = lost = 0
    tiles_w = (before.width + 63)//64
    tile_ink = [0] * (tiles_w * ((before.height + 63)//64))
    tile_lost = [0] * len(tile_ink)
    # Every rendered pixel, at bounded resolution. Foreground loss catches thin
    # lines / small stamps that a whole-page average can conceal.
    for i in range(0, len(a), 3):
        delta = max(abs(a[i] - b[i]), abs(a[i+1] - b[i+1]), abs(a[i+2] - b[i+2]))
        total += delta
        changed += delta > 32
        if min(a[i:i+3]) < 160:
            ink += 1
            erased = min(b[i:i+3]) > min(a[i:i+3]) + 60
            lost += erased
            pixel = i//3
            tile = (pixel//before.width//64)*tiles_w + (pixel%before.width//64)
            tile_ink[tile] += 1
            tile_lost[tile] += erased
    count = len(a) // 3
    return {"mae": total / (count * 255), "changed": changed / count,
            "ink_loss": lost / max(1, ink),
            "tile_ink_loss": max((l/n for l,n in zip(tile_lost,tile_ink) if n >= 8), default=0)}


def analyze(path):
    pdf = library()
    size = Path(path).stat().st_size
    with pdf.open(path) as doc:
        if not doc.is_pdf or doc.needs_pass or not len(doc):
            raise ValueError("unreadable or encrypted PDF")
        page_info, unique, occurrences, areas, modes = [], set(), 0, [], {"color": 0, "gray": 0, "bitonal": 0}
        sampled = sample_pages(len(doc), ANALYSIS_SAMPLE_LIMIT)
        scan_pages, photo_pages, protected_pages = [], [], []
        for index, page in enumerate(doc):
            digest, chars = text_digest(page)
            page_info.append({"geometry": geometry(page), "text": digest, "chars": chars,
                              "links": len(page.get_links()),
                              "annots": sum(1 for _ in page.annots() or []),
                              "widgets": sum(1 for _ in page.widgets() or [])})
            # Metadata only: no image decoding and no full-page raster analysis.
            images = page.get_images(full=True)
            occurrences += len(images)
            for img in images:
                if img[0] in unique:
                    continue
                unique.add(img[0])
                modes["bitonal" if img[4] == 1 else "gray" if img[5] == "DeviceGray" else "color"] += 1
            if index not in sampled:
                continue
            placements = page.get_image_info()
            area = min(1.0, sum((pdf.Rect(im["bbox"]) & page.rect).get_area()
                                for im in placements) / max(1, page.rect.get_area()))
            areas.append(area)
            try:
                drawings = page.get_drawings()
            except Exception as exc:
                raise AdaptiveStageError("analysis.drawings", exc, page_index=index) from exc
            if any(is_thin_stroke(drawing) for drawing in drawings):
                protected_pages.append(index)
            if area >= 0.45:
                thumb = page.get_pixmap(matrix=pdf.Matrix(0.35, 0.35), colorspace=pdf.csGRAY)
                white = sum(v > 225 for v in thumb.samples) / len(thumb.samples)
                if white >= 0.55 or modes["bitonal"]:
                    scan_pages.append(index)
                else:
                    photo_pages.append(index)
        coverage = sum(areas) / max(1, len(areas))
        kind = "text" if not unique or coverage < 0.15 else "scan" if scan_pages else "photo" if photo_pages else "mixed"
        if protected_pages:
            kind = "drawing"
        visual_pages = sorted(set([0, len(doc)//2, len(doc)-1] + (scan_pages + protected_pages + photo_pages)[:2]))
        return {"pages": len(doc), "bytes": size, "bytes_per_page": size / len(doc),
                "images": len(unique), "image_occurrences": occurrences,
                "image_page_ratio": sum(bool(doc[i].get_images()) for i in sampled) / len(sampled),
                "sampled_image_coverage": coverage, "image_modes": modes,
                "text_chars": sum(p["chars"] for p in page_info), "kind": kind,
                "analysis_pages": sampled, "visual_pages": visual_pages,
                "protected_pages": protected_pages, "page_info": page_info,
                "interactive": any(p["links"] or p["annots"] or p["widgets"] for p in page_info),
                "embedded_files": doc.embfile_count(), "toc": doc.get_toc()}


def candidate_plan(profile, ghostscript=None):
    names = ["A_lossless"]
    if profile["image_modes"]["bitonal"]:
        names.append("C_bitonal_fax")
    if profile["images"] and profile["sampled_image_coverage"] >= 0.15:
        if profile["kind"] == "photo" and profile["image_modes"]["color"] and not profile["image_modes"]["bitonal"]:
            names.append("B_balanced")
        elif profile["image_modes"]["color"]:
            names.append("C_document")
    # pdfwrite can flatten or remove interactive features. Keep those in MuPDF.
    if ghostscript and not profile["interactive"] and not profile["embedded_files"]:
        names.append("D_ghostscript")
    return names


def create_candidate(source, output, name, ghostscript=None):
    pdf = library()
    if name == "D_ghostscript":
        subprocess.run([str(ghostscript), "-dSAFER", "-dBATCH", "-dNOPAUSE", "-dQUIET",
                        "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.7", "-dPDFSETTINGS=/ebook",
                        "-dAutoRotatePages=/None", "-dDetectDuplicateImages=true",
                        f"-sOutputFile={output}", "-f", str(source)],
                       capture_output=True, check=True, timeout=CANDIDATE_TIMEOUT - 5,
                       creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        return
    with pdf.open(source) as doc:
        if name == "B_balanced":
            doc.rewrite_images(dpi_threshold=220, dpi_target=150, quality=82, bitonal=False, gray=False)
        elif name == "C_document":
            # Preserve scan glyphs / stamps: 300 dpi, high quality, original color.
            # Gray JPEG rewriting is disabled: benchmark E exposed raster
            # corruption in the pinned library. Lossless / bitonal stay enabled.
            doc.rewrite_images(dpi_threshold=320, dpi_target=300, quality=92, bitonal=False, gray=False)
        elif name == "C_bitonal_fax":
            opts = pdf.mupdf.PdfImageRewriterOptions()
            opts.bitonal_image_recompress_method = pdf.mupdf.FZ_RECOMPRESS_FAX
            # No downsampling and no symbol substitution (JBIG2).
            doc.rewrite_images(options=opts)
        elif name != "A_lossless":
            raise ValueError("unknown candidate")
        doc.save(output, garbage=4, clean=True, deflate=True, deflate_images=True,
                 deflate_fonts=True, use_objstms=1)


def verify_candidate(source, output, profile, name):
    pdf = library()
    if not Path(output).is_file() or Path(output).stat().st_size <= 0:
        raise ValueError("empty output")
    with pdf.open(source) as original, pdf.open(output) as result:
        if not result.is_pdf or result.needs_pass or len(result) != profile["pages"]:
            raise ValueError("page count or PDF format changed")
        if result.embfile_count() != profile["embedded_files"] or result.get_toc() != profile["toc"]:
            raise ValueError("document navigation changed")
        for index, page in enumerate(result):
            expected = profile["page_info"][index]
            if any(abs(a-b) > 0.01 for a,b in zip(geometry(page), expected["geometry"])):
                raise ValueError(f"geometry changed on page {index+1}")
            if text_digest(page)[0] != expected["text"]:
                raise ValueError(f"text changed on page {index+1}")
            if (len(page.get_links()), sum(1 for _ in page.annots() or []), sum(1 for _ in page.widgets() or [])) != (expected["links"], expected["annots"], expected["widgets"]):
                raise ValueError(f"interactive content changed on page {index+1}")
        differences = [dict(page=i+1, **visual_difference(render(original[i]), render(result[i])))
                       for i in profile["visual_pages"]]
        worst = {key: max(d[key] for d in differences) for key in ("mae", "changed", "ink_loss")}
        level = "lossless" if name in ("A_lossless", "C_bitonal_fax") else "photo" if profile["kind"] == "photo" else "document"
        limits = QUALITY_LIMITS[level]
        if any(worst[k] > limit for k,limit in zip(worst, limits)):
            raise ValueError(f"visual quality rejected: {worst}")
        tile_loss = max(d["tile_ink_loss"] for d in differences)
        if tile_loss > (0.45 if level == "photo" else 0.2):
            raise ValueError(f"local ink loss rejected: {tile_loss}")
        return {"checks": "passed", "visual": differences, "worst": worst, "quality_level": level}


def _candidate_worker(connection, source, output, profile, name, gs):
    try:
        start = time.monotonic()
        create_candidate(source, output, name, gs)
        connection.send(("phase", "status_verifying"))
        check = verify_candidate(source, output, profile, name)
        connection.send(("result", {"name": name, "size": Path(output).stat().st_size,
                                   "seconds": time.monotonic()-start, **check}))
    except Exception as exc:
        connection.send(("error", f"{type(exc).__name__}: {exc}"))
    finally:
        connection.close()


def run_candidate(source, output, profile, name, gs, phase, timeout):
    # Bound native calls as well as Ghostscript. A bad PDF cannot hang the GUI.
    ctx = mp.get_context("spawn")
    receive, send = ctx.Pipe(duplex=False)
    process = ctx.Process(target=_candidate_worker, args=(send, source, output, profile, name, gs))
    start = time.monotonic()
    process.start()
    send.close()
    try:
        while time.monotonic() - start < timeout:
            if receive.poll(0.05):
                event, data = receive.recv()
                if event == "phase":
                    phase(data)
                elif event == "result":
                    data["seconds"] = time.monotonic() - start
                    return data
                else:
                    raise ValueError(data)
            if not process.is_alive():
                raise ValueError(f"candidate process exited: {process.exitcode}")
        raise TimeoutError("candidate time budget exceeded")
    finally:
        if process.is_alive():
            process.join(0.15)
        if process.is_alive():
            process.terminate()
        process.join(2)
        receive.close()
        process.close()


def choose_candidate(candidates, original_size):
    acceptable = [c for c in candidates if c.get("checks") == "passed" and 0 < c["size"] < original_size]
    for c in acceptable:
        reduction = 1 - c["size"] / original_size
        quality_cost = c["worst"]["mae"] * 4 + c["worst"]["ink_loss"] * 0.5
        time_cost = min(c["seconds"] / TOTAL_CANDIDATE_BUDGET, 1) * 0.03
        # A lossless result wins close calls; minimum size alone is not the goal.
        c["score"] = reduction - quality_cost - time_cost + (0.015 if c["name"] in ("A_lossless", "C_bitonal_fax") else 0)
    return max(acceptable, key=lambda c: c["score"]) if acceptable else None


def compress(source, ghostscript=None, phase=lambda key: None):
    source = Path(source).absolute()
    start = time.monotonic()
    fingerprint = (source.stat().st_size, source.stat().st_mtime_ns)
    phase("status_analyzing")
    profile = analyze(source)
    candidates = []
    with tempfile.TemporaryDirectory(prefix=".dake_pdf_compress_", dir=source.parent) as folder:
        paths = {}
        for name in candidate_plan(profile, ghostscript):
            base = next((c for c in candidates if c["name"] == "A_lossless" and c.get("checks") == "passed"), None)
            if name in ("B_balanced", "C_document", "D_ghostscript") and base and base["size"] < profile["bytes"] * 0.1 and base["size"] / profile["pages"] < 65536:
                candidates.append({"name": name, "skipped": "lossless already compact (<64 KiB/page, >90% reduction)"})
                continue
            remaining = TOTAL_CANDIDATE_BUDGET - (time.monotonic()-start)
            if remaining < 1:
                candidates.append({"name": name, "rejected": "total time budget"})
                continue
            phase("status_optimizing" if name == "A_lossless" else "status_images")
            path = Path(folder) / (name + ".pdf")
            paths[name] = path
            try:
                candidates.append(run_candidate(source, path, profile, name, ghostscript, phase,
                                                min(CANDIDATE_TIMEOUT, remaining)))
            except Exception as exc:
                candidates.append({
                    "name": name,
                    "rejected": str(exc),
                    "stage": getattr(exc, "stage", "candidate"),
                    "page_index": getattr(exc, "page_index", None),
                    "candidate": getattr(exc, "candidate", name),
                    "exception_type": getattr(exc, "exception_type", type(exc).__name__),
                    "exception_message": getattr(exc, "exception_message", str(exc)),
                })
        chosen = choose_candidate(candidates, profile["bytes"])
        if chosen is None:
            return None, profile, candidates
        if fingerprint != (source.stat().st_size, source.stat().st_mtime_ns):
            raise ValueError("source changed during compression")
        phase("status_saving")
        # Atomic no-clobber publication: link complete temp inode into a new
        # name. Unlike exists()+replace(), this also protects concurrent runs.
        index = 1
        while True:
            suffix = "" if index == 1 else f"_{index}"
            target = source.with_name(f"{source.stem}_compressed{suffix}.pdf")
            try:
                os.link(paths[chosen["name"]], target)
                break
            except FileExistsError:
                index += 1
            except OSError:
                # Windows rename is atomic and refuses to replace an existing
                # target, including on filesystems without hardlinks.
                if os.name != "nt":
                    raise
                try:
                    os.rename(paths[chosen["name"]], target)
                    break
                except FileExistsError:
                    index += 1
        return {"path": str(target), **chosen, "total_seconds": time.monotonic()-start}, profile, candidates
