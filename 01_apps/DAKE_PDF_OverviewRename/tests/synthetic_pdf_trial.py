# -*- coding: utf-8 -*-
"""Generate non-confidential PDFs and exercise scan/render/rename/undo at scale."""

from __future__ import annotations

import argparse
import hashlib
import json
import queue
import sys
import tempfile
import threading
import time
from pathlib import Path

from PIL import Image, ImageDraw

APP_DIR = Path(__file__).resolve().parents[1]
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

from main import RenderPool, RenderRequest, scan_pdf_folder
from rename_core import FileSnapshot, RenameRequest, rename_batch, undo_rename


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def make_synthetic_pdf(path: Path, index: int, total: int) -> None:
    image = Image.new("RGB", (320, 450), "white")
    draw = ImageDraw.Draw(image)
    draw.rectangle((18, 18, 302, 432), outline="#2F6FED", width=4)
    draw.text((38, 50), "DAKE synthetic PDF", fill="#1E2430")
    draw.text((38, 82), f"document {index:04d} / {total:04d}", fill="#1E2430")
    draw.text((38, 118), "No confidential data", fill="#667085")
    image.save(path, "PDF", resolution=96.0)


def run_trial(root: Path, count: int) -> dict[str, object]:
    folder = root / f"pdf_{count:03d}"
    folder.mkdir(parents=True)
    started = time.perf_counter()
    for index in range(1, count + 1):
        make_synthetic_pdf(folder / f"scan_{index:04d}.pdf", index, count)
    generated_seconds = time.perf_counter() - started

    started = time.perf_counter()
    snapshots = scan_pdf_folder(folder)
    scan_seconds = time.perf_counter() - started
    assert len(snapshots) == count
    before_hashes = {snapshot.path.name: digest(snapshot.path) for snapshot in snapshots}

    pool = RenderPool(worker_count=3)
    generation = count
    requests = [
        RenderRequest(generation, "thumbnail", index, snapshot, (270, 350))
        for index, snapshot in enumerate(snapshots)
    ]
    started = time.perf_counter()
    pool.replace(generation, requests)
    rendered = 0
    failures: list[str] = []
    deadline = time.monotonic() + max(30.0, count * 1.5)
    while rendered < count and time.monotonic() < deadline:
        try:
            result = pool.results.get(timeout=0.5)
        except queue.Empty:
            continue
        rendered += 1
        if result.error is not None or result.image is None or result.page_count != 1:
            failures.append(result.request.snapshot.path.name)
    render_seconds = time.perf_counter() - started
    pool.shutdown()
    worker_threads_stopped = not any(
        thread.is_alive() and thread.name.startswith("overview-thumb-")
        for thread in threading.enumerate()
    )
    assert rendered == count
    assert not failures
    assert worker_threads_stopped

    requests_for_rename = [
        RenameRequest(snapshot, f"document_{index:04d}")
        for index, snapshot in enumerate(snapshots, start=1)
    ]
    started = time.perf_counter()
    _plan, undo = rename_batch(requests_for_rename)
    rename_seconds = time.perf_counter() - started
    renamed_paths = sorted(folder.glob("*.pdf"))
    assert len(renamed_paths) == count
    after_hashes = {
        f"scan_{index:04d}.pdf": digest(folder / f"document_{index:04d}.pdf")
        for index in range(1, count + 1)
    }
    assert after_hashes == before_hashes

    started = time.perf_counter()
    undo_rename(undo)
    undo_seconds = time.perf_counter() - started
    restored = sorted(folder.glob("*.pdf"))
    assert len(restored) == count
    assert {path.name: digest(path) for path in restored} == before_hashes
    assert not any(path.name.startswith(".__dake_overview_") for path in folder.iterdir())

    return {
        "count": count,
        "generated_seconds": round(generated_seconds, 3),
        "scan_seconds": round(scan_seconds, 3),
        "render_seconds": round(render_seconds, 3),
        "rendered": rendered,
        "render_failures": len(failures),
        "rename_seconds": round(rename_seconds, 3),
        "undo_seconds": round(undo_seconds, 3),
        "content_hashes_preserved": True,
        "temporary_files_remaining": 0,
        "worker_threads_stopped": worker_threads_stopped,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("counts", nargs="*", type=int, default=[1, 48, 100, 300])
    args = parser.parse_args()
    with tempfile.TemporaryDirectory(prefix="dake_overview_trial_") as temporary:
        root = Path(temporary)
        results = [run_trial(root, count) for count in args.counts]
    print(json.dumps(results, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
