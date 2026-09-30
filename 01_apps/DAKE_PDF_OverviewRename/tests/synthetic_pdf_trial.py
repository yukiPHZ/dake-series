# -*- coding: utf-8 -*-
"""Generate non-confidential PDFs and exercise scan/render/rename/undo at scale."""

from __future__ import annotations

import argparse
import ctypes
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

from main import THUMB_RENDER_BOX, RenderPool, RenderRequest, scan_pdf_folder
from rename_core import FileSnapshot, RenameRequest, rename_batch, undo_rename


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def working_set_bytes() -> int | None:
    if not sys.platform.startswith("win"):
        return None

    class ProcessMemoryCounters(ctypes.Structure):
        _fields_ = [
            ("cb", ctypes.c_ulong),
            ("PageFaultCount", ctypes.c_ulong),
            ("PeakWorkingSetSize", ctypes.c_size_t),
            ("WorkingSetSize", ctypes.c_size_t),
            ("QuotaPeakPagedPoolUsage", ctypes.c_size_t),
            ("QuotaPagedPoolUsage", ctypes.c_size_t),
            ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
            ("QuotaNonPagedPoolUsage", ctypes.c_size_t),
            ("PagefileUsage", ctypes.c_size_t),
            ("PeakPagefileUsage", ctypes.c_size_t),
        ]

    counters = ProcessMemoryCounters()
    counters.cb = ctypes.sizeof(counters)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    psapi = ctypes.WinDLL("psapi", use_last_error=True)
    kernel32.GetCurrentProcess.restype = ctypes.c_void_p
    psapi.GetProcessMemoryInfo.argtypes = [
        ctypes.c_void_p,
        ctypes.POINTER(ProcessMemoryCounters),
        ctypes.c_ulong,
    ]
    psapi.GetProcessMemoryInfo.restype = ctypes.c_int
    process = kernel32.GetCurrentProcess()
    ok = psapi.GetProcessMemoryInfo(process, ctypes.byref(counters), counters.cb)
    return int(counters.WorkingSetSize) if ok else None


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
    expected_names = {f"scan_{index:04d}.pdf" for index in range(1, count + 1)}
    scanned_names = {snapshot.path.name for snapshot in snapshots}
    assert scanned_names == expected_names
    before_hashes = {snapshot.path.name: digest(snapshot.path) for snapshot in snapshots}

    pool = RenderPool(worker_count=3)
    memory_before_render = working_set_bytes()
    peak_working_set = memory_before_render
    generation = count
    requests = [
        RenderRequest(generation, "thumbnail", index, snapshot, THUMB_RENDER_BOX)
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
        current_memory = working_set_bytes()
        if current_memory is not None:
            peak_working_set = max(peak_working_set or 0, current_memory)
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

    rename_seconds = 0.0
    undo_seconds = 0.0
    if snapshots:
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
        "path_sets_match": scanned_names == expected_names,
        "rename_seconds": round(rename_seconds, 3),
        "undo_seconds": round(undo_seconds, 3),
        "content_hashes_preserved": True,
        "temporary_files_remaining": 0,
        "worker_threads_stopped": worker_threads_stopped,
        "memory_before_render_mb": (
            round(memory_before_render / 1024**2, 1) if memory_before_render is not None else None
        ),
        "peak_working_set_mb": (
            round(peak_working_set / 1024**2, 1) if peak_working_set is not None else None
        ),
        "render_working_set_delta_mb": (
            round((peak_working_set - memory_before_render) / 1024**2, 1)
            if peak_working_set is not None and memory_before_render is not None
            else None
        ),
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
