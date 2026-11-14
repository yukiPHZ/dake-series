"""Serialized PDF rendering with latest-request replacement and bounded PIL cache."""
from __future__ import annotations

from collections import OrderedDict
from pathlib import Path
import queue
import threading


CACHE_PAGES = 48
CACHE_BYTES = 24 * 1024 * 1024
RENDER_WIDTH = 360
RENDER_HEIGHT = 512


class ThumbnailWorker:
    def __init__(self, results: queue.Queue):
        self.results = results
        self.condition = threading.Condition()
        self.loads: dict[str, tuple[int, Path]] = {}
        self.request = None
        self.current_source = None
        self.closed = False
        self.cache = OrderedDict()
        self.cache_bytes = 0
        self.cache_peak_bytes = 0
        self.render_count = 0
        self.thread = threading.Thread(target=self._run, name="pdf-thumbnails", daemon=True)
        self.thread.start()

    def load(self, kind: str, load_id: int, path: Path):
        with self.condition:
            self.loads[kind] = (load_id, path)
            if kind == "source":
                self.current_source = (load_id, path)
                self.request = None
            self.condition.notify()

    def render(self, load_id: int, request_id: int, path: Path, pages: tuple[int, ...], width: int, height: int):
        with self.condition:
            self.request = (load_id, request_id, path, pages, width, height)
            self.condition.notify()

    def invalidate(self):
        with self.condition:
            self.request = None

    def close(self):
        with self.condition:
            self.closed = True
            self.request = None
            self.loads.clear()
            self.condition.notify()

    def _current(self, request):
        with self.condition:
            return not self.closed and self.request is request and not self.loads

    def _publish(self, message, request=None):
        while not self.closed and (request is None or self._current(request)):
            try:
                self.results.put(message, timeout=.025)
                return True
            except queue.Full:
                continue
        return False

    def _run(self):
        document = None
        document_key = None
        try:
            while True:
                with self.condition:
                    self.condition.wait_for(lambda: self.closed or self.loads or self.request is not None)
                    if self.closed:
                        return
                    loads = self.loads
                    self.loads = {}
                    request = self.request
                # Heavy imports happen only after the first accepted PDF.
                try:
                    import fitz
                    from PIL import Image
                except Exception as exc:
                    for kind, (load_id, _path) in loads.items():
                        self._publish(("pdf_load_failed", kind, load_id, exc))
                    with self.condition:
                        if self.request is request:
                            self.request = None
                    continue
                for kind, (load_id, path) in loads.items():
                    try:
                        with fitz.open(str(path)) as info:
                            if info.needs_pass:
                                self._publish(("pdf_rejected", kind, load_id, "error_pdf_protected"))
                                continue
                            count = info.page_count
                            if count < 1:
                                self._publish(("pdf_rejected", kind, load_id, "error_pdf_empty"))
                                continue
                        self._publish(("pdf_metadata", kind, load_id, path, count))
                    except Exception as exc:
                        self._publish(("pdf_load_failed", kind, load_id, exc))
                if request is None or not self._current(request):
                    continue
                load_id, request_id, path, pages, width, height = request
                key = (load_id, path)
                try:
                    if key != document_key:
                        if document is not None:
                            document.close()
                        document = fitz.open(str(path))
                        document_key = key
                        self.cache.clear()
                        self.cache_bytes = 0
                    for page_index in pages:
                        if not self._current(request):
                            break
                        image = self.cache.get(page_index)
                        if image is None:
                            page = document.load_page(page_index)
                            rect = page.rect
                            zoom = min(RENDER_WIDTH/max(rect.width, 1), RENDER_HEIGHT/max(rect.height, 1))
                            pixmap = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False, colorspace=fitz.csRGB)
                            image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
                            self.render_count += 1
                            self.cache[page_index] = image
                            self.cache_bytes += image.width * image.height * 3
                            while len(self.cache) > CACHE_PAGES or self.cache_bytes > CACHE_BYTES:
                                _, victim = self.cache.popitem(last=False)
                                self.cache_bytes -= victim.width * victim.height * 3
                            self.cache_peak_bytes = max(self.cache_peak_bytes, self.cache_bytes)
                        else:
                            self.cache.move_to_end(page_index)
                        resized = image.copy()
                        resized.thumbnail((width, height), Image.Resampling.LANCZOS)
                        if self._current(request):
                            # Backpressure bounds queued PIL images if the UI is busy.
                            self._publish(("view_thumbnail", load_id, request_id, page_index, width, height, resized), request)
                    if self._current(request):
                        self._publish(("view_done", load_id, request_id), request)
                except Exception as exc:
                    if self._current(request):
                        self._publish(("view_failed", load_id, request_id, exc), request)
                finally:
                    with self.condition:
                        if self.request is request:
                            self.request = None
        finally:
            if document is not None:
                document.close()
            self.cache.clear()
            self.cache_bytes = 0
