"""One owner for the PDF handle; no Tk objects cross this boundary."""

import queue
import threading
from collections import deque


def render_page_ppm(document, index, width, height):
    import pypdfium2.raw as pdfium_c

    page = document[index]
    bitmap = None
    try:
        page_width, page_height = page.get_size()
        if page_width <= 0 or page_height <= 0:
            raise ValueError("PDF page has invalid dimensions")
        scale = min(width / page_width, height / page_height)
        bitmap = page.render(
            scale=scale,
            force_bitmap_format=pdfium_c.FPDFBitmap_BGR,
            rev_byteorder=True,
        )
        row_bytes = bitmap.width * 3
        if bitmap.stride < row_bytes:
            raise ValueError("PDF bitmap has invalid stride")
        pixels = memoryview(bitmap.buffer).cast("B")
        rgb = b"".join(
            pixels[row * bitmap.stride:row * bitmap.stride + row_bytes]
            for row in range(bitmap.height)
        )
        return f"P6\n{bitmap.width} {bitmap.height}\n255\n".encode("ascii") + rgb
    finally:
        if bitmap is not None:
            bitmap.close()
        page.close()


class PdfRenderWorker:
    def __init__(self, events, width, height):
        self.events = events
        self.width = width
        self.height = height
        self.condition = threading.Condition()
        self.pending = deque()
        self.cancel = threading.Event()
        self.view_cancel = threading.Event()
        self.stopping = False
        self.closed = threading.Event()
        self.closed.set()
        self.rendered_count = 0
        self.render_page = render_page_ppm
        self.thread = threading.Thread(target=self._run, name="pdf-render", daemon=True)
        self.thread.start()

    def reset(self, generation, path=None):
        with self.condition:
            self.cancel.set()
            self.view_cancel.set()
            self.cancel = threading.Event()
            self.view_cancel = threading.Event()
            self.pending.clear()
            self.closed.clear()
            self.pending.append(("close", generation, None, self.cancel, None, None))
            if path is not None:
                self.pending.append(("open", generation, path, self.cancel, None, None))
            self.condition.notify()

    def request(self, generation, request_id, pages):
        with self.condition:
            self.view_cancel.set()
            self.view_cancel = threading.Event()
            self.pending = deque(task for task in self.pending if task[0] != "render")
            for page in pages:
                self.pending.append(("render", generation, page, self.cancel, self.view_cancel, request_id))
            self.condition.notify()

    def stop(self):
        with self.condition:
            self.stopping = True
            self.cancel.set()
            self.view_cancel.set()
            self.pending.clear()
            self.condition.notify()

    def _emit(self, event, cancel, view_cancel=None):
        while not cancel.is_set() and not (view_cancel and view_cancel.is_set()):
            try:
                self.events.put(event, timeout=0.05)
                return
            except queue.Full:
                continue

    def _run(self):
        document = None
        document_generation = None
        try:
            while True:
                with self.condition:
                    self.condition.wait_for(lambda: self.pending or self.stopping)
                    if self.stopping:
                        return
                    kind, generation, value, cancel, view_cancel, request_id = self.pending.popleft()
                try:
                    if kind == "close":
                        if document is not None:
                            document.close()
                            document = None
                        document_generation = None
                        self.closed.set()
                        continue
                    if cancel.is_set() or (view_cancel and view_cancel.is_set()):
                        continue
                    if kind == "open":
                        import pypdfium2 as pdfium

                        if document is not None:
                            document.close()
                            document = None
                        try:
                            document = pdfium.PdfDocument(value)
                        except pdfium.PdfiumError as error:
                            if "password" in str(error).lower():
                                raise ValueError("PDF password required") from error
                            raise
                        self.closed.clear()
                        document_generation = generation
                        self.rendered_count = 0
                        page_count = len(document)
                        if page_count == 0:
                            raise ValueError("PDF document contains no pages")
                        self._emit(("pdf_loaded", generation, value, page_count), cancel)
                    elif kind == "render":
                        if document is None or document_generation != generation:
                            raise ValueError("PDF document is not open")
                        ppm = self.render_page(document, value, self.width, self.height)
                        self.rendered_count += 1
                        self._emit(("thumbnail_ready", generation, request_id, value, ppm), cancel, view_cancel)
                except Exception as error:
                    if document is not None:
                        document.close()
                        document = None
                    self.closed.set()
                    event = (("pdf_load_failed", generation, str(error)) if kind == "open" else
                             ("thumbnail_failed", generation, request_id, value, str(error)))
                    # A broken document must not generate one failure per pending page.
                    with self.condition:
                        self.pending = deque(task for task in self.pending if task[1] != generation or task[0] == "close")
                    # Closing the document affects the session, even if the viewport moved.
                    self._emit(event, cancel)
                finally:
                    if cancel.is_set() and document is not None:
                        document.close()
                        document = None
                        self.closed.set()
        finally:
            if document is not None:
                document.close()
            self.closed.set()
