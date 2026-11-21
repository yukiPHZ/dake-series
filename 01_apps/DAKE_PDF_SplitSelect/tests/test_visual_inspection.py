"""Real Tk/PDFium regression; injected input is not physical mouse/DnD evidence."""

import hashlib
import gc
import json
import sys
import time
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parent))
import test_lifecycle as lifecycle
import main


class VisualInspectionTests(unittest.TestCase):
    setUpClass = classmethod(lifecycle.LifecycleTests.setUpClass.__func__)
    tearDownClass = classmethod(lifecycle.LifecycleTests.tearDownClass.__func__)
    spin = lifecycle.LifecycleTests.spin
    settle = lifecycle.LifecycleTests.settle
    load = lifecycle.LifecycleTests.load

    def setUp(self):
        # Multiple disposable Tk interpreters must be collected on the test/UI thread.
        gc.collect()
        lifecycle.LifecycleTests.setUp(self)

    def tearDown(self):
        self.app._close_preview()
        lifecycle.LifecycleTests.tearDown(self)
        self.app = None
        gc.collect()

    def snapshot(self):
        app = self.app
        return (set(app.selected_pages), app.range_var.get(), app.thumbnail_selection_anchor,
                app.save_dir, app.current_pdf_path)

    def open_preview(self, page=0):
        start = time.perf_counter()
        self.app._open_preview(page)
        self.spin(lambda: self.app.preview is not None and self.app.preview.photo is not None)
        print(json.dumps({"preview_seconds": time.perf_counter() - start,
                          "pixels": self.app.preview.photo.width() * self.app.preview.photo.height()}))
        return self.app.preview

    def test_sizes_and_anchor_eviction(self):
        for count in (100, 300):
            self.load(count)
            vp = self.app.thumbnail_viewport
            self.app._toggle_thumbnail_page(5)
            self.app.range_var.set("1-3,5,8-10")
            before = self.snapshot()
            vp._on_scrollbar("moveto", .45)
            self.spin(lambda: vp._redraw_job is None)
            self.settle()
            anchor = vp._capture_anchor()[0]
            for key in ("small", "standard", "large", "extra", "small"):
                start = time.perf_counter()
                self.app.thumbnail_size_var.set(key)
                self.app._change_thumbnail_size()
                self.spin(lambda: vp._redraw_job is None)
                self.settle()
                self.assertEqual(self.snapshot(), before)
                self.assertLessEqual(abs(vp._capture_anchor()[0] - anchor), max(6, vp.columns))
                self.assertTrue(all(image.width() <= vp.thumb_width and image.height() <= vp.thumb_height
                                    for image in vp.thumbnail_cache.values()))
                print(json.dumps({"size": key, "pages": count, "seconds": time.perf_counter() - start,
                                  "cache_bytes": vp.cache_bytes}))
            for step in range(41):
                vp._on_scrollbar("moveto", step / 40)
                self.spin(lambda: vp._redraw_job is None)
                self.settle()
                self.assertLessEqual(len(vp.thumbnail_cache), 96)
                self.assertLessEqual(len(vp.thumbnail_cache) + len(vp._size_cache), 96)
                self.assertLessEqual(vp.cache_bytes, main.THUMBNAIL_CACHE_BYTES)
            self.assertIn(count - 1, vp.visible_pages)
            self.assertNotIn(0, vp.thumbnail_cache)
            vp._on_scrollbar("moveto", 0)
            self.spin(lambda: vp._redraw_job is None)
            self.settle()
            self.assertIn(0, vp.thumbnail_cache)
            self.assertEqual(self.snapshot(), before)
            for key in list(main.THUMBNAIL_SIZES) * 8:
                self.app.thumbnail_size_var.set(key)
                self.app._change_thumbnail_size()
            self.spin(lambda: vp._redraw_job is None)
            self.settle()
            self.assertLessEqual(len(self.app.render_worker.pending), len(vp.visible_pages) + 16)
            self.app.refresh_all()

    def test_reselecting_current_size_keeps_pending_render(self):
        self.load(300)
        vp = self.app.thumbnail_viewport
        vp._on_scrollbar("moveto", .8)
        vp._redraw()
        request_id = self.app.render_request_id
        self.app._change_thumbnail_size()
        self.assertEqual(self.app.render_request_id, request_id)
        self.settle()

    def test_cached_resolution_is_reused(self):
        self.load(30)
        self.settle()
        for key in ("small", "standard"):
            before = self.app.render_worker.rendered_count
            self.app.thumbnail_size_var.set(key)
            self.app._change_thumbnail_size()
            self.spin(lambda: self.app.thumbnail_viewport._redraw_job is None)
            self.settle()
            if key == "standard":
                self.assertEqual(self.app.render_worker.rendered_count, before)
        vp = self.app.thumbnail_viewport
        expected = sum(image.width() * image.height() * 4
                       for image in list(vp.thumbnail_cache.values()) + list(vp._size_cache.values()))
        self.assertEqual(vp.cache_bytes, expected)

    def test_preview_bindings_and_read_only(self):
        self.load(30)
        app, vp = self.app, self.app.thumbnail_viewport
        app.range_var.set("1-3,5,8-10")
        vp.canvas.event_generate("<Button-1>", x=40, y=40)
        self.assertEqual(app.thumbnail_selected_pages, {1})
        before = self.snapshot()
        vp.canvas.event_generate("<Control-Button-1>", x=40, y=40)
        self.spin(lambda: app.preview is not None and app.preview.photo is not None)
        preview = app.preview
        self.assertEqual(self.snapshot(), before)
        window = preview.window
        vp.canvas.event_generate("<Button-3>", x=vp.card_width + main.CARD_GAP_X + 40, y=40)
        self.spin(lambda: preview.photo is not None and preview.page == 1)
        self.assertIs(preview.window, window)
        self.assertEqual(self.snapshot(), before)
        vp._on_click(SimpleNamespace(x=40, y=40, state=main.CONTROL_MASK))
        self.assertEqual(self.snapshot(), before)
        app._close_preview()
        vp.canvas.event_generate("<Button-3>", x=2, y=2)
        self.assertIsNone(app.preview)
        vp.canvas.event_generate("<Shift-Button-1>", x=vp.card_width + main.CARD_GAP_X + 40, y=40)
        self.assertEqual(app.thumbnail_selected_pages, {1, 2})
        vp.canvas.event_generate("<Button-1>", x=40, y=40)
        self.assertEqual(app.thumbnail_selected_pages, {2})

    def test_high_resolution_zoom_pan_resize_and_close(self):
        self.load(3)
        before = self.snapshot()
        preview = self.open_preview()
        base_width = preview.photo.width()
        for _ in range(15):
            preview.on_wheel(SimpleNamespace(x=300, y=300, delta=120))
        self.spin(lambda: preview.job is None and preview.photo.width() > base_width * 2)
        self.assertEqual(preview.zoom, 3)
        self.assertLessEqual(preview.photo.width() * preview.photo.height(), 12_010_000)
        preview.pan_start(SimpleNamespace(x=300, y=300))
        preview.pan_move(SimpleNamespace(x=-200, y=-200))
        self.assertLessEqual(preview.offset[0], 0)
        preview.window.geometry("700x550")
        self.spin(lambda: preview.job is None and preview.canvas.winfo_width() < 750)
        for _ in range(25):
            preview.on_wheel(SimpleNamespace(x=200, y=200, delta=-120))
        self.spin(lambda: preview.job is None and preview.photo.width() < base_width)
        self.assertEqual(preview.zoom, .5)
        self.assertEqual(self.snapshot(), before)

    def test_maximum_window_retains_real_300_percent(self):
        self.load(3)
        before = self.snapshot()
        preview = self.open_preview()
        preview.window.geometry("1400x1000")
        self.spin(lambda: preview.job is None and abs(preview.photo.height() - preview.canvas.winfo_height()) <= 1)
        fit_height = preview.photo.height()
        x, y = preview.canvas.winfo_width() / 2, preview.canvas.winfo_height() / 2
        anchor = preview.point_anchor(x, y)
        for _ in range(15):
            preview.on_wheel(SimpleNamespace(x=x, y=y, delta=120))
        self.spin(lambda: preview.job is None and preview.photo.height() >= fit_height * 2.99)
        self.assertAlmostEqual(preview.photo.height() / fit_height, 3, delta=.01)
        after = preview.point_anchor(x, y)
        self.assertAlmostEqual(anchor[0], after[0], delta=.01)
        self.assertAlmostEqual(anchor[1], after[1], delta=.01)
        self.assertLessEqual(preview.photo.width() * preview.photo.height(), 12_010_000)
        preview.window.event_generate("<Escape>")
        self.app.update()
        self.assertIsNone(self.app.preview)
        self.assertEqual(self.snapshot(), before)

    def test_preview_cancel_refresh_switch_close_handle(self):
        self.load(300)
        self.open_preview()
        old_generation = self.app.render_generation
        old_id = self.app.preview_request_id
        original = self.app.render_worker.render_page
        started = main.threading.Event()
        release = main.threading.Event()
        def controlled(*args):
            if args[2] > main.THUMBNAIL_WIDTH:
                started.set()
                release.wait(3)
            return original(*args)
        with patch.object(self.app.render_worker, "render_page", controlled):
            self.app.preview.on_wheel(SimpleNamespace(x=100, y=100, delta=120))
            self.spin(started.is_set)
            self.app.refresh_all()
            self.assertIsNone(self.app.preview)
            release.set()
            self.load(3)
            self.settle()
            self.assertEqual(set(self.app.thumbnail_viewport.thumbnail_cache), {0, 1, 2})
        self.app._handle_worker_event(("preview_ready", old_generation, old_id, 0, b"stale"))
        self.assertIsNone(self.app.preview)
        self.open_preview()
        self.app._load_pdf_async(str(self.pdfs[30]))
        self.assertIsNone(self.app.preview)
        self.settle()
        self.open_preview()
        self.app._on_close()
        self.app.render_worker.thread.join(5)
        self.assertTrue(self.app.render_worker.closed.is_set())
        # Reopening exclusively after worker shutdown checks Windows sharing restrictions.
        self.pdfs[300].rename(self.root / "renamed-300.pdf").rename(self.pdfs[300])
        self.assertEqual(hashlib.sha256(self.pdfs[300].read_bytes()).hexdigest(), self.hashes[300])

    def test_refresh_releases_loaded_copy_for_delete(self):
        copy = self.root / "handle-test.pdf"
        copy.write_bytes(self.pdfs[300].read_bytes())
        self.app._load_pdf_async(str(copy))
        self.settle()
        self.open_preview()
        self.app.refresh_all()
        self.spin(self.app.render_worker.closed.is_set)
        moved = self.root / "moved-handle"
        moved.mkdir(exist_ok=True)
        renamed = copy.rename(self.root / "renamed-handle.pdf")
        renamed.rename(moved / "delete-test.pdf").unlink()
        self.assertFalse(copy.exists())

    def test_preview_priority_and_memory_measurement(self):
        import ctypes
        from ctypes import wintypes
        self.load(300)
        self.settle()
        calls = []
        entered, release = main.threading.Event(), main.threading.Event()
        original = self.app.render_worker.render_page
        def record(doc, page, width, height):
            calls.append((page, width, main.threading.get_ident()))
            if len(calls) == 1:
                entered.set()
                release.wait(3)
            return original(doc, page, width, height)
        worker = self.app.render_worker
        with patch.object(worker, "render_page", record):
            worker.request(self.app.render_generation, 999, list(range(100, 124)))
            self.spin(entered.is_set)
            worker.request_preview(self.app.render_generation, 999, 2, (900, 700))
            release.set()
            self.spin(lambda: len(calls) >= 3)
        self.assertEqual(calls[1][:2], (2, 900))
        self.assertEqual({item[2] for item in calls}, {worker.thread.ident})
        self.assertLessEqual(worker.preview_results.qsize(), 1)
        self.app.refresh_all()
        for count in (3, 30, 100, 300):
            self.load(count)
            self.open_preview()
            preview = self.app.preview
            start = time.perf_counter()
            for _ in range(15):
                preview.on_wheel(SimpleNamespace(x=300, y=300, delta=120))
            self.spin(lambda: preview.job is None and preview.photo.height() > 1000)
            class Counters(ctypes.Structure):
                _fields_ = [("cb", wintypes.DWORD), ("PageFaultCount", wintypes.DWORD)] + [
                    (name, ctypes.c_size_t) for name in ("PeakWorkingSetSize", "WorkingSetSize",
                    "QuotaPeakPagedPoolUsage", "QuotaPagedPoolUsage", "QuotaPeakNonPagedPoolUsage",
                    "QuotaNonPagedPoolUsage", "PagefileUsage", "PeakPagefileUsage", "PrivateUsage")]
            counters = Counters()
            counters.cb = ctypes.sizeof(counters)
            kernel = ctypes.WinDLL("kernel32")
            kernel.GetCurrentProcess.restype = wintypes.HANDLE
            psapi = ctypes.WinDLL("psapi")
            psapi.GetProcessMemoryInfo.argtypes = [wintypes.HANDLE, ctypes.c_void_p, wintypes.DWORD]
            self.assertTrue(psapi.GetProcessMemoryInfo(kernel.GetCurrentProcess(), ctypes.byref(counters), counters.cb))
            print(json.dumps({"pages": count, "zoom_seconds": time.perf_counter() - start,
                              "working_set_bytes": counters.WorkingSetSize,
                              "peak_working_set_bytes": counters.PeakWorkingSetSize,
                              "private_bytes": counters.PrivateUsage,
                              "preview_pixels": preview.photo.width() * preview.photo.height(),
                              "cache_bytes": self.app.thumbnail_viewport.cache_bytes}), flush=True)
            self.app.refresh_all()

    def test_preview_failure_does_not_close_main_document(self):
        self.load(3)
        original = self.app.render_worker.render_page
        def fail_preview(doc, page, width, height):
            if width > main.THUMBNAIL_WIDTH:
                raise RuntimeError("controlled preview failure")
            return original(doc, page, width, height)
        with patch.object(self.app.render_worker, "render_page", fail_preview):
            self.app._open_preview(0)
            self.spin(lambda: any("controlled preview failure" in str(self.app.preview.canvas.itemcget(i, "text"))
                                  for i in self.app.preview.canvas.find_withtag("message")))
        self.assertFalse(self.app.render_worker.closed.is_set())
        self.assertEqual(self.app.page_count, 3)
        self.app._close_preview()
        self.open_preview(1)

    def test_shortcuts_and_preview_window_manager_close(self):
        with patch.object(main.filedialog, "askopenfilename", return_value=str(self.pdfs[3])) as dialog:
            self.app.focus_force()
            self.app.update()
            self.app.event_generate("<Control-o>")
            self.spin(lambda: self.app.document_ready)
            dialog.assert_called_once()
        preview = self.open_preview()
        before = self.snapshot()
        # Execute the registered WM protocol, rather than bypassing its callback.
        preview.window.tk.call(preview.window.protocol("WM_DELETE_WINDOW"))
        self.assertIsNone(self.app.preview)
        self.assertEqual(self.snapshot(), before)
        preview = self.open_preview()
        preview.window.focus_force()
        self.app.update()
        preview.window.event_generate("<F5>")
        self.app.update()
        self.assertIsNone(self.app.preview)
        self.assertIsNone(self.app.current_pdf_path)


if __name__ == "__main__":
    unittest.main(verbosity=2)
