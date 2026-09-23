"""Run with the verified Windows interpreter; generated PDFs stay in TemporaryDirectory."""

import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import time
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import main


def write_fixture(path, count, *, landscape=False, rotated=False, image=False):
    from pypdf import PdfWriter
    from pypdf.generic import DecodedStreamObject, DictionaryObject, NameObject, NumberObject

    writer = PdfWriter()
    font = DictionaryObject({
        NameObject("/Type"): NameObject("/Font"),
        NameObject("/Subtype"): NameObject("/Type1"),
        NameObject("/BaseFont"): NameObject("/Helvetica"),
    })
    font_ref = writer._add_object(font)
    for index in range(count):
        page = writer.add_blank_page(width=842 if landscape else 595, height=595 if landscape else 842)
        resources = DictionaryObject({NameObject("/Font"): DictionaryObject({NameObject("/F1"): font_ref})})
        drawing = f"{index / max(count, 1):.3f} 0.4 0.8 rg 50 100 250 300 re f\n"
        drawing += f"BT /F1 36 Tf 72 510 Td (PAGE {index + 1:03}) Tj ET\n"
        if image:
            pixels = bytes(component for y in range(8) for x in range(8)
                           for component in (x * 32, y * 32, 128))
            bitmap = DecodedStreamObject()
            bitmap.set_data(pixels)
            bitmap = bitmap.flate_encode()
            bitmap.update({
                NameObject("/Type"): NameObject("/XObject"),
                NameObject("/Subtype"): NameObject("/Image"),
                NameObject("/Width"): NumberObject(8),
                NameObject("/Height"): NumberObject(8),
                NameObject("/ColorSpace"): NameObject("/DeviceRGB"),
                NameObject("/BitsPerComponent"): NumberObject(8),
            })
            resources[NameObject("/XObject")] = DictionaryObject({NameObject("/Im1"): writer._add_object(bitmap)})
            drawing += "q 200 0 0 200 310 250 cm /Im1 Do Q\n"
        page[NameObject("/Resources")] = resources
        stream = DecodedStreamObject()
        stream.set_data(drawing.encode("ascii"))
        page[NameObject("/Contents")] = writer._add_object(stream)
        if rotated:
            page.rotate(90)
    with path.open("wb") as output:
        writer.write(output)


class LifecycleTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp = tempfile.TemporaryDirectory(prefix="dake-regression-")
        cls.root = Path(cls.temp.name)
        # Importing the UI module must not initialize any PDF/image backend.
        check = subprocess.run(
            [sys.executable, "-c", "import main, sys; assert not any(name in sys.modules for name in "
             "('pypdfium2', 'pypdf', 'PIL.Image', 'PIL.ImageTk'))"],
            cwd=Path(main.__file__).parent, capture_output=True,
        )
        assert check.returncode == 0, check.stderr
        cls.pdfs = {}
        for count in (3, 30, 100, 300):
            path = cls.root / f"pages-{count}.pdf"
            write_fixture(path, count)
            cls.pdfs[count] = path
        cls.visual_pdfs = {}
        for name, options in (("landscape", {"landscape": True}),
                              ("rotated", {"rotated": True}),
                              ("scanned", {"image": True})):
            path = cls.root / f"{name}.pdf"
            write_fixture(path, 1, **options)
            cls.visual_pdfs[name] = path
        cls.hashes = {n: hashlib.sha256(p.read_bytes()).hexdigest() for n, p in cls.pdfs.items()}

    @classmethod
    def tearDownClass(cls):
        cls.temp.cleanup()

    def setUp(self):
        self.app = main.DakePdfSplitSelectApp()
        self.app.update()
        self.max_queue = 0
        self.requests = []
        request = self.app.render_worker.request
        def record_request(generation, request_id, pages):
            self.requests.append(tuple(pages))
            request(generation, request_id, pages)
        self.app.render_worker.request = record_request

    def tearDown(self):
        self.app.render_worker.stop()
        self.app.render_worker.thread.join(5)
        self.assertFalse(self.app.render_worker.thread.is_alive())
        try:
            if not self.app.winfo_exists():
                return
        except main.tk.TclError:
            return
        if self.app._poll_job:
            self.app.after_cancel(self.app._poll_job)
        if self.app.thumbnail_viewport._redraw_job:
            self.app.after_cancel(self.app.thumbnail_viewport._redraw_job)
        self.app.destroy()

    def spin(self, predicate, timeout=15):
        deadline = time.perf_counter() + timeout
        while time.perf_counter() < deadline:
            self.app.update()
            with self.app.render_worker.condition:
                self.max_queue = max(self.max_queue, len(self.app.render_worker.pending))
            if predicate():
                return
            time.sleep(0.003)
        self.fail(f"Timeout: {self.app.current_status}, pending={list(self.app.render_worker.pending)}")

    def settle(self):
        self.spin(lambda: self.app.document_ready and not self.app.queued_thumbnail_pages)
        self.app.update()

    def load(self, count):
        self.requests.clear()
        started = time.perf_counter()
        self.app._load_pdf_async(str(self.pdfs[count]))
        self.spin(lambda: self.app.document_ready)
        elapsed = time.perf_counter() - started
        worker = self.app.render_worker
        vp = self.app.thumbnail_viewport
        self.assertTrue(all(page in vp.thumbnail_cache for page in vp.visible_pages))
        self.assertLessEqual(worker.rendered_count, len(vp.visible_pages) + 16)
        self.assertEqual(self.requests[0][:len(vp.visible_pages)], vp.visible_pages)
        print(json.dumps({"pages": count, "ready_seconds": round(elapsed, 4),
                          "rendered_at_ready": worker.rendered_count,
                          "visible": len(vp.visible_pages), "initial_requested": len(self.requests[0]),
                          "max_pending_observed": self.max_queue}), flush=True)

    def test_page_sizes_cache_and_selection(self):
        for count in (3, 30, 100, 300):
            self.load(count)
            self.settle()
            self.assertEqual(self.app.current_status.state, main.AppState.READY)
            self.app._toggle_thumbnail_page(1)
            self.assertEqual(self.app.current_status.state, main.AppState.SELECTING)
            self.app._toggle_thumbnail_page(3, True)
            self.assertEqual(self.app.selected_pages, {1, 2, 3})
            self.app._toggle_thumbnail_page(1, True)
            self.assertFalse(self.app.selected_pages)
            self.app._toggle_thumbnail_page(1)
            self.app.range_var.set("1-3" if count == 3 else "1-3,5,8-10")
            selected_range = set(self.app.range_selected_pages)
            vp = self.app.thumbnail_viewport
            for step in range(1, 41):
                vp._on_scrollbar("moveto", step / 40)
                self.spin(lambda: vp._redraw_job is None)
                self.settle()
                self.assertLessEqual(len(vp.thumbnail_cache), main.THUMBNAIL_CACHE_LIMIT)
            if count >= 100:
                self.assertNotIn(0, vp.thumbnail_cache)
            before = self.app.render_worker.rendered_count
            vp._on_scrollbar("moveto", 0)
            self.spin(lambda: vp._redraw_job is None)
            self.settle()
            self.assertIn(0, vp.thumbnail_cache)
            if count >= 100:
                self.assertGreater(self.app.render_worker.rendered_count, before)
            self.assertEqual(self.app.thumbnail_selected_pages, {1})
            self.assertEqual(self.app.range_selected_pages, selected_range)
            self.app.range_var.set("1--3")
            self.assertEqual(str(self.app.extract_merged_button["state"]), "disabled")
            self.app.clear_selection()
            self.assertEqual(self.app.current_status.state, main.AppState.READY)
            print(json.dumps({"cache_pages": count, "size": len(vp.thumbnail_cache),
                              "total_renders": self.app.render_worker.rendered_count}), flush=True)
            self.app.refresh_all()

    def test_refresh_cancel_and_handle(self):
        original = self.app.render_worker.render_page

        def slow_page(*args, **kwargs):
            time.sleep(0.08)
            return original(*args, **kwargs)

        for immediate_new_pdf in (False, True):
            with patch.object(self.app.render_worker, "render_page", slow_page):
                self.app._load_pdf_async(str(self.pdfs[300]))
                self.spin(lambda: bool(self.app.queued_thumbnail_pages))
                old_generation = self.app.render_generation
                old_cancel = self.app.render_worker.cancel
                self.app.save_dir = str(self.root)
                self.app.save_dir_is_manual = True
                self.app.range_var.set("invalid")
                started = time.perf_counter()
                self.app.refresh_all()
                self.assertLess(time.perf_counter() - started, 0.1)
                self.assertTrue(old_cancel.is_set())
                self.assertIsNone(self.app.current_pdf_path)
                self.assertEqual(self.app.state_var.get(), main.UI_TEXT["status_idle"])
                self.assertEqual(self.app.range_var.get(), "")
                self.assertEqual(self.app.save_dir, str(self.root))
                self.assertFalse(self.app.thumbnail_viewport.thumbnail_cache)
                self.assertIsNone(self.app.thumbnail_selection_anchor)
                if immediate_new_pdf:
                    self.load(3)
                    self.settle()
                    self.assertEqual(self.app.page_count, 3)
                    self.assertEqual(set(self.app.thumbnail_viewport.thumbnail_cache), {0, 1, 2})
                else:
                    self.spin(self.app.render_worker.closed.is_set)
                    count = self.app.render_worker.rendered_count
                    time.sleep(0.15)
                    self.app.update()
                    self.assertEqual(count, self.app.render_worker.rendered_count)
                    self.assertFalse(self.app.thumbnail_viewport.thumbnail_cache)
                self.app._handle_worker_event(("pdf_loaded", old_generation, "stale.pdf", 300))
                self.assertNotEqual(self.app.current_pdf_path, "stale.pdf")
                renamed = self.root / "renamed.pdf"
                moved_dir = self.root / "moved"
                moved_dir.mkdir(exist_ok=True)
                self.pdfs[300].rename(renamed)
                moved = renamed.rename(moved_dir / "moved.pdf")
                moved.rename(self.pdfs[300])
                self.assertEqual(hashlib.sha256(self.pdfs[300].read_bytes()).hexdigest(), self.hashes[300])
            self.app.refresh_all()

    def test_save_and_cli(self):
        from pypdf import PdfReader
        self.load(30)
        patterns = ([1], [3], [5], [1, 2, 3, 4, 5], [1, 2, 3, 5, 8, 9, 10])
        for index, pages in enumerate(patterns):
            for mode in ("merged", "single"):
                directory = self.root / f"save-{index}-{mode}"
                self.app.save_dir = str(directory)
                self.app.save_dir_is_manual = True
                self.app.clear_selection()
                if index == 3:
                    self.app._toggle_thumbnail_page(1)
                    self.app._toggle_thumbnail_page(5, True)
                else:
                    self.app.range_var.set(",".join(map(str, pages)))
                with patch.object(main.messagebox, "showinfo") as dialog, patch.object(main, "open_directory") as opened:
                    self.app._start_extract(mode)
                    self.spin(lambda: not self.app.is_processing)
                    dialog.assert_called_once()
                    opened.assert_called_once_with(str(directory))
                files = sorted(directory.glob("*.pdf"))
                self.assertEqual(len(files), 1 if mode == "merged" else len(pages))
                actual = [int(page.extract_text().split()[1]) for path in files for page in PdfReader(path).pages]
                self.assertEqual(actual, pages)
                if mode == "single":
                    self.assertTrue(all(len(PdfReader(path).pages) == 1 for path in files))
        for expression, code in (("1-3,5,8-10", 0), ("1--3", 1)):
            output = self.root / f"cli-{code}.pdf"
            result = subprocess.run([sys.executable, str(Path(main.__file__)), "--from-shimarisu",
                                     "--inputs", str(self.pdfs[30]), "--pages", expression,
                                     "--output", str(output), "--silent"], capture_output=True)
            self.assertEqual(result.returncode, code, result.stderr)
            if code == 0:
                self.assertEqual(len(PdfReader(output).pages), 7)
        for count, path in self.pdfs.items():
            self.assertEqual(hashlib.sha256(path.read_bytes()).hexdigest(), self.hashes[count])

    def test_fast_scroll_drop_and_error_recovery(self):
        from types import SimpleNamespace
        self.app._on_drop(SimpleNamespace(data=self.app.tk.call("list", str(self.pdfs[300]))))
        self.spin(lambda: self.app.document_ready)
        vp = self.app.thumbnail_viewport
        for position in (1, 0, .8, .1, .9, .2, 1, 0):
            vp._on_scrollbar("moveto", position)
            vp._redraw()
        self.settle()
        self.assertIn(0, vp.thumbnail_cache)
        self.assertLessEqual(len(self.app.render_worker.pending), len(vp.visible_pages) + 16)
        self.app.refresh_all()
        self.spin(self.app.render_worker.closed.is_set)
        with patch.object(self.app.render_worker, "render_page", side_effect=RuntimeError("test render failure")):
            self.app._load_pdf_async(str(self.pdfs[3]))
            self.spin(lambda: self.app.current_status.state == main.AppState.ERROR)
            self.assertIn("test render failure", self.app.status_message_var.get())
            self.assertTrue(self.app.render_worker.closed.is_set())
        self.app.refresh_all()
        self.load(3)
        missing = self.root / "missing.pdf"
        self.app._load_pdf_async(str(missing))
        self.spin(lambda: self.app.current_status.state == main.AppState.ERROR)
        self.assertTrue(self.app.render_worker.closed.is_set())
        self.app.refresh_all()
        self.assertEqual(self.app.state_var.get(), main.UI_TEXT["status_idle"])

    def test_ppm_orientation_color_and_bitmap_lifetime(self):
        import pypdfium2 as pdfium
        from pdf_backend import render_page_ppm

        for name, path in self.visual_pdfs.items():
            document = pdfium.PdfDocument(str(path))
            try:
                ppm = render_page_ppm(document, 0, main.THUMBNAIL_WIDTH, main.THUMBNAIL_HEIGHT)
            finally:
                document.close()
            header, dimensions, max_value, pixels = ppm.split(b"\n", 3)
            width, height = map(int, dimensions.split())
            self.assertEqual((header, max_value), (b"P6", b"255"))
            self.assertEqual(len(pixels), width * height * 3)
            self.assertLessEqual(width, main.THUMBNAIL_WIDTH)
            self.assertLessEqual(height, main.THUMBNAIL_HEIGHT)
            self.assertGreater(len(set(pixels)), 2, name)
            if name == "landscape" or name == "rotated":
                self.assertGreater(width, height, name)
            else:
                self.assertLess(width, height, name)
            if name == "scanned":
                colors = set(zip(pixels[0::3], pixels[1::3], pixels[2::3]))
                self.assertGreater(len(colors), 16)
                offset = (140 * width + 40) * 3
                red, green, blue = pixels[offset:offset + 3]
                self.assertLess(red, green)
                self.assertLess(green, blue)

    def test_password_and_invalid_pdf_recovery(self):
        from pypdf import PdfWriter

        protected = self.root / "protected.pdf"
        writer = PdfWriter()
        writer.add_blank_page(width=100, height=100)
        writer.encrypt("secret")
        with protected.open("wb") as output:
            writer.write(output)
        self.app._load_pdf_async(str(protected))
        self.spin(lambda: self.app.current_status.state == main.AppState.ERROR)
        self.assertIn("password", self.app.status_message_var.get().lower())
        self.assertTrue(self.app.render_worker.closed.is_set())
        self.app.refresh_all()
        invalid = self.root / "invalid.pdf"
        invalid.write_bytes(b"not a PDF")
        self.app._load_pdf_async(str(invalid))
        self.spin(lambda: self.app.current_status.state == main.AppState.ERROR)
        self.assertTrue(self.app.render_worker.closed.is_set())
        self.app.refresh_all()
        self.load(3)

    def test_shutdown_and_save_guard(self):
        self.load(3)
        path = self.app.current_pdf_path
        self.app.is_processing = True
        self.app.refresh_all()
        self.app._load_pdf_async(str(self.pdfs[300]))
        self.app._on_close()
        self.assertEqual(self.app.current_pdf_path, path)
        self.assertFalse(self.app._closing)
        self.app.is_processing = False
        self.app._on_close()
        self.app.render_worker.thread.join(5)
        self.assertFalse(self.app.render_worker.thread.is_alive())
        self.assertTrue(self.app.render_worker.closed.is_set())


if __name__ == "__main__":
    unittest.main(verbosity=2)
