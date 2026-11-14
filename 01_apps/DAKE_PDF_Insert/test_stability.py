from __future__ import annotations

import ast
import hashlib
import queue
import tempfile
import threading
import time
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import fitz
from PIL import ImageTk
from pypdf import PdfReader, PdfWriter

import main as appmod
from thumbnail_worker import CACHE_BYTES, CACHE_PAGES, ThumbnailWorker


def make_pdf(path: Path, count: int, prefix="SRC"):
    with fitz.open() as doc:
        for index in range(count):
            page = doc.new_page(width=595, height=842)
            page.insert_text((45, 80), f"{prefix}{index+1}", fontsize=28)
        doc.save(path)


class PdfTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.source = self.folder / "元 PDF {確認}.pdf"
        self.insert = self.folder / "差込 PDF.pdf"
        make_pdf(self.source, 4)
        make_pdf(self.insert, 2, "INS")

    def test_order_and_non_destructive(self):
        hashes = [hashlib.sha256(path.read_bytes()).digest() for path in (self.source, self.insert)]
        for index, expected in ((0, ["INS1", "INS2", "SRC1", "SRC2", "SRC3", "SRC4"]), (2, ["SRC1", "SRC2", "INS1", "INS2", "SRC3", "SRC4"]), (4, ["SRC1", "SRC2", "SRC3", "SRC4", "INS1", "INS2"])):
            out = self.folder / f"完成 {index}.pdf"
            self.assertEqual(appmod.insert_pdf_pages(self.source, self.insert, out, index), 6)
            with fitz.open(out) as doc:
                self.assertEqual([page.get_text().strip() for page in doc], expected)
        self.assertEqual(hashes, [hashlib.sha256(path.read_bytes()).digest() for path in (self.source, self.insert)])

    def test_failure_has_no_partial_pdf(self):
        out = self.folder / "failed.pdf"
        def broken_write(_writer, stream):
            stream.write(b"partial pdf")
            raise OSError("injected disk failure")
        with patch.object(PdfWriter, "write", broken_write), self.assertRaises(OSError):
            appmod.insert_pdf_pages(self.source, self.insert, out, 2)
        self.assertFalse(out.exists())
        self.assertEqual(list(self.folder.glob(".dake-insert-*.tmp")), [])

    def test_cancel_and_existing_output(self):
        cancelled = threading.Event()
        cancelled.set()
        out = self.folder / "cancelled.pdf"
        with self.assertRaises(appmod.SaveCancelled):
            appmod.insert_pdf_pages(self.source, self.insert, out, 2, cancelled)
        self.assertFalse(out.exists())
        out.write_bytes(b"existing")
        with self.assertRaises(FileExistsError):
            appmod.insert_pdf_pages(self.source, self.insert, out, 2)
        self.assertEqual(out.read_bytes(), b"existing")
        self.assertEqual(list(self.folder.glob(".dake-insert-*.tmp")), [])

    def test_invalid_inputs(self):
        bad = self.folder / "broken.pdf"
        bad.write_bytes(b"not a PDF")
        with self.assertRaises(appmod.PdfInsertError):
            appmod.read_pdf_info(bad)
        with self.assertRaises(appmod.PdfInsertError):
            appmod.read_pdf_info(self.folder/"missing.pdf")
        with self.assertRaises(appmod.PdfInsertError):
            appmod.read_pdf_info(self.folder/"not-pdf.txt")
        with self.assertRaises(appmod.PdfInsertError):
            appmod.insert_pdf_pages(self.source, self.insert, self.source, 0)

    def test_cancel_during_write_cleans_temporary_file(self):
        cancelled = threading.Event()
        original = PdfWriter.write
        def cancel_write(writer, stream):
            result = original(writer, stream)
            cancelled.set()
            return result
        out = self.folder / "mid-write.pdf"
        with patch.object(PdfWriter, "write", cancel_write), self.assertRaises(appmod.SaveCancelled):
            appmod.insert_pdf_pages(self.source, self.insert, out, 2, cancelled)
        self.assertFalse(out.exists())
        self.assertEqual(list(self.folder.glob(".dake-insert-*.tmp")), [])

    def test_worker_close_with_full_result_queue(self):
        results = queue.Queue(maxsize=1)
        worker = ThumbnailWorker(results)
        try:
            worker.load("source", 1, self.source)
            self.assertEqual(results.get(timeout=5)[0], "pdf_metadata")
            worker.render(1, 1, self.source, (0,), 165, 230)
            deadline = time.perf_counter()+5
            while results.empty() and time.perf_counter() < deadline:
                time.sleep(.002)
            self.assertFalse(results.empty())
            worker.close()
            worker.thread.join(timeout=2)
            self.assertFalse(worker.thread.is_alive())
        finally:
            worker.close()


class GuiTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.folder = Path(self.temp.name)
        self.source = self.folder / ("日本語 {契約} " + "長い名前" * 20 + ".pdf")
        self.insert = self.folder / "差込 PDF.pdf"
        make_pdf(self.source, 300)
        make_pdf(self.insert, 2, "INS")
        self.root = appmod.BASE_WINDOW()
        self.app = appmod.DakePdfInsertApp(self.root)
        self.errors = []
        self.root.report_callback_exception = lambda *args: self.errors.append(args)
        self.thread_id = threading.get_ident()
        original_photo = ImageTk.PhotoImage
        def checked_photo(*args, **kwargs):
            self.assertEqual(threading.get_ident(), self.thread_id)
            return original_photo(*args, **kwargs)
        self.photo_patch = patch.object(ImageTk, "PhotoImage", checked_photo)
        self.photo_patch.start()
        self.max_gap = 0
        self.last_tick = time.perf_counter()
        self.beat_job = None
        def heartbeat():
            now = time.perf_counter()
            self.max_gap = max(self.max_gap, now-self.last_tick)
            self.last_tick = now
            if not self.app._closing:
                self.beat_job = self.root.after(20, heartbeat)
        self.beat_job = self.root.after(0, heartbeat)

    def tearDown(self):
        self.close()
        if self.app.thumbnail_worker is not None:
            self.app.thumbnail_worker.thread.join(timeout=5)
            self.assertFalse(self.app.thumbnail_worker.thread.is_alive())
        self.photo_patch.stop()
        self.temp.cleanup()
        self.assertEqual(self.errors, [])

    def close(self):
        if not self.app._closing:
            if self.beat_job is not None:
                self.root.after_cancel(self.beat_job)
                self.beat_job = None
            self.app._on_close()

    def wait(self, predicate, timeout=8):
        deadline = time.perf_counter()+timeout
        while not predicate():
            self.root.update()
            if time.perf_counter() > deadline:
                self.fail("GUI timeout")
            time.sleep(.002)
        self.root.update()

    def settled(self):
        a = self.app
        return not a._view_pending and a._view_job is None and a.resize_debounce_job is None and all(a._photo_sizes.get(p) == (a.thumbnail_display_size, a._image_height) for p in a._visible_pages)

    def drop(self, kind, *paths):
        data = self.root.tk.call("format", "%s", self.root.tk.call("list", *(str(p) for p in paths)))
        self.app._on_drop(SimpleNamespace(data=data), kind)

    def test_300_pages_resize_scroll_reload_and_close(self):
        a = self.app
        self.assertEqual(str(a.thumbnail_scale.cget("state")), "disabled")
        self.assertIsNone(a.thumbnail_worker)
        self.drop("source", self.source)
        self.assertIn("日本語", a.source_name_var.get())
        self.assertEqual(str(a.thumbnail_scale.cget("state")), "disabled")
        self.drop("insert", self.insert)
        self.wait(lambda: a.thumbnail_ready and a.insert_info is not None and self.settled())
        self.assertEqual(str(a.thumbnail_scale.cget("state")), "normal")
        self.assertEqual(a.source_info.pages, 300)
        self.assertLess(a.thumbnail_worker.render_count, 12)
        for value in [110, 260]*10+[210]:
            a._on_thumbnail_scale(str(value))
            self.root.update()
        self.wait(lambda: self.settled() and a.thumbnail_display_size == 210)
        for fraction in (.2, .4, .6, .8, .95, .5, .25):
            a._on_scrollbar("moveto", fraction)
            self.wait(self.settled)
        old_page = a._visible_pages[0]
        a._on_thumbnail_scale("140")
        self.wait(lambda: self.settled() and a.thumbnail_display_size == 140)
        self.assertLessEqual(abs(a._visible_pages[0]-old_page), 1)
        self.assertLessEqual(len(a.thumbnail_photos), 16)
        self.assertLessEqual(len(a.thumbnail_worker.cache), CACHE_PAGES)
        self.assertLessEqual(a.thumbnail_worker.cache_peak_bytes, CACHE_BYTES)
        for position in (0, 2, 300, 1):
            a._select_position(position)
            self.assertEqual(a.selected_position, position)
        a._on_thumbnail_click(SimpleNamespace(x=200, y=0))
        self.assertEqual(a.selected_position, 1)
        old_id = a.source_load_id
        self.drop("source", self.source, self.insert)
        self.assertEqual(a.source_load_id, old_id)
        self.assertEqual(a.status_var.get(), appmod.UI_TEXT["status_error"])
        short = self.folder / "replacement.pdf"
        make_pdf(short, 4)
        self.drop("source", self.source)
        self.drop("source", short)
        self.wait(lambda: a.thumbnail_ready and a.source_info.pages == 4 and self.settled())
        photos = dict(a.thumbnail_photos)
        a._handle_queue_message(("view_thumbnail", old_id, 0, 0, 110, 150, None))
        self.assertEqual(photos, a.thumbnail_photos)
        for geometry in ("1062x740", "920x680", "1400x900", "1062x740"):
            self.root.geometry(geometry)
            self.root.update()
            self.wait(self.settled)
            for widget in (a.execute_button, a.folder_button, a.footer_left, a.footer_right, a.status_badge):
                y = widget.winfo_rooty()-self.root.winfo_rooty()
                self.assertGreaterEqual(y, 0)
                self.assertLessEqual(y+widget.winfo_height(), self.root.winfo_height())
                self.assertGreaterEqual(widget.winfo_width(), widget.winfo_reqwidth())
        self.assertLess(self.max_gap, .5)
        print(f"GUI 300 pages max heartbeat gap {self.max_gap:.3f}s; cache peak {a.thumbnail_worker.cache_peak_bytes/1024**2:.2f} MiB")
        self.drop("source", self.source)
        self.close()

    def test_close_during_save(self):
        a = self.app
        self.drop("source", self.source)
        self.drop("insert", self.insert)
        self.wait(lambda: a.thumbnail_ready and a.insert_info is not None)
        a._select_position(2)
        a._execute_insert()
        self.assertTrue(a.is_busy)
        self.close()
        for thread in threading.enumerate():
            if thread.name == "pdf-save":
                thread.join(timeout=8)
                self.assertFalse(thread.is_alive())
        self.assertEqual(list(self.folder.glob(".dake-insert-*.tmp")), [])
        self.assertFalse(list(self.folder.glob("*_差し込み済み.pdf")))


class TextTests(unittest.TestCase):
    def test_ui_text_and_worker_boundaries(self):
        import re
        source = Path(appmod.__file__).read_text(encoding="utf-8")
        tree = ast.parse(source)
        japanese = re.compile(r"[ぁ-んァ-ン一-龥]")
        for node in ast.walk(tree):
            if isinstance(node, ast.Call):
                for kw in node.keywords:
                    if kw.arg in {"text", "title", "message"} and isinstance(kw.value, ast.Constant) and isinstance(kw.value.value, str):
                        self.assertIsNone(japanese.search(kw.value.value), (node.lineno, kw.value.value))
        worker = (Path(appmod.__file__).parent/"thumbnail_worker.py").read_text(encoding="utf-8")
        self.assertNotIn("tkinter", worker)
        self.assertNotIn("PhotoImage", worker)


class ScalingTests(unittest.TestCase):
    def test_layout_at_font_scaling_100_125_150(self):
        for factor in (1, 1.25, 1.5):
            root = appmod.BASE_WINDOW()
            root.tk.call("tk", "scaling", 96/72*factor)
            app = appmod.DakePdfInsertApp(root)
            try:
                root.geometry("1062x740")
                root.update()
                for widget in (app.folder_button, app.execute_button, app.footer_left, app.footer_right, app.status_badge):
                    self.assertLessEqual(widget.winfo_rooty()+widget.winfo_height(), root.winfo_rooty()+root.winfo_height(), factor)
                    self.assertGreaterEqual(widget.winfo_rootx(), root.winfo_rootx(), factor)
                    self.assertLessEqual(widget.winfo_rootx()+widget.winfo_width(), root.winfo_rootx()+root.winfo_width(), factor)
                    self.assertGreaterEqual(widget.winfo_width(), widget.winfo_reqwidth(), factor)
            finally:
                app._on_close()


if __name__ == "__main__":
    unittest.main(verbosity=2)
