# -*- coding: utf-8 -*-
import hashlib
import importlib.util
import queue
import tempfile
import threading
import time
import unittest
from pathlib import Path
from unittest import mock

import sys


APP_DIR = Path(__file__).resolve().parents[1]
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

spec = importlib.util.spec_from_file_location("dake_pdf_splitone_main", APP_DIR / "main.py")
assert spec and spec.loader
splitone = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = splitone
spec.loader.exec_module(splitone)


def create_fixture_pdf(path: Path, page_count: int, width_base: int = 200) -> None:
    from pypdf import PdfWriter

    writer = PdfWriter()
    for index in range(page_count):
        writer.add_blank_page(width=width_base + index, height=300 + index)
    with path.open("wb") as stream:
        writer.write(stream)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def work_folders(root: Path) -> list[Path]:
    return sorted(root.glob(".*_split_work_*"))


def wait_until(predicate, timeout: float = 10.0) -> None:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if predicate():
            return
        time.sleep(0.01)
    raise AssertionError("timed out waiting for condition")


class FakeConfig:
    def __init__(self, save_folder: Path):
        self.last_save_folder = str(save_folder)


class FakeUi:
    def __init__(self):
        self.idle_count = 0
        self.statuses: list[tuple[str, str]] = []
        self.progress: list[str] = []
        self.completions: list[Path | None] = []
        self.errors: list[str] = []
        self.contexts: list[tuple[Path | None, Path | None]] = []
        self.enabled = True

    def update_save_folder(self, _folder: Path) -> None:
        pass

    def show_idle(self) -> None:
        self.idle_count += 1
        self.statuses.clear()
        self.progress.clear()
        self.contexts.clear()
        self.enabled = True

    def set_interaction_enabled(self, enabled: bool) -> None:
        self.enabled = enabled

    def show_file_context(
        self,
        source_path: Path | None,
        output_folder: Path | None,
    ) -> None:
        self.contexts.append((source_path, output_folder))

    def update_status(self, state_key: str, detail: str) -> None:
        self.statuses.append((state_key, detail))

    def update_progress(self, detail: str) -> None:
        self.progress.append(detail)

    def show_completion(self, output_folder: Path | None) -> None:
        self.completions.append(output_folder)

    def show_error(self, message: str) -> None:
        self.errors.append(message)


class LifecycleTests(unittest.TestCase):
    def split(self, source: Path, save_root: Path) -> Path:
        return splitone.PdfSplitService().split_all_pages(
            source_pdf=source,
            save_root=save_root,
            cancel_event=threading.Event(),
            on_loaded=lambda _total, _output: None,
            on_progress=lambda _current, _total, _name: None,
        )

    def test_normal_split_3_30_100_300(self) -> None:
        from pypdf import PdfReader

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            for page_count in (3, 30, 100, 300):
                source = root / f"sample_{page_count}.pdf"
                create_fixture_pdf(source, page_count)
                source_before = sha256(source)

                output = self.split(source, root)

                files = sorted(output.glob("p*.pdf"))
                self.assertEqual(page_count, len(files))
                for index, output_pdf in enumerate(files):
                    reader = PdfReader(str(output_pdf))
                    self.assertEqual(1, len(reader.pages))
                    self.assertEqual(
                        200 + index,
                        int(float(reader.pages[0].mediabox.width)),
                    )
                    reader.stream.close()
                self.assertEqual(source_before, sha256(source))
                self.assertEqual([], work_folders(root))

    def test_repeat_is_non_destructive(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "sample.pdf"
            create_fixture_pdf(source, 3)

            first = self.split(source, root)
            first_hashes = {item.name: sha256(item) for item in first.glob("*.pdf")}
            second = self.split(source, root)

            self.assertEqual("sample_split", first.name)
            self.assertEqual("sample_split_2", second.name)
            self.assertEqual(
                first_hashes,
                {item.name: sha256(item) for item in first.glob("*.pdf")},
            )
            self.assertEqual([], work_folders(root))

    def test_cancel_cleans_temp_without_formal_output(self) -> None:
        from pypdf import PdfWriter

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "cancel.pdf"
            create_fixture_pdf(source, 300)
            cancel_event = threading.Event()
            write_started = threading.Event()
            release_write = threading.Event()
            original_write = PdfWriter.write

            def delayed_write(writer, stream):
                write_started.set()
                release_write.wait(5.0)
                return original_write(writer, stream)

            errors: queue.Queue[BaseException] = queue.Queue()

            def run_split() -> None:
                try:
                    splitone.PdfSplitService().split_all_pages(
                        source,
                        root,
                        cancel_event,
                        lambda _total, _output: None,
                        lambda _current, _total, _name: None,
                    )
                except BaseException as exc:
                    errors.put(exc)

            with mock.patch.object(PdfWriter, "write", delayed_write):
                worker = threading.Thread(target=run_split)
                worker.start()
                self.assertTrue(write_started.wait(5.0))
                cancel_event.set()
                release_write.set()
                worker.join(10.0)

            self.assertFalse(worker.is_alive())
            self.assertIsInstance(errors.get_nowait(), splitone.JobCancelled)
            self.assertFalse((root / "cancel_split").exists())
            self.assertEqual([], work_folders(root))

    def test_immediate_replacement_ignores_old_job_events(self) -> None:
        from pypdf import PdfWriter

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            old_source = root / "old.pdf"
            new_source = root / "new.pdf"
            create_fixture_pdf(old_source, 300, width_base=200)
            create_fixture_pdf(new_source, 3, width_base=700)

            ui = FakeUi()
            controller = splitone.SplitController(
                config=FakeConfig(root),
                notifier=splitone.WorkerNotifier(),
                service=splitone.PdfSplitService(),
            )
            controller.attach_ui(ui)

            old_write_started = threading.Event()
            release_old_write = threading.Event()
            original_write = PdfWriter.write

            def delay_old_job_only(writer, stream):
                width = int(float(writer.pages[0].mediabox.width))
                if width < 700:
                    old_write_started.set()
                    release_old_write.wait(5.0)
                return original_write(writer, stream)

            with mock.patch.object(PdfWriter, "write", delay_old_job_only):
                controller.start_from_path(old_source)
                self.assertTrue(old_write_started.wait(5.0))
                old_generation = controller.job_generation

                controller.refresh()
                self.assertFalse(controller.busy)
                self.assertIsNone(controller.current_source)
                self.assertGreater(controller.job_generation, old_generation)

                controller.start_from_path(new_source)
                controller.notifier.publish(
                    old_generation,
                    "progress",
                    {"current": 299, "total": 300, "name": "p299.pdf"},
                )
                controller.notifier.publish(
                    old_generation,
                    "done",
                    {"output_folder": str(root / "stale_output")},
                )
                controller.notifier.publish(
                    old_generation,
                    "error",
                    "stale error",
                )
                release_old_write.set()

                def replacement_completed() -> bool:
                    controller.process_worker_events()
                    return bool(ui.completions)

                wait_until(replacement_completed)
                wait_until(lambda: not controller.has_active_workers())
                controller.process_worker_events()

            self.assertEqual(1, len(ui.completions))
            replacement_output = ui.completions[0]
            self.assertIsNotNone(replacement_output)
            self.assertEqual("new_split", replacement_output.name)
            self.assertEqual(3, len(list(replacement_output.glob("p*.pdf"))))
            self.assertFalse((root / "old_split").exists())
            self.assertFalse(any("/300" in detail for detail in ui.progress))
            self.assertEqual([], ui.errors)
            self.assertEqual([], work_folders(root))

    def test_close_cancel_cleans_temp(self) -> None:
        from pypdf import PdfWriter

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "close.pdf"
            create_fixture_pdf(source, 300)
            ui = FakeUi()
            controller = splitone.SplitController(
                config=FakeConfig(root),
                notifier=splitone.WorkerNotifier(),
                service=splitone.PdfSplitService(),
            )
            controller.attach_ui(ui)

            write_started = threading.Event()
            release_write = threading.Event()
            original_write = PdfWriter.write

            def delayed_write(writer, stream):
                write_started.set()
                release_write.wait(5.0)
                return original_write(writer, stream)

            with mock.patch.object(PdfWriter, "write", delayed_write):
                controller.start_from_path(source)
                self.assertTrue(write_started.wait(5.0))
                controller.cancel_all_jobs()
                release_write.set()
                wait_until(lambda: not controller.has_active_workers())

            self.assertFalse((root / "close_split").exists())
            self.assertEqual([], work_folders(root))

    def test_error_cleans_temp_without_formal_output(self) -> None:
        from pypdf import PdfWriter

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "error.pdf"
            create_fixture_pdf(source, 3)

            with mock.patch.object(PdfWriter, "write", side_effect=OSError("boom")):
                with self.assertRaises(splitone.PdfServiceError):
                    self.split(source, root)

            self.assertFalse((root / "error_split").exists())
            self.assertEqual([], work_folders(root))


if __name__ == "__main__":
    unittest.main()
