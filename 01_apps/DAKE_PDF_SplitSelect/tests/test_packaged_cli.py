"""Exercise an unpublished onedir or extracted onedir with disposable PDFs."""

import hashlib
import os
from pathlib import Path
import subprocess
import tempfile
import unittest

from pypdf import PdfReader

from test_lifecycle import write_fixture


APP_DIR = Path(__file__).resolve().parents[1]
DEFAULT_EXE = APP_DIR / "dist" / "DakePDF_Split_Select" / "DakePDF_Split_Select.exe"


class PackagedCliTests(unittest.TestCase):
    def test_nine_cli_cases(self):
        exe = Path(os.environ.get("DAKE_SPLITSELECT_EXE", str(DEFAULT_EXE)))
        if not exe.is_file():
            self.skipTest(f"onedir candidate not built: {exe}")
        with tempfile.TemporaryDirectory(prefix="dake-splitselect-package-") as directory:
            root = Path(directory)
            source = root / "input.pdf"
            write_fixture(source, 3)
            before = hashlib.sha256(source.read_bytes()).hexdigest()

            def run(*args):
                return subprocess.run(
                    [str(exe), "--from-shimarisu", *map(str, args), "--silent"],
                    cwd=exe.parent, capture_output=True, timeout=15,
                )

            exact = root / "exact.pdf"
            self.assertEqual(run("--inputs", source, "--pages", "1,3", "--output", exact).returncode, 0)
            self.assertEqual(len(PdfReader(exact).pages), 2)

            output_dir = root / "output-dir"
            output_dir.mkdir()
            self.assertEqual(run("--inputs", source, "--pages", "2", "--output", output_dir).returncode, 0)
            self.assertEqual(len(PdfReader(next(output_dir.glob("*.pdf"))).pages), 1)

            self.assertEqual(run("--inputs", source, "--pages", "1").returncode, 0)
            self.assertEqual(len(list(root.glob("input_extract_*.pdf"))), 1)

            exact.write_bytes(b"preserve existing output")
            self.assertEqual(run("--inputs", source, "--pages", "3", "--output", exact).returncode, 0)
            self.assertEqual(exact.read_bytes(), b"preserve existing output")
            self.assertEqual(len(PdfReader(root / "exact_1.pdf").pages), 1)

            for args in (
                ("--inputs", source, "--pages", "1--3"),
                ("--inputs", source, "--pages", "4"),
                ("--inputs", root / "missing.pdf", "--pages", "1"),
                ("--pages", "1"),
                ("--inputs", source),
            ):
                with self.subTest(args=args):
                    self.assertEqual(run(*args).returncode, 1)

            self.assertEqual(hashlib.sha256(source.read_bytes()).hexdigest(), before)


if __name__ == "__main__":
    unittest.main(verbosity=2)
