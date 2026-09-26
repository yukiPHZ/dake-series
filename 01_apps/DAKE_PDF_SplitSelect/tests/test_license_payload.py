"""Compare shipped notices with this build environment's installed wheel."""

import hashlib
from importlib.metadata import distribution
from pathlib import Path
import re
import unittest


APP_DIR = Path(__file__).resolve().parents[1]
DIST_DIR = APP_DIR / "dist" / "DakePDF_Split_Select"
LICENSE_DIR = Path("third_party_licenses") / "pypdfium2-5.13.0"
FORBIDDEN = re.compile(r"fitz|pymupdf|mupdfcpp|_mupdf|pillow|^pil$", re.IGNORECASE)


def sha256(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


class LicensePayloadTests(unittest.TestCase):
    def test_installed_wheel_license_files(self):
        wheel = distribution("pypdfium2")
        self.assertEqual(wheel.version, "5.13.0")
        entries = wheel.metadata.get_all("License-File")
        self.assertEqual(len(entries), 19)
        shipped = set()
        for entry in entries:
            source = Path(wheel.locate_file(f"pypdfium2-5.13.0.dist-info/licenses/{entry}"))
            self.assertTrue(source.is_file(), entry)
            relative = entry.removeprefix("data/windows_x64/")
            shipped.add(relative)
            for root in (APP_DIR, DIST_DIR) if DIST_DIR.is_dir() else (APP_DIR,):
                target = root / LICENSE_DIR / relative
                self.assertTrue(target.is_file(), target)
                self.assertEqual(sha256(target), sha256(source), relative)
        for root in (APP_DIR, DIST_DIR) if DIST_DIR.is_dir() else (APP_DIR,):
            actual = {p.relative_to(root / LICENSE_DIR).as_posix()
                      for p in (root / LICENSE_DIR).rglob("*") if p.is_file()}
            self.assertEqual(actual, shipped)
            self.assertEqual(sha256(root / "THIRD_PARTY_NOTICES.txt"),
                             sha256(APP_DIR / "THIRD_PARTY_NOTICES.txt"))

    def test_no_legacy_renderer_in_source_or_onedir(self):
        for name in ("main.py", "pdf_backend.py", "tests/test_lifecycle.py", "tests/test_packaged_cli.py"):
            text = (APP_DIR / name).read_text(encoding="utf-8")
            self.assertNotRegex(text, r"\bimport\s+fitz\b")
        self.assertNotRegex((APP_DIR / "requirements.txt").read_text(), r"(?i)pymupdf|pillow")
        toc = APP_DIR / "build" / "DakePDF_Split_Select" / "Analysis-00.toc"
        if toc.is_file():
            self.assertIsNone(FORBIDDEN.search(toc.read_text(encoding="utf-8")))
        if not DIST_DIR.is_dir():
            return
        files = [p for p in DIST_DIR.rglob("*") if p.is_file()]
        self.assertFalse([p for p in files if FORBIDDEN.search(p.name)])
        self.assertEqual(len([p for p in files if p.name.lower() == "pdfium.dll"]), 1)


if __name__ == "__main__":
    unittest.main(verbosity=2)
