"""Compare shipped notices with this build environment's installed wheel."""

import hashlib
from importlib.metadata import distribution
from pathlib import Path
import re
import unittest


APP_DIR = Path(__file__).resolve().parents[1]
ONEFILE_EXE = APP_DIR / "dist" / "DakePDF_Split_Select.exe"
DIST_DIR = ONEFILE_EXE.parent if ONEFILE_EXE.is_file() else APP_DIR / "dist" / "DakePDF_Split_Select"
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

    def test_onefile_embeds_runtime_and_licenses(self):
        if not ONEFILE_EXE.is_file():
            self.skipTest("onefile candidate not built")
        from PyInstaller.archive.readers import CArchiveReader

        archive = CArchiveReader(str(ONEFILE_EXE))
        names = set(archive.toc)
        self.assertEqual(len([name for name in names if Path(name).name.lower() == "pdfium.dll"]), 1)
        self.assertIn("python312.dll", names)
        self.assertTrue(any("tkdnd" in name.lower() and name.endswith(".dll") for name in names))
        self.assertFalse([name for name in names if FORBIDDEN.search(Path(name).name)])
        for source in (APP_DIR / LICENSE_DIR).rglob("*"):
            if source.is_file():
                member = str(source.relative_to(APP_DIR))
                self.assertEqual(archive.extract(member), source.read_bytes(), member)
        self.assertEqual(archive.extract("THIRD_PARTY_NOTICES.txt"),
                         (APP_DIR / "THIRD_PARTY_NOTICES.txt").read_bytes())
        self.assertEqual(archive.extract("dake_icon.ico"),
                         (APP_DIR.parents[1] / "02_assets" / "dake_icon.ico").read_bytes())

    def test_no_legacy_renderer_in_source_or_distribution(self):
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
        if not ONEFILE_EXE.is_file():
            self.assertEqual(len([p for p in files if p.name.lower() == "pdfium.dll"]), 1)


if __name__ == "__main__":
    unittest.main(verbosity=2)
