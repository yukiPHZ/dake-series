from __future__ import annotations

import hashlib
import tempfile
import unittest
from pathlib import Path

import fitz

from main import StampPosition, save_stamped_pdf


class PdfSmokeTests(unittest.TestCase):
    def test_stamp_is_saved_without_changing_source_pdf(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source = root / "source.pdf"
            output = root / "stamped.pdf"

            with fitz.open() as document:
                document.new_page(width=595, height=842)
                document.save(source)
            source_hash = hashlib.sha256(source.read_bytes()).hexdigest()

            save_stamped_pdf(
                source,
                output,
                StampPosition(page_index=0, x=297.5, y=421),
                "菊田",
                "20261121",
                80,
            )

            self.assertTrue(output.is_file())
            self.assertEqual(hashlib.sha256(source.read_bytes()).hexdigest(), source_hash)
            with fitz.open(output) as stamped:
                self.assertEqual(stamped.page_count, 1)


if __name__ == "__main__":
    unittest.main()
