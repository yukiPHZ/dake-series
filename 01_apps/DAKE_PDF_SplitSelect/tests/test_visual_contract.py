"""Static invariants and build metadata compatibility for the Phase 2 candidate."""

import ast
import importlib.util
from pathlib import Path
import re
import unittest

APP_DIR = Path(__file__).resolve().parents[1]


class VisualContractTests(unittest.TestCase):
    def test_no_inline_widget_text(self):
        for filename in ("main.py", "page_preview.py"):
            tree = ast.parse((APP_DIR / filename).read_text(encoding="utf-8"))
            for node in ast.walk(tree):
                if isinstance(node, ast.keyword) and node.arg in ("text", "title", "message"):
                    if isinstance(node.value, ast.Constant) and isinstance(node.value.value, str):
                        self.assertIsNone(re.search(r"[\u3040-\u30ff\u3400-\u9fff]", node.value.value),
                                          f"{filename}:{node.lineno}")

    def test_metadata_override_does_not_change_other_apps_default(self):
        spec = importlib.util.spec_from_file_location("version_info", APP_DIR.parents[1] / "tools" / "generate_version_info.py")
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        meta = module.read_meta(APP_DIR)
        default = module.build_version_info(meta)
        self.assertIn(module.COMPANY_NAME, default)
        self.assertIn(module.COPYRIGHT, default)
        candidate = module.build_version_info(meta, "PEAKHEADZ", "Copyright Yukihiko Kikuta")
        self.assertIn("StringStruct('CompanyName', 'PEAKHEADZ')", candidate)
        self.assertIn("StringStruct('LegalCopyright', 'Copyright Yukihiko Kikuta')", candidate)
        self.assertIn("StringStruct('FileVersion', '1.0.2')", candidate)


if __name__ == "__main__":
    unittest.main(verbosity=2)
