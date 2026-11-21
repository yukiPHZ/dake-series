from __future__ import annotations

import ast
import unittest
from pathlib import Path


PRODUCT_ROOT = Path(__file__).resolve().parents[1]


def top_level_constants() -> dict[str, object]:
    tree = ast.parse((PRODUCT_ROOT / "main.py").read_text(encoding="utf-8-sig"))
    values: dict[str, object] = {}

    def evaluate(node: ast.AST) -> object:
        if isinstance(node, ast.Constant):
            return node.value
        if isinstance(node, ast.Name) and node.id in values:
            return values[node.id]
        if isinstance(node, ast.Dict):
            return {evaluate(key): evaluate(value) for key, value in zip(node.keys, node.values)}
        if isinstance(node, ast.List):
            return [evaluate(item) for item in node.elts]
        if isinstance(node, ast.Tuple):
            return tuple(evaluate(item) for item in node.elts)
        raise ValueError

    for node in tree.body:
        if not isinstance(node, ast.Assign) or len(node.targets) != 1 or not isinstance(node.targets[0], ast.Name):
            continue
        try:
            values[node.targets[0].id] = evaluate(node.value)
        except (ValueError, TypeError, KeyError):
            continue
    return values


class ProductQualityTests(unittest.TestCase):
    def test_user_facing_branding_and_error_heading(self) -> None:
        values = top_level_constants()
        self.assertEqual(values["COPYRIGHT"], "© Yukihiko Kikuta / PEAKHEADZ")
        ui_text = values["UI_TEXT"]
        self.assertEqual(ui_text["status_error"], "内容をご確認ください")
        self.assertEqual(ui_text["dialog_error_title"], "確認してください")


if __name__ == "__main__":
    unittest.main()
