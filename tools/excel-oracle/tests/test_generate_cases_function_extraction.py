from __future__ import annotations

import importlib.util
import sys
import unittest
from pathlib import Path


class GenerateCasesFunctionExtractionTests(unittest.TestCase):
    def _load_generator(self):
        repo_root = Path(__file__).resolve().parents[3]
        generator_path = repo_root / "tools/excel-oracle/generate_cases.py"
        self.assertTrue(generator_path.is_file(), f"generate_cases.py not found at {generator_path}")

        # `generate_cases.py` imports `case_generators` as a top-level package, so ensure the script
        # directory is importable.
        sys_path_before = list(sys.path)
        sys.path.insert(0, str(generator_path.parent))
        try:
            spec = importlib.util.spec_from_file_location("_excel_oracle_generate_cases", generator_path)
            self.assertIsNotNone(spec)
            self.assertIsNotNone(spec.loader)
            module = importlib.util.module_from_spec(spec)
            sys.modules[spec.name] = module
            spec.loader.exec_module(module)  # type: ignore[union-attr]
        finally:
            sys.path[:] = sys_path_before
        return module

    def test_extract_function_names_ignores_string_literals(self) -> None:
        module = self._load_generator()

        names = module._extract_function_names('=TEXT("SUMM(1,2)")')
        self.assertEqual(names, ["TEXT"], "Function-like tokens inside string literals must be ignored")

        names = module._extract_function_names('=IF(A1="NOW()",1,2)')
        self.assertEqual(names, ["IF"], "Function-like tokens inside string literals must be ignored")

        # Escaped quotes (\"\") should also be handled.
        names = module._extract_function_names('=TEXT("a\"\"b SUM(1)")')
        self.assertEqual(names, ["TEXT"])

    def test_extract_function_names_ignores_sheet_and_bracket_references(self) -> None:
        module = self._load_generator()

        # Parentheses inside quoted sheet names must not be treated as function calls.
        names = module._extract_function_names("='Sheet(1)'!A1")
        self.assertEqual(names, [])

        # Bracketed workbook references can also contain parentheses.
        names = module._extract_function_names("=SUM([Book(1)]Sheet1!A1)")
        self.assertEqual(names, ["SUM"])

        # Structured references can contain parentheses in column names.
        names = module._extract_function_names("=SUM(Table1[Column(1)])")
        self.assertEqual(names, ["SUM"])

        # Nested structured-reference brackets are common in Excel.
        names = module._extract_function_names("=SUM(Table1[[#Headers],[Column(1)]])")
        self.assertEqual(names, ["SUM"])


if __name__ == "__main__":
    unittest.main()
