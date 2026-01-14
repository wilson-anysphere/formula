from __future__ import annotations

import unittest

from tools.corpus.ingest import _normalize_workbook_extension


class IngestFilenamePrivacyTests(unittest.TestCase):
    def test_normalize_workbook_extension_preserves_only_true_extension(self) -> None:
        self.assertEqual(_normalize_workbook_extension("book.xlsx"), ".xlsx")
        self.assertEqual(_normalize_workbook_extension("book.xlsm"), ".xlsm")
        self.assertEqual(_normalize_workbook_extension("book.xlsb"), ".xlsb")

        # Dot-separated identifiers should not be preserved as part of the extension.
        self.assertEqual(_normalize_workbook_extension("acme.com.xlsx"), ".xlsx")

        # Case-insensitive.
        self.assertEqual(_normalize_workbook_extension("BOOK.XLSB"), ".xlsb")

        # Unknown/missing extensions fall back to `.xlsx`.
        self.assertEqual(_normalize_workbook_extension("book"), ".xlsx")
        self.assertEqual(_normalize_workbook_extension("book.weird"), ".xlsx")


if __name__ == "__main__":
    unittest.main()

