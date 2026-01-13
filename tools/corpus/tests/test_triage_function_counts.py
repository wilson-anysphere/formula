from __future__ import annotations

import io
import unittest
import zipfile

from tools.corpus.triage import _extract_function_counts


class TriageFunctionCountTests(unittest.TestCase):
    def test_extract_function_counts_detects_mixed_case_worksheet_paths(self) -> None:
        # Some malformed XLSX archives store worksheet parts with mixed-case folder/file names.
        # Function fingerprinting should still scan those parts.
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as z:
            z.writestr(
                "XL/WORKSHEETS/SHEET1.XML",
                """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1">
        <f>SUM(1,2)</f>
      </c>
    </row>
  </sheetData>
</worksheet>
""",
            )

        buf.seek(0)
        with zipfile.ZipFile(buf, "r") as z:
            counts = _extract_function_counts(z)

        self.assertEqual(counts.get("SUM"), 1)


if __name__ == "__main__":
    unittest.main()

