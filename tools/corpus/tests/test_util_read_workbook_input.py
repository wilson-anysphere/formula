from __future__ import annotations

import base64
import io
import tempfile
import unittest
import zipfile
from pathlib import Path

from tools.corpus.util import read_workbook_input


def _make_minimal_xlsx() -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
</Types>
""",
        )
        z.writestr(
            "xl/workbook.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
</workbook>
""",
        )
    return buf.getvalue()


class ReadWorkbookInputTests(unittest.TestCase):
    def test_reads_b64_case_insensitively(self) -> None:
        with tempfile.TemporaryDirectory(prefix="read-workbook-input-") as td:
            tmp = Path(td)
            xlsx = _make_minimal_xlsx()
            path = tmp / "case.xlsx.B64"
            path.write_bytes(base64.encodebytes(xlsx))

            wb = read_workbook_input(path)
            self.assertEqual(wb.display_name, "case.xlsx")
            self.assertEqual(wb.data, xlsx)

    def test_detects_enc_case_insensitively(self) -> None:
        with tempfile.TemporaryDirectory(prefix="read-workbook-input-") as td:
            tmp = Path(td)
            path = tmp / "case.xlsx.ENC"
            path.write_bytes(b"not-really-encrypted")

            with self.assertRaises(ValueError):
                read_workbook_input(path)


if __name__ == "__main__":
    unittest.main()

