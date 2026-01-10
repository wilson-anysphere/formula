from __future__ import annotations

import io
import unittest
import zipfile

from tools.corpus.sanitize import SanitizeOptions, sanitize_xlsx_bytes


def _make_minimal_xlsx_with_secrets() -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/connections.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"/>
  <Override PartName="/customXml/item1.xml" ContentType="application/xml"/>
</Types>
""",
        )
        z.writestr(
            "_rels/.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
""",
        )
        z.writestr(
            "xl/workbook.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
""",
        )
        z.writestr(
            "xl/_rels/workbook.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections" Target="connections.xml"/>
</Relationships>
""",
        )
        z.writestr(
            "xl/worksheets/sheet1.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1"><v>123</v></c>
      <c r="B1" t="s"><v>0</v></c>
    </row>
    <row r="2">
      <c r="A2"><f>SUM(A1:A1)</f><v>123</v></c>
    </row>
  </sheetData>
</worksheet>
""",
        )
        z.writestr(
            "xl/worksheets/_rels/sheet1.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/secret" TargetMode="External"/>
</Relationships>
""",
        )
        z.writestr(
            "xl/sharedStrings.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><t>SecretValue</t></si>
</sst>
""",
        )
        z.writestr(
            "xl/connections.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <connection id="1" name="Conn" description="password=hunter2"/>
</connections>
""",
        )
        z.writestr("customXml/item1.xml", "<root>token=abcd</root>")
        z.writestr(
            "docProps/core.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>Alice</dc:creator>
  <cp:lastModifiedBy>Alice</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:modified>
</cp:coreProperties>
""",
        )
    return buf.getvalue()


class SanitizeTests(unittest.TestCase):
    def test_sanitize_removes_common_secret_bearing_parts(self) -> None:
        original = _make_minimal_xlsx_with_secrets()
        sanitized, summary = sanitize_xlsx_bytes(original, options=SanitizeOptions())

        self.assertIn("xl/connections.xml", summary.removed_parts)
        self.assertIn("customXml/item1.xml", summary.removed_parts)

        with zipfile.ZipFile(io.BytesIO(sanitized), "r") as z:
            names = set(z.namelist())
            self.assertNotIn("xl/connections.xml", names)
            self.assertNotIn("customXml/item1.xml", names)

            # Content types should not reference removed parts.
            ct = z.read("[Content_Types].xml").decode("utf-8")
            self.assertNotIn("/xl/connections.xml", ct)
            self.assertNotIn("/customXml/item1.xml", ct)

    def test_sanitize_scrubs_external_relationship_targets(self) -> None:
        original = _make_minimal_xlsx_with_secrets()
        sanitized, _ = sanitize_xlsx_bytes(original, options=SanitizeOptions())

        with zipfile.ZipFile(io.BytesIO(sanitized), "r") as z:
            rels = z.read("xl/worksheets/_rels/sheet1.xml.rels").decode("utf-8")
            self.assertNotIn("example.com", rels)
            self.assertIn("https://redacted.invalid/", rels)

    def test_sanitize_redacts_cell_values_but_preserves_formulas(self) -> None:
        original = _make_minimal_xlsx_with_secrets()
        sanitized, _ = sanitize_xlsx_bytes(original, options=SanitizeOptions())

        with zipfile.ZipFile(io.BytesIO(sanitized), "r") as z:
            from xml.etree import ElementTree as ET

            sheet_root = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
            ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

            # Constant numeric values are normalized.
            a1 = sheet_root.find(".//m:c[@r='A1']/m:v", ns)
            self.assertIsNotNone(a1)
            self.assertEqual(a1.text, "0")

            # Formula text is preserved and cached results removed.
            a2_f = sheet_root.find(".//m:c[@r='A2']/m:f", ns)
            self.assertIsNotNone(a2_f)
            self.assertEqual(a2_f.text, "SUM(A1:A1)")
            a2_v = sheet_root.find(".//m:c[@r='A2']/m:v", ns)
            self.assertIsNone(a2_v)

            sst = z.read("xl/sharedStrings.xml").decode("utf-8")
            self.assertNotIn("SecretValue", sst)
            self.assertIn("REDACTED", sst)


if __name__ == "__main__":
    unittest.main()
