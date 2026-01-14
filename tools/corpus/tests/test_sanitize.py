from __future__ import annotations

import io
import unittest
import zipfile
from pathlib import Path
from unittest import mock

from tools.corpus import sanitize as sanitize_mod
from tools.corpus.sanitize import SanitizeOptions, sanitize_xlsx_bytes, scan_xlsx_bytes_for_leaks
from tools.corpus.util import read_workbook_input


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


def _make_minimal_xlsx_with_cell_images_metadata() -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/cellImages.xml" ContentType="application/vnd.ms-excel.cellImages+xml"/>
</Types>
""",
        )
        z.writestr(
            "xl/cellImages.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cellImages xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:pic>
    <xdr:nvPicPr>
      <xdr:cNvPr id="1" name="Alice" descr="alice@example.com"/>
    </xdr:nvPicPr>
  </xdr:pic>
  <a:txBody>
    <a:p><a:r><a:t>Secret</a:t></a:r></a:p>
  </a:txBody>
</cellImages>
""",
        )
    return buf.getvalue()


def _make_minimal_xlsx_with_cell_images_metadata_numeric_suffix() -> bytes:
    """Minimal XLSX with a forward-compatible `cellimages*.xml` part name.

    Some producers (and our parser for forward compatibility) accept `cellimages1.xml`,
    `cellimages2.xml`, ... in addition to the canonical `cellImages.xml`. The sanitizer
    should treat these as the same leak surface and scrub metadata accordingly.
    """

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/richData/cellimages1.xml" ContentType="application/vnd.ms-excel.cellImages+xml"/>
</Types>
""",
        )
        z.writestr(
            "xl/richData/cellimages1.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cellImages xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:pic>
    <xdr:nvPicPr>
      <xdr:cNvPr id="1" name="Alice" descr="alice@example.com"/>
    </xdr:nvPicPr>
  </xdr:pic>
  <a:txBody>
    <a:p><a:r><a:t>Secret</a:t></a:r></a:p>
  </a:txBody>
</cellImages>
""",
        )
    return buf.getvalue()


def _make_minimal_xlsx_for_sheet_rename() -> bytes:
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
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
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
    <sheet name="Sensitive Sheet" sheetId="1" r:id="rId1"/>
    <sheet name="Other" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>
""",
        )
        z.writestr(
            "xl/_rels/workbook.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>
""",
        )
        z.writestr(
            "xl/worksheets/sheet1.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>123</v></c>
    </row>
  </sheetData>
</worksheet>
""",
        )
        z.writestr(
            "xl/worksheets/sheet2.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f>'Sensitive Sheet'!A1</f><v>123</v></c>
    </row>
  </sheetData>
</worksheet>
""",
        )
        z.writestr(
            "docProps/app.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Excel</Application>
  <Company>ACME Corp</Company>
  <TitlesOfParts>
    <vt:vector size="2" baseType="lpstr">
      <vt:lpstr>Sensitive Sheet</vt:lpstr>
      <vt:lpstr>Other</vt:lpstr>
    </vt:vector>
  </TitlesOfParts>
</Properties>
""",
        )
    return buf.getvalue()


def _make_minimal_xlsx_with_protection_and_sharing() -> bytes:
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
  <fileSharing userName="alice@example.com" readOnlyRecommended="1"/>
  <workbookProtection workbookPassword="ABCDEF" lockStructure="1"/>
  <workbookPr codeName="MySecretWorkbook"/>
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
</Relationships>
""",
        )
        z.writestr(
            "xl/worksheets/sheet1.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetPr codeName="MySecretSheetCode"/>
  <sheetProtection password="ABCDEF" sheet="1"/>
  <sheetData>
    <row r="1"><c r="A1"><v>123</v></c></row>
  </sheetData>
</worksheet>
""",
        )
    return buf.getvalue()


def _make_minimal_xlsx_with_vml_imagedata() -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
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
</Relationships>
""",
        )
        z.writestr(
            "xl/worksheets/sheet1.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
  <legacyDrawing r:id="rId1"/>
</worksheet>
""",
        )
        z.writestr(
            "xl/worksheets/_rels/sheet1.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
</Relationships>
""",
        )
        z.writestr(
            "xl/drawings/vmlDrawing1.vml",
            """<?xml version="1.0" encoding="UTF-8"?>
<xml xmlns:v="urn:schemas-microsoft-com:vml"
     xmlns:o="urn:schemas-microsoft-com:office:office"
     xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <v:shape id="_x0000_s1025" type="#_x0000_t75">
    <v:imagedata o:relid="rId1"/>
  </v:shape>
</xml>
""",
        )
        z.writestr(
            "xl/drawings/_rels/vmlDrawing1.vml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.jpeg"/>
</Relationships>
""",
        )
        z.writestr("xl/media/image1.jpeg", b"JPEGDATA alice@example.com")

    return buf.getvalue()


def _make_minimal_xlsx_with_printer_settings() -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/octet-stream"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/printerSettings/printerSettings1.bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/>
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
</Relationships>
""",
        )
        z.writestr(
            "xl/worksheets/sheet1.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>
""",
        )
        z.writestr(
            "xl/worksheets/_rels/sheet1.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" Target="../printerSettings/printerSettings1.bin"/>
</Relationships>
""",
        )
        z.writestr("xl/printerSettings/printerSettings1.bin", b"PRINTER \\\\server\\\\printer alice@example.com")
    return buf.getvalue()


def _make_minimal_xlsx_with_cell_images() -> bytes:
    """Construct a minimal XLSX containing the in-cell images part (`xl/cellimages.xml`).

    This is a known leak surface because it can embed DrawingML `r:embed` relationship IDs
    pointing at `xl/media/**`. When `remove_secrets=True` the sanitizer drops `xl/media/**`,
    so it must also drop (or sanitize) this part to avoid dangling references.
    """

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/cellImages.xml" ContentType="application/vnd.ms-excel.cellimages+xml"/>
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
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2022/relationships/cellImages" Target="cellImages.xml"/>
</Relationships>
""",
        )
        z.writestr(
            "xl/worksheets/sheet1.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>
""",
        )
        z.writestr(
            "xl/cellimages.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cellImages xmlns="http://schemas.microsoft.com/office/spreadsheetml/2023/cellimages"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <a:blip r:embed="rId1"/>
</cellImages>
""",
        )
        z.writestr(
            "xl/_rels/cellimages.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>
""",
        )
        z.writestr("xl/media/image1.png", b"\x89PNG\r\n\x1a\nPNGDATA")

    return buf.getvalue()


def _make_minimal_xlsx_with_nested_cell_images() -> bytes:
    """Like `_make_minimal_xlsx_with_cell_images`, but stores cellimages under a subdir.

    Some producers may not use the canonical `xl/cellimages.xml` location. The sanitizer
    should still remove the part (and its rels) when `remove_secrets=True`.
    """

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/richData/cellimages.xml" ContentType="application/vnd.ms-excel.cellimages+xml"/>
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
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2022/relationships/cellImages" Target="richData/cellimages.xml"/>
</Relationships>
""",
        )
        z.writestr(
            "xl/worksheets/sheet1.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>
""",
        )
        z.writestr(
            "xl/richData/cellimages.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cellImages xmlns="http://schemas.microsoft.com/office/spreadsheetml/2023/cellimages"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <a:blip r:embed="rId1"/>
</cellImages>
""",
        )
        z.writestr(
            "xl/richData/_rels/cellimages.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>
""",
        )
        z.writestr("xl/media/image1.png", b"\x89PNG\r\n\x1a\nPNGDATA")

    return buf.getvalue()


def _make_minimal_xlsx_with_nested_cell_images_numeric_suffix() -> bytes:
    """Like `_make_minimal_xlsx_with_nested_cell_images`, but uses `cellimages1.xml`.

    Excel may introduce new `cellimages*.xml` part names for forward compatibility.
    When `remove_secrets=True`, media is dropped, so these parts must also be dropped
    to avoid leaving dangling `r:embed` references.
    """

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/richData/cellimages1.xml" ContentType="application/vnd.ms-excel.cellimages+xml"/>
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
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2022/relationships/cellImages" Target="richData/cellimages1.xml"/>
</Relationships>
""",
        )
        z.writestr(
            "xl/worksheets/sheet1.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>
""",
        )
        z.writestr(
            "xl/richData/cellimages1.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cellImages xmlns="http://schemas.microsoft.com/office/spreadsheetml/2023/cellimages"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <a:blip r:embed="rId1"/>
</cellImages>
""",
        )
        z.writestr(
            "xl/richData/_rels/cellimages1.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>
""",
        )
        z.writestr("xl/media/image1.png", b"\x89PNG\r\n\x1a\nPNGDATA")

    return buf.getvalue()


def _make_minimal_xlsx_with_cell_images_in_cell() -> bytes:
    """Minimal XLSX exercising Excel "Images in Cell" (xl/cellImages.xml).

    When remove_secrets drops xl/media/**, the sanitizer must also remove the
    cellImages part, its rels, and any references to it.
    """

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/cellImages.xml" ContentType="application/vnd.ms-excel.cellimage+xml"/>
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
  <Relationship Id="rId2" Type="http://example.invalid/relationships/cellImages" Target="cellImages.xml"/>
</Relationships>
""",
        )
        z.writestr(
            "xl/worksheets/sheet1.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>
""",
        )

        z.writestr(
            "xl/cellImages.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cellImages xmlns="http://schemas.microsoft.com/office/spreadsheetml/2019/11/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cellImage r:id="rId1"/>
</cellImages>
""",
        )
        z.writestr(
            "xl/_rels/cellImages.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>
""",
        )
        z.writestr("xl/media/image1.png", b"PNGDATA")

    return buf.getvalue()


def _make_minimal_xlsx_with_dialogsheet_inline_str() -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/dialogsheets/sheet1.xml" ContentType="application/vnd.ms-excel.dialogsheet+xml"/>
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
    <sheet name="Dialog" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
""",
        )
        z.writestr(
            "xl/_rels/workbook.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet" Target="dialogsheets/sheet1.xml"/>
</Relationships>
""",
        )
        z.writestr(
            "xl/dialogsheets/sheet1.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dialogsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>alice@example.com</t></is></c>
    </row>
  </sheetData>
</dialogsheet>
""",
        )
    return buf.getvalue()


def _make_minimal_xlsx_with_macrosheet_inline_str() -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/macrosheets/sheet1.xml" ContentType="application/vnd.ms-excel.macrosheet+xml"/>
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
    <sheet name="MacroSheet" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
""",
        )
        z.writestr(
            "xl/_rels/workbook.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet" Target="macrosheets/sheet1.xml"/>
</Relationships>
""",
        )
        z.writestr(
            "xl/macrosheets/sheet1.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<macroSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>alice@example.com</t></is></c>
    </row>
  </sheetData>
</macroSheet>
""",
        )
    return buf.getvalue()

def _rewrite_zip_with_leading_slash_entry_names(data: bytes) -> bytes:
    zin_buf = io.BytesIO(data)
    zout_buf = io.BytesIO()
    with zipfile.ZipFile(zin_buf, "r") as zin:
        with zipfile.ZipFile(zout_buf, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                if info.is_dir():
                    continue
                name = info.filename
                new_name = name if name.startswith("/") else f"/{name}"
                zout.writestr(new_name, zin.read(name))
    return zout_buf.getvalue()


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

    def test_sanitize_tolerates_leading_slash_zip_entry_names(self) -> None:
        original = _make_minimal_xlsx_with_secrets()
        rewritten = _rewrite_zip_with_leading_slash_entry_names(original)
        sanitized, summary = sanitize_xlsx_bytes(rewritten, options=SanitizeOptions())

        self.assertIn("xl/connections.xml", summary.removed_parts)
        self.assertIn("customXml/item1.xml", summary.removed_parts)

        with zipfile.ZipFile(io.BytesIO(sanitized), "r") as z:
            names = set(z.namelist())
            # Sanitizer writes normalized part names (no leading '/').
            self.assertIn("xl/workbook.xml", names)
            self.assertNotIn("/xl/workbook.xml", names)
            self.assertNotIn("xl/connections.xml", names)
            self.assertNotIn("customXml/item1.xml", names)

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

    def test_sanitize_scrubs_cell_images_metadata(self) -> None:
        original = _make_minimal_xlsx_with_cell_images_metadata()
        sanitized, _ = sanitize_xlsx_bytes(
            original, options=SanitizeOptions(scrub_metadata=True, remove_secrets=False)
        )

        with zipfile.ZipFile(io.BytesIO(sanitized), "r") as z:
            self.assertIn("xl/cellImages.xml", z.namelist())
            cell_images = z.read("xl/cellImages.xml").decode("utf-8", errors="ignore")
            self.assertNotIn("Alice", cell_images)
            self.assertNotIn("alice@example.com", cell_images)
            self.assertNotIn("Secret", cell_images)

    def test_sanitize_scrubs_cell_images_metadata_with_numeric_suffix(self) -> None:
        original = _make_minimal_xlsx_with_cell_images_metadata_numeric_suffix()
        sanitized, _ = sanitize_xlsx_bytes(
            original, options=SanitizeOptions(scrub_metadata=True, remove_secrets=False)
        )

        with zipfile.ZipFile(io.BytesIO(sanitized), "r") as z:
            self.assertIn("xl/richData/cellimages1.xml", z.namelist())
            cell_images = z.read("xl/richData/cellimages1.xml").decode("utf-8", errors="ignore")
            self.assertNotIn("Alice", cell_images)
            self.assertNotIn("alice@example.com", cell_images)
            self.assertNotIn("Secret", cell_images)

    def test_sanitize_scrubs_pii_surfaces_and_leak_scan_passes(self) -> None:
        fixture_path = Path(__file__).parent / "fixtures" / "pii-surfaces.xlsx.b64"
        wb = read_workbook_input(fixture_path)
        original = wb.data

        sanitized, summary = sanitize_xlsx_bytes(original, options=SanitizeOptions())

        # High-level summary: removed binary/custom UI surfaces + rewrote XML parts.
        self.assertIn("xl/vbaProject.bin", summary.removed_parts)
        self.assertIn("xl/vbaProjectSignature.bin", summary.removed_parts)
        self.assertIn("customUI/customUI.xml", summary.removed_parts)
        self.assertIn("docProps/custom.xml", summary.removed_parts)
        self.assertIn("docProps/thumbnail.jpeg", summary.removed_parts)
        self.assertIn("xl/workbook.xml", summary.rewritten_parts)
        self.assertIn("xl/worksheets/sheet1.xml", summary.rewritten_parts)
        self.assertIn("xl/comments1.xml", summary.rewritten_parts)
        self.assertIn("xl/charts/chart1.xml", summary.rewritten_parts)
        self.assertIn("xl/drawings/drawing1.xml", summary.rewritten_parts)
        self.assertIn("xl/tables/table1.xml", summary.rewritten_parts)
        self.assertIn("xl/pivotCache/pivotCacheDefinition1.xml", summary.rewritten_parts)
        self.assertIn("xl/pivotCache/pivotCacheRecords1.xml", summary.rewritten_parts)

        sensitive_tokens = [
            "alice@example.com",
            "leaky.example.com",
            "ACME_SECRET_NAME",
            "ACME_SECRET_TOKEN",
            "ACME_TABLE_SECRET",
            "ACME_COLUMN_SECRET",
            "PIVOT_TOKEN_SECRET",
            "PIVOT_FIELD_SECRET",
            "CHART_TOKEN_123",
            "DRAWING_TOKEN",
            "VBASECRET",
            "SIGNATURESECRET",
            "THUMBSECRET",
        ]

        with zipfile.ZipFile(io.BytesIO(sanitized), "r") as z:
            # Output ZIP remains readable and contains workbook.xml.
            self.assertIn("xl/workbook.xml", z.namelist())

            # Removed parts should not exist and must be removed from [Content_Types].xml.
            names = set(z.namelist())
            self.assertNotIn("xl/vbaProject.bin", names)
            self.assertNotIn("xl/vbaProjectSignature.bin", names)
            self.assertNotIn("customUI/customUI.xml", names)
            self.assertNotIn("docProps/custom.xml", names)
            self.assertNotIn("docProps/thumbnail.jpeg", names)
            ct = z.read("[Content_Types].xml").decode("utf-8")
            self.assertNotIn("/xl/vbaProject.bin", ct)
            self.assertNotIn("/xl/vbaProjectSignature.bin", ct)
            self.assertNotIn("/customUI/customUI.xml", ct)
            self.assertNotIn("/docProps/custom.xml", ct)
            self.assertNotIn("/docProps/thumbnail.jpeg", ct)

            # Ensure common leak surfaces no longer contain the injected tokens.
            for part in z.namelist():
                data = z.read(part)
                text = data.decode("utf-8", errors="ignore")
                for token in sensitive_tokens:
                    self.assertNotIn(token, text, msg=f"Token {token!r} leaked via {part}")

            # External hyperlink target should be scrubbed.
            rels = z.read("xl/worksheets/_rels/sheet1.xml.rels").decode("utf-8")
            self.assertIn("https://redacted.invalid/", rels)
            self.assertNotIn("leaky.example.com", rels)

        # Leak scanner should detect issues in the original, but not in the sanitized output.
        original_scan = scan_xlsx_bytes_for_leaks(original, plaintext_strings=sensitive_tokens)
        self.assertFalse(original_scan.ok)
        sanitized_scan = scan_xlsx_bytes_for_leaks(sanitized, plaintext_strings=sensitive_tokens)
        self.assertTrue(sanitized_scan.ok)

    def test_hash_strings_is_deterministic_across_runs(self) -> None:
        fixture_path = Path(__file__).parent / "fixtures" / "pii-surfaces.xlsx.b64"
        wb = read_workbook_input(fixture_path)

        options = SanitizeOptions(hash_strings=True, hash_salt="unit-test-salt")
        sanitized_a, _ = sanitize_xlsx_bytes(wb.data, options=options)
        sanitized_b, _ = sanitize_xlsx_bytes(wb.data, options=options)

        with zipfile.ZipFile(io.BytesIO(sanitized_a), "r") as za, zipfile.ZipFile(
            io.BytesIO(sanitized_b), "r"
        ) as zb:
            for part in (
                "xl/sharedStrings.xml",
                "xl/worksheets/sheet1.xml",
                "xl/comments1.xml",
                "xl/tables/table1.xml",
                "xl/pivotCache/pivotCacheDefinition1.xml",
            ):
                self.assertEqual(za.read(part), zb.read(part), msg=f"{part} should be stable")

            # Spot check: known plaintext string literal in a formula becomes a stable H_<digest> token.
            expected = sanitize_mod._hash_text("alice@example.com", salt="unit-test-salt")
            sheet_xml = za.read("xl/worksheets/sheet1.xml").decode("utf-8", errors="ignore")
            self.assertIn(expected, sheet_xml)

    def test_sanitize_output_is_deterministic_wrt_zip_timestamps(self) -> None:
        # ZIP entry timestamps should not depend on wall-clock time; otherwise repeated sanitization
        # produces noisy diffs and can leak ingest time.
        original = _make_minimal_xlsx_with_secrets()

        with mock.patch("zipfile.time.time", return_value=1_000_000_000):
            sanitized_a, _ = sanitize_xlsx_bytes(original, options=SanitizeOptions())
        with mock.patch("zipfile.time.time", return_value=2_000_000_000):
            sanitized_b, _ = sanitize_xlsx_bytes(original, options=SanitizeOptions())

        self.assertEqual(sanitized_a, sanitized_b)

        with zipfile.ZipFile(io.BytesIO(sanitized_a), "r") as z:
            for info in z.infolist():
                if info.is_dir():
                    continue
                self.assertEqual(info.date_time, (1980, 1, 1, 0, 0, 0))

    def test_rename_sheets_updates_workbook_and_formulas(self) -> None:
        original = _make_minimal_xlsx_for_sheet_rename()
        sanitized, _ = sanitize_xlsx_bytes(original, options=SanitizeOptions(rename_sheets=True))

        with zipfile.ZipFile(io.BytesIO(sanitized), "r") as z:
            from xml.etree import ElementTree as ET

            wb = ET.fromstring(z.read("xl/workbook.xml"))
            sheets = [el for el in wb.iter() if el.tag.split("}")[-1] == "sheet"]
            self.assertEqual([s.attrib.get("name") for s in sheets], ["Sheet1", "Sheet2"])

            sheet2 = ET.fromstring(z.read("xl/worksheets/sheet2.xml"))
            ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
            f = sheet2.find(".//m:c[@r='A1']/m:f", ns)
            self.assertIsNotNone(f)
            self.assertEqual(f.text, "Sheet1!A1")

            # app.xml sheet titles should be rewritten too (avoid leaking old names).
            app = z.read("docProps/app.xml").decode("utf-8", errors="ignore")
            self.assertNotIn("Sensitive Sheet", app)
            self.assertNotIn("Other", app)
            self.assertIn("Sheet1", app)
            self.assertIn("Sheet2", app)

    def test_sanitize_strips_protection_and_sharing_metadata(self) -> None:
        original = _make_minimal_xlsx_with_protection_and_sharing()
        sanitized, _ = sanitize_xlsx_bytes(original, options=SanitizeOptions())

        with zipfile.ZipFile(io.BytesIO(sanitized), "r") as z:
            wb_xml = z.read("xl/workbook.xml").decode("utf-8", errors="ignore")
            self.assertNotIn("fileSharing", wb_xml)
            self.assertNotIn("workbookProtection", wb_xml)
            self.assertNotIn("alice@example.com", wb_xml)
            self.assertNotIn("ABCDEF", wb_xml)
            self.assertNotIn("MySecretWorkbook", wb_xml)

            sheet_xml = z.read("xl/worksheets/sheet1.xml").decode("utf-8", errors="ignore")
            self.assertNotIn("sheetProtection", sheet_xml)
            self.assertNotIn("ABCDEF", sheet_xml)
            self.assertNotIn("MySecretSheetCode", sheet_xml)

    def test_remove_secrets_sanitizes_worksheets_even_without_redaction(self) -> None:
        original = _make_minimal_xlsx_with_protection_and_sharing()
        sanitized, _ = sanitize_xlsx_bytes(
            original,
            options=SanitizeOptions(
                redact_cell_values=False,
                remove_secrets=True,
                scrub_metadata=False,
                remove_external_links=False,
            ),
        )

        with zipfile.ZipFile(io.BytesIO(sanitized), "r") as z:
            sheet_xml = z.read("xl/worksheets/sheet1.xml").decode("utf-8", errors="ignore")
            self.assertNotIn("sheetProtection", sheet_xml)

    def test_remove_secrets_strips_vml_imagedata_when_media_removed(self) -> None:
        original = _make_minimal_xlsx_with_vml_imagedata()
        sanitized, _ = sanitize_xlsx_bytes(
            original,
            options=SanitizeOptions(
                redact_cell_values=False,
                remove_secrets=True,
                scrub_metadata=False,
                remove_external_links=False,
            ),
        )

        with zipfile.ZipFile(io.BytesIO(sanitized), "r") as z:
            names = set(z.namelist())
            self.assertNotIn("xl/media/image1.jpeg", names)
            vml = z.read("xl/drawings/vmlDrawing1.vml").decode("utf-8", errors="ignore")
            self.assertNotIn("imagedata", vml)
            rels = z.read("xl/drawings/_rels/vmlDrawing1.vml.rels").decode("utf-8", errors="ignore")
            self.assertNotIn("image1.jpeg", rels)

    def test_leak_scanner_detects_high_risk_patterns(self) -> None:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
            z.writestr("xl/workbook.xml", "alice@example.com https://evil.example.com AKIA0123456789ABCDEF eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiIxMjM0NTY3ODkwIn0.SflKxwRJSMeKKF2QT4fwpMeJf36POk6yJV_adQssw5c \\\\server\\\\share")

        scan = scan_xlsx_bytes_for_leaks(buf.getvalue())
        self.assertFalse(scan.ok)
        kinds = {f.kind for f in scan.findings}
        self.assertIn("email", kinds)
        self.assertIn("url", kinds)
        self.assertIn("aws_key", kinds)
        self.assertIn("jwt", kinds)

    def test_remove_secrets_drops_printer_settings(self) -> None:
        original = _make_minimal_xlsx_with_printer_settings()
        sanitized, summary = sanitize_xlsx_bytes(original, options=SanitizeOptions(remove_secrets=True))

        self.assertIn("xl/printerSettings/printerSettings1.bin", summary.removed_parts)
        with zipfile.ZipFile(io.BytesIO(sanitized), "r") as z:
            names = set(z.namelist())
            self.assertNotIn("xl/printerSettings/printerSettings1.bin", names)
            rels = z.read("xl/worksheets/_rels/sheet1.xml.rels").decode("utf-8", errors="ignore")
            self.assertNotIn("printerSettings1.bin", rels)

    def test_remove_secrets_drops_cellimages_part_to_avoid_dangling_embeds(self) -> None:
        original = _make_minimal_xlsx_with_cell_images()
        sanitized, summary = sanitize_xlsx_bytes(original, options=SanitizeOptions(remove_secrets=True))

        self.assertIn("xl/media/image1.png", summary.removed_parts)
        self.assertIn("xl/cellimages.xml", summary.removed_parts)
        self.assertIn("xl/_rels/cellimages.xml.rels", summary.removed_parts)

        with zipfile.ZipFile(io.BytesIO(sanitized), "r") as z:
            names = set(z.namelist())
            # Media is always removed when remove_secrets=True.
            self.assertFalse(any(n.startswith("xl/media/") for n in names))

            # `xl/cellimages.xml` and its rels part should be removed so we don't leave
            # dangling `r:embed` IDs referencing missing relationships.
            self.assertNotIn("xl/cellimages.xml", names)
            self.assertNotIn("xl/_rels/cellimages.xml.rels", names)

            # Workbook relationships and content types should not reference removed parts.
            wb_rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", errors="ignore")
            self.assertNotIn("cellimages.xml", wb_rels.lower())
            ct = z.read("[Content_Types].xml").decode("utf-8", errors="ignore")
            self.assertNotIn("/xl/cellimages.xml", ct.lower())

    def test_remove_secrets_drops_nested_cellimages_part_to_avoid_dangling_embeds(self) -> None:
        original = _make_minimal_xlsx_with_nested_cell_images()
        sanitized, summary = sanitize_xlsx_bytes(original, options=SanitizeOptions(remove_secrets=True))

        self.assertIn("xl/media/image1.png", summary.removed_parts)
        self.assertIn("xl/richData/cellimages.xml", summary.removed_parts)
        self.assertIn("xl/richData/_rels/cellimages.xml.rels", summary.removed_parts)

        with zipfile.ZipFile(io.BytesIO(sanitized), "r") as z:
            names = set(z.namelist())
            self.assertFalse(any(n.startswith("xl/media/") for n in names))
            self.assertNotIn("xl/richData/cellimages.xml", names)
            self.assertNotIn("xl/richData/_rels/cellimages.xml.rels", names)

            wb_rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", errors="ignore")
            self.assertNotIn("richdata/cellimages.xml", wb_rels.lower())
            ct = z.read("[Content_Types].xml").decode("utf-8", errors="ignore")
            self.assertNotIn("/xl/richdata/cellimages.xml", ct.lower())

    def test_remove_secrets_drops_nested_cellimages_numeric_suffix_part_to_avoid_dangling_embeds(self) -> None:
        original = _make_minimal_xlsx_with_nested_cell_images_numeric_suffix()
        sanitized, summary = sanitize_xlsx_bytes(original, options=SanitizeOptions(remove_secrets=True))

        self.assertIn("xl/media/image1.png", summary.removed_parts)
        self.assertIn("xl/richData/cellimages1.xml", summary.removed_parts)
        self.assertIn("xl/richData/_rels/cellimages1.xml.rels", summary.removed_parts)

        with zipfile.ZipFile(io.BytesIO(sanitized), "r") as z:
            names = set(z.namelist())
            self.assertFalse(any(n.startswith("xl/media/") for n in names))
            self.assertNotIn("xl/richData/cellimages1.xml", names)
            self.assertNotIn("xl/richData/_rels/cellimages1.xml.rels", names)

            wb_rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", errors="ignore")
            self.assertNotIn("richdata/cellimages1.xml", wb_rels.lower())
            ct = z.read("[Content_Types].xml").decode("utf-8", errors="ignore")
            self.assertNotIn("/xl/richdata/cellimages1.xml", ct.lower())

    def test_remove_secrets_removes_cell_images_in_cell_parts(self) -> None:
        original = _make_minimal_xlsx_with_cell_images_in_cell()
        sanitized, _ = sanitize_xlsx_bytes(original, options=SanitizeOptions(remove_secrets=True))

        with zipfile.ZipFile(io.BytesIO(sanitized), "r") as z:
            names = set(z.namelist())
            self.assertNotIn("xl/cellImages.xml", names)
            self.assertNotIn("xl/_rels/cellImages.xml.rels", names)
            self.assertNotIn("xl/media/image1.png", names)

            from xml.etree import ElementTree as ET

            ct_root = ET.fromstring(z.read("[Content_Types].xml"))
            ct_ns = {"ct": "http://schemas.openxmlformats.org/package/2006/content-types"}
            overrides = ct_root.findall("ct:Override", ct_ns)
            part_names = {o.attrib.get("PartName", "") for o in overrides}
            self.assertNotIn("/xl/cellImages.xml", part_names)

            wb_rels_root = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
            rel_ns = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
            targets = {rel.attrib.get("Target", "") for rel in wb_rels_root.findall("r:Relationship", rel_ns)}
            self.assertNotIn("cellImages.xml", targets)

    def test_dialogsheet_cell_values_are_sanitized(self) -> None:
        original = _make_minimal_xlsx_with_dialogsheet_inline_str()
        sanitized, _ = sanitize_xlsx_bytes(original, options=SanitizeOptions())
        with zipfile.ZipFile(io.BytesIO(sanitized), "r") as z:
            dialog_xml = z.read("xl/dialogsheets/sheet1.xml").decode("utf-8", errors="ignore")
            self.assertNotIn("alice@example.com", dialog_xml)
            self.assertIn("REDACTED", dialog_xml)

    def test_macrosheet_cell_values_are_sanitized(self) -> None:
        original = _make_minimal_xlsx_with_macrosheet_inline_str()
        sanitized, _ = sanitize_xlsx_bytes(original, options=SanitizeOptions())
        with zipfile.ZipFile(io.BytesIO(sanitized), "r") as z:
            macro_xml = z.read("xl/macrosheets/sheet1.xml").decode("utf-8", errors="ignore")
            self.assertNotIn("alice@example.com", macro_xml)
            self.assertIn("REDACTED", macro_xml)


if __name__ == "__main__":
    unittest.main()
