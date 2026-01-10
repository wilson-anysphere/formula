#!/usr/bin/env python3
"""
Generate a tiny XLSX fixture corpus without external dependencies.

The goal of these fixtures is *round-trip* validation: load → save → diff at the
OpenXML part level. We keep the workbooks intentionally small so they can live
in-repo and run quickly in CI.
"""

from __future__ import annotations

import pathlib
import zipfile


ROOT = pathlib.Path(__file__).resolve().parent
EPOCH = (1980, 1, 1, 0, 0, 0)


def _zip_write(zf: zipfile.ZipFile, name: str, data: str) -> None:
    info = zipfile.ZipInfo(name, date_time=EPOCH)
    info.compress_type = zipfile.ZIP_DEFLATED
    info.create_system = 0
    zf.writestr(info, data.encode("utf-8"))


def write_xlsx(
    path: pathlib.Path,
    sheet_xmls: list[str],
    styles_xml: str,
    *,
    sheet_names: list[str] | None = None,
    shared_strings_xml: str | None = None,
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        path.unlink()

    if sheet_names is None:
        sheet_names = [f"Sheet{i+1}" for i in range(len(sheet_xmls))]
    if len(sheet_names) != len(sheet_xmls):
        raise ValueError("sheet_names must match sheet_xmls length")

    with zipfile.ZipFile(path, "w") as zf:
        _zip_write(
            zf,
            "[Content_Types].xml",
            content_types_xml(
                sheet_count=len(sheet_xmls),
                include_shared_strings=shared_strings_xml is not None,
            ),
        )
        _zip_write(zf, "_rels/.rels", package_rels_xml())
        _zip_write(zf, "docProps/core.xml", core_props_xml())
        _zip_write(zf, "docProps/app.xml", app_props_xml(sheet_names))
        _zip_write(zf, "xl/workbook.xml", workbook_xml(sheet_names))
        _zip_write(
            zf,
            "xl/_rels/workbook.xml.rels",
            workbook_rels_xml(
                sheet_count=len(sheet_xmls),
                include_shared_strings=shared_strings_xml is not None,
            ),
        )
        for idx, sheet_xml in enumerate(sheet_xmls, start=1):
            _zip_write(zf, f"xl/worksheets/sheet{idx}.xml", sheet_xml)
        _zip_write(zf, "xl/styles.xml", styles_xml)
        if shared_strings_xml is not None:
            _zip_write(zf, "xl/sharedStrings.xml", shared_strings_xml)


def content_types_xml(*, sheet_count: int, include_shared_strings: bool) -> str:
    overrides = [
        '  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    ]
    for idx in range(1, sheet_count + 1):
        overrides.append(
            f'  <Override PartName="/xl/worksheets/sheet{idx}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
    overrides.append(
        '  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    )
    if include_shared_strings:
        overrides.append(
            '  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
        )
    overrides.append(
        '  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
    )
    overrides.append(
        '  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
    )

    return (
        """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
%s
</Types>
"""
        % "\n".join(overrides)
    )


def package_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""


def workbook_xml(sheet_names: list[str]) -> str:
    sheets = []
    for idx, name in enumerate(sheet_names, start=1):
        sheets.append(f'    <sheet name="{xml_escape(name)}" sheetId="{idx}" r:id="rId{idx}"/>')
    return (
        """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
%s
  </sheets>
</workbook>
"""
        % "\n".join(sheets)
    )


def workbook_rels_xml(*, sheet_count: int, include_shared_strings: bool) -> str:
    rels = []
    for idx in range(1, sheet_count + 1):
        rels.append(
            f'  <Relationship Id="rId{idx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{idx}.xml"/>'
        )

    next_id = sheet_count + 1
    rels.append(
        f'  <Relationship Id="rId{next_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    )
    next_id += 1

    if include_shared_strings:
        rels.append(
            f'  <Relationship Id="rId{next_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
        )

    return (
        """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
%s
</Relationships>
"""
        % "\n".join(rels)
    )


def core_props_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>fixtures</dc:creator>
  <cp:lastModifiedBy>fixtures</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00Z</dcterms:modified>
</cp:coreProperties>
"""


def app_props_xml(sheet_names: list[str]) -> str:
    titles = []
    for name in sheet_names:
        titles.append(f"      <vt:lpstr>{xml_escape(name)}</vt:lpstr>")

    return (
        """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Formula Fixtures</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs>
    <vt:vector size="2" baseType="variant">
      <vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>
      <vt:variant><vt:i4>%d</vt:i4></vt:variant>
    </vt:vector>
  </HeadingPairs>
  <TitlesOfParts>
    <vt:vector size="%d" baseType="lpstr">
%s
    </vt:vector>
  </TitlesOfParts>
  <Company></Company>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>1.0</AppVersion>
</Properties>
"""
        % (len(sheet_names), len(sheet_names), "\n".join(titles))
    )


def xml_escape(value: str) -> str:
    return (
        value.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def shared_strings_xml(strings: list[str]) -> str:
    sis = []
    for s in strings:
        sis.append(f"  <si><t>{xml_escape(s)}</t></si>")
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="%d" uniqueCount="%d">
%s
</sst>
""" % (len(strings), len(strings), "\n".join(sis))


def styles_minimal_xml() -> str:
    # A conservative minimal style sheet: enough for Excel/LibreOffice to accept
    # the file while keeping the structure small.
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="minor"/>
    </font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>
"""


def styles_bold_cell_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="minor"/>
    </font>
    <font>
      <b/>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="minor"/>
    </font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>
"""


def sheet_basic_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1" t="inlineStr"><is><t>Hello</t></is></c>
    </row>
  </sheetData>
</worksheet>
"""


def sheet_formulas_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><v>2</v></c>
      <c r="C1">
        <f>A1+B1</f>
        <v>3</v>
      </c>
    </row>
  </sheetData>
</worksheet>
"""


def sheet_conditional_formatting_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>5</v></c>
      <c r="A2"><v>150</v></c>
    </row>
  </sheetData>
  <conditionalFormatting sqref="A1:A2">
    <cfRule type="cellIs" priority="1" operator="greaterThan">
      <formula>100</formula>
    </cfRule>
  </conditionalFormatting>
</worksheet>
"""


def sheet_styles_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" s="1" t="inlineStr"><is><t>Bold</t></is></c>
    </row>
  </sheetData>
</worksheet>
"""


def sheet_shared_strings_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
    </row>
  </sheetData>
</worksheet>
"""


def sheet_two_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>2</v></c>
    </row>
  </sheetData>
</worksheet>
"""


def sheet_chart_data_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Category</t></is></c>
      <c r="B1" t="inlineStr"><is><t>Value</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>A</t></is></c>
      <c r="B2"><v>10</v></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>B</t></is></c>
      <c r="B3"><v>20</v></c>
    </row>
    <row r="4">
      <c r="A4" t="inlineStr"><is><t>C</t></is></c>
      <c r="B4"><v>30</v></c>
    </row>
  </sheetData>
  <drawing r:id="rId1"/>
</worksheet>
"""


def sheet1_drawing_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>
"""


def drawing1_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>2</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>1</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>8</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>15</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="Chart 1"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm>
        <a:off x="0" y="0"/>
        <a:ext cx="0" cy="0"/>
      </xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                   xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                   r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>
"""


def drawing1_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>
"""


def chart1_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:autoTitleDeleted val="1"/>
    <c:plotArea>
      <c:layout/>
      <c:pieChart>
        <c:varyColors val="1"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx>
            <c:strRef>
              <c:f>Sheet1!$B$1</c:f>
            </c:strRef>
          </c:tx>
          <c:cat>
            <c:strRef>
              <c:f>Sheet1!$A$2:$A$4</c:f>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$B$2:$B$4</c:f>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:firstSliceAng val="0"/>
      </c:pieChart>
    </c:plotArea>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="zero"/>
    <c:showDLblsOverMax val="0"/>
  </c:chart>
</c:chartSpace>
"""


def content_types_chart_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""


def write_chart_xlsx(path: pathlib.Path) -> None:
    sheet_names = ["Sheet1"]
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        path.unlink()

    with zipfile.ZipFile(path, "w") as zf:
        _zip_write(zf, "[Content_Types].xml", content_types_chart_xml())
        _zip_write(zf, "_rels/.rels", package_rels_xml())
        _zip_write(zf, "docProps/core.xml", core_props_xml())
        _zip_write(zf, "docProps/app.xml", app_props_xml(sheet_names))
        _zip_write(zf, "xl/workbook.xml", workbook_xml(sheet_names))
        _zip_write(
            zf,
            "xl/_rels/workbook.xml.rels",
            workbook_rels_xml(sheet_count=1, include_shared_strings=False),
        )
        _zip_write(zf, "xl/worksheets/sheet1.xml", sheet_chart_data_xml())
        _zip_write(zf, "xl/worksheets/_rels/sheet1.xml.rels", sheet1_drawing_rels_xml())
        _zip_write(zf, "xl/drawings/drawing1.xml", drawing1_xml())
        _zip_write(zf, "xl/drawings/_rels/drawing1.xml.rels", drawing1_rels_xml())
        _zip_write(zf, "xl/charts/chart1.xml", chart1_xml())
        _zip_write(zf, "xl/styles.xml", styles_minimal_xml())


def main() -> None:
    write_xlsx(
        ROOT / "basic" / "basic.xlsx",
        [sheet_basic_xml()],
        styles_minimal_xml(),
    )
    write_xlsx(
        ROOT / "basic" / "shared-strings.xlsx",
        [sheet_shared_strings_xml()],
        styles_minimal_xml(),
        shared_strings_xml=shared_strings_xml(["Hello", "World"]),
    )
    write_xlsx(
        ROOT / "basic" / "multi-sheet.xlsx",
        [sheet_basic_xml(), sheet_two_xml()],
        styles_minimal_xml(),
        sheet_names=["Sheet1", "Sheet2"],
    )
    write_xlsx(
        ROOT / "formulas" / "formulas.xlsx",
        [sheet_formulas_xml()],
        styles_minimal_xml(),
    )
    write_xlsx(
        ROOT / "conditional-formatting" / "conditional-formatting.xlsx",
        [sheet_conditional_formatting_xml()],
        styles_minimal_xml(),
    )
    write_xlsx(
        ROOT / "styles" / "styles.xlsx",
        [sheet_styles_xml()],
        styles_bold_cell_xml(),
    )
    write_chart_xlsx(ROOT / "charts" / "basic-chart.xlsx")

    # Directory scaffold for future corpora (kept empty for now).
    for name in ["charts", "pivots", "macros"]:
        (ROOT / name).mkdir(parents=True, exist_ok=True)

    print("Generated XLSX fixtures under", ROOT)


if __name__ == "__main__":
    main()
